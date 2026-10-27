using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using GraphMobileSync.Services;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Appointments;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace GraphMobileSync
{
    public sealed partial class MainPage : Page
    {
        private string _accessToken;
        private bool _hasWriteAccess;
        private bool _isProActive;
        private bool _loadingSettings;
        private bool _loadingCalendarForm;
        private readonly DispatcherTimer _activityRefreshTimer;

        public MainPage()
        {
            InitializeComponent();
            _activityRefreshTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(5)
            };
            _activityRefreshTimer.Tick += ActivityRefreshTimer_Tick;
            InitializeNewEventForm();
            BuildVersionTextBlock.Text = GetBuildVersionText();
            LoadSettings();
            Loaded += MainPage_Loaded;
            Unloaded += MainPage_Unloaded;
        }

        private void LoadSettings()
        {
            LoadSyncSettings();
            bool hasSavedSession = OAuthService.HasSavedSession();
            AccountStatusTextBlock.Text = hasSavedSession
                ? "Restoring your saved Microsoft account sign-in..."
                : "You are not signed in.";
            SignInButton.IsEnabled = !hasSavedSession;
            SignOutButton.IsEnabled = hasSavedSession;
            string backgroundStatus = BackgroundSyncDiagnostics.GetLastStatus();

            if (!string.IsNullOrEmpty(backgroundStatus))
            {
                StatusTextBlock.Text = backgroundStatus;
            }

            RefreshSyncHistory();
            RefreshBackgroundTaskStatus();
        }

        private async void SignInButton_Click(object sender, RoutedEventArgs e)
        {
            SetBusy(true, "Opening Microsoft sign-in...");

            try
            {
                string clientId = AppConfiguration.ClientId;
                string tenant = AppConfiguration.Tenant;
                OAuthService.ValidateSettings(clientId, tenant);

                _accessToken = await OAuthService.SignInAsync(
                    clientId,
                    tenant,
                    promptForAccountSelection: !OAuthService.HasSavedSession());
                _hasWriteAccess = false;
                SyncButton.IsEnabled = true;
                AccountStatusTextBlock.Text = "You are signed in to your Microsoft account.";
                StatusTextBlock.Text = "Signed in. You can now sync contacts and calendars.";
            }
            catch (Exception exception)
            {
                AccountStatusTextBlock.Text = "Sign-in failed. Please try again.";
                StatusTextBlock.Text = exception.Message;
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            SetBusy(true, "Signing out...");

            try
            {
                OAuthService.SignOut();
                _accessToken = null;
                _hasWriteAccess = false;
                SetTwoWaySyncToggle(false);
                SaveSyncSettings();
                BackgroundSyncRegistration.Unregister();
                AccountStatusTextBlock.Text =
                    "Signed out. Existing LiveSync contacts and calendar events remain on this phone.";
                StatusTextBlock.Text =
                    "Signed out. Sign in to sync a different Microsoft account.";
                BackgroundSyncDiagnostics.Record(
                    "Signed out and removed LiveSync's saved Microsoft Graph authorization.");
                await RefreshNewEventCalendarsAsync();
            }
            catch (Exception exception)
            {
                StatusTextBlock.Text = "Sign-out failed: " + exception.Message;
            }
            finally
            {
                RefreshSyncHistory();
                RefreshBackgroundTaskStatus();
                SetBusy(false, null);
            }
        }

        private async void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= MainPage_Loaded;
            await RefreshProStatusAsync();
            await RefreshNewEventCalendarsAsync();

            if (!OAuthService.HasSavedSession())
            {
                AccountStatusTextBlock.Text = "You are not signed in.";
                return;
            }

            SetBusy(true, "Restoring saved sign-in...");

            try
            {
                await RestoreAccessTokenAsync();
                AccountStatusTextBlock.Text = "You are signed in to your Microsoft account.";
                StatusTextBlock.Text = _hasWriteAccess
                    ? "Signed in with your saved session. Two-way sync is ready."
                    : "Signed in with your saved session. You can sync contacts and calendars.";

                if (SyncSettings.Load().BackgroundSyncEnabled)
                {
                    await BackgroundSyncRegistration.RegisterAsync();
                    RefreshBackgroundTaskStatus();
                }
            }
            catch (Exception exception)
            {
                if (!string.IsNullOrEmpty(_accessToken))
                {
                    StatusTextBlock.Text = "Signed in, but background sync could not be registered: " + exception.Message;
                    BackgroundSyncDiagnostics.Record("Background registration failed: " + exception.Message);
                }
                else
                {
                    AccountStatusTextBlock.Text =
                        "Your saved sign-in needs attention. Sign in again to continue.";
                    StatusTextBlock.Text =
                        "Your saved session could not be restored. Check your connection, then sign in again if needed.";
                }
            }
            finally
            {
                await RefreshNewEventCalendarsAsync();
                RefreshBackgroundTaskStatus();
                SetBusy(false, null);
            }
        }

        private async void SyncButton_Click(object sender, RoutedEventArgs e)
        {
            SetBusy(true, "Downloading Microsoft Graph data...");

            try
            {
                SyncSettings syncSettings = SaveSyncSettings();
                Progress<SyncProgress> syncProgress = new Progress<SyncProgress>(UpdateSyncProgress);
                SyncSummary summary;

                using (GraphClient graphClient = new GraphClient(_accessToken))
                {
                    if (syncSettings.TwoWaySyncEnabled && _hasWriteAccess)
                    {
                        StatusTextBlock.Text =
                            "Reconciling Microsoft Graph changes before the full sync...";
                        await new DeltaSyncService(graphClient).PollAndApplyAsync(
                            ignoreBackgroundSetting: true,
                            outboundWritesAvailable: true);
                    }

                    GraphSyncService syncService = new GraphSyncService(graphClient);
                    summary = await syncService.SyncAsync(syncProgress);
                    StatusTextBlock.Text = "Establishing Microsoft Graph change tracking...";
                    await DeltaCheckpointService.CreateAsync(graphClient);

                    if (syncSettings.TwoWaySyncEnabled && _hasWriteAccess)
                    {
                        await new TwoWaySyncService(graphClient).InitializeAsync(syncSettings);
                    }
                }
                await RefreshNewEventCalendarsAsync();

                if (syncSettings.BackgroundSyncEnabled)
                {
                    await BackgroundSyncRegistration.RegisterAsync();
                }
                else
                {
                    BackgroundSyncRegistration.Unregister();
                }

                int duplicateContactListCount = summary.CleanupPlan.ContactListIds.Distinct().Count();

                if (duplicateContactListCount > 0)
                {
                    await GraphSyncService.CleanupDuplicateContactListsAsync(
                        summary.CleanupPlan.ContactListIds);
                }

                string summaryText = string.Format(
                    "Sync complete: {0} contacts ({1} photos downloaded, {2} saved to the People app; {3} contacts have no Graph photo), {4} events added to {5} calendars, {6} duplicate contacts removed, and {7} duplicate LiveSync lists cleared.",
                    summary.Contacts,
                    summary.ContactPhotosDownloaded,
                    summary.ContactPhotosStored,
                    summary.ContactsWithoutPhotos,
                    summary.Events,
                    summary.Calendars,
                    summary.DuplicateContactsRemoved,
                    duplicateContactListCount);
                BackgroundSyncDiagnostics.Record(syncSettings.BackgroundSyncEnabled
                    ? "Background change checks are registered."
                    : "Manual sync completed; background sync is disabled.");
                StatusTextBlock.Text = summaryText;
                await ConfirmDuplicateCalendarCleanupAsync(summary.CleanupPlan);
            }
            catch (Exception exception)
            {
                StatusTextBlock.Text = exception.Message;
            }
            finally
            {
                RefreshSyncHistory();
                RefreshBackgroundTaskStatus();
                SetBusy(false, null);
            }
        }

        private async void BackgroundCheckButton_Click(object sender, RoutedEventArgs e)
        {
            SetBusy(true, "Checking Microsoft Graph for changes...");

            try
            {
                SaveSyncSettings();
                DeltaSyncResult result;

                using (GraphClient graphClient = new GraphClient(_accessToken))
                {
                    result = await new DeltaSyncService(graphClient).PollAndApplyAsync(
                        ignoreBackgroundSetting: true,
                        outboundWritesAvailable: _hasWriteAccess);
                }

                BackgroundSyncDiagnostics.Record("Manual background check: " + result.Message);
                StatusTextBlock.Text = result.Message;
            }
            catch (Exception exception)
            {
                string message = "Graph change check failed: " + exception.Message;
                BackgroundSyncDiagnostics.Record(message);
                StatusTextBlock.Text = message;
            }
            finally
            {
                RefreshSyncHistory();
                RefreshBackgroundTaskStatus();
                SetBusy(false, null);
            }
        }

        private async void ProUpgradeButton_Click(object sender, RoutedEventArgs e)
        {
            SetProBusy(true);
            ProStatusTextBlock.Text = "Opening Microsoft Store...";

            try
            {
                LiveSyncProPurchaseResult result = await LiveSyncProService.PurchaseAsync();

                if (result.IsSuccessful)
                {
                    await RefreshProStatusAsync();
                    await RefreshNewEventCalendarsAsync();
                }
                else
                {
                    ProStatusTextBlock.Text = result.Message;
                }
            }
            catch (Exception exception)
            {
                ProStatusTextBlock.Text =
                    "LiveSync Pro could not be purchased: " + exception.Message;
            }
            finally
            {
                SetProBusy(false);
            }
        }

        private async void ProRefreshButton_Click(object sender, RoutedEventArgs e)
        {
            SetProBusy(true);

            try
            {
                await RefreshProStatusAsync();
                await RefreshNewEventCalendarsAsync();
            }
            finally
            {
                SetProBusy(false);
            }
        }

        private async void TwoWaySyncToggleSwitch_Toggled(object sender, RoutedEventArgs e)
        {
            if (_loadingSettings)
            {
                return;
            }

            if (!TwoWaySyncToggleSwitch.IsOn)
            {
                SaveSyncSettings();
                StatusTextBlock.Text =
                    "Two-way sync is disabled. LiveSync will continue to download Microsoft Graph changes.";
                return;
            }

            SetBusy(true, "Checking LiveSync Pro and requesting Microsoft Graph write access...");
            TwoWaySyncToggleSwitch.IsEnabled = false;
            ProStatusTextBlock.Text =
                "Authorizing Microsoft Graph write access and preparing two-way sync...";

            try
            {
                LiveSyncProEntitlement entitlement =
                    await LiveSyncProService.GetEntitlementAsync();

                if (!entitlement.IsActive)
                {
                    SetTwoWaySyncToggle(false);
                    SaveSyncSettings();
                    ProStatusTextBlock.Text =
                        "LiveSync Pro is required to enable two-way synchronization.";
                    StatusTextBlock.Text =
                        "Subscribe to LiveSync Pro, then refresh your subscription status to enable two-way sync.";
                    return;
                }

                string clientId = AppConfiguration.ClientId;
                string tenant = AppConfiguration.Tenant;
                OAuthService.ValidateSettings(clientId, tenant);
                _accessToken = await OAuthService.SignInAsync(
                    clientId,
                    tenant,
                    requestWriteAccess: true);
                _hasWriteAccess = true;
                AccountStatusTextBlock.Text = "You are signed in to your Microsoft account.";
                SyncSettings settings = SaveSyncSettings();

                using (GraphClient graphClient = new GraphClient(_accessToken))
                {
                    await new TwoWaySyncService(graphClient).InitializeAsync(settings);
                }

                StatusTextBlock.Text =
                    "Two-way sync is enabled. LiveSync will now sync supported local edits back to Microsoft Graph.";
                ProStatusTextBlock.Text =
                    "LiveSync Pro is active. Two-way synchronization is enabled.";
                await RefreshNewEventCalendarsAsync();
            }
            catch (Exception exception)
            {
                _hasWriteAccess = false;
                SetTwoWaySyncToggle(false);
                SaveSyncSettings();
                StatusTextBlock.Text =
                    "Two-way sync could not be enabled: " + exception.Message;
                ProStatusTextBlock.Text =
                    "Two-way sync could not be enabled: " + exception.Message;
                await RefreshNewEventCalendarsAsync();
            }
            finally
            {
                TwoWaySyncToggleSwitch.IsEnabled = _isProActive;
                SetBusy(false, null);
            }
        }

        private async Task RestoreAccessTokenAsync()
        {
            SyncSettings settings = SyncSettings.Load();

            if (!settings.TwoWaySyncEnabled)
            {
                _accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                _hasWriteAccess = false;
                return;
            }

            try
            {
                LiveSyncProEntitlement entitlement =
                    await LiveSyncProService.GetEntitlementAsync();

                if (!entitlement.IsActive)
                {
                    _accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                    _hasWriteAccess = false;
                    BackgroundSyncDiagnostics.Record(
                        "LiveSync Pro is not active; Microsoft Graph write access was not requested.");
                    return;
                }

                try
                {
                    _accessToken = await OAuthService.GetBackgroundAccessTokenAsync(
                        requestWriteAccess: true);
                    _hasWriteAccess = true;
                }
                catch (Exception exception)
                {
                    _accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                    _hasWriteAccess = false;
                    BackgroundSyncDiagnostics.Record(
                        "Microsoft Graph write authorization needs attention: " + exception.Message);
                }
            }
            catch (Exception exception)
            {
                _accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                _hasWriteAccess = false;
                BackgroundSyncDiagnostics.Record(
                    "LiveSync Pro status could not be verified; Microsoft Graph write access was not requested: "
                    + exception.Message);
            }
        }

        private void SetBusy(bool isBusy, string status)
        {
            SignInButton.IsEnabled = !isBusy && string.IsNullOrEmpty(_accessToken);
            SignOutButton.IsEnabled = !isBusy && OAuthService.HasSavedSession();
            SyncButton.IsEnabled = !isBusy && !string.IsNullOrEmpty(_accessToken);
            BackgroundCheckButton.IsEnabled = !isBusy && !string.IsNullOrEmpty(_accessToken);
            SyncProgressBar.IsIndeterminate = isBusy;
            SyncProgressBar.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
            SyncProgressTextBlock.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

            if (!string.IsNullOrEmpty(status))
            {
                StatusTextBlock.Text = status;
            }
        }

        private async void CreateEventButton_Click(object sender, RoutedEventArgs e)
        {
            SetCalendarBusy(true, "Creating calendar event...");

            try
            {
                SyncSettings settings = SyncSettings.Load();

                if (!settings.TwoWaySyncEnabled || !_hasWriteAccess)
                {
                    throw new InvalidOperationException(
                        "Enable LiveSync Pro two-way synchronization before creating events here.");
                }

                LiveSyncProEntitlement entitlement =
                    await LiveSyncProService.GetEntitlementAsync();

                if (!entitlement.IsActive)
                {
                    throw new InvalidOperationException(
                        "LiveSync Pro is not active, so the event cannot be uploaded.");
                }

                ComboBoxItem selectedCalendar =
                    NewEventCalendarComboBox.SelectedItem as ComboBoxItem;
                string graphCalendarId = selectedCalendar == null
                    ? string.Empty
                    : selectedCalendar.Tag as string;

                if (string.IsNullOrEmpty(graphCalendarId))
                {
                    throw new InvalidOperationException("Select a LiveSync calendar.");
                }

                string subject = NewEventSubjectTextBox.Text.Trim();

                if (string.IsNullOrEmpty(subject))
                {
                    throw new InvalidOperationException("Enter an event subject.");
                }

                DateTimeOffset startTime;
                DateTimeOffset endTime;

                if (NewEventAllDayToggleSwitch.IsOn)
                {
                    startTime = CreateNewEventDateTime(NewEventStartDatePicker, TimeSpan.Zero);
                    endTime = CreateNewEventDateTime(NewEventEndDatePicker, TimeSpan.Zero).AddDays(1);
                }
                else
                {
                    startTime = CreateNewEventDateTime(
                        NewEventStartDatePicker,
                        NewEventStartTimePicker.Time);
                    endTime = CreateNewEventDateTime(
                        NewEventEndDatePicker,
                        NewEventEndTimePicker.Time);
                }

                if (endTime <= startTime)
                {
                    throw new InvalidOperationException("The event end must be after its start.");
                }

                SyncMappings mappings = await SyncMappingsStore.LoadAsync();
                CalendarMapping calendarMapping;

                if (!mappings.Calendars.TryGetValue(graphCalendarId, out calendarMapping))
                {
                    throw new InvalidOperationException(
                        "The selected calendar is no longer available. Run a full sync and try again.");
                }

                AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                    AppointmentStoreAccessType.AppCalendarsReadWrite);

                if (appointmentStore == null)
                {
                    throw new InvalidOperationException(
                        "Calendar access was not granted. Enable calendar access for LiveSync and try again.");
                }

                AppointmentCalendar calendar = (await appointmentStore.FindAppointmentCalendarsAsync())
                    .FirstOrDefault(candidate => candidate.LocalId == calendarMapping.LocalId);

                if (calendar == null)
                {
                    throw new InvalidOperationException(
                        "The selected LiveSync calendar is unavailable. Run a full sync and try again.");
                }

                Appointment appointment = new Appointment
                {
                    Subject = subject,
                    Details = NewEventDetailsTextBox.Text.Trim(),
                    Location = NewEventLocationTextBox.Text.Trim(),
                    StartTime = startTime,
                    Duration = endTime - startTime,
                    AllDay = NewEventAllDayToggleSwitch.IsOn
                };
                await calendar.SaveAppointmentAsync(appointment);

                using (GraphClient graphClient = new GraphClient(_accessToken))
                {
                    await new TwoWaySyncService(graphClient).CreateAppointmentAsync(
                        graphCalendarId,
                        appointment);
                }

                NewEventSubjectTextBox.Text = string.Empty;
                NewEventLocationTextBox.Text = string.Empty;
                NewEventDetailsTextBox.Text = string.Empty;
                NewEventStatusTextBlock.Text =
                    "Event created and uploaded to Microsoft Graph.";
                BackgroundSyncDiagnostics.Record(
                    "Created a calendar event and uploaded it to Microsoft Graph.");
                RefreshSyncHistory();
            }
            catch (Exception exception)
            {
                NewEventStatusTextBlock.Text =
                    "The event could not be uploaded: " + exception.Message;
            }
            finally
            {
                SetCalendarBusy(false, null);
            }
        }

        private async Task RefreshNewEventCalendarsAsync()
        {
            SyncMappings mappings = await SyncMappingsStore.LoadAsync();
            string selectedCalendarId = (NewEventCalendarComboBox.SelectedItem as ComboBoxItem)?
                .Tag as string;
            _loadingCalendarForm = true;

            try
            {
                NewEventCalendarComboBox.Items.Clear();

                foreach (KeyValuePair<string, CalendarMapping> mapping in mappings.Calendars
                    .OrderBy(entry => entry.Value.DisplayName))
                {
                    NewEventCalendarComboBox.Items.Add(new ComboBoxItem
                    {
                        Content = "Microsoft Graph - " + mapping.Value.DisplayName,
                        Tag = mapping.Key
                    });
                }

                ComboBoxItem matchingCalendar = NewEventCalendarComboBox.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(item => item.Tag as string == selectedCalendarId);
                NewEventCalendarComboBox.SelectedItem = matchingCalendar
                    ?? NewEventCalendarComboBox.Items.OfType<ComboBoxItem>().FirstOrDefault();
            }
            finally
            {
                _loadingCalendarForm = false;
            }

            bool hasCalendars = NewEventCalendarComboBox.Items.Count > 0;
            bool canCreate = hasCalendars
                && _isProActive
                && SyncSettings.Load().TwoWaySyncEnabled
                && _hasWriteAccess;
            CreateEventButton.IsEnabled = canCreate;

            if (!hasCalendars)
            {
                NewEventStatusTextBlock.Text =
                    "Run a full sync to add a LiveSync calendar before creating events.";
            }
            else if (!canCreate)
            {
                NewEventStatusTextBlock.Text =
                    "Enable LiveSync Pro two-way synchronization to create and upload events.";
            }
            else if (!_loadingCalendarForm)
            {
                NewEventStatusTextBlock.Text =
                    "New events are uploaded to Microsoft Graph immediately.";
            }
        }

        private void InitializeNewEventForm()
        {
            DateTimeOffset now = DateTimeOffset.Now;
            DateTimeOffset today = new DateTimeOffset(
                now.Year,
                now.Month,
                now.Day,
                0,
                0,
                0,
                now.Offset);
            TimeSpan startTime = new TimeSpan(now.Hour, now.Minute, 0);
            TimeSpan endTime = startTime.Add(TimeSpan.FromHours(1));

            if (endTime >= TimeSpan.FromDays(1))
            {
                endTime = TimeSpan.FromHours(23) + TimeSpan.FromMinutes(59);
            }

            NewEventStartDatePicker.Date = today;
            NewEventEndDatePicker.Date = today;
            NewEventStartTimePicker.Time = startTime;
            NewEventEndTimePicker.Time = endTime;
        }

        private void NewEventAllDayToggleSwitch_Toggled(object sender, RoutedEventArgs e)
        {
            if (_loadingCalendarForm)
            {
                return;
            }

            bool isAllDay = NewEventAllDayToggleSwitch.IsOn;
            NewEventStartTimePicker.IsEnabled = !isAllDay;
            NewEventEndTimePicker.IsEnabled = !isAllDay;
        }

        private static DateTimeOffset CreateNewEventDateTime(
            DatePicker datePicker,
            TimeSpan time)
        {
            DateTimeOffset date = datePicker.Date;
            DateTime localDateTime = new DateTime(
                date.Year,
                date.Month,
                date.Day,
                time.Hours,
                time.Minutes,
                0,
                DateTimeKind.Unspecified);
            return new DateTimeOffset(
                localDateTime,
                TimeZoneInfo.Local.GetUtcOffset(localDateTime));
        }

        private void SetCalendarBusy(bool isBusy, string status)
        {
            NewEventCalendarComboBox.IsEnabled = !isBusy;
            NewEventSubjectTextBox.IsEnabled = !isBusy;
            NewEventAllDayToggleSwitch.IsEnabled = !isBusy;
            NewEventStartDatePicker.IsEnabled = !isBusy;
            NewEventStartTimePicker.IsEnabled = !isBusy && !NewEventAllDayToggleSwitch.IsOn;
            NewEventEndDatePicker.IsEnabled = !isBusy;
            NewEventEndTimePicker.IsEnabled = !isBusy && !NewEventAllDayToggleSwitch.IsOn;
            NewEventLocationTextBox.IsEnabled = !isBusy;
            NewEventDetailsTextBox.IsEnabled = !isBusy;
            CreateEventButton.IsEnabled = !isBusy
                && _isProActive
                && NewEventCalendarComboBox.Items.Count > 0
                && SyncSettings.Load().TwoWaySyncEnabled
                && _hasWriteAccess;

            if (!string.IsNullOrEmpty(status))
            {
                NewEventStatusTextBlock.Text = status;
            }
        }

        private async Task RefreshProStatusAsync()
        {
            try
            {
                LiveSyncProEntitlement entitlement =
                    await LiveSyncProService.GetEntitlementAsync();
                _isProActive = entitlement.IsActive;
                CalendarPivotItem.IsEnabled = _isProActive;

                if (!entitlement.IsActive)
                {
                    ProStatusTextBlock.Text = string.IsNullOrEmpty(entitlement.StatusMessage)
                        ? "LiveSync Pro is not active. Subscribe to unlock Pro features."
                        : entitlement.StatusMessage;
                    ProUpgradeButton.Visibility = Visibility.Visible;
                    TwoWaySyncToggleSwitch.IsEnabled = false;
                    return;
                }

                if (entitlement.ExpirationDate == DateTimeOffset.MaxValue)
                {
                    ProStatusTextBlock.Text = "LiveSync Pro is active.";
                }
                else
                {
                    ProStatusTextBlock.Text = string.Format(
                        "LiveSync Pro is active through {0:d}.",
                        entitlement.ExpirationDate.ToLocalTime());
                }

                ProUpgradeButton.Visibility = Visibility.Collapsed;
                TwoWaySyncToggleSwitch.IsEnabled = true;
            }
            catch (Exception exception)
            {
                _isProActive = false;
                CalendarPivotItem.IsEnabled = false;
                ProStatusTextBlock.Text =
                    "LiveSync Pro status is unavailable: " + exception.Message;
                ProUpgradeButton.Visibility = Visibility.Visible;
                TwoWaySyncToggleSwitch.IsEnabled = false;
            }
        }

        private void SetProBusy(bool isBusy)
        {
            ProUpgradeButton.IsEnabled = !isBusy;
            ProRefreshButton.IsEnabled = !isBusy;
        }

        private void UpdateSyncProgress(SyncProgress progress)
        {
            StatusTextBlock.Text = progress.Status;
            SyncProgressTextBlock.Text = progress.CounterText;
            SyncProgressBar.IsIndeterminate = progress.IsIndeterminate;

            if (!progress.IsIndeterminate)
            {
                SyncProgressBar.Maximum = Math.Max(progress.Total, 1);
                SyncProgressBar.Value = Math.Min(progress.Current, SyncProgressBar.Maximum);
            }
        }

        private void LoadSyncSettings()
        {
            SyncSettings syncSettings = SyncSettings.Load();
            _loadingSettings = true;

            try
            {
                BackgroundSyncToggleSwitch.IsOn = syncSettings.BackgroundSyncEnabled;
                FailureNotificationToggleSwitch.IsOn = syncSettings.FailureNotificationsEnabled;
                TwoWaySyncToggleSwitch.IsOn = syncSettings.TwoWaySyncEnabled;
                CalendarNamesTextBox.Text = syncSettings.CalendarNames;
                ComboBoxItem matchingWindow = EventWindowComboBox.Items
                    .OfType<ComboBoxItem>()
                    .FirstOrDefault(item => item.Tag as string == syncSettings.EventWindowMonths.ToString());

                if (matchingWindow != null)
                {
                    EventWindowComboBox.SelectedItem = matchingWindow;
                }
            }
            finally
            {
                _loadingSettings = false;
            }
        }

        private SyncSettings SaveSyncSettings()
        {
            ComboBoxItem selectedWindow = EventWindowComboBox.SelectedItem as ComboBoxItem;
            int eventWindowMonths;

            if (selectedWindow == null || !int.TryParse(selectedWindow.Tag as string, out eventWindowMonths))
            {
                throw new InvalidOperationException("Select a valid event window.");
            }

            SyncSettings syncSettings = new SyncSettings
            {
                BackgroundSyncEnabled = BackgroundSyncToggleSwitch.IsOn,
                FailureNotificationsEnabled = FailureNotificationToggleSwitch.IsOn,
                TwoWaySyncEnabled = TwoWaySyncToggleSwitch.IsOn,
                CalendarNames = CalendarNamesTextBox.Text.Trim(),
                EventWindowMonths = eventWindowMonths
            };
            syncSettings.Save();
            return syncSettings;
        }

        private void SetTwoWaySyncToggle(bool isOn)
        {
            _loadingSettings = true;

            try
            {
                TwoWaySyncToggleSwitch.IsOn = isOn;
            }
            finally
            {
                _loadingSettings = false;
            }
        }

        private async System.Threading.Tasks.Task ConfirmDuplicateCalendarCleanupAsync(
            SyncCleanupPlan cleanupPlan)
        {
            if (cleanupPlan == null || cleanupPlan.CalendarIds.Count == 0)
            {
                return;
            }

            ContentDialog dialog = new ContentDialog
            {
                Title = "Remove duplicate LiveSync calendars?",
                Content = string.Format(
                    "LiveSync found {0} duplicate calendar(s) created by an earlier version. Remove those duplicate calendar containers and their copied events?",
                    cleanupPlan.CalendarIds.Distinct().Count()),
                PrimaryButtonText = "Remove duplicate calendars",
                CloseButtonText = "Keep for now"
            };

            if (await dialog.ShowAsync() == ContentDialogResult.Primary)
            {
                await GraphSyncService.CleanupDuplicatesAsync(cleanupPlan);
                BackgroundSyncDiagnostics.Record("Removed duplicate LiveSync calendars.");
                StatusTextBlock.Text = "Sync complete and duplicate LiveSync calendars were removed.";
            }
        }

        private void RefreshSyncHistory()
        {
            SyncHistoryTextBlock.Text = string.Join(Environment.NewLine, BackgroundSyncDiagnostics.GetHistory());
        }

        private void MainPivot_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MainPivot.SelectedItem == ActivityPivotItem)
            {
                RefreshSyncHistory();
                _activityRefreshTimer.Start();
            }
            else
            {
                _activityRefreshTimer.Stop();
            }
        }

        private void ActivityRefreshTimer_Tick(object sender, object e)
        {
            RefreshSyncHistory();
        }

        private void MainPage_Unloaded(object sender, RoutedEventArgs e)
        {
            _activityRefreshTimer.Stop();
        }

        private void RefreshBackgroundTaskStatus()
        {
            SyncSettings syncSettings = SyncSettings.Load();

            if (!syncSettings.BackgroundSyncEnabled)
            {
                BackgroundTaskStatusTextBlock.Text = "Background sync is disabled.";
            }
            else if (BackgroundSyncRegistration.IsRegistered)
            {
                BackgroundTaskStatusTextBlock.Text =
                    "Background sync is registered. Windows checks for changes at least 15 minutes apart.";
            }
            else
            {
                BackgroundTaskStatusTextBlock.Text =
                    "Background sync is not registered. Complete a manual sync to register it.";
            }
        }

        private static string GetBuildVersionText()
        {
            PackageVersion version = Package.Current.Id.Version;
            return string.Format(
                "Version {0}.{1}.{2}.{3}",
                version.Major,
                version.Minor,
                version.Build,
                version.Revision);
        }
    }
}
