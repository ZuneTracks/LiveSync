using System;
using System.Linq;
using GraphMobileSync.Services;
using Windows.ApplicationModel;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace GraphMobileSync
{
    public sealed partial class MainPage : Page
    {
        private string _accessToken;

        public MainPage()
        {
            InitializeComponent();
            BuildVersionTextBlock.Text = GetBuildVersionText();
            LoadSettings();
            Loaded += MainPage_Loaded;
        }

        private void LoadSettings()
        {
            LoadSyncSettings();
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

                _accessToken = await OAuthService.SignInAsync(clientId, tenant);
                SyncButton.IsEnabled = true;
                StatusTextBlock.Text = "Signed in. You can now sync contacts and calendars.";
            }
            catch (Exception exception)
            {
                StatusTextBlock.Text = exception.Message;
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        private async void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= MainPage_Loaded;

            if (!OAuthService.HasSavedSession())
            {
                return;
            }

            SetBusy(true, "Restoring saved sign-in...");

            try
            {
                _accessToken = await OAuthService.GetBackgroundAccessTokenAsync();
                StatusTextBlock.Text = "Signed in with your saved session. You can sync contacts and calendars.";

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
                    StatusTextBlock.Text =
                        "Your saved session could not be restored. Check your connection, then sign in again if needed.";
                }
            }
            finally
            {
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
                    GraphSyncService syncService = new GraphSyncService(graphClient);
                    summary = await syncService.SyncAsync(syncProgress);
                    StatusTextBlock.Text = "Establishing Microsoft Graph change tracking...";
                    await DeltaCheckpointService.CreateAsync(graphClient);
                }

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
                    result = await new DeltaSyncService(graphClient).PollAndApplyAsync(true);
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

        private void SetBusy(bool isBusy, string status)
        {
            SignInButton.IsEnabled = !isBusy;
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
            BackgroundSyncToggleSwitch.IsOn = syncSettings.BackgroundSyncEnabled;
            FailureNotificationToggleSwitch.IsOn = syncSettings.FailureNotificationsEnabled;
            CalendarNamesTextBox.Text = syncSettings.CalendarNames;
            ComboBoxItem matchingWindow = EventWindowComboBox.Items
                .OfType<ComboBoxItem>()
                .FirstOrDefault(item => item.Tag as string == syncSettings.EventWindowMonths.ToString());

            if (matchingWindow != null)
            {
                EventWindowComboBox.SelectedItem = matchingWindow;
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
                CalendarNames = CalendarNamesTextBox.Text.Trim(),
                EventWindowMonths = eventWindowMonths
            };
            syncSettings.Save();
            return syncSettings;
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
                    "LiveSync found {0} duplicate calendar(s) created by an earlier version. Clear only the data inside those duplicate LiveSync calendars?",
                    cleanupPlan.CalendarIds.Count),
                PrimaryButtonText = "Clear duplicate calendars",
                CloseButtonText = "Keep for now"
            };

            if (await dialog.ShowAsync() == ContentDialogResult.Primary)
            {
                await GraphSyncService.CleanupDuplicatesAsync(cleanupPlan);
                BackgroundSyncDiagnostics.Record("Cleared duplicate LiveSync calendars.");
                StatusTextBlock.Text = "Sync complete and duplicate LiveSync calendars were cleared.";
            }
        }

        private void RefreshSyncHistory()
        {
            SyncHistoryTextBlock.Text = string.Join(Environment.NewLine, BackgroundSyncDiagnostics.GetHistory());
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
