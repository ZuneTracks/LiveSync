using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.Appointments;
using Windows.ApplicationModel.Contacts;
using Windows.Data.Json;
using Windows.Security.Cryptography;
using Windows.Security.Cryptography.Core;
using Windows.Storage;
using Windows.Storage.Streams;

namespace GraphMobileSync.Services
{
    internal enum ContactPhotoResult
    {
        NoPhoto,
        Downloaded,
        Reused
    }

    internal sealed class SyncSummary
    {
        public int Contacts { get; set; }

        public int ContactsUnchanged { get; set; }

        public int ContactPhotosDownloaded { get; set; }

        public int ContactPhotosReused { get; set; }

        public int ContactsWithoutPhotos { get; set; }

        public int DuplicateContactsRemoved { get; set; }

        public int Calendars { get; set; }

        public int Events { get; set; }

        public int EventsUnchanged { get; set; }

        public SyncCleanupPlan CleanupPlan { get; set; }
    }

    internal sealed class CalendarDeltaBatch
    {
        public CalendarDeltaBatch(string graphId, string displayName, IList<JsonObject> items)
        {
            GraphId = graphId;
            DisplayName = displayName;
            Items = items;
        }

        public string GraphId { get; private set; }

        public string DisplayName { get; private set; }

        public IList<JsonObject> Items { get; private set; }
    }

    internal sealed class SyncProgress
    {
        public SyncProgress(string status, string counterText, int current, int total, bool isIndeterminate)
        {
            Status = status;
            CounterText = counterText;
            Current = current;
            Total = total;
            IsIndeterminate = isIndeterminate;
        }

        public string Status { get; private set; }

        public string CounterText { get; private set; }

        public int Current { get; private set; }

        public int Total { get; private set; }

        public bool IsIndeterminate { get; private set; }
    }

    internal sealed partial class GraphSyncService
    {
        private const string ContactListName = "Microsoft Graph";

        private const string CalendarPrefix = "Microsoft Graph - ";

        private const string ContactPhotoFolderName = "ContactPhotos";

        private readonly GraphClient _graphClient;

        public GraphSyncService(GraphClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<SyncSummary> SyncAsync(IProgress<SyncProgress> progress)
        {
            SyncSettings settings = SyncSettings.Load();
            ReportDownloadProgress(progress, "Downloading contacts from Microsoft Graph...", "Downloading contacts...");
            IList<JsonObject> graphContacts = await _graphClient.GetCollectionAsync(
                "/v1.0/me/contacts?$select=id,changeKey,givenName,surname,displayName,emailAddresses,mobilePhone,businessPhones&$top=500");
            IList<JsonObject> graphCalendars = await _graphClient.GetCollectionAsync(
                "/v1.0/me/calendars?$select=id,name&$top=50");
            List<CalendarSyncData> calendarData = await GetCalendarDataAsync(graphCalendars, settings, progress);
            int totalItems = graphContacts.Count + calendarData.Sum(data => data.Events.Count);
            int completedItems = 0;
            SyncMappings mappings = await SyncMappingsStore.LoadAsync();
            SyncCleanupPlan cleanupPlan = new SyncCleanupPlan();
            SyncSummary summary = new SyncSummary { CleanupPlan = cleanupPlan };

            ReportSyncProgress(progress, "Preparing contacts on this phone...", completedItems, totalItems);
            ContactStore contactStore = await ContactManager.RequestStoreAsync(ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException("Contact access was not granted. Enable contact access for LiveSync and try again.");
            }

            ContactList contactList = await GetContactListAsync(contactStore, mappings, cleanupPlan);
            IReadOnlyList<Contact> existingContacts = await GetContactsInListAsync(contactList);
            completedItems = await SyncContactsAsync(
                contactList,
                existingContacts,
                graphContacts,
                mappings,
                summary,
                progress,
                totalItems,
                completedItems);

            ReportSyncProgress(progress, "Preparing calendars on this phone...", completedItems, totalItems);
            AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                AppointmentStoreAccessType.AppCalendarsReadWrite);

            if (appointmentStore == null)
            {
                throw new InvalidOperationException("Calendar access was not granted. Enable calendar access for LiveSync and try again.");
            }

            HashSet<string> syncedCalendarIds = new HashSet<string>();
            List<AppointmentCalendar> localCalendars =
                (await appointmentStore.FindAppointmentCalendarsAsync()).ToList();

            foreach (CalendarSyncData sourceCalendar in calendarData)
            {
                syncedCalendarIds.Add(sourceCalendar.GraphId);
                AppointmentCalendar calendar = await GetCalendarAsync(
                    appointmentStore,
                    localCalendars,
                    sourceCalendar.GraphId,
                    sourceCalendar.DisplayName,
                    mappings,
                    cleanupPlan);
                CalendarMapping calendarMapping = mappings.Calendars[sourceCalendar.GraphId];
                IReadOnlyList<Appointment> existingAppointments = await calendar.FindAppointmentsAsync(
                    settings.GetStartDate(),
                    settings.GetEndDate() - settings.GetStartDate());
                completedItems = await SyncAppointmentsAsync(
                    calendar,
                    existingAppointments,
                    sourceCalendar,
                    calendarMapping,
                    summary,
                    progress,
                    totalItems,
                    completedItems);
                summary.Calendars++;
            }

            foreach (KeyValuePair<string, CalendarMapping> calendar in mappings.Calendars.ToList())
            {
                if (!syncedCalendarIds.Contains(calendar.Key) && !string.IsNullOrEmpty(calendar.Value.LocalId))
                {
                    cleanupPlan.CalendarIds.Add(calendar.Value.LocalId);
                    mappings.Calendars.Remove(calendar.Key);
                }
            }

            await SyncMappingsStore.SaveAsync(mappings);
            ReportSyncProgress(progress, "Finishing sync...", completedItems, totalItems);
            return summary;
        }

        // The caller owns loading and saving the mappings so that one delta run, which may also push
        // local changes, reads and writes the mappings file only once.
        public async Task<SyncSummary> ApplyDeltaAsync(
            IList<JsonObject> contactChanges,
            IList<CalendarDeltaBatch> calendarChanges,
            ISet<string> activeCalendarIds,
            SyncSettings settings,
            SyncMappings mappings)
        {
            SyncSummary summary = new SyncSummary
            {
                CleanupPlan = new SyncCleanupPlan()
            };

            if (contactChanges != null && contactChanges.Count > 0)
            {
                await ApplyContactDeltaAsync(contactChanges, mappings, summary);
            }

            bool hasCalendarMappingChanges = mappings.Calendars.Keys.Any(
                calendarId => !activeCalendarIds.Contains(calendarId));

            if ((calendarChanges != null && calendarChanges.Count > 0) || hasCalendarMappingChanges)
            {
                AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                    AppointmentStoreAccessType.AppCalendarsReadWrite);

                if (appointmentStore == null)
                {
                    throw new InvalidOperationException(
                        "Calendar access was not granted. Enable calendar access for LiveSync and try again.");
                }

                if (calendarChanges != null)
                {
                    List<AppointmentCalendar> localCalendars =
                        (await appointmentStore.FindAppointmentCalendarsAsync()).ToList();

                    foreach (CalendarDeltaBatch calendarChange in calendarChanges)
                    {
                        AppointmentCalendar calendar = await GetCalendarAsync(
                            appointmentStore,
                            localCalendars,
                            calendarChange.GraphId,
                            calendarChange.DisplayName,
                            mappings,
                            summary.CleanupPlan);
                        CalendarMapping mapping = mappings.Calendars[calendarChange.GraphId];
                        IReadOnlyList<Appointment> existingAppointments =
                            await calendar.FindAppointmentsAsync(
                                settings.GetStartDate(),
                                settings.GetEndDate() - settings.GetStartDate());
                        await ApplyAppointmentDeltaAsync(
                            calendar,
                            existingAppointments,
                            calendarChange.Items,
                            mapping,
                            summary);
                    }
                }
            }

            foreach (string calendarId in mappings.Calendars.Keys.ToList())
            {
                if (!activeCalendarIds.Contains(calendarId))
                {
                    mappings.Calendars.Remove(calendarId);
                }
            }

            return summary;
        }

        private static bool IsGraphItemRemoved(JsonObject graphItem)
        {
            return JsonValueReader.GetObject(graphItem, "@removed") != null;
        }

        private static void ReportDownloadProgress(
            IProgress<SyncProgress> progress,
            string status,
            string counterText)
        {
            if (progress != null)
            {
                progress.Report(new SyncProgress(status, counterText, 0, 0, true));
            }
        }

        private static void ReportSyncProgress(
            IProgress<SyncProgress> progress,
            string status,
            int current,
            int total)
        {
            if (progress != null)
            {
                progress.Report(new SyncProgress(
                    status,
                    string.Format("{0} of {1} items synced", current, total),
                    current,
                    total,
                    false));
            }
        }
    }
}
