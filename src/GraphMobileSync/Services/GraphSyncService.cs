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
        Stored,
        NotStored
    }

    internal sealed class SyncSummary
    {
        public int Contacts { get; set; }

        public int ContactPhotosDownloaded { get; set; }

        public int ContactPhotosStored { get; set; }

        public int ContactsWithoutPhotos { get; set; }

        public int DuplicateContactsRemoved { get; set; }

        public int Calendars { get; set; }

        public int Events { get; set; }

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

    internal sealed class GraphSyncService
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
                "/v1.0/me/contacts?$select=id,givenName,surname,displayName,emailAddresses,mobilePhone,businessPhones&$top=50");
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

            foreach (CalendarSyncData sourceCalendar in calendarData)
            {
                syncedCalendarIds.Add(sourceCalendar.GraphId);
                AppointmentCalendar calendar = await GetCalendarAsync(
                    appointmentStore,
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

        public async Task<SyncSummary> ApplyDeltaAsync(
            IList<JsonObject> contactChanges,
            IList<CalendarDeltaBatch> calendarChanges,
            ISet<string> activeCalendarIds,
            SyncSettings settings)
        {
            SyncMappings mappings = await SyncMappingsStore.LoadAsync();
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
                    foreach (CalendarDeltaBatch calendarChange in calendarChanges)
                    {
                        AppointmentCalendar calendar = await GetCalendarAsync(
                            appointmentStore,
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

            await SyncMappingsStore.SaveAsync(mappings);
            return summary;
        }

        public static async Task CleanupDuplicatesAsync(SyncCleanupPlan cleanupPlan)
        {
            if (cleanupPlan == null || !cleanupPlan.HasItems)
            {
                return;
            }

            await CleanupDuplicateContactListsAsync(cleanupPlan.ContactListIds);

            if (cleanupPlan.CalendarIds.Count == 0)
            {
                return;
            }

            AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                AppointmentStoreAccessType.AppCalendarsReadWrite);

            if (appointmentStore == null)
            {
                throw new InvalidOperationException("Calendar access was not granted. Duplicate calendars were not removed.");
            }

            IReadOnlyList<AppointmentCalendar> calendars = await appointmentStore.FindAppointmentCalendarsAsync();

            foreach (string calendarId in cleanupPlan.CalendarIds.Distinct())
            {
                AppointmentCalendar duplicateCalendar = calendars.FirstOrDefault(calendar => calendar.LocalId == calendarId);

                if (duplicateCalendar == null)
                {
                    continue;
                }

                IReadOnlyList<Appointment> duplicateAppointments = await duplicateCalendar.FindAppointmentsAsync(
                    DateTimeOffset.Now.AddYears(-20),
                    TimeSpan.FromDays(14600));

                foreach (Appointment appointment in duplicateAppointments)
                {
                    await duplicateCalendar.DeleteAppointmentAsync(appointment.LocalId);
                }
            }
        }

        public static async Task CleanupDuplicateContactListsAsync(
            IEnumerable<string> duplicateContactListIds)
        {
            if (duplicateContactListIds == null)
            {
                return;
            }

            ContactStore contactStore = await ContactManager.RequestStoreAsync(ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException("Contact access was not granted. Duplicate contacts were not removed.");
            }

            IReadOnlyList<ContactList> contactLists = await contactStore.FindContactListsAsync();

            foreach (string contactListId in duplicateContactListIds.Distinct())
            {
                ContactList duplicateList = contactLists.FirstOrDefault(list => list.Id == contactListId);

                if (duplicateList == null)
                {
                    continue;
                }

                IReadOnlyList<Contact> duplicateContacts = await GetContactsInListAsync(duplicateList);

                foreach (Contact contact in duplicateContacts)
                {
                    await duplicateList.DeleteContactAsync(contact);
                }

                await RemoveEmptyContactListAsync(duplicateList);
            }
        }

        private static async Task RemoveEmptyContactListAsync(ContactList contactList)
        {
            try
            {
                await contactList.DeleteAsync();
                return;
            }
            catch (Exception)
            {
                // Windows 10 Mobile 15063 does not let an app delete its own contact list container.
            }

            try
            {
                // Hiding the emptied container keeps it out of the People app, including the
                // "Link contacts" picker, which still lists app-owned lists that remain visible.
                contactList.IsHidden = true;
                await contactList.SaveAsync();
            }
            catch (Exception)
            {
            }
        }

        private async Task<int> SyncContactsAsync(
            ContactList contactList,
            IReadOnlyList<Contact> existingContacts,
            IList<JsonObject> graphContacts,
            SyncMappings mappings,
            SyncSummary summary,
            IProgress<SyncProgress> progress,
            int totalItems,
            int completedItems)
        {
            Dictionary<string, Contact> existingById = existingContacts.ToDictionary(contact => contact.Id);
            Dictionary<string, List<Contact>> existingByFingerprint =
                GroupContactsByFingerprint(existingContacts);
            HashSet<string> graphContactIds = new HashSet<string>();
            HashSet<string> claimedContactIds = new HashSet<string>();
            HashSet<string> mappedContactIds = new HashSet<string>(
                mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));

            foreach (JsonObject graphContact in graphContacts)
            {
                string graphContactId = JsonValueReader.GetString(graphContact, "id", string.Empty);

                if (string.IsNullOrEmpty(graphContactId))
                {
                    continue;
                }

                graphContactIds.Add(graphContactId);
                Contact localContact = FindExistingContact(
                    graphContactId,
                    graphContact,
                    mappings,
                    existingById,
                    existingByFingerprint,
                    claimedContactIds,
                    mappedContactIds);

                if (localContact == null)
                {
                    localContact = new Contact();
                }

                ContactPhotoResult photoResult =
                    await SaveContactAsync(contactList, localContact, graphContact);
                RecordContactPhotoResult(summary, photoResult);

                string savedContactId = localContact.Id;

                if (string.IsNullOrEmpty(savedContactId))
                {
                    savedContactId = await FindSavedContactIdAsync(
                        contactList,
                        GetContactFingerprint(graphContact));
                }

                if (string.IsNullOrEmpty(savedContactId))
                {
                    throw new InvalidOperationException("LiveSync could not identify a saved contact.");
                }

                mappings.ContactIds[graphContactId] = savedContactId;
                claimedContactIds.Add(savedContactId);
                summary.Contacts++;
                completedItems++;
                ReportSyncProgress(progress, "Syncing contacts...", completedItems, totalItems);
            }

            foreach (KeyValuePair<string, string> mapping in mappings.ContactIds.ToList())
            {
                if (!graphContactIds.Contains(mapping.Key))
                {
                    Contact removedContact;

                    if (existingById.TryGetValue(mapping.Value, out removedContact))
                    {
                        await contactList.DeleteContactAsync(removedContact);
                    }

                    await DeleteContactPhotoAsync(mapping.Key);
                    mappings.ContactIds.Remove(mapping.Key);
                }
            }

            ISet<string> mappedLocalContactIds = new HashSet<string>(
                mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));

            foreach (Contact existingContact in existingContacts)
            {
                if (!mappedLocalContactIds.Contains(existingContact.Id))
                {
                    await contactList.DeleteContactAsync(existingContact);
                    summary.DuplicateContactsRemoved++;
                }
            }

            return completedItems;
        }

        private async Task ApplyContactDeltaAsync(
            IList<JsonObject> contactChanges,
            SyncMappings mappings,
            SyncSummary summary)
        {
            ContactStore contactStore = await ContactManager.RequestStoreAsync(
                ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException(
                    "Contact access was not granted. Enable contact access for LiveSync and try again.");
            }

            ContactList contactList = await GetContactListAsync(
                contactStore,
                mappings,
                summary.CleanupPlan);
            IReadOnlyList<Contact> existingContacts = await GetContactsInListAsync(contactList);
            Dictionary<string, Contact> existingById = existingContacts.ToDictionary(contact => contact.Id);
            Dictionary<string, List<Contact>> existingByFingerprint =
                GroupContactsByFingerprint(existingContacts);
            HashSet<string> claimedContactIds = new HashSet<string>();
            HashSet<string> mappedContactIds = new HashSet<string>(
                mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));

            foreach (JsonObject graphContact in contactChanges)
            {
                string graphContactId = JsonValueReader.GetString(graphContact, "id", string.Empty);

                if (string.IsNullOrEmpty(graphContactId))
                {
                    continue;
                }

                string localContactId;

                if (IsGraphItemRemoved(graphContact))
                {
                    if (mappings.ContactIds.TryGetValue(graphContactId, out localContactId))
                    {
                        Contact localContact;

                        if (existingById.TryGetValue(localContactId, out localContact))
                        {
                            await contactList.DeleteContactAsync(localContact);
                            existingById.Remove(localContactId);
                        }

                        await DeleteContactPhotoAsync(graphContactId);
                        mappings.ContactIds.Remove(graphContactId);
                    }

                    summary.Contacts++;
                    continue;
                }

                Contact updatedContact = FindExistingContact(
                    graphContactId,
                    graphContact,
                    mappings,
                    existingById,
                    existingByFingerprint,
                    claimedContactIds,
                    mappedContactIds);

                if (updatedContact == null)
                {
                    updatedContact = new Contact();
                }

                ContactPhotoResult photoResult =
                    await SaveContactAsync(contactList, updatedContact, graphContact);
                RecordContactPhotoResult(summary, photoResult);

                string savedContactId = updatedContact.Id;

                if (string.IsNullOrEmpty(savedContactId))
                {
                    savedContactId = await FindSavedContactIdAsync(
                        contactList,
                        GetContactFingerprint(graphContact));
                }

                if (string.IsNullOrEmpty(savedContactId))
                {
                    throw new InvalidOperationException(
                        "LiveSync could not identify a contact updated from Microsoft Graph.");
                }

                mappings.ContactIds[graphContactId] = savedContactId;
                claimedContactIds.Add(savedContactId);
                existingById[savedContactId] = updatedContact;
                summary.Contacts++;
            }
        }

        private static async Task ApplyAppointmentDeltaAsync(
            AppointmentCalendar calendar,
            IReadOnlyList<Appointment> existingAppointments,
            IList<JsonObject> eventChanges,
            CalendarMapping calendarMapping,
            SyncSummary summary)
        {
            Dictionary<string, Appointment> existingById =
                existingAppointments.ToDictionary(appointment => appointment.LocalId);
            Dictionary<string, Appointment> unmappedByFingerprint = existingAppointments
                .GroupBy(GetAppointmentFingerprint)
                .Where(group => !string.IsNullOrEmpty(group.Key) && group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.First());
            ISet<string> appointmentIdsInWindow = new HashSet<string>(existingById.Keys);

            foreach (JsonObject graphEvent in eventChanges)
            {
                string graphEventId = JsonValueReader.GetString(graphEvent, "id", string.Empty);

                if (string.IsNullOrEmpty(graphEventId))
                {
                    continue;
                }

                string localAppointmentId;

                if (IsGraphItemRemoved(graphEvent))
                {
                    if (calendarMapping.EventIds.TryGetValue(graphEventId, out localAppointmentId))
                    {
                        Appointment existingAppointment;

                        if (existingById.TryGetValue(localAppointmentId, out existingAppointment))
                        {
                            await calendar.DeleteAppointmentAsync(localAppointmentId);
                            existingById.Remove(localAppointmentId);
                            appointmentIdsInWindow.Remove(localAppointmentId);
                        }

                        calendarMapping.EventIds.Remove(graphEventId);
                    }

                    summary.Events++;
                    continue;
                }

                Appointment sourceAppointment = CreateAppointment(graphEvent);

                if (sourceAppointment == null)
                {
                    continue;
                }

                Appointment updatedAppointment = null;

                if (calendarMapping.EventIds.TryGetValue(graphEventId, out localAppointmentId))
                {
                    existingById.TryGetValue(localAppointmentId, out updatedAppointment);
                }

                if (updatedAppointment == null)
                {
                    unmappedByFingerprint.TryGetValue(
                        GetAppointmentFingerprint(sourceAppointment),
                        out updatedAppointment);
                }

                if (updatedAppointment == null)
                {
                    updatedAppointment = sourceAppointment;
                }
                else
                {
                    PopulateAppointment(updatedAppointment, sourceAppointment);
                }

                await calendar.SaveAppointmentAsync(updatedAppointment);
                localAppointmentId = updatedAppointment.LocalId;

                if (string.IsNullOrEmpty(localAppointmentId))
                {
                    localAppointmentId = await FindSavedAppointmentIdAsync(calendar, sourceAppointment);
                }

                if (string.IsNullOrEmpty(localAppointmentId))
                {
                    throw new InvalidOperationException(
                        "LiveSync could not identify a calendar event updated from Microsoft Graph.");
                }

                calendarMapping.EventIds[graphEventId] = localAppointmentId;
                appointmentIdsInWindow.Add(localAppointmentId);
                existingById[localAppointmentId] = updatedAppointment;
                summary.Events++;
            }

            foreach (KeyValuePair<string, string> mapping in calendarMapping.EventIds.ToList())
            {
                if (string.IsNullOrEmpty(mapping.Value)
                    || !appointmentIdsInWindow.Contains(mapping.Value))
                {
                    calendarMapping.EventIds.Remove(mapping.Key);
                }
            }
        }

        private static async Task<int> SyncAppointmentsAsync(
            AppointmentCalendar calendar,
            IReadOnlyList<Appointment> existingAppointments,
            CalendarSyncData sourceCalendar,
            CalendarMapping calendarMapping,
            SyncSummary summary,
            IProgress<SyncProgress> progress,
            int totalItems,
            int completedItems)
        {
            Dictionary<string, Appointment> existingById = existingAppointments.ToDictionary(appointment => appointment.LocalId);
            Dictionary<string, Appointment> unmappedByFingerprint = existingAppointments
                .GroupBy(GetAppointmentFingerprint)
                .Where(group => !string.IsNullOrEmpty(group.Key) && group.Count() == 1)
                .ToDictionary(group => group.Key, group => group.First());
            HashSet<string> graphEventIds = new HashSet<string>();

            foreach (EventSyncData sourceEvent in sourceCalendar.Events)
            {
                graphEventIds.Add(sourceEvent.GraphId);
                string localAppointmentId;
                Appointment localAppointment = null;

                if (calendarMapping.EventIds.TryGetValue(sourceEvent.GraphId, out localAppointmentId))
                {
                    existingById.TryGetValue(localAppointmentId, out localAppointment);
                }

                if (localAppointment == null)
                {
                    unmappedByFingerprint.TryGetValue(GetAppointmentFingerprint(sourceEvent.Appointment), out localAppointment);
                }

                if (localAppointment == null)
                {
                    localAppointment = sourceEvent.Appointment;
                }
                else
                {
                    PopulateAppointment(localAppointment, sourceEvent.Appointment);
                }

                await calendar.SaveAppointmentAsync(localAppointment);
                localAppointmentId = localAppointment.LocalId;

                if (string.IsNullOrEmpty(localAppointmentId))
                {
                    localAppointmentId = await FindSavedAppointmentIdAsync(calendar, sourceEvent.Appointment);
                }

                if (string.IsNullOrEmpty(localAppointmentId))
                {
                    throw new InvalidOperationException("LiveSync could not identify a saved calendar event.");
                }

                calendarMapping.EventIds[sourceEvent.GraphId] = localAppointmentId;
                summary.Events++;
                completedItems++;
                ReportSyncProgress(progress, "Syncing calendar events...", completedItems, totalItems);
            }

            foreach (KeyValuePair<string, string> mapping in calendarMapping.EventIds.ToList())
            {
                if (!graphEventIds.Contains(mapping.Key))
                {
                    if (!string.IsNullOrEmpty(mapping.Value))
                    {
                        await calendar.DeleteAppointmentAsync(mapping.Value);
                    }

                    calendarMapping.EventIds.Remove(mapping.Key);
                }
            }

            ISet<string> mappedAppointmentIds = new HashSet<string>(
                calendarMapping.EventIds.Values.Where(id => !string.IsNullOrEmpty(id)));

            foreach (Appointment existingAppointment in existingAppointments)
            {
                if (!mappedAppointmentIds.Contains(existingAppointment.LocalId))
                {
                    await calendar.DeleteAppointmentAsync(existingAppointment.LocalId);
                }
            }

            return completedItems;
        }

        private static async Task<ContactList> GetContactListAsync(
            ContactStore contactStore,
            SyncMappings mappings,
            SyncCleanupPlan cleanupPlan)
        {
            IReadOnlyList<ContactList> lists = await contactStore.FindContactListsAsync();

            // The store is opened with AppContactsReadWrite, so every list returned here is already
            // owned by LiveSync. Matching on display name alone stranded app-owned lists saved under
            // any other name: they were never treated as duplicates, so their contacts stayed visible
            // in the People app and kept showing up in its "Link contacts" picker.
            ContactList contactList = lists.FirstOrDefault(list => list.Id == mappings.ContactListId)
                ?? lists.FirstOrDefault(list => list.DisplayName == ContactListName)
                ?? lists.FirstOrDefault()
                ?? await contactStore.CreateContactListAsync(ContactListName);

            if (contactList.DisplayName != ContactListName || contactList.IsHidden)
            {
                // Un-hide as well as rename: a container hidden by an earlier cleanup pass must never
                // stay hidden once it is the canonical list, or synced contacts would be invisible.
                contactList.DisplayName = ContactListName;
                contactList.IsHidden = false;
                await contactList.SaveAsync();
            }

            mappings.ContactListId = contactList.Id;

            foreach (ContactList duplicateList in lists.Where(list => list.Id != contactList.Id))
            {
                cleanupPlan.ContactListIds.Add(duplicateList.Id);
            }

            return contactList;
        }

        private static Dictionary<string, List<Contact>> GroupContactsByFingerprint(
            IEnumerable<Contact> contacts)
        {
            Dictionary<string, List<Contact>> contactsByFingerprint =
                new Dictionary<string, List<Contact>>();

            foreach (Contact contact in contacts)
            {
                string fingerprint = GetContactFingerprint(contact);

                if (string.IsNullOrEmpty(fingerprint))
                {
                    continue;
                }

                List<Contact> contactsWithFingerprint;

                if (!contactsByFingerprint.TryGetValue(fingerprint, out contactsWithFingerprint))
                {
                    contactsWithFingerprint = new List<Contact>();
                    contactsByFingerprint[fingerprint] = contactsWithFingerprint;
                }

                contactsWithFingerprint.Add(contact);
            }

            return contactsByFingerprint;
        }

        private static Contact FindExistingContact(
            string graphContactId,
            JsonObject graphContact,
            SyncMappings mappings,
            IDictionary<string, Contact> existingById,
            IDictionary<string, List<Contact>> existingByFingerprint,
            ISet<string> claimedContactIds,
            ISet<string> mappedContactIds)
        {
            string mappedContactId = string.Empty;
            Contact mappedContact;

            if (mappings.ContactIds.TryGetValue(graphContactId, out mappedContactId)
                && existingById.TryGetValue(mappedContactId, out mappedContact)
                && !claimedContactIds.Contains(mappedContact.Id))
            {
                return mappedContact;
            }

            List<Contact> contactsWithFingerprint;

            if (!existingByFingerprint.TryGetValue(
                    GetContactFingerprint(graphContact),
                    out contactsWithFingerprint))
            {
                return null;
            }

            foreach (Contact contact in contactsWithFingerprint)
            {
                if (claimedContactIds.Contains(contact.Id))
                {
                    continue;
                }

                if (mappedContactIds.Contains(contact.Id)
                    && !string.Equals(contact.Id, mappedContactId, StringComparison.Ordinal))
                {
                    continue;
                }

                return contact;
            }

            return null;
        }

        private static async Task<IReadOnlyList<Contact>> GetContactsInListAsync(ContactList contactList)
        {
            // ContactStore.FindContactsAsync(string) is a full-text search over contact fields, not a
            // lookup by contact list ID. Passing a list ID to it always matched zero contacts, so the
            // sync treated every list as empty: nothing was ever reused and no stale entry was ever
            // deleted. A ContactReader is the only way to enumerate the contacts in one list.
            List<Contact> contacts = new List<Contact>();
            ContactReader reader = contactList.GetContactReader();

            while (true)
            {
                ContactBatch batch = await reader.ReadBatchAsync();

                if (batch.Status != ContactBatchStatus.Success || batch.Contacts.Count == 0)
                {
                    break;
                }

                contacts.AddRange(batch.Contacts);
            }

            return contacts;
        }

        private static async Task<string> FindSavedContactIdAsync(
            ContactList contactList,
            string fingerprint)
        {
            IReadOnlyList<Contact> savedContacts = await GetContactsInListAsync(contactList);
            List<Contact> matchingContacts = savedContacts
                .Where(contact => GetContactFingerprint(contact) == fingerprint)
                .ToList();

            return matchingContacts.Count == 1 ? matchingContacts[0].Id : string.Empty;
        }

        private static async Task<string> FindSavedAppointmentIdAsync(
            AppointmentCalendar calendar,
            Appointment appointment)
        {
            IReadOnlyList<Appointment> savedAppointments = await calendar.FindAppointmentsAsync(
                appointment.StartTime,
                appointment.Duration);
            string fingerprint = GetAppointmentFingerprint(appointment);
            List<Appointment> matchingAppointments = savedAppointments
                .Where(candidate => GetAppointmentFingerprint(candidate) == fingerprint)
                .ToList();

            return matchingAppointments.Count == 1 ? matchingAppointments[0].LocalId : string.Empty;
        }

        private static async Task<AppointmentCalendar> GetCalendarAsync(
            AppointmentStore appointmentStore,
            string graphCalendarId,
            string graphCalendarDisplayName,
            SyncMappings mappings,
            SyncCleanupPlan cleanupPlan)
        {
            IReadOnlyList<AppointmentCalendar> calendars = await appointmentStore.FindAppointmentCalendarsAsync();
            CalendarMapping mapping;

            if (!mappings.Calendars.TryGetValue(graphCalendarId, out mapping))
            {
                mapping = new CalendarMapping();
                mappings.Calendars[graphCalendarId] = mapping;
            }

            string displayName = CalendarPrefix + graphCalendarDisplayName;
            List<AppointmentCalendar> matchingCalendars = calendars
                .Where(candidate => candidate.DisplayName == displayName)
                .ToList();
            AppointmentCalendar calendar = matchingCalendars.FirstOrDefault(item => item.LocalId == mapping.LocalId)
                ?? matchingCalendars.FirstOrDefault()
                ?? await appointmentStore.CreateAppointmentCalendarAsync(displayName);
            mapping.LocalId = calendar.LocalId;
            mapping.DisplayName = graphCalendarDisplayName;

            foreach (AppointmentCalendar duplicateCalendar in matchingCalendars.Where(item => item.LocalId != calendar.LocalId))
            {
                cleanupPlan.CalendarIds.Add(duplicateCalendar.LocalId);
            }

            return calendar;
        }

        private static bool IsGraphItemRemoved(JsonObject graphItem)
        {
            return JsonValueReader.GetObject(graphItem, "@removed") != null;
        }

        private async Task<List<CalendarSyncData>> GetCalendarDataAsync(
            IList<JsonObject> graphCalendars,
            SyncSettings settings,
            IProgress<SyncProgress> progress)
        {
            List<CalendarSyncData> calendars = new List<CalendarSyncData>();

            for (int index = 0; index < graphCalendars.Count; index++)
            {
                JsonObject graphCalendar = graphCalendars[index];
                string calendarId = JsonValueReader.GetString(graphCalendar, "id", string.Empty);
                string calendarName = JsonValueReader.GetString(graphCalendar, "name", "Calendar");

                if (string.IsNullOrEmpty(calendarId) || !settings.IncludesCalendar(calendarName))
                {
                    continue;
                }

                ReportDownloadProgress(
                    progress,
                    string.Format("Downloading calendar {0} of {1}...", index + 1, graphCalendars.Count),
                    string.Format("Downloading calendar {0} of {1}", index + 1, graphCalendars.Count));
                IList<JsonObject> graphEvents = await _graphClient.GetCollectionAsync(
                    GetCalendarViewPath(calendarId, settings));
                List<EventSyncData> events = new List<EventSyncData>();

                foreach (JsonObject graphEvent in graphEvents)
                {
                    string eventId = JsonValueReader.GetString(graphEvent, "id", string.Empty);
                    Appointment appointment = CreateAppointment(graphEvent);

                    if (!string.IsNullOrEmpty(eventId) && appointment != null)
                    {
                        events.Add(new EventSyncData(eventId, appointment));
                    }
                }

                calendars.Add(new CalendarSyncData(calendarId, calendarName, events));
            }

            return calendars;
        }

        private static string GetCalendarViewPath(string calendarId, SyncSettings settings)
        {
            return string.Format(
                "/v1.0/me/calendars/{0}/calendarView?startDateTime={1}&endDateTime={2}&$select=id,subject,bodyPreview,start,end,location,isAllDay&$top=50",
                Uri.EscapeDataString(calendarId),
                Uri.EscapeDataString(settings.GetStartDate().ToString("o")),
                Uri.EscapeDataString(settings.GetEndDate().ToString("o")));
        }

        private static void PopulateContact(Contact contact, JsonObject graphContact)
        {
            contact.FirstName = JsonValueReader.GetString(graphContact, "givenName", string.Empty);
            contact.LastName = JsonValueReader.GetString(graphContact, "surname", string.Empty);
            contact.Emails.Clear();
            contact.Phones.Clear();
            JsonArray emailAddresses = JsonValueReader.GetArray(graphContact, "emailAddresses");

            foreach (IJsonValue emailAddress in emailAddresses)
            {
                if (emailAddress.ValueType != JsonValueType.Object)
                {
                    continue;
                }

                string address = JsonValueReader.GetString(emailAddress.GetObject(), "address", string.Empty);

                if (!string.IsNullOrWhiteSpace(address))
                {
                    contact.Emails.Add(new ContactEmail { Address = address, Kind = ContactEmailKind.Personal });
                }
            }

            AddPhoneNumber(
                contact,
                JsonValueReader.GetString(graphContact, "mobilePhone", string.Empty),
                ContactPhoneKind.Mobile);
            JsonArray businessPhones = JsonValueReader.GetArray(graphContact, "businessPhones");

            foreach (IJsonValue businessPhone in businessPhones)
            {
                if (businessPhone.ValueType == JsonValueType.String)
                {
                    AddPhoneNumber(contact, businessPhone.GetString(), ContactPhoneKind.Work);
                }
            }
        }

        private async Task<ContactPhotoResult> SaveContactAsync(
            ContactList contactList,
            Contact contact,
            JsonObject graphContact)
        {
            string graphContactId = JsonValueReader.GetString(graphContact, "id", string.Empty);

            if (string.IsNullOrEmpty(graphContactId))
            {
                throw new InvalidOperationException(
                    "LiveSync could not save a contact without a Microsoft Graph ID.");
            }

            PopulateContact(contact, graphContact);
            byte[] photoBytes = await _graphClient.GetContactPhotoAsync(graphContactId);

            if (photoBytes == null)
            {
                await contactList.SaveContactAsync(contact);
                return ContactPhotoResult.NoPhoto;
            }

            if (photoBytes.Length == 0)
            {
                throw new InvalidOperationException(
                    "Microsoft Graph returned an empty contact photo.");
            }

            StorageFile photoFile = await SaveContactPhotoAsync(graphContactId, photoBytes);
            IRandomAccessStreamReference photoReference =
                RandomAccessStreamReference.CreateFromFile(photoFile);

            // Windows generates SmallDisplayPicture, LargeDisplayPicture and Thumbnail from
            // SourceDisplayPicture, and the People app renders those generated images. Setting
            // Thumbnail on its own leaves the display pictures empty, so the photo never appears.
            contact.SourceDisplayPicture = photoReference;
            contact.Thumbnail = photoReference;
            await contactList.SaveContactAsync(contact);

            return await HasStoredDisplayPictureAsync(contactList, contact)
                ? ContactPhotoResult.Stored
                : ContactPhotoResult.NotStored;
        }

        private static void RecordContactPhotoResult(SyncSummary summary, ContactPhotoResult photoResult)
        {
            if (photoResult == ContactPhotoResult.NoPhoto)
            {
                summary.ContactsWithoutPhotos++;
                return;
            }

            summary.ContactPhotosDownloaded++;

            if (photoResult == ContactPhotoResult.Stored)
            {
                summary.ContactPhotosStored++;
            }
        }

        private static async Task<bool> HasStoredDisplayPictureAsync(
            ContactList contactList,
            Contact contact)
        {
            if (string.IsNullOrEmpty(contact.Id))
            {
                return false;
            }

            try
            {
                Contact savedContact = await contactList.GetContactAsync(contact.Id);
                return savedContact != null
                    && (savedContact.SourceDisplayPicture != null
                        || savedContact.SmallDisplayPicture != null
                        || savedContact.Thumbnail != null);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static async Task<StorageFile> SaveContactPhotoAsync(
            string graphContactId,
            byte[] photoBytes)
        {
            StorageFolder photoFolder = await ApplicationData.Current.LocalFolder.CreateFolderAsync(
                ContactPhotoFolderName,
                CreationCollisionOption.OpenIfExists);
            StorageFile photoFile = await photoFolder.CreateFileAsync(
                GetContactPhotoFileName(graphContactId),
                CreationCollisionOption.ReplaceExisting);
            await FileIO.WriteBytesAsync(photoFile, photoBytes);
            return photoFile;
        }

        private static async Task DeleteContactPhotoAsync(string graphContactId)
        {
            try
            {
                StorageFolder photoFolder = await ApplicationData.Current.LocalFolder.GetFolderAsync(
                    ContactPhotoFolderName);
                StorageFile photoFile = await photoFolder.GetFileAsync(
                    GetContactPhotoFileName(graphContactId));
                await photoFile.DeleteAsync();
            }
            catch (System.IO.FileNotFoundException)
            {
            }
        }

        private static string GetContactPhotoFileName(string graphContactId)
        {
            HashAlgorithmProvider algorithm = HashAlgorithmProvider.OpenAlgorithm(HashAlgorithmNames.Sha256);
            IBuffer graphContactIdBuffer = CryptographicBuffer.ConvertStringToBinary(
                graphContactId,
                BinaryStringEncoding.Utf8);
            return CryptographicBuffer.EncodeToHexString(
                algorithm.HashData(graphContactIdBuffer)) + ".jpg";
        }

        private static void AddPhoneNumber(Contact contact, string number, ContactPhoneKind kind)
        {
            if (!string.IsNullOrWhiteSpace(number))
            {
                contact.Phones.Add(new ContactPhone { Number = number, Kind = kind });
            }
        }

        private static Appointment CreateAppointment(JsonObject graphEvent)
        {
            JsonObject start = JsonValueReader.GetObject(graphEvent, "start");
            JsonObject end = JsonValueReader.GetObject(graphEvent, "end");

            if (start == null || end == null)
            {
                return null;
            }

            DateTimeOffset startTime;
            DateTimeOffset endTime;

            if (!DateTimeOffset.TryParse(
                    JsonValueReader.GetString(start, "dateTime", string.Empty),
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out startTime)
                || !DateTimeOffset.TryParse(
                    JsonValueReader.GetString(end, "dateTime", string.Empty),
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out endTime))
            {
                return null;
            }

            startTime = startTime.ToLocalTime();
            endTime = endTime.ToLocalTime();
            JsonObject location = JsonValueReader.GetObject(graphEvent, "location") ?? new JsonObject();
            return new Appointment
            {
                Subject = JsonValueReader.GetString(graphEvent, "subject", "(No subject)"),
                Details = JsonValueReader.GetString(graphEvent, "bodyPreview", string.Empty),
                Location = JsonValueReader.GetString(location, "displayName", string.Empty),
                StartTime = startTime,
                Duration = endTime - startTime,
                AllDay = JsonValueReader.GetBoolean(graphEvent, "isAllDay", false)
            };
        }

        private static void PopulateAppointment(Appointment destination, Appointment source)
        {
            destination.Subject = source.Subject;
            destination.Details = source.Details;
            destination.Location = source.Location;
            destination.StartTime = source.StartTime;
            destination.Duration = source.Duration;
            destination.AllDay = source.AllDay;
        }

        private static string GetContactFingerprint(JsonObject graphContact)
        {
            JsonArray emails = JsonValueReader.GetArray(graphContact, "emailAddresses");

            foreach (IJsonValue email in emails)
            {
                if (email.ValueType == JsonValueType.Object)
                {
                    string address = JsonValueReader.GetString(email.GetObject(), "address", string.Empty);

                    if (!string.IsNullOrEmpty(address))
                    {
                        return "email:" + address.ToLowerInvariant();
                    }
                }
            }

            string mobilePhone = JsonValueReader.GetString(graphContact, "mobilePhone", string.Empty);
            return !string.IsNullOrEmpty(mobilePhone)
                ? "phone:" + NormalizePhoneNumber(mobilePhone)
                : "name:" + JsonValueReader.GetString(graphContact, "givenName", string.Empty).ToLowerInvariant()
                    + "|" + JsonValueReader.GetString(graphContact, "surname", string.Empty).ToLowerInvariant();
        }

        private static string GetContactFingerprint(Contact contact)
        {
            ContactEmail email = contact.Emails.FirstOrDefault();

            if (email != null && !string.IsNullOrEmpty(email.Address))
            {
                return "email:" + email.Address.ToLowerInvariant();
            }

            ContactPhone phone = contact.Phones.FirstOrDefault();
            return phone != null && !string.IsNullOrEmpty(phone.Number)
                ? "phone:" + NormalizePhoneNumber(phone.Number)
                : "name:" + (contact.FirstName ?? string.Empty).ToLowerInvariant()
                    + "|" + (contact.LastName ?? string.Empty).ToLowerInvariant();
        }

        private static string GetAppointmentFingerprint(Appointment appointment)
        {
            return string.Format(
                "{0}|{1}|{2}",
                appointment.Subject ?? string.Empty,
                appointment.StartTime.UtcDateTime.Ticks,
                appointment.Duration.Ticks);
        }

        private static string NormalizePhoneNumber(string number)
        {
            return new string(number.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
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

        private sealed class CalendarSyncData
        {
            public CalendarSyncData(string graphId, string displayName, IList<EventSyncData> events)
            {
                GraphId = graphId;
                DisplayName = displayName;
                Events = events;
            }

            public string GraphId { get; private set; }

            public string DisplayName { get; private set; }

            public IList<EventSyncData> Events { get; private set; }
        }

        private sealed class EventSyncData
        {
            public EventSyncData(string graphId, Appointment appointment)
            {
                GraphId = graphId;
                Appointment = appointment;
            }

            public string GraphId { get; private set; }

            public Appointment Appointment { get; private set; }
        }
    }
}
