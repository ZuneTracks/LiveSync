using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Windows.ApplicationModel.Appointments;
using Windows.ApplicationModel.Contacts;
using Windows.Data.Json;

namespace GraphMobileSync.Services
{
    internal sealed class TwoWaySyncResult
    {
        public int ContactsWritten { get; set; }

        public int EventsWritten { get; set; }

        public int ConflictsResolvedByGraph { get; set; }

        public int CalendarChangesIgnored { get; set; }
    }

    internal sealed class TwoWaySyncService
    {
        private readonly GraphClient _graphClient;

        public TwoWaySyncService(GraphClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task InitializeAsync(SyncSettings settings)
        {
            SyncMappings mappings = await SyncMappingsStore.LoadAsync();

            if (string.IsNullOrEmpty(mappings.ContactListId) && mappings.Calendars.Count == 0)
            {
                throw new InvalidOperationException(
                    "Run a full sync before enabling two-way synchronization.");
            }

            await InitializeContactTrackingAsync(mappings);
            await InitializeAppointmentTrackingAsync(mappings, settings);
            await SyncMappingsStore.SaveAsync(mappings);
        }

        // The caller loads and saves the mappings so a delta run touches the mappings file only once.
        public async Task<TwoWaySyncResult> PushLocalChangesAsync(
            SyncMappings mappings,
            ISet<string> graphChangedContactIds,
            IDictionary<string, ISet<string>> graphChangedEventIds)
        {
            TwoWaySyncResult result = new TwoWaySyncResult();

            await PushContactChangesAsync(mappings, graphChangedContactIds, result);
            await PushAppointmentChangesAsync(mappings, graphChangedEventIds, result);
            return result;
        }

        public async Task CreateAppointmentAsync(string graphCalendarId, Appointment appointment)
        {
            if (string.IsNullOrEmpty(graphCalendarId))
            {
                throw new ArgumentException("A Microsoft Graph calendar ID is required.", nameof(graphCalendarId));
            }

            if (appointment == null || string.IsNullOrEmpty(appointment.LocalId))
            {
                throw new InvalidOperationException(
                    "LiveSync could not identify the new local calendar event.");
            }

            SyncMappings mappings = await SyncMappingsStore.LoadAsync();
            CalendarMapping calendarMapping;

            if (!mappings.Calendars.TryGetValue(graphCalendarId, out calendarMapping))
            {
                throw new InvalidOperationException(
                    "The selected LiveSync calendar is no longer available. Run a full sync and try again.");
            }

            JsonObject createdEvent = await _graphClient.CreateEventAsync(
                graphCalendarId,
                CreateGraphEvent(appointment));
            string graphEventId = JsonValueReader.GetString(createdEvent, "id", string.Empty);

            if (string.IsNullOrEmpty(graphEventId))
            {
                throw new InvalidOperationException(
                    "Microsoft Graph did not return an ID for the new calendar event.");
            }

            calendarMapping.EventIds[graphEventId] = appointment.LocalId;
            calendarMapping.EventFingerprints[graphEventId] =
                SyncContentFingerprint.ForAppointment(appointment);
            await SyncMappingsStore.SaveAsync(mappings);
        }

        private static async Task InitializeContactTrackingAsync(SyncMappings mappings)
        {
            if (string.IsNullOrEmpty(mappings.ContactListId))
            {
                return;
            }

            ContactStore contactStore = await ContactManager.RequestStoreAsync(
                ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException(
                    "Contact access was not granted. Enable contact access for LiveSync and try again.");
            }

            ContactList contactList = (await contactStore.FindContactListsAsync())
                .FirstOrDefault(list => list.Id == mappings.ContactListId);

            if (contactList == null)
            {
                throw new InvalidOperationException(
                    "LiveSync's contact list is unavailable. Run a full sync before enabling two-way synchronization.");
            }

            foreach (KeyValuePair<string, string> mapping in mappings.ContactIds.ToList())
            {
                Contact contact = await contactList.GetContactAsync(mapping.Value);

                if (contact == null)
                {
                    mappings.RemoveContact(mapping.Key);
                    continue;
                }

                mappings.ContactFingerprints[mapping.Key] = SyncContentFingerprint.ForContact(contact);
            }

            ContactChangeTracker tracker = contactList.ChangeTracker;
            tracker.Enable();
        }

        private static async Task InitializeAppointmentTrackingAsync(
            SyncMappings mappings,
            SyncSettings settings)
        {
            if (mappings.Calendars.Count == 0)
            {
                return;
            }

            AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                AppointmentStoreAccessType.AppCalendarsReadWrite);

            if (appointmentStore == null)
            {
                throw new InvalidOperationException(
                    "Calendar access was not granted. Enable calendar access for LiveSync and try again.");
            }

            IReadOnlyList<AppointmentCalendar> calendars =
                await appointmentStore.FindAppointmentCalendarsAsync();

            foreach (CalendarMapping mapping in mappings.Calendars.Values)
            {
                AppointmentCalendar calendar = calendars.FirstOrDefault(
                    candidate => candidate.LocalId == mapping.LocalId);

                if (calendar == null)
                {
                    continue;
                }

                IReadOnlyList<Appointment> appointments = await calendar.FindAppointmentsAsync(
                    settings.GetStartDate(),
                    settings.GetEndDate() - settings.GetStartDate());
                Dictionary<string, Appointment> appointmentsById = appointments.ToDictionary(
                    appointment => appointment.LocalId);

                foreach (KeyValuePair<string, string> eventMapping in mapping.EventIds.ToList())
                {
                    Appointment appointment;

                    if (!appointmentsById.TryGetValue(eventMapping.Value, out appointment))
                    {
                        mapping.EventIds.Remove(eventMapping.Key);
                        mapping.EventFingerprints.Remove(eventMapping.Key);
                        continue;
                    }

                    mapping.EventFingerprints[eventMapping.Key] =
                        SyncContentFingerprint.ForAppointment(appointment);
                }
            }

            AppointmentStoreChangeTracker tracker = appointmentStore.ChangeTracker;
            tracker.Enable();
        }

        private async Task PushContactChangesAsync(
            SyncMappings mappings,
            ISet<string> graphChangedContactIds,
            TwoWaySyncResult result)
        {
            if (string.IsNullOrEmpty(mappings.ContactListId))
            {
                return;
            }

            ContactStore contactStore = await ContactManager.RequestStoreAsync(
                ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException(
                    "Contact access was not granted. Enable contact access for LiveSync and try again.");
            }

            ContactList contactList = (await contactStore.FindContactListsAsync())
                .FirstOrDefault(list => list.Id == mappings.ContactListId);

            if (contactList == null)
            {
                return;
            }

            ContactChangeTracker tracker = contactList.ChangeTracker;
            ContactChangeReader reader = tracker.GetChangeReader();
            Dictionary<string, string> graphIdsByLocalId = CreateReverseLookup(mappings.ContactIds);

            while (true)
            {
                IReadOnlyList<ContactChange> changes = await reader.ReadBatchAsync();

                if (changes.Count == 0)
                {
                    break;
                }

                foreach (ContactChange change in changes)
                {
                    if (change.ChangeType == ContactChangeType.ChangeTrackingLost)
                    {
                        tracker.Reset();
                        tracker.Enable();
                        throw new InvalidOperationException(
                            "Local contact change tracking was reset. Run a full sync before enabling two-way synchronization again.");
                    }

                    Contact contact = change.Contact;

                    if (contact == null || string.IsNullOrEmpty(contact.Id))
                    {
                        throw new InvalidOperationException(
                            "LiveSync could not identify a changed local contact.");
                    }

                    string graphContactId = FindGraphId(graphIdsByLocalId, contact.Id);

                    if (!string.IsNullOrEmpty(graphContactId)
                        && graphChangedContactIds.Contains(graphContactId))
                    {
                        result.ConflictsResolvedByGraph++;
                        continue;
                    }

                    if (change.ChangeType == ContactChangeType.Deleted)
                    {
                        if (!string.IsNullOrEmpty(graphContactId))
                        {
                            await _graphClient.DeleteContactAsync(graphContactId);
                            mappings.RemoveContact(graphContactId);
                            graphIdsByLocalId.Remove(contact.Id);
                            result.ContactsWritten++;
                        }

                        continue;
                    }

                    string fingerprint = SyncContentFingerprint.ForContact(contact);
                    string storedFingerprint;

                    if (!string.IsNullOrEmpty(graphContactId)
                        && mappings.ContactFingerprints.TryGetValue(
                            graphContactId,
                            out storedFingerprint)
                        && storedFingerprint == fingerprint)
                    {
                        continue;
                    }

                    JsonObject graphContact = CreateGraphContact(contact);

                    if (string.IsNullOrEmpty(graphContactId))
                    {
                        JsonObject createdContact = await _graphClient.CreateContactAsync(graphContact);
                        graphContactId = JsonValueReader.GetString(createdContact, "id", string.Empty);

                        if (string.IsNullOrEmpty(graphContactId))
                        {
                            throw new InvalidOperationException(
                                "Microsoft Graph did not return an ID for the new contact.");
                        }

                        mappings.ContactIds[graphContactId] = contact.Id;
                        graphIdsByLocalId[contact.Id] = graphContactId;
                    }
                    else
                    {
                        await _graphClient.UpdateContactAsync(graphContactId, graphContact);
                    }

                    // The upload gives the Graph contact a new changeKey; forget the old one so the
                    // next inbound sync re-evaluates this contact.
                    mappings.ContactChangeKeys.Remove(graphContactId);
                    mappings.ContactFingerprints[graphContactId] = fingerprint;
                    result.ContactsWritten++;
                }

                reader.AcceptChanges();
            }
        }

        private async Task PushAppointmentChangesAsync(
            SyncMappings mappings,
            IDictionary<string, ISet<string>> graphChangedEventIds,
            TwoWaySyncResult result)
        {
            if (mappings.Calendars.Count == 0)
            {
                return;
            }

            AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                AppointmentStoreAccessType.AppCalendarsReadWrite);

            if (appointmentStore == null)
            {
                throw new InvalidOperationException(
                    "Calendar access was not granted. Enable calendar access for LiveSync and try again.");
            }

            AppointmentStoreChangeTracker tracker = appointmentStore.ChangeTracker;
            AppointmentStoreChangeReader reader = tracker.GetChangeReader();
            Dictionary<string, string> graphCalendarIdsByLocalId = new Dictionary<string, string>();
            Dictionary<string, Dictionary<string, string>> graphEventIdsByCalendar =
                new Dictionary<string, Dictionary<string, string>>();

            foreach (KeyValuePair<string, CalendarMapping> entry in mappings.Calendars)
            {
                if (!string.IsNullOrEmpty(entry.Value.LocalId)
                    && !graphCalendarIdsByLocalId.ContainsKey(entry.Value.LocalId))
                {
                    graphCalendarIdsByLocalId[entry.Value.LocalId] = entry.Key;
                }
            }

            while (true)
            {
                IReadOnlyList<AppointmentStoreChange> changes = await reader.ReadBatchAsync();

                if (changes.Count == 0)
                {
                    break;
                }

                foreach (AppointmentStoreChange change in changes)
                {
                    if (change.ChangeType == AppointmentStoreChangeType.ChangeTrackingLost)
                    {
                        tracker.Reset();
                        tracker.Enable();
                        throw new InvalidOperationException(
                            "Local calendar change tracking was reset. Run a full sync before enabling two-way synchronization again.");
                    }

                    if (change.ChangeType != AppointmentStoreChangeType.AppointmentCreated
                        && change.ChangeType != AppointmentStoreChangeType.AppointmentModified
                        && change.ChangeType != AppointmentStoreChangeType.AppointmentDeleted)
                    {
                        continue;
                    }

                    AppointmentCalendar localCalendar = change.AppointmentCalendar;
                    Appointment appointment = change.Appointment;

                    if (localCalendar == null
                        || appointment == null
                        || string.IsNullOrEmpty(appointment.LocalId))
                    {
                        // Windows can retain a change record after the event or its calendar has
                        // been removed. It cannot be uploaded without an ID, so acknowledge it
                        // instead of replaying the same unusable record on every sync.
                        result.CalendarChangesIgnored++;
                        continue;
                    }

                    string graphCalendarId;

                    if (string.IsNullOrEmpty(localCalendar.LocalId)
                        || !graphCalendarIdsByLocalId.TryGetValue(localCalendar.LocalId, out graphCalendarId))
                    {
                        continue;
                    }

                    CalendarMapping calendarMapping = mappings.Calendars[graphCalendarId];
                    Dictionary<string, string> graphEventIdsByLocalId;

                    if (!graphEventIdsByCalendar.TryGetValue(graphCalendarId, out graphEventIdsByLocalId))
                    {
                        graphEventIdsByLocalId = CreateReverseLookup(calendarMapping.EventIds);
                        graphEventIdsByCalendar[graphCalendarId] = graphEventIdsByLocalId;
                    }

                    string graphEventId = FindGraphId(graphEventIdsByLocalId, appointment.LocalId);
                    ISet<string> changedEvents;

                    if (!string.IsNullOrEmpty(graphEventId)
                        && graphChangedEventIds.TryGetValue(graphCalendarId, out changedEvents)
                        && changedEvents.Contains(graphEventId))
                    {
                        result.ConflictsResolvedByGraph++;
                        continue;
                    }

                    if (change.ChangeType == AppointmentStoreChangeType.AppointmentDeleted)
                    {
                        if (!string.IsNullOrEmpty(graphEventId))
                        {
                            await _graphClient.DeleteEventAsync(graphCalendarId, graphEventId);
                            calendarMapping.EventIds.Remove(graphEventId);
                            calendarMapping.EventFingerprints.Remove(graphEventId);
                            graphEventIdsByLocalId.Remove(appointment.LocalId);
                            result.EventsWritten++;
                        }

                        continue;
                    }

                    string fingerprint = SyncContentFingerprint.ForAppointment(appointment);
                    string storedFingerprint;

                    if (!string.IsNullOrEmpty(graphEventId)
                        && calendarMapping.EventFingerprints.TryGetValue(
                            graphEventId,
                            out storedFingerprint)
                        && storedFingerprint == fingerprint)
                    {
                        continue;
                    }

                    JsonObject graphEvent = CreateGraphEvent(appointment);

                    if (string.IsNullOrEmpty(graphEventId))
                    {
                        JsonObject createdEvent = await _graphClient.CreateEventAsync(
                            graphCalendarId,
                            graphEvent);
                        graphEventId = JsonValueReader.GetString(createdEvent, "id", string.Empty);

                        if (string.IsNullOrEmpty(graphEventId))
                        {
                            throw new InvalidOperationException(
                                "Microsoft Graph did not return an ID for the new calendar event.");
                        }

                        calendarMapping.EventIds[graphEventId] = appointment.LocalId;
                        graphEventIdsByLocalId[appointment.LocalId] = graphEventId;
                    }
                    else
                    {
                        await _graphClient.UpdateEventAsync(
                            graphCalendarId,
                            graphEventId,
                            graphEvent);
                    }

                    calendarMapping.EventFingerprints[graphEventId] = fingerprint;
                    result.EventsWritten++;
                }

                reader.AcceptChanges();
            }
        }

        private static Dictionary<string, string> CreateReverseLookup(IDictionary<string, string> mappings)
        {
            Dictionary<string, string> graphIdsByLocalId = new Dictionary<string, string>();

            foreach (KeyValuePair<string, string> mapping in mappings)
            {
                if (!string.IsNullOrEmpty(mapping.Value) && !graphIdsByLocalId.ContainsKey(mapping.Value))
                {
                    graphIdsByLocalId[mapping.Value] = mapping.Key;
                }
            }

            return graphIdsByLocalId;
        }

        private static string FindGraphId(
            IDictionary<string, string> graphIdsByLocalId,
            string localId)
        {
            string graphId;
            return graphIdsByLocalId.TryGetValue(localId, out graphId) ? graphId : string.Empty;
        }

        private static JsonObject CreateGraphContact(Contact contact)
        {
            JsonObject graphContact = new JsonObject();
            SetString(graphContact, "givenName", contact.FirstName);
            SetString(graphContact, "surname", contact.LastName);
            SetString(graphContact, "displayName", contact.DisplayName);
            JsonArray emails = new JsonArray();

            foreach (ContactEmail email in contact.Emails)
            {
                if (string.IsNullOrWhiteSpace(email.Address))
                {
                    continue;
                }

                JsonObject graphEmail = new JsonObject();
                SetString(graphEmail, "address", email.Address);
                SetString(graphEmail, "name", contact.DisplayName);
                emails.Add(graphEmail);
            }

            graphContact.SetNamedValue("emailAddresses", emails);
            ContactPhone mobilePhone = contact.Phones.FirstOrDefault(
                phone => phone.Kind == ContactPhoneKind.Mobile);
            SetString(
                graphContact,
                "mobilePhone",
                mobilePhone == null ? string.Empty : mobilePhone.Number);
            JsonArray businessPhones = new JsonArray();

            foreach (ContactPhone phone in contact.Phones.Where(
                candidate => candidate.Kind == ContactPhoneKind.Work
                    && !string.IsNullOrWhiteSpace(candidate.Number)))
            {
                businessPhones.Add(JsonValue.CreateStringValue(phone.Number));
            }

            graphContact.SetNamedValue("businessPhones", businessPhones);
            return graphContact;
        }

        private static JsonObject CreateGraphEvent(Appointment appointment)
        {
            JsonObject graphEvent = new JsonObject();
            SetString(graphEvent, "subject", appointment.Subject);
            graphEvent.SetNamedValue("isAllDay", JsonValue.CreateBooleanValue(appointment.AllDay));
            JsonObject body = new JsonObject();
            SetString(body, "contentType", "text");
            SetString(body, "content", appointment.Details);
            graphEvent.SetNamedValue("body", body);
            JsonObject location = new JsonObject();
            SetString(location, "displayName", appointment.Location);
            graphEvent.SetNamedValue("location", location);

            if (appointment.AllDay)
            {
                // Graph rejects all-day events whose start or end is not midnight, so send whole dates.
                DateTime startDate = appointment.StartTime.ToLocalTime().Date;
                int days = Math.Max(1, (int)Math.Round(appointment.Duration.TotalDays));
                graphEvent.SetNamedValue("start", CreateGraphDateTime(startDate));
                graphEvent.SetNamedValue("end", CreateGraphDateTime(startDate.AddDays(days)));
                return graphEvent;
            }

            graphEvent.SetNamedValue("start", CreateGraphDateTime(appointment.StartTime));
            graphEvent.SetNamedValue(
                "end",
                CreateGraphDateTime(appointment.StartTime + appointment.Duration));
            return graphEvent;
        }

        private static JsonObject CreateGraphDateTime(DateTimeOffset value)
        {
            return CreateGraphDateTime(value.ToLocalTime().DateTime);
        }

        private static JsonObject CreateGraphDateTime(DateTime localDateTime)
        {
            JsonObject graphDateTime = new JsonObject();
            SetString(
                graphDateTime,
                "dateTime",
                localDateTime.ToString(
                    "yyyy-MM-ddTHH:mm:ss",
                    CultureInfo.InvariantCulture));
            SetString(graphDateTime, "timeZone", TimeZoneInfo.Local.Id);
            return graphDateTime;
        }

        private static void SetString(JsonObject target, string name, string value)
        {
            target.SetNamedValue(name, JsonValue.CreateStringValue(value ?? string.Empty));
        }
    }
}
