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
    internal sealed partial class GraphSyncService
    {
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
            Dictionary<string, Appointment> pendingAppointments = new Dictionary<string, Appointment>();

            foreach (JsonObject graphEvent in eventChanges)
            {
                string graphEventId = JsonValueReader.GetString(graphEvent, "id", string.Empty);

                if (string.IsNullOrEmpty(graphEventId))
                {
                    continue;
                }

                if (pendingAppointments.ContainsKey(graphEventId))
                {
                    AddResolvedAppointments(
                        existingById,
                        appointmentIdsInWindow,
                        await ResolvePendingAppointmentIdsAsync(
                            calendar,
                            pendingAppointments,
                            calendarMapping,
                            existingById.Keys));
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
                        calendarMapping.EventFingerprints.Remove(graphEventId);
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

                updatedAppointment = await UpsertAppointmentAsync(
                    calendar,
                    updatedAppointment,
                    sourceAppointment,
                    summary);
                localAppointmentId = updatedAppointment.LocalId;
                summary.Events++;

                if (string.IsNullOrEmpty(localAppointmentId))
                {
                    pendingAppointments[graphEventId] = updatedAppointment;
                    continue;
                }

                calendarMapping.EventIds[graphEventId] = localAppointmentId;
                calendarMapping.EventFingerprints[graphEventId] =
                    SyncContentFingerprint.ForAppointment(updatedAppointment);
                appointmentIdsInWindow.Add(localAppointmentId);
                existingById[localAppointmentId] = updatedAppointment;
            }

            AddResolvedAppointments(
                existingById,
                appointmentIdsInWindow,
                await ResolvePendingAppointmentIdsAsync(
                    calendar,
                    pendingAppointments,
                    calendarMapping,
                    existingById.Keys));

            foreach (KeyValuePair<string, string> mapping in calendarMapping.EventIds.ToList())
            {
                if (string.IsNullOrEmpty(mapping.Value)
                    || !appointmentIdsInWindow.Contains(mapping.Value))
                {
                    calendarMapping.EventIds.Remove(mapping.Key);
                    calendarMapping.EventFingerprints.Remove(mapping.Key);
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
            Dictionary<string, Appointment> pendingAppointments = new Dictionary<string, Appointment>();

            foreach (EventSyncData sourceEvent in sourceCalendar.Events)
            {
                if (pendingAppointments.ContainsKey(sourceEvent.GraphId))
                {
                    await ResolvePendingAppointmentIdsAsync(
                        calendar,
                        pendingAppointments,
                        calendarMapping,
                        existingById.Keys);
                }

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

                localAppointment = await UpsertAppointmentAsync(
                    calendar,
                    localAppointment,
                    sourceEvent.Appointment,
                    summary);
                localAppointmentId = localAppointment.LocalId;

                if (string.IsNullOrEmpty(localAppointmentId))
                {
                    pendingAppointments[sourceEvent.GraphId] = localAppointment;
                }
                else
                {
                    calendarMapping.EventIds[sourceEvent.GraphId] = localAppointmentId;
                    calendarMapping.EventFingerprints[sourceEvent.GraphId] =
                        SyncContentFingerprint.ForAppointment(localAppointment);
                }

                summary.Events++;
                completedItems++;
                ReportSyncProgress(progress, "Syncing calendar events...", completedItems, totalItems);
            }

            await ResolvePendingAppointmentIdsAsync(
                calendar,
                pendingAppointments,
                calendarMapping,
                existingById.Keys);

            foreach (KeyValuePair<string, string> mapping in calendarMapping.EventIds.ToList())
            {
                if (!graphEventIds.Contains(mapping.Key))
                {
                    if (!string.IsNullOrEmpty(mapping.Value))
                    {
                        await calendar.DeleteAppointmentAsync(mapping.Value);
                    }

                    calendarMapping.EventIds.Remove(mapping.Key);
                    calendarMapping.EventFingerprints.Remove(mapping.Key);
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

        private static async Task<Appointment> UpsertAppointmentAsync(
            AppointmentCalendar calendar,
            Appointment localAppointment,
            Appointment sourceAppointment,
            SyncSummary summary)
        {
            if (localAppointment == null)
            {
                await calendar.SaveAppointmentAsync(sourceAppointment);
                return sourceAppointment;
            }

            if (SyncContentFingerprint.ForAppointment(localAppointment)
                == SyncContentFingerprint.ForAppointment(sourceAppointment))
            {
                summary.EventsUnchanged++;
                return localAppointment;
            }

            PopulateAppointment(localAppointment, sourceAppointment);
            await calendar.SaveAppointmentAsync(localAppointment);
            return localAppointment;
        }

        private static void AddResolvedAppointments(
            IDictionary<string, Appointment> existingById,
            ISet<string> appointmentIdsInWindow,
            IEnumerable<Appointment> resolvedAppointments)
        {
            foreach (Appointment appointment in resolvedAppointments)
            {
                existingById[appointment.LocalId] = appointment;
                appointmentIdsInWindow.Add(appointment.LocalId);
            }
        }

        // Resolves every new appointment whose LocalId was left empty by SaveAppointmentAsync with a
        // single calendar query, instead of one query per saved appointment.
        private static async Task<List<Appointment>> ResolvePendingAppointmentIdsAsync(
            AppointmentCalendar calendar,
            IDictionary<string, Appointment> pendingAppointments,
            CalendarMapping calendarMapping,
            IEnumerable<string> preexistingAppointmentIds)
        {
            List<Appointment> resolvedAppointments = new List<Appointment>();

            if (pendingAppointments.Count == 0)
            {
                return resolvedAppointments;
            }

            TimeSpan padding = TimeSpan.FromMinutes(1);
            DateTimeOffset rangeStart = pendingAppointments.Values.Min(item => item.StartTime) - padding;
            DateTimeOffset rangeEnd = pendingAppointments.Values.Max(item => item.StartTime + item.Duration)
                + padding;
            IReadOnlyList<Appointment> savedAppointments =
                await calendar.FindAppointmentsAsync(rangeStart, rangeEnd - rangeStart);
            HashSet<string> unavailableIds = new HashSet<string>(preexistingAppointmentIds);
            unavailableIds.UnionWith(calendarMapping.EventIds.Values.Where(id => !string.IsNullOrEmpty(id)));
            List<Appointment> candidates = savedAppointments
                .Where(appointment => !unavailableIds.Contains(appointment.LocalId))
                .ToList();

            foreach (KeyValuePair<string, Appointment> pendingAppointment in pendingAppointments)
            {
                string contentFingerprint = SyncContentFingerprint.ForAppointment(pendingAppointment.Value);
                string identityFingerprint = GetAppointmentFingerprint(pendingAppointment.Value);
                Appointment savedAppointment =
                    FindUniqueAppointment(
                        candidates,
                        candidate => SyncContentFingerprint.ForAppointment(candidate) == contentFingerprint)
                    ?? FindUniqueAppointment(
                        candidates,
                        candidate => GetAppointmentFingerprint(candidate) == identityFingerprint);

                if (savedAppointment == null)
                {
                    throw new InvalidOperationException("LiveSync could not identify a saved calendar event.");
                }

                candidates.Remove(savedAppointment);
                calendarMapping.EventIds[pendingAppointment.Key] = savedAppointment.LocalId;
                calendarMapping.EventFingerprints[pendingAppointment.Key] = contentFingerprint;
                resolvedAppointments.Add(savedAppointment);
            }

            pendingAppointments.Clear();
            return resolvedAppointments;
        }

        private static Appointment FindUniqueAppointment(
            IEnumerable<Appointment> appointments,
            Func<Appointment, bool> predicate)
        {
            List<Appointment> matches = appointments.Where(predicate).Take(2).ToList();
            return matches.Count == 1 ? matches[0] : null;
        }

        private static async Task<AppointmentCalendar> GetCalendarAsync(
            AppointmentStore appointmentStore,
            IList<AppointmentCalendar> calendars,
            string graphCalendarId,
            string graphCalendarDisplayName,
            SyncMappings mappings,
            SyncCleanupPlan cleanupPlan)
        {
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
                ?? matchingCalendars.FirstOrDefault();

            if (calendar == null)
            {
                calendar = await appointmentStore.CreateAppointmentCalendarAsync(displayName);
                calendars.Add(calendar);
            }

            if (calendar.OtherAppWriteAccess != AppointmentCalendarOtherAppWriteAccess.SystemOnly)
            {
                calendar.OtherAppWriteAccess = AppointmentCalendarOtherAppWriteAccess.SystemOnly;
                await calendar.SaveAsync();
            }

            mapping.LocalId = calendar.LocalId;
            mapping.DisplayName = graphCalendarDisplayName;

            foreach (AppointmentCalendar duplicateCalendar in matchingCalendars.Where(item => item.LocalId != calendar.LocalId))
            {
                cleanupPlan.CalendarIds.Add(duplicateCalendar.LocalId);
            }

            return calendar;
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
                "/v1.0/me/calendars/{0}/calendarView?startDateTime={1}&endDateTime={2}&$select=id,subject,bodyPreview,start,end,location,isAllDay&$top=500",
                Uri.EscapeDataString(calendarId),
                Uri.EscapeDataString(settings.GetStartDate().ToString("o")),
                Uri.EscapeDataString(settings.GetEndDate().ToString("o")));
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

        private static string GetAppointmentFingerprint(Appointment appointment)
        {
            return string.Format(
                "{0}|{1}|{2}",
                appointment.Subject ?? string.Empty,
                appointment.StartTime.UtcDateTime.Ticks,
                appointment.Duration.Ticks);
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
