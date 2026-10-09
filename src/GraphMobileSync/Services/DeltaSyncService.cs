using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Windows.Data.Json;

namespace GraphMobileSync.Services
{
    internal sealed class DeltaSyncResult
    {
        public DeltaSyncResult(bool changesApplied, string message)
        {
            ChangesApplied = changesApplied;
            Message = message;
        }

        public bool ChangesApplied { get; private set; }

        public string Message { get; private set; }
    }

    internal sealed class DeltaSyncService
    {
        private readonly GraphClient _graphClient;

        public DeltaSyncService(GraphClient graphClient)
        {
            _graphClient = graphClient;
        }

        public async Task<DeltaSyncResult> PollAndApplyAsync(bool ignoreBackgroundSetting = false)
        {
            try
            {
                SyncSettings settings = SyncSettings.Load();

                if (!ignoreBackgroundSetting && !settings.BackgroundSyncEnabled)
                {
                    return new DeltaSyncResult(false, "Background sync is disabled.");
                }

                DeltaCheckpoint checkpoint = await DeltaCheckpointService.LoadAsync();

                if (checkpoint == null)
                {
                    return new DeltaSyncResult(
                        false,
                        "Waiting for the first foreground sync to establish change tracking.");
                }

                GraphDeltaResult contactDelta = await _graphClient.GetDeltaAsync(checkpoint.ContactDeltaLink);
                IList<JsonObject> graphCalendars = await _graphClient.GetCollectionAsync(
                    "/v1.0/me/calendars?$select=id,name&$top=50");
                Dictionary<string, string> updatedCalendarLinks =
                    new Dictionary<string, string>();
                HashSet<string> currentCalendarIds = new HashSet<string>();
                List<CalendarDeltaBatch> calendarChanges = new List<CalendarDeltaBatch>();
                bool graphChangesDetected = contactDelta.Items.Count > 0;
                bool pruneMappings = SyncMappingsStore.IsPruneDue();

                foreach (JsonObject graphCalendar in graphCalendars)
                {
                    string calendarId = JsonValueReader.GetString(graphCalendar, "id", string.Empty);
                    string calendarName = JsonValueReader.GetString(graphCalendar, "name", "Calendar");

                    if (string.IsNullOrEmpty(calendarId) || !settings.IncludesCalendar(calendarName))
                    {
                        continue;
                    }

                    currentCalendarIds.Add(calendarId);
                    string calendarDeltaLink;
                    GraphDeltaResult eventDelta;

                    if (!checkpoint.CalendarDeltaLinks.TryGetValue(calendarId, out calendarDeltaLink))
                    {
                        eventDelta = await _graphClient.GetDeltaAsync(
                            DeltaCheckpointService.GetCalendarDeltaPath(calendarId, settings));
                        graphChangesDetected = true;
                        calendarChanges.Add(new CalendarDeltaBatch(
                            calendarId,
                            calendarName,
                            eventDelta.Items));
                    }
                    else
                    {
                        eventDelta = await _graphClient.GetDeltaAsync(calendarDeltaLink);

                        if (eventDelta.Items.Count > 0)
                        {
                            graphChangesDetected = true;
                            calendarChanges.Add(new CalendarDeltaBatch(
                                calendarId,
                                calendarName,
                                eventDelta.Items));
                        }
                    }

                    updatedCalendarLinks[calendarId] = eventDelta.DeltaLink;

                    if (pruneMappings
                        && !calendarChanges.Any(change => change.GraphId == calendarId))
                    {
                        calendarChanges.Add(new CalendarDeltaBatch(
                            calendarId,
                            calendarName,
                            new List<JsonObject>()));
                    }
                }

                if (!checkpoint.CalendarDeltaLinks.Keys.All(currentCalendarIds.Contains))
                {
                    graphChangesDetected = true;
                }

                if (!graphChangesDetected && !pruneMappings)
                {
                    await DeltaCheckpointService.SaveAsync(
                        new DeltaCheckpoint(contactDelta.DeltaLink, updatedCalendarLinks));
                    return new DeltaSyncResult(false, "No Microsoft Graph changes detected.");
                }

                await new GraphSyncService(_graphClient).ApplyDeltaAsync(
                    contactDelta.Items,
                    calendarChanges,
                    currentCalendarIds,
                    settings);
                await DeltaCheckpointService.SaveAsync(
                    new DeltaCheckpoint(contactDelta.DeltaLink, updatedCalendarLinks));

                if (pruneMappings)
                {
                    SyncMappingsStore.MarkPruned();
                }

                return new DeltaSyncResult(
                    true,
                    graphChangesDetected
                        ? "Applied Microsoft Graph contact or calendar changes."
                        : "Pruned stale local synchronization mappings.");
            }
            catch (GraphThrottledException exception)
            {
                return new DeltaSyncResult(false, exception.Message);
            }
        }
    }
}
