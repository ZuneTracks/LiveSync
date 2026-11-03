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

        public async Task<DeltaSyncResult> PollAndApplyAsync(
            bool ignoreBackgroundSetting = false,
            bool outboundWritesAvailable = true,
            LiveSyncProEntitlement verifiedEntitlement = null)
        {
            SyncMappings mappings = null;
            bool mappingsSaved = false;

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

                TwoWaySyncResult twoWayResult = null;
                string twoWayStatus = string.Empty;

                if (settings.TwoWaySyncEnabled && outboundWritesAvailable)
                {
                    LiveSyncProEntitlement entitlement = verifiedEntitlement;

                    if (entitlement == null)
                    {
                        try
                        {
                            entitlement = await LiveSyncProService.GetEntitlementAsync();
                        }
                        catch (Exception exception)
                        {
                            twoWayStatus =
                                "LiveSync Pro status could not be verified; local edits were not uploaded: "
                                + exception.Message;
                            BackgroundSyncDiagnostics.Record(twoWayStatus);
                        }
                    }

                    if (entitlement != null && entitlement.IsActive)
                    {
                        mappings = await SyncMappingsStore.LoadAsync();
                        twoWayResult = await new TwoWaySyncService(_graphClient)
                            .PushLocalChangesAsync(
                                mappings,
                                GetChangedIds(contactDelta.Items),
                                GetChangedEventIds(calendarChanges));
                    }
                    else if (entitlement != null)
                    {
                        twoWayStatus =
                            "LiveSync Pro is not active; local edits were not uploaded.";
                    }
                }
                else if (settings.TwoWaySyncEnabled)
                {
                    twoWayStatus =
                        "Microsoft Graph write authorization needs attention; local edits were not uploaded.";
                }

                SyncSummary deltaSummary = null;

                if (graphChangesDetected || pruneMappings)
                {
                    if (mappings == null)
                    {
                        mappings = await SyncMappingsStore.LoadAsync();
                    }

                    deltaSummary = await new GraphSyncService(_graphClient).ApplyDeltaAsync(
                        contactChanges: contactDelta.Items,
                        calendarChanges: calendarChanges,
                        activeCalendarIds: currentCalendarIds,
                        settings: settings,
                        mappings: mappings);

                    // Background runs detect duplicate app-owned contact lists too. Discarding the plan
                    // left their contacts in the People app until the user ran a manual sync.
                    await GraphSyncService.CleanupDuplicateContactListsAsync(
                        deltaSummary.CleanupPlan.ContactListIds);
                }

                // Mappings must be persisted before the checkpoint advances; otherwise a failed save
                // would leave applied items unmapped and the next delta would sweep them as orphans.
                if (mappings != null)
                {
                    await SyncMappingsStore.SaveAsync(mappings);
                    mappingsSaved = true;
                }

                await DeltaCheckpointService.SaveAsync(
                    new DeltaCheckpoint(contactDelta.DeltaLink, updatedCalendarLinks));

                if (pruneMappings)
                {
                    SyncMappingsStore.MarkPruned();
                }

                int localWrites = twoWayResult == null
                    ? 0
                    : twoWayResult.ContactsWritten + twoWayResult.EventsWritten;
                int graphConflicts = twoWayResult == null
                    ? 0
                    : twoWayResult.ConflictsResolvedByGraph;
                int ignoredCalendarChanges = twoWayResult == null
                    ? 0
                    : twoWayResult.CalendarChangesIgnored;
                string message;

                if (graphChangesDetected)
                {
                    message = "Applied Microsoft Graph contact or calendar changes.";
                }
                else if (pruneMappings)
                {
                    message = "Pruned stale local synchronization mappings.";
                }
                else if (localWrites > 0)
                {
                    message = "Uploaded local contact or calendar edits to Microsoft Graph.";
                }
                else
                {
                    message = "No Microsoft Graph or local changes detected.";
                }

                if (graphConflicts > 0)
                {
                    message += " Microsoft Graph was kept for "
                        + graphConflicts
                        + " same-interval conflict(s).";
                }

                if (ignoredCalendarChanges > 0)
                {
                    message += " Ignored "
                        + ignoredCalendarChanges
                        + " stale local calendar change record(s) reported by Windows.";
                }

                if (!string.IsNullOrEmpty(twoWayStatus))
                {
                    message += " " + twoWayStatus;
                }

                return new DeltaSyncResult(
                    graphChangesDetected || pruneMappings || localWrites > 0,
                    message);
            }
            catch (GraphThrottledException exception)
            {
                return new DeltaSyncResult(false, exception.Message);
            }
            finally
            {
                // Local and Microsoft Graph writes made before a failure are already permanent, so the
                // mappings that describe them must be saved even when the run does not complete.
                if (mappings != null && !mappingsSaved)
                {
                    try
                    {
                        await SyncMappingsStore.SaveAsync(mappings);
                    }
                    catch (Exception exception)
                    {
                        BackgroundSyncDiagnostics.Record(
                            "Synchronization mappings could not be saved after a failed run: "
                            + exception.Message);
                    }
                }
            }
        }

        private static ISet<string> GetChangedIds(IEnumerable<JsonObject> changes)
        {
            return new HashSet<string>(
                changes
                    .Select(change => JsonValueReader.GetString(change, "id", string.Empty))
                    .Where(id => !string.IsNullOrEmpty(id)));
        }

        private static IDictionary<string, ISet<string>> GetChangedEventIds(
            IEnumerable<CalendarDeltaBatch> calendarChanges)
        {
            Dictionary<string, ISet<string>> changes =
                new Dictionary<string, ISet<string>>();

            foreach (CalendarDeltaBatch calendarChange in calendarChanges)
            {
                changes[calendarChange.GraphId] = GetChangedIds(calendarChange.Items);
            }

            return changes;
        }
    }
}
