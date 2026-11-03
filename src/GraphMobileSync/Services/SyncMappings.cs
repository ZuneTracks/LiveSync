using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Json;
using Windows.Storage;

namespace GraphMobileSync.Services
{
    internal sealed class SyncMappings
    {
        public string ContactListId { get; set; }

        public IDictionary<string, string> ContactIds { get; private set; } =
            new Dictionary<string, string>();

        public IDictionary<string, string> ContactFingerprints { get; private set; } =
            new Dictionary<string, string>();

        // Graph contact changeKey values from the last successful download. An unchanged key means
        // the contact, including its photo, has not changed in Microsoft Graph since then.
        public IDictionary<string, string> ContactChangeKeys { get; private set; } =
            new Dictionary<string, string>();

        public IDictionary<string, CalendarMapping> Calendars { get; private set; } =
            new Dictionary<string, CalendarMapping>();

        internal string PersistedJson { get; set; }

        public void RemoveContact(string graphContactId)
        {
            ContactIds.Remove(graphContactId);
            ContactFingerprints.Remove(graphContactId);
            ContactChangeKeys.Remove(graphContactId);
        }
    }

    internal sealed class CalendarMapping
    {
        public string LocalId { get; set; }

        public string DisplayName { get; set; }

        public IDictionary<string, string> EventIds { get; private set; } =
            new Dictionary<string, string>();

        public IDictionary<string, string> EventFingerprints { get; private set; } =
            new Dictionary<string, string>();
    }

    internal sealed class SyncCleanupPlan
    {
        public IList<string> ContactListIds { get; private set; } = new List<string>();

        public IList<string> CalendarIds { get; private set; } = new List<string>();

        public bool HasItems
        {
            get
            {
                return ContactListIds.Count > 0 || CalendarIds.Count > 0;
            }
        }
    }

    internal static class SyncMappingsStore
    {
        private const string LegacyMappingsKey = "GraphDeviceMappings";
        private const string MappingsFileName = "GraphDeviceMappings.json";
        private const string LastPruneKey = "GraphDeviceMappingsLastPruneUtc";

        public static async Task<SyncMappings> LoadAsync()
        {
            string serializedMappings = await ReadMappingsFileAsync();
            bool loadedFromFile = serializedMappings != null;

            if (!loadedFromFile)
            {
                serializedMappings =
                    ApplicationData.Current.LocalSettings.Values[LegacyMappingsKey] as string;
            }

            if (string.IsNullOrEmpty(serializedMappings))
            {
                return new SyncMappings();
            }

            try
            {
                JsonObject root = JsonObject.Parse(serializedMappings);
                SyncMappings mappings = new SyncMappings
                {
                    ContactListId = JsonValueReader.GetString(root, "contactListId", string.Empty)
                };

                ReadStringMap(JsonValueReader.GetObject(root, "contacts"), mappings.ContactIds);
                ReadStringMap(
                    JsonValueReader.GetObject(root, "contactFingerprints"),
                    mappings.ContactFingerprints);
                ReadStringMap(
                    JsonValueReader.GetObject(root, "contactChangeKeys"),
                    mappings.ContactChangeKeys);
                JsonObject calendars = JsonValueReader.GetObject(root, "calendars");

                if (calendars != null)
                {
                    foreach (KeyValuePair<string, IJsonValue> calendarEntry in calendars)
                    {
                        if (calendarEntry.Value.ValueType != JsonValueType.Object)
                        {
                            continue;
                        }

                        JsonObject serializedCalendar = calendarEntry.Value.GetObject();
                        CalendarMapping calendar = new CalendarMapping
                        {
                            LocalId = JsonValueReader.GetString(serializedCalendar, "localId", string.Empty),
                            DisplayName = JsonValueReader.GetString(serializedCalendar, "displayName", string.Empty)
                        };
                        ReadStringMap(JsonValueReader.GetObject(serializedCalendar, "events"), calendar.EventIds);
                        ReadStringMap(
                            JsonValueReader.GetObject(serializedCalendar, "eventFingerprints"),
                            calendar.EventFingerprints);
                        mappings.Calendars[calendarEntry.Key] = calendar;
                    }
                }

                if (loadedFromFile)
                {
                    mappings.PersistedJson = serializedMappings;
                }

                return mappings;
            }
            catch (Exception exception)
            {
                BackgroundSyncDiagnostics.Record("Stored device mapping is invalid: " + exception.Message);
                return new SyncMappings();
            }
        }

        public static async Task SaveAsync(SyncMappings mappings)
        {
            JsonObject root = new JsonObject();
            root.SetNamedValue("contactListId", JsonValue.CreateStringValue(mappings.ContactListId ?? string.Empty));
            root.SetNamedValue("contacts", WriteStringMap(mappings.ContactIds));
            root.SetNamedValue("contactFingerprints", WriteStringMap(mappings.ContactFingerprints));
            root.SetNamedValue("contactChangeKeys", WriteStringMap(mappings.ContactChangeKeys));
            JsonObject calendars = new JsonObject();

            foreach (KeyValuePair<string, CalendarMapping> calendarEntry in mappings.Calendars)
            {
                CalendarMapping calendar = calendarEntry.Value;
                JsonObject serializedCalendar = new JsonObject();
                serializedCalendar.SetNamedValue("localId", JsonValue.CreateStringValue(calendar.LocalId ?? string.Empty));
                serializedCalendar.SetNamedValue("displayName", JsonValue.CreateStringValue(calendar.DisplayName ?? string.Empty));
                serializedCalendar.SetNamedValue("events", WriteStringMap(calendar.EventIds));
                serializedCalendar.SetNamedValue(
                    "eventFingerprints",
                    WriteStringMap(calendar.EventFingerprints));
                calendars.SetNamedValue(calendarEntry.Key, serializedCalendar);
            }

            root.SetNamedValue("calendars", calendars);
            string serializedMappings = root.Stringify();

            // Most background checks change nothing, so skip rewriting an identical file.
            if (serializedMappings == mappings.PersistedJson)
            {
                return;
            }

            StorageFile mappingsFile = await ApplicationData.Current.LocalFolder.CreateFileAsync(
                MappingsFileName,
                CreationCollisionOption.ReplaceExisting);
            await FileIO.WriteTextAsync(mappingsFile, serializedMappings);
            mappings.PersistedJson = serializedMappings;
            ApplicationData.Current.LocalSettings.Values.Remove(LegacyMappingsKey);
        }

        public static bool IsPruneDue()
        {
            string serializedLastPrune =
                ApplicationData.Current.LocalSettings.Values[LastPruneKey] as string;
            DateTimeOffset lastPrune;

            return !DateTimeOffset.TryParse(
                    serializedLastPrune,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out lastPrune)
                || DateTimeOffset.UtcNow - lastPrune >= TimeSpan.FromDays(1);
        }

        public static void MarkPruned()
        {
            ApplicationData.Current.LocalSettings.Values[LastPruneKey] =
                DateTimeOffset.UtcNow.UtcDateTime.ToString("o", CultureInfo.InvariantCulture);
        }

        private static async Task<string> ReadMappingsFileAsync()
        {
            try
            {
                StorageFile mappingsFile = await ApplicationData.Current.LocalFolder.GetFileAsync(MappingsFileName);
                return await FileIO.ReadTextAsync(mappingsFile);
            }
            catch (FileNotFoundException)
            {
                return null;
            }
        }

        private static void ReadStringMap(JsonObject serializedMap, IDictionary<string, string> target)
        {
            if (serializedMap == null)
            {
                return;
            }

            foreach (KeyValuePair<string, IJsonValue> entry in serializedMap)
            {
                if (entry.Value.ValueType == JsonValueType.String)
                {
                    target[entry.Key] = entry.Value.GetString();
                }
            }
        }

        private static JsonObject WriteStringMap(IDictionary<string, string> map)
        {
            JsonObject serializedMap = new JsonObject();

            foreach (KeyValuePair<string, string> entry in map)
            {
                serializedMap.SetNamedValue(entry.Key, JsonValue.CreateStringValue(entry.Value));
            }

            return serializedMap;
        }
    }
}
