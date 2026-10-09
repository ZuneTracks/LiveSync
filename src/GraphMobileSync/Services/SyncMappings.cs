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

        public IDictionary<string, CalendarMapping> Calendars { get; private set; } =
            new Dictionary<string, CalendarMapping>();
    }

    internal sealed class CalendarMapping
    {
        public string LocalId { get; set; }

        public string DisplayName { get; set; }

        public IDictionary<string, string> EventIds { get; private set; } =
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
            string serializedMappings = await ReadMappingsAsync();

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
                        mappings.Calendars[calendarEntry.Key] = calendar;
                    }
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
            JsonObject calendars = new JsonObject();

            foreach (KeyValuePair<string, CalendarMapping> calendarEntry in mappings.Calendars)
            {
                CalendarMapping calendar = calendarEntry.Value;
                JsonObject serializedCalendar = new JsonObject();
                serializedCalendar.SetNamedValue("localId", JsonValue.CreateStringValue(calendar.LocalId ?? string.Empty));
                serializedCalendar.SetNamedValue("displayName", JsonValue.CreateStringValue(calendar.DisplayName ?? string.Empty));
                serializedCalendar.SetNamedValue("events", WriteStringMap(calendar.EventIds));
                calendars.SetNamedValue(calendarEntry.Key, serializedCalendar);
            }

            root.SetNamedValue("calendars", calendars);
            StorageFile mappingsFile = await ApplicationData.Current.LocalFolder.CreateFileAsync(
                MappingsFileName,
                CreationCollisionOption.ReplaceExisting);
            await FileIO.WriteTextAsync(mappingsFile, root.Stringify());
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

        private static async Task<string> ReadMappingsAsync()
        {
            try
            {
                StorageFile mappingsFile = await ApplicationData.Current.LocalFolder.GetFileAsync(MappingsFileName);
                return await FileIO.ReadTextAsync(mappingsFile);
            }
            catch (FileNotFoundException)
            {
                return ApplicationData.Current.LocalSettings.Values[LegacyMappingsKey] as string;
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
