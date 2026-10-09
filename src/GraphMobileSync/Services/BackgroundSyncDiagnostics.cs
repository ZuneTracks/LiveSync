using System;
using System.Collections.Generic;
using System.Linq;
using Windows.Data.Json;
using Windows.Storage;

namespace GraphMobileSync.Services
{
    internal static class BackgroundSyncDiagnostics
    {
        private const string LastStatusKey = "LastBackgroundSyncStatus";
        private const string HistoryKey = "SyncHistory";

        public static string GetLastStatus()
        {
            return ApplicationData.Current.LocalSettings.Values[LastStatusKey] as string;
        }

        public static void Record(string message)
        {
            string entry = string.Format(
                "{0:g}: {1}",
                DateTimeOffset.Now,
                (message ?? string.Empty).Substring(0, Math.Min((message ?? string.Empty).Length, 300)));
            ApplicationData.Current.LocalSettings.Values[LastStatusKey] = entry;
            List<string> history = GetHistory().ToList();
            history.Insert(0, entry);
            JsonArray serializedHistory = new JsonArray();

            foreach (string historyEntry in history.Take(10))
            {
                serializedHistory.Add(JsonValue.CreateStringValue(historyEntry));
            }

            ApplicationData.Current.LocalSettings.Values[HistoryKey] = serializedHistory.Stringify();
        }

        public static IReadOnlyList<string> GetHistory()
        {
            string serializedHistory = ApplicationData.Current.LocalSettings.Values[HistoryKey] as string;

            if (string.IsNullOrEmpty(serializedHistory))
            {
                return new List<string>();
            }

            try
            {
                return JsonArray.Parse(serializedHistory)
                    .Where(value => value.ValueType == JsonValueType.String)
                    .Select(value => value.GetString())
                    .ToList();
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }
    }
}
