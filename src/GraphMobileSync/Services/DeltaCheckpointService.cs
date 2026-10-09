using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Json;
using Windows.Storage;

namespace GraphMobileSync.Services
{
    internal sealed class DeltaCheckpoint
    {
        public DeltaCheckpoint(string contactDeltaLink, IDictionary<string, string> calendarDeltaLinks)
        {
            ContactDeltaLink = contactDeltaLink;
            CalendarDeltaLinks = calendarDeltaLinks;
        }

        public string ContactDeltaLink { get; private set; }

        public IDictionary<string, string> CalendarDeltaLinks { get; private set; }
    }

    internal static class DeltaCheckpointService
    {
        private const string LegacyContactDeltaLinkKey = "ContactDeltaLink";
        private const string LegacyCalendarDeltaLinksKey = "CalendarDeltaLinks";
        private const string CheckpointFileName = "GraphDeltaCheckpoint.json";
        private const string ContactDeltaPath =
            "/v1.0/me/contacts/delta?$select=id,givenName,surname,displayName,emailAddresses,mobilePhone,businessPhones";

        public static async Task CreateAsync(GraphClient graphClient)
        {
            SyncSettings settings = SyncSettings.Load();
            GraphDeltaResult contactDelta = await graphClient.GetDeltaAsync(ContactDeltaPath);
            IList<JsonObject> graphCalendars = await graphClient.GetCollectionAsync(
                "/v1.0/me/calendars?$select=id,name&$top=50");
            Dictionary<string, string> calendarDeltaLinks = new Dictionary<string, string>();

            foreach (JsonObject graphCalendar in graphCalendars)
            {
                string calendarId = JsonValueReader.GetString(graphCalendar, "id", string.Empty);

                if (string.IsNullOrEmpty(calendarId))
                {
                    continue;
                }

                string calendarName = JsonValueReader.GetString(graphCalendar, "name", "Calendar");

                if (!settings.IncludesCalendar(calendarName))
                {
                    continue;
                }

                GraphDeltaResult eventDelta = await graphClient.GetDeltaAsync(GetCalendarDeltaPath(calendarId, settings));
                calendarDeltaLinks[calendarId] = eventDelta.DeltaLink;
            }

            await SaveAsync(new DeltaCheckpoint(contactDelta.DeltaLink, calendarDeltaLinks));
        }

        public static async Task<DeltaCheckpoint> LoadAsync()
        {
            string serializedCheckpoint = await ReadCheckpointAsync();

            if (string.IsNullOrEmpty(serializedCheckpoint))
            {
                return null;
            }

            try
            {
                JsonObject root = JsonObject.Parse(serializedCheckpoint);
                string contactDeltaLink = JsonValueReader.GetString(root, "contactDeltaLink", string.Empty);
                JsonObject serializedLinks = JsonValueReader.GetObject(root, "calendarDeltaLinks");

                if (string.IsNullOrEmpty(contactDeltaLink) || serializedLinks == null)
                {
                    return null;
                }

                Dictionary<string, string> calendarDeltaLinks = new Dictionary<string, string>();

                foreach (KeyValuePair<string, IJsonValue> entry in serializedLinks)
                {
                    if (entry.Value.ValueType == JsonValueType.String)
                    {
                        calendarDeltaLinks[entry.Key] = entry.Value.GetString();
                    }
                }

                return new DeltaCheckpoint(contactDeltaLink, calendarDeltaLinks);
            }
            catch (Exception exception)
            {
                BackgroundSyncDiagnostics.Record("Stored delta state is invalid: " + exception.Message);
                return null;
            }
        }

        public static async Task SaveAsync(DeltaCheckpoint checkpoint)
        {
            if (checkpoint == null || string.IsNullOrEmpty(checkpoint.ContactDeltaLink))
            {
                throw new ArgumentException("A complete delta checkpoint is required.", nameof(checkpoint));
            }

            JsonObject serializedLinks = new JsonObject();

            foreach (KeyValuePair<string, string> entry in checkpoint.CalendarDeltaLinks)
            {
                serializedLinks.SetNamedValue(entry.Key, JsonValue.CreateStringValue(entry.Value));
            }

            JsonObject root = new JsonObject();
            root.SetNamedValue("contactDeltaLink", JsonValue.CreateStringValue(checkpoint.ContactDeltaLink));
            root.SetNamedValue("calendarDeltaLinks", serializedLinks);
            StorageFile checkpointFile = await ApplicationData.Current.LocalFolder.CreateFileAsync(
                CheckpointFileName,
                CreationCollisionOption.ReplaceExisting);
            await FileIO.WriteTextAsync(checkpointFile, root.Stringify());
            ApplicationData.Current.LocalSettings.Values.Remove(LegacyContactDeltaLinkKey);
            ApplicationData.Current.LocalSettings.Values.Remove(LegacyCalendarDeltaLinksKey);
        }

        private static async Task<string> ReadCheckpointAsync()
        {
            try
            {
                StorageFile checkpointFile = await ApplicationData.Current.LocalFolder.GetFileAsync(CheckpointFileName);
                return await FileIO.ReadTextAsync(checkpointFile);
            }
            catch (FileNotFoundException)
            {
                ApplicationDataContainer settings = ApplicationData.Current.LocalSettings;
                string contactDeltaLink = settings.Values[LegacyContactDeltaLinkKey] as string;
                string calendarDeltaLinks = settings.Values[LegacyCalendarDeltaLinksKey] as string;

                if (string.IsNullOrEmpty(contactDeltaLink) || string.IsNullOrEmpty(calendarDeltaLinks))
                {
                    return string.Empty;
                }

                JsonObject root = new JsonObject();
                root.SetNamedValue("contactDeltaLink", JsonValue.CreateStringValue(contactDeltaLink));
                root.SetNamedValue("calendarDeltaLinks", JsonObject.Parse(calendarDeltaLinks));
                return root.Stringify();
            }
        }

        public static string GetCalendarDeltaPath(string calendarId, SyncSettings settings)
        {
            return string.Format(
                "/v1.0/me/calendars/{0}/calendarView/delta?startDateTime={1}&endDateTime={2}&$select=id,subject,bodyPreview,start,end,location,isAllDay",
                Uri.EscapeDataString(calendarId),
                Uri.EscapeDataString(settings.GetStartDate().ToString("o")),
                Uri.EscapeDataString(settings.GetEndDate().ToString("o")));
        }
    }
}
