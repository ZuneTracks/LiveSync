using System;
using System.Collections.Generic;
using System.Linq;
using Windows.Storage;

namespace GraphMobileSync.Services
{
    internal sealed class SyncSettings
    {
        private const string BackgroundEnabledKey = "BackgroundSyncEnabled";
        private const string FailureNotificationsEnabledKey = "FailureNotificationsEnabled";
        private const string TwoWaySyncEnabledKey = "TwoWaySyncEnabled";
        private const string CalendarNamesKey = "SelectedCalendarNames";
        private const string EventWindowKey = "EventWindowMonths";

        private HashSet<string> _selectedCalendarNames;

        private string _selectedCalendarNamesSource;

        public bool BackgroundSyncEnabled { get; set; }

        public bool FailureNotificationsEnabled { get; set; }

        public bool TwoWaySyncEnabled { get; set; }

        public string CalendarNames { get; set; }

        public int EventWindowMonths { get; set; }

        public static SyncSettings Load()
        {
            ApplicationDataContainer settings = ApplicationData.Current.LocalSettings;
            return new SyncSettings
            {
                BackgroundSyncEnabled = GetBoolean(settings, BackgroundEnabledKey, true),
                FailureNotificationsEnabled = GetBoolean(settings, FailureNotificationsEnabledKey, false),
                TwoWaySyncEnabled = GetBoolean(settings, TwoWaySyncEnabledKey, false),
                CalendarNames = settings.Values[CalendarNamesKey] as string ?? string.Empty,
                EventWindowMonths = GetInt32(settings, EventWindowKey, 12)
            };
        }

        public void Save()
        {
            ApplicationDataContainer settings = ApplicationData.Current.LocalSettings;
            settings.Values[BackgroundEnabledKey] = BackgroundSyncEnabled;
            settings.Values[FailureNotificationsEnabledKey] = FailureNotificationsEnabled;
            settings.Values[TwoWaySyncEnabledKey] = TwoWaySyncEnabled;
            settings.Values[CalendarNamesKey] = CalendarNames ?? string.Empty;
            settings.Values[EventWindowKey] = EventWindowMonths;
        }

        public bool IncludesCalendar(string calendarName)
        {
            string calendarNames = CalendarNames ?? string.Empty;

            if (_selectedCalendarNames == null
                || !string.Equals(_selectedCalendarNamesSource, calendarNames, StringComparison.Ordinal))
            {
                _selectedCalendarNames = new HashSet<string>(
                    calendarNames
                        .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(name => name.Trim()),
                    StringComparer.OrdinalIgnoreCase);
                _selectedCalendarNamesSource = calendarNames;
            }

            return _selectedCalendarNames.Count == 0 || _selectedCalendarNames.Contains(calendarName);
        }

        public DateTimeOffset GetStartDate()
        {
            return DateTimeOffset.UtcNow.AddMonths(-Math.Min(EventWindowMonths, 6));
        }

        public DateTimeOffset GetEndDate()
        {
            return DateTimeOffset.UtcNow.AddMonths(EventWindowMonths);
        }

        private static bool GetBoolean(ApplicationDataContainer settings, string key, bool defaultValue)
        {
            object value;
            return settings.Values.TryGetValue(key, out value) && value is bool
                ? (bool)value
                : defaultValue;
        }

        private static int GetInt32(ApplicationDataContainer settings, string key, int defaultValue)
        {
            object value;
            return settings.Values.TryGetValue(key, out value) && value is int
                ? (int)value
                : defaultValue;
        }
    }
}
