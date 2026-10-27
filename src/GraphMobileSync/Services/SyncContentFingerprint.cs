using System;
using System.Collections.Generic;
using System.Linq;
using Windows.ApplicationModel.Appointments;
using Windows.ApplicationModel.Contacts;

namespace GraphMobileSync.Services
{
    internal static class SyncContentFingerprint
    {
        public static string ForContact(Contact contact)
        {
            IEnumerable<string> emails = contact.Emails
                .Select(email => string.Format("{0}:{1}", email.Kind, email.Address ?? string.Empty))
                .OrderBy(value => value, StringComparer.Ordinal);
            IEnumerable<string> phones = contact.Phones
                .Select(phone => string.Format("{0}:{1}", phone.Kind, phone.Number ?? string.Empty))
                .OrderBy(value => value, StringComparer.Ordinal);

            return Join(
                contact.FirstName,
                contact.LastName,
                contact.DisplayNameOverride,
                string.Join(",", emails),
                string.Join(",", phones));
        }

        public static string ForAppointment(Appointment appointment)
        {
            return Join(
                appointment.Subject,
                appointment.Details,
                appointment.Location,
                appointment.StartTime.UtcDateTime.Ticks.ToString(),
                appointment.Duration.Ticks.ToString(),
                appointment.AllDay.ToString());
        }

        private static string Join(params string[] values)
        {
            return string.Concat(values.Select(value =>
            {
                string safeValue = value ?? string.Empty;
                return safeValue.Length + ":" + safeValue;
            }));
        }
    }
}
