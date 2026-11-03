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
        public static async Task CleanupDuplicatesAsync(SyncCleanupPlan cleanupPlan)
        {
            if (cleanupPlan == null || !cleanupPlan.HasItems)
            {
                return;
            }

            await CleanupDuplicateContactListsAsync(cleanupPlan.ContactListIds);

            if (cleanupPlan.CalendarIds.Count == 0)
            {
                return;
            }

            AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(
                AppointmentStoreAccessType.AppCalendarsReadWrite);

            if (appointmentStore == null)
            {
                throw new InvalidOperationException("Calendar access was not granted. Duplicate calendars were not removed.");
            }

            IReadOnlyList<AppointmentCalendar> calendars = await appointmentStore.FindAppointmentCalendarsAsync();

            foreach (string calendarId in cleanupPlan.CalendarIds.Distinct())
            {
                AppointmentCalendar duplicateCalendar = calendars.FirstOrDefault(calendar => calendar.LocalId == calendarId);

                if (duplicateCalendar == null)
                {
                    continue;
                }

                await duplicateCalendar.DeleteAsync();
            }
        }

        public static async Task CleanupDuplicateContactListsAsync(
            IEnumerable<string> duplicateContactListIds)
        {
            if (duplicateContactListIds == null)
            {
                return;
            }

            ContactStore contactStore = await ContactManager.RequestStoreAsync(ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException("Contact access was not granted. Duplicate contacts were not removed.");
            }

            IReadOnlyList<ContactList> contactLists = await contactStore.FindContactListsAsync();

            foreach (string contactListId in duplicateContactListIds.Distinct())
            {
                ContactList duplicateList = contactLists.FirstOrDefault(list => list.Id == contactListId);

                if (duplicateList == null)
                {
                    continue;
                }

                IReadOnlyList<Contact> duplicateContacts = await GetContactsInListAsync(duplicateList);

                foreach (Contact contact in duplicateContacts)
                {
                    await duplicateList.DeleteContactAsync(contact);
                }

                await RemoveEmptyContactListAsync(duplicateList);
            }
        }

        private static async Task RemoveEmptyContactListAsync(ContactList contactList)
        {
            try
            {
                await contactList.DeleteAsync();
                return;
            }
            catch (Exception)
            {
                // Windows 10 Mobile 15063 does not let an app delete its own contact list container.
            }

            try
            {
                // Hiding the emptied container keeps it out of the People app, including the
                // "Link contacts" picker, which still lists app-owned lists that remain visible.
                contactList.IsHidden = true;
                await contactList.SaveAsync();
            }
            catch (Exception)
            {
            }
        }

        private static async Task<int> RemoveUnmappedContactsAsync(
            ContactList contactList,
            IEnumerable<Contact> existingContacts,
            SyncMappings mappings,
            HashSet<string> deletedContactIds)
        {
            ISet<string> mappedLocalContactIds = new HashSet<string>(
                mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));
            int removedContacts = 0;

            foreach (Contact existingContact in existingContacts)
            {
                // deletedContactIds keeps this sweep from deleting a contact that the caller already
                // removed in the same run. Those contacts are still in the snapshot taken before the
                // sync and their mapping is gone, so they would otherwise be deleted twice and
                // double-counted as duplicates.
                if (mappedLocalContactIds.Contains(existingContact.Id)
                    || deletedContactIds.Contains(existingContact.Id))
                {
                    continue;
                }

                await contactList.DeleteContactAsync(existingContact);
                deletedContactIds.Add(existingContact.Id);
                removedContacts++;
            }

            return removedContacts;
        }
    }
}
