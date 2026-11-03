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
        private async Task<int> SyncContactsAsync(
            ContactList contactList,
            IReadOnlyList<Contact> existingContacts,
            IList<JsonObject> graphContacts,
            SyncMappings mappings,
            SyncSummary summary,
            IProgress<SyncProgress> progress,
            int totalItems,
            int completedItems)
        {
            Dictionary<string, Contact> existingById = existingContacts.ToDictionary(contact => contact.Id);
            Dictionary<string, List<Contact>> existingByFingerprint =
                GroupContactsByFingerprint(existingContacts);
            HashSet<string> graphContactIds = new HashSet<string>();
            HashSet<string> claimedContactIds = new HashSet<string>();
            HashSet<string> mappedContactIds = new HashSet<string>(
                mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));
            Dictionary<string, Contact> pendingContacts = new Dictionary<string, Contact>();

            foreach (JsonObject graphContact in graphContacts)
            {
                string graphContactId = JsonValueReader.GetString(graphContact, "id", string.Empty);

                if (string.IsNullOrEmpty(graphContactId))
                {
                    continue;
                }

                if (pendingContacts.ContainsKey(graphContactId))
                {
                    await ResolvePendingContactIdsAsync(
                        contactList,
                        pendingContacts,
                        mappings,
                        claimedContactIds);
                }

                graphContactIds.Add(graphContactId);
                Contact localContact = FindExistingContact(
                    graphContactId,
                    graphContact,
                    mappings,
                    existingById,
                    existingByFingerprint,
                    claimedContactIds,
                    mappedContactIds);
                localContact = await UpsertContactAsync(
                    contactList,
                    localContact,
                    graphContactId,
                    graphContact,
                    mappings,
                    summary);

                if (string.IsNullOrEmpty(localContact.Id))
                {
                    pendingContacts[graphContactId] = localContact;
                }
                else
                {
                    RecordContactMapping(mappings, graphContactId, localContact);
                    claimedContactIds.Add(localContact.Id);
                }

                summary.Contacts++;
                completedItems++;
                ReportSyncProgress(progress, "Syncing contacts...", completedItems, totalItems);
            }

            await ResolvePendingContactIdsAsync(
                contactList,
                pendingContacts,
                mappings,
                claimedContactIds);
            HashSet<string> deletedContactIds = new HashSet<string>();

            foreach (KeyValuePair<string, string> mapping in mappings.ContactIds.ToList())
            {
                if (!graphContactIds.Contains(mapping.Key))
                {
                    Contact removedContact;

                    if (existingById.TryGetValue(mapping.Value, out removedContact))
                    {
                        await contactList.DeleteContactAsync(removedContact);
                        deletedContactIds.Add(removedContact.Id);
                    }

                    await DeleteContactPhotoAsync(mapping.Key);
                    mappings.RemoveContact(mapping.Key);
                }
            }

            summary.DuplicateContactsRemoved += await RemoveUnmappedContactsAsync(
                contactList,
                existingContacts,
                mappings,
                deletedContactIds);

            return completedItems;
        }

        private async Task ApplyContactDeltaAsync(
            IList<JsonObject> contactChanges,
            SyncMappings mappings,
            SyncSummary summary)
        {
            ContactStore contactStore = await ContactManager.RequestStoreAsync(
                ContactStoreAccessType.AppContactsReadWrite);

            if (contactStore == null)
            {
                throw new InvalidOperationException(
                    "Contact access was not granted. Enable contact access for LiveSync and try again.");
            }

            string mappedContactListId = mappings.ContactListId;
            ContactList contactList = await GetContactListAsync(
                contactStore,
                mappings,
                summary.CleanupPlan);
            IReadOnlyList<Contact> existingContacts = await GetContactsInListAsync(contactList);
            Dictionary<string, Contact> existingById = existingContacts.ToDictionary(contact => contact.Id);
            Dictionary<string, List<Contact>> existingByFingerprint =
                GroupContactsByFingerprint(existingContacts);
            HashSet<string> claimedContactIds = new HashSet<string>();
            HashSet<string> deletedContactIds = new HashSet<string>();
            HashSet<string> mappedContactIds = new HashSet<string>(
                mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));
            Dictionary<string, Contact> pendingContacts = new Dictionary<string, Contact>();

            foreach (JsonObject graphContact in contactChanges)
            {
                string graphContactId = JsonValueReader.GetString(graphContact, "id", string.Empty);

                if (string.IsNullOrEmpty(graphContactId))
                {
                    continue;
                }

                if (pendingContacts.ContainsKey(graphContactId))
                {
                    // Delta pages can repeat an item. Resolve the earlier save first so the repeat
                    // updates that contact instead of creating a second one.
                    AddResolvedContacts(
                        existingById,
                        await ResolvePendingContactIdsAsync(
                            contactList,
                            pendingContacts,
                            mappings,
                            claimedContactIds));
                }

                string localContactId;

                if (IsGraphItemRemoved(graphContact))
                {
                    if (mappings.ContactIds.TryGetValue(graphContactId, out localContactId))
                    {
                        Contact localContact;

                        if (existingById.TryGetValue(localContactId, out localContact))
                        {
                            await contactList.DeleteContactAsync(localContact);
                            existingById.Remove(localContactId);
                            deletedContactIds.Add(localContactId);
                        }

                        await DeleteContactPhotoAsync(graphContactId);
                        mappings.RemoveContact(graphContactId);
                    }

                    summary.Contacts++;
                    continue;
                }

                Contact updatedContact = FindExistingContact(
                    graphContactId,
                    graphContact,
                    mappings,
                    existingById,
                    existingByFingerprint,
                    claimedContactIds,
                    mappedContactIds);
                updatedContact = await UpsertContactAsync(
                    contactList,
                    updatedContact,
                    graphContactId,
                    graphContact,
                    mappings,
                    summary);

                if (string.IsNullOrEmpty(updatedContact.Id))
                {
                    pendingContacts[graphContactId] = updatedContact;
                }
                else
                {
                    RecordContactMapping(mappings, graphContactId, updatedContact);
                    claimedContactIds.Add(updatedContact.Id);
                    existingById[updatedContact.Id] = updatedContact;
                }

                summary.Contacts++;
            }

            AddResolvedContacts(
                existingById,
                await ResolvePendingContactIdsAsync(
                    contactList,
                    pendingContacts,
                    mappings,
                    claimedContactIds));

            // A delta response only describes changed contacts, so it can never prove that an
            // unreferenced contact is stale the way a full sync can. The mapping table can, because it
            // lists every contact LiveSync owns in this list. Only sweep when that table is populated
            // and was built against this exact list: after a mapping reset, or when the canonical list
            // changed, every contact would look unmapped and a delta sync would delete contacts it is
            // not going to re-create.
            if (mappings.ContactIds.Count > 0 && mappedContactListId == contactList.Id)
            {
                summary.DuplicateContactsRemoved += await RemoveUnmappedContactsAsync(
                    contactList,
                    existingContacts,
                    mappings,
                    deletedContactIds);
            }
        }

        private static async Task<ContactList> GetContactListAsync(
            ContactStore contactStore,
            SyncMappings mappings,
            SyncCleanupPlan cleanupPlan)
        {
            IReadOnlyList<ContactList> lists = await contactStore.FindContactListsAsync();

            // The store is opened with AppContactsReadWrite, so every list returned here is already
            // owned by LiveSync. Matching on display name alone stranded app-owned lists saved under
            // any other name: they were never treated as duplicates, so their contacts stayed visible
            // in the People app and kept showing up in its "Link contacts" picker.
            ContactList contactList = lists.FirstOrDefault(list => list.Id == mappings.ContactListId)
                ?? lists.FirstOrDefault(list => list.DisplayName == ContactListName)
                ?? lists.FirstOrDefault()
                ?? await contactStore.CreateContactListAsync(ContactListName);

            if (contactList.DisplayName != ContactListName || contactList.IsHidden)
            {
                // Un-hide as well as rename: a container hidden by an earlier cleanup pass must never
                // stay hidden once it is the canonical list, or synced contacts would be invisible.
                contactList.DisplayName = ContactListName;
                contactList.IsHidden = false;
                await contactList.SaveAsync();
            }

            mappings.ContactListId = contactList.Id;

            foreach (ContactList duplicateList in lists.Where(list => list.Id != contactList.Id))
            {
                cleanupPlan.ContactListIds.Add(duplicateList.Id);
            }

            return contactList;
        }

        private static Dictionary<string, List<Contact>> GroupContactsByFingerprint(
            IEnumerable<Contact> contacts)
        {
            Dictionary<string, List<Contact>> contactsByFingerprint =
                new Dictionary<string, List<Contact>>();

            foreach (Contact contact in contacts)
            {
                string fingerprint = GetContactFingerprint(contact);

                if (string.IsNullOrEmpty(fingerprint))
                {
                    continue;
                }

                List<Contact> contactsWithFingerprint;

                if (!contactsByFingerprint.TryGetValue(fingerprint, out contactsWithFingerprint))
                {
                    contactsWithFingerprint = new List<Contact>();
                    contactsByFingerprint[fingerprint] = contactsWithFingerprint;
                }

                contactsWithFingerprint.Add(contact);
            }

            return contactsByFingerprint;
        }

        private static Contact FindExistingContact(
            string graphContactId,
            JsonObject graphContact,
            SyncMappings mappings,
            IDictionary<string, Contact> existingById,
            IDictionary<string, List<Contact>> existingByFingerprint,
            ISet<string> claimedContactIds,
            ISet<string> mappedContactIds)
        {
            string mappedContactId = string.Empty;
            Contact mappedContact;

            if (mappings.ContactIds.TryGetValue(graphContactId, out mappedContactId)
                && existingById.TryGetValue(mappedContactId, out mappedContact)
                && !claimedContactIds.Contains(mappedContact.Id))
            {
                return mappedContact;
            }

            List<Contact> contactsWithFingerprint;

            if (!existingByFingerprint.TryGetValue(
                    GetContactFingerprint(graphContact),
                    out contactsWithFingerprint))
            {
                return null;
            }

            foreach (Contact contact in contactsWithFingerprint)
            {
                if (claimedContactIds.Contains(contact.Id))
                {
                    continue;
                }

                if (mappedContactIds.Contains(contact.Id)
                    && !string.Equals(contact.Id, mappedContactId, StringComparison.Ordinal))
                {
                    continue;
                }

                return contact;
            }

            return null;
        }

        private static async Task<IReadOnlyList<Contact>> GetContactsInListAsync(ContactList contactList)
        {
            // ContactStore.FindContactsAsync(string) is a full-text search over contact fields, not a
            // lookup by contact list ID. Passing a list ID to it always matched zero contacts, so the
            // sync treated every list as empty: nothing was ever reused and no stale entry was ever
            // deleted. A ContactReader is the only way to enumerate the contacts in one list.
            List<Contact> contacts = new List<Contact>();
            ContactReader reader = contactList.GetContactReader();

            while (true)
            {
                ContactBatch batch = await reader.ReadBatchAsync();

                if (batch.Status != ContactBatchStatus.Success || batch.Contacts.Count == 0)
                {
                    break;
                }

                contacts.AddRange(batch.Contacts);
            }

            return contacts;
        }

        private async Task<Contact> UpsertContactAsync(
            ContactList contactList,
            Contact localContact,
            string graphContactId,
            JsonObject graphContact,
            SyncMappings mappings,
            SyncSummary summary)
        {
            string changeKey = JsonValueReader.GetString(graphContact, "changeKey", string.Empty);
            bool sameGraphRevision = false;

            if (localContact != null && !string.IsNullOrEmpty(changeKey))
            {
                string mappedContactId;
                string storedChangeKey;
                string storedFingerprint;

                sameGraphRevision =
                    mappings.ContactIds.TryGetValue(graphContactId, out mappedContactId)
                    && string.Equals(mappedContactId, localContact.Id, StringComparison.Ordinal)
                    && mappings.ContactChangeKeys.TryGetValue(graphContactId, out storedChangeKey)
                    && string.Equals(storedChangeKey, changeKey, StringComparison.Ordinal);

                if (sameGraphRevision
                    && mappings.ContactFingerprints.TryGetValue(graphContactId, out storedFingerprint)
                    && storedFingerprint == SyncContentFingerprint.ForContact(localContact))
                {
                    summary.ContactsUnchanged++;
                    return localContact;
                }
            }

            if (localContact == null)
            {
                localContact = new Contact();
            }

            // An unchanged changeKey means the Graph photo has not changed either, so a locally
            // edited contact can be restored from the cached photo without another download.
            ContactPhotoResult photoResult = await SaveContactAsync(
                contactList,
                localContact,
                graphContactId,
                graphContact,
                sameGraphRevision);
            RecordContactPhotoResult(summary, photoResult);

            if (string.IsNullOrEmpty(changeKey))
            {
                mappings.ContactChangeKeys.Remove(graphContactId);
            }
            else
            {
                mappings.ContactChangeKeys[graphContactId] = changeKey;
            }

            return localContact;
        }

        private static void RecordContactMapping(
            SyncMappings mappings,
            string graphContactId,
            Contact contact)
        {
            mappings.ContactIds[graphContactId] = contact.Id;
            mappings.ContactFingerprints[graphContactId] = SyncContentFingerprint.ForContact(contact);
        }

        private static void AddResolvedContacts(
            IDictionary<string, Contact> existingById,
            IEnumerable<Contact> resolvedContacts)
        {
            foreach (Contact contact in resolvedContacts)
            {
                existingById[contact.Id] = contact;
            }
        }

        // Some Windows 10 Mobile builds leave Contact.Id empty after saving a new contact. Rather
        // than re-reading the whole list after every such save, those contacts are resolved together
        // with a single read once the batch has been written.
        private static async Task<List<Contact>> ResolvePendingContactIdsAsync(
            ContactList contactList,
            IDictionary<string, Contact> pendingContacts,
            SyncMappings mappings,
            ISet<string> claimedContactIds)
        {
            List<Contact> resolvedContacts = new List<Contact>();

            if (pendingContacts.Count == 0)
            {
                return resolvedContacts;
            }

            IReadOnlyList<Contact> savedContacts = await GetContactsInListAsync(contactList);
            HashSet<string> unavailableIds = new HashSet<string>(claimedContactIds);
            unavailableIds.UnionWith(mappings.ContactIds.Values.Where(id => !string.IsNullOrEmpty(id)));
            List<Contact> candidates = savedContacts
                .Where(contact => !unavailableIds.Contains(contact.Id))
                .ToList();

            foreach (KeyValuePair<string, Contact> pendingContact in pendingContacts)
            {
                string contentFingerprint = SyncContentFingerprint.ForContact(pendingContact.Value);
                string identityFingerprint = GetContactFingerprint(pendingContact.Value);
                Contact savedContact =
                    FindUniqueContact(
                        candidates,
                        contact => SyncContentFingerprint.ForContact(contact) == contentFingerprint)
                    ?? FindUniqueContact(
                        candidates,
                        contact => GetContactFingerprint(contact) == identityFingerprint);

                if (savedContact == null)
                {
                    throw new InvalidOperationException("LiveSync could not identify a saved contact.");
                }

                candidates.Remove(savedContact);
                mappings.ContactIds[pendingContact.Key] = savedContact.Id;
                mappings.ContactFingerprints[pendingContact.Key] = contentFingerprint;
                claimedContactIds.Add(savedContact.Id);
                resolvedContacts.Add(savedContact);
            }

            pendingContacts.Clear();
            return resolvedContacts;
        }

        private static Contact FindUniqueContact(IEnumerable<Contact> contacts, Func<Contact, bool> predicate)
        {
            List<Contact> matches = contacts.Where(predicate).Take(2).ToList();
            return matches.Count == 1 ? matches[0] : null;
        }

        private static void PopulateContact(Contact contact, JsonObject graphContact)
        {
            contact.FirstName = JsonValueReader.GetString(graphContact, "givenName", string.Empty);
            contact.LastName = JsonValueReader.GetString(graphContact, "surname", string.Empty);
            contact.Emails.Clear();
            contact.Phones.Clear();
            JsonArray emailAddresses = JsonValueReader.GetArray(graphContact, "emailAddresses");

            foreach (IJsonValue emailAddress in emailAddresses)
            {
                if (emailAddress.ValueType != JsonValueType.Object)
                {
                    continue;
                }

                string address = JsonValueReader.GetString(emailAddress.GetObject(), "address", string.Empty);

                if (!string.IsNullOrWhiteSpace(address))
                {
                    contact.Emails.Add(new ContactEmail { Address = address, Kind = ContactEmailKind.Personal });
                }
            }

            AddPhoneNumber(
                contact,
                JsonValueReader.GetString(graphContact, "mobilePhone", string.Empty),
                ContactPhoneKind.Mobile);
            JsonArray businessPhones = JsonValueReader.GetArray(graphContact, "businessPhones");

            foreach (IJsonValue businessPhone in businessPhones)
            {
                if (businessPhone.ValueType == JsonValueType.String)
                {
                    AddPhoneNumber(contact, businessPhone.GetString(), ContactPhoneKind.Work);
                }
            }
        }

        private async Task<ContactPhotoResult> SaveContactAsync(
            ContactList contactList,
            Contact contact,
            string graphContactId,
            JsonObject graphContact,
            bool reuseCachedPhoto)
        {
            PopulateContact(contact, graphContact);
            ContactPhotoResult photoResult = ContactPhotoResult.Reused;
            StorageFile photoFile = reuseCachedPhoto
                ? await GetCachedContactPhotoAsync(graphContactId)
                : null;

            if (photoFile == null)
            {
                byte[] photoBytes = await _graphClient.GetContactPhotoAsync(graphContactId);

                if (photoBytes == null)
                {
                    await contactList.SaveContactAsync(contact);
                    return ContactPhotoResult.NoPhoto;
                }

                if (photoBytes.Length == 0)
                {
                    throw new InvalidOperationException(
                        "Microsoft Graph returned an empty contact photo.");
                }

                photoFile = await SaveContactPhotoAsync(graphContactId, photoBytes);
                photoResult = ContactPhotoResult.Downloaded;
            }

            IRandomAccessStreamReference photoReference =
                RandomAccessStreamReference.CreateFromFile(photoFile);

            // Windows generates SmallDisplayPicture, LargeDisplayPicture and Thumbnail from
            // SourceDisplayPicture, and the People app renders those generated images. Setting
            // Thumbnail on its own leaves the display pictures empty, so the photo never appears.
            contact.SourceDisplayPicture = photoReference;
            contact.Thumbnail = photoReference;
            await contactList.SaveContactAsync(contact);
            return photoResult;
        }

        private static void RecordContactPhotoResult(SyncSummary summary, ContactPhotoResult photoResult)
        {
            switch (photoResult)
            {
                case ContactPhotoResult.NoPhoto:
                    summary.ContactsWithoutPhotos++;
                    break;
                case ContactPhotoResult.Downloaded:
                    summary.ContactPhotosDownloaded++;
                    break;
                case ContactPhotoResult.Reused:
                    summary.ContactPhotosReused++;
                    break;
            }
        }

        private static async Task<StorageFile> GetCachedContactPhotoAsync(string graphContactId)
        {
            StorageFolder photoFolder = await ApplicationData.Current.LocalFolder.TryGetItemAsync(
                ContactPhotoFolderName) as StorageFolder;

            if (photoFolder == null)
            {
                return null;
            }

            return await photoFolder.TryGetItemAsync(GetContactPhotoFileName(graphContactId)) as StorageFile;
        }

        private static async Task<StorageFile> SaveContactPhotoAsync(
            string graphContactId,
            byte[] photoBytes)
        {
            StorageFolder photoFolder = await ApplicationData.Current.LocalFolder.CreateFolderAsync(
                ContactPhotoFolderName,
                CreationCollisionOption.OpenIfExists);
            StorageFile photoFile = await photoFolder.CreateFileAsync(
                GetContactPhotoFileName(graphContactId),
                CreationCollisionOption.ReplaceExisting);
            await FileIO.WriteBytesAsync(photoFile, photoBytes);
            return photoFile;
        }

        private static async Task DeleteContactPhotoAsync(string graphContactId)
        {
            try
            {
                StorageFolder photoFolder = await ApplicationData.Current.LocalFolder.GetFolderAsync(
                    ContactPhotoFolderName);
                StorageFile photoFile = await photoFolder.GetFileAsync(
                    GetContactPhotoFileName(graphContactId));
                await photoFile.DeleteAsync();
            }
            catch (System.IO.FileNotFoundException)
            {
            }
        }

        private static string GetContactPhotoFileName(string graphContactId)
        {
            HashAlgorithmProvider algorithm = HashAlgorithmProvider.OpenAlgorithm(HashAlgorithmNames.Sha256);
            IBuffer graphContactIdBuffer = CryptographicBuffer.ConvertStringToBinary(
                graphContactId,
                BinaryStringEncoding.Utf8);
            return CryptographicBuffer.EncodeToHexString(
                algorithm.HashData(graphContactIdBuffer)) + ".jpg";
        }

        private static void AddPhoneNumber(Contact contact, string number, ContactPhoneKind kind)
        {
            if (!string.IsNullOrWhiteSpace(number))
            {
                contact.Phones.Add(new ContactPhone { Number = number, Kind = kind });
            }
        }

        private static string GetContactFingerprint(JsonObject graphContact)
        {
            JsonArray emails = JsonValueReader.GetArray(graphContact, "emailAddresses");

            foreach (IJsonValue email in emails)
            {
                if (email.ValueType == JsonValueType.Object)
                {
                    string address = JsonValueReader.GetString(email.GetObject(), "address", string.Empty);

                    if (!string.IsNullOrEmpty(address))
                    {
                        return "email:" + address.ToLowerInvariant();
                    }
                }
            }

            string mobilePhone = JsonValueReader.GetString(graphContact, "mobilePhone", string.Empty);
            return !string.IsNullOrEmpty(mobilePhone)
                ? "phone:" + NormalizePhoneNumber(mobilePhone)
                : "name:" + JsonValueReader.GetString(graphContact, "givenName", string.Empty).ToLowerInvariant()
                    + "|" + JsonValueReader.GetString(graphContact, "surname", string.Empty).ToLowerInvariant();
        }

        private static string GetContactFingerprint(Contact contact)
        {
            ContactEmail email = contact.Emails.FirstOrDefault();

            if (email != null && !string.IsNullOrEmpty(email.Address))
            {
                return "email:" + email.Address.ToLowerInvariant();
            }

            ContactPhone phone = contact.Phones.FirstOrDefault();
            return phone != null && !string.IsNullOrEmpty(phone.Number)
                ? "phone:" + NormalizePhoneNumber(phone.Number)
                : "name:" + (contact.FirstName ?? string.Empty).ToLowerInvariant()
                    + "|" + (contact.LastName ?? string.Empty).ToLowerInvariant();
        }

        private static string NormalizePhoneNumber(string number)
        {
            return new string(number.Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();
        }
    }
}
