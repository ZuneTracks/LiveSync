# LiveSync

Prototype UWP package for Windows 10 Mobile build 15063 or later. It signs in to Microsoft Graph, imports the signed-in user's contacts and available contact photos to an app-owned **Microsoft Graph** People list, and imports each Graph calendar to an app-owned **Microsoft Graph - _calendar name_** calendar. Those lists and calendars are visible in the built-in People and Calendar apps. During sync, LiveSync displays download status, then an item-level `synced / total` progress counter.

LiveSync targets Windows 10 SDK 10.0.15063.0 with the same 10.0.15063.0 minimum version. This matches the proven LiveDrive and UniTube Store configuration: release and Store builds explicitly set `UseDotNetNativeToolchain` to use the .NET Native 1.7 toolchain from `Microsoft.NETCore.UniversalWindowsPlatform` 6.2.14, instead of the incompatible in-box SDK fallback. Store ARM submissions must declare the Windows 10 Mobile-compatible .NET Native Framework and Runtime 1.7 dependencies. Keep the platform-version values at 15063 and retain this toolchain setting unless support for W10M 15063 is intentionally dropped.

LiveSync stores Graph contact, event, and calendar IDs alongside their device counterparts so upgrades and later delta syncs update existing items rather than create duplicates. On the first sync after upgrade, it migrates legacy app-owned data using unique email/phone and event-time matches. Contacts already saved in an app-owned list are enumerated with `ContactList.GetContactReader()` and `ContactReader.ReadBatchAsync()`. The `ContactStore.FindContactsAsync(String)` overload is a full-text search over contact fields and not a lookup by contact list ID, so passing a list ID to it matched zero contacts and made every list look empty: existing contacts were never reused, stale entries were never deleted, and each sync added another copy of every contact. After each complete sync, it removes unmapped, stale, and duplicate entries from its app-owned contact list and within the configured event window of its app-owned calendars. Delta syncs apply contacts that Microsoft Graph reports as removed and also sweep contacts that no mapping references, so orphans no longer wait for the next manual sync. That sweep is skipped when the mapping table is empty or was built against a different contact list, because a delta response only describes changed contacts and would otherwise delete contacts it is not going to re-create. It also automatically clears the contents of duplicate app-owned Microsoft Graph contact lists, retaining one canonical list. Because the contact store is opened for app-owned contacts only, every contact list it returns belongs to LiveSync, so all non-canonical lists are treated as duplicates regardless of their display name. Each emptied duplicate list is deleted, and when Windows 10 Mobile refuses to delete an app-created container it is hidden instead, so it no longer appears in the People app or its "Link contacts" picker. Background change checks perform this same contact-list cleanup rather than deferring it to the next manual sync. If duplicate LiveSync calendars are found, it asks before clearing their contents. This cleans incorrect legacy entries such as pre-timezone-fix events. Downloaded contact photos are stored in LiveSync's local app storage and assigned from durable file references. Each photo is written to `Contact.SourceDisplayPicture` as well as `Contact.Thumbnail`, because Windows generates the `SmallDisplayPicture` and `LargeDisplayPicture` images that the People app actually renders from `SourceDisplayPicture`; setting `Thumbnail` alone leaves contacts without a visible picture. After saving, LiveSync reads each contact back from the store to confirm the picture persisted. The completion status separately reports how many contact photos were downloaded, how many were saved to the People app, and how many contacts do not have a Graph photo. Windows 10 Mobile may refuse to delete app-created list/calendar containers through the 15063 APIs; calendar containers therefore remain as empty containers after cleanup.

Graph identity mappings and delta cursors are stored as files in LiveSync's local app storage rather than as `LocalSettings` values. This supports large contact/calendar collections without hitting the Windows application-settings value limit.

Calendar times are requested from Microsoft Graph in UTC, then converted to the Windows 10 Mobile device's current local timezone before saving. Change the phone's timezone in Settings if its displayed local time is incorrect.

After the first successful manual sync, LiveSync stores Microsoft Graph delta cursors and registers an in-process UWP timer background task. Windows schedules that task no more often than every 15 minutes, subject to battery saver, data-saver, and system scheduling policy. The task refreshes its Graph token from Windows Credential Locker and applies returned contact and calendar delta records directly to the app-owned People and Calendar data. It also prunes stale event mappings at most once a day. If Microsoft Graph returns `429 Too Many Requests`, LiveSync honors its `Retry-After` value and defers later checks until that window has elapsed. Its latest result is shown when LiveSync next opens. The app shows whether Windows has registered the task; **Check Graph changes now** immediately runs the same delta-check logic for troubleshooting without waiting for the OS scheduler.

LiveSync also restores the saved sign-in automatically whenever the app opens. The refresh token stays in Windows Credential Locker; the app silently exchanges it for a short-lived Graph access token. Interactive sign-in is needed again only when the saved token has been removed, revoked, or expires, or when the phone cannot reach Microsoft Entra to renew it.

The **About LiveSync** flyout includes an app description, a link to the ZuneTracks GitHub profile, and the installed package version read directly from the package manifest.

## Privacy

LiveSync's Store-ready privacy policy is available in [PRIVACY.md](PRIVACY.md). Publish this document at a stable, publicly accessible HTTPS URL before submitting LiveSync to Microsoft Store.

## Sync controls

- **Sync in the background** enables or disables the periodic Graph delta poll.
- **Notify me when background sync fails** is opt-in and raises a toast only when LiveSync needs attention.
- **Calendars to sync** accepts exact, comma-separated Graph calendar names; leave blank to sync all calendars.
- **Event window** controls the past/future date range included in calendar syncs.
- **Recent sync activity** records the last ten manual/background results. Select **Sign in / sign in again** if the history reports refresh-token or consent failure.

## Configure Microsoft Entra

1. Create a Microsoft Entra application registration for public/native clients.
2. Under **Authentication**, add the mobile and desktop redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient` and enable public client flows.
3. Add delegated Microsoft Graph permissions: `User.Read`, `Contacts.Read`, and `Calendars.Read`. Grant consent if the tenant requires it.
4. Set the client ID in the LiveSync solution, then launch the app and sign in with the user's Microsoft account.

For a ZuneTracks-distributed build, set `MicrosoftEntraClientId` in `src\GraphMobileSync\Configuration\LiveSyncConfiguration.cs` before building. Users only need to sign in with their own Microsoft account; they do not create or register an Entra application.

The app requests only read permissions from Graph. `Contacts.Read` also permits LiveSync to download each available contact photo; no additional Entra permission is required. It recreates only the contacts and appointments inside its own app-owned lists during every sync; it does not alter other device data.

## Build an ARM `.appx`

Open `GraphMobileSync.sln` in Visual Studio with the Windows 10 SDK 10.0.15063.0 installed, select **Release | ARM**, then use **Project > Publish > Create App Packages** and select sideloading.

From a Developer PowerShell with Visual Studio installed:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe' .\GraphMobileSync.sln /t:Rebuild /p:Configuration=Release /p:Platform=ARM /p:AppxBundle=Never /p:AppxPackage=true
```

The package is written below `src\GraphMobileSync\AppPackages`. The ready-to-test ARM package is:

```text
src\GraphMobileSync\AppPackages\GraphMobileSync_1.0.0.11_ARM_Test\GraphMobileSync_1.0.0.11_ARM.appx
```

It is signed with the accompanying `LiveSync-TestCertificate.cer`. Before sideloading that package to a developer-unlocked phone, trust the certificate in the device's Trusted People store. The test certificate expires after one year; it is intentionally for development only.
