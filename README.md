# LiveSync

Prototype UWP package for Windows 10 Mobile build 15063 or later. It signs in to Microsoft Graph, imports the signed-in user's contacts to an app-owned **Microsoft Graph** People list, and imports each Graph calendar to an app-owned **Microsoft Graph - _calendar name_** calendar. Those lists and calendars are visible in the built-in People and Calendar apps. During sync, LiveSync displays download status, then an item-level `synced / total` progress counter.

LiveSync stores Graph contact, event, and calendar IDs alongside their device counterparts so upgrades and later delta syncs update existing items rather than create duplicates. On the first sync after upgrade, it migrates legacy app-owned data using unique email/phone and event-time matches. After each complete sync, it also removes unmapped/stale entries from its app-owned contact list and within the configured event window of its app-owned calendars. This cleans incorrect legacy entries such as pre-timezone-fix events. If multiple old LiveSync lists/calendars are found, it asks before clearing only their duplicate contents. Windows 10 Mobile cannot delete those app-created list/calendar containers through the 15063 APIs, so they remain as empty containers after cleanup.

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

The app requests only read permissions from Graph. It recreates only the contacts and appointments inside its own app-owned lists during every sync; it does not alter other device data.

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
