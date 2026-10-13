# LiveSync Privacy Policy

**Effective date:** October 8, 2026

ZuneTracks ("ZuneTracks," "we," "us," or "our") provides the LiveSync application ("LiveSync" or the "app"). This Privacy Policy explains how LiveSync handles information when you use the app.

## Summary

LiveSync is a read-only Microsoft Graph synchronization app. It copies selected contact and calendar data from the Microsoft account that you choose to connect into LiveSync-owned contact lists and calendars on your Windows device. LiveSync does not operate a ZuneTracks server and does not send your contacts, calendar content, or credentials to ZuneTracks.

## Information LiveSync accesses

After you sign in and grant permission, LiveSync requests the following delegated Microsoft Graph permissions:

- `User.Read`, used as part of Microsoft Entra sign-in;
- `Contacts.Read`, to read your contacts;
- `Calendars.Read`, to read your calendars and calendar events; and
- `offline_access`, to keep a refresh token so that background synchronization can continue after you close the app.

Depending on the information available in your Microsoft account, LiveSync reads:

- contact identifiers, names, email addresses, mobile phone numbers, business phone numbers, and contact photos;
- calendar names and identifiers;
- event identifiers, subjects, body previews, start and end times, locations, and all-day status; and
- synchronization preferences, locally generated identifiers that associate Graph items with device items, Graph delta links, and a limited recent synchronization-status history.

LiveSync does not request permission to create, edit, or delete data in your Microsoft account. It does not read email, files, messages, or device location.

## How LiveSync uses information

LiveSync uses the information it accesses solely to:

- sign you in through Microsoft Entra;
- copy selected Microsoft Graph contacts into the LiveSync-owned People list on your device;
- copy selected Microsoft Graph calendar events into LiveSync-owned calendars on your device;
- update or remove the local copies when the corresponding Graph data changes or is removed;
- detect changes through Microsoft Graph delta queries; and
- show synchronization progress, status, and error information in the app.

LiveSync may check Microsoft Graph for changes in the background, subject to Windows scheduling, battery, data-saver, and permission settings.

## Where information is stored

LiveSync stores synchronization mappings, delta checkpoints, downloaded contact photos, preferences, and limited diagnostic history in the app's local storage on your device. A Microsoft Graph refresh token is stored in the Windows Credential Locker, an operating-system-managed credential store.

Contacts and calendar events that LiveSync copies are stored in LiveSync-owned lists and calendars in the Windows People and Calendar stores. These local copies may be visible through the Windows People and Calendar apps.

LiveSync communicates with Microsoft Entra and Microsoft Graph to authenticate you and retrieve data. LiveSync does not transmit this information to a ZuneTracks-operated server, and ZuneTracks does not retain a copy of your contacts, calendar data, or refresh token.

Your device, Windows, Microsoft account, and any other apps or account services that access the Windows People or Calendar stores may process data under their own privacy policies. This Privacy Policy covers LiveSync only.

## Sharing and sale of information

ZuneTracks does not sell, rent, trade, or use your personal information for advertising. LiveSync does not include third-party advertising, analytics, or tracking SDKs.

Information is sent to Microsoft only as necessary to use Microsoft Entra sign-in and Microsoft Graph. Microsoft's handling of information is governed by the [Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement).

## Retention and your choices

LiveSync retains app-local synchronization state for as long as it remains installed or until the operating system removes it. The app updates and deletes the local copies it manages during normal synchronization when corresponding Microsoft Graph items are changed or removed.

To stop future access to your Microsoft account, revoke LiveSync's permission in your Microsoft account or Microsoft Entra ID account's application-permissions settings. You can also uninstall LiveSync. If LiveSync-created People or Calendar entries remain visible after uninstalling, remove the entries from their LiveSync-owned list or calendar through the Windows People or Calendar app, where supported by the device.

Because ZuneTracks does not maintain a server-side copy of LiveSync data, requests to access, correct, or delete app-local data generally must be completed on your device or through the Microsoft account and Windows controls that manage that data.

## Security

LiveSync uses Microsoft Entra's authorization flow and stores its refresh token in Windows Credential Locker. No method of transmission or storage is completely secure, and you should protect access to your device and Microsoft account.

## Children's privacy

LiveSync is not directed to children. Do not use LiveSync if you are not permitted to authorize access to the Microsoft account you connect.

## Changes to this policy

We may update this Privacy Policy when LiveSync changes or when legal requirements change. We will post the updated policy with a revised effective date. Your continued use of LiveSync after an update takes effect means you acknowledge the updated policy.

## Contact us

For questions about this Privacy Policy or LiveSync's privacy practices, contact ZuneTracks at [app.support@ndtech.dev](mailto:app.support@ndtech.dev).
