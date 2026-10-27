# LiveSync Privacy Policy

**Effective date:** October 26, 2026

ZuneTracks ("ZuneTracks," "we," "us," or "our") provides the LiveSync application ("LiveSync" or the "app"). This Privacy Policy explains how LiveSync handles information when you use the app.

## Summary

LiveSync's standard features are read-only: the app copies selected contact and calendar data from the Microsoft account that you choose to connect into LiveSync-owned contact lists and calendars on your Windows device.

LiveSync Pro is an optional paid Microsoft Store add-on. When you enable two-way synchronization, supported changes you make to LiveSync-managed contacts and events on your device can be sent to Microsoft Graph, and corresponding changes made in Microsoft Graph can be applied on your device.

LiveSync does not operate a ZuneTracks server and does not send your contacts, calendar content, credentials, or payment information to ZuneTracks.

## Information LiveSync accesses

After you sign in and grant permission, the standard LiveSync features request the following delegated Microsoft Graph permissions:

- `User.Read`, used as part of Microsoft Entra sign-in;
- `Contacts.Read`, to read your contacts;
- `Calendars.Read`, to read your calendars and calendar events; and
- `offline_access`, to keep a refresh token so that background synchronization can continue after you close the app.

When you purchase and enable LiveSync Pro two-way synchronization, LiveSync requests the following elevated delegated Microsoft Graph permissions instead of the corresponding read-only permissions:

- `Contacts.ReadWrite`, to read, create, update, and delete supported contacts in your Microsoft account; and
- `Calendars.ReadWrite`, to read, create, update, and delete supported calendar events in your Microsoft account.

`User.Read` and `offline_access` remain necessary for LiveSync Pro. Microsoft Entra may require you to sign in again and grant consent before the elevated permissions take effect.

LiveSync also checks the Microsoft Store license or entitlement needed to determine whether LiveSync Pro features are available. Microsoft processes the purchase. LiveSync does not receive or store your payment-card or other payment-account information.

Depending on the information available in your Microsoft account and the features you enable, LiveSync reads:

- contact identifiers, names, email addresses, mobile phone numbers, business phone numbers, and contact photos;
- calendar names and identifiers;
- event identifiers, subjects, body previews, start and end times, locations, and all-day status; and
- synchronization preferences, locally generated identifiers that associate Graph items with device items, Graph delta links, and a limited recent synchronization-status history.

When LiveSync Pro two-way synchronization is enabled, LiveSync may send to Microsoft Graph the supported contact or event fields that you create or edit in LiveSync-owned People or Calendar containers. It may also send a deletion when you delete a corresponding LiveSync-managed contact or event on your device. LiveSync Pro writes only items it creates, imports, or has linked to a LiveSync mapping; it is not intended to edit unrelated Graph contacts or calendar events.

The supported outbound contact fields are given name, surname, display name, email addresses, mobile phone, and work phone. The supported outbound event fields are subject, details, location, start/end times, and all-day status. Contact photos, recurrence, attendees, categories, response state, calendar creation/removal, and other fields are not uploaded by LiveSync Pro.

To detect supported edits made in the built-in People and Calendar apps, LiveSync uses Windows restricted `contactsSystem` and `appointmentsSystem` capabilities. Microsoft Store approval is required before a public Store release can use those capabilities.

LiveSync does not request or use permission to access email, files, messages, or device location.

## How LiveSync uses information

LiveSync uses the information it accesses solely to:

- sign you in through Microsoft Entra;
- copy selected Microsoft Graph contacts into the LiveSync-owned People list on your device;
- copy selected Microsoft Graph calendar events into LiveSync-owned calendars on your device;
- update or remove the local copies when the corresponding Graph data changes or is removed;
- detect changes through Microsoft Graph delta queries; and
- show synchronization progress, status, and error information in the app.

When LiveSync Pro two-way synchronization is enabled, LiveSync additionally uses the information it accesses to:

- verify your LiveSync Pro entitlement with Microsoft Store;
- identify supported changes to LiveSync-managed contacts and events made in the Windows People and Calendar apps;
- create, update, or delete the corresponding Microsoft Graph contacts and calendar events;
- prevent a change downloaded from Graph from being immediately uploaded again; and
- resolve an ambiguous same-interval local and Graph edit in favor of the Microsoft Graph version.

LiveSync may check Microsoft Graph for changes in the background, subject to Windows scheduling, battery, data-saver, and permission settings.

## Where information is stored

LiveSync stores synchronization mappings, delta checkpoints, downloaded contact photos, preferences, and limited diagnostic history in the app's local storage on your device. When LiveSync Pro two-way synchronization is enabled, the app also stores local synchronization state needed to identify supported local changes, associate them with Graph items, avoid synchronization loops, and record conflict or write-back results. This state can include local and Graph item identifiers, local content fingerprints, and synchronization status.

A Microsoft Graph refresh token is stored in the Windows Credential Locker, an operating-system-managed credential store. LiveSync may cache Microsoft Store license information needed to determine whether Pro features are available; it does not store payment-card or payment-account information.

Contacts and calendar events that LiveSync copies are stored in LiveSync-owned lists and calendars in the Windows People and Calendar stores. These local copies may be visible through the Windows People and Calendar apps.

LiveSync communicates with Microsoft Entra and Microsoft Graph to authenticate you and retrieve data. LiveSync does not transmit this information to a ZuneTracks-operated server, and ZuneTracks does not retain a copy of your contacts, calendar data, or refresh token.

Your device, Windows, Microsoft account, and any other apps or account services that access the Windows People or Calendar stores may process data under their own privacy policies. This Privacy Policy covers LiveSync only.

## Sharing and sale of information

ZuneTracks does not sell, rent, trade, or use your personal information for advertising. LiveSync does not include third-party advertising, analytics, or tracking SDKs.

Information is sent to Microsoft only as necessary to use Microsoft Entra sign-in, Microsoft Store licensing, and Microsoft Graph synchronization. Microsoft's handling of information is governed by the [Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement).

## Retention and your choices

LiveSync retains app-local synchronization state for as long as it remains installed or until the operating system removes it. The app updates and deletes the local copies it manages during normal synchronization when corresponding Microsoft Graph items are changed or removed. When Pro two-way synchronization is enabled, the app uses the related local synchronization state while processing supported changes and write-backs.

To stop two-way synchronization, turn off the Pro synchronization feature in the app. To stop future access to your Microsoft account, revoke LiveSync's permission in your Microsoft account or Microsoft Entra ID account's application-permissions settings. You can also uninstall LiveSync. If LiveSync-created People or Calendar entries remain visible after uninstalling, remove the entries from their LiveSync-owned list or calendar through the Windows People or Calendar app, where supported by the device.

Because ZuneTracks does not maintain a server-side copy of LiveSync data, requests to access, correct, or delete app-local data generally must be completed on your device or through the Microsoft account and Windows controls that manage that data.

## Security

LiveSync uses Microsoft Entra's authorization flow and stores its refresh token in Windows Credential Locker. No method of transmission or storage is completely secure, and you should protect access to your device and Microsoft account.

## Children's privacy

LiveSync is not directed to children. Do not use LiveSync if you are not permitted to authorize access to the Microsoft account you connect.

## Changes to this policy

We may update this Privacy Policy when LiveSync changes or when legal requirements change. We will post the updated policy with a revised effective date. Your continued use of LiveSync after an update takes effect means you acknowledge the updated policy.

## Contact us

For questions about this Privacy Policy or LiveSync's privacy practices, contact ZuneTracks at [app.support@ndtech.dev](mailto:app.support@ndtech.dev).
