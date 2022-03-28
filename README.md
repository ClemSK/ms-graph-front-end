# ms-graph-front-end

## Summary

This is a front-end consuming the Microsoft Graph API.

The solution creates a web part which provides a user with an experience which displays the recordings information in an easy-to-understand format.

If a user selects to record their meeting can be viewed in a folder called ‘Recordings’.

---

## Features

- [x] A user can access sharepoint
- [x] A user can see the list of the last 10 meeting recordings
- [x] The files are displayed in an easy-to-understand format
- [ ] Webhook to update the Recordings folder as meetings are recorded
- [ ] In the user's OneDrive is a special folder called 'Recordings' - more on this below
- [ ] If a user selects to record their meeting the mp4 file is stored in the user's OneDrive

---

## The MVP

A web part which displays the last 10 meeting recordings, displaying the meetings by endDateTime in descending order.

The [Special folder](https://docs.microsoft.com/en-us/graph/api/drive-get-specialfolder?view=graph-rest-1.0&tabs=http) has this denomination because:

> Special folders provide simple aliases to access well-known folders in OneDrive without the need to look up the folder by path (which would require localization), or reference the folder with an ID. If a special folder is renamed or moved to another location within the drive, this syntax will continue to find that folder. Special folders are automatically created the first time an application attempts to write to one, if it doesn't already exist. If a user deletes one, it is recreated when written to again.

tldr: They are folders with specific aliases which are automatically written to and have more flexible localisation.

Special folders take the following aliases:

- Documents
- Photos
- Camera Roll
- App Root
- Music

While the recordings might go in documents, I didn't find a way to rename the special folders to 'Recordings'.

To automatically add files to a conventional folder would require a webhook, setting up a subscription as an alternative to regularly send recording data and files the 'Recordings' folder.

---

## If I had more time

1. I would complete a webhook feature to create a subscription so the user's folder would get updated as they create new recordings.

2. Pairing the callRecords with the 'callRecordingEventMessageDetail.callRecordingDisplayName' would provide the user with a descriptive title for the call record.

3. Additionally, calling the callRecords.mediaStream endpoint could provide further information about the recording to the user, potentially supplying the recording file or link to the stream.

---

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

> You will need to register a Microsoft developer account and create a Sharepoint tenant to deploy this webpart and view it.

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
