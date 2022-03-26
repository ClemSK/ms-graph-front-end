# ms-graph-front-end

## Summary

This is a front-end consuming the Microsoft Graph API.

The solution creates a webpart which provides a user with an experience which displays the recordings information in an easy-to-understand format.

If a user selects to record their meeting the mp4 file is stored in that users OneDrive in a special folder called ‘Recordings’.

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

## Features

- A user can access OneDrive.
- In the user's OneDrive is a special folder called 'Recordings'.
- If a user selects to record their meeting the mp4 file is stored in the user's OneDrive.
- The files are displayed in an easy-to-understand format.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
