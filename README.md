# Governor Sharing

## Summary

SPFx WebPart shows documents which have been explicitly shared.

It does this by using the following steps:
- Issueing a Search Query (KQL) against the Graph API to retrieve documents where the managed property SharedWithUsersOWSUSER contains a value
- Iterate through the result of the search query to get the permissions per file (/permissions endpoint of driveItems on GraphAPI)
- Show the results in a ShimmeredDetailsList and the Pagination control for paging the results
- By selecting a document and clicking on the Sharing Settings button will open the Manage Access pane for further review of the sharing

Here is an example of document shared with an external user, notice the tooltip & icon in front of the document
![Example Image](screenshot.png)

Next, when clicking on the Sharing Settings, the Manage Access page tells you that a Sharing Link was created for an external user
![Example Image](screenshot2.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| Governor Sharing | Robin Meure MSFT                                        |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | October 27, 2023 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
