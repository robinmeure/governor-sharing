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

# SharePoint App Deployment

## Prerequisites
- A copy of the solution .sppkg package.
- The user deploying an app must be a SharePoint Administrator or Global Administrator in Microsoft 365. 
- A Global Administrator need to be approved and provide consent for the API permissions.

## Step 1 - Add the app to the SharePoint App catalog

Follow the steps below to add the app to the SharePoint App catalog:

- Go to [More features](https://go.microsoft.com/fwlink/?linkid=2185077) in the SharePoint admin center, and sign in with an account that has the SharePoint Administrator or Global Administrator for your organization.
- Under Apps, select Open. If you didn’t have an app catalog before, it might take a few minutes to load.

![Example Image](screenshots/SharePoint_Admin_Center_Manage_apps.png)

- On the Manage apps page, select Upload, and browse to the app. The package file should have .sppkg extension.
- There should be a popup named Enable app: We recommend keeping the option “Enable this app and add it to all sites”. This will automatically add the app to the sites, so that site owners will not need to do it themselves. This will minimize the adoption efforts for the app. Unchecked the box “Add to Teams”. If you want to add the App to Teams you need to follow these instructions. Click Enable app at the bottom of the side panel.

![Example Image](screenshots/SharePoint_Admin_Center_Enable_app.png)

## Step 2 - Provide API consent

After the API is Enable you will need to provide consent. For this step you need the Global Administrator role.
You will provide delegated permissions that will allow the application to act on a user's behalf. The application will never be able to access anything the signed in user themselves couldn't access. To learn more about delegated permissions see: https://learn.microsoft.com/en-us/entra/identity-platform/permissions-consent-overview#types-of-permissions

- Click on Go to the API access page.

![Example Image](screenshots/SharePoint_Admin_Center_API_Consent.png)

- Click Approve to provide consent.

![Example Image](screenshots/SharePoint_Admin_Center_API_Consent_Approve.png)

## Step 3 - Adding the app to a SharePoint site

- On the site where you want to use the app go to a page and open it for editing or create a new page for this purpose.
- Click on the “+” to add a new web part and search for “sharing”. The app with the name “Governor Sharing” will be available. Click on it to add it to the page.

![Example Image](screenshots/Govenor_Sharing_AddtoSharePointSite.png)

![Example Image](screenshots/Govenor_Sharing_SharedItemsExample.png)

- Save or Republish the page to see the changes applied.

# Teams App Deployment

For the Teams App deployment, the app needs to be deployed to the SharePoint App Catalog first.

## Prerequisites
- A copy of the Teams Apps solution .zip package.
- The user deploying the app must be a Teams Administrator or Global Administrator in Microsoft 365.









# Troubleshooting 

Debug mode??



