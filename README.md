# Governor Sharing

## Summary

SPFx WebPart shows documents which have been explicitly shared within a SharePoint site or Team.

It does this by using the following steps:

- Issue a Search Query (KQL) against the Graph API to retrieve documents where the managed property `SharedWithUsersOWSUSER` contains a value.
- Iterate through the results of the search query to get the permissions (e.g., sharing information) for each file using the `/permissions` endpoint of `driveItems` on GraphAPI.
- Display the results in a `ShimmeredDetailsList` and use the `Pagination` control for paging through the results.
- By selecting a document and clicking on the Sharing Settings button, the Manage Access pane will open for further review of the sharing settings.
- The panel includes a Manage Access pivot and an Activity pivot.
- Enable debug mode as a webpart property to see the query that the Graph Search API is triggering or by adding query string `debug=true`.

Here is an example with a list of shared documents, with a clear distinction when they are shared with external users (notice the tooltip & icon in front of the document):
![Example Image](/assets/Screenshot_v2.png)

When you want to know more about the sharing settings of a particular document, you can click on the view icon of the document. This will open up the Side Panel which has the Manage Access component rendered through an iFrame, indicating that a sharing link was created for the external user.
![Example Image](/assets/screenshot2_v2.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to the [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Minimal Path to Awesome

- Clone this repository
- Move to the right solution folder
- In the command line run:
  - `npm install`
  - `gulp serve` or `npm run serve`

## Solution

| Solution         | Author(s)           |
| ---------------- | ------------------- |
| Governor Sharing | Robin Meure MSFT    |
| Governor Sharing | Ahamed Fazil Buhari |

## Version history

| Version | Date              | Comments                                                |
| ------- | ----------------- | ------------------------------------------------------- |
| 1.0     | October 27, 2023  | Initial release                                         |
| 2.0     | November 11, 2024 | SPFx 1.20 upgraded, React hooks, Server filter & Search |

## Deployment Overview

- [SharePoint App Deployment](#sharepoint-app-deployment)
  - [Prerequisites](#prerequisites-1)
  - [Step 1 - Add the app to the SharePoint App catalog](#step-1---add-the-app-to-the-sharepoint-app-catalog)
  - [Step 2 - Provide API consent](#step-2---provide-api-consent)
  - [Step 3 - Adding the app to a SharePoint site](#step-3---adding-the-app-to-a-sharepoint-site)
- [Teams App Deployment](#teams-app-deployment)
  - [Prerequisites](#prerequisites-2)
  - [Step 1 - Add the app to Teams App Catalog](#step-1---add-the-app-to-teams-app-catalog)
  - [Step 2 - Add the app to a Teams tab](#step-2---add-the-app-to-a-teams-tab)

# SharePoint App Deployment

## Prerequisites

- Ensure you have a copy of the solution package file with a `.sppkg` extension.
- The user deploying the app must have SharePoint Administrator or Global Administrator permissions in Microsoft 365.
- The same user must approve and provide consent for the required API permissions to call the Graph Search endpoint.

## Step 1 - Add the app to the SharePoint App catalog

Follow the steps below to add the app to the SharePoint App catalog:

- Go to [More features](https://go.microsoft.com/fwlink/?linkid=2185077) in the SharePoint admin center, and sign in with an account that has the SharePoint Administrator or Global Administrator role for your organization.
- Under Apps, select Open. If you didn’t have an app catalog before, it might take a few minutes to load.

<img src="assets/SharePoint_Admin_Center_Manage_apps.png" width="1000"/>

- On the Manage apps page, click **Upload**, and browse to the location of the app package. The package file should have a `.sppkg` extension.
- Select **Enable this app and add it to all sites**. This will automatically add the app to the sites, so that site owners will not need to do it themselves. Uncheck the box **Add to Teams**. If you want to add the App to Teams, you need to follow these instructions. Click **Enable app** at the bottom of the side panel.

<img src="assets/SharePoint_Admin_Center_Enable_app.png" width="300"/>

## Step 2 - Provide API consent

After the API is enabled, you will need to provide consent. For this step, you need the Global Administrator role.
You will provide delegated permissions that will allow the application to act on a user's behalf. The application will never be able to access anything the signed-in user themselves couldn't access. To learn more about delegated permissions, see: [Types of permissions](https://learn.microsoft.com/en-us/entra/identity-platform/permissions-consent-overview#types-of-permissions).

- Click on **Go to the API access** page.

<img src="assets/SharePoint_Admin_Center_API_Consent.png" width="300"/>

- Click **Approve** to provide consent.

<img src="assets/SharePoint_Admin_Center_API_Consent_Approve.png" width="600"/>

## Step 3 - Adding the app to a SharePoint site

- On the site where you want to use the app, go to a page and open it for editing or create a new page for this purpose.
- Click on the **“+”** to add a new web part and search for “Governor sharing”. Click on it to add it to the page.

<img src="assets/Govenor_Sharing_AddtoSharePointSite.png" width="600"/>

- The web part should now be added to your page.
- Save or Republish the page to see the changes applied.

# Teams App Deployment

For the Teams App deployment, the app needs to be deployed to the SharePoint App Catalog first (Step 1 and Step 2).

## Prerequisites

- A copy of the Teams Apps solution [package](/assets/governorsharing_teamspackage.zip)
- The user deploying the app must be a Teams Administrator or Global Administrator in Microsoft 365.

## Step 1 - Add the app to Teams App Catalog

- Browse to the Manage Apps page in the Teams Admin Center: [Manage Apps](https://admin.teams.microsoft.com/policies/manage-apps)
- Click **Upload new App**, Click **Upload** and browse to the Teams app package location. The package file should have a `.zip` extension. After selecting the package, click **Open**. The app will be uploaded.

<img src="assets/Teams_Admin_Center_Manage_apps.png" width="500"/>

<img src="assets/Teams_Admin_Center_Manage_apps_Upload.png" width="500"/>

<img src="assets/Teams_Admin_Center_Manage_apps_Uploaded.png" width="500"/>

- You may need to adjust your Teams App policies to make the app available for your organization. For more information, see [Teams App Permission Policies](https://learn.microsoft.com/en-us/microsoftteams/teams-app-permission-policies).

## Step 2 - Add the app to a Teams tab

- Go to MS Teams and click on the **Apps** on the left bar to open the App store of Teams.
- On the left menu, choose **Built for your Org** option to prefilter the apps and select “Governor sharing”. Click **Add**.

<img src="assets/Govenor_Sharing_AddtoTeam.png" width="500"/>

- Click on **Add to a team**, choose a team and a channel where you want the app to be added and click **Set up a tab** on the bottom right of the pop-up window.

<img src="assets/Govenor_Sharing_AddtoTeamTab.png" width="500"/>

<img src="assets/Govenor_Sharing_AddtoTeam_SelectTeam.png" width="500"/>

- Click on **Save**

<img src="assets/Govenor_Sharing_AddtoTeam_Save.png" width="500"/>

- The app has been added to a Team. The settings panel on the right side can be closed.

<img src="assets/Govenor_Sharing_AddedtoTeam.png" width="500"/>

# Troubleshooting

If you face any other errors, you can enable the debugging mode from the configuration pane. When this is enabled, there is a lot more details being outputted to the console.

- In green, you see the search (KQL) query that is used to retrieve documents.
- In yellow, you see the search results.
- In blue, you see the transformation of combining the search results and the permission calls.

<img src="assets/debug.png" width="500"/>

## Known errors

**Issue:** We can't upload the app because there's already an app in the catalog with the same app ID. To upload a new app, change the app ID and try again. To update an existing app, go to the app details page.

**Solution:** Delete the app in the Teams Apps overview and re-add the package.

More information about deleting apps in Teams can be found here: [Delete custom apps from your organization's catalog](https://learn.microsoft.com/en-us/microsoftteams/teams-custom-app-policies-and-settings#delete-custom-apps-from-your-organizations-catalog).

<img src="https://pnptelemetry.azurewebsites.net/sp-dev-solutions/solutions/governorsharing" />

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
