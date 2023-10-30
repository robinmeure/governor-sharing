import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, spGet, SPQueryable, SPFx as spSPFx } from '@pnp/sp'
import { ISearchResultExtended } from "./ISearchResultExtended";

import { ISharingResult } from "./ISharingResult";
import { Utils } from "./Utils";
import "@pnp/sp/webs";
import "@pnp/sp/search";
import "@pnp/sp/sharing";
import "@pnp/sp/presets/all";
import { Caching } from "@pnp/queryable";
import { Logger, LogLevel,PnPLogging  } from "@pnp/logging";
import { graphfi, GraphFI, graphGet, GraphQueryable, SPFx as graphSPFx } from "@pnp/graph";
import "@pnp/graph/search";
import "@pnp/graph/users";
import "@pnp/graph/onedrive";
import "@pnp/graph/batching";
import { IFacepilePersona } from "@fluentui/react";
import { search } from "@microsoft/teams-js";

export default  interface IDataProvider {
    getSharingLinks(listItems:Record<string, any>): Promise<ISharingResult[]>;
    //fetchSearchResultsAll(startRow:Number): Promise<Record<string, any>>;
    getSearchResults(): Promise<Record<string, any>>;
    getassociatedGroups(): Promise<void>;
}

export default class DataProvider implements IDataProvider{
    private webpartContext: WebPartContext;
    private isTeams:boolean;
    private isPrivateChannel:boolean;
    private siteUrl: string;
    private serverRelativeUrl: string;
    private tenantId:string;
    private groupId:string;
    private channelName:string;
    private channelRelativeUrl:string;
    private teamSitePath:string;
    private sp: any;
    private graph: any;
    private standardGroups:string[]=[];
    private utility: Utils = new Utils();

    constructor(context: WebPartContext) 
    {
        this.webpartContext = context;
        
        this.sp = spfi().using(SPFx(this.webpartContext), Caching);
        this.graph = graphfi().using(graphSPFx(this.webpartContext), Caching).using(PnPLogging(LogLevel.Warning));

        if (this.webpartContext.sdks.microsoftTeams) {
           this.isTeams = true;
        }

        if (this.isTeams)
        {
            this.siteUrl =  this.webpartContext.sdks.microsoftTeams.context.teamSiteUrl;
            this.groupId = this.webpartContext.sdks.microsoftTeams.context.groupId;
            this.tenantId = this.webpartContext.sdks.microsoftTeams.context.tid;
            this.isPrivateChannel = (this.webpartContext.sdks.microsoftTeams.context.channelType == "Private");
            this.channelName =  this.webpartContext.sdks.microsoftTeams.context.channelName;
            this.channelRelativeUrl = this.webpartContext.sdks.microsoftTeams.context.channelRelativeUrl;
            this.teamSitePath = this.webpartContext.sdks.microsoftTeams.context.teamSitePath;
            this.serverRelativeUrl = this.webpartContext.sdks.microsoftTeams.context.teamSitePath;
        }
        else
        { 
            this.siteUrl = this.webpartContext.pageContext.web.absoluteUrl;
            this.serverRelativeUrl = this.webpartContext.pageContext.web.serverRelativeUrl; 
            this.tenantId = this.webpartContext.pageContext.aadInfo.tenantId;
        }
    }

    public async getAssociatedGroups(): Promise<void>
    {
         // Gets the associated visitors group of a web
         const visitorsGroup = await this.sp.web.associatedVisitorGroup.select("Title")();
         this.standardGroups.push(visitorsGroup.Title);

         // Gets the associated members group of a web
         const membersGroup = await this.sp.web.associatedMemberGroup.select("Title")();
         this.standardGroups.push(membersGroup.Title);
         
         // Gets the associated owners group of a web
         const ownersGroup = await this.sp.web.associatedOwnerGroup.select("Title")();        
         this.standardGroups.push(ownersGroup.Title);
    }

    private async getDriveItemsBySearchResult(listItems:Record<string, any>): Promise<Record<string, any>>
    {
        let driveItems:Record<string, any>= {};

        const [batchedGraph, execute] = this.graph.batched();
        batchedGraph.using(Caching());

        // for each file, we need to get the permissions
        for (let fileId in listItems) 
        {
            let file = (listItems as any)[fileId];
            // the permissions endpoint on the driveItem is not (yet?) exposed in pnpjs, so we need to use the graphQueryable
            const driveItemQuery = batchedGraph.drives.getById(file.DriveId).getItemById(file.DriveItemId);
            // adding the permissions endpoint
            const graphQueryable = GraphQueryable(driveItemQuery, "permissions")
            // getting the permissions and adding the request to the batch
            graphGet(GraphQueryable(graphQueryable)).then(r =>  {
                driveItems[fileId] = r;
            });
        }

        // Executes the batched calls
        await execute();

        return driveItems;
    }


    public async getSharingLinks(listItems:Record<string, any>): Promise<ISharingResult[]>
    {
        const sharedResults: ISharingResult[] = [];
        const driveItems = await this.getDriveItemsBySearchResult(listItems);
        //const standardGroups:string[] = ["Retail Owners", "Retail Members", "Retail Visitors"];

        // now we have all the data we need, we can start building up the result
        for (let fileId in driveItems) 
        {
            let driveItem = driveItems[fileId];
            let file = listItems[fileId];
            
            let sharedWithUser:  IFacepilePersona[] = [];
            let sharingUserType = "Member";

            // Getting all the details of the file and in which folder is lives
            let folderUrl = file.Path.replace(`/${file.FileName}`, '');
            let folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;

            // for certain filetypes we get the dispform.aspx link back instead of the full path, so we need to fix that
            if (folderName.indexOf("DispForm.aspx") > -1)
            {
                folderUrl = folderUrl.substring(0, folderUrl.lastIndexOf("/Forms/DispForm.aspx"));
                folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;
                file.FileExtension = file.FileName.substring(file.FileName.lastIndexOf(".") + 1);
            }

            file.FileUrl =  file.Path;
            file.FolderUrl = folderUrl;
            file.FolderName = folderName;
            file.FileId = fileId;

            // if (file.FileName == "zalando-fossil-watch.pdf")
            //     debugger;

            // if a file has inherited permissions, the propery is returned as "inheritedFrom": {}
            // if a file has unique permissions, the propery is not returned at all
            driveItem.forEach(permission =>
            {
                // if (permission.inheritedFrom != null)
                //     return;
                // only interested in the sharing links (perhaps we should also add the normal permissions?)
                if(permission.link)
                {
                    switch (permission.link.scope)
                    {
                        case "anonymous":break;
                        case "organization":
                            let _user: IFacepilePersona = {};     
                            _user.personaName = permission.link.scope + " " + permission.link.type;
                            if (sharedWithUser.indexOf(_user) == -1)
                            { 
                                sharedWithUser.push(_user);
                            }
                        break; 
                        case "users":
                            let _users = this.utility.convertToFacePilePersona(permission.grantedToIdentitiesV2);
                            sharedWithUser.push(..._users);
                        break;
                        default: break;
                    }
                }
                else // checking the normal permissions as well, other than the sharing links
                {
                    // if the permission is not the same as the default associated spo groups, we need to add it to the sharedWithUser array
                    if (this.standardGroups.indexOf(permission.grantedTo.user.displayName) == -1)
                    {
                        let _users = this.utility.convertUserToFacePilePersona(permission.grantedToV2);
                        sharedWithUser.push(_users);          
                    }
                    else // otherwise, we're gonna add these groups and mark it as inherited permissions
                    {
                        let _user: IFacepilePersona = {};     
                        _user.personaName = permission.grantedTo.user.displayName;
                        _user.data = "Inherited";
                        if (sharedWithUser.indexOf(_user) == -1)
                        { 
                            sharedWithUser.push(_user);
                        }
                    }         
                }
            });

            if (file.SharedWithUsersOWSUSER != null)
            {
                let _users = this.utility.processUsers(file.SharedWithUsersOWSUSER);
                sharedWithUser.push(..._users);
            }

            // if there are any duplicates, this will remove them (e.g. multiple organization links)
            sharedWithUser = this.utility.uniqForObject(sharedWithUser);
            if (sharedWithUser.length == 0)
                continue;
            
            for (const user of sharedWithUser) 
            {
                if (user.data == "Guest")
                {
                    // this is most important, once we found guest users, we need to set the sharingUserType to Guest (and thus break out of the loop)
                    sharingUserType = "Guest";
                    break;
                }
                if (user.personaName.indexOf("organization") > -1)
                {
                    // this is the next most important, once we found organization links, we need to set the sharingUserType to Link (and thus break out of the loop)
                    sharingUserType = "Link";
                    break;
                }
                if (user.data == "Inherited")
                {
                    sharingUserType = "Inherited";
                }
            }
            
            // building up the result to be returned
            const sharedResult: ISharingResult = 
            {
                FileExtension : (file.FileExtension == null) ? "folder" : file.FileExtension,
                FileName : file.FileName,
                Channel: file.FolderName,
                LastModified: file.LastModifiedTime,
                SharedWith: sharedWithUser,
                ListId: file.ListId,
                ListItemId: file.ListItemId,
                Url: file.FileUrl,
                FolderUrl : file.FolderUrl,
                SharingUserType : sharingUserType,
                FileId: file.FileId
            };
            sharedResults.push(sharedResult);
        }
        return sharedResults;
    }

    public async getSearchResults(): Promise<Record<string, any>>
    {
        let listItems:Record<string, any> = {};
        let searchResults:any[] = [];
        searchResults = await this.fetchSearchResultsAll(0, searchResults);

        searchResults.forEach(results =>
        {
            results.forEach(result =>
            {
                result.hitsContainers.forEach(hits => 
                {   
                    hits.hits.forEach(hit=>
                    {
                        let SharedWithUsersOWSUser = (hit.resource['listItem']['fields']['sharedWithUsersOWSUSER'] != undefined) ? hit.resource['listItem']['fields']['sharedWithUsersOWSUSER'] : null;
                        const result:ISearchResultExtended = {
                            DriveItemId : hit.resource.id,
                            FileName : hit.resource['listItem']['fields']['fileName'],
                            FileExtension : hit.resource['listItem']['fields']['fileExtension'],
                            ListId : hit.resource['listItem']['fields']['listId'],
                            FileId : hit.resource['listItem']['id'],
                            DriveId: hit.resource['listItem']['fields']['driveId'],
                            ListItemId : hit.resource['listItem']['fields']['listItemId'],
                            Path : hit.resource['webUrl'],
                            LastModifiedTime : hit.resource['lastModifiedDateTime'],
                            SharedWithUsersOWSUSER : SharedWithUsersOWSUser
                        }
                        listItems[result.FileId] = result;
                    });
                });
            });
        });

        return listItems;
    }

    private async fetchSearchResultsAll(page:number, searchResults?:any[]): Promise<any>
    {
        if (page == 0)
        {
            searchResults = [];
        }

        let everyoneExceptExternalsUserName = `spo-grid-all-users/${this.tenantId}`;
        let query = `(IsDocument:TRUE OR IsContainer:TRUE) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone)) AND (SPSiteUrl:${this.siteUrl}) `;

        Logger.write(`Issueing search query: ${query}`, LogLevel.Verbose);
        const results = await this.graph.query({
            entityTypes: ["driveItem", "listItem"],
            query: {
                queryString: `${query}`
            },
            fields: ["path", "id", "driveId","driveItemId", "listId", "listItemId", "fileName", "fileExtension", "webUrl","lastModifiedDateTime","lastModified","SharedWithUsersOWSUSER"],
            from: `${page}`,
            size: 500           
        });
        
        searchResults.push(results);

        if (results[0].hitsContainers[0].moreResultsAvailable)
        {
            searchResults = await this.fetchSearchResultsAll(page + 500, searchResults)
        }
            
        return searchResults;
    }
}