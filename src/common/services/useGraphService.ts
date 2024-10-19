// /* eslint-disable @typescript-eslint/no-explicit-any */
// import { useState, useEffect, useCallback } from 'react';
// import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { spfi, SPFx } from '@pnp/sp';
// import { ISearchResultExtended } from "./ISearchResultExtended";

// import { IFacepilePersona } from "@fluentui/react";
// import { graphfi, graphGet, GraphQueryable, SPFx as graphSPFx } from "@pnp/graph";
// import "@pnp/graph/batching";
// import "@pnp/graph/onedrive";
// import "@pnp/graph/search";
// import "@pnp/graph/users";
// import { Logger, LogLevel, PnPLogging } from "@pnp/logging";
// import { Caching } from "@pnp/queryable";
// import "@pnp/sp/presets/all";
// import "@pnp/sp/search";
// import "@pnp/sp/sharing";
// import "@pnp/sp/webs";
// import { ISharingResult } from "./ISharingResult";
// import { convertToFacePilePersona, convertUserToFacePilePersona, processUsers, uniqForObject } from "./Utils";

// export interface IDataProvider {
//     getSharingLinks(listItems: Record<string, any>): Promise<ISharingResult[]>;
//     getSearchResults(): Promise<Record<string, any>>;
//     loadAssociatedGroups(): Promise<void>;
// }

// export const useGraphService = (context: WebPartContext) => {
//     const [standardGroups, setStandardGroups] = useState<string[]>([]);
//     const [sp, setSp] = useState<any>(null);
//     const [graph, setGraph] = useState<any>(null);
//     const [isTeams, setIsTeams] = useState<boolean>(false);
//     const [siteUrl, setSiteUrl] = useState<string>('');
//     const [tenantId, setTenantId] = useState<string>('');
//     const [groupId, setGroupId] = useState<string>('');
//     const [isPrivateChannel, setIsPrivateChannel] = useState<boolean>(true);

//     useEffect(() => {
//         const init = () => {
//             const spInstance = spfi().using(SPFx(context), Caching);
//             const graphInstance = graphfi().using(graphSPFx(context), Caching).using(PnPLogging(LogLevel.Warning));
//             setSp(spInstance);
//             setGraph(graphInstance);

//             if (context.sdks.microsoftTeams) {
//                 setIsTeams(true);
//                 setSiteUrl(context.sdks.microsoftTeams.context.teamSiteUrl);
//                 setTenantId(context.sdks.microsoftTeams.context.tid);
//                 setGroupId(context.sdks.microsoftTeams.context.groupId);
//                 setIsPrivateChannel(context.sdks.microsoftTeams.context.channelType === "Private");
//             } else {
//                 setSiteUrl(context.pageContext.web.absoluteUrl);
//                 setTenantId(context.pageContext.aadInfo.tenantId);
//             }
//         };

//         init();
//     }, [context]);

//     const loadAssociatedGroups = useCallback(async () => {
//         if (!sp) return;

//         const visitorsGroup = await sp.web.associatedVisitorGroup.select("Title")();
//         const membersGroup = await sp.web.associatedMemberGroup.select("Title")();
//         const ownersGroup = await sp.web.associatedOwnerGroup.select("Title")();

//         setStandardGroups([visitorsGroup.Title, membersGroup.Title, ownersGroup.Title]);
//     }, [sp]);

//     const getDriveItemsBySearchResult = useCallback(async (listItems: Record<string, any>): Promise<Record<string, any>> => {
//         if (!graph) return {};

//         const driveItems: Record<string, any> = {};
//         const [batchedGraph, execute] = graph.batched();
//         batchedGraph.using(Caching());

//         for (const fileId in listItems) {
//             const file = listItems[fileId];
//             const driveItemQuery = batchedGraph.drives.getById(file.DriveId).getItemById(file.DriveItemId);
//             const graphQueryable = GraphQueryable(driveItemQuery, "permissions");
//             const r = await graphGet(GraphQueryable(graphQueryable));
//             driveItems[fileId] = r;
//         }

//         await execute();
//         return driveItems;
//     }, [graph]);

//     const getSharingLinks = useCallback(async (listItems: Record<string, any>): Promise<ISharingResult[]> => {
//         const sharedResults: ISharingResult[] = [];
//         const driveItems = await getDriveItemsBySearchResult(listItems);

//         for (const fileId in driveItems) {
//             const driveItem = driveItems[fileId];
//             const file = listItems[fileId];

//             let sharedWithUser: IFacepilePersona[] = [];
//             let sharingUserType = "Member";

//             let folderUrl = file.Path.replace(`/${file.FileName}`, '');
//             let folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;

//             if (folderName.indexOf("DispForm.aspx") > -1) {
//                 folderUrl = folderUrl.substring(0, folderUrl.lastIndexOf("/Forms/DispForm.aspx"));
//                 folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;
//                 file.FileExtension = file.FileName.substring(file.FileName.lastIndexOf(".") + 1);
//             }

//             file.FileUrl = file.Path;
//             file.FolderUrl = folderUrl;
//             file.FolderName = folderName;
//             file.FileId = fileId;

//             driveItem.forEach(permission => {
//                 if (permission.link) {
//                     switch (permission.link.scope) {
//                         case "anonymous":
//                             break;
//                         case "organization": {
//                             const _user: IFacepilePersona = {};
//                             _user.personaName = permission.link.scope + " " + permission.link.type;
//                             _user.data = "Organization";
//                             if (sharedWithUser.indexOf(_user) === -1) {
//                                 sharedWithUser.push(_user);
//                             }
//                             break;
//                         }
//                         case "users": {
//                             const _users = convertToFacePilePersona(permission.grantedToIdentitiesV2);
//                             sharedWithUser.push(..._users);
//                             break;
//                         }
//                         default:
//                             break;
//                     }
//                 } else {
//                     if (standardGroups.indexOf(permission.grantedTo.user.displayName) === -1) {
//                         const _users = convertUserToFacePilePersona(permission.grantedToV2);
//                         sharedWithUser.push(_users);
//                     } else {
//                         const _user: IFacepilePersona = {};
//                         _user.personaName = permission.grantedTo.user.displayName;
//                         _user.data = "Inherited";
//                         if (sharedWithUser.indexOf(_user) === -1) {
//                             sharedWithUser.push(_user);
//                         }
//                     }
//                 }
//             });

//             if (file.SharedWithUsersOWSUSER !== null) {
//                 const _users = processUsers(file.SharedWithUsersOWSUSER);
//                 sharedWithUser.push(..._users);
//             }

//             sharedWithUser = uniqForObject(sharedWithUser);
//             if (sharedWithUser.length === 0) continue;

//             let isGuest = false;
//             let isLink = false;
//             let isInherited = false;

//             for (const user of sharedWithUser) {
//                 switch (user.data) {
//                     case "Guest": isGuest = true; break;
//                     case "Organization": isLink = true; break;
//                     case "Inherited": isInherited = true; break;
//                 }
//             }

//             if (isGuest) {
//                 sharingUserType = "Guest";
//             } else if (isLink) {
//                 sharingUserType = "Link";
//             } else if (isInherited) {
//                 sharingUserType = "Inherited";
//             }

//             const sharedResult: ISharingResult = {
//                 FileExtension: (file.FileExtension === null) ? "folder" : file.FileExtension,
//                 FileName: file.FileName,
//                 Channel: file.FolderName,
//                 LastModified: file.LastModifiedTime,
//                 SharedWith: sharedWithUser,
//                 ListId: file.ListId,
//                 ListItemId: file.ListItemId,
//                 Url: file.FileUrl,
//                 FolderUrl: file.FolderUrl,
//                 SharingUserType: sharingUserType,
//                 FileId: file.FileId,
//                 SiteUrl: file.SiteUrl
//             };
//             sharedResults.push(sharedResult);
//             Logger.writeJSON(sharedResult, LogLevel.Verbose);
//         }
//         return sharedResults;
//     }, [getDriveItemsBySearchResult, standardGroups]);

//     const getSearchResults = useCallback(async (): Promise<Record<string, any>> => {
//         const listItems: Record<string, any> = {};
//         let searchResults: any[] = [];
//         searchResults = await fetchSearchResultsAll(0, searchResults);

//         searchResults.forEach(results => {
//             results.forEach(result => {
//                 result.hitsContainers.forEach(hits => {
//                     hits.hits.forEach(hit => {
//                         const SharedWithUsersOWSUser = (hit.resource.listItem.fields.sharedWithUsersOWSUSER !== undefined) ? hit.resource.listItem.fields.sharedWithUsersOWSUSER : null;

//                         if (hit.resource.listItem.fields.driveId === undefined) return;

//                         const result: ISearchResultExtended = {
//                             DriveItemId: hit.resource.id,
//                             FileName: hit.resource.listItem.fields.fileName,
//                             FileExtension: hit.resource.listItem.fields.fileExtension,
//                             ListId: hit.resource.listItem.fields.listId,
//                             FileId: hit.resource.listItem.id,
//                             DriveId: hit.resource.listItem.fields.driveId,
//                             ListItemId: hit.resource.listItem.fields.listItemId,
//                             Path: hit.resource.webUrl,
//                             LastModifiedTime: hit.resource.lastModifiedDateTime,
//                             SharedWithUsersOWSUSER: SharedWithUsersOWSUser,
//                             SiteUrl: hit.resource.listItem.fields.spSiteUrl
//                         }
//                         listItems[result.FileId] = result;
//                         Logger.writeJSON(result, LogLevel.Verbose);
//                     });
//                 });
//             });
//         });

//         return listItems;
//     }, []);

//     const fetchSearchResultsAll = useCallback(async (page: number, searchResults?: any[]): Promise<any> => {
//         if (page === 0) {
//             searchResults = [];
//         }

//         const everyoneExceptExternalsUserName = `spo-grid-all-users/${tenantId}`;

//         const query = (isTeams && !isPrivateChannel) ?
//             `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone)) AND (GroupId:${groupId} OR RelatedGroupId:${groupId})`
//             : `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone)) AND (SPSiteUrl:${siteUrl})`

//         Logger.write(`Issuing search query: ${query}`, LogLevel.Verbose);
//         const results = await graph.query({
//             entityTypes: ["driveItem", "listItem"],
//             query: {
//                 queryString: `${query}`
//             },
//             fields: ["path", "id", "driveId", "driveItemId", "listId", "listItemId", "fileName", "fileExtension", "webUrl", "lastModifiedDateTime", "lastModified", "SharedWithUsersOWSUSER", "SPSiteUrl"],
//             from: `${page}`,
//             size: 500
//         });

//         searchResults.push(results);

//         if (results[0].hitsContainers[0].moreResultsAvailable) {
//             searchResults = await fetchSearchResultsAll(page + 500, searchResults)
//         }

//         return searchResults;
//     }, [graph, isTeams, isPrivateChannel, siteUrl, tenantId, groupId]);

//     return {
//         loadAssociatedGroups,
//         getSharingLinks,
//         getSearchResults
//     };
// };
