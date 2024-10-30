import { IFacepilePersona } from "@fluentui/react";
import { SearchResponse, EntityType, SearchHit, ListItem, DriveItem, Site } from "@microsoft/microsoft-graph-types";
import { IListItemSearchResponse, ISiteSearchResponse, IFileSharingResponse, SharedType, ISharedUser } from "../model";
import { convertToFacePilePersona, convertUserToFacePilePersona, processUsers } from "../utils/Utils";
import { DrivePermissionResponse } from "../services/useGraphService";



export const GraphSearchResponseMapper = <T>(searchResponse: SearchResponse[], entityType: EntityType[]): T[] => {

    try {
        const locMappedVal: unknown[] = [];

        searchResponse.forEach(results => {
            results.hitsContainers?.forEach(hitsContainer => {
                hitsContainer?.hits?.forEach((hit: SearchHit) => {

                    if (entityType.includes("driveItem") || entityType.includes("listItem")) {

                        if (!hit.resource || !("listItem" in hit.resource)) return;
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const resource = hit.resource as any;
                        const listItem = resource.listItem as ListItem | DriveItem;
                        // eslint-disable-next-line @typescript-eslint/no-explicit-any
                        const fields = (listItem as { fields: { [key: string]: any } }).fields;
                        const SharedWithUsersOWSUser = (fields.sharedWithUsersOWSUSER !== undefined) ? fields.sharedWithUsersOWSUSER : null;

                        // if we don't get a driveId back (e.g. documentlibrary), then skip the returned item
                        if (fields.driveId === undefined)
                            return;

                        const result: IListItemSearchResponse = {
                            ItemId: resource.id ?? '',
                            FileName: fields.fileName,
                            FileExtension: fields.fileExtension ? fields.fileExtension : "folder",
                            ListId: fields.listId,
                            FileId: listItem.id ?? '',
                            DriveId: fields.driveId,
                            ListItemId: fields.listItemId,
                            Path: fields.path ?? '',
                            LastModifiedTime: resource?.lastModifiedDateTime ? new Date(resource.lastModifiedDateTime) : undefined,
                            SharedWithUsersOWSUSER: SharedWithUsersOWSUser,
                            SiteUrl: fields.spSiteUrl,
                            ViewableByExternalUsers: fields?.viewableByExternalUsers ?? false,
                            LastModifiedBy: {
                                displayName: resource.lastModifiedBy?.user?.displayName,
                                id: resource?.lastModifiedBy?.user.email
                            }
                        }
                        locMappedVal.push(result as unknown);
                    } else if (entityType.includes("site")) {
                        if (!hit.resource) return;
                        const site = hit.resource as Site;
                        if (!site.displayName || !site.webUrl) return;

                        const result: ISiteSearchResponse = {
                            name: site.displayName,
                            url: site.webUrl
                        }
                        locMappedVal.push(result as unknown);

                    }
                });
            });
        });


        return locMappedVal as T[];
    } catch (error) {
        console.log("GraphSearchResponseMapper ~ error:", error);
        throw error;
    }
}

export const DrivePermissionResponseMapper = (file: IListItemSearchResponse, driveItem: DrivePermissionResponse): IFileSharingResponse => {
    const sharedWithUser: IFacepilePersona[] = [];
    const sharedWithUser2: ISharedUser[] = [];

    let sharingUserType: SharedType = "Member";

    // Getting all the details of the file and in which folder is lives
    let folderUrl = file.Path ? file.Path.replace(`/${file.FileName}`, '') : '';
    let folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;

    // for certain filetypes we get the dispform.aspx link back instead of the full path, so we need to fix that
    if (folderName.indexOf("DispForm.aspx") > -1) {
        folderUrl = folderUrl.substring(0, folderUrl.lastIndexOf("/Forms/DispForm.aspx"));
        folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;
        if (file.FileName) {
            file.FileExtension = file.FileName.substring(file.FileName.lastIndexOf(".") + 1);
        }
    }

    driveItem.permissions.map(perm => {
        if (perm.link) {
            switch (perm.link.scope) {
                case "anonymous":
                    break;
                case "organization": {
                    // Shared through links
                    const _user: IFacepilePersona = {};
                    _user.personaName = perm.link.scope + " " + perm.link.type;
                    _user.data = "Organization";
                    if (sharedWithUser.indexOf(_user) === -1) {
                        sharedWithUser.push(_user);
                    }
                    break;
                }
                case "users": {
                    // find the the user is group or actual user
                    const _users = perm.grantedToIdentitiesV2 ? convertToFacePilePersona(perm.grantedToIdentitiesV2) : [];
                    sharedWithUser.push(..._users);
                    break;
                }
                default:
                    break;
            }
        }
        else // checking the normal permissions as well, other than the sharing links
        {
            // if the permission is not the same as the default associated spo groups, we need to add it to the sharedWithUser array
            if (perm.grantedTo?.user && perm.grantedToV2) {
                const _users = convertUserToFacePilePersona(perm.grantedToV2);
                sharedWithUser.push(_users);
            }
            else // otherwise, we're gonna add these groups and mark it as inherited permissions
            {
                const _user: IFacepilePersona = {};
                _user.personaName = perm.grantedTo?.user?.displayName ?? undefined;
                _user.data = "Inherited";
                if (sharedWithUser.indexOf(_user) === -1) {
                    sharedWithUser.push(_user);
                }
            }
        }

        if (file.SharedWithUsersOWSUSER !== null) {
            const _users = file.SharedWithUsersOWSUSER ? processUsers(file.SharedWithUsersOWSUSER) : [];
            sharedWithUser.push(..._users);
        }
    });
    // if there are any duplicates, this will remove them (e.g. multiple organization links)
    if (sharedWithUser?.length > 0) {

        sharedWithUser.forEach(element => {
            const name = element.personaName || element.name || "";

            sharedWithUser2.push({
                id: element.id || name,
                displayName: name,
                type: element.data
            })
        });



        let isGuest = false;
        let isLink = false;
        let isInherited = false;

        for (const user of sharedWithUser) {
            switch (user.data) {
                case "Guest": isGuest = true; break;
                case "Organization": isLink = true; break;
                case "Inherited": isInherited = true; break;
            }
        }

        // if we found a guest user, we need to set the sharingUserType to Guest
        if (isGuest) {
            sharingUserType = "Guest";
        }
        else if (isLink) {
            sharingUserType = "Link";
        }
        else if (isInherited) {
            sharingUserType = "Inherited";
        }
    }

    // sharedWithUser2 = uniqForObject(sharedWithUser2);
    const uniqueUsers = Array.from(
        new Map(sharedWithUser2.map(user => [user.displayName, user])).values()
    );

    // building up the result to be returned
    const sharedResult: IFileSharingResponse =
    {
        FileExtension: file?.FileExtension ? file.FileExtension : "folder",
        FileName: file.FileName ?? '',
        Channel: folderName,
        LastModified: file.LastModifiedTime ?? new Date(),
        SharedWith: uniqueUsers,
        ListId: file.ListId ?? '',
        ListItemId: file.ListItemId ?? 0,
        Url: file.Path ?? '',
        FolderUrl: folderUrl,
        SharedType: sharingUserType,
        FileId: file.FileId ?? '',
        SiteUrl: file.SiteUrl ?? '',
        LastModifiedBy: file.LastModifiedBy
    };
    return sharedResult;

}