import { IFacepilePersona } from "@fluentui/react";
import { Identity } from "@microsoft/microsoft-graph-types";

export interface IDriveItems {
    driveId: string;
    itemId: string;
}



export type SharingUserType = "Link" | "Inherited" | "Owner" | "Member" | "Guest" | "Everyone";


//#region Graph API Response Interfaces


export interface IFileSharingResponse {
    FileExtension: string;
    FileName: string;
    LastModified: Date;
    SharedWith: IFacepilePersona[];
    ListId: string;
    ListItemId: number;
    Url: string;
    FolderUrl: string;
    Channel: string;
    FileId: string;
    SharingUserType: SharingUserType;
    SiteUrl: string;
    LastModifiedBy: Identity | undefined;
}

export interface IListItemSearchResponse {
    ListItemId: number;
    ListId: string;
    SharedWithUsersOWSUSER: string;
    FileName: string;
    ViewableByExternalUsers: boolean;
    ItemId: string;
    FileId: string;
    FileExtension: string;
    DriveId: string;
    SiteUrl: string;
    Path: string;
    LastModifiedTime: Date | undefined;
    LastModifiedBy: Identity | undefined;

    // SharedWithDetails: string;
    // Rank: number;
    // DocId: number;
    // WorkId: number;
}

export interface ISiteSearchResponse {
    name: string;
    url: string;
}

//#endregion