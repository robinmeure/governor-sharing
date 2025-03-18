import { Identity } from "@microsoft/microsoft-graph-types";

export interface IDriveItems {
    driveId: string;
    itemId: string;
}

export interface ISharedUser {
    displayName: string,
    id: string,
    type: SharedType
}

export type SharedType = "Link" | "Inherited" | "Member" | "Guest" | "Everyone" | "Group";

export interface IFileSharingResponse {
    DriveId: string;
    FileExtension: string;
    FileName: string;
    LastModified: Date;
    SharedWith: ISharedUser[];
    ListId: string;
    ListItemId: number;
    Url: string;
    FolderUrl: string;
    Channel: string;
    FileId: string;
    SharedType: SharedType;
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
}

export interface IGraphResponseMetadata {
    moreResultsAvailable: boolean;
    totalResults: number;
}

export interface ISiteSearchResponse {
    name: string;
    url: string;
}

//#endregion