import { Identity } from "@microsoft/microsoft-graph-types";
import { DeleteAction } from "@microsoft/microsoft-graph-types-beta";

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

    // SharedWithDetails: string;
    // Rank: number;
    // DocId: number;
    // WorkId: number;
}

export interface IGraphResponseMetadata {
    moreResultsAvailable: boolean;
    totalResults: number;
}

export interface ISiteSearchResponse {
    name: string;
    url: string;
}

export interface IItemActivityAction {
    comment: {
        isReply?: false,
        parentAuthor?: Identity,
        participants?: Identity[]
    };
    create: {};
    delete: DeleteAction;
    edit: {};
    mention: {
        mentionees: Identity[];
    };
    move: {
        from?: string;
        to?: string;
    };
    rename: {
        oldName?: string;
        newName?: string;
    };
    restore: {};
    share: {
        recipients: Identity[]
    };
    version: {
        newVersion: string;
    };
}

export interface IItemActivity {
    id: string;
    action: string;
    actor: Identity | undefined;
    time: Date;
}

//#endregion