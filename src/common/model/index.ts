import { Permission } from "@microsoft/microsoft-graph-types";



export interface IDrivePermissionParams {
    driveId: string;
    driveItemId: string;
}

export interface IDrivePermissionResponse {
    fileId: string;
    permissions: Permission[];
}

export interface ISiteData {
    name: string;
    url: string;
}