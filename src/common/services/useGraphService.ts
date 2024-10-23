/* eslint-disable */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { BatchRequestContent, BatchRequestStep, BatchResponseContent } from "@microsoft/microsoft-graph-client";
import { Permission, SearchRequest, SearchResponse } from "@microsoft/microsoft-graph-types";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IDrivePermissionParams, IDrivePermissionResponse } from "../model";

interface IGraphService {
    getDriveItemsPermission(driveItems: Record<string, IDrivePermissionParams>): Promise<IDrivePermissionResponse[]>;
    getByGraphSearch(searchReq: SearchRequest): Promise<SearchResponse[]>;

}

export const useGraphService = (spfxContext: WebPartContext | ApplicationCustomizerContext): IGraphService => {

    let graphClient: MSGraphClientV3;

    const initializeGraphClient = async () => {
        if (!graphClient) {
            graphClient = await spfxContext.msGraphClientFactory.getClient("3") as MSGraphClientV3;
        }
    };

    const getByGraphSearch = async (searchReq: SearchRequest): Promise<SearchResponse[]> => {
        try {
            await initializeGraphClient();
            const response = await graphClient.api('/search/query')
                .post({
                    requests: [
                        searchReq
                    ]
                });
            console.log("FazLog ~ getByGraphSearch ~ response:", response);
            return response?.value;
        } catch (error) {
            console.log("FazLog ~ getDocsByGraphSearch ~ error:", error);
            throw error;
        }
    }

    const getDriveItemsPermission = async (listItems: Record<string, IDrivePermissionParams>): Promise<IDrivePermissionResponse[]> => {
        try {
            await initializeGraphClient();
            const driveItemPermissions: IDrivePermissionResponse[] = [];

            const batchReq: BatchRequestStep[] = [];
            Object.entries(listItems).forEach(([key, file]) => {
                batchReq.push({
                    id: key,
                    request: new Request(`https://graph.microsoft.com/drives/${file.driveId}/items/${file.driveItemId}/permissions`, {
                        method: "GET"
                    })
                });
            });

            const batchRequestContent = new BatchRequestContent(batchReq);
            const content = await batchRequestContent.getContent();

            // POST the batch request content to the /$batch endpoint
            const batchResponse = await graphClient.api('/$batch').post(content);
            // Create a BatchResponseContent object to parse the response
            const batchResponseContent = new BatchResponseContent(batchResponse);
            const driveItemsPromises = Object.entries(listItems).map(async ([key]) => {
                const driveResponse = batchResponseContent.getResponseById(key);
                if (driveResponse.ok) {
                    const driveItemPermission: Permission = (await driveResponse.json())?.value as Permission;
                    driveItemPermissions.push({
                        fileId: key,
                        permission: driveItemPermission
                    });
                }
            });

            await Promise.all(driveItemsPromises);
            return driveItemPermissions;
        } catch (error) {
            console.log("FazLog ~ getDriveItemsBySearchResult ~ error:", error);
            throw error;
        }
    }

    return {
        getDriveItemsPermission,
        getByGraphSearch
    };

}