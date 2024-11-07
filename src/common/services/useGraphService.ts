import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { BatchRequestContent, BatchRequestStep, BatchResponseContent } from "@microsoft/microsoft-graph-client";
import { Permission, SearchRequest, SearchResponse } from "@microsoft/microsoft-graph-types";
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IDriveItems } from "../model";
import { ItemActivityOLD } from "@microsoft/microsoft-graph-types-beta";


export interface DrivePermissionResponse {
    fileId: string;
    permissions: Permission[];
}

interface IGraphService {
    getDriveItemsPermission(driveItems: Record<string, IDriveItems>): Promise<DrivePermissionResponse[]>;
    getByGraphSearch(searchReq: SearchRequest): Promise<SearchResponse[]>;
    getItemsActivityBETA(file: { driveId: string, itemId: string }): Promise<ItemActivityOLD[] | undefined>;
}

export const useGraphService = (spfxContext: WebPartContext | ApplicationCustomizerContext): IGraphService => {

    let graphClient: MSGraphClientV3;

    /**
     * Initializes the MSGraphClientV3 instance if it is not already initialized.
     * @returns {Promise<void>} A promise that resolves when the client is initialized.
    */
    const initializeGraphClient = async (): Promise<void> => {
        if (!graphClient) {
            const client = await spfxContext.msGraphClientFactory.getClient("3") as MSGraphClientV3;
            if (!graphClient) {
                graphClient = client;
            }
        }
    };

    /**
     * Executes a search query using Microsoft Graph API.
     * @param {SearchRequest} searchReq - The search request object.
     * @returns {Promise<SearchResponse[]>} A promise that resolves to an array of search responses.
     * @throws Will throw an error if the search request fails.
     */
    const getByGraphSearch = async (searchReq: SearchRequest): Promise<SearchResponse[]> => {
        try {
            await initializeGraphClient();
            const response = await graphClient.api('/search/query')
                .post({
                    requests: [
                        searchReq
                    ]
                });
            return response?.value;
        } catch (error) {
            console.log("getByGraphSearch ~ error:", error);
            throw error;
        }
    }

    /**
     * Retrieves the permissions for a list of drive items using Microsoft Graph API.
     * @param {Record<string, IDriveItems>} listItems - A record of drive items with their IDs as keys.
     * @returns {Promise<DrivePermissionResponse[]>} A promise that resolves to an array of drive permission responses.
     * @throws Will throw an error if the request for permissions fails.
     */
    const getDriveItemsPermission = async (listItems: Record<string, IDriveItems>): Promise<DrivePermissionResponse[]> => {
        try {
            await initializeGraphClient();
            const driveItemPermissions: DrivePermissionResponse[] = [];

            const batchReq: BatchRequestStep[] = [];
            Object.entries(listItems).forEach(([key, file]) => {
                batchReq.push({
                    id: key,
                    request: new Request(`https://graph.microsoft.com/drives/${file.driveId}/items/${file.itemId}/permissions`, {
                        method: "GET"
                    })
                });
            });

            const batchRequestContent = new BatchRequestContent(batchReq);
            const content = await batchRequestContent.getContent();

            const batchResponse = await graphClient.api('/$batch').post(content);
            // Create a BatchResponseContent object to parse the response
            const batchResponseContent = new BatchResponseContent(batchResponse);
            const driveItemsPromises = Object.entries(listItems).map(async ([key]) => {
                const driveResponse = batchResponseContent.getResponseById(key);
                if (driveResponse.ok) {
                    const driveItemPermission: Permission[] = (await driveResponse.json())?.value as Permission[];
                    driveItemPermissions.push({
                        fileId: key,
                        permissions: driveItemPermission
                    });
                }
            });

            await Promise.all(driveItemsPromises);
            return driveItemPermissions;
        } catch (error) {
            console.error("getDriveItemsPermission ~ error", error);
            throw error;
        }
    }

    /**
     * Retrieves the activity for a specific drive item using the beta version of Microsoft Graph API.
     * @param {{ driveId: string, itemId: string }} file - An object containing the drive ID and item ID.
     * @returns {Promise<ItemActivityOLD[] | undefined>} A promise that resolves to an array of item activities or undefined if no activities are found.
     * @throws Will throw an error if the request for item activities fails.
     */
    const getItemsActivityBETA = async (file: { driveId: string, itemId: string }): Promise<ItemActivityOLD[] | undefined> => {
        try {
            await initializeGraphClient();

            const activities = await graphClient.api(`/drives/${file.driveId}/items/${file.itemId}/activities`)
                .version('beta')
                .get();
            if (activities) {
                return activities.value as ItemActivityOLD[];
            }
            return undefined;
        } catch (error) {
            console.error("getItemsActivityBETA ~ error", error);
            throw error;
        }
    }

    return {
        getDriveItemsPermission,
        getByGraphSearch,
        getItemsActivityBETA
    };

}