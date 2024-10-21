/* eslint-disable */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { BatchRequestContent, BatchRequestStep, BatchResponseContent } from "@microsoft/microsoft-graph-client";
import { Permission } from "@microsoft/microsoft-graph-types";


interface IGraphService {
    getDriveItemsBySearchResult(listItems: Record<string, any>): Promise<Record<string, any>>;

}

export const useGraphService = (context: WebPartContext | ApplicationCustomizerContext): IGraphService => {

    const getDriveItemsBySearchResult = async (listItems: Record<string, any>): Promise<Record<string, any>> => {
        try {
            console.log("FazLog ~ getDriveItemsBySearchResult ~ listItems:", listItems);
            const driveItems: Record<string, any> = {};

            const graphClient = await context.msGraphClientFactory
                .getClient('3');

            const batchReq: BatchRequestStep[] = [];
            for (const fileId in listItems) {
                if (Object.prototype.hasOwnProperty.call(listItems, fileId)) {
                    const file = listItems[fileId];
                    batchReq.push({
                        id: fileId,
                        request: new Request(`https://graph.microsoft.com/drives/${file.DriveId}/items/${file.DriveItemId}/permissions`, {
                            method: "GET"
                        })
                    });
                }
            }

            const batchRequestContent = new BatchRequestContent(batchReq);
            const content = await batchRequestContent.getContent();

            // POST the batch request content to the /$batch endpoint
            const batchResponse = await graphClient.api('/$batch').post(content);
            // Create a BatchResponseContent object to parse the response
            const batchResponseContent = new BatchResponseContent(batchResponse);
            for (const fileId in listItems) {
                if (Object.prototype.hasOwnProperty.call(listItems, fileId)) {
                    const driveResponse = batchResponseContent.getResponseById(fileId);
                    if (driveResponse.ok) {
                        const drivePermissionItem: Permission = (await driveResponse.json())?.value as Permission;
                        driveItems[fileId] = drivePermissionItem;
                    }
                }
            }
            return driveItems;
        } catch (error) {
            console.log("FazLog ~ getDriveItemsBySearchResult ~ error:", error);
        }
    }

    return {
        getDriveItemsBySearchResult
    };

}