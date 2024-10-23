/* eslint-disable */

import { EntityType } from "@microsoft/microsoft-graph-types";

export const _CONST = Object.freeze({
    GraphSearch: {
        DocsSearch: {
            EntityType: ["driveItem", "listItem"] as EntityType[],
            Fields: ["path", "id", "driveId", "driveItemId", "listId", "listItemId", "fileName", "fileExtension", "webUrl", "lastModifiedDateTime", "lastModified", "SharedWithUsersOWSUSER", "SPSiteUrl", "name"]
        },
        SiteSearch: {
            EntityType: ["site"] as EntityType[],
        }
    }
});