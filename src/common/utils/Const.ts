

import { EntityType } from "@microsoft/microsoft-graph-types";

export const _CONST = Object.freeze({
    GraphSearch: {
        DocsSearch: {
            EntityType: ["driveItem", "listItem"] as EntityType[],
            Fields: ["path", "id", "driveId", "listId", "listItemId", "fileName", "fileExtension", "webUrl", "lastModifiedDateTime", "lastModifiedBy", "SharedWithUsersOWSUSER", "SPSiteUrl", "viewableByExternalUsers"]
        },
        SiteSearch: {
            EntityType: ["site"] as EntityType[],
        }
    }
});