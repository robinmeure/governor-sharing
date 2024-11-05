
import { WebPartContext } from "@microsoft/sp-webpart-base";


//#region Hooks
export interface ISharingWebPartContext {
    webpartContext: WebPartContext;
    isTeams: boolean;
    pageLimit: number;
}

//#endregion
