
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISharingWebPartProps } from "../SharingWebPart";


//#region Hooks
export interface ISharingWebPartContext {
    webpartContext: WebPartContext;
    webpartProperties: ISharingWebPartProps;
    isTeams: boolean;
    pageLimit: number;
}

//#endregion
