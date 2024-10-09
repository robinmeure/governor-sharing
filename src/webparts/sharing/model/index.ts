import { WebPartContext } from "@microsoft/sp-webpart-base";
import IDataProvider from "../components/SharingView/DataProvider";


export interface ISharingWebPartContext {
    webpartContext: WebPartContext;
    isTeams: boolean;
    pageLimit: number;
    dataProvider: IDataProvider;
}