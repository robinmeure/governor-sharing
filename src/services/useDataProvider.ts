/* eslint-disable @typescript-eslint/no-explicit-any */
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files/web";
import "@pnp/sp/items/get-all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Logger, LogLevel } from '@pnp/logging';
import { Caching } from '@pnp/queryable';

interface IDataProvider {
    // getSharingLinks(listItems: Record<string, any>): Promise<ISharingResult[]>;
    // getSearchResults(): Promise<Record<string, any>>;
    loadAssociatedGroups(siteUrl?: string): Promise<void>;
}
/** Represents all calls to SharePoint with help of Graph API
 * @param {WebPartContext} spfxContext - is used to make Graph API calls
 */
export const useDataProvider = (spfxContext: WebPartContext): IDataProvider => {


    const loadAssociatedGroups = async (siteUrl?: string): Promise<void> => {

        try {
            const sp = spfi(siteUrl).using(SPFx(spfxContext), Caching);
            console.log("FazLog ~ loadAssociatedGroups ~ sp:", sp);
            const { Title } = await sp.web.select("Title")()
            console.log(`Web title: ${Title}`);
        }
        catch (error) {
            Logger.write(`getPageReviewItems in useSPService | Error: ${error}`, LogLevel.Error);
            throw error;
        }
    };

    // Return functions
    return {
        loadAssociatedGroups
    };
};
