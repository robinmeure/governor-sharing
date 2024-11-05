
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPFI, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import { Logger, LogLevel } from '@pnp/logging';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { GraphFI, graphfi } from "@pnp/graph";
import { Caching } from "@pnp/queryable";
import { getGraph, getSP } from "../config/PnPjsConfig";

interface IPnPService {
    getSiteGroups(targetSite: string): Promise<string[]>;
}
/** Represents all calls to SharePoint with help of Graph API
 * @param {WebPartContext} spfxContext - is used to make Graph API calls
 */

export const usePnPService = (spfxContext: WebPartContext | ApplicationCustomizerContext, siteUrl?: string): IPnPService => {

    const sp: SPFI = getSP(spfxContext, siteUrl);
    const graph: GraphFI = getGraph(spfxContext, siteUrl);

    const getCacheFI = <T extends "SP" | "Graph">(type: T): T extends "SP" ? SPFI : GraphFI => {
        const cache = type === "SP" ? spfi(sp) : graphfi(graph);
        return cache.using(Caching({ store: "session" })) as T extends "SP" ? SPFI : GraphFI;
    }


    //NOTE: This function is not used in the current implementation
    /** 
     * Gets the associated groups of a site
     * @returns {Promise<string[]>} - An array of associated groups
    */
    const getSiteGroups = async (): Promise<string[]> => {

        try {
            const localSP = getCacheFI("SP"); //spfi(targetSite).using(SPFx(spfxContext as ISPFXContext));
            const { Title } = await localSP.web.select("Title")();
            console.log(`Web title: ${Title}`);
            const locStandardGroups: string[] = [];

            // Gets the associated visitors group of a web
            const visitorsGroup = await localSP.web.associatedVisitorGroup.select("Title")();
            locStandardGroups.push(visitorsGroup.Title);

            // Gets the associated members group of a web
            const membersGroup = await localSP.web.associatedMemberGroup.select("Title")();
            locStandardGroups.push(membersGroup.Title);

            // Gets the associated owners group of a web
            const ownersGroup = await localSP.web.associatedOwnerGroup.select("Title")();
            locStandardGroups.push(ownersGroup.Title);
            return locStandardGroups;
        }
        catch (error) {
            Logger.write(`loadAssociatedGroups in usePnPService | Error: ${error}`, LogLevel.Error);
            throw error;
        }
    };

    return {
        getSiteGroups
    };
};
