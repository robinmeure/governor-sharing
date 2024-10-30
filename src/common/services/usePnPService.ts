
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { ISPFXContext, spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import { Logger, LogLevel } from '@pnp/logging';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IPnPService {
    getSiteGroups(targetSite: string): Promise<string[]>;
}
/** Represents all calls to SharePoint with help of Graph API
 * @param {WebPartContext} spfxContext - is used to make Graph API calls
 */

export const usePnPService = (spfxContext: WebPartContext | ApplicationCustomizerContext, siteUrl?: string): IPnPService => {

    // const sp: SPFI = getSP(spfxContext, siteUrl);
    // const graph: GraphFI = getGraph(spfxContext, siteUrl);

    // const getCacheFI = <T extends "SP" | "Graph">(type: T): T extends "SP" ? SPFI : GraphFI => {
    //     const cache = type === "SP" ? spfi(sp) : graphfi(graph);
    //     return cache.using(Caching({ store: "session" })) as T extends "SP" ? SPFI : GraphFI;
    // }


    const getSiteGroups = async (targetSite: string): Promise<string[]> => {

        try {
            const localSP = spfi(targetSite).using(SPFx(spfxContext as ISPFXContext));
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

    // const getDocsByGraphSearch = async (searchReq: SearchRequest): Promise<ISearchResultExtended[]> => {
    //     try {
    //         const locSearchResults: ISearchResultExtended[] = [];
    //         const graphCache = getCacheFI("Graph");

    //         Logger.write(`Issuing search query: ${searchReq.query.queryString}`, LogLevel.Verbose);
    //         const results = await graphCache.query(searchReq);

    //         locSearchResults.push(...GraphResponseToSearchResultMapper(results));

    //         if (results[0].hitsContainers[0].moreResultsAvailable) {
    //             //TODO handle pagination
    //             // locSearchResults = await fetchSearchResultsAll(page + 500, searchResults)
    //         }
    //         return locSearchResults;
    //     } catch (error) {
    //         throw error;
    //     }
    // }

    // Return functions
    return {
        getSiteGroups
    };
};
