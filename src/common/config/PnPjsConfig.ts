

import { WebPartContext } from "@microsoft/sp-webpart-base";
// import pnp and pnp logging system
import { ISPFXContext, spfi, SPFI, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";

let _sp: SPFI;
let _graph: GraphFI;

export const getSP = (context: WebPartContext | ApplicationCustomizerContext, siteUrl?: string): SPFI => {
    if (_sp === undefined || _sp === null) {
        //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
        // The LogLevel set's at what level a message will be written to the console
        _sp = spfi(siteUrl).using(spSPFx(context as ISPFXContext));
    }
    return _sp;
};

export const getGraph = (context: WebPartContext | ApplicationCustomizerContext, siteUrl?: string): GraphFI => {
    if (_graph === undefined || _graph === null) {
        //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
        // The LogLevel set's at what level a message will be written to the console
        _graph = graphfi(siteUrl).using(graphSPFx(context as ISPFXContext));
    }
    return _graph;
};