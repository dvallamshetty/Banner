import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp and pnp logging system
import { SPBrowser, spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";


// eslint-disable-next-line no-var
var _sp: SPFI;

export const getSP = (context: WebPartContext, sourceURL?: string): SPFI => {
  if (!!context) { // eslint-disable-line eqeqeq
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    if (sourceURL !== undefined) {
      _sp = spfi().using(SPBrowser({ baseUrl: sourceURL }));
    } else {
      _sp = spfi().using(SPFx(context));
    }
  }
  return _sp;
};
