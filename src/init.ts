import { ISPFXContext, spfi, SPFI, SPFx } from "@pnp/sp";

/**
 * Initializes the Sharepoint factory interface based on the Webpart context.
 * @param context The context of the Sharepoint Webpart.
 * @returns The Sharepoint factory interface.
 */
export function initSharepoint(context: ISPFXContext): SPFI {
  return spfi().using(SPFx(context));
}
