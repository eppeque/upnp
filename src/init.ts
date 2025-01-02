import { ISPFXContext, spfi, SPFI, SPFx } from "@pnp/sp";

export function initSharepoint(context: ISPFXContext): SPFI {
  return spfi().using(SPFx(context));
}
