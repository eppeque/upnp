import { spfi, SPFx } from "@pnp/sp";
export function initSharepoint(context) {
    return spfi().using(SPFx(context));
}
//# sourceMappingURL=init.js.map