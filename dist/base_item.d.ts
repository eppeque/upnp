import type { IAttachmentInfo } from "@pnp/sp/attachments";
/**
 * The base item with default properties of a Sharepoint list item.
 */
export interface BaseItem {
    ID: number;
    Title: string;
    AttachmentFiles?: IAttachmentInfo[];
}
