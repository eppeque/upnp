import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import type { BaseItem } from "./base_item";
/**
 * A base class that exposes the most basic operations possible.
 *
 * This class should be inherited by each service of each list.
 */
export declare class Service<T extends BaseItem> {
    /**
     * The Sharepoint object provided by the `@pnp/sp` package.
     */
    private sp;
    /**
     * The title of the list for this service.
     */
    private listTitle;
    /**
     * The cached items of the list for this service.
     */
    private items;
    /**
     * Creates a base service object with basic operations.
     * @param sp The Sharepoint object provided by the `@pnp/sp` package.
     * @param listTitle The title of the list for this service.
     */
    constructor(sp: SPFI, listTitle: string);
    /**
     * Returns all the items contained in the list.
     * @param fetchAttachmentFiles Tells whether this method must fetch the attachment files of each item in the list.
     * @returns All the items contained in the list.
     */
    getAllItems(fetchAttachmentFiles?: boolean): Promise<T[]>;
    /**
     * Returns all the attachment files for the item with the given `id`.
     * @param id The ID of the item for which the attachment files must be fetched.
     * @returns An array of the attachment files for the item with the given `id`.
     */
    private fetchAttachmentFiles;
    /**
     * Returns the items in the list that respect the `filter` callback.
     * @param filter The filter callback that determines whether an item should be included in the result or not.
     * @returns The items for which the `filter` callback returned `true`
     */
    getItemsWhere(filter: (item: T) => boolean): Promise<T[]>;
    /**
     * Returns the item in the list with the given `id`.
     * @param id The ID of the item to get.
     * @returns The item in the list with the given `id`.
     */
    getItemById(id: number): Promise<T | undefined>;
    /**
     * Creates a new `item` in the list.
     * @param item The item to create.
     * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
     */
    createItem(item: T): Promise<void>;
    /**
     * Updates the given `item` in the list.
     * @param item The item to update.
     * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
     */
    updateItem(item: T): Promise<void>;
    /**
     * Deletes the item with the given `id`.
     * @param id The ID of the item to delete
     * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
     */
    deleteItem(id: number): Promise<void>;
}
