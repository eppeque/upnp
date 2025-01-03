import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
/**
 * A base class that exposes the most basic operations possible.
 *
 * This class should be inherited by each service of each list.
 */
export class Service {
    /**
     * The Sharepoint object provided by the `@pnp/sp` package.
     */
    sp;
    /**
     * The title of the list for this service.
     */
    listTitle;
    /**
     * Creates a base service object with basic operations.
     * @param sp The Sharepoint object provided by the `@pnp/sp` package.
     * @param listTitle The title of the list for this service.
     */
    constructor(sp, listTitle) {
        this.sp = sp;
        this.listTitle = listTitle;
    }
    /**
     * Returns all the items contained in the list.
     * @param fetchAttachmentFiles Tells whether this method must fetch the attachment files of each item in the list.
     * @returns All the items contained in the list.
     */
    async getAllItems(fetchAttachmentFiles) {
        const items = await this.sp.web.lists
            .getByTitle(this.listTitle)
            .items.top(5000)();
        // If the argument is not defined, the value is false by default.
        fetchAttachmentFiles ??= false;
        if (fetchAttachmentFiles) {
            items.forEach((item) => this.fetchAttachmentFiles(item.ID).then((attachmentFiles) => (item.AttachmentFiles = attachmentFiles)));
        }
        return items;
    }
    /**
     * Returns all the attachment files for the item with the given `id`.
     * @param id The ID of the item for which the attachment files must be fetched.
     * @returns An array of the attachment files for the item with the given `id`.
     */
    async fetchAttachmentFiles(id) {
        return this.sp.web.lists
            .getByTitle(this.listTitle)
            .items.getById(id)
            .attachmentFiles();
    }
    /**
     * Returns the items in the list that respect the `filter` callback.
     * @param filter The filter callback that determines whether an item should be included in the result or not.
     * @returns The items for which the `filter` callback returned `true`
     */
    async getItemsWhere(filter) {
        const items = await this.getAllItems();
        return items.filter(filter);
    }
    /**
     * Returns the item in the list with the given `id`.
     * @param id The ID of the item to get.
     * @returns The item in the list with the given `id`.
     */
    async getItemById(id) {
        return this.sp.web.lists.getByTitle(this.listTitle).items.getById(id)();
    }
    /**
     * Creates a new `item` in the list.
     * @param item The item to create.
     * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
     */
    async createItem(item) {
        return this.sp.web.lists.getByTitle(this.listTitle).items.add(item);
    }
    /**
     * Updates the given `item` in the list.
     * @param item The item to update.
     * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
     */
    async updateItem(item) {
        return this.sp.web.lists
            .getByTitle(this.listTitle)
            .items.getById(item.ID)
            .update(item);
    }
    /**
     * Deletes the item with the given `id`.
     * @param id The ID of the item to delete
     * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
     */
    async deleteItem(id) {
        return this.sp.web.lists
            .getByTitle(this.listTitle)
            .items.getById(id)
            .delete();
    }
}
//# sourceMappingURL=service.js.map