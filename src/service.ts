import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import type { BaseItem } from "./base_item";
import type { IAttachmentInfo } from "@pnp/sp/attachments";

/**
 * A base class that exposes the most basic operations possible.
 *
 * This class should be inherited by each service of each list.
 */
export class Service<T extends BaseItem> {
  /**
   * The Sharepoint object provided by the `@pnp/sp` package.
   */
  private sp: SPFI;

  /**
   * The title of the list for this service.
   */
  private listTitle: string;

  /**
   * The cached items of the list for this service.
   */
  private items: T[] | undefined;

  /**
   * Creates a base service object with basic operations.
   * @param sp The Sharepoint object provided by the `@pnp/sp` package.
   * @param listTitle The title of the list for this service.
   */
  constructor(sp: SPFI, listTitle: string) {
    this.sp = sp;
    this.listTitle = listTitle;
  }

  /**
   * Returns all the items contained in the list.
   * @param fetchAttachmentFiles Tells whether this method must fetch the attachment files of each item in the list.
   * @returns All the items contained in the list.
   */
  async getAllItems(fetchAttachmentFiles?: boolean): Promise<T[]> {
    if (this.items) return this.items;

    this.items = await this.sp.web.lists
      .getByTitle(this.listTitle)
      .items.top(5000)<T[]>();

    // If the argument is not defined, the value is false by default.
    fetchAttachmentFiles ??= false;

    if (fetchAttachmentFiles) {
      this.items.forEach((item) =>
        this.fetchAttachmentFiles(item.ID).then(
          (attachmentFiles) => (item.AttachmentFiles = attachmentFiles)
        )
      );
    }

    return this.items;
  }

  /**
   * Returns all the attachment files for the item with the given `id`.
   * @param id The ID of the item for which the attachment files must be fetched.
   * @returns An array of the attachment files for the item with the given `id`.
   */
  private async fetchAttachmentFiles(id: number): Promise<IAttachmentInfo[]> {
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
  async getItemsWhere(filter: (item: T) => boolean): Promise<T[]> {
    if (!this.items) {
      await this.getAllItems();
    }

    const items = this.items!;
    return items.filter(filter);
  }

  /**
   * Returns the item in the list with the given `id`.
   * @param id The ID of the item to get.
   * @returns The item in the list with the given `id`.
   */
  async getItemById(id: number): Promise<T | undefined> {
    if (this.items) {
      return this.items.filter((item) => item.ID === id)[0];
    }

    return this.sp.web.lists.getByTitle(this.listTitle).items.getById(id)<T>();
  }

  /**
   * Creates a new `item` in the list.
   * @param item The item to create.
   * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
   */
  async createItem(item: T): Promise<void> {
    return this.sp.web.lists.getByTitle(this.listTitle).items.add(item);
  }

  /**
   * Updates the given `item` in the list.
   * @param item The item to update.
   * @returns A `Promise` that returns nothing and that will resolve when the operation is over.
   */
  async updateItem(item: T): Promise<void> {
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
  async deleteItem(id: number): Promise<void> {
    return this.sp.web.lists
      .getByTitle(this.listTitle)
      .items.getById(id)
      .delete();
  }
}
