import { 
    IAttachmentFileInfo, 
    ICamlQuery, 
    IFolder, 
    IFolderAddResult, 
    IListItemFormUpdateValue, 
    PagedItemCollection, 
    sp } from "@pnp/sp/presets/all";
import { IBaseModel } from "../Models/IBaseModel";
import { CustomProperties } from "../Enums/Enums";
import { IListItemAttachmentFile } from "../Models/IListItemAttachmentFile";
import { ISPDataProvider } from "../Providers/ISPDataProvider";

export abstract class BaseServices {
  public abstract itemData: IBaseModel;
  public abstract itemsData: Array<IBaseModel>;
  public abstract listInternalName: string;

  protected _rootWeb: boolean;

  constructor(protected spDataProvider: ISPDataProvider, rootWeb: boolean = false) {
    this._rootWeb = rootWeb;
  }

  public get dataProvider(): ISPDataProvider{
    return this.spDataProvider;
  }

  get siteAbsoluteUrl(): string {
    return this.spDataProvider.siteAbsoluteUrl;
  }

  get listAbsoluteUrl(): string {
    return `${this.spDataProvider.siteAbsoluteUrl}/Lists/${this.listInternalName}`;
  }

  get listRelativeUrl(): string {
    return this._rootWeb
      ? `${this.spDataProvider.context.pageContext.site.serverRelativeUrl}/Lists/${this.listInternalName}`
      : `${this.spDataProvider.serverRelativeUrl}/Lists/${this.listInternalName}`;
  }

  public async loadItemData(ID: number): Promise<void> {
    let item = await this.spDataProvider.spList.getById(ID, this.listRelativeUrl, this._rootWeb);
    this.itemData = item;
  }

  public async loadItemsData(filter?: string): Promise<void> {
    let items = await this.spDataProvider.spList.getItemsByFilter(this.listRelativeUrl, filter, this._rootWeb);
    this.itemsData = items;
  }

  public async loadItemsDataLargeList(filter: string): Promise<void> {
    let items = await this.spDataProvider.spList.getItemsByFilterInLargeLists(
      this.listRelativeUrl,
      filter
    );
    this.itemsData = items;
  }

  public async saveItems(): Promise<IBaseModel[]> {
    try {
      let promisses = [];
      let indexArray = [];
      for (let index = 0; index < this.itemsData.length; index++) {
        let itemChanged = CustomProperties.ItemChanged;
        if (this.itemsData[index][itemChanged]) {
          delete this.itemsData[index][CustomProperties.ItemChanged]; //Control to update only changed item
          promisses.push(
            this.spDataProvider.spList.save(this.itemsData[index], this.listRelativeUrl, this._rootWeb)
          );
          indexArray.push(index);
          this.itemsData[index][CustomProperties.ItemChanged] = true;
        }
      }
      await Promise.all(promisses).then(item => {
        for (let index = 0; index < item.length; index++) {
          this.itemsData[indexArray[index]] = item[index];
        }
      });
      return this.itemsData;
    } catch (err) {
      console.log(err);
      return this.itemsData;
    }
  }

  public async saveItemsInFolder(folderName: string | number): Promise<IBaseModel[]> {
    try {
      let promissesNewItem = [];
      let promissesUpdateItem = [];
      for (let index = 0; index < this.itemsData.length; index++) {
        if (!this.itemsData[index].ID) {
          promissesNewItem.push(
            this.spDataProvider.spList.save(this.itemsData[index], this.listRelativeUrl, this._rootWeb)
          );
        } else {
          promissesUpdateItem.push(
            this.spDataProvider.spList.save(this.itemsData[index], this.listRelativeUrl, this._rootWeb)
          );
        }
      }
      let promissesMoveToFolder = [];
      await Promise.all(promissesNewItem).then(item => {
        for (let index = 0; index < item.length; index++) {
          this.itemsData[index] = item[index];
          promissesMoveToFolder.push(
            this.spDataProvider.spList.moveItemToFolder(
              this.itemsData[index],
              `${folderName}`,
              this.listRelativeUrl,
              "",
              this._rootWeb
            )
          );
        }
      });
      if (promissesMoveToFolder.length > 0) await Promise.all(promissesMoveToFolder);
      return this.itemsData;
    } catch {
      return this.itemsData;
    }
  }

  public async save(attachments?: IAttachmentFileInfo []): Promise<void> {
    let itemSaved = await this.spDataProvider.spList.save(this.itemData, this.listRelativeUrl, this._rootWeb);
    this.itemData = itemSaved;

    if (attachments)
      await this.spDataProvider.spList.addAttachment(
        this.itemData,
        this.listRelativeUrl,
        attachments,
        this._rootWeb
      );
  }

  public deleteAttachments(attachments: string[]): Promise<void> {
    return this.spDataProvider.spList.deleteAttachment(
      this.itemData,
      this.listRelativeUrl,
      attachments,
      this._rootWeb
    );
  }

  public async getAttachments(): Promise<IListItemAttachmentFile[]> {
    let result: IListItemAttachmentFile[] = [];
    let files: any[] = await this.spDataProvider.spList.getAttachments(
      this.itemData,
      this.listRelativeUrl,
      this._rootWeb
    );
    if (files && files.length > 0) {
      for (let index = 0; index < files.length; index++) {
        const file = files[index];
        result.push({
          FileName: file.FileName,
          ServerRelativeUrl: file.ServerRelativeUrl
        });
      }
    }
    return result;
  }

  public async saveInFolder(
    folderName: string | number,
    attachments?: IAttachmentFileInfo[]
  ): Promise<IBaseModel> {
    let notMove = this.itemData.ID > 0;
    let itemSaved = await this.spDataProvider.spList.save(this.itemData, this.listRelativeUrl, this._rootWeb);
    this.itemData = itemSaved;
    if (!notMove) await this.moveItemToFolder(`${folderName}`);
    if (attachments)
      await this.spDataProvider.spList.addAttachment(
        this.itemData,
        this.listRelativeUrl,
        attachments,
        this._rootWeb
      );
    return this.itemData;
  }

  public async moveItemToFolder(toFolderName: string, originalFolderName?: string): Promise<void> {
    return this.spDataProvider.spList.moveItemToFolder(
      this.itemData,
      toFolderName,
      this.listRelativeUrl,
      originalFolderName,
      this._rootWeb
    );
  }

  public async createFolder(folderName: string): Promise<Boolean> {
    return this.spDataProvider.spList.createFolder(
      folderName,
      this.siteAbsoluteUrl,
      this.listRelativeUrl,
      this._rootWeb
    );
  }

  public async breakFolderPermission(
    folderName: string,
    copyRoleAssignments?: boolean,
    clearSubscopes?: boolean,
    permissionsAdd?: { [key: number]: number }[],
    permissionsRemove?: { [key: number]: number }[]
  ): Promise<void> {
    return this.spDataProvider.spList.breakFolderPermission(
      folderName,
      this.listRelativeUrl,
      copyRoleAssignments,
      clearSubscopes,
      permissionsAdd,
      permissionsRemove,
      this._rootWeb
    );
  }

  public async createFolderDocumentLibrary(folderName: string): Promise<IFolder> {
    return this.spDataProvider.spList.createFolderDocumentLibrary(folderName, this.listRelativeUrl, this._rootWeb);
  }

  public async deleteFolderDocumentLibrary(folderName: string): Promise<void> {
    return this.spDataProvider.spList.deleteFolderDocumentLibrary(folderName, this.listRelativeUrl, this._rootWeb);
  }

  public async getItemsWithAttachmentsFiltered(expanded: string, filter?: string): Promise<void> {
    let items = await this.spDataProvider.spList.getItemsWithAttachmentsFiltered(this.listRelativeUrl, expanded, filter, this._rootWeb);
    this.itemsData = items;
  }

  public async loadItemsTopPaginate(top: number, filter?: string): Promise<void> {
    let items = await this.spDataProvider.spList.getItemsPaged(this.listRelativeUrl, top, filter, true, "ID", this._rootWeb);
    this.itemsData = items.results;
    while (items.hasNext) {
      items = await items.getNext();
      Array.prototype.push.apply(this.itemsData, items.results);
    }
  }

  public async loadItemsDataCAMLQuery(CAMLQuery: ICamlQuery): Promise<void> {
    let items = await this.spDataProvider.spList.getItemsByCAMLQueryXML(this.listRelativeUrl, CAMLQuery, this._rootWeb);
    this.itemsData = items;
  }
}