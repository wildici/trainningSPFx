import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPRest } from "@pnp/sp";
import { IAttachmentFileInfo, IFolder, IFolderAddResult, sp } from "@pnp/sp/presets/all";

import { ISPListProvider } from "./ISPListProvider";
import { IWeb } from "@pnp/sp/webs";
import { IItemAddResult, Item, PagedItemCollection } from "@pnp/sp/items";
import { ICamlQuery, IListItemFormUpdateValue } from "@pnp/sp/lists";
import { IBaseModel } from "../Models/IBaseModel";

export class SPListProvider implements ISPListProvider{

    constructor(protected readonly sp: SPRest, protected readonly webPartContext: WebPartContext) { }
    public readonly restItemLimit: number = 5000;

    public getById(ID: number, listRelativeUrl: string, rootWeb: boolean = false): Promise<IBaseModel>{
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return spWeb
            .getList(listRelativeUrl)
            .items.getById(ID)
            .get();
    }

    public getItemsByFilter(listRelativeUrl: string, filter?: string, rootWeb: boolean = false): Promise<Array<IBaseModel>> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return filter
          ? spWeb
            .getList(listRelativeUrl)
            .items.filter(filter)
            .get()
          : spWeb.getList(listRelativeUrl).items.get();
      }

      public getItems(listRelativeUrl: string, filter?: string, select?: string[], expand?: string[], order: boolean = true, elementOrder: string = "ID", rootWeb: boolean = false): Promise<IBaseModel[]> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        if (filter && select && expand) {
          return spWeb.getList(listRelativeUrl).items
            .filter(filter)
            .select(...select)
            .expand(...expand)
            .orderBy(elementOrder, order)
            .get();
        }
        else if (filter && select && !expand) {
          return spWeb.getList(listRelativeUrl).items
            .filter(filter)
            .select(...select)
            .orderBy(elementOrder, order)
            .get();
        }
        else if (filter && !select && !expand) {
          return spWeb.getList(listRelativeUrl).items
            .filter(filter)
            .orderBy(elementOrder, order)
            .get();
        }
    
        else if (filter && !select && expand) {
          return spWeb.getList(listRelativeUrl).items
            .filter(filter)
            .expand(...expand)
            .orderBy(elementOrder, order)
            .get();
        }
        else if (!filter && select && expand) {
          return spWeb.getList(listRelativeUrl).items
            .select(...select)
            .expand(...expand)
            .orderBy(elementOrder, order)
            .get();
        }
        else {
          return spWeb.getList(listRelativeUrl).items
            .expand(...expand)
            .orderBy(elementOrder, order)
            .get();
        }
      }

      public async getListItemsCount(
        listRelativeUrl: string,
        rootWeb: boolean = false
      ): Promise<number>{
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        
          const result = await spWeb.getList(listRelativeUrl).get();
          return result.ItemCount;
      }

      public async getItemsByFilterInLargeLists(
        listRelativeUrl: string,
        filter: string
      ): Promise<Array<IBaseModel>> {
        let results: Array<IBaseModel> = [];
        let lastID = await this.getLastItemId(listRelativeUrl);
        if (lastID < this.restItemLimit) {
          results = await this.getItemsByFilter(listRelativeUrl, filter);
          return results;
        } else {
          let firstIdBatch = 0;
          while (firstIdBatch < lastID) {
            let lastIdBatch =
              firstIdBatch + this.restItemLimit > lastID
                ? lastID
                : firstIdBatch + this.restItemLimit;
            let filterBatch = `ID ${
              lastIdBatch == lastID ? "le" : "lt"
              } ${lastIdBatch} and ID gt ${firstIdBatch} and (${filter})`;
            let batchResults = await this.getItemsByFilter(
              listRelativeUrl,
              filterBatch
            );
            firstIdBatch = lastIdBatch;
            results.push(...batchResults);
          }
          return results;
        }
      }

      public async getLastItemId(listRelativeUrl: string, rootWeb: boolean = false): Promise<number> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        let itemsResult = await spWeb
          .getList(listRelativeUrl)
          .items.select("ID")
          .top(1)
          .orderBy("ID", false)
          .get();
        return itemsResult && itemsResult.length > 0 ? itemsResult[0].ID : 0;
      }

      public async getLastItem(listRelativeUrl: string, filter?: string, rootWeb: boolean = false): Promise<IBaseModel> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        const results = filter
          ? await spWeb
            .getList(listRelativeUrl)
            .items.filter(filter)
            .top(1)
            .orderBy("ID", false)
            .get()
          : await spWeb
            .getList(listRelativeUrl)
            .items.top(1)
            .orderBy("ID", false)
            .get();
        return results ? results[0] : null;
      }

      public async save(item: IBaseModel, listRelativeUrl: string, rootWeb: boolean = false): Promise<IBaseModel> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        let propsToDelete: string[] = [];
        Object.keys(item).forEach(key => {
          if (key.indexOf("StringId") >= 0) {
            propsToDelete.push(key);
          }
        });
    
        for (let index = propsToDelete.length - 1; index >= 0; index--) {
          delete item[propsToDelete[index]];
        }
    
        if (!item.ID || item.ID <= 0) {
          const resultAdd: IItemAddResult = await spWeb
            .getList(listRelativeUrl)
            .items.add({ ...item });
          item = resultAdd.data;
        } else {
          try {
            await spWeb
              .getList(listRelativeUrl)
              .items.getById(item.ID)
              .update({ ...item });
          } catch (err) {
            console.log(err);
          }
        }
        return item;
      }

      public breakListPermission(
        listRelativeUrl: string,
        copyRoleAssignments?: boolean,
        clearSubscopes?: boolean,
        rootWeb: boolean = false
      ): Promise<any> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return spWeb
          .getList(listRelativeUrl)
          .breakRoleInheritance(copyRoleAssignments, clearSubscopes);
      }

      public async saveToFolder(
        item: IBaseModel,
        listRelativeUrl: string,
        folderName: string,
        rootWeb: boolean = false
      ): Promise<IListItemFormUpdateValue[]> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        let propsToDelete: string[] = [];
        Object.keys(item).forEach(key => {
          if (key.indexOf("StringId") >= 0) {
            propsToDelete.push(key);
          }
        });
    
        for (let index = propsToDelete.length - 1; index >= 0; index--) {
          delete item[propsToDelete[index]];
        }
    
        if (!item.ID || item.ID <= 0) {
          return spWeb
            .getList(listRelativeUrl)
            .addValidateUpdateItemUsingPath(
              this.convertObjToListItemFormUpdateValue(item),
              `${listRelativeUrl}/${folderName}`
            );
        } else {
          return spWeb
            .getList(listRelativeUrl)
            .items.getById(item.ID)
            .validateUpdateListItem(this.convertObjToListItemFormUpdateValue(item));
        }
      }

      public async moveItemToFolder(
        item: IBaseModel,
        toFolderName: string,
        listRelativeUrl: string,
        originalFolderName?: string,
        rootWeb: boolean = false
      ): Promise<void> {
        const spWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return spWeb
          .getFileByServerRelativeUrl(
            originalFolderName
              ? `${listRelativeUrl}/${originalFolderName}/${item.ID}_.000`
              : `${listRelativeUrl}/${item.ID}_.000`
          )
          .moveTo(`${listRelativeUrl}/${toFolderName}/${item.ID}_.000`);
      }

      public convertObjToListItemFormUpdateValue(item: any, rootWeb: boolean = false): IListItemFormUpdateValue[] {
        let props: IListItemFormUpdateValue[] = [];
        Object.keys(item).forEach(e => props.push({ FieldName: e, FieldValue: item[e] }));
        return props;
      }

      public async createFolder(
        folderName: string,
        siteAbsoluteUrl: string,
        listRelativeUrl: string,
        rootWeb: boolean = false
      ): Promise<boolean> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        try {
          await spWeb
            .getList(listRelativeUrl)
            .rootFolder.folders.getByName(folderName)
            .getItem();
        } catch {
          await spWeb.getList(listRelativeUrl).rootFolder.folders.add(folderName).then(response => {
            //return true;
            return response.data != null;
          }).catch(err => {
            console.log(err);
            return false;
          });
        }
        return false;
      }

      public async breakFolderPermission(
        folderName: string,
        listRelativeUrl: string,
        copyRoleAssignments?: boolean,
        clearSubscopes?: boolean,
        permissionsAdd?: { [key: number]: number }[],
        permissionsRemove?: { [key: number]: number }[],
        rootWeb: boolean = false
      ): Promise<any> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        await this.resetFolderPermission(folderName, listRelativeUrl);
        const folderItem = await spWeb
          .getList(listRelativeUrl)
          .rootFolder.folders.getByName(folderName)
          .getItem();
        await folderItem.breakRoleInheritance(copyRoleAssignments, clearSubscopes);
        let arrayPromises = [];
        if (permissionsAdd) {
          permissionsAdd.forEach(p => {
            Object.keys(p).forEach(e =>
              arrayPromises.push(folderItem.roleAssignments.add(parseInt(e), p[e]))
            );
          });
        }
        if (permissionsRemove) {
          permissionsRemove.forEach(p => {
            Object.keys(p).forEach(e =>
              arrayPromises.push(folderItem.roleAssignments.remove(parseInt(e), p[e]))
            );
          });
        }
        if (arrayPromises.length > 0) return Promise.all(arrayPromises);
        else return new Promise<any>(resolve => resolve({}));
      }

      public async resetFolderPermission(folderName: string, listRelativeUrl: string, rootWeb: boolean = false): Promise<void> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        const folderItem = await spWeb
          .getList(listRelativeUrl)
          .rootFolder.folders.getByName(folderName)
          .getItem();
        await folderItem.resetRoleInheritance();
      }

      public async createFolderDocumentLibrary(
        folderName: string,
        listRelativeUrl: string,
        rootWeb: boolean = false
      ): Promise<IFolder> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        let newFolder: IFolder;
        try {
          newFolder = await spWeb
            .getList(listRelativeUrl)
            .rootFolder.folders.getByName(folderName)
            .get();
        } catch (e) {
          let folderAddResult: IFolderAddResult = await spWeb
            .getList(listRelativeUrl)
            .rootFolder.folders.add(folderName);
          newFolder = folderAddResult.folder;
        }
    
        return newFolder;
      }

      public async deleteFolderDocumentLibrary(
        folderName: string,
        listRelativeUrl: string,
        rootWeb: boolean = false
      ): Promise<void> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return spWeb
          .getList(listRelativeUrl)
          .rootFolder.folders.getByName(folderName)
          .delete();
      }

      public async delete(
        itemID: number,
        listRelativeUrl: string,
        rootWeb: boolean = false
      ): Promise<void> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return spWeb
          .getList(listRelativeUrl)
          .items.getById(itemID)
          .delete();
      }

      public async addAttachment(
        item: IBaseModel,
        listRelativeUrl: string,
        attachments: IAttachmentFileInfo[],
        rootWeb: boolean = false
      ): Promise<void> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        if (item.ID > 0) {
          if (attachments.length > 0) {
            await spWeb
              .getList(listRelativeUrl)
              .items.getById(item.ID)
              .attachmentFiles.addMultiple(attachments);
          }
        }
      }

      public async deleteAttachment(
        item: IBaseModel,
        listRelativeUrl: string,
        attachments: string[],
        rootWeb: boolean = false
      ): Promise<void> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        if (item.ID > 0) {
          if (attachments.length > 0) {
            await spWeb
              .getList(listRelativeUrl)
              .items.getById(item.ID)
              .attachmentFiles.recycleMultiple(...attachments);
          }
        }
      }

      public async getAttachments(item: IBaseModel, listRelativeUrl: string, rootWeb: boolean = false): Promise<any[]> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        if (item.ID > 0) {
          return await spWeb
            .getList(listRelativeUrl)
            .items.getById(item.ID)
            .attachmentFiles.get();
        }
        return null;
    }

    public getItemsWithAttachmentsFiltered(listRelativeUrl: string, expanded: string, filter?: string, rootWeb: boolean = false): Promise<Array<IBaseModel>> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return filter
          ? spWeb
            .getList(listRelativeUrl)
            .items.filter(filter)
            .expand(expanded)
            .get()
          : spWeb.getList(listRelativeUrl).items.expand(expanded).get();
    }

    public getItemsPaged(listRelativeUrl: string, top: number, filter?: string, order: boolean = true, elementOrder: string = "ID", rootWeb: boolean = false): Promise<PagedItemCollection<IBaseModel[]>> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return filter
          ? spWeb.getList(listRelativeUrl).items
            .top(top)
            .filter(filter)
            .orderBy(elementOrder, order)
            .getPaged()
          : spWeb.getList(listRelativeUrl).items
            .top(top)
            .orderBy(elementOrder, order)
            .getPaged();
    }

    public getItemsByCAMLQueryXML(listRelativeUrl: string, CAMLQuery: ICamlQuery, rootWeb: boolean = false): Promise<IBaseModel[]> {
        const spWeb: IWeb = rootWeb ? this.sp.site.rootWeb : this.sp.web;
        return spWeb.getList(listRelativeUrl).getItemsByCAMLQuery(CAMLQuery);
    }
}