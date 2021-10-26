import { 
    IAttachmentFileInfo, 
    ICamlQuery, 
    IFolder, 
    IFolderAddResult, 
    IListItemFormUpdateValue, 
    PagedItemCollection, 
    sp 
} from "@pnp/sp/presets/all";

import { IBaseModel } from "../Models/IBaseModel";

export interface ISPListProvider{
    getById(
        ID: number, 
        listRelativeUrl: string, 
        rootWeb?: boolean
    ): Promise<IBaseModel>;

    getItemsByFilter(
        listRelativeUrl: string,
        filter?: string,
        rootWeb?: boolean
    ): Promise<Array<IBaseModel>>

    getItems(
        listRelativeUrl: string,
        filter?: string,
        select?: string[],
        expand?:string[],
        order?: boolean,
        elementOrder?: string,
        rootWeb?: boolean
    ): Promise<IBaseModel[]>

    getListItemsCount(
        listRelativeUrl: string,
        rootWeb?: boolean
    ): Promise<number>;

    getItemsByFilterInLargeLists(
        listRelativeUrl: string,
        filter: string
    ): Promise<Array<IBaseModel>>;

    getItemsPaged(
        listRelativeUrl: string,
        top: number,
        filter?: string,
        order?: boolean,
        elementOrder?: string,
        rootWeb?: boolean
    ): Promise<PagedItemCollection<IBaseModel[]>>;

    getLastItemId(
        listRelativeUrl: string, 
        rootWeb?: boolean
    ): Promise<number>;

    getLastItem(
        listRelativeUrl: string, 
        filter?: string, 
        rootWeb?: boolean
    );

    save(
        item: IBaseModel, 
        listRelativeUrl: string, 
        rootWeb?: boolean
    ): Promise<IBaseModel>;

    addAttachment(
        item: IBaseModel,
        listRelativeUrl: string,
        attachments: IAttachmentFileInfo[],
        rootWeb?: boolean
    ): Promise<void>;

    getAttachments(
        item: IBaseModel, 
        listRelativeUrl: string, 
        rootWeb?: boolean
    ): Promise<any[]>;

    deleteAttachment(
        item: IBaseModel,
        listRelativeUrl: string,
        attachments: string[],
        rootWeb?: boolean
    ): Promise<void>;

    breakListPermission(
        listRelativeUrl: string,
        copyRoleAssignments?: boolean,
        clearSubscopes?: boolean,
        rootWeb?: boolean
    ): Promise<any>;

    saveToFolder(
        item: IBaseModel,
        listRelativeUrl: string,
        folderName: string,
        rootWeb?: boolean
    ): Promise<IListItemFormUpdateValue[]>;

    moveItemToFolder(
        item: IBaseModel,
        toFolderName: string,
        listRelativeUrl: string,
        originalFolderName?: string,
        rootWeb?: boolean
    ): Promise<void>;

    createFolder(
        folderName: string,
        siteAbsoluteUrl: string,
        listRelativeUrl: string,
        rootWeb?: boolean
    ): Promise<boolean>;

    breakFolderPermission(
        folderName: string,
        listRelativeUrl: string,
        copyRoleAssignments?: boolean,
        clearSubscopes?: boolean,
        permissions?: { [key: number]: number }[],
        permissionsRemove?: { [key: number]: number }[],
        rootWeb?: boolean
    ): Promise<any>;

    createFolderDocumentLibrary(
        folderName: string,
        listRelativeUrl: string,
        rootWeb?: boolean
    ): Promise<IFolder>;

    deleteFolderDocumentLibrary(
        folderName: string,
        listRelativeUrl: string,
        rootWeb?: boolean
    ): Promise<void>;

    delete(
        itemID: number,
        listRelativeUrl: string,
        rootWeb?: boolean
    ): Promise<void>;

    getItemsWithAttachmentsFiltered(
        listRelativeUrl: string, 
        expanded: string, 
        filter?: string, 
        rootWeb?: boolean
    ): Promise<Array<IBaseModel>>;

    getItemsByCAMLQueryXML(
        listRelativeUrl: string, 
        CAMLQuery: ICamlQuery, 
        rootWeb?: boolean
    ): Promise<Array<IBaseModel>>;
}
