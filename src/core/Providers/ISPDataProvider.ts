import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPListProvider } from "./ISPListProvider";

export interface ISPDataProvider{
    siteAbsoluteUrl: string;
    serverRelativeUrl: string;
    spList: ISPListProvider;
    context: WebPartContext;
}
