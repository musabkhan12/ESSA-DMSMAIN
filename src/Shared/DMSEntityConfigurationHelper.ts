//Declare Imprts

import { BaseWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { getSP } from "../webparts/advancedSearch/loc/pnpjsConfig";
import { IListInfo } from "@pnp/sp/lists";
import "@pnp/sp/sites";
import "@pnp/sp/sites";



///

//Constants

export const LIST_TITLE_DMSEntitySitesMaster="MasterSiteURL";
export const LIST_TITLE_DMSFolderMaster="DMSFolderMaster";
export const LIST_TITLE_DMSEntityDocumentLibs="DMSPreviewFormMaster";

//Interface 

export interface IDMSEntity
{
    Title?:string;
    SiteID?:string;
    Active?:string;
    ID?:number;  
    FileMasterList?:string;     
    SiteURL?:string;
}

export interface IDMSEntityDocLib
{
    Title?:string;
    // SiteID?:string;
    ID?:number;  
    SiteTitle?:string;
    DocumentLibraryName?:string;
    FolderPath?:string;
    ListID?:string;

}

export interface IDMSEntityList
{
    ColumnName?:string;
    ColumnType?:string;
    SiteName?:string;
    DocumentLibraryName?:string;
    SearchMappedManagedProperty?:string;
    ID?:number;
     
}

export class DMSEntityConfigurationHelper
{
     sp: SPFI;
    constructor(wpcontext?:BaseWebPartContext)
    {
        this.sp=getSP(wpcontext as WebPartContext);
    }

    GetActivEnitities=async():Promise<IDMSEntity[]>=>
    {
        const FilesItems:IDMSEntity[] = await this.sp.web.lists.getByTitle(LIST_TITLE_DMSEntitySitesMaster)
         .items//.select("Title", "SiteID", "FileMasterList", "Active")
         .filter(`Active eq 'Yes'`)();

         return FilesItems;

    }

    
    

    GetDocLibs=async(siteid:string):Promise<IListInfo[]>=>
    {
        
        const subsiteWeb = await this.sp.site.openWebById(siteid);
        const libraries = await subsiteWeb.web.lists
        .select('ID', 'Title', 'BaseTemplate')
        .filter("BaseTemplate eq 101 and Hidden eq false")();
        return libraries;
    }

    GetActiveEntityListsAndColumns=async():Promise<IDMSEntityList[]>=>
    {
        let actent=await this.GetActivEnitities();
           // Collect promises for fetching document libraries
        let promises = actent.map(async (ent) => {
            return await this.sp.web.lists
                .getByTitle(LIST_TITLE_DMSEntityDocumentLibs)
                .items.select("ColumnName", "ID", "SearchMappedManagedProperty")
                .filter(`SiteName eq '${ent.Title}'`)();
        });

        // Resolve all promises and flatten the results
        let results = (await Promise.all(promises)).flat();

         // Filter results based on SearchMappedManagedProperty
        return results.filter(ent => ent.SearchMappedManagedProperty);
    }

    GetActiveEntityDocLibs=async():Promise<IDMSEntityDocLib[]>=>
        {
            let actent=await this.GetActivEnitities();
               // Collect promises for fetching document libraries
            let promises = actent.map(async (ent) => {
                let doclibs= await this.sp.web.lists
                    .getByTitle(LIST_TITLE_DMSFolderMaster)
                    .items.select("SiteTitle", "ID", "DocumentLibraryName")
                    .filter(`IsLibrary eq 1 and SiteTitle eq ${ent.Title}`)();
                 return doclibs.map(d=>{
                    let d1=d;
                    d1['SiteID']=ent.SiteID;
                    return d1;
                })   

            });
    
            // Resolve all promises and flatten the results
            let results = (await Promise.all(promises)).flat();
    
             // Filter results based on SearchMappedManagedProperty
            return results;
        }

        GetActiveEntityDocLibsBySiteId=async(sitetitle:string,siteId?:string):Promise<IDMSEntityDocLib[]>=>
        {
            
            let doclibs:IDMSEntityDocLib[]= await this.sp.web.lists
                .getByTitle(LIST_TITLE_DMSFolderMaster)
                .items//.select("SiteTitle", "ID", "DocumentLibraryName")
                .filter(`IsLibrary eq 1 and SiteTitle eq '${sitetitle}'`)();
            
            let alldoclibsofsite:IListInfo[]=[];
            if(siteId) 
             {
                alldoclibsofsite=await this.GetDocLibs(siteId);
                doclibs=doclibs.map(d=>{
                    let ad=alldoclibsofsite.find(a=>a.Title==d.DocumentLibraryName);
                    if(ad) d.ListID=ad.Id;
                    return d;
                });
             }
            return doclibs
        }

        GetActiveEntityDocLibsSearchFields=async(sitetitle:string,doclibtitle:string):Promise<IDMSEntityList[]>=>
        {
            
            let doclibsfields= await this.sp.web.lists
                .getByTitle(LIST_TITLE_DMSEntityDocumentLibs)
                .items.select("SiteName", "ID", "DocumentLibraryName", "ColumnName","ColumnType","SearchMappedManagedProperty")
                .filter(`DocumentLibraryName eq '${doclibtitle}' and SiteName eq '${sitetitle}'`)();
                
            return doclibsfields.filter(d=>d.SearchMappedManagedProperty);
        }

}



