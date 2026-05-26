import { SearchRequest } from "@microsoft/microsoft-graph-types"

export interface ISearchHitResource {
    "@odata.type": string
    size: number
    fileSystemInfo: FileSystemInfo
    listItem: ListItem
    id: string
    createdBy: CreatedBy
    createdDateTime: string
    lastModifiedBy: LastModifiedBy
    lastModifiedDateTime: string
    name: string
    parentReference: ParentReference
    webUrl: string
  }
  
  export interface FileSystemInfo {
    createdDateTime: string
    lastModifiedDateTime: string
  }
  
  export interface ListItem {
    "@odata.type": string
    id: string
    fields: Fields
  }
  
  export interface Fields {}
  
  export interface CreatedBy {
    user: User
  }
  
  export interface User {
    displayName: string
    email: string
  }
  
  export interface LastModifiedBy {
    user: User2
  }
  
  export interface User2 {
    displayName: string
    email: string
  }
  
  export interface ParentReference {
    driveId: string
    id: string
    sharepointIds: SharepointIds
    siteId: string
  }
  
  export interface SharepointIds {
    listId: string
    listItemId: string
    listItemUniqueId: string
  }
  
  export interface ISearchQuery
  {
    requests:SearchRequest[];
  }