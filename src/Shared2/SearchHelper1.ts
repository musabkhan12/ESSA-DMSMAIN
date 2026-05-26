import { AggregationOption, SearchHit, SearchHitsContainer, SearchQuery, SearchRequest, SearchResponse } from '@microsoft/microsoft-graph-types';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { ISearchHitResource, ISearchQuery } from './SearchHelperInterfaces';
//import * as msgraphtypes from '@microsoft/microsoft-graph-types';
//import { version } from 'os';

export interface ISearchResult {
  name: string;
  webUrl: string;
  lastModifiedDateTime: string;
  createdBy: string;
}

export class GraphSearchHelper {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;    
  }

  /**
   * Searches SharePoint files using Microsoft Graph Search API
   * @param query - Search query string
   * @param size - Number of results to return
   * @returns Promise<ISearchResult[]> - List of search results
   */
  // public async searchFiles(query: string, size: number = 100,refiners:string[]=[]): Promise<SearchHit[]> {
  //   try {
  //     const searchaggegators:AggregationOption[]=['lastModifiedDateTime'].concat(refiners).map(ref=>{return  ({
  //       "field": ref,
  //       "size": 100,
  //       "bucketDefinition": {
  //           "sortBy": "count",                  
  //           "minimumCount": 1,
  //           "isDescending":true                    
  //       }          
  //     })} );
  //    const requestBody:ISearchQuery = {
  //       requests: [
  //         {
  //           entityTypes: ['driveItem'],
  //           query: {
  //             queryString: query
  //           },
  //           from: 0,
  //           size: size,            
  //           fields:['name','size','lastModifiedDateTime','weburl'].concat(refiners),
  //           aggregations:searchaggegators
  //           // [
  //           //   {
  //           //     "field": "RefinableString00",
  //           //     "size": 100,
  //           //     "bucketDefinition": {
  //           //         "sortBy": "count",                  
  //           //         "minimumCount": 1,
  //           //         "isDescending":true                    
  //           //     }          
  //           //   },
  //           //   {
  //           //     "field": "lastModifiedDateTime",
  //           //     "size": 100,
  //           //     "bucketDefinition": {
  //           //         "sortBy": "count",                  
  //           //         "minimumCount": 1,
  //           //         "isDescending":true                    
  //           //     }          
  //           //   }  

  //           // ],
  //           //aggregationFilters:["RefinableString00:\"ǂǂ7465737435\""]

  //           //fields: ['title', "path","url"],
  //           // sortProperties: [{ name: 'lastModifiedDateTime', isDescending: true }]
  //         }
  //       ]
  //     };

  //     const response = await this.graphClient.api('/search/query').post(requestBody);
  //     console.log(response);
      

  //     let hitcont=(response.value[0] as SearchResponse).hitsContainers[0];
  //     const resulthits=(hitcont.hits)?hitcont.hits:[]
  //     const results: Partial<ISearchHitResource>[] = (hitcont.hits)?hitcont.hits.map((hit) => {
  //       const resource:Partial<ISearchHitResource> = hit.resource as ISearchHitResource;   
             
  //       return resource;
  //       // return {
  //       //   name: resource.name,
  //       //   webUrl: resource.webUrl,
  //       //   lastModifiedDateTime: resource.lastModifiedDateTime,
  //       //   createdBy: resource.createdBy?.user?.displayName || 'Unknown'
  //       // };
  //     }):[];

  //     // return results;
  //     return resulthits;
  //   } catch (error) {
  //     console.error('Error searching files:', error);
  //     throw error;
  //   }
  // }

  public async searchFiles(query: string, size: number = 100): Promise<SearchHit[]> {
    try {
      const requestBody = {
        requests: [
          {
            entityTypes: ['driveItem'],
            query: {
              queryString: query
            },
            from: 0,
            size: size,
            trimDuplicates: true
            //fields: ['title', "path","url"],
            // sortProperties: [{ name: 'lastModifiedDateTime', isDescending: true }]
          }
        ]
      };
      console.log(requestBody , "DMS Search helper ");
      const response = await this.graphClient.api('/search/query').post(requestBody);
      
      debugger
      console.log(response);
      

      let hitcont=(response.value[0] as SearchResponse).hitsContainers[0];
      const resulthits=(hitcont.hits)?hitcont.hits:[]
      const results: Partial<ISearchHitResource>[] = (hitcont.hits)?hitcont.hits.map((hit) => {
        const resource:Partial<ISearchHitResource> = hit.resource as ISearchHitResource;   
             
        return resource;
        // return {
        //   name: resource.name,
        //   webUrl: resource.webUrl,
        //   lastModifiedDateTime: resource.lastModifiedDateTime,
        //   createdBy: resource.createdBy?.user?.displayName || 'Unknown'
        // };
      }):[];

      // return results;
      return resulthits;
    } catch (error) {
      console.error('Error searching files:', error);
      throw error;
    }
  }

  public async searchAll(query: string, size: number = 100,refiners:string[]=[],refinerfilters:string[]=[]): Promise<SearchHitsContainer> {
    try {
      const searchaggegators:AggregationOption[]=[].concat(refiners).map(ref=>{return  ({
        "field": ref,
        "size": 100,
        "bucketDefinition": {
            "sortBy": "count",                  
            "minimumCount": 1,
            "isDescending":true                    
        }          
      })} );
     
     let searchrequestparam:SearchRequest= {
      entityTypes: ['driveItem'],
      query: {
        queryString: query
      },
      from: 0,
      size: size,            
      fields:['name','size','lastModifiedDateTime','webUrl','createdBy','FileType','FileExtension'].concat(refiners),
      aggregations:searchaggegators,
      // [
      //   {
      //     "field": "RefinableString00",
      //     "size": 100,
      //     "bucketDefinition": {
      //         "sortBy": "count",                  
      //         "minimumCount": 1,
      //         "isDescending":true                    
      //     }          
      //   },
      //   {
      //     "field": "lastModifiedDateTime",
      //     "size": 100,
      //     "bucketDefinition": {
      //         "sortBy": "count",                  
      //         "minimumCount": 1,
      //         "isDescending":true                    
      //     }          
      //   }  

      // ],
      //aggregationFilters:refinerfilters

      //fields: ['title', "path","url"],
      // sortProperties: [{ name: 'lastModifiedDateTime', isDescending: true }]
    };

    if(refinerfilters && refinerfilters.length>0) searchrequestparam['aggregationFilters']=refinerfilters;

     const requestBody:ISearchQuery = {
        requests: [searchrequestparam]
      };

      const response = await this.graphClient.api('/search/query').post(requestBody);
      console.log(response);
      

      let hitcont=(response.value[0] as SearchResponse).hitsContainers[0];
      // const resulthits=(hitcont.hits)?hitcont.hits:[]
      // const results: Partial<ISearchHitResource>[] = (hitcont.hits)?hitcont.hits.map((hit) => {
      //   const resource:Partial<ISearchHitResource> = hit.resource as ISearchHitResource;   
             
      //   return resource;
      //   // return {
      //   //   name: resource.name,
      //   //   webUrl: resource.webUrl,
      //   //   lastModifiedDateTime: resource.lastModifiedDateTime,
      //   //   createdBy: resource.createdBy?.user?.displayName || 'Unknown'
      //   // };
      // }):[];

      // return results;
      return hitcont;
    } catch (error) {
      console.error('Error searching files:', error);
      throw error;
    }
  }

  /**
   * Helper function to format results for display
   * @param results - Array of ISearchResult
   * @returns string - Formatted results
   */
  public formatResults(results: ISearchResult[]): string {
    return results.map(result => `Name: ${result.name}\nURL: ${result.webUrl}\nLast Modified: ${result.lastModifiedDateTime}\nCreated By: ${result.createdBy}\n`).join('\n');
  }
}


//-----------------------------------------------
