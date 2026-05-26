import { SearchHit, SearchResponse } from '@microsoft/microsoft-graph-types';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { ISearchHitResource } from './SearchHelperInterfaces';
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
