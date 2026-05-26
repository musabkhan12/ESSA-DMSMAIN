import * as React from 'react';
//import styles from './AdvancedSearch.module.scss';
//import type { IAdvancedSearchProps } from './IAdvancedSearchProps';
import { escape } from '@microsoft/sp-lodash-subset';
import "bootstrap/dist/css/bootstrap.min.css";
//import { GraphSearchHelper } from '../../../Shared/SearchHelper1';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { IDocumentDisplayFields, ResultsViewMode } from './Interfaces';
import { GraphSearchHelper } from '../../../../Shared/SearchHelper1';
import { ISearchHitResource } from '../../../../Shared/SearchHelperInterfaces';
import { DMSResultsGrid } from './DMSResultsGrid';
//import { IDocumentDisplayFields } from '../../essadmsMain/components/DMSSearch/Interfaces';
//import { ISearchHitResource } from '../../../Shared/SearchHelperInterfaces';


export interface IDMSEntitySearchProps {
    context:BaseWebPartContext
    searchtext: string;
    searchfilter:string;
    searchpath:string;
}

export interface IDMSEntitySearchState {       
    searchresult: IDocumentDisplayFields[];
}

export const getUrlParameter=(name:string)=>{
  const urlParams = new URLSearchParams(window.location.search);
  return urlParams.get(name);
}

export default class DMSEntitySearch extends React.Component<IDMSEntitySearchProps,IDMSEntitySearchState> {
  
  public constructor(props: IDMSEntitySearchProps) {
    super(props);
    this.state = {
    //   searchtext: "",
    //   searchfilter: "",
    
      searchresult: []
    };
  }

   
   componentDidMount(): void {
   
    // let _searchquery,_searchpath;
    // let searchqueryfromurl=getUrlParameter("searchquery");
    // let searchpathfromurl=getUrlParameter("searchpath");
    // _searchquery=(searchqueryfromurl)?decodeURIComponent(searchqueryfromurl):this.state.searchfilter;
    // _searchpath=(searchpathfromurl)?decodeURIComponent(searchpathfromurl):this.state.searchpath;
     this.runSearch(this.props.searchtext,this.props.searchfilter,this.props.searchpath);
      
   }

   componentDidUpdate(prevProps: Readonly<IDMSEntitySearchProps>, prevState: Readonly<IDMSEntitySearchState>, snapshot?: any): void {
     //this.runSearch(this.props.searchtext,this.props.searchfilter,this.props.searchpath);     
     if(this.props.searchfilter.toLowerCase()!=prevProps.searchfilter.toLowerCase() 
      || this.props.searchpath.toLowerCase()!=prevProps.searchpath.toLowerCase() 
      || this.props.searchtext.toLowerCase()!=prevProps.searchtext.toLowerCase() )
     {
       this.runSearch(this.props.searchtext,this.props.searchfilter,this.props.searchpath);  
     }
   }

  runSearch=async (searchtext:string,searchFilters:string,searchPath:string)=>{

    
    let qyerytext=`${searchtext} IsDocument:True ${searchFilters} Path:"${searchPath}"`;
    let graphcl=await (this.props.context as BaseWebPartContext).msGraphClientFactory.getClient("3");
    let mssearch=new GraphSearchHelper(graphcl);
    
    let searchres=await mssearch.searchFiles(qyerytext,500);

    let resultsdoc: IDocumentDisplayFields[] = searchres.map(filehit => {
      let file:Partial<ISearchHitResource>=filehit.resource;
      let tRes: IDocumentDisplayFields = {
        Title: file.name,
        Size: file.size,
        Extension: file.name.split('.').pop(),
        Path:file.webUrl,
        Summary:filehit.summary
      }
      return tRes;
    })

    this.setState({searchresult:resultsdoc});
  } 

  SearchClickHandler:React.MouseEventHandler=(ev)=>{

    ev.preventDefault();
    //this.setState({searchtext:(ev.target as HTMLInputElement).value});
    //this.runSearch(this.state.searchtext,this.state.searchfilter,this.state.searchpath);


  }

  SearchTextChangeHandler:React.ChangeEventHandler=(ev)=>
  {
    //this.setState({searchtext:(ev.target as HTMLInputElement).value});

  }
  public render(): React.ReactElement<IDMSEntitySearchProps> {
    //let st =this.props.searchtext;
    //let st1 =this.props.searchpath;

    //this.runSearch(st,"",st1);
    return (
        <>
        <DMSResultsGrid ViewMode={ResultsViewMode.Grid} results={this.state.searchresult} />
        </>
      
    );
  }
}
