import * as React from 'react';
import { useState, useEffect } from 'react';
// import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/sites";
import "bootstrap/dist/css/bootstrap.min.css";
import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
// import { getSP } from '../../essadmsMain/loc/pnpjsConfig';
import { getSP } from '../../../webparts/advancedSearch/loc/pnpjsConfig';
import { DMSEntityConfigurationHelper, IDMSEntity, IDMSEntityDocLib } from '../../../Shared/DMSEntityConfigurationHelper';
import { IContextInfo } from '@pnp/sp/context-info';
// import { CheckedDropdownWithSearch, ICheckedBoxItem } from './CheckedDropDown';
// import CheckedFilterDropDown from './CheckedFilterDropDown';
import CheckedFilterDropDown from './CheckedFilterDropDown';
import { fieldnamesmapping } from './Common';
// import { IContextInfo } from "@pnp/sp/sites";

export interface IDropDownOption {
  key: string;
  text: string;
  children?: any[];
  checked?: boolean;
  data?:any  
}

export enum enumfieldtype
{
  SingleLineofText='Single Line of Text',
  MultipleLineofText='Multiple Line of Text',
  YesNo='Yes or No',
  DateTime='Date & Time',
  Number='Number'
}



export interface IDMSEntitySearchDropDownsProps
{
  context:BaseWebPartContext  
  onFieldSelect?: (field: string,fieldtype?:enumfieldtype) => void;
  onMultiFieldSelect?: (selectedfields:{field: string,fieldtype?:enumfieldtype}[]) => void;
  onSiteSelect?:(selectedsites:IDMSEntity[]) => void; 
  onLibrarySelect?:(selectedlibraries:IDMSEntityDocLib[]) => void; 
  
}

export const DMSSearchFilterCheckedDropDown: React.FC<IDMSEntitySearchDropDownsProps> = (props:IDMSEntitySearchDropDownsProps) => {
  // const [treeData, setTreeData] = useState<TreeNode[]>([]);
  // const [selectedLibrary, setSelectedLibrary] = useState<string | null>(null);
  // const [fields, setFields] = useState<{ key: string; text: string, type?:string }[]>([]);
  // const [selectedField, setSelectedField] = useState<string | null>(null);

  const [dropdown1, setDropdown1] = useState<IDropDownOption[]>([]);
  const [dropdown2, setDropdown2] = useState<IDropDownOption[]>([]);
  const [dropdown3, setDropdown3] = useState([]);
  const [dropdown4, setDropdown4] = useState<IDropDownOption[]>([]);
  const [selected1, setSelected1] = useState("");
  const [selected2, setSelected2] = useState("");
  const [selected3, setSelected3] = useState("");
  const [selected4, setSelected4] = useState("");


  const sp: SPFI = getSP(props.context as WebPartContext);
  const confighelper=new DMSEntityConfigurationHelper(props.context);
  useEffect(() => {
    // Fetch sites as top-level nodes
    // let confighelper=new DMSEntityConfigurationHelper(props.context);
    confighelper.GetActivEnitities().then(items => {
        const sites = items.map(item => ({ key: ''+item.SiteID, text: item.Title, children: [], checked: false, data:item }));
        setDropdown1(sites);
    });
  }, []);

  

  // Mock function to get values dynamically
  // const getNextDropdownValues = (value) => {
  //   return [`${value} - Option 1`, `${value} - Option 2`, `${value} - Option 3`];
  // };

  const fetchDocumentLibraries = async (sitetitle:string,siteId: string):Promise<any[]> => {
    console.log("siteId",siteId);   
    let  libs=await confighelper.GetActiveEntityDocLibsBySiteId(sitetitle,siteId);
    return libs.map(lib => ({ key: sitetitle+"|"+lib.DocumentLibraryName, text: lib.DocumentLibraryName, children: [], checked: false, data:lib }));
  };

  // Handle selection and adding to the next dropdown
  const handleAddDropdown2 = async (selectedent:{value: string, label: string}[]) => {
    let dropdown2t:any[]=[];

    // selectedent.forEach(async element => {
    //   let doclibs= await fetchDocumentLibraries(element.value);
    //   dropdown2t=dropdown2t.concat(doclibs.filter(d=>dropdown2.filter(drp=>drp.data==d.data && drp.text==d.text).length==0 ));
    //   setDropdown2(dropdown2t);
    // });
    const doclibs = await Promise.all(selectedent.map(element => fetchDocumentLibraries(element.label, element.value)));
    let c=doclibs.flat();
    //let drpdwn=[...dropdown2];
    // dropdown2t = dropdown2t.concat(doclibs.flat().filter(d => dropdown2.filter(drp => drp.data == d.data && drp.text == d.text).length == 0));
    //dropdown2t = c.filter(d => drpdwn.filter(drp => drp.key == d.key && drp.text == d.text).length == 0)
    setDropdown2(c);
    let allsites= [...dropdown1];
    // let sleectedistes= allsites.filter((a)=>selectedent.some(s=>s.value==a.text));
    let sleectedistes= allsites.filter((a)=>selectedent.some(s=>s.label==a.text));
    if(props.onSiteSelect) props.onSiteSelect(sleectedistes.map(a=>a.data));
  };

  const handleAddDropdown3 = async (selectedent:{value: string, label: string}[]) => {
  
      const fields = await Promise.all(selectedent.map(element => {
        let selecteddoclibs=element.value.split('|');
        return fetchFields(selecteddoclibs[0],selecteddoclibs[1]);        
      }));

      let nonaddedfields=fields.flat();//.filter(fld=>dropdown3.filter(d=>d.key==fld.key).length==0)
      
      //let dropdown3t=dropdown3.concat(nonaddedfields);
      setDropdown3(nonaddedfields);
      nonaddedfields.forEach(f=>{
        if(!fieldnamesmapping[f.key])fieldnamesmapping[f.key]=f.text;
      })
      let alllibs= [...dropdown2];
      let sleectedlibs= alllibs.filter((a)=>selectedent.some(s=>s.value==a.key));
      if(props.onLibrarySelect) props.onLibrarySelect(sleectedlibs.map(a=>a.data));
  };

  const handleFieldSelect = async (selectedent:{value: string, label: string}[]) => {
  
    let t=selectedent.map(fielddrpdonw=>({field: fielddrpdonw.value.split('|')[0],fieldtype:fielddrpdonw.value.split('|')[1] as enumfieldtype}));
     if(props.onMultiFieldSelect) props.onMultiFieldSelect(t);
  };

  const handleAddDropdown4 = () => {
    if (selected3) {
      //setDropdown4(getNextDropdownValues(selected3));
    }
  };

  
  // const fetchFields = async (siteid:string,libraryId: string) => {
  //   let web1=await sp.site.openWebById(siteid);
  //   const fields = await web1.web.lists.getById(libraryId).fields.filter("Hidden eq false").select('InternalName', 'Title')();
  //   setFields(fields.map(field => ({ key: field.InternalName, text: field.Title })));
  // };
  const fetchFields = async (siteid:string,libraryId: string) => {
   
    const _fields = await confighelper.GetActiveEntityDocLibsSearchFields(siteid,libraryId);
    let fldnodes=_fields.map(field => ({ key: field.SearchMappedManagedProperty, text: field.ColumnName,type:field.ColumnType, data:siteid+" - "+libraryId }))
    //fldnodes.push({key: 'lastModifiedDateTime', text: 'Published Date',type:enumfieldtype.DateTime});
    //setFields(fields.concat(fldnodes));
    return fldnodes
  };

  // const handleFieldSelect = (event: React.ChangeEvent<HTMLSelectElement>) => {
  //   //setSelectedField(event.target.value);
  // };

  const handleButtonClick = () => {
    let fielddrpdonw=document.getElementById('FieldSelectDropdown') as HTMLSelectElement;
    if (fielddrpdonw.value) {
      let t=fielddrpdonw.value.split('|')
      if(props.onFieldSelect) props.onFieldSelect(t[0],t[1] as enumfieldtype);
    }
  };

 
  
  

  return (
    <div className="p-0 pt-1">
       <div className="mb-2 border-bottom">
        <label className="form-label">Select Entities</label>
        <div className="d-flex">
           {/* <CheckedFilterDropDown options={dropdown1.map(d=>({label:d.text,value:d.text }))} onChange={handleAddDropdown2} ></CheckedFilterDropDown>  */}
           <CheckedFilterDropDown options={dropdown1.map(d=>({label:d.text,value:d.key }))} onChange={handleAddDropdown2} ></CheckedFilterDropDown> 
        </div>
      </div>

      {/* Second Dropdown */}
      <div className="mb-2 border-bottom">
        <label className="form-label">Select Document Libraries</label>
        <div className="d-flex">
          <CheckedFilterDropDown options={dropdown2.map(d=>({label:d.data.SiteTitle+" - "+d.text,value:d.key }))} onChange={handleAddDropdown3}></CheckedFilterDropDown>         </div>
      </div>

      {/* Third Dropdown */}
      <div className="mb-2">
        <label className="form-label">Select Fields</label>
        <div className="d-flex">
          <CheckedFilterDropDown options={dropdown3.map(d=>({label:d.data+" - "+d.text,value:d.key }))} onChange={handleFieldSelect}></CheckedFilterDropDown> 
        </div>
      </div>
    </div>
  );
};

