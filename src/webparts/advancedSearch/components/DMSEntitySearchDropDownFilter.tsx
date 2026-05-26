import * as React from 'react';
import { useState, useEffect } from 'react';
// import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/sites";
import "bootstrap/dist/css/bootstrap.min.css";
import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';
// import { getSP } from '../../essadmsMain/loc/pnpjsConfig';
import { getSP } from '../../../webparts/advancedSearch/loc/pnpjsConfig';
import { DMSEntityConfigurationHelper } from '../../../Shared/DMSEntityConfigurationHelper';
import { IContextInfo } from '@pnp/sp/context-info';
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

}

export const DMSEntitySearchDropDowns: React.FC<IDMSEntitySearchDropDownsProps> = (props:IDMSEntitySearchDropDownsProps) => {
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
        const sites = items.map(item => ({ key: ''+item.SiteID, text: item.Title, children: [], checked: false }));
        setDropdown1(sites);
    });
  }, []);

  

  // Mock function to get values dynamically
  // const getNextDropdownValues = (value) => {
  //   return [`${value} - Option 1`, `${value} - Option 2`, `${value} - Option 3`];
  // };

  const fetchDocumentLibraries = async (siteId: string):Promise<any[]> => {
    console.log("siteId",siteId);   
    let  libs=await confighelper.GetActiveEntityDocLibsBySiteId(siteId);
    return libs.map(lib => ({ key: siteId+"|"+lib.DocumentLibraryName, text: lib.DocumentLibraryName, children: [], checked: false, data:siteId }));
  };

  // Handle selection and adding to the next dropdown
  const handleAddDropdown2 = async () => {
    if (selected1) {
      let doclibs= await fetchDocumentLibraries(selected1);
      let dropdown2t=dropdown2.concat(doclibs.filter(d=>dropdown2.filter(drp=>drp.data==d.data && drp.text==d.text).length==0 ));
      setDropdown2(dropdown2t);
    }
  };

  const handleAddDropdown3 = async () => {
    if (selected2) {
      let selecteddoclibs=selected2.split('|');
      let fields=await fetchFields(selecteddoclibs[0],selecteddoclibs[1])
      let nonaddedfields=fields.filter(fld=>dropdown3.filter(d=>d.key==fld.key).length==0)
      let dropdown3t=dropdown3.concat(nonaddedfields);
      setDropdown3(dropdown3t);
    }
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

  const handleFieldSelect = (event: React.ChangeEvent<HTMLSelectElement>) => {
    //setSelectedField(event.target.value);
  };

  const handleButtonClick = () => {
    let fielddrpdonw=document.getElementById('FieldSelectDropdown') as HTMLSelectElement;
    if (fielddrpdonw.value) {
      let t=fielddrpdonw.value.split('|')
      if(props.onFieldSelect) props.onFieldSelect(t[0],t[1] as enumfieldtype);
    }
  };

//   const handleLibrarySelect = async (event: React.ChangeEvent<HTMLSelectElement>) => {
//     const libraryId = event.target.value;
//     setSelectedLibrary(libraryId);

//     if (libraryId) {
//       await fetchFields(libraryId);
//     } else {
//       setFields([]);
//     }
//   };

  
//   const fetchFolders = async (libraryId: string) => {
//     const folders = await sp.web.lists.getById(libraryId).items.filter("FSObjType eq 1").select('ID', 'Title').get();
//     return folders.map(folder => ({ key: folder.ID, text: folder.Title, checked: false }));
//   };

  // const handleSiteCheck = async (siteIndex: number) => {
  //   const updatedTree = [...treeData];
  //   const site = updatedTree[siteIndex];
  //   site.checked = !site.checked;

  //   if (site.checked && site.children && site.children.length === 0) {
  //     site.children = await fetchDocumentLibraries(site.text);
  //   }

  //   setTreeData(updatedTree);
  // };

  

  // const handleLibraryCheck = async (siteIndex: number, libraryIndex: number,siteid:string,libraryid:string) => {
  //    const updatedTree = [...treeData];
  //    console.log(siteid,libraryid);
  //   const library = updatedTree[siteIndex].children![libraryIndex];
  //   library.checked = !library.checked;

  //   if (library.checked && library.children && library.children.length === 0) {
  //     //library.children = await fetchFolders(library.key);
  //   }

  //    setTreeData(updatedTree);
  //    if (libraryid) {
  //       await fetchFields(siteid,libraryid);
  //     } else {
  //       setFields([]);
  //     }
  // };

  return (
    <div className="container p-3">
       <div className="mb-3">
        <label className="form-label">Select Entities</label>
        <div className="d-flex">
          <select
            className="form-select me-2"
            value={selected1}
            onChange={(e) => setSelected1(e.target.value)}
          >
            <option value="">Select an option</option>
            {dropdown1.map((item, index) => (
              <option key={index} value={item.text}>
                {item.text}
              </option>
            ))}
          </select>
          <button type='button' className="btn btn-primary" onClick={(e)=>handleAddDropdown2()}>
            Add Entity
          </button>
        </div>
      </div>

      {/* Second Dropdown */}
      <div className="mb-3">
        <label className="form-label">Select Document Libraries</label>
        <div className="d-flex">
          <select
            className="form-select me-2"
            value={selected2}
            onChange={(e) => setSelected2(e.target.value)}
            disabled={dropdown2.length === 0}
          >
            <option value="">Select an option</option>
            {dropdown2.map((item, index) => (
              <option key={index} value={item.key}>
                {item.data}-{item.text}
              </option>
            ))}
          </select>
          <button type='button'
            className="btn btn-primary"
            onClick={handleAddDropdown3}
            disabled={!selected2}
          >
            Add Document Library
          </button>
        </div>
      </div>

      {/* Third Dropdown */}
      <div className="mb-3">
        <label className="form-label">Select Fields</label>
        <div className="d-flex">
          <select
            className="form-select me-2"
            id='FieldSelectDropdown' 
            value={selected3}
            onChange={(e) => setSelected3(e.target.value)}
            disabled={dropdown3.length === 0}
          >
            <option value="">Select an option</option>
            {dropdown3.map((item, index) => (
              <option key={index} value={item.key+"|"+item.type}>
                 {item.data}-{item.text}
              </option>
            ))}
          </select>
          <button type='button'
            className="btn btn-primary"
            onClick={handleButtonClick}
            disabled={!selected3}
          >
            Add Field
          </button>
        </div>
      </div>
      

      {/* {treeData.map((site, siteIndex) => (
        // <div key={site.key} className="form-check">
        <div key={site.text} className="form-check">
          <input
            className="form-check-input"
            type="checkbox"
            checked={site.checked}
            onChange={() => handleSiteCheck(siteIndex)}
          />
       
          <label className="form-check-label">{site.text}</label>
          <div className='newline'>
          {site.children && site.children.map((library, libraryIndex) => (
            <div key={library.key} className="ms-3 form-check">
              <input
                className="form-check-input"
                type="checkbox"
                checked={library.checked}
                onChange={() => handleLibraryCheck(siteIndex, libraryIndex,site.text,library.text)}
              />
              <label className="form-check-label">{library.text}</label>
              {library.children && library.children.map(folder => (
                <div key={folder.key} className="ms-4 form-check">
                  <input className="form-check-input" type="checkbox" checked={folder.checked} readOnly />
                  <label className="form-check-label">{folder.text}</label>
                </div>
              ))}
            </div>
          ))}
        </div></div>
      ))}    */}
      {/* <div className='row'>  
      <div className="mt-3 mb-3 col-sm-4">
        <label className="form-label">Fields</label>
        <select id='FieldSelectDropdown' className="form-select" onChange={handleFieldSelect} value={selectedField || ''}>
          {fields.map(field => (
            <option key={field.key} value={field.key+'|'+field.type} data-type={field.type} >{field.text}</option>
          ))}
        </select>
      </div>
      <div className="mt-3 mb-3 col-sm-4">
      <button type='button' style={{marginTop:'27px',padding:'8px 12px'}} className="btn btn-primary mb-0" onClick={handleButtonClick}>Submit Field</button></div>
    </div> */}
    </div>
  );
};

