import { IField } from "./AdvancedSearch";


export enum enumfieldtype
{
  SingleLineofText='Single Line of Text',
  MultipleLineofText='Multiple Line of Text',
  YesNo='Yes or No',
  DateTime='Date & Time',
  Number='Number'
}

export var fieldnamesmapping:any ={
    // 'lastModifiedDateTime':"Published Date",
    'lastmodifiedtime':"Published Date",
    'filetype':'Document Type',
    'createdby':'Uploaded By',
    'fileextension':"File Extension",
}

export const GetFieldName=(fieldname:string)=>{
    console.log('fieldname',fieldname);
    console.log('fieldnamesmapping',fieldnamesmapping);
    return (fieldnamesmapping[fieldname.toLowerCase()])?fieldnamesmapping[fieldname.toLowerCase()]:fieldname;
}

export const removeDuplicates=<T>(array: T[], key: keyof T): T[]=> {
    const seen = new Set();
    return array.filter(item => {
        const value = item[key];
        if (seen.has(value)) {
            return false;
        }
        seen.add(value);
        return true;
    });
}

export const GetFieldValue=(searchfields:IField[], fieldname:string,fieldval:any)=>{
       
    try{
        
        let fldtype=GetFieldType(searchfields,fieldname);
        if(fldtype==enumfieldtype.DateTime)
        {
            return (new Date(fieldval)).toLocaleString();
        }
        else
        {
            return fieldval;
        }
    }
    catch(ex)
    {
        console.log("error",ex);
    }

}

const GetFieldType = (searchfields:IField[], fieldname: string) => {
    try{
        console.log('fieldname',fieldname);
        console.log('searchfields',searchfields);
        return searchfields.filter(f => f.fieldname.toLowerCase() == fieldname.toLowerCase())[0].fieldtype;
    }
    catch(ex)
    {
        
        console.log("error",ex);
    }
}

