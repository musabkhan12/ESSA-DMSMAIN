import { IDMSResultsGridProps } from "./Interfaces";
import * as React from "react";
import '../../components/dmscss.css'

const getFileIcon = (fileName: any) => {
  const fileExtension = fileName?.split(".").pop().toLowerCase();
  let fileIcon;
  switch (fileExtension) {
    case "doc":
    case "docx":
      fileIcon = require("../../assets/DOC.png");
      break;
    case "txt":
      fileIcon = require("../../assets/TXT.png");
      break;
    case "pdf":
      fileIcon = require("../../assets/PDF.png");
      break;
    case "xls":
    case "xlsx":
      fileIcon = require("../../assets/XLS.png");
      break;
    case "zip":
      fileIcon = require("../../assets/ZIP.png");
      break;
    default:
      fileIcon = require("../../assets/DOC.png"); // Default icon if no match
      break;
  }
  return { fileIcon, fileExtension };
};

export const DMSResultsGrid=(props:IDMSResultsGridProps)=>
{
    
  //  const [ficon,setfileicon]=React.useState(require("../../../assets/DOC.png"));
  //  React.useEffect(()=>{
  //     let t=getFileIcon(props.results)
  //  })

    return ( 
        <div id="results-container">
          {
            props.results?.map((resdoc)=>
                <div className="card"> 
                  <div className="row"> 
                    <div className="col-md-2 pe-0"> 
                      <div className="IMGContainer">        
                          <img className="filextension" src={getFileIcon(resdoc.Title).fileIcon} alt="doc icon" data-themekey="#" />
                      </div>
                    </div>
                    <div className="col-md-10"> 
                      <div className="CardTextContainer">
                      <p className="p1st">{resdoc.Title}</p>
                      <p className="p2nd">{resdoc.Extension}</p>
                      <p className="p3rd ">{(resdoc.Size/(1024 * 1024)).toFixed(2)} MB</p>
                      <p className="p3rd ">{resdoc.Summary}</p>
                      </div>     
                    </div> 
                  </div>       
             </div>
            )
          }
        </div>
    )
}

export const ResultsFilter=(props:any)=>
{

}

