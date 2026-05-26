declare global {
  interface Window {
    // managePermissionSearch: (DocumentLibraryName: string, SiteTilte: string, SiteID: string, folderName: any, folderPath: any) => void;
    manageWorkflow: (DocumentLibraryName: string, SiteTilte: string, SiteID: string) => void;
    view: (message: string) => void;
    PreviewFile: (path: string, siteID: string, docLibName: any, filemasterlist: any, filepreview: any) => void;
    Download: (path: string, siteID: string, docLibName: any, filemasterlist: any, filepreview: any) => void;
    deleteFile: (fileId: string, siteID: string, IsHardDelete: any, listToUpdate: any) => void;
    shareFile: (fileID: string, siteId: string, currentFolderPathForFile: string, fileName: string, flag: string, FileVersion: any, FileSize: any, Status: any, FilePreviewURL: any, DocumentLibraryName: any) => void;
    editFile: (siteName: string, documentLibraryName: string) => void;
    documentLibraryPopUp: (fileId: string, siteID: any, FolderPath: any, FileName: any, permission: any) => void;
    undo: (fileId: any, siteId: any, FileMasterList: any, documentLibraryName: any, ID: any, folderPath: any, fileName: any) => void;
    confirmUndo: (fileId: any, siteId: any, FileMasterList: any, documentLibraryName: any, ID: any, folderPath: any, fileName: any) => void;
    hideSharePopUp: () => void;
  }

}

// props for Manage work flow
// props for Manage work flow
const propsForManageWorkFlow = {
  SiteTitle: "",
  DocumentLibraryName: "",
  SiteID: ""
}
// props for managePermissionSearch
const managePermissionProps = {
  SiteTitle: "",
  DocumentLibraryName: "",
  SiteID: "",
  FolderName: "",
  FolderPath: ""
}
interface UploadFileProps {
  currentfolderpath: {
    CurrentEntity: string;
    currentEntityURL: string;
    currentsiteID: string;
    // ... other properties
  };
}
// export interface IEssadmsMainProps {
//   description: string;
//   isDarkTheme: boolean;
//   environmentMessage: string;
//   hasTeamsContext: boolean;
//   userDisplayName: string;
//   context: any;
//   siteUrl: string;
// }

// @ts-ignore
import * as React from "react";
// import { getSP } from "../../loc/pnpjsConfig";
import { getSP } from '../../../advancedSearch/loc/pnpjsConfig';
import { SPFI } from "@pnp/sp";
import "bootstrap/dist/css/bootstrap.min.css";
// import "bootstrap//dist/"


import "../../../verticalSideBar/components/VerticalSidebar2.scss";
import VerticalSideBar from "../../../verticalSideBar/components/VerticalSideBar";
import UserContext from "../../../../GlobalContext/context";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import {
  faUser,
  faShareAlt,
  faListSquares,
  faTableCells,
  faList
  // faTrash, 
  // faEdit, 
  // faEye  
} from "@fortawesome/free-solid-svg-icons";
import {
  faStar as faStarRegular,
  faFolder as faFolderRegular,

} from "@fortawesome/free-regular-svg-icons";
// import { useState , useEffect } from "react";
import Provider from '../../../../GlobalContext/provider';
import { useMediaQuery } from "react-responsive";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all"
import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/site-users/web";
import { PermissionKind } from "@pnp/sp/security";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../../CustomCss/mainCustom.scss";
import "../../../verticalSideBar/components/VerticalSidebar2.scss";
import "./DMSSearch.scss";
// import "./DmsMain.css";
import "../../../../CustomCss/mainCustom.scss"
import { useState, useRef, useEffect } from "react";
import UploadFile from "../UploadFile";
import CreateFolder from "../CreateFolder";
// import Table from "../Table";
import { IFileInfo } from "@pnp/sp/files";
import { Popup } from "@fluentui/react";


import { IEssadmsMainProps } from '../IEssadmsMainProps'
import HorizontalNavbar from "../../../horizontalNavBar/components/HorizontalNavBar";
import ManageWorkFlow from "../ManageWorkflow";
import ManageFolderPermission from "../ManageFolderPermission";
import { folderFromPath } from "@pnp/sp/folders";
import Swal from "sweetalert2";
import { SearchQueryInit } from "@pnp/sp/presets/all";
import { BaseWebPartContext, WebPartContext } from "@microsoft/sp-webpart-base";
import { DMSResultsGrid } from "./DMSResultsGrid";
import { IDocumentDisplayFields, ResultsViewMode } from "./Interfaces";
import { GraphSearchHelper } from "../../../../Shared/SearchHelper1";
import DMSSearchFilter from "./DMSSearchFilter";
import { ISearchHitResource } from "../../../../Shared/SearchHelperInterfaces";



// import { documentLibraryPopUp, getFileIcon } from "./SearchResultsViewHelper";

let Undo = require('../../assets/Undo.svg');
let sharewithmeicon = require('../../assets/nodes.png')
let recyclebin = require('../../assets/recycle-bin.png')
let sharewithothericon = require('../../assets/share.png')
let starticon = require('../../assets/star.png')
let listicon = require('../../assets/list.png')
let downloadicon = require('../../assets/download.png')
let foldericon = require('../../assets/foldericon.png')
let Docicon = require("../../assets/DOC.png");
let Txticon = require("../../assets/TXT.png");
let Pdficon = require("../../assets/PDF.png");
let Xlsicon = require("../../assets/XLS.png");
let Zipicon = require("../../assets/ZIP.png");
let MainRounteVariable = 'MyRequest'

let managePermissionIcon = require('../../assets/managePermissionSearch.svg')
// import managePermissionIcon from '../../assets/managePermissionSearch.svg';
let manageWorkFlowIcon = require('../../assets/ManageWorkflow.svg')
// import manageWorkFlowIcon from '../../assets/ManageWorkflow.svg';
let viewIcon = require('../../assets/View.svg')
// import viewIcon from '../../assets/View.svg';
let editIcon = require('../../assets/Edit.svg')
// import editIcon from '../../assets/Edit.svg';
let deleteIcon = require('../../assets/Delete.svg')
// import deleteIcon from '../../assets/Delete.svg';
let FillFavouriteFile = require('../../assets/FillFavourite.svg')
let ShareFile = require('../../assets/Edit.svg')
let UnFillFavouriteFile = require('../../assets/UnFillFavourite.svg')
let myfolderdata: any = []

let currentDocumentLibrary = "";
let currentFolder = ""
let currentfolderpath = "";
let ismyrequordoclibforfilepreview: any
// @ts-ignore
let parentfolder = ""
let currentDevision = "";
// @ts-ignore
let currentDepartment = ""
let currentEntityURL = "";
// @ts-ignore
let currentEntity = ""
let currentsiteID = ""
let iscontribute = ""
let isadmin = ""
let mydatacard = ""
let mydata: string[] = [];

// start
// let searchArray:any=[];
let routeToDiffSideBar = "";
// end



const ArgPoc = ({ props }: any) => {
  const sp: SPFI = getSP();
  // console.log(sp, "sp");
  let locationPath=window.location.pathname.match(/\/sites\/[^\/]+/)[0];
  const [showDeletepopup, setShowDeletepopup] = useState(false);
  const [activeButton] = React.useState<string>("");
  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);
  const [showFirstDiv, setShowFirstDiv] = useState(true);
  const [showworkflowdiv, setshowworkflowdiv] = useState('');
  const [showWorkflow, setShowWorkflow] = useState(false);
  const [showfolderpermission, setShowfolderpermission] = useState(false);
  let cleanUrlInMyRequest = false;
  // const handleButtonClickShow = () => {
  //   setShowFirstDiv(false);
  // };


  React.useEffect(() => {
    // console.log("This function is called only once", useHide);

    const showNavbar = (
      toggleId: string,
      navId: string,
      bodyId: string,
      headerId: string
    ) => {
      const toggle = document.getElementById(toggleId);
      const nav = document.getElementById(navId);
      const bodypd = document.getElementById(bodyId);
      const headerpd = document.getElementById(headerId);

      if (toggle && nav && bodypd && headerpd) {
        toggle.addEventListener("click", () => {
          nav.classList.toggle("show");
          toggle.classList.toggle("bx-x");
          bodypd.classList.toggle("body-pd");
          headerpd.classList.toggle("body-pd");
        });
      }
    };

    showNavbar("header-toggle", "nav-bar", "body-pd", "header");

    const linkColor = document.querySelectorAll(".nav_link");

    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach((l) => l.classList.remove("active"));
        this.classList.add("active");
      }
    }

    linkColor.forEach((l) => l.addEventListener("click", colorLink));
  }, [useHide]);
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: "(max-width: 768px)" });

  React.useEffect(() => {
    // console.log("This function is called only once", useHide);

    const showNavbar = (
      toggleId: string,
      navId: string,
      bodyId: string,
      headerId: string
    ) => {
      const toggle = document.getElementById(toggleId);
      const nav = document.getElementById(navId);
      const bodypd = document.getElementById(bodyId);
      const headerpd = document.getElementById(headerId);

      if (toggle && nav && bodypd && headerpd) {
        toggle.addEventListener("click", () => {
          nav.classList.toggle("show");
          toggle.classList.toggle("bx-x");
          bodypd.classList.toggle("body-pd");
          headerpd.classList.toggle("body-pd");
        });
      }
    };

    showNavbar("header-toggle", "nav-bar", "body-pd", "header");

    const linkColor = document.querySelectorAll(".nav_link");

    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach((l) => l.classList.remove("active"));
        this.classList.add("active");
      }
    }

    linkColor.forEach((l) => l.addEventListener("click", colorLink));
  }, [useHide]);
  React.useEffect(() => {
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        if (document.fullscreenElement) {
          document.exitFullscreen();
        }
      }
    };

    window.addEventListener("keydown", handleEscape);
    return () => window.removeEventListener("keydown", handleEscape);
  }, []);


  // code to route to different document library and folder start
  useEffect(() => {
    // const params = new URLSearchParams(window.location.search);
    const url = window.location.href;
    // const matches = url.match(/\/([^\/]+)\.aspx/);
    let extractedPart = url.split('.aspx')[1];
    let parameters = extractedPart.split('?')
    // alert( parameters);
    console.log("parameters", parameters);
    let path = "";
    let siteId = "";
    let folderName = "";
    let devision = "";
    let department = "";
    if (parameters.length > 1) {
      parameters.forEach((items, index) => {
        console.log(`items[${index}]`, items)

        if (index == 1) {
          if (items.includes('%20')) {
            console.log("Clean Url")
            const cleanUrl = items.replace(/%20/g, ' ');
            path = cleanUrl;
          } else {
            path = items;
          }

        }
        if (index == 2) {
          if (items.includes('%20')) {
            console.log("Clean path")
            const cleanUrl = items.replace(/%20/g, ' ');
            folderName = cleanUrl;
          } else {
            folderName = items;
          }
          // folderName=items;
        }
        if (index == 3) {
          siteId = items;
        }
        if (index == 4) {
          if (items.includes('%20')) {
            console.log("Clean devision")
            const cleanDevision = items.replace(/%20/g, ' ');
            devision = cleanDevision;
          } else {
            devision = items;
          }
        }
        if (index == 5) {
          if (items.includes('%20')) {
            console.log("Clean deaprtment")
            const cleanDepartment = items.replace(/%20/g, ' ');
            department = cleanDepartment;
          } else {
            department = items;
          }
        }
      })
      console.log("path", path)
      console.log("siteId", siteId)
      console.log("folderName", folderName)
      console.log("department", department)
      console.log("devision", devision)
      currentDepartment = department;
      currentDevision = devision;
      cleanUrlInMyRequest = true;
      getdoclibdata(path, siteId, folderName);
    }

  }, []);
  // end
  /////////////////// DMS Code start / ////////////////////////////////////
  const buttonDivRef = useRef<HTMLDivElement>(null);
  const [showMyrequButtons, setShowMyrequButtons] = useState(true); // Initially hidden
  const [showMyfavButtons, setShowMyfavButtons] = useState(false); // Initially hidden
  const [displayuploadfileandcreatefolder, setdisplayuploadfileandcreatefolder] = useState(false); // Initially hidden
  const [Myreqormyfav, setMyreqormyfav] = useState(''); // Initially hidden
  const [searchQueryPath, setSearchQueryPath] = useState<string>((props.siteUrl));
  const [searchDocResults, setSearchDocResults] = useState<IDocumentDisplayFields[]>([]);



  // console.log(Myreqormyfav , "Myreqormyfav")
  // console.log("This is current side ID",currentsiteID)
  const currentUserEmailRef = useRef('');
  useEffect(() => {
    getcurrentuseremail();
    getdata();
    runSearch();

  }, []);


  const getdata = async () => {
    // debugger
    const ids = window.location.search;
    //  alert(ids)
    const originalString = ids;
    // alert(originalString)
    const idNum2: any = originalString.substring(1);
    // alert(idNum2)
    const getgroup = await sp.web.lists
      .getByTitle("ARGGroupandTeam")
      .items.select("*,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType").expand("InviteMemebers")()
      .then((res) => {
        // arr=res;
        console.log(res, ":response")
        // debugger
        console.log("res------", res)
        //  setArrDetails(res)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
  }

  const myrequestbuttonclick = () => {
    const musa = document.getElementById('Myrequestbutton')
    if (musa) {
      // alert("enter")
      musa.click();
      // alert("click")
    }

  }

  const getcurrentuseremail = async () => {
    const userProfile = await sp.profiles.myProperties();
    console.log(userProfile, "userProfile")
    console.log(userProfile.Title, "userProfile userProfile.Title")
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    myrequestbuttonclick()
    // console.log(currentUserEmailRef.current, "currentuser")
  }


  const fetchAndBuildTree2 = async () => {
    // event.preventDefault()
    // event.stopImmediatePropagation()
    // event.stopPropagation()
    try {
      //Old working code
      //  Fetch data from EntityDivisionDepartmentMappingMasterList
      const entityItems = await sp.web.lists
        .getByTitle("EntityDivisionDepartmentMappingMasterList")
        .items.select(
          "Entitylookup/Title, Entitylookup/SiteURL", "Entitylookup/SiteID",
          "Devisionlookup/Title",
          "Departmentlookup/Title",
          "Devisionlookup/Active",
          "Departmentlookup/Active"
        )
        .expand("Entitylookup", "Devisionlookup", "Departmentlookup")
        .filter("Entitylookup/Active eq 'Yes'")();
      console.log(entityItems, "entityItems 1")
      const uniqueEntityMap = new Map();
      const uniqueEntitiesWithAccess: any = [];

      // Loop through each item and check permissions
      for (const item of entityItems) {
        const entityTitle = item.Entitylookup.Title;
        try {
          const subsiteWeb = await sp.site.openWebById(item.Entitylookup.SiteID);
          const hasAccess = await subsiteWeb.web.currentUserHasPermissions(PermissionKind.ViewListItems);

          if (hasAccess) {
            // Add to uniqueEntitiesWithAccess only if user has access
            uniqueEntityMap.set(entityTitle, item); // Store the item or any required data
            uniqueEntitiesWithAccess.push(item);  // Add the item to the list of entities with access
            console.log(`User has access to site: ${entityTitle}`, item);
          } else {
            console.log(`User does not have access to site: ${entityTitle}`);
          }
        } catch (error) {
          console.error(`Error while checking access for site: ${entityTitle}`, error);
        }
      }
      // console.log(uniqueEntityMap, "uniqueEntityMap ......")
      // console.log(uniqueEntitiesWithAccess, "uniqueEntitiesWithAccess");
      /// New Code 


      // Fetch data from DMSFolderMaster
      const folderItems = await sp.web.lists
        .getByTitle("DMSFolderMaster")
        .items.getAll();
      console.log("folderItems", folderItems);

      const myButton = document.getElementById("mybutton");
      const createFileButton = document.getElementById("createFileButton");
      const createFileButton2 = document.getElementById("createFileButton2");
      const createFolderButton = document.getElementById("createFolderButton");
      // Create a map to hold folder data by SiteTitle, Devision, Department
      const folderMap = new Map();
      folderItems.forEach((folderItem) => {
        const {
          SiteTitle,
          Devision,
          Department,
          DocumentLibraryName,
          FolderName,
          ParentFolderId,
          FolderPath,
          IsRename
        } = folderItem;

        if (SiteTitle) {
          const key = `${SiteTitle.trim()}::${Devision?.trim() || ""}::${Department?.trim() || ""}`;
          if (!folderMap.has(key)) {
            folderMap.set(key, []);
          }
          if (DocumentLibraryName) {
            folderMap
              .get(key)
              .push({
                IsRename,
                FolderPath,
                ParentFolderId,
                DocumentLibraryName,
                FolderName: Array.isArray(FolderName)
                  ? FolderName
                  : [FolderName],
              });
          }
        }
      });
      // console.log(folderMap, "folderMap");
      // const entitiesMap = new Map();
      const entitiesMap: any = new Map();

      uniqueEntitiesWithAccess.forEach((item: any) => {
        const entityTitle = item.Entitylookup.Title;
        const siteURL = item.Entitylookup.SiteURL;
        const siteID = item.Entitylookup.SiteID;

        if (!entitiesMap.has(entityTitle)) {
          entitiesMap.set(entityTitle, {
            siteURL: siteURL,
            entityTitle: entityTitle,
            siteID: siteID,
            devisions: new Map(),
          });
        }

        const entry = entitiesMap.get(entityTitle);
        const devisionTitle = item.Devisionlookup?.Title;
        const departmentTitle = item.Departmentlookup?.Title;
        const isDevisionActive = item.Devisionlookup?.Active === "Yes";
        const isDepartmentActive = item.Departmentlookup?.Active === "Yes";

        if (devisionTitle && isDevisionActive) {
          if (!entry.devisions.has(devisionTitle)) {
            entry.devisions.set(devisionTitle, {
              departments: new Set(),
              docLibs: new Set(),
            });
          }
          const devisionEntry = entry.devisions.get(devisionTitle);
          if (departmentTitle && isDepartmentActive) {
            devisionEntry.departments.add(departmentTitle);
          } else if (!departmentTitle || !isDepartmentActive) {
            const nullDeptKey = `${entityTitle.trim()}::${devisionTitle.trim()}::`;
            // Handle case where department is null or inactive
          }
        }
      });
      const buildFolderStructure = (
        folderList: HTMLElement,
        folders: any[],
        parentFolderId: string | null
      ) => {
        const filteredFolders = folders.filter(
          (folder) => folder.ParentFolderId === parentFolderId
        );
        filteredFolders.forEach((folder) => {
          const folderElement = document.createElement("li");
          folderElement.textContent = folder.FolderName;
          folderList.appendChild(folderElement);

          const childFolderList = document.createElement("ul");
          childFolderList.style.display = "none";
          folderElement.appendChild(childFolderList);

          folderElement.addEventListener("click", (event) => {
            event.stopPropagation();
            // currentFolder = folder.FolderName;
            toggleVisibility(childFolderList);
          });

          // Recursively build the structure for subfolders
          buildFolderStructure(childFolderList, folders, folder.FolderName);
        });
      };
      // Build the folder tree structure in the DOM
      const container = document.getElementById("folderContainer2");

      if (container) {
        container.innerHTML = ""; // Clear previous contents
      } else {
        console.error("Container element not found");
      }
      // container.innerHTML = ''; // Clear previous contents

      const toggleVisibility = (element: any, forceShow = false) => {
        const isVisible = element.style.display === "block";
        element.style.display = isVisible && !forceShow ? "none" : "block";
      };
      const createImageElement = (src: string, alt: string) => {
        const img = document.createElement("img");
        img.src = require("../../assets/add-folder.png");
        img.alt = alt;
        img.style.float = "left";
        img.style.width = "20px"; // Adjust the size as needed
        img.style.height = "20px"; // Adjust the size as needed
        img.style.marginRight = "5px"; // Space between image and text
        return img;
      };
      //     const createToggleButton = () => {
      //     const link = document.createElement("a");
      //     link.textContent = "+"; // Initial text
      //     link.className="toggle-button"
      //     link.style.cursor = "pointer";
      //     link.style.textDecoration = "none";

      //     link.addEventListener("click", (e) => {
      //         e.preventDefault()
      //         console.log("Button clicked +/-");
      //         if (link.textContent === "+") {
      //             link.textContent = "-"; // Change to minus when content is visible
      //         } else if(link.textContent){
      //             link.textContent = "+"; // Change to plus when content is hidden
      //         }
      //     });

      //     return link;
      // };
      const createToggleButton = (siteurl: string) => {
        const link = document.createElement("a");
        // link.id="toggle-plus-minus";
        link.textContent = "+"; // Initial text
        link.className = "toggle-button"
        link.style.cursor = "pointer";
        link.style.textDecoration = "none";
        link.setAttribute("siteurl", siteurl);
        link.addEventListener("click", (e) => {
          e.preventDefault();
          let siteurl = (e.target as HTMLElement).getAttribute("siteurl");
          setSearchQueryPath(`${siteurl}`);
          runSearch(`${siteurl}`);
        });

        // link.addEventListener("click", (e) => {
        //     e.preventDefault()
        //     console.log("Button clicked +/-");
        //     if (link.textContent === "+") {
        //         link.textContent = "-"; // Change to minus when content is visible
        //     } else if(link.textContent){
        //         link.textContent = "+"; // Change to plus when content is hidden
        //     }
        // });

        return link;
      };
      entitiesMap.forEach((value: any, entityTitle: any) => {
        const titleElement = document.createElement("p");

        // titleElement.textContent = entityTitle;
        titleElement.classList.add("folder", "icon");
        titleElement.style.cursor = "pointer";
        // const entityImage = createImageElement(
        //   "icons/entity-icon.png",
        //   "Entity Icon"
        // );
        const toggleButton = createToggleButton(value.siteURL);
        titleElement.appendChild(toggleButton);
        titleElement.appendChild(document.createTextNode(entityTitle));

        if (container) {
          container.appendChild(titleElement);
        } else {
          console.error("Container element not found");
        }

        const documentList = document.createElement("ul");
        titleElement.appendChild(documentList);
        documentList.style.display = "none";
        /////start: Display Document library with recursive folder under Enitiy directly when Devision and Department Null /////
        const nullKey = `${entityTitle.trim()}::::`;
        if (folderMap.has(nullKey)) {
          const documentLibraries = folderMap.get(nullKey) || [];

          // Create a map to store unique DocumentLibraryNames and their details
          const uniqueDocLibs = new Map();

          // Iterate over document libraries and populate the map with unique DocumentLibraryNames
          documentLibraries.forEach((item: any) => {
            if (!uniqueDocLibs.has(item.DocumentLibraryName)) {
              uniqueDocLibs.set(item.DocumentLibraryName, {
                folders: [],
                folderPath: item.FolderPath, // Store FolderPath with other details
              });
            }
            uniqueDocLibs.get(item.DocumentLibraryName).folders.push(item);
          });

          // Now render each unique DocumentLibraryName and its associated folders
          uniqueDocLibs.forEach((data, docLibName) => {
            const docLibElement = document.createElement("li");
            // New code to check the Document library name is Rename or not start
            const checkIsRename = data.folders.filter((item: any) =>
              item.FolderName.length === 1 && item.FolderName[0] === null
            );
            // console.log("checkIsRename", checkIsRename);
            let renameText = docLibName;
            if (checkIsRename[0].IsRename !== null) {
              renameText = checkIsRename[0].IsRename;
            }
            // End
            docLibElement.textContent = renameText;
            // docLibElement.textContent = docLibName;
            // console.log("Data of doclib/folder", data);
            // Optionally display the FolderPath in the docLibElement
            const pathText = document.createElement("span");
            // pathText.textContent = ` (${data.folderPath})`; // Display FolderPath
            docLibElement.appendChild(pathText);

            documentList.appendChild(docLibElement);

            const folderList = document.createElement("ul");
            folderList.style.display = "none";
            folderList.style.width = "240px";
            const entityImage = createImageElement(
              "icons/entity-icon.png",
              "Entity Icon"
            );
            docLibElement.appendChild(entityImage);
            docLibElement.appendChild(folderList);

            // Handle click to toggle the visibility of the folder list
            docLibElement.addEventListener("click", (event: any) => {
              event.preventDefault()
              event.stopPropagation();
              // setlistorgriddata('')
              // setShowMyrequButtons(false)
              // setShowMyfavButtons(false)
              handleNavigation(value.entityTitle, null, null, docLibName, null)
              toggleVisibility(folderList);
              getdoclibdata(data.folderPath, value.siteID, docLibName);
              currentfolderpath = data.folderPath
              currentDocumentLibrary = docLibName;
              currentEntityURL = value.siteURL;
              currentEntity = value.entityTitle
              currentsiteID = value.siteID
              currentDevision = ''
              currentDepartment = ''
              currentFolder = ''

              // console.log(currentEntityURL, "currentEntityURL")
              // console.log(currentsiteID, "currentsiteID")
              // console.log(currentEntity, "currentEntity")
              // console.log(currentDocumentLibrary, "currentFolder")
              // console.log(currentfolderpath, "currentfolderpath")
              // console.log(currentDevision, "currentfolderpath")
              // console.log(currentDepartment, "currentfolderpath")
              // createFileButton.style.display = "block";
              // createFileButton2.style.display = "block";
              if (createFolderButton) {
                createFolderButton.style.display = "block"
              }

              if (createFileButton) {
                createFileButton.style.display = "block";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "block";
              }

              if (myButton) {
                myButton.textContent = `Create Folder under ${docLibName}`;
              } else {
                console.error();
              }
            });

            // Handle double-click to hide the folder list
            docLibElement.addEventListener("dblclick", (event) => {
              event.stopPropagation();
              toggleVisibility(folderList, false);
            });

            // Function to build the folder structure recursively
            const buildFolderStructure = (
              parentFolderId: any,
              parentElement: any
            ) => {
              data.folders.forEach((item: any) => {
                const folderNamesArray = Array.isArray(item.FolderName)
                  ? item.FolderName
                  : [item.FolderName];

                folderNamesArray.forEach((folderName: any) => {
                  if (folderName && item.ParentFolderId === parentFolderId) {
                    // Only display non-null folder names
                    const folderElement = document.createElement("li");
                    // New code to check the folder library name is Rename or not start
                    let folderRenameText = folderName;
                    if (item.IsRename !== null) {
                      folderRenameText = item.IsRename
                    }
                    // End
                    // folderElement.textContent = folderName;
                    folderElement.textContent = folderRenameText;
                    parentElement.appendChild(folderElement);
                    const entityImage = createImageElement(
                      "icons/entity-icon.png",
                      "Entity Icon"
                    );
                    folderElement.appendChild(entityImage);
                    const subFolderList = document.createElement("ul");
                    subFolderList.style.display = "none";
                    subFolderList.style.width = "240px";
                    folderElement.appendChild(subFolderList);

                    folderElement.addEventListener("click", (event: any) => {
                      event.preventDefault();  // Prevent default action
                      event.stopPropagation();  // Stop event bubbling
                      console.log("Event listener triggered");
                      currentEntityURL = value.siteURL;
                      currentsiteID = value.siteID
                      currentEntity = value.entityTitle
                      currentDocumentLibrary = docLibName;
                      currentFolder = folderName;
                      parentfolder = item.ParentFolderId;
                      currentfolderpath = item.FolderPath;
                      currentDevision = ''
                      currentDepartment = ''
                      console.log(currentEntityURL, "currentEntityURL")
                      console.log(currentsiteID, "currentsiteID")
                      console.log(currentEntity, "currentEntity")
                      console.log(currentDocumentLibrary, "currentDocumentLibrary")
                      console.log(currentFolder, "currentFolder")
                      console.log(parentfolder, "parentfolder")
                      console.log(currentfolderpath, "currentfolderpath");
                      handleNavigation(value.entityTitle, null, null, docLibName, folderName)
                      event.stopPropagation();
                      getdoclibdata(item.FolderPath, currentsiteID, docLibName)
                      if (myButton) {
                        myButton.textContent = `Create Folder under ${folderName}`;
                      } else {
                        console.error();
                      }


                      toggleVisibility(subFolderList);

                      // Clear existing sub-folder list to avoid duplications
                      subFolderList.innerHTML = "";

                      // Recursively build the sub-folder structure
                      buildFolderStructure(folderName, subFolderList);
                    });
                  }
                });
              });
            };

            // Start building the folder structure from the root level (null ParentFolderId)
            // buildFolderStructure(null, folderList);
          });
        }
        /////End Display Document library with recursive folder under Enitiy directly when Devision and Department Null /////
        const devisionList = document.createElement("ul");
        devisionList.style.display = "none";
        titleElement.appendChild(devisionList);

        value.devisions.forEach((devisionValue: any, devisionTitle: any) => {
          const devisionElement = document.createElement("li");
          devisionElement.textContent = devisionTitle;
          devisionElement.classList.add("folder", "icon");
          devisionElement.style.cursor = "pointer";
          devisionList.appendChild(devisionElement);

          const docLibList = document.createElement("ul");
          docLibList.style.display = "none";
          const entityImage = createImageElement(
            "icons/entity-icon.png",
            "Entity Icon"
          );
          devisionElement.appendChild(entityImage);
          devisionElement.appendChild(docLibList);

          // Display unique DocumentLibraryName under Devision
          console.log("devisionValue.docLibs", devisionValue.docLibs);
          devisionValue.docLibs.forEach((docLibName: any) => {
            const docLibElement = document.createElement("li");
            docLibElement.textContent = docLibName;
            docLibElement.classList.add("file-icon", "icon");
            docLibList.appendChild(docLibElement);

            const folderList = document.createElement("ul");
            folderList.style.display = "none";

            docLibElement.appendChild(folderList);

            const docLibKey = `${entityTitle.trim()}::${devisionTitle.trim()}::`;
            const docLibFolders = folderMap.get(docLibKey) || [];
            docLibFolders.forEach((folderItem: any) => {
              console.log("Folder under divisions", folderItem);
              const folderElement = document.createElement("li");
              folderElement.textContent = folderItem.FolderName;

              folderList.appendChild(folderElement);
            });

            docLibElement.addEventListener("click", (event) => {
              console.log(devisionValue, "devisionValue");
              event.stopPropagation();
              currentDocumentLibrary = docLibName;
              // currentFolder = '';
              currentDevision = devisionTitle;
              // currentDepartment = '';
              currentEntityURL = value.siteURL;
              currentEntity = value.entityTitle
              currentsiteID = value.siteID

              console.log("currentEntityURL", currentEntityURL);
              console.log("currentEntity", currentEntity);
              console.log("currentsiteID", currentsiteID);
              console.log("currentDevision", currentDevision);
              console.log("currentDocumentLibrary", currentDocumentLibrary);
              if (myButton) {
                myButton.textContent = `Create Library under ${docLibName}`;
              } else {
                console.error();
              }

              toggleVisibility(folderList);
            });

            docLibElement.addEventListener("dblclick", (event) => {
              event.stopPropagation();
              toggleVisibility(folderList, false);
            });
          });

          const departmentList = document.createElement("ul");

          departmentList.style.display = "none";
          devisionElement.appendChild(departmentList);

          devisionValue.departments.forEach((departmentTitle: any) => {
            const departmentElement = document.createElement("li");
            departmentElement.textContent = departmentTitle;
            departmentElement.classList.add("folder");
            departmentElement.style.cursor = "pointer";
            departmentList.appendChild(departmentElement);

            const documentList = document.createElement("ul");
            documentList.style.display = "none";
            documentList.style.width = "300px";
            const entityImage = createImageElement(
              "icons/entity-icon.png",
              "Entity Icon"
            );
            departmentElement.appendChild(entityImage);
            departmentElement.appendChild(documentList);

            departmentElement.addEventListener("click", (event) => {
              currentEntityURL = value.siteURL;
              currentsiteID = value.siteID
              currentEntity = value.entityTitle;
              currentDevision = devisionTitle;
              currentDepartment = departmentTitle;
              currentDocumentLibrary = ''
              currentFolder = ''
              currentfolderpath = ''
              console.log("currentEntityURL", currentEntityURL);
              console.log("currentsiteID", currentsiteID);
              console.log("currentEntity", currentEntity);
              console.log("currentDevision", currentDevision);
              console.log("currentDepartment", currentDepartment);
              handleNavigation(value.entityTitle, devisionTitle, departmentTitle, null, null)
              event.stopPropagation();
              if (myButton) {
                myButton.textContent = `Create Library under ${departmentTitle}`;
              } else {
                console.error();
              }

              // Prevent toggling visibility before the list is populated
              if (documentList.innerHTML === "") {
                const key = `${entityTitle.trim()}::${devisionTitle.trim()}::${departmentTitle.trim()}`;
                const documentLibraries = folderMap.get(key) || [];
                documentList.innerHTML = "";
                const uniqueDocLibs = new Map();

                documentLibraries.forEach((item: any) => {
                  if (!uniqueDocLibs.has(item.DocumentLibraryName)) {
                    uniqueDocLibs.set(item.DocumentLibraryName, {
                      folders: [],
                      folderPath: item.FolderPath, // Store FolderPath
                    });
                  }
                  uniqueDocLibs
                    .get(item.DocumentLibraryName)
                    .folders.push(item);
                });

                uniqueDocLibs.forEach((data, docLibName) => {
                  console.log(uniqueDocLibs, "uniqueDocLibs")
                  const docLibElement = document.createElement("li");
                  docLibElement.textContent = docLibName;

                  // Optionally display the FolderPath in the docLibElement
                  // const pathText = document.createElement("span");
                  // pathText.textContent = ` (${data.folderPath})`; // Display FolderPath
                  // docLibElement.appendChild(pathText);

                  documentList.appendChild(docLibElement);

                  const folderList = document.createElement("ul");
                  folderList.style.display = "none";
                  folderList.style.width = "351px";
                  const entityImage = createImageElement(
                    "icons/entity-icon.png",
                    "Entity Icon"
                  );
                  docLibElement.appendChild(entityImage);
                  docLibElement.appendChild(folderList);

                  docLibElement.addEventListener("click", (event) => {
                    event.stopPropagation();
                    currentEntityURL = value.siteURL;
                    currentsiteID = value.siteID
                    currentEntity = value.entityTitle;
                    currentDevision = devisionTitle;
                    currentDocumentLibrary = docLibName;
                    currentDepartment = departmentTitle;
                    currentfolderpath = data.folderPath,
                      currentFolder = ''
                    // console.log(data, data, "data")
                    // console.log("currentEntityURL", currentEntityURL);
                    // console.log("currentsiteID", currentsiteID);
                    // console.log("currentEntity", currentEntity);
                    // console.log("currentDevision", currentDevision);
                    // console.log("currentDepartment", currentDepartment);
                    // console.log("currentDocumentLibrary", currentDocumentLibrary);
                    // console.log("currentfolderpath", currentfolderpath);
                    // console.log("parentfolder", parentfolder);
                    getdoclibdata(data.folderPath, value.siteID, docLibName)
                    handleNavigation(value.entityTitle, devisionTitle, departmentTitle, docLibName, null)
                    console.log(
                      "FolderPath for document library:",
                      data.folderPath
                    );
                    toggleVisibility(folderList);
                    const createFileButton = document.getElementById("createFileButton")
                    const createFileButton2 = document.getElementById("createFileButton")
                    createFileButton.style.display = "block";
                    createFileButton2.style.display = "block";
                    if (myButton) {
                      myButton.textContent = `Create Folder under ${docLibName}`;
                    } else {
                      console.error();
                    }
                  });

                  docLibElement.addEventListener("dblclick", (event) => {
                    event.stopPropagation();
                    toggleVisibility(folderList, false);
                  });
                  const buildFolderStructure = (
                    parentFolderId: any,
                    parentElement: any
                  ) => {
                    data.folders.forEach((item: any) => {

                      const folderNamesArray = Array.isArray(item.FolderName)
                        ? item.FolderName
                        : [item.FolderName];

                      folderNamesArray.forEach((folderName: any) => {

                        if (
                          folderName &&
                          item.ParentFolderId === parentFolderId
                        ) {
                          const folderElement = document.createElement("li");
                          folderElement.textContent = folderName;
                          parentElement.appendChild(folderElement);
                          const entityImage = createImageElement(
                            "icons/entity-icon.png",
                            "Entity Icon"
                          );
                          folderElement.appendChild(entityImage);
                          const subFolderList = document.createElement("ul");
                          subFolderList.style.display = "none";
                          folderElement.appendChild(subFolderList);

                          folderElement.addEventListener("click", (event) => {
                            currentEntityURL = value.siteURL;
                            currentEntity = value.entityTitle;
                            currentsiteID = value.siteID
                            currentDevision = devisionTitle;
                            currentDepartment = departmentTitle;
                            currentDocumentLibrary = docLibName;
                            currentFolder = folderName

                            // console.log("currentEntityURL", currentEntityURL);
                            // console.log("currentEntity", currentEntity);
                            // console.log("currentsiteID", currentsiteID);
                            // console.log("currentDevision", currentDevision);
                            // console.log("currentDepartment", currentDepartment);
                            // console.log("currentDocumentLibrary", currentDocumentLibrary);
                            // console.log("currentfolderpath", item.FolderPath);
                            getdoclibdata(item.FolderPath, currentsiteID, docLibName)
                            handleNavigation(value.entityTitle, devisionTitle, departmentTitle, docLibName, folderName)
                            const createFileButton = document.getElementById("createFileButton")
                            createFileButton.style.display = "block";
                            const createFileButton2 = document.getElementById("createFileButton")
                            createFileButton2.style.display = "block";
                            if (myButton) {
                              myButton.textContent = `Create Folder under ${folderName}`;
                            } else {
                              console.error();
                            }
                            event.stopPropagation();
                            toggleVisibility(subFolderList);
                            subFolderList.innerHTML = "";
                            buildFolderStructure(folderName, subFolderList);
                          });
                        }
                      });
                    });
                  };
                  buildFolderStructure(null, folderList);
                });
              }

              toggleVisibility(documentList);
            });

            departmentElement.addEventListener("dblclick", (event) => {
              event.stopPropagation();
              toggleVisibility(documentList, false);
            });
          });

          ///Start: display all Document libraries under Devision directly if Department null with nested folder //////
          const keyForDevisionOnly = `${entityTitle.trim()}::${devisionTitle.trim()}::`;

          if (folderMap.has(keyForDevisionOnly)) {
            const documentLibraries = folderMap.get(keyForDevisionOnly) || [];
            // console.log(documentLibraries, "documentLibraries");
            const uniqueDocLibNames = new Set();

            documentLibraries.forEach((item: any) => {
              const normalizedDocLibName =
                item.DocumentLibraryName.trim().toLowerCase();
              console.log("item of doclib under division", item);

              if (!uniqueDocLibNames.has(normalizedDocLibName)) {
                uniqueDocLibNames.add(normalizedDocLibName);

                const docLibElement = document.createElement("li");
                // New code to check the Document library name is Rename or not start
                // const checkIsRename = item.folders.filter((item:any) => 
                //   item.FolderName.length === 1 && item.FolderName[0] === null
                // );
                // console.log("checkIsRename",checkIsRename);
                // let renameText=item.DocumentLibraryName;
                // if(checkIsRename[0].IsRename !== null){
                //   renameText=checkIsRename[0].IsRename;
                // }
                // End
                // docLibElement.textContent = renameText;
                docLibElement.textContent = item.DocumentLibraryName;

                departmentList.appendChild(docLibElement);

                const folderList = document.createElement("ul");
                folderList.style.display = "none";
                const entityImage = createImageElement(
                  "icons/entity-icon.png",
                  "Entity Icon"
                );
                docLibElement.appendChild(entityImage);

                docLibElement.appendChild(folderList);

                docLibElement.addEventListener("click", (event) => {
                  event.stopPropagation();
                  currentEntityURL = value.siteURL; // Use the SiteURL from entitiesMap
                  currentsiteID = value.siteID
                  currentEntity = value.entityTitle
                  currentDevision = devisionTitle;
                  currentDepartment = ''
                  currentFolder = ''
                  currentDocumentLibrary = item.DocumentLibraryName;
                  currentfolderpath = item.FolderPath;
                  console.log("currentEntityURL", currentEntityURL);
                  console.log("currentsiteID", currentsiteID);
                  console.log("currentEntity", currentEntity);
                  console.log("currentDevision", currentDevision);
                  console.log("currentDepartment", currentDepartment);
                  console.log("currentDocumentLibrary", currentDocumentLibrary);
                  console.log("currentfolderpath", currentfolderpath);
                  getdoclibdata(item.FolderPath, value.siteID, item.DocumentLibraryName)
                  handleNavigation(value.entityTitle, devisionTitle, null, item.DocumentLibraryName)
                  const createFileButton = document.getElementById("createFileButton")
                  createFileButton.style.display = "block";
                  const createFileButton2 = document.getElementById("createFileButton2")
                  createFileButton2.style.display = "block";
                  if (myButton) {
                    myButton.textContent = `Create Folder under ${item.DocumentLibraryName}`;
                  } else {
                    console.error();
                  }
                  toggleVisibility(folderList);
                  folderList.innerHTML = "";
                  const buildFolderStructure = (
                    parentFolderId: any,
                    parentElement: any
                  ) => {
                    const createImageElement = (src: string, alt: string) => {
                      const img = document.createElement("img");
                      img.src = require("../../assets/add-folder.png");
                      img.alt = alt;
                      img.style.float = "left";
                      img.style.width = "20px"; // Adjust the size as needed
                      img.style.height = "20px"; // Adjust the size as needed
                      img.style.marginRight = "5px"; // Space between image and text
                      return img;
                    };
                    documentLibraries.forEach((libItem: any) => {

                      if (
                        libItem.DocumentLibraryName.trim().toLowerCase() ===
                        normalizedDocLibName
                      ) {
                        const folderNamesArray = Array.isArray(
                          libItem.FolderName
                        )
                          ? libItem.FolderName
                          : [libItem.FolderName];

                        folderNamesArray.forEach((folderName: any) => {

                          if (
                            folderName &&
                            libItem.ParentFolderId === parentFolderId
                          ) {
                            // Only display non-null folder names
                            const folderElement2 = document.createElement("li");
                            folderElement2.textContent = folderName;
                            parentElement.appendChild(folderElement2);
                            const folderPath = libItem.FolderPath;
                            const entityImage = createImageElement(
                              "icons/entity-icon.png",
                              "Entity Icon"
                            );
                            folderElement2.appendChild(entityImage);
                            const subFolderList2 = document.createElement("ul");
                            subFolderList2.style.display = "none";

                            // const entityImage = createImageElement('icons/entity-icon.png', 'Entity Icon')
                            folderElement2.appendChild(entityImage);
                            subFolderList2.appendChild(entityImage);
                            folderElement2.appendChild(subFolderList2);

                            folderElement2.addEventListener(

                              "click",
                              (event) => {
                                currentEntityURL = value.siteURL; // Use the SiteURL from entitiesMap
                                currentsiteID = value.siteID
                                currentEntity = value.entityTitle
                                currentDevision = devisionTitle;
                                currentDepartment = null
                                currentDocumentLibrary = item.DocumentLibraryName;
                                currentFolder = folderName
                                currentDepartment = ''
                                currentFolder = folderPath
                                // currentfolderpath = item.FolderPath;
                                parentfolder = parentFolderId
                                console.log("currentEntityURL", currentEntityURL);
                                console.log("currentsiteID", currentsiteID);

                                console.log("currentEntity", currentEntity);
                                console.log("currentDevision", currentDevision);
                                console.log("currentDepartment", currentDepartment);
                                console.log("currentDocumentLibrary", currentDocumentLibrary);
                                console.log("currentFolder", currentFolder);
                                console.log("currentfolderpath", folderPath);
                                console.log("parentfolder", parentfolder);
                                handleNavigation(value.entityTitle, devisionTitle, null, item.DocumentLibraryName, folderName)
                                event.stopPropagation();
                                toggleVisibility(subFolderList2);
                                console.log("enter ee");
                                getdoclibdata(folderPath, currentsiteID, item.DocumentLibraryName)
                                const createFileButton = document.getElementById("createFileButton")
                                createFileButton.style.display = "block";
                                const createFileButton2 = document.getElementById("createFileButton")
                                createFileButton2.style.display = "block";
                                if (myButton) {
                                  myButton.textContent = `Create Folder under ${folderName}`;
                                } else {
                                  console.error();
                                }

                                // Clear existing sub-folder list to avoid duplications
                                subFolderList2.innerHTML = "";

                                // Recursively build the sub-folder structure
                                buildFolderStructure(
                                  folderName,
                                  subFolderList2
                                );
                              }
                            );
                          }
                        });
                      }
                    });
                  };

                  // Start building the folder structure from the root level (null ParentFolderId)
                  buildFolderStructure(null, folderList);
                });

                // Optionally, expand the folder structure by default
                // buildFolderStructure(folderList, documentLibraries, null);
              }
            });
          }

          ///End: display all Document libraries under Devision directly if Department null with nested folder //////

          devisionElement.addEventListener("click", (event) => {
            const breadcrumbElement = document.getElementById("breadcrumb");
            if (breadcrumbElement) {
              breadcrumbElement.style.display = "none";
            }
            event.stopPropagation();
            currentDevision = devisionTitle;
            currentEntityURL = value.siteURL;
            currentEntity = value.entityTitle
            currentsiteID = value.siteID
            currentDepartment = ''
            currentDocumentLibrary = ''
            currentFolder = ''
            currentfolderpath = ''
            console.log("currentEntityURL", currentEntityURL);
            console.log("currentsiteID", currentsiteID);
            console.log("currentEntity", currentEntity);
            console.log("currentDevision", currentDevision);
            handleNavigation(value.entityTitle, devisionTitle, null, null, null)
            toggleVisibility(departmentList);
            // Toggle plus/minus icon
            devisionElement.classList.remove("expanded");
            // const //createFileButton=document.getElementById("createFileButton")
            // createFileButton.style.display="block";
            if (myButton) {
              myButton.textContent = `Create Library under ${devisionTitle}`;
            } else {
              console.error();
            }
          });

          devisionElement.addEventListener("dblclick", (event) => {
            event.stopPropagation();
            toggleVisibility(departmentList, false);
            // Toggle plus/minus icon
            devisionElement.classList.remove("expanded");
          });
        });

        let clickTimer: any;
        titleElement.addEventListener("click", async (event) => {
          const breadcrumbElement = document.getElementById("breadcrumb");
          if (breadcrumbElement) {
            breadcrumbElement.style.display = "none";
          }
          // alert("in first")
          // new code added.
          // toggle createfolder button based on the permission
          // Get the users in the group
          // const subsiteContext=await sp.site.openWebById(value.siteID);
          // const usersFromAdmin = await subsiteContext.web.siteGroups.getByName(`${value.entityTitle}_Admin`).users();
          // const usersFromInitiator=await subsiteContext.web.siteGroups.getByName(`${value.entityTitle}_Initiator`).users();
          // console.log("usersFromAdmin",usersFromAdmin);
          // console.log("usersFromInitiator",usersFromInitiator);
          try {
            const currentUser = await sp.web.currentUser();
            const userGroups = await sp.web.siteUsers.getById(currentUser.Id).groups();
            const isMemberOfGroup = userGroups.some(group => group.Title === `${currentEntity}_Admin`);
            const isMemberOfSuperAdmin = userGroups.some(group => group.Title === `DMSSuper_Admin`);
            console.log("isMemberOfSuperAdmin", isMemberOfSuperAdmin);
            console.log(`Is member of ${currentEntity}_Admin:`, isMemberOfGroup);
            // console.log(`User is a member of the group: ${currentEntity}_Admin`);
            if (isMemberOfGroup || isMemberOfSuperAdmin) {
              console.log(`User is a member of the group: ${currentEntity}_Admin`);
              if (createFileButton) {
                createFileButton.style.display = "none";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "block";
              }
            } else {
              console.log(`User is not a member of the group: ${currentEntity}_Admin`);
              if (createFileButton) {
                createFileButton.style.display = "none";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "none";
              }


            }
          } catch (error) {
            console.log(`User is not a member of the group: ${currentEntity}_Admin`);
            if (createFileButton) {
              createFileButton.style.display = "none";
            }
            if (createFileButton2) {
              createFileButton2.style.display = "none";
            }


          }
        })
        titleElement.addEventListener("click", async (event) => {
          const breadcrumbElement = document.getElementById("breadcrumb");
          if (breadcrumbElement) {
            breadcrumbElement.style.display = "none";
          }
          setdisplayuploadfileandcreatefolder(true)
          // alert("in second")
          // Toggle +/- button
          // const plusMinus = document.getElementById("toggle-plus/minus");
          if (toggleButton.textContent === "+") {
            toggleButton.textContent = "-";
          } else if (toggleButton.textContent) {
            toggleButton.textContent = "+";
          }
          const createFileButton = document.getElementById("createFileButton")
          const createFileButton2 = document.getElementById("createFileButton2")
          try {
            const currentUser = await sp.web.currentUser();
            const userGroups = await sp.web.siteUsers.getById(currentUser.Id).groups();
            const isMemberOfGroup = userGroups.some(group => group.Title === `${currentEntity}_Admin`);
            const isMemberOfSuperAdmin = userGroups.some(group => group.Title === `DMSSuper_Admin`);
            console.log("isMemberOfSuperAdmin", isMemberOfSuperAdmin);
            console.log(`Is member of ${currentEntity}_Admin:`, isMemberOfGroup);
            // console.log(`User is a member of the group: ${currentEntity}_Admin`);
            if (isMemberOfGroup || isMemberOfSuperAdmin) {
              console.log(`User is a member of the group: ${currentEntity}_Admin`);
              if (createFileButton) {
                createFileButton.style.display = "none";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "block";
              }
            } else {
              console.log(`User is not a member of the group: ${currentEntity}_Admin`);
              if (createFileButton) {
                createFileButton.style.display = "none";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "none";
              }


            }
          } catch (error) {
            console.log(`User is not a member of the group: ${currentEntity}_Admin`);
            if (createFileButton) {
              createFileButton.style.display = "none";
            }
            if (createFileButton2) {
              createFileButton2.style.display = "none";
            }


          }
          event.stopPropagation();
          // const createFileButton2 = document.getElementById("createFileButton2");
          // Clear any existing timer
          clearTimeout(clickTimer);

          // Set a new timer
          clickTimer = setTimeout(() => {
            setlistorgriddata('');
            currentEntity = value.entityTitle
            currentEntityURL = value.siteURL;
            currentsiteID = value.siteID;
            currentDevision = ""
            currentDepartment = ''
            currentDocumentLibrary = ""
            currentFolder = ""
            currentfolderpath = ""
            console.log(value.entityTitle, "value");
            console.log(currentsiteID, "currentsiteID");
            console.log("currentEntityURL", currentEntityURL);
            mydata.push(value.siteURL);
            console.log(mydata, "my mydata");
            toggleVisibility(devisionList);
            toggleVisibility(documentList);
            const hidegidvewlistviewbutton = document.getElementById("hidegidvewlistviewbutton");
            const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2");
            if (hidegidvewlistviewbutton) {
              console.log("enter here .....................");
              hidegidvewlistviewbutton.style.display = 'none';
            }
            if (hidegidvewlistviewbutton2) {
              console.log("enter here .....................");
              hidegidvewlistviewbutton2.style.display = 'none';
            }
            handleNavigation(value.entityTitle, null, null, null, null);
            // Toggle plus/minus icon
            titleElement.classList.toggle("expanded");
            console.log(value, "value");
            const createFileButton = document.getElementById("createFileButton");
            const createFileButton2 = document.getElementById("createFileButton2");
            // if (createFolderButton) {
            //     createFolderButton.style.display = "block";
            // }
            if (createFileButton) {
              createFileButton.style.display = "none";
            }
            if (createFileButton2) {
              createFileButton2.style.display = "block";
            }
            if (myButton) {
              myButton.textContent = `Create Library under ${entityTitle}`;
            } else {
              console.error();
            }
            // fetchData(currentEntityURL);
          }, 300); // Adjust the delay as needed
        });

        titleElement.addEventListener("dblclick", async (event) => {
          const breadcrumbElement = document.getElementById("breadcrumb");
          if (breadcrumbElement) {
            breadcrumbElement.style.display = "none";
          }

          setdisplayuploadfileandcreatefolder(true)
          event.stopPropagation();
          try {
            const currentUser = await sp.web.currentUser();
            const userGroups = await sp.web.siteUsers.getById(currentUser.Id).groups();
            const isMemberOfGroup = userGroups.some(group => group.Title === `${currentEntity}_Admin`);
            const isMemberOfSuperAdmin = userGroups.some(group => group.Title === `DMSSuper_Admin`);
            console.log("isMemberOfSuperAdmin", isMemberOfSuperAdmin);
            console.log(`Is member of ${currentEntity}_Admin:`, isMemberOfGroup);
            // console.log(`User is a member of the group: ${currentEntity}_Admin`);
            if (isMemberOfGroup || isMemberOfSuperAdmin) {
              console.log(`User is a member of the group: ${currentEntity}_Admin`);
              if (createFileButton) {
                createFileButton.style.display = "none";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "block";
              }
            } else {
              console.log(`User is not a member of the group: ${currentEntity}_Admin`);
              if (createFileButton) {
                createFileButton.style.display = "none";
              }
              if (createFileButton2) {
                createFileButton2.style.display = "none";
              }


            }
          } catch (error) {
            console.log(`User is not a member of the group: ${currentEntity}_Admin`);
            if (createFileButton) {
              createFileButton.style.display = "none";
            }
            if (createFileButton2) {
              createFileButton2.style.display = "none";
            }


          }
          // Clear the single click timer
          clearTimeout(clickTimer);

          setlistorgriddata('');
          toggleVisibility(devisionList, false);
          toggleVisibility(documentList, false);
          // Toggle plus/minus icon
          titleElement.classList.remove("expanded");
        });
      });
    } catch (error) {
      console.error("Error fetching or building folder tree:", error);
    }
  };
  useEffect(() => {
    //searctest();
    fetchAndBuildTree2();
  }, [])



  // code to route to different document library and folder start
  useEffect(() => {
    // const params = new URLSearchParams(window.location.search);
    const url = window.location.href;
    // const matches = url.match(/\/([^\/]+)\.aspx/);
    let extractedPart = url.split('.aspx')[1];
    let parameters = extractedPart.split('?')
    // alert( parameters);
    console.log("parameters", parameters);
    let path = "";
    let siteId = "";
    let folderName = "";
    let devision = "";
    let department = "";
    if (parameters.length > 1) {
      parameters.forEach((items, index) => {
        console.log(`items[${index}]`, items)

        if (index == 1) {
          if (items.includes('%20')) {
            console.log("Clean Url")
            const cleanUrl = items.replace(/%20/g, ' ');
            path = cleanUrl;
          } else {
            path = items;
          }

        }
        if (index == 2) {
          if (items.includes('%20')) {
            console.log("Clean path")
            const cleanUrl = items.replace(/%20/g, ' ');
            folderName = cleanUrl;
          } else {
            folderName = items;
          }
          // folderName=items;
        }
        if (index == 3) {
          siteId = items;
        }
        if (index == 4) {
          if (items.includes('%20')) {
            console.log("Clean devision")
            const cleanDevision = items.replace(/%20/g, ' ');
            devision = cleanDevision;
          } else {
            devision = items;
          }
        }
        if (index == 5) {
          if (items.includes('%20')) {
            console.log("Clean deaprtment")
            const cleanDepartment = items.replace(/%20/g, ' ');
            department = cleanDepartment;
          } else {
            department = items;
          }
        }
      })
      console.log("path", path)
      console.log("siteId", siteId)
      console.log("folderName", folderName)
      console.log("department", department)
      console.log("devision", devision)
      currentDepartment = department;
      currentDevision = devision;
      cleanUrlInMyRequest = true;
      getdoclibdata(path, siteId, folderName);
    }
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton) {
      hidegidvewlistviewbutton.style.display = 'none'
    }
  }, []);
  // end
  const getdoclibdata = async (FolderPath: any, siteID: any, docLibName: any) => { return Promise.resolve(); }
  const getdoclibdata1 = async (FolderPath: any, siteID: any, docLibName: any) => {
    setdisplayuploadfileandcreatefolder(true)
    //  ismyrequordoclibforfilepreview = "getdoclibdata"
    ismyrequordoclibforfilepreview = "getdoclibdata"
    console.log('path   ', FolderPath)
    console.log('SiteID :    ', siteID)
    console.log('docLibName :    ', docLibName);
    console.log('currentEntity :    ', currentEntity);

    // Hide the list and grid view start
    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2")
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'none'

    }
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'

    }
    // End
    // Extract the current entity from url
    const segments = FolderPath.split('/');
    const currentSubsite = segments[3];
    console.log("segments", segments);
    console.log(currentSubsite);
    console.log("Devision", currentDevision);
    console.log("Department", currentDepartment);
    // set current entity ,current document library and folder name
    const folName = segments[segments.length - 1];
    if (folName === docLibName) {
      console.log("its document libray", folName)
    } else {
      console.log("its folder", folName)
      currentFolder = folName
    }
    currentEntity = currentSubsite;
    currentDocumentLibrary = docLibName;
    // Update the url  start
    if (currentDevision !== "") {
      console.log("Devision present", currentDevision);
      if (currentDepartment !== "") {
        console.log("Department present", currentDepartment);
        const newUrl = `${window.location.origin}${window.location.pathname}?${FolderPath}?${docLibName}?${siteID}?${currentDevision}?${currentDepartment}`;
        //window.history.pushState(null, '', newUrl);
      } else {
        const newUrl = `${window.location.origin}${window.location.pathname}?${FolderPath}?${docLibName}?${siteID}?${currentDevision}`;
        // window.history.pushState(null, '', newUrl);
      }
    } else {
      const newUrl = `${window.location.origin}${window.location.pathname}?${FolderPath}?${docLibName}?${siteID}`;
      //window.history.pushState(null, '', newUrl);
    }

    // end

    routeToDiffSideBar = "";

    const testidsub = await sp.site.openWebById(siteID);
    let files: any = [];
    let batchSize = 5000;
    let nextLink = null;
    let hasMoreItems = true;
    currentsiteID = siteID;
    currentfolderpath = FolderPath;
    const container = document.getElementById("files-container");
    container.innerHTML = "";
    console.log("folderpath:", FolderPath);
    try {
      while (hasMoreItems) {
        let response;
        if (nextLink) {
          response = await sp.web(nextLink);
        } else {
          response = await testidsub.web
            .getFolderByServerRelativePath(FolderPath)
            .files.select("Name", "Length", "ServerRelativeUrl", "UniqueId", "MajorVersion", "ListItemAllFields/Status", "ListItemAllFields/IsDeleted").expand("ListItemAllFields")
            .top(batchSize)();
          myfolderdata = response
          console.log(response, "response")
        }
        // Add the current batch of files to the files array
        files = [...files, ...response as IFileInfo[]];
        // Check if there is a nextLink for more items
        if ("@odata.nextLink" in response) {
          nextLink = response["@odata.nextLink"];
        } else {
          hasMoreItems = false; // No more items, exit loop
        }
      }
      console.log("All files fetched:", files);

      // Get the details of the users permission.
      // start
      // const library = testidsub.web.lists.getByTitle(docLibName);
      const library = testidsub.web.getList(`${FolderPath}`);
      const permissions = await library.getCurrentUserEffectivePermissions();
      // console.log("permissions",permissions);

      // Check for all permissions
      //   const userPermissions = {
      //     canViewPages: testidsub.web.hasPermissions(permissions, PermissionKind.ViewPages),
      //     canView: testidsub.web.hasPermissions(permissions, PermissionKind.ViewListItems),
      //     canEdit: testidsub.web.hasPermissions(permissions, PermissionKind.EditListItems),
      //     canAdd: testidsub.web.hasPermissions(permissions, PermissionKind.AddListItems),          
      //     canFullControl: testidsub.web.hasPermissions(permissions, PermissionKind.FullMask),
      //     canFullControl1: testidsub.web.hasPermissions(permissions, PermissionKind.ManagePermissions),
      //     canDelete: testidsub.web.hasPermissions(permissions, PermissionKind.DeleteListItems),
      //     canApprove: testidsub.web.hasPermissions(permissions, PermissionKind.ApproveItems),
      //     canOpen: testidsub.web.hasPermissions(permissions, PermissionKind.OpenItems),
      //     canViewVersions: testidsub.web.hasPermissions(permissions, PermissionKind.ViewVersions),
      //     canDeleteVersions: testidsub.web.hasPermissions(permissions, PermissionKind.DeleteVersions),
      //     canManagePermissions: testidsub.web.hasPermissions(permissions, PermissionKind.ManagePermissions),
      //     canViewFormPages: testidsub.web.hasPermissions(permissions, PermissionKind.ViewFormPages),
      //     canEditMyUserInfo: testidsub.web.hasPermissions(permissions, PermissionKind.EditMyUserInfo)
      // };
      const userPermissions = {
        hasFullControl: testidsub.web.hasPermissions(permissions, PermissionKind.FullMask) || testidsub.web.hasPermissions(permissions, PermissionKind.ManagePermissions),
        hasContribute: testidsub.web.hasPermissions(permissions, PermissionKind.AddListItems) &&
          testidsub.web.hasPermissions(permissions, PermissionKind.EditListItems) &&
          testidsub.web.hasPermissions(permissions, PermissionKind.DeleteListItems),
        hasEdit: testidsub.web.hasPermissions(permissions, PermissionKind.EditListItems),
        // hasEdit1: testidsub.web.hasPermissions(permissions, PermissionKind.),
        hasRead: testidsub.web.hasPermissions(permissions, PermissionKind.ViewListItems),
        hasView: testidsub.web.hasPermissions(permissions, PermissionKind.ViewPages)
      };
      console.log("userPermissions", userPermissions);
      // End

      // Belong to admin or not start
      // Toggle the createFile and createFolder button based on the permission
      // let permission:string;
      const createFileButton = document.getElementById("createFileButton");
      const createFileButton2 = document.getElementById("createFileButton2");
      if (userPermissions.hasFullControl) {
        console.log(`Current User has full control on the library/Folder`);
        if (createFileButton) {
          createFileButton.style.display = "block";
        }
        if (createFileButton2) {
          createFileButton2.style.display = "block";
        }
      } else if (userPermissions.hasContribute || userPermissions.hasEdit) {
        console.log(`Current User has Contribute/Edit permission on the library/Folder`);
        if (createFileButton) {
          createFileButton.style.display = "block";
        }
        if (createFileButton2) {
          createFileButton2.style.display = "none";
        }
      } else {
        console.log(`Current User has no permission on the library/Folder`);
        if (createFileButton) {
          createFileButton.style.display = "none";
        }
        if (createFileButton2) {
          createFileButton2.style.display = "none";
        }
      }

      const DMSEntityFileMasterList = `DMS${currentEntity}FileMaster`;
      console.log(DMSEntityFileMasterList);

      const filesData = await sp.web.lists
        .getByTitle(`${DMSEntityFileMasterList}`)
        .items.select("FileUID", "IsFavourite")
        .filter(
          `IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`
        )();

      // Create a map for quick lookup of IsFavourite status by FileUID
      const favouriteMap = new Map(
        filesData.map((item: any) => [item.FileUID, item.IsFavourite])
      );


      // console.log("FavouriteMap",favouriteMap)
      console.log("Files", filesData);
      // Add breadCrumb start
      handleNavigation(currentSubsite, currentDevision, currentDepartment, currentDocumentLibrary, currentFolder);
      // End
      const container = document.getElementById("files-container");
      container.innerHTML = "";
      files.forEach(async (file: any) => {
        const isFavourite = favouriteMap.get(file.UniqueId) || 0;
        const favouriteText = isFavourite ? "Unmark as Favourite" : "Mark as Favourite";

        // Set display properties based on favorite status
        const displayPropertyforFillFavourite = isFavourite ? "block" : "none";
        const displayPropertyforUnFillFavourite = isFavourite ? "none" : "block";

        if (file.ListItemAllFields.IsDeleted === null) {
          if (file.ListItemAllFields.Status !== "Pending") {
            let permission = file.ListItemAllFields.Status;
            const { fileIcon } = getFileIcon(file.Name);
            const card = createFileCardForDocumentLibrary(file, fileIcon, siteID, false, docLibName, displayPropertyforUnFillFavourite, displayPropertyforFillFavourite, favouriteText, permission, FolderPath,);
            container.appendChild(card);
          }
        }
      });
    } catch (error) {
      console.error("Error fetching Doclib data:", error);
    }

  };

  //   window.documentLibraryPopUp = async function(fileId: string , siteID:any , FolderPath:any , FileName:any,permission:any) {
  //     console.log("Inside the documentLibraryPopUp");
  //     console.log(siteID, "siteID")
  //     console.log(fileId , "fileId")
  //     console.log(FolderPath , "folderPath")
  //     console.log(FileName , "fileName")
  //     console.log(typeof siteID, "siteID typeof")
  //     console.log(typeof fileId , "fileId type of")
  //     console.log(typeof FolderPath , "folderPath typeof")
  //     console.log(typeof FileName , "fileName typeof")
  //     // check user permission on item start
  //     const testidsub =await sp.site.openWebById(siteID)
  //     let filePermission:string;
  //     let filePath=`${FolderPath}/${FileName}`;
  //     console.log("filePath",filePath);
  //     const fileServerRelativePath = testidsub.web.getFileByServerRelativePath(filePath);
  //     // Retrieve the list item associated with the file
  //     const item = await fileServerRelativePath.getItem();
  //     console.log("items",item);
  //     // Get current user permissions on the item (file)
  //     const filePermissions = await item.getCurrentUserEffectivePermissions(); 
  //     console.log("File permissions:", filePermissions);
  //     // console.log("file listItems All field",file.ListItemAllFields);

  //     const hasFullControl = testidsub.web.hasPermissions(filePermissions, PermissionKind.ManageWeb);
  //     const hasEdit = testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
  //     const hasContribute = testidsub.web.hasPermissions(filePermissions, PermissionKind.AddListItems) && testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
  //     const hasRead = testidsub.web.hasPermissions(filePermissions, PermissionKind.ViewListItems);
  //     console.log(hasFullControl , "hasFullControl")
  //     console.log(hasEdit , "hasEdit")
  //     console.log(hasContribute , "hasContribute")
  //     console.log(hasRead , "hasRead")

  //     if (hasFullControl) {
  //       filePermission ="Full Control";
  //     } else if (hasEdit) {
  //       filePermission ="Edit";
  //     } else if (hasContribute) {
  //       filePermission = "Contribute";
  //     } else if (hasRead) {
  //       filePermission = "Read";
  //     } else {
  //       filePermission = "No Access";
  //     }

  //     console.log("filePermission",filePermission);
  //     let statusPermission=false;
  //     if(permission === "Approved"){
  //         statusPermission=true;
  //     }
  //     // check user permission on item End

  //   // console.log("enter here i n menu card")
  //   const allMenus = document.querySelectorAll('.popup-menu');
  //   console.log(allMenus , "allMenus")
  //   allMenus.forEach(menu => {
  //     // console.log(menu , "menu")
  //     // console.log(menu.id , "menu.id")
  //     // console.log(fileId , "fileId")
  //     if (menu.id !== `menu-${fileId}`) {
  //       menu.classList.remove("show");
  //     }
  //   });

  //   // Toggle the menu for the clicked card
  //   const menu = document.getElementById(`menu-${fileId}`);
  //   if (menu) {
  //     // console.log("Toggle the menu for the clicked card")
  //     // if(filePermission === "Edit" || filePermission === "Contribute" || filePermission === "Read" ){
  //     //     menu.children[0].children[0].remove();
  //     //     // check for read only
  //     //     if(filePermission === "Read"){
  //     //       menu.children[0].children[0].remove();
  //     //       menu.children[0].children[1].remove();
  //     //     }
  //     // }
  //     // menu.classList.toggle("show");

  //   const menu = document.getElementById(`menu-${fileId}`);
  //   if (!menu) return; 
  //   if (filePermission === "Edit" || filePermission === "Contribute" || filePermission === "Read") {
  //     const firstItem = menu.children[0]?.children[0] as HTMLElement;
  //     const secondItem = menu.children[0]?.children[1] as HTMLElement;
  //     const secondItem3 = menu.children[0]?.children[3] as HTMLElement;
  //     const secondItem4 = menu.children[0]?.children[4] as HTMLElement;

  //     if (firstItem && firstItem.style.display !== "none") {
  //           firstItem.style.display = "none";
  //     }
  //     if (filePermission === "Read" && secondItem && secondItem.style.display !== "none") {
  //           secondItem.style.display = "none";
  //     }
  //     if (filePermission === "Read" && secondItem3 && secondItem3.style.display !== "none") {
  //         secondItem3.style.display = "none";
  //     }
  //     if (filePermission === "Read" && secondItem4 && secondItem4.style.display !== "none") {
  //       secondItem4.style.display = "none";
  //    }
  //   }

  //   if(statusPermission === true){
  //     const firstItem = menu.children[0]?.children[0] as HTMLElement;
  //     if (firstItem && firstItem.style.display !== "none") {
  //         firstItem.style.display = "none";
  //     }
  //   }


  //   menu.classList.toggle("show");
  //   }


  //   document.addEventListener('click', (event) => {

  //     // console.log("Outside click Event Called");

  //     const target = event.target as HTMLElement;

  //     // Check if the click was inside any menu or three-dot icon
  //     const isClickInsideMenu = target.closest('.popup-menu');
  //     const isClickInsideThreeDots = target.closest('.three-dots');

  //     // console.log("This is nested folder",isClickInsideThreeDots);

  //     if (!isClickInsideMenu && !isClickInsideThreeDots) {
  //       const allMenus = document.querySelectorAll('.popup-menu');
  //       allMenus.forEach(menu => {
  //         menu.classList.remove('show');
  //       });
  //     }
  //   });
  // }
  const createFileCardForDocumentLibrary = (file: any, fileIcon: any, siteID: string, IsHardDelete: boolean, docLibName: string, displayPropertyforUnFillFavourite: any, displayPropertyforFillFavourite: any, favouriteText: any, permission: any, FolderPath: any) => {
    // console.log("permission",permission);
    const card = document.createElement("div");
    card.className = "card";
    card.dataset.fileId = file.UniqueId;
    card.innerHTML = `  
        <div class="row">
          <div class="col-md-2 pe-0">
        <img class="filextension" src=${fileIcon} alt="File icon"/>
        </div>
         <div class="col-md-10 pe-0">
        <p class="p1st">${file.Name}</p>
          <p class="p3rd">${((file.Length as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
         </div>
         </div>
          <div id="three-dots" class="three-dots" onclick="documentLibraryPopUp('${file.UniqueId}', '${siteID}','${FolderPath}','${file.Name}','${permission}')">
          <span>...</span>
          </div>
        `;

    const menu = document.createElement("div");
    menu.id = `menu-${file.UniqueId}`;
    menu.className = "popup-menu";
    // Conditionally add the delete option based on permission
    // let deleteOptionHTML = "";
    // if (permission === "Admin") {
    //     deleteOptionHTML = `
    //         <li onclick="confirmDeleteFile('${file.UniqueId}', '${siteID}', '${IsHardDelete}', '${null}')">
    //             <img src=${deleteIcon} alt="Delete"/>
    //             Delete
    //         </li>
    //     `;
    // }
    menu.innerHTML = `
        <ul>
          <li onclick="confirmDeleteFile('${file.UniqueId}', '${siteID}', '${IsHardDelete}', '${null}')">
                  <img src=${deleteIcon} alt="Delete"/>
                  Delete
          </li>
          <li onclick="auditHistory('${file.UniqueId}', '${siteID}','${currentDocumentLibrary}','${currentEntity}')">
          <img src=${editIcon} alt="Edit"/>
                      Audit History
          </li>
          <li onclick="PreviewFile('${file.ServerRelativeUrl}', '${siteID}' , '${docLibName}')">
          <img src=${editIcon} alt="Preview"/>
                      Preview File
          </li>
          <li id="favouriteToggle-${file.UniqueId}" onclick="toggleFavourite('${file.UniqueId}', '${siteID}')">
          <img src=${UnFillFavouriteFile} alt="Mark as Favourite" class="mark-as-favourite" style="display:${displayPropertyforUnFillFavourite};"/>
          <img src=${FillFavouriteFile} alt="Unmark as Favourite" class="unmark-as-favourite" style="display:${displayPropertyforFillFavourite};"/>
          <span class="favourite-text">${favouriteText}</span>
          </li>
          <li onclick="shareFile('${file.UniqueId}','${siteID}','${FolderPath}','${file.Name}','DocumentLibrary','${file.MajorVersion}','${((file.Length as unknown as number) / (1024 * 1024)).toFixed(2)}','${file.ListItemAllFields.Status}','','${currentDocumentLibrary}')">
          <img src=${ShareFile} alt="Share"/> Share
          </li>  
        </ul>
      `;
    card.appendChild(menu);
    return card;
  }
  // Helper function to determine the file icon based on file extension
  // const getFileIcon = (fileName: string) => {
  //   console.log(fileName , "filenmae")
  //   const fileExtension = fileName.split(".").pop().toLowerCase();
  //   switch (fileExtension) {
  //     case "doc":
  //     case "docx":
  //       return require("../../assets/DOC.png");
  //     case "txt":
  //       return require("../../assets/TXT.png");
  //     case "pdf":
  //       return require("../../assets/PDF.png");
  //     case "xls":
  //     case "xlsx":
  //       return require("../../assets/XLS.png");
  //     case "zip":
  //       return require("../../assets/ZIP.png");
  //     default:
  //       return require("../../assets/DOC.png");
  //   }
  // };

  // This function give the File Icon
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

  window.PreviewFile = function (path: any, SiteID: any, docLibName: any, flag: string, filepreviewurl) {
    console.log(docLibName, "docLibName")
    console.log(filepreviewurl, "filepreviewurl")
    console.log("path", path);
    const segments = path.split('/');
    // extarct the current entity start
    const currentSubsite = segments[3];
    // end
    // Find the index of 'sites'
    const sitesIndex = segments.indexOf('sites');

    // If 'sites' is found and there are enough segments after it
    let myactualdoclib
    if (sitesIndex !== -1 && segments.length > sitesIndex + 3) {
      myactualdoclib = segments[sitesIndex + 3];
      console.log(myactualdoclib, "myactualdoclib")
      // return segments[sitesIndex + 3];  // The document library is the 4th segment after 'sites'
    } else {
      // return null;  // Return null if not enough segments are available
    }
    event.preventDefault()
    event.stopPropagation()
    const createpreviewdiv = document.createElement('div')
    createpreviewdiv.style.display = 'grid'
    const previewfileframe = document.createElement('iframe')
    previewfileframe.id = 'filePreview'
    previewfileframe.style.width = '930px'
    previewfileframe.style.height = '500px'
    const librarydiv = document.getElementById('files-container')
    const createbutton = document.createElement('button')
    createbutton.textContent = 'Close File Preview';
    console.log("enter here in preview : ", path)
    const encodedFilePath = encodeURIComponent(path);
    console.log(encodedFilePath, "encodedFilePath");

    // Extract the parent folder correctly
    const parentFolder = path.substring(0, path.lastIndexOf('/'));
    console.log(parentFolder, "parentFolder");

    // Correctly encode the parent folder
    const encodedParentFolder = encodeURIComponent(parentFolder);

    // Get the base site URL
    const siteUrl = window.location.origin;
    console.log(siteUrl, "siteUrl");

    console.log(path, ".....path")
    if (ismyrequordoclibforfilepreview === "myRequest" || ismyrequordoclibforfilepreview === "sharewithme") {
      const previewUrl = filepreviewurl

      console.log(previewUrl, "Generated preview URL");

      console.log("Generated Preview URL:", previewUrl);
      if (previewUrl) {
        librarydiv.innerHTML = "";
        previewfileframe.src = previewUrl;
        createpreviewdiv.appendChild(createbutton)
        createpreviewdiv.appendChild(previewfileframe);
        librarydiv.appendChild(createpreviewdiv)
        createbutton.addEventListener('click', function () {
          event.preventDefault()
          event.stopPropagation()
          // alert('Button was clicked!');
          if (ismyrequordoclibforfilepreview === "myRequest") {
            myRequest();
          }
          if (ismyrequordoclibforfilepreview === "sharewithme") {
            ShareWithMe();
          }
          // if(flag === "shareWithMe"){
          //     ShareWithMe(null,null);
          // }
          // if(flag === "documentLibrary"){
          //   getdoclibdata(currentfolderpath , currentsiteID , currentDocumentLibrary)
          // }

        });
      }
    }
    if (ismyrequordoclibforfilepreview === "getdoclibdata") {
      // Generate the correct preview URL
      // const previewUrl = `${siteUrl}/sites/IntranetUAT/${currentEntity}/${myactualdoclib}/Forms/AllItems.aspx?id=${path}&parent=${encodedParentFolder}`;
      const previewUrl = `${siteUrl}/sites/IntranetUAT/${currentEntity}/${myactualdoclib}/Forms/AllItems.aspx?id=${path}&parent=${encodedParentFolder}`;
      // const previewUrl = `${siteUrl}/sites/SPFXDemo/${currentEntity}/${myactualdoclib}/Forms/AllItems.aspx?id=${path}&parent=${encodedParentFolder}`;

      console.log(previewUrl, "Generated preview URL");

      console.log("Generated Preview URL:", previewUrl);
      if (previewUrl) {
        librarydiv.innerHTML = "";
        previewfileframe.src = previewUrl;
        createpreviewdiv.appendChild(createbutton)
        createpreviewdiv.appendChild(previewfileframe);
        librarydiv.appendChild(createpreviewdiv)
        createbutton.addEventListener('click', function () {
          event.preventDefault()
          event.stopPropagation()
          // alert('Button was clicked!');
          // if(flag === "shareWithMe"){
          //     ShareWithMe(null,null);
          // }
          // if(flag === "documentLibrary"){
          //   getdoclibdata(currentfolderpath , currentsiteID , currentDocumentLibrary)
          // }
          getdoclibdata(currentfolderpath, currentsiteID, currentDocumentLibrary)
        });
      }
    }


  }

  const searchFiles = async (event: React.FormEvent) => {
    event.preventDefault();
    event.stopPropagation();

    // const searchInput = document.getElementById('searchinput') as HTMLInputElement;
    // const searchText = searchInput.value;
    // console.log(searchText , "searchText")


    // if (searchText !== "" ) {
    //     try {
    //       console.log(currentfolderpath, "currentfolderpath");
    //       console.log(searchQueryPath, "currentfolderpath");
    //         const searchQuery:SearchQueryInit = {
    //              //Querytext:`${searchText} AND Path:"https://officeindia.sharepoint.com${currentfolderpath}"`,
    //             Querytext:`${searchText} AND Path:"${searchQueryPath}"`,
    //             // Querytext: `"${searchText}"`,
    //             RowLimit: 500,
    //             SelectProperties: ["Title", "Path", "FileExtension", "UniqueId", "Size", "Created", "Modified","Author","ModifiedBy"], 
    //             TrimDuplicates:true
    //             // Additional file properties
    //             // Refiners: 'FileExtension',
    //             // RefinementFilters: ['FileExtension:equals("docx")',
    //             //                     'FileExtension:equals("pdf")',
    //             //                     'FileExtension:equals("pptx")',
    //             //                   ],  
    //             // TrimDuplicates: false
    //         };
    //         // Performing the search
    //         const searchResults = await sp.search(searchQuery);
    //         const files = searchResults.PrimarySearchResults;


    //         // console.log("routeToDiffSideBar",routeToDiffSideBar);

    //         console.log(files, "files");
    //         // Clear the previous results
    //         const container = document.getElementById("files-container");
    //         container.innerHTML = "";

    //         // Display the search results
    //         // start
    //       if( routeToDiffSideBar === "" ){
    //             files.forEach((file: any) => {
    //                 const card = document.createElement("div");
    //                 const {fileIcon} = getFileIcon(file.Title);  
    //                 card.className = "card";
    //                 card.dataset.fileId = file.UniqueId;
    //                 // console.log(file.UniqueId , "file.UniqueId")
    //                 card.innerHTML = `
    //                    <div class="row">
    //           <div class="col-md-2 pe-0">
    //                     <img class="filextension" src=${fileIcon} alt="File icon"/>
    //                    </div>
    //                     <div class="col-md-10 pe-0">
    //                    <p class="p1st">${file.Title}</p>
    //                     <p class="p3rd">${((file.Size as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
    //                    </div></div>
    //                     <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.UniqueId}', '${currentsiteID}')">
    //                       <span>...</span>
    //                     </div>
    //                 `;
    //                 const menu = document.createElement("div");
    //                 menu.id = `menu-${file.UniqueId}`;
    //                 menu.className = "popup-menu";
    //                 menu.innerHTML = `
    //                   <ul>
    //                     <li onclick="confirmDeleteFile('${file.UniqueId}', '${currentsiteID}')">
    //                       <img src=${deleteIcon} alt="Delete"/>
    //                       Delete
    //                     </li>
    //                     <li onclick="auditHistory('${file.UniqueId}', '${currentsiteID}','${file.Title}')">
    //                       <img src=${editIcon} alt="Edit"/>
    //                       Audit History
    //                     </li>
    //                     <li onclick="PreviewFile('${file.ServerRelativeUrl}', '${currentsiteID}' , '${currentDocumentLibrary}')">
    //                       <img src=${editIcon} alt="Preview"/>
    //                       Preview File
    //                     </li>
    //                     <li id="favouriteToggle-${file.UniqueId}" onclick="toggleFavourite('${file.UniqueId}', '${currentsiteID}')">
    //                       <img src=${UnFillFavouriteFile} alt="Mark as Favourite" class="mark-as-favourite"/>
    //                       <img src=${FillFavouriteFile} alt="Unmark as Favourite" class="unmark-as-favourite" style="display:none;"/>
    //                       <span class="favourite-text">Mark as Favourite</span>
    //                     </li>  
    //                   </ul>
    //                 `;

    //                 card.appendChild(menu);
    //                 container.appendChild(card);
    //             });
    //       }else{
    //           if( routeToDiffSideBar === "myRequest" ){
    //               myRequest(null,null,searchInput);
    //           }

    //           else if( routeToDiffSideBar === "myFavourite" ){

    //                 // console.log("myFavourite");
    //                 myFavorite(null,null,searchInput);

    //           }
    //           else if( routeToDiffSideBar === "myFolder"){
    //                 // console.log("Inside search => myFolder");
    //                 mycreatedfolders(event,searchInput);
    //           }
    //           else if(routeToDiffSideBar === "shareWithOthers"){
    //               ShareWithOther(null,searchInput);
    //           }
    //           else if(routeToDiffSideBar === "shareWithMe"){
    //               ShareWithMe(null,searchInput);
    //           }
    //       }
    //       // end
    //     } catch (error) {
    //         console.error("Error searching files: ", error);
    //     }
    // }

    await runSearch();

  };

  const searctest=async()=>{

    debugger
      let graphcl=await (props.context as BaseWebPartContext).msGraphClientFactory.getClient("3");
      let mssearch=new GraphSearchHelper(graphcl);
   
      debugger
  }
  const runSearch = async (searhcpath?: string,filters:string='') => {

    const searchInput = document.getElementById('searchinput') as HTMLInputElement;
    const searchText = searchInput.value;
    console.log(searchText, "searchText")

    try {

      // const searchQuery: SearchQueryInit = {
      //   //Querytext:`${searchText} AND Path:"https://officeindia.sharepoint.com${currentfolderpath}"`,
      //   // Querytext: `${searchText} IsDocument:True Path:"${(searhcpath) ? searhcpath : searchQueryPath}"`,

      //   Querytext:"",
   
      //   RowLimit: 20,
      //   SelectProperties: ["Title", "Path", "FileExtension", "UniqueId", "Size", "Created", "Modified", "Author", "ModifiedBy"],
      //   SourceId: "eff79881-1eb9-4f5e-bba5-86d48c22960c",
      //   // TrimDuplicates: false
      // };
      // // Performing the search
      // const searchResults = await sp.search(searchQuery);
      // const files = searchResults.PrimarySearchResults;

      let qyerytext=`${searchText} IsDocument:True ${filters} Path:"${(searhcpath) ? searhcpath : searchQueryPath}"`;
      let graphcl=await (props.context as BaseWebPartContext).msGraphClientFactory.getClient("3");
      let mssearch=new GraphSearchHelper(graphcl);
    
      let searchres=await mssearch.searchFiles(qyerytext,500);

      let resultsdoc: IDocumentDisplayFields[] = searchres.map(filehit => {
        let file:Partial<ISearchHitResource> =filehit.resource; 
        let tRes: IDocumentDisplayFields = {
          Title: file.name,
          Size: file.size,
          Extension: file.name.split('.').pop(),
          Path:file.webUrl

        }
        return tRes;
      })

      setSearchDocResults(resultsdoc);

      

    } catch (error) {
       console.log("Graph search error",error);
    }

  }

  const runSearch1 = async () => {
    const searchInput = document.getElementById('searchinput') as HTMLInputElement;
    const searchText = searchInput.value;
    console.log(searchText, "searchText")



    //if (searchText !== "" ) {
    try {
      console.log(currentfolderpath, "currentfolderpath");
      console.log(searchQueryPath, "currentfolderpath");
      const searchQuery: SearchQueryInit = {
        //Querytext:`${searchText} AND Path:"https://officeindia.sharepoint.com${currentfolderpath}"`,
        Querytext: `${searchText} IsDocument:True Path:"${searchQueryPath}"`,
        // Querytext: `"${searchText}"`,
        RowLimit: 5000,
        SelectProperties: ["Title", "Path", "FileExtension", "UniqueId", "Size", "Created", "Modified", "Author", "ModifiedBy"],
        TrimDuplicates: false
        // Additional file properties
        // Refiners: 'FileExtension',
        // RefinementFilters: ['FileExtension:equals("docx")',
        //                     'FileExtension:equals("pdf")',
        //                     'FileExtension:equals("pptx")',
        //                   ],  
        // TrimDuplicates: false
      };
      // Performing the search
      const searchResults = await sp.search(searchQuery);
      const files = searchResults.PrimarySearchResults;


      // console.log("routeToDiffSideBar",routeToDiffSideBar);

      console.log(files, "files");
      // Clear the previous results
      const container = document.getElementById("files-container1");
      container.innerHTML = "";

      // Display the search results
      // start
      if (routeToDiffSideBar === "") {
        files.forEach((file: any) => {
          const card = document.createElement("div");
          const { fileIcon } = getFileIcon(file.Title);
          card.className = "card";
          card.dataset.fileId = file.UniqueId;
          // console.log(file.UniqueId , "file.UniqueId")
          card.innerHTML = `
                     <div class="row">
            <div class="col-md-2 pe-0">
                      <img class="filextension" src=${fileIcon} alt="File icon"/>
                     </div>
                      <div class="col-md-10 pe-0">
                     <p class="p1st">${file.Title}</p>
                      <p class="p3rd">${((file.Size as unknown as number) / (1024 * 1024)).toFixed(2)} MB</p>
                     </div></div>
                      <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.UniqueId}', '${currentsiteID}')">
                        <span>...</span>
                      </div>
                  `;
          const menu = document.createElement("div");
          menu.id = `menu-${file.UniqueId}`;
          menu.className = "popup-menu";
          menu.innerHTML = `
                    <ul>
                      <li onclick="confirmDeleteFile('${file.UniqueId}', '${currentsiteID}')">
                        <img src=${deleteIcon} alt="Delete"/>
                        Delete
                      </li>
                      <li onclick="auditHistory('${file.UniqueId}', '${currentsiteID}','${file.Title}')">
                        <img src=${editIcon} alt="Edit"/>
                        Audit History
                      </li>
                      <li onclick="PreviewFile('${file.ServerRelativeUrl}', '${currentsiteID}' , '${currentDocumentLibrary}')">
                        <img src=${editIcon} alt="Preview"/>
                        Preview File
                      </li>
                      <li id="favouriteToggle-${file.UniqueId}" onclick="toggleFavourite('${file.UniqueId}', '${currentsiteID}')">
                        <img src=${UnFillFavouriteFile} alt="Mark as Favourite" class="mark-as-favourite"/>
                        <img src=${FillFavouriteFile} alt="Unmark as Favourite" class="unmark-as-favourite" style="display:none;"/>
                        <span class="favourite-text">Mark as Favourite</span>
                      </li>  
                    </ul>
                  `;

          card.appendChild(menu);
          container.appendChild(card);
        });
      } else {
        if (routeToDiffSideBar === "myRequest") {
          myRequest(null, null, searchInput);
        }

        else if (routeToDiffSideBar === "myFavourite") {

          // console.log("myFavourite");
          myFavorite(null, null, searchInput);

        }
        else if (routeToDiffSideBar === "myFolder") {
          // console.log("Inside search => myFolder");
          mycreatedfolders(event, searchInput);
        }
        else if (routeToDiffSideBar === "shareWithOthers") {
          ShareWithOther(null, searchInput);
        }
        else if (routeToDiffSideBar === "shareWithMe") {
          ShareWithMe(null, searchInput);
        }
      }
      // end
    } catch (error) {
      console.error("Error searching files: ", error);
    }
    //}

  }


  const ShareWithOther = async (event: React.MouseEvent<HTMLButtonElement> = null, searchText: HTMLInputElement = null) => {
    setdisplayuploadfileandcreatefolder(false)
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }
    // Hide the list and grid view start
    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2")
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'none'

    }
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'

    }
    // End
    // clean the url start
    const newUrl = `${window.location.origin}${window.location.pathname}`;
    window.history.pushState(null, '', newUrl)
    // New code to hide the create file and folder button start
    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }
    //End 
    // end
    // New Code start
    const DMSShareWithOtherMaster = await sp.web.lists
      .getByTitle("DMSShareWithOtherMaster")
      .items.select("FileName", "FileUID", "FileVersion", "FileSize", "CurrentUser", "DocumentLibraryName", "CurrentFolderPath", "ShareWithOthers", "Status", "SiteID", "SiteName", "FilePreviewURL", "ShareAt", "UserID", "PermissionType")
      .filter(`CurrentUser eq '${currentUserEmailRef.current}'`).orderBy("Created", false)();

    // Mapped the file with the users
    // Mapped the file with the users
    const groupedData = DMSShareWithOtherMaster.reduce((acc, item) => {
      const key = `${item.FileUID}-${item.FileName}`;

      if (!acc[key]) {
        acc[key] = {
          FileUID: item.FileUID,
          FileName: item.FileName,
          FileVersion: item.FileVersion,
          SiteID: item.SiteID,
          FileSize: item.FileSize,
          Users: []
        };
      }

      acc[key].Users.push({
        User: item.ShareWithOthers,
        UserID: item.UserID,
        PermissionType: item.PermissionType,
        ShareAt: item.ShareAt
      });

      return acc;
    }, {})

    // Convert the result back to an array
    const result = [];
    for (let key in groupedData) {
      result.push(groupedData[key]);
    }

    console.log("DMSShareWithOtherMaster", DMSShareWithOtherMaster);
    console.log("result", result);
    const container = document.getElementById("files-container");
    container.innerHTML = "";

    routeToDiffSideBar = "shareWithOthers"
    let filteredFileData = [];
    if (searchText !== null) {
      filteredFileData = result.filter((file: any) => file?.FileName?.toLowerCase().includes(searchText?.value?.toLowerCase()))
      if (filteredFileData.length === 0 && searchText !== null) {
        console.log("combineArray", filteredFileData);
        fileNotFound(`No file match ${searchText.value}`);
      }
    } else {
      filteredFileData = result;
    }

    filteredFileData.forEach((file) => {

      // Get the first two users
      const firstTwoUsers = file.Users.slice(0, 2);
      // Remaining users count
      const moreUsersCount = file.Users.length - 2;

      // Create shared users HTML for the first two users
      let sharedUsersHTML = firstTwoUsers.map((user: any) => {
        let firstNameInitial = "";
        let lastNameInitial = "";

        if (user.User) {
          const nameParts = user.User.split(" ");
          // Assign initials based on the number of name parts
          if (nameParts.length > 0) {
            firstNameInitial = nameParts[0].charAt(0).toUpperCase();
          }
          if (nameParts.length > 1) {
            // Use the last part as the last name initial
            lastNameInitial = nameParts[nameParts.length - 1].charAt(0).toUpperCase();
          }
        }

        return `<span class="shared-user">${firstNameInitial}${lastNameInitial}</span>`;
      })
        .join("");

      // If there are more users, add "+more"
      if (moreUsersCount > 0) {
        sharedUsersHTML += `<span class="more-users">+${moreUsersCount} more</span>`;
      }
      const { fileIcon, fileExtension } = getFileIcon(file.FileName);
      const card = document.createElement("div");
      card.className = "card";
      card.dataset.fileId = file.FileUID;
      card.dataset.listId = file.SiteID;

      card.innerHTML = `  
<div class="row">

    <div class="col-md-2 pe-0">
    <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
    </div>
    <div class="col-md-10 pe-0">
    <p class="p1st">${file.FileName}</p>
    <div class="fileSizeAndVersion">
    <p class="p3rd">${file.FileSize} MB</p>
    </div>
    </div>
</div>
<div class="sharedFile">
      ${sharedUsersHTML}
</div>
  `;
      container.appendChild(card);
    })
    // End



  }

  //Toggle the menu card for share with me
  // @ts-ignore
  const ShareWithMe = async (event: React.MouseEvent<HTMLButtonElement> = null, searchText: HTMLInputElement = null) => {
    setdisplayuploadfileandcreatefolder(false)
    ismyrequordoclibforfilepreview = "sharewithme"

    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }

    // Hide the list and grid view start
    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2")
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'none'

    }
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'

    }
    // End

    // clean the url start
    const newUrl = `${window.location.origin}${window.location.pathname}`;
    window.history.pushState(null, '', newUrl)
    // New code to hide the create file and folder button start
    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }
    //End 
    // end
    // New Code Start
    const DMSShareWithOtherMaster = await sp.web.lists
      .getByTitle("DMSShareWithOtherMaster")
      .items.select("FileName", "FileUID", "FileVersion", "FileSize", "CurrentUser", "DocumentLibraryName", "CurrentFolderPath", "ShareWithMe", "Status", "SiteID", "SiteName", "FilePreviewURL", "ShareAt", "UserID", "PermissionType")
      .orderBy("Created", false)
      ();
    const filteredFiles = DMSShareWithOtherMaster.filter(file => file.ShareWithMe === currentUserEmailRef.current);

    const uniqueItems = filteredFiles.filter((item, index, self) =>
      index === self.findIndex((i) => i.FileUID === item.FileUID)
    );
    console.log("uniqueItems", uniqueItems);
    // console.log("filteredFileData",filteredFiles);



    const container = document.getElementById("files-container");
    container.innerHTML = "";

    console.log("Files Share with me", filteredFiles);
    routeToDiffSideBar = "shareWithMe";
    let filteredFileData = [];
    if (searchText !== null) {
      filteredFileData = uniqueItems.filter((file: any) => file?.FileName?.toLowerCase().includes(searchText?.value?.toLowerCase()))
      if (filteredFileData.length === 0 && searchText !== null) {
        console.log("combineArray", filteredFileData);
        fileNotFound(`No file match ${searchText.value}`);
      }
    } else {
      filteredFileData = uniqueItems;
    }
    filteredFileData.forEach((file) => {
      const { fileIcon, fileExtension } = getFileIcon(file.FileName);
      // console.log("file-Details",file);
      const card = document.createElement("div");
      card.className = "card";
      card.dataset.fileId = file.FileUID;
      card.dataset.listId = file.SiteID;


      card.innerHTML = `   
  <div class="row">
          <div class="col-md-2 pe-0">     
    <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
    </div>
    <div class="col-md-10 pe-0">
    <p class="p1st">${file.FileName}</p>
    <div class="fileSizeAndVersion">
    <p class="p3rd">${file.FileSize} MB</p>
    </div>  </div>  </div>
    <div id="three-dots" class="three-dots" onclick="shareWithMePopUp('${file.FileUID}','${file.SiteID}','${file.CurrentFolderPath}','${file.FileName}')">
    <span>...</span>
    </div>
  `;
      const menu = document.createElement("div");
      menu.id = `menu-${file.FileUID}`;
      menu.className = "popup-menu";
      // menu.innerHTML = `
      //   <ul>
      //     <li onclick="PreviewFile('${file.CurrentFolderPath}/${file.FileName}', '${file.SiteID}','${file.DocumentLibraryName}')">
      //       <img src=${ShareFile} alt="Share"/> File Preview
      //     </li>
      //     <li onclick="shareFile('${file.FileUID}', '${file.SiteID}','${file.CurrentFolderPath}','${file.FileName}','ShareWithMe','${file.FileVersion}','${file.FileSize}','${file.Status}','${file.FilePreviewURL}','${file.DocumentLibraryName}')">
      //       <img src=${ShareFile} alt="Share"/> Share
      //     </li>
      //     <li onclick="DownloadFile('${file.FileUID}', '${file.SiteID}')">
      //       <img src=${ShareFile} alt="Share"/> Download File                
      //     </li>
      //   </ul>
      // `;
      // menu.innerHTML = `
      //   <ul>
      //     <li onclick="PreviewFile('${file.CurrentFolderPath}/${file.FileName}', '${file.SiteID}','${file.DocumentLibraryName}')">
      //       <img src=${ShareFile} alt="Share"/> File Preview
      //     </li>

      //     <li onclick="DownloadFile('${file.FileUID}', '${file.SiteID}')">
      //       <img src=${ShareFile} alt="Share"/> Download File                
      //     </li>
      //   </ul>
      // `;
      menu.innerHTML = `
  <ul>
    <li onclick="PreviewFile('${file.CurrentFolderPath}' , '${file.FileName}', '${file.SiteID}','${file.DocumentLibraryName}','${file.FilePreviewURL}')">
      <img src=${ShareFile} alt="Share"/> File Preview
    </li>

    <li onclick="Download('${file.FileUID}', '${file.SiteID}')">
      <img src=${downloadicon} alt="Share"/> Download File                
    </li>
  </ul>
`;
      card.appendChild(menu);
      container.appendChild(card);
    })
    // End




  }

  //@ts-ignore

  window.shareWithMePopUp = async function (fileId: string, siteID: any, FolderPath: any, FileName: any) {
    console.log("Inside the shareWithMePopUp");
    console.log(siteID, "siteID")
    console.log(fileId, "fileId")
    console.log(FolderPath, "folderPath")
    console.log(FileName, "fileName")

    // check user permission on item start
    const testidsub = await sp.site.openWebById(siteID);
    let filePermission: string;
    let filePath = `${FolderPath}/${FileName}`;
    console.log("filePath", filePath);
    const fileServerRelativePath = testidsub.web.getFileByServerRelativePath(filePath);
    // Retrieve the list item associated with the file
    const item = await fileServerRelativePath.getItem();
    console.log("items", item);
    // Get current user permissions on the item (file)
    const filePermissions = await item.getCurrentUserEffectivePermissions();
    console.log("File permissions:", filePermissions);

    const hasFullControl = testidsub.web.hasPermissions(filePermissions, PermissionKind.ManageWeb);
    const hasEdit = testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
    const hasContribute = testidsub.web.hasPermissions(filePermissions, PermissionKind.AddListItems) && testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
    const hasRead = testidsub.web.hasPermissions(filePermissions, PermissionKind.ViewListItems);
    console.log(hasFullControl, "hasFullControl")
    console.log(hasEdit, "hasEdit")
    console.log(hasContribute, "hasContribute")
    console.log(hasRead, "hasRead")

    if (hasFullControl) {
      filePermission = "Full Control";
    } else if (hasEdit) {
      filePermission = "Edit";
    } else if (hasContribute) {
      filePermission = "Contribute";
    } else if (hasRead) {
      filePermission = "Read";
    } else {
      filePermission = "No Access";
    }

    console.log("filePermission", filePermission);
    // check user permission on item End

    // console.log("enter here i n menu card")
    const allMenus = document.querySelectorAll('.popup-menu');
    console.log(allMenus, "allMenus")
    allMenus.forEach(menu => {
      // console.log(menu , "menu")
      // console.log(menu.id , "menu.id")
      // console.log(fileId , "fileId")
      if (menu.id !== `menu-${fileId}`) {
        menu.classList.remove("show");
      }
    });

    // Toggle the menu for the clicked card
    const menu = document.getElementById(`menu-${fileId}`);
    if (menu) {
      const menu = document.getElementById(`menu-${fileId}`);
      if (!menu) return;
      // if (filePermission === "Edit" || filePermission === "Contribute" || filePermission === "Read") {
      // }
      const secondItem = menu.children[0]?.children[1] as HTMLElement;
      // const secondItem3 = menu.children[0]?.children[2] as HTMLElement;
      if (filePermission === "Read" && secondItem && secondItem.style.display !== "none") {
        secondItem.style.display = "none";
      }
      // if (filePermission === "Read" && secondItem3 && secondItem3.style.display !== "none") {
      //     secondItem3.style.display = "none";
      // }
      menu.classList.toggle("show");
    }


    document.addEventListener('click', (event) => {

      // console.log("Outside click Event Called");

      const target = event.target as HTMLElement;

      // Check if the click was inside any menu or three-dot icon
      const isClickInsideMenu = target.closest('.popup-menu');
      const isClickInsideThreeDots = target.closest('.three-dots');

      // console.log("This is nested folder",isClickInsideThreeDots);

      if (!isClickInsideMenu && !isClickInsideThreeDots) {
        const allMenus = document.querySelectorAll('.popup-menu');
        allMenus.forEach(menu => {
          menu.classList.remove('show');
        });
      }
    });
  }




  const Recyclebin = async (event: React.MouseEvent<HTMLButtonElement> = null, siteIdToUpdate: string = null, searchText: any = null) => {
    setdisplayuploadfileandcreatefolder(false)
    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }

    // Hide the list and grid view start
    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2")
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'none'

    }
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'

    }
    // End

    const container = document.getElementById("files-container");
    if (siteIdToUpdate === null) {
      container.innerHTML = "";
    }

    const FilesItems = await sp.web.lists
      .getByTitle("MasterSiteURL")
      .items.select("Title", "SiteID", "FileMasterList", "Active")
      .filter(`Active eq 'Yes'`)();

    FilesItems.forEach(async (fileItem) => {
      if (fileItem.FileMasterList !== null) {

        if (siteIdToUpdate && fileItem.SiteID !== siteIdToUpdate) {
          return;
        }

        console.log("fileItem.FileMasterList", fileItem.FileMasterList);
        const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "Status", "SiteID", "CurrentFolderPath", "DocumentLibraryName", "SiteName", "FilePreviewURL", "IsDeleted")
          .filter(
            `CurrentUser eq '${currentUserEmailRef.current}'`
          )();

        const listElements = document.querySelectorAll(
          `[data-list-id='${fileItem.SiteID}']`
        );
        console.log("ListElemet To update", listElements)
        listElements.forEach((el) => el.remove());

        console.log("files", filesData);

        routeToDiffSideBar = "recyclebin";
        let filteredFileData;
        if (searchText !== null) {
          filteredFileData = filesData.filter((file: any) => file.FileName.toLowerCase().includes(searchText.value.toLowerCase()))
        } else {
          filteredFileData = filesData;
        }

        filteredFileData.forEach((file) => {
          if (file.IsDeleted !== null) {
            const card = document.createElement("div");
            let fileIcon;
            const fileExtension = file.FileName?.split(".").pop().toLowerCase();
            switch (fileExtension) {
              case "doc":
              case "docx":
                fileIcon = Docicon;
                break;
              case "txt":
                fileIcon = Txticon;
                break;
              case "pdf":
                fileIcon = Pdficon;
                break;
              case "xls":
              case "xlsx":
                fileIcon = Xlsicon;
                break;
              case "zip":
                fileIcon = Zipicon;
                break;
              default:
                fileIcon = Docicon;
                break;
            }

            card.className = "card";
            card.dataset.listId = file.SiteID;
            card.innerHTML = `  
        <div class="row">
          <div class="col-md-2 pe-0"> 
          <div class="IMGContainer"> 
          <div class="CardTextContainer">       
          <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
         </div></div></div>
         <div class="col-md-10 pe-0">
          <p class="p1st">${file.FileName}</p>
          <p class="p2nd"></p>
          <p class="p3rd">${file.FileSize}</p>
          </div></div>
          <div class="three-dots" onclick="toggleMenu2('${file.FileUID}','${fileItem.SiteID}','${file.ID}' , '${fileItem.FileMasterList}')  ">
              <span>...</span>
          </div>
        `;

            const menu = document.createElement("div");
            menu.id = `menu-${file.FileUID}`;
            menu.className = "popup-menu";
            // const showaudit = <FontAwesomeIcon style={{color: "black"}} icon={faListSquares}/>
            menu.innerHTML = `
         <ul>
          <li onclick="confirmUndo('${file.FileUID}','${file.SiteID}','${fileItem.FileMasterList}','${file.DocumentLibraryName}','${file.ID}','${file.CurrentFolderPath}','${file.FileName}')">
            <img src=${Undo} alt="undo"/>Undo
          </li>
        </ul>
        `;
            card.appendChild(menu);
            container.appendChild(card);
          }
        })

      }
    })


  }

  window.undo = async (fileId: any, siteId: any, FileMasterList: any, documentLibraryName: any, ID: any, folderPath: any, fileName: any) => {
    console.log("Undo function called");
    console.log("fileId", fileId);
    console.log("siteId", siteId);
    console.log("FileMasterList", FileMasterList);
    console.log("documentLibraryName", documentLibraryName);
    console.log("ID", ID);
    console.log("folderPath", folderPath);
    console.log("fileName", fileName);

    try {
      // update the corresponding IsDeleted Column of DMSEntityFileMaster to which file belong to.
      // const Id=Number(ID)
      // await sp.web.lists.getByTitle(FileMasterList).items.getById(Id).update({
      //   IsDeleted:null,
      // });
      // console.log(`Item ID ${ID} updated successfully in list "${FileMasterList}".`);
      const currentFile = await sp.web.lists
        .getByTitle(FileMasterList).items.filter(`FileUID eq '${fileId}'`)();

      currentFile.forEach(async (file) => {
        await sp.web.lists.getByTitle(FileMasterList).items.getById(file.Id).update({
          IsDeleted: null
        });
      })

      // update the correponding IsDeleted Column of the document libray to which file belong to.
      try {
        const subsiteContext = await sp.site.openWebById(siteId);
        const fileItem = await subsiteContext.web.getFileByServerRelativePath(`${folderPath}/${fileName}`).getItem();
        let payload: any = {
          IsDeleted: null
        }
        const itemData = await fileItem.update(payload)
        console.log("column updated successfully", itemData);
        await Recyclebin(null, siteId);
      } catch (error) {
        console.log(`Error in updating the columns of document library ${documentLibraryName}`, error);
      }

    } catch (error) {
      console.log(`Error in Updating the Columns of ${FileMasterList}`, error);
    }

  }
  window.confirmUndo = (fileId: any, siteId: any, FileMasterList: any, documentLibraryName: any, ID: any, folderPath: any, fileName: any) => {
    // Create a container for the popup
    const popupContainer = document.createElement("div");
    popupContainer.id = "dynamicPopup";
    popupContainer.style.position = "fixed";
    popupContainer.style.top = "0";
    popupContainer.style.left = "0";
    popupContainer.style.width = "100%";
    popupContainer.style.height = "100%";
    popupContainer.style.backgroundColor = "rgba(0, 0, 0, 0.5)";
    popupContainer.style.display = "flex";
    popupContainer.style.justifyContent = "center";
    popupContainer.style.alignItems = "center";
    popupContainer.style.zIndex = "1000";

    // Create popup content
    const popupContent = document.createElement("div");
    popupContent.style.backgroundColor = "white";
    popupContent.style.padding = "20px";
    popupContent.style.borderRadius = "8px";
    popupContent.style.textAlign = "center";
    popupContent.style.width = "300px";

    // Add message text to the popup
    const message = document.createElement("p");
    message.textContent = "Are you sure you want to undo?";
    popupContent.appendChild(message);

    // Create Yes button
    const yesButton = document.createElement("button");
    yesButton.textContent = "Yes";
    yesButton.style.marginRight = "10px";
    yesButton.onclick = function () {
      window.undo(fileId, siteId, FileMasterList, documentLibraryName, ID, folderPath, fileName);
      closePopup(); // Close the popup after confirming
    };
    popupContent.appendChild(yesButton);

    // Create No button
    const noButton = document.createElement("button");
    noButton.textContent = "No";
    noButton.onclick = closePopup; // Just close the popup if canceled
    popupContent.appendChild(noButton);

    // Append popup content to popup container
    popupContainer.appendChild(popupContent);

    // Append the popup container to the document body
    document.body.appendChild(popupContainer);

    // Function to remove the popup from the DOM
    function closePopup() {
      document.body.removeChild(popupContainer);
    }
  }

  window.shareFile = async (fileID: string, siteId: string, currentFolderPathForFile: string, fileName: string, flag: string, FileVersion: any, FileSize: any, Status: any, FilePreviewURL: any, DocumentLibraryName: any) => {
    console.log("Share File called");
    console.log("flag", flag);
    console.log("file Id", fileID);
    console.log("site Id", siteId);
    console.log("FileName", fileName);
    console.log("currentFolderPath", currentFolderPathForFile);

    // Check permission of file when it come from the myrequest start
    const testidsub = await sp.site.openWebById(siteId)

    let filePath = `${currentFolderPathForFile}/${fileName}`;
    console.log("filePath", filePath);
    const fileServerRelativePath = testidsub.web.getFileByServerRelativePath(filePath);
    // Retrieve the list item associated with the file
    const item = await fileServerRelativePath.getItem();
    console.log("items", item);
    // Get current user permissions on the item (file)
    const filePermissions = await item.getCurrentUserEffectivePermissions();
    console.log("File permissions:", filePermissions);
    // console.log("file listItems All field",file.ListItemAllFields);

    const hasFullControl = testidsub.web.hasPermissions(filePermissions, PermissionKind.ManageWeb);
    const hasEdit = testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
    const hasContribute = testidsub.web.hasPermissions(filePermissions, PermissionKind.AddListItems) && testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
    const hasRead = testidsub.web.hasPermissions(filePermissions, PermissionKind.ViewListItems);
    console.log(hasFullControl, "hasFullControl")
    console.log(hasEdit, "hasEdit")
    console.log(hasContribute, "hasContribute")
    console.log(hasRead, "hasRead")
    let filePermission: string;
    if (hasFullControl) {
      filePermission = "Full Control";
    } else if (hasEdit) {
      filePermission = "Edit";
    } else if (hasContribute) {
      filePermission = "Contribute";
    } else if (hasRead) {
      filePermission = "Read";
    } else {
      filePermission = "No Access";
    }

    console.log("filePermission", filePermission);

    // exreact the Entity from folder path
    const parts = currentFolderPathForFile.split("/");
    const entity = parts[3];
    console.log(entity);

    const fetchUser = async (entity: any) => {
      // const [
      //   users,
      //   users1,
      //   users2,
      //   users3,
      //   users4,
      // ] = await Promise.all([
      //   sp.web.siteGroups.getByName(`${entity}_Read`).users(),
      //   sp.web.siteGroups.getByName(`${entity}_Initiator`).users(),
      //   sp.web.siteGroups.getByName(`${entity}_Contribute`).users(),
      //   sp.web.siteGroups.getByName(`${entity}_Admin`).users(),
      //   sp.web.siteGroups.getByName(`${entity}_View`).users(),
      // ]);
      // console.log(users, "users ", users1,users2,users3,users4);
      // const combineArray = [
      //   ...(users || []),
      //   ...(users1 || []),
      //   ...(users2 || []),
      //   ...(users3 || []),
      //   ...(users4 || []),
      // ];

      // const siteContext = await sp.site.openWebById(OthProps.siteID);
      const user0 = await sp.web.siteUsers();
      const combineUsersArray = user0.map((user) => (
        {
          id: String(user.Id),
          value: user.Title,
          email: user.Email,
        }
      ))
      console.log("Sub site users", combineUsersArray);

      // const resultArray=combineUsersArray.map((user) => ( 
      //   {
      //     id:String(user.Id),
      //     value: user.Title,
      //     email: user.Email
      //   }
      // ))
      // console.log("combineArray", combineArray);
      // console.log("resultArray",resultArray)

      return combineUsersArray;
    }

    const users = await fetchUser(entity);
    console.log("UserArray", users);


    // Check if a popup already exists, if so, remove it before creating a new one
    const existingPopup = document.getElementById('share-popup');
    if (existingPopup) {
      existingPopup.remove();
    }

    // Dummy data
    // const users = [
    //   { value: 'Test1', id: '14',email:"User1@officeindia.onmicrosoft.com" },
    //   { value: 'Test2', id: '31',email:"User2@officeindia.onmicrosoft.com" },
    //   { value: 'Test3', id: '137',email:"User3@officeindia.onmicrosoft.com"},
    //   { value: 'Test4', id: '33',email:"User4@officeindia.onmicrosoft.com" },
    //   { value: 'Test5', id: '32',email:"User5@officeindia.onmicrosoft.com" },
    //   { value: 'Test6', id: '34',email:"User6@officeindia.onmicrosoft.com" },
    //   { value: 'Test User1', id: '39',email:"User7@officeindia.onmicrosoft.com" },
    //   ];


    // Declare selectedUsers with an explicit type, assuming user IDs are of type string for selecting the user for share
    let selectedUsers: { id: string; value: string; email: string }[] = [];
    // Create the pop-up element
    const popup = document.createElement("div");
    popup.id = 'share-popup';
    popup.className = "share-popup";

    // Show permissions options.
    let options = ''
    if (filePermission === "Full Control") {
      options = `
  <option value="Full Control">Full Control</option>
  <option value="Contribute">Contribute</option>
  <option value="Edit">Edit</option>
  <option value="Read">Read</option>
`
    } else if (filePermission === "Contribute" || filePermission === "Edit") {
      options = `
<option value="Contribute">Contribute</option>
<option value="Edit">Edit</option>
<option value="Read">Read</option>
`
    } else if (filePermission === "Read") {
      options = `
<option value="Read">Read</option>
`
    }


    // Add HTML structure for the pop-up with a dropdown and a close "X" button
    popup.innerHTML = `
<div class="share-popup-content">
<div class="share-popup-header">
<h4>Share</h4>
<span class="share-close-popup" onClick="hideSharePopUp()">x</span>
</div>
<div class="share-popup-body">
<div id="share-reactSelect">
    <input type="text" id="userInput" placeholder="Add a Name, Group, or Email" style="
    width: 100%; 
    padding: 10px;
    font-size: 14px;
    border-radius: 4px;
    border: 1px solid #ccc;
  "/>
  <div id="userDropdown" class="user-dropdown" style="
    display: none;
    position: absolute;
    width: 100%;
    max-height: 150px;
    overflow-y: auto;
    background-color: white;
    border: 1px solid #ccc;
    border-radius: 4px;
    z-index: 1000;
  ">
  </div>
</div>
 <div>
  <select id="permissionSelect" style="
    margin-bottom:10px;
    width: 100%; 
    padding: 10px;
    font-size: 14px;
    border-radius: 4px;
    border: 1px solid #ccc;
    margin-top: 10px;
  ">
    <option value="" disabled selected>Permission</option>
    ${options}
  </select>
</div>
<textarea id="share-message" placeholder="Write a message..." >
</textarea>
</div>
<div class="share-popup-footer">
<button id="share-shareFileButton">Share</button>
</div>
</div>
`;

    // Append the  popup to the body
    document.body.appendChild(popup);

    // Get references to the input box and dropdown
    const userInput = document.getElementById('userInput') as HTMLInputElement;
    const userDropdown = document.getElementById('userDropdown');

    // Function to render dropdown options based on user input
    function renderDropdown(users: { id: string, value: string, email: string }[]) {
      // Clear previous options
      userDropdown.innerHTML = '';
      users.forEach(user => {
        const option = document.createElement('div');
        option.className = 'dropdown-item';
        option.style.padding = '8px';
        option.style.cursor = 'pointer';
        option.textContent = user.value;
        option.onclick = () => selectUser(user);
        userDropdown.appendChild(option);
      });
    }

    // Function to show the dropdown when the input is clicked
    userInput.addEventListener('focus', () => {
      userDropdown.style.display = 'block';

      // Display all users initially
      renderDropdown(users);
    });
    userInput.addEventListener('paste', (event) => {
      setTimeout(() => {
        const inputValue = userInput.value.toLowerCase();
        const matchedUser = users.find(user => user.email.toLowerCase() === inputValue);

        if (matchedUser) {
          renderDropdown([matchedUser]);
          userDropdown.style.display = 'none';
        }
      }, 0); // Use setTimeout to ensure the paste event updates the input value first
    });
    // Filter dropdown based on input value
    userInput.addEventListener('input', () => {
      const searchValue = userInput.value.toLowerCase();
      const filteredUsers = users.filter(user => user.value.toLowerCase().includes(searchValue));
      renderDropdown(filteredUsers);
    });

    // Function to select a user and display it inside the input
    function selectUser(user: { id: string, value: string, email: string }) {
      console.log("selected user", selectedUsers)
      if (!selectedUsers.some(selectedUser => selectedUser.id === user.id)) {

        selectedUsers.push(user);

        // Create a span for the selected user with a close button
        const selectedUserDiv = document.createElement('span');
        selectedUserDiv.className = 'selected-user';
        selectedUserDiv.style.display = 'inline-block';
        selectedUserDiv.style.padding = '2px 6px';
        selectedUserDiv.style.backgroundColor = '#e0e0e0';
        selectedUserDiv.style.borderRadius = '12px';
        selectedUserDiv.style.marginRight = '5px';
        selectedUserDiv.style.position = 'relative';

        selectedUserDiv.textContent = user.value;

        // Create close button for deselecting the user
        const closeButton = document.createElement('span');
        closeButton.textContent = 'x';
        closeButton.style.cursor = 'pointer';
        closeButton.style.marginLeft = '5px';
        closeButton.onclick = () => deselectUser(user.id, selectedUserDiv);
        selectedUserDiv.appendChild(closeButton);

        // Append the selected user to the input field
        userInput.parentNode!.insertBefore(selectedUserDiv, userInput);
        userInput.value = '';
      }
      userDropdown.style.display = 'none';
    }

    // Function to deselect a user
    function deselectUser(userId: string, selectedUserDiv: HTMLElement) {
      // selectedUsers = selectedUsers.filter(id => id !== userId);
      selectedUsers = selectedUsers.filter(selectedUser => selectedUser.id !== userId);
      console.log("selected user", selectedUsers);
      selectedUserDiv.remove();
    }

    // Hide the dropdown if clicked outside
    document.addEventListener('click', (event) => {
      if (!userInput.contains(event.target as Node) && !userDropdown.contains(event.target as Node)) {
        userDropdown.style.display = 'none';
      }
    });

    // Capture selected permission
    let selectedPermission = "";
    document.getElementById('permissionSelect').addEventListener('change', (event) => {
      selectedPermission = (event.target as HTMLSelectElement).value;
      console.log("Selected Permission:", selectedPermission);
    });

    // Adding event listener to the "Share" button
    document.getElementById('share-shareFileButton').addEventListener('click', async function () {
      console.log("selectedUserArray", selectedUsers);
      console.log("Entity", entity);
      console.log("FileId", fileID);
      console.log("SiteId", siteId);
      console.log("currentFolderPathForFile", currentFolderPathForFile);
      console.log("FileName", fileName);
      console.log("filesize", FileSize);
      console.log("FileVersion", FileVersion);
      console.log("Status", Status);
      console.log("FilePreviewURL", FilePreviewURL);
      console.log("DocumentLibraryName", DocumentLibraryName)

      // New Code push the data into the DMSShareWithOtherMaster Start
      try {
        const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
        const payloadForDMSShareWithOtherMaster = {
          FileName: fileName,
          FileUID: fileID,
          CurrentUser: currentUserEmailRef.current,
          CurrentFolderPath: currentFolderPathForFile,
          SiteName: entity,
          PermissionType: selectedPermission,
          ShareAt: isoDate,
          FileVersion: FileVersion,
          FileSize: FileSize,
          Status: Status,
          FilePreviewURL: FilePreviewURL,
          SiteID: siteId,
          DocumentLibraryName: DocumentLibraryName
        }
        selectedUsers.forEach(async (user) => {
          (payloadForDMSShareWithOtherMaster as any).UserID = user.id;
          (payloadForDMSShareWithOtherMaster as any).ShareWithOthers = user.value;
          (payloadForDMSShareWithOtherMaster as any).ShareWithMe = user.email;
          const newItem = await sp.web.lists.getByTitle(`DMSShareWithOtherMaster`).items.add(payloadForDMSShareWithOtherMaster)
          console.log("Data added successfully in the", newItem);
        })

      } catch (error) {
        console.log("Error in adding data to the DMSShareWithOtherMaster", error);
      }


    });


  }
  // hide the share popup
  // @ts-ignore
  window.hideSharePopUp = () => {
    const popup = document.querySelector('.share-popup');

    if (popup) {
      popup.remove();
    }
  }


  // Sharewith Me And Share With Others
  // Toggle the menu card
  // @ts-ignore
  window.toggleMenu2 = async function (fileId: any, siteID: any, listitemid: any, Listname: any) {




    const myfunction = async () => {
      const subsiteContext = await sp.site.openWebById(siteID);

      // Fetch all the groups in the subsite
      interface IMember {
        PrincipalType: number;
        Title: string;
        Id: number;
      }

      interface IRoleAssignmentInfo {
        Member?: IMember;
      }

      const groups3: IRoleAssignmentInfo[] = await subsiteContext.web.roleAssignments.expand("Member")();
      console.log("groups3", groups3);

      // Filter the groups for current user roles (_View, _Read, _Contribute, _Admin, etc.)
      const filteredMembers = groups3.filter((roleAssignment) => {
        return roleAssignment.Member.PrincipalType === 8; // Group PrincipalType
      });

      const filteredGroups = filteredMembers.map((object) => ({
        value: object.Member.Title,
        label: object.Member.Title,
        Id: object.Member.Id,
      }));
      console.log("filteredGroups", filteredGroups);
      mydatacard = "12"
      // Check if current user is in the _Admin or _Contribute group
      const isAdmin = filteredGroups.some((group) => group.value.includes("_Admin"));
      const isContribute = filteredGroups.some((group) => group.value.includes("_Contribute"));
      if (isAdmin) {
        isadmin = "Admin"
        console.log("User is Admin")
      }
      if (isContribute) {
        isadmin = "Contribute"
        console.log("User is Contribute")
      }
    }
    myfunction()

    console.log("Inside the toggleMenu2");
    console.log(siteID, "siteID")
    console.log(fileId, "fileId")
    console.log("enter here i n menu card")
    const allMenus = document.querySelectorAll('.popup-menu');
    console.log(allMenus, "allMenus")
    allMenus.forEach(menu => {
      console.log(menu, "menu")
      console.log(menu.id, "menu.id")
      console.log(fileId, "fileId")
      if (menu.id !== `menu-${fileId}`) {
        menu.classList.remove("show");
      }
    });

    // Toggle the menu for the clicked card
    const menu = document.getElementById(`menu-${fileId}`);
    if (menu) {
      console.log("Toggle the menu for the clicked card")
      menu.classList.toggle("show");
    }
    document.addEventListener('click', (event) => {

      // console.log("Outside click Event Called");

      const target = event.target as HTMLElement;

      // Check if the click was inside any menu or three-dot icon
      const isClickInsideMenu = target.closest('.popup-menu');
      const isClickInsideThreeDots = target.closest('.three-dots');

      // console.log("This is nested folder",isClickInsideThreeDots);

      if (!isClickInsideMenu && !isClickInsideThreeDots) {
        const allMenus = document.querySelectorAll('.popup-menu');
        allMenus.forEach(menu => {
          menu.classList.remove('show');
        });
      }
    });
  }


  //   Edit file action
  //    @ts-ignore



  //   Delete file action
  //  @ts-ignore

  //  / this is pop up function
  window.confirmDeleteFile = async (fileId: string, siteID: string, IsHardDelete: any, ListToUpdate: any = null) => {
    console.log("list name is ", ListToUpdate)
    // console.log(listToUpdate , "listAnme")
    event.preventDefault();
    event.stopPropagation();

    console.log("Inside The confirmDeleteFile");
    console.log("Is Hard Delete ", IsHardDelete);
    console.log("FileID", fileId);
    console.log("siteId", siteID);

    const popupData = await sp.web.lists.getByTitle('DMSPopupMaster').items
      .select('PopupText', 'Typeofpopup', 'Isrequires')();

    console.log("popupitems", popupData);

    popupData.forEach(async (popItems) => {

      // Check the type of popup and if it is required
      switch (popItems.Typeofpopup) {

        case 'Delete':
          if (popItems.Isrequires) {
            console.log(popItems.Typeofpopup, "popItems.Typeofpopup ")
            console.log("TypeOfPopUp: Delete and Isrequires is true");

            // Create Pop
            const deleteConfirmationPop = document.createElement('div');
            deleteConfirmationPop.className = "popup-modal";
            deleteConfirmationPop.innerHTML = `
                        <div class="popup-content">
                          <p id="confirmation-text">${popItems.PopupText}</p>
                          <div class="popup-actions">
                              <button id="confirm-yes">Yes</button>
                              <button id="confirm-no">No</button>
                          </div>
                        </div>
                    `;

            document.body.appendChild(deleteConfirmationPop);

            // Handle Yes button click (confirmation for Delete)
            const yesButton = document.getElementById('confirm-yes');
            yesButton.addEventListener('click', async () => {
              const confirmationText = document.getElementById('confirmation-text');
              confirmationText.innerHTML = 'Loading...';

              try {
                console.log("Calling deleteFile from confirm delete");
                await window.deleteFile(fileId, siteID, IsHardDelete, ListToUpdate);

                // console.log("Updating List inside the confirem Delete");
                // if(ListToUpdate){       
                //         const items999 = await sp.web.lists
                //         .getByTitle(ListToUpdate).items.filter(`FileUID eq '${fileId}'`).top(1)();
                //         alert(items999)

                //         if (items999.length > 0) {
                //         const itemId = items999[0].ID;
                //         console.log(itemId , "itemId")
                //         console.log(items999 , "item9999")
                //         // Delete the item by ID
                //         const mylist = ListToUpdate
                //         console.log(mylist, "mylist")
                //         await sp.web.lists.getByTitle(mylist).items.getById(itemId).delete();

                //         console.log(`Item with FileUid ${fileId} has been deleted.`);
                //         }

                // }
                confirmationText.innerHTML = 'Your file was deleted successfully.';

              } catch (error) {
                confirmationText.innerHTML = 'Something went wrong. Your file was not deleted.';
              }

              // Remove the popup after 1 second
              setTimeout(() => document.body.removeChild(deleteConfirmationPop), 1000);
            });

            // Handle No button click (cancel deletion)
            document.getElementById('confirm-no').addEventListener('click', () => {
              document.body.removeChild(deleteConfirmationPop); // Close the popup
            });

          } else {
            console.log("TypeOfPopUp: Delete and Isrequires is false");
            // Directly delete the file if no popup is required
            try {
              await window.deleteFile(fileId, siteID, IsHardDelete, ListToUpdate);
              // alert('Your file was deleted successfully.');
            } catch (error) {
              // alert('Error deleting file.');
            }
          }
          break;

        case 'CreateFile':
          if (popItems.Isrequires) {
            console.log("TypeOfPopUp: CreateFile and Isrequires is true");

            // Show popup for CreateFile
            const createFileConfirmationPop = document.createElement('div');
            createFileConfirmationPop.className = "popup-modal";
            createFileConfirmationPop.innerHTML = `
                        <div class="popup-content">
                          <p id="confirmation-text">${popItems.PopupText}</p>
                          <div class="popup-actions">
                              <button id="confirm-yes">Yes</button>
                              <button id="confirm-no">No</button>
                          </div>
                        </div>
                    `;

            document.body.appendChild(createFileConfirmationPop);
          } else {
            // Logic without Pop
          }
          break;

        // Add more cases here for other types like 'Edit', 'Upload', etc.
        default:
          console.log("Unknown TypeOfpopup: ", popItems.Typeofpopup);
      }
    });

  }



  window.deleteFile = async (fileId: string, siteID: string, IsHardDelete: any, ListToUpdate: any = null) => {
    console.log("Inside the deleteFile");
    console.log("ListToUpdate", ListToUpdate)
    console.log(siteID, "siteID")
    console.log(`Delete file with ID: ${fileId}`);
    console.log("ISHard delete inside delete ", IsHardDelete);
    console.log("ISHard delete type of ", typeof (IsHardDelete));


    const { web } = await sp.site.openWebById(siteID)
    const file = await web.getFileById(fileId);
    const listItem = await file.getItem();

    const isoDate = new Date().toISOString().slice(0, 19) + 'Z';



    if (IsHardDelete === "true") {
      // alert( `in true IsHardDelete is ${IsHardDelete}`)
      try {
        const deleteffile = await web.getFileById(fileId).delete();
        console.log(deleteffile, "deleteffile");
      } catch (error) {
        console.log(error, "in true IsHardDelete is")
      }

    } else if (IsHardDelete === "false") {
      // alert( `in false IsHardDelete is ${IsHardDelete}`)
      try {
        const updatedData = await listItem.update({
          IsDeleted: isoDate
        });
        console.log("Updated data", updatedData);
      } catch (error) {
        console.log(error, "in  IsHardDelete is")
      }

    }

    //  alert(`File with ID: ${fileId} has been deleted successfully.`);
    console.log(currentfolderpath, "currentfolderpath")
    console.log("currentEntity", currentEntity);

    //start
    if (ListToUpdate || currentEntity) {
      console.log("Inside The check Of Entity->", currentEntity, "->", ListToUpdate);
      let currentList: any;
      if (ListToUpdate) {
        currentList = ListToUpdate;
      }
      if (currentEntity) {
        currentList = `DMS${currentEntity}FileMaster`;
      }
      console.log("selected List", currentList);
      // New Code Start create or update the DMSEntityFileMaster for delete operation.
      const currentFileDataRelatedToUser = await sp.web.lists
        .getByTitle(currentList).items.filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}'`)();
      console.log("currentFileDataRelatedToUser", currentFileDataRelatedToUser);
      if (currentFileDataRelatedToUser.length > 0) {
        console.log("Current file data present corresponding to the current user");
        const currentFileData = await sp.web.lists
          .getByTitle(currentList).items.filter(`FileUID eq '${fileId}'`)();
        currentFileData.forEach(async (file) => {
          // console.log("file",file);
          await sp.web.lists.getByTitle(currentList).items.getById(file.Id).update({
            IsDeleted: isoDate
          });
        })
      } else {
        console.log("Current file data does not present corresponding to the current user");
        const payload = {
          FileName: "",
          FileUID: fileId,
          FileVersion: "",
          FileSize: "",
          CurrentUser: currentUserEmailRef.current,
          CurrentFolderPath: currentfolderpath,
          DocumentLibraryName: currentDocumentLibrary,
          FolderName: currentFolder,
          SiteName: currentEntity,
          SiteID: siteID,
          Status: "",
          FilePreviewURL: ""
        }
        const file: any = await web.getFileById(fileId).select("Name", "Length", "ServerRelativeUrl", "UniqueId", "MajorVersion", "ListItemAllFields/Status", "ListItemAllFields/IsDeleted").expand('ListItemAllFields')();
        console.log("file", file);

        // Encode the file name and construct the preview URL
        const encodedFilePath = encodeURIComponent(file.ServerRelativeUrl);

        const parentFolder = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf('/'));
        const siteUrl = window.location.origin;
        // const previewUrl = `${siteUrl}/sites/AlRostmani/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
        //  const previewUrl = `${siteUrl}/sites/AlRostmanispfx2/${currentEntity}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
        const previewUrl = `${siteUrl}${locationPath}/${currentEntity}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
        console.log("previewUrl", previewUrl);
        payload.FilePreviewURL = previewUrl

        payload.FileName = file.Name;
        payload.FileSize = ((file.Length as unknown as number) / (1024 * 1024)).toFixed(2);
        payload.FileVersion = String(file.MajorVersion)
        payload.Status = file.ListItemAllFields.Status

        const addData = await sp.web.lists.getByTitle(currentList).items.add(payload);
        console.log("addData", addData);


        const currentFile = await sp.web.lists
          .getByTitle(currentList).items.filter(`FileUID eq '${fileId}'`)();

        currentFile.forEach(async (file) => {
          await sp.web.lists.getByTitle(currentList).items.getById(file.Id).update({
            IsDeleted: isoDate
          });
        })
      }
      // End


    }
    // end
    console.log("currentfolderpath", currentfolderpath, "currentsiteID", currentsiteID, "currentDocumentLibrary", currentDocumentLibrary)
    getdoclibdata(currentfolderpath, currentsiteID, currentDocumentLibrary)
    //  getfolderdata(currentfolderpath,currentsiteID)
  };


  //Manage Folder Permission Action 
  // window.managePermissionSearch=(message:string)=>{
  //   console.log(message);
  // }

  // Manage Folder WorkFlow Action

  // Manage Folder View Action
  window.view = (message: string) => {
    console.log(message);
  }



  // My ctreated folder 
  const createFileButton2 = document.getElementById('createFileButton2')
  const createFileButton = document.getElementById('createFileButton')

  const mycreatedfolders = async (event: any = null, searchText: any = null) => {
    setdisplayuploadfileandcreatefolder(false)
    const wait = document.getElementById('files-container')
    wait.classList?.remove('hidemydatacards')
    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }
    setlistorgriddata('')
    setlistorgriddata('')
    setShowMyrequButtons(false)
    setShowMyfavButtons(false)

    if (event) {
      event.preventDefault()
      event.stopPropagation()
    }
    // clean the url start
    const newUrl = `${window.location.origin}${window.location.pathname}`;
    window.history.pushState(null, '', newUrl)
    // end

    // start
    // call this function onClick of the myFolder Button
    // handleShowContent(event)
    // end
    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2");
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'

    }
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................");
      hidegidvewlistviewbutton2.style.display = 'none';
    }

    // Check if the user is super Admin start
    let superAdmin = false;
    let folderItems: any[] = []
    const currentUser = await sp.web.currentUser();
    const userGroups = await sp.web.siteUsers.getById(currentUser.Id).groups();
    const isMemberOfSuperAdmin = userGroups.some(group => group.Title === `DMSSuper_Admin`);
    if (isMemberOfSuperAdmin) {
      // alert("Current user is super admin")
      superAdmin = true;
      folderItems = await sp.web.lists
        .getByTitle("DMSFolderMaster")
        .items.select("CurrentUser", "IsFolder", "FolderPath", "DocumentLibraryName", "SiteTitle", "ID", "IsPrivate", "IsLibrary", "FolderName", "IsRename")
        .orderBy("Created", false).getAll();
    } else {
      // alert("Current user not a super admin")
      folderItems = await sp.web.lists
        .getByTitle("DMSFolderMaster")
        .items.select("CurrentUser", "IsFolder", "FolderPath", "DocumentLibraryName", "SiteTitle", "ID", "IsPrivate", "IsLibrary", "FolderName", "IsRename")
        .filter(`CurrentUser eq '${currentUserEmailRef.current}'`).orderBy("Created", false).getAll();
    }
    // end

    // const folderItems = await sp.web.lists
    // .getByTitle("DMSFolderMaster")
    // .items.select("CurrentUser" , "IsFolder" , "FolderPath" , "DocumentLibraryName","SiteTitle","ID" , "IsPrivate","IsLibrary","FolderName")
    // .filter(`CurrentUser eq '${currentUserEmailRef.current}'`).orderBy("Created", false)();
    // console.log(folderItems , "folderItems");

    // new code to fetch the siteId from the masterSiteURl and map this siteid with corresponding siteTitle forEach folder in the folderData that fetch from the DMSFolderMaster
    const dataFromMasterSiteURL = await sp.web.lists.getByTitle("MasterSiteURL").items.select("Title", "SiteID").filter(`Active eq 'Yes'`)();
    console.log("dataFromMasterSiteURL", dataFromMasterSiteURL);

    // check the folder permission start

    // if(folderItems[0].IsLibrary){
    //   console.log(`folder path - ${folderItems[0].FolderPath}`)
    //   try {
    //     const library = await sp.web.getList(`${folderItems[0].FolderPath}`).roleAssignments.expand("Member", "RoleDefinitionBindings")();

    //     library.forEach((assignment:any) => {
    //         console.log("Assigned to:", assignment.Member.Title);
    //         console.log(
    //             "Roles:",
    //             assignment.RoleDefinitionBindings.map((role:any) => role.Name).join(", ")
    //         );
    //     });
    // } catch (error) {
    //     console.error("Error fetching library permissions:", error);
    // }
    // }


    // end

    const siteMap = new Map();
    dataFromMasterSiteURL.forEach(site => {
      siteMap.set(site.Title, site.SiteID);
    });

    const folderDataWithSiteId = folderItems.map(folder => {
      // Get the SiteID or null if not found
      const siteID = siteMap.get(folder.SiteTitle) || null;
      return {
        ...folder,
        // Append SiteID to the folder object
        SiteID: siteID
      };
    });

    console.log("Resultant folder data", folderDataWithSiteId);
    // end new code

    const container = document.getElementById("files-container");
    container.innerHTML = "";
    const folderimg = require('../../assets/Folder.png')

    // start
    console.log("searchInput", searchText);
    routeToDiffSideBar = "myFolder";
    let filteredFileData;
    if (searchText !== null) {
      // here we change the array to new siteId containing array
      filteredFileData = folderDataWithSiteId.filter((folder: any) =>
        folder.DocumentLibraryName?.toLowerCase().includes(searchText.value.toLowerCase())
        || folder.FolderName?.toLowerCase().includes(searchText.value.toLowerCase())
        // ||   folder.ParentFolder.toLowerCase().includes(searchText.value.toLowerCase())
      )

      if (filteredFileData.length === 0 && searchText !== null) {
        console.log("combineArray", filteredFileData);
        fileNotFound(`No folder match ${searchText.value}`);
      }
    } else {
      // here we change the array to new siteId containing array
      filteredFileData = folderDataWithSiteId;
      console.log("filteredFileData", filteredFileData)
    }
    // end
    // change the array name in the for loop
    for (const files of filteredFileData) {
      // console.log("FolderName",files.FolderName);
      let folderName = '';
      if (files.IsLibrary === true) {
        folderName = files.DocumentLibraryName;
      } else if (files.IsFolder === true) {
        folderName = files.FolderName;
      }
      if (files.IsRename !== null) {
        folderName = files.IsRename;
      }
      let folderisprivateorpublic: any = ""
      if (files.IsPrivate === true) {
        folderisprivateorpublic = "Private"
      } else if (files.IsPrivate === false) {
        folderisprivateorpublic = "Public"
      } else if (files.IsPrivate === null) {
        folderisprivateorpublic = "Null"
      }
      // console.log("files111",files);
      const card = document.createElement("div");

      card.className = "card";
      card.innerHTML = `
   <div class="row"> 
    <div class="col-md-2 pe-0">   
    <div class="IMGContainer">  
     <div class="CardTextContainer">
  <img class="filextension" src=${folderimg} icon"/>
  </div></div></div>
  <div class="col-md-10"> 
  <p class="p1st p1stfolder">${folderName}</p>
  <p class="p3rd">${files.SiteTitle} </p>
  <div class="mycreatedfolderpublicorlibrary"> <p class="filestatus">${folderisprivateorpublic} </p> 
  
  <p class="filestatus2 ${files.IsLibrary === true ? 'root-folder' : 'sub-folder'}"> ${files.IsLibrary === true ? 'Root Folder' : 'Sub Folder'} </p> </div>
  </div>
  <div class="three-dots folderthreedots" onclick="toggleMenu2('${files.ID}','${files.SiteID}')">
      <span>...</span>
  </div> </div>
           </div>
`;
      const menu = document.createElement("div");
      menu.id = `menu-${files.ID}`;
      menu.className = "popup-menu";
      menu.innerHTML = `
<ul>
     <li onclick="managePermissionSearch('${files.DocumentLibraryName}','${files.SiteTitle}','${files.SiteID}','${files.FolderName}','${files.FolderPath}')">
      <img src=${managePermissionIcon} alt="managePermissionSearch"/>
      Manage Permission
  </li>
  <li onclick="manageWorkflow('${files.DocumentLibraryName}','${files.SiteTitle}','${files.SiteID}')">
    <img src=${manageWorkFlowIcon} alt="ManageWorkFlow"/>
    Manage Workflow
  </li>
  <li onclick="editFile('${files.SiteTitle}','${files.DocumentLibraryName}')">
    <img src=${editIcon} alt="Edit"/>
    Edit
  </li>
  ${superAdmin === true ?
          `<li onclick="deleteFolder('${files.SiteTitle}','${folderName}','${files.ID}','${files.SiteID}','${files.FolderPath}','${files.IsLibrary}')">
    <img src=${deleteIcon} alt="Edit"/>
    Delete
    </li>
    <li onclick="renameFolder('${files.SiteTitle}','${folderName}','${files.ID}','${files.SiteID}')">
      <img src=${editIcon} alt="Edit"/>
      Rename
    </li>
    ${files.IsLibrary === true ? `<li onclick="renameColumn('${files.SiteTitle}','${files.DocumentLibraryName}')">
      <img src=${editIcon} alt="Edit"/>
      Rename Column
    </li>`: ''}
    `
          : ''}
</ul>
`;


      card.appendChild(menu);
      const isfolderordoclib = card.querySelector(".filestatus2") as HTMLElement;
      switch (files.IsLibrary === true) {
        case true:
          isfolderordoclib.style.backgroundColor = "gray";
          isfolderordoclib.style.color = "white";
      }
      const fileStatusElement = card.querySelector(".filestatus") as HTMLElement;
      switch (files.IsPrivate) {
        case false:
          fileStatusElement.style.backgroundColor = "#b5e7d3";
          fileStatusElement.style.color = "#008751";
          break;
        case true:
          fileStatusElement.style.backgroundColor = "rgba(241, 85, 108, 0.1)";
          fileStatusElement.style.color = "#f1556c";
          break;
        case null:
          fileStatusElement.style.backgroundColor = "gray";
          fileStatusElement.style.color = "white";
          break;
        default:
          fileStatusElement.style.backgroundColor = "gray";
          fileStatusElement.style.color = "white";
          break;
      }

      container.appendChild(card);
      const menu1 = document.getElementById(`menu-${files.ID}`);
      if (files.IsFolder === true) {
        const secondItem = menu1.children[0]?.children[1] as HTMLElement;
        const thirdItem = menu1.children[0]?.children[2] as HTMLElement;
        if (secondItem && secondItem.style.display !== "none") {
          secondItem.style.display = "none";
        }
        if (thirdItem && thirdItem.style.display !== "none") {
          thirdItem.style.display = "none";
        }
      }
    }

  }
  // Function to delete the folder
  // @ts-ignore
  // window.deleteFolder=(siteName:any,documentLibraryName:any)=>{
  //   console.log("siteName",siteName)
  //   console.log("documentLibraryName",documentLibraryName)
  // }

  // Function to delete the folder
  // @ts-ignore
  window.deleteFolder = (siteName: any, folderName: any, itemId: any, siteId: any, folderpath: any, IsLibrary: any) => {
    console.log("siteName", siteName)
    console.log("folderName", folderName)
    console.log("itemId", itemId)
    console.log("siteId", siteId)
    console.log("folderpath", folderpath)
    console.log("IsLibrary", IsLibrary)
    Swal.fire({
      title: "Are you sure you want to delete this folder?",
      text: "Deleting this folder will also permanently delete all files and subfolders inside it.",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, delete it!"
    }).then(async (result) => {
      if (result.isConfirmed) {
        // Delete the Library and folder in sharepoint
        try {

          const { web } = await sp.site.openWebById(siteId);
          if (IsLibrary === 'true') {
            const data = await web.lists.getByTitle(folderName).delete();
            console.log("Library deleted succesfully", data);
          } else {
            const data = await web.getFolderByServerRelativePath(`${folderpath}`).delete();
            console.log("Folder deleted succesfully", data);
          }

          // Delete the data related to the folder/library inside the DMSFolderMster
          try {
            await sp.web.lists.getByTitle(`DMSFolderMaster`).items.getById(itemId).delete();
            console.log(`Item with ID ${itemId} deleted successfully from list DMSFolderMaster.`);
          } catch (error) {
            console.error(`Error deleting item: ${error.message}`);
          }

          try {
            // Query the list for items where the 'FolderPath' column matches the given value
            const items = await sp.web.lists.getByTitle(`DMS${siteName}FileMaster`)
              .items.filter(`CurrentFolderPath eq '${folderpath}'`)
              .select("Id")();

            // Check if any items match the query
            if (items.length > 0) {
              // Delete all matching items
              for (const item of items) {
                await sp.web.lists.getByTitle(`DMS${siteName}FileMaster`).items.getById(item.Id).delete();
                console.log(`Deleted item with ID: ${item.Id}`);
              }

              console.log("All matching items deleted successfully.");
            }
          } catch (error) {
            console.error(`Error deleting items in DMS${siteName}FileMaster: ${error.message}`);
          }
        } catch (error) {
          console.log("Error in deleteing Folder ", error);
        }


        Swal.fire({
          title: "Deleted!",
          text: "Your folder has been deleted.",
          icon: "success"
        });
        mycreatedfolders(null, null);
      }
    });

  }

  // Function to rename the folder
  // @ts-ignore
  window.renameFolder = (siteName: any, folderName: any, itemId: any, siteId: any) => {
    console.log("siteName", siteName)
    console.log("documentLibraryName", folderName)
    console.log("itemId", itemId)

    // Check if a popup already exists, if so, remove it
    const existingPopup = document.getElementById("rename-popup");
    if (existingPopup) {
      existingPopup.remove();
    }

    // Create the popup container
    const popup = document.createElement("div");
    popup.id = "rename-popup";
    popup.style.position = "fixed";
    popup.style.top = "50%";
    popup.style.left = "50%";
    popup.style.transform = "translate(-50%, -50%)";
    popup.style.padding = "20px";
    popup.style.backgroundColor = "#fff";
    popup.style.boxShadow = "0px 4px 6px rgba(0,0,0,0.1)";
    popup.style.borderRadius = "8px";
    popup.style.zIndex = "1000";

    // Add the heading
    const heading = document.createElement("h3");
    heading.innerText = "Rename Folder";
    heading.style.marginBottom = "15px";
    popup.appendChild(heading);

    // Add a close button
    const closeButton = document.createElement("span");
    closeButton.innerText = "×";
    closeButton.style.position = "absolute";
    closeButton.style.top = "10px";
    closeButton.style.right = "10px";
    closeButton.style.cursor = "pointer";
    closeButton.style.fontSize = "18px";
    closeButton.onclick = () => popup.remove();
    popup.appendChild(closeButton);

    // Add the input box with the current folder name as the default value
    const input = document.createElement("input");
    input.type = "text";
    input.value = folderName; // Pre-fill with current name
    input.style.width = "100%";
    input.style.marginBottom = "15px";
    input.style.padding = "8px";
    input.style.border = "1px solid #ccc";
    input.style.borderRadius = "4px";
    popup.appendChild(input);

    // Add the submit button
    const submitButton = document.createElement("button");
    submitButton.innerText = "Submit";
    submitButton.style.padding = "10px 20px";
    submitButton.style.backgroundColor = "#0078d4";
    submitButton.style.color = "#fff";
    submitButton.style.border = "none";
    submitButton.style.borderRadius = "4px";
    submitButton.style.cursor = "pointer";
    submitButton.onclick = async () => {
      const newName = input.value.trim();
      if (newName) {
        console.log("New folder name:", newName);
        // submit the new name to the list
        try {
          if (newName === '') {
            console.log("required")
            return;
          }
          // const {web}=await sp.site.openWebById(siteId);
          await sp.web.lists.getByTitle('DMSFolderMaster').items.getById(itemId).update({
            IsRename: newName
          });
          console.log("Folder Rename successfully");
          mycreatedfolders()
          popup.remove();
        } catch (error) {
          console.log("Error in rename the folders ", error)
        }

      } else {

      }
    };
    popup.appendChild(submitButton);

    // Add the popup to the document body
    document.body.appendChild(popup);
  }

  // This function called when we click on the Rename Column option inside the mycreated Folder
  // @ts-ignore
  window.renameColumn = async (siteName: string, documentLibraryName: string) => {
    console.log("siteName", siteName)
    console.log("documentLibraryName", documentLibraryName)

    // Fetch the existing columns and types from the list
    const existingColumns = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "ColumnType", "ID", "IsRename").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibraryName}' and IsDocumentLibrary eq 0`)();

    console.log("existingColumns", existingColumns);


    // Create the popup container
    const popup = document.createElement("div");
    popup.id = "renamePopup";
    popup.style.cssText = `
    position: fixed; 
    top: 50%; 
    left: 50%; 
    transform: translate(-50%, -50%); 
    background: white; 
    padding: 20px; 
    border: 1px solid #ccc; 
    border-radius: 8px;
    box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.2); 
    z-index: 1000; 
    width: 400px;
  `;

    // Add the close button
    const closeButton = document.createElement("span");
    closeButton.innerHTML = "&times;";
    closeButton.style.cssText = `
    position: absolute; 
    top: 10px; 
    right: 15px; 
    font-size: 20px; 
    font-weight: bold; 
    color: #333; 
    cursor: pointer;
  `;
    closeButton.onclick = () => document.body.removeChild(popup);

    // Generate the form dynamically
    const formContent = existingColumns
      .map(
        (column) => `
        <div style="margin-bottom: 15px;">
          <input 
            type="text" 
            id="col-${column.ID}" 
            value="${column.IsRename !== null ? column.IsRename : column.ColumnName}" 
            data-id="${column.ID}" 
            style="width: calc(100% - 10px); padding: 8px; border: 1px solid #ccc; border-radius: 4px;" />
        </div>
      `
      )
      .join("");

    popup.innerHTML = `
    <h3 style="margin-top: 0; text-align: center; font-size: 18px;">Rename Columns</h3>
    <form id="renameForm">
      ${formContent}
      <div style="margin-top: 20px; text-align: right;">
        <button type="button" id="cancelBtn" style="
          background: #ccc; 
          border: none; 
          padding: 8px 15px; 
          border-radius: 4px; 
          cursor: pointer; 
          margin-right: 10px;">Cancel</button>
        <button type="submit" style="
          background: #0078d4; 
          color: white; 
          border: none; 
          padding: 8px 15px; 
          border-radius: 4px; 
          cursor: pointer;">Save</button>
      </div>
    </form>
  `;

    popup.prepend(closeButton); // Add the close button to the popup
    document.body.appendChild(popup);

    // Handle form submission
    document
      .getElementById("renameForm")!
      .addEventListener("submit", async (event) => {
        event.preventDefault();

        const inputs = popup.querySelectorAll("input");
        const updates = Array.from(inputs).map((input) => ({
          ID: input.getAttribute("data-id"),
          ColumnName: (input as HTMLInputElement).value,
        }));

        console.log("Updates to save:", updates);

        // Update the list with new column names
        // for (const update of updates) {
        //   await sp.web.lists
        //     .getByTitle("DMSPreviewFormMaster")
        //     .items.getById(update.ID)
        //     .update({
        //       ColumnName: update.ColumnName,
        //     });
        // }

        // Filter only updated items
        const updatedItems = updates.filter(updatedItem => {
          const oldItem = existingColumns.find(old => old.ID === parseInt(updatedItem.ID, 10));
          if (oldItem.IsRename !== null) {
            return oldItem && oldItem.IsRename !== updatedItem.ColumnName
          } else if (oldItem.IsRename === null) {
            return oldItem && oldItem.ColumnName !== updatedItem.ColumnName
          }
        });

        if (updatedItems.length > 0) {
          for (const item of updatedItems) {

            try {
              await sp.web.lists
                .getByTitle("DMSPreviewFormMaster")
                .items.getById(Number(item.ID))
                .update({
                  IsRename: item.ColumnName,
                });
              console.log("Column Updated successfully")
            } catch (error) {
              console.log("Error in Column Updated", error)
            }

          }
          Swal.fire('Success', 'Column Updated successfully', 'success');
        } else {
          Swal.fire('Warning', 'Please update atleast one column', 'warning');
          return;
        }

        console.log("Updated items:", updatedItems);



        // alert("Column names updated successfully!");
        document.body.removeChild(popup);
      });

    // Handle cancel button
    document.getElementById("cancelBtn")!.addEventListener("click", () => {
      document.body.removeChild(popup);
    });

  }


  const myFavorite = async (event: any = null, siteIdToUpdate: string = null, searchText: any = null) => {
    // // alert()
    // setlistorgriddata('')
    // setMyreqormyfav('Myfavourite')
    // // setShowButtons(true)
    // setShowMyrequButtons(false)
    // setShowMyfavButtons(true)

    // clean the url start
    const newUrl = `${window.location.origin}${window.location.pathname}`;
    window.history.pushState(null, '', newUrl)
    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }
    // end
    setTimeout(() => {
      // alert("set timer")
      setlistorgriddata('');  // Update state to '' after a delay

      console.log(listorgriddata, "list")
    }, 100);

    const wait = document.getElementById('files-container')
    wait.classList?.remove('hidemydatacards')
    setShowMyrequButtons(false)
    setShowMyfavButtons(true)
    setMyreqormyfav((previous) => 'Myfavourite')

    // setlistorgriddata('')
    // const hidegidvewlistviewbutton=document.getElementById("hidegidvewlistviewbutton")
    // if (hidegidvewlistviewbutton) {
    //   console.log("enter here .....................")
    //   hidegidvewlistviewbutton.style.display = 'flex'

    // }

    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2")
    const hidegidvewlistviewbutton = document.getElementById('hidegidvewlistviewbutton')
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'flex'

    }
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'none'

    }

    // const hidegidvewlistviewbutton2=document.getElementById("hidegidvewlistviewbutton2")
    // if (hidegidvewlistviewbutton2) {
    //   console.log("enter here .....................")
    //   hidegidvewlistviewbutton2.style.display = 'none'

    // }

    // const hidegidvewlistviewbutton2=document.getElementById("hidegidvewlistviewbutton2")
    // if (hidegidvewlistviewbutton2) {
    //   console.log("enter here .....................")
    //   hidegidvewlistviewbutton2.style.display = 'flex'

    // }

    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }

    console.log("myFavorite Function is called");

    const container = document.getElementById("files-container");
    if (siteIdToUpdate === null) {
      container.innerHTML = "";
    }

    let combineArray: any[] = [];

    // Fetch the list of active lists
    const FilesItems = await sp.web.lists
      .getByTitle("MasterSiteURL")
      .items.select("Title", "SiteID", "FileMasterList", "Active")
      .filter(`Active eq 'Yes'`)();

    console.log("Files items", FilesItems);
    console.log("searchInput", searchText);
    FilesItems.forEach(async (fileItem, index) => {
      if (fileItem.FileMasterList !== null) {

        console.log("siteIdToUpdate", siteIdToUpdate)
        // Skip rendering if we're updating only a specific list
        if (siteIdToUpdate && fileItem.SiteID !== siteIdToUpdate) {
          return;
        }

        console.log("SiteIddd", fileItem.SiteID);

        // Fetch files marked as favorite
        const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "IsDeleted", "DocumentLibraryName", "CurrentFolderPath", "SiteName", "Status", "SiteID", "FilePreviewURL")
          .filter(
            `IsFavourite eq 1 and CurrentUser eq '${currentUserEmailRef.current}'`
          ).orderBy("Modified", false)();
        // ("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID","CurrentFolderPath","DocumentLibraryName","SiteName","FilePreviewURL")

        console.log("Files", filesData);

        // Remove existing content for this specific list to avoid duplication
        const listElements = document.querySelectorAll(
          `[data-list-id='${fileItem.SiteID}']`
        );
        console.log("ListElemet To update", listElements)
        listElements.forEach((el) => el.remove());

        // start
        routeToDiffSideBar = "myFavourite";
        let filteredFileData;
        if (searchText !== null) {
          filteredFileData = filesData.filter((file: any) => file.FileName.toLowerCase().includes(searchText.value.toLowerCase()))

          combineArray = [...combineArray, ...filteredFileData]
          if (combineArray.length === 0 && searchText !== null && FilesItems.length === index + 1) {
            console.log("combineArray", combineArray);
            fileNotFound(`No files match ${searchText.value}`);
          }
          // console.log("this is filtered data",filteredFileData)
        } else {
          filteredFileData = filesData;
        }
        // end

        // change the array name
        // Render only the updated list's items
        console.log("fl data", filteredFileData)
        filteredFileData.forEach((file) => {
          console.log("hello---> ")
          console.log("file.IsDeleted", file.IsDeleted);
          console.log("file.Status", file.Status);
          if (file.IsDeleted === null) {
            const { fileIcon, fileExtension } = getFileIcon(file.FileName);
            const card = createFileCard(file, fileIcon, fileItem.SiteID, fileItem.FileMasterList, fileExtension, file.CurrentFolderPath, file.FileName);
            container.appendChild(card);
          }
        });
      }
    });

    return;
  };
  // This Function create the File card
  // This Function create the File card


  const createFileCard = (file: any, fileIcon: any, siteId: any, listToUpdate: any, fileExtension: any, FolderPath: string, fileName: string) => {
    // fileID:string,siteId:string,currentFolderPathForFile:string,fileName:string,flag:string
    const card = document.createElement("div");
    card.className = "card";
    card.dataset.fileId = file.FileUID; // Store file ID in the card element
    card.dataset.listId = siteId; // Store site ID

    card.innerHTML = `   
  <div class="row">
          <div class="col-md-2 pe-0">     
    <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
     </div>
         <div class="col-md-10 pe-0">
    <p class="p1st">${file.FileName}</p>

    <p class="p3rd">${file.FileSize} MB</p>
    </div></div>
    <div id="three-dots" class="three-dots" onclick="toggleMenu2('${file.FileUID}', '${siteId}')">
      <span>...</span>
    </div>
  `;

    const menu = document.createElement("div");
    menu.id = `menu-${file.FileUID}`;
    menu.className = "popup-menu";
    menu.innerHTML = `
    <ul>
      <li onclick="confirmDeleteFile('${file.FileUID}', '${siteId}','${false}','${listToUpdate}' )">
        <img src=${deleteIcon} alt="Delete"/> Delete
      </li>
      <li onclick="unMarkAsFavorite('${file.FileUID}', '${siteId}','${listToUpdate}')">
        <img src=${FillFavouriteFile} alt="Unmark as Favorite"/> Unmark as Favorite
      </li>
      <li onclick="shareFile('${file.FileUID}', '${file.SiteID}','${file.CurrentFolderPath}','${file.FileName}','MyFavourite','${file.FileVersion}','${file.FileSize}','${file.Status}','${file.FilePreviewURL}','${file.DocumentLibraryName}')">
        <img src=${ShareFile} alt="Share"/> Share
      </li>
    </ul>
  `;

    card.appendChild(menu);

    return card;
  };

  //Manage UnMark File  
  // @ts-ignore
  window.unMarkAsFavorite = async (fileId: number, siteId: string, listToUpdate: string) => {
    console.log("File Id is ", fileId);
    console.log("siteId is ", siteId,);
    console.log("List ", listToUpdate);

    try {
      const list = sp.web.lists.getByTitle(`${listToUpdate}`);
      console.log("List", list);
      const isFavourite = false;
      const items = await list.items.filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 0`)();
      console.log("File Data", items)
      if (items.length > 0) {
        if (items[0].IsFavourite && items[0].CurrentUser === currentUserEmailRef.current) {
          const itemId = items[0].Id;
          await list.items.getById(itemId).update({
            IsFavourite: isFavourite
          });
          console.log(`Item with FileUID '${fileId}' updated successfully.`);
          // Re-render only the modified list
          await myFavorite(null, siteId);
        }

      } else {
        console.log(`No item found with FileUID '${fileId}'.`);
      }

    } catch (error) {
      console.log("This error is from unMarkAsFavorite function ", error);
    }
  };

  // function to toggle between Favourite and UnFavourite
  // @ts-ignore
  window.toggleFavourite = async (fileId, siteId) => {

    console.log("SiteId", siteId)

    const favouriteToggle = document.getElementById(`favouriteToggle-${fileId}`);
    const markAsFavouriteIcon = favouriteToggle?.querySelector('.mark-as-favourite') as HTMLElement;
    const unMarkAsFavouriteIcon = favouriteToggle?.querySelector('.unmark-as-favourite') as HTMLElement;
    const textElement = favouriteToggle?.querySelector('.favourite-text') as HTMLElement;

    console.log("current Entity", currentEntity);
    let listToUpdate = `DMS${currentEntity}FileMaster`;

    async function markAsFavourite(fileId: any, siteId: any) {

      const siteContext = await sp.site.openWebById(siteId);
      const folderData = await siteContext.web.getFolderByServerRelativePath(currentfolderpath).files.select("Name", "Length", "ServerRelativeUrl", "UniqueId", "MajorVersion", "ListItemAllFields/Status", "ListItemAllFields/IsDeleted").expand('ListItemAllFields')();
      console.log("folderData", folderData);

      const isFavourite = true;
      const payload = {
        FileName: "",
        FileUID: fileId,
        FileVersion: "",
        FileSize: "",
        IsFavourite: isFavourite,
        CurrentUser: currentUserEmailRef.current,
        CurrentFolderPath: currentfolderpath,
        DocumentLibraryName: currentDocumentLibrary,
        FolderName: currentFolder,
        SiteName: currentEntity,
        SiteID: siteId,
        Status: "",
        FilePreviewURL: ""
      }

      folderData.forEach(async (file: any) => {
        if (file.UniqueId === fileId) {
          payload.FileName = file.Name;
          payload.FileSize = ((file.Length as unknown as number) / (1024 * 1024)).toFixed(2);
          payload.FileVersion = String(file.MajorVersion)
          payload.Status = file.ListItemAllFields.Status
          const encodedFilePath = encodeURIComponent(file.ServerRelativeUrl);
          const parentFolder = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf('/'));
          const siteUrl = window.location.origin;
          // const previewUrl = `${siteUrl}/sites/AlRostmani/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          const previewUrl = `${siteUrl}${locationPath}/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          //  const previewUrl = `${siteUrl}/sites/AlRostmanispfx2/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          console.log("previewUrl", previewUrl);

          payload.FilePreviewURL = previewUrl
        }
      })
      console.log(payload);

      // Get the list by name
      const list = sp.web.lists.getByTitle(listToUpdate);

      const data = await sp.web.lists.getByTitle(listToUpdate).items
        .filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 0`)();
      console.log("Data", data);

      // Add the new item to the list
      if (data.length > 0) {
        const itemId = data[0].Id;
        console.log("items ID", itemId);
        if (!data[0].IsFavourite && currentUserEmailRef.current === data[0].CurrentUser) {

          const updatedData = await sp.web.lists.getByTitle(listToUpdate).items.getById(itemId).update({
            IsFavourite: true
          });
          console.log("Updated data", updatedData)
        }

      } else {
        const addedItem = await list.items.add(payload);
        console.log("New item added successfully:", addedItem);
      }

    }


    async function UnmarkAsFavourite(fileId: any) {


      try {

        const data = await sp.web.lists.getByTitle(listToUpdate).items
          .filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 0 `)();

        console.log("Data", data);
        const isFavourite = false;

        if (data.length > 0) {
          const itemId = data[0].Id;
          console.log("items ID", itemId);
          if (data[0].IsFavourite && data[0].CurrentUser === currentUserEmailRef.current) {
            const updatedData = await sp.web.lists.getByTitle(listToUpdate).items.getById(itemId).update({
              IsFavourite: isFavourite
            });

            console.log("Updated data", updatedData);
          } else {
            console.log("Can not find item relataed to current user to unmark");
          }


        } else {
          console.log("No items found with FileUID:", fileId);
        }

      } catch (error) {
        console.error("Error updating the list item:", error);
      }
    }

    try {

      if (markAsFavouriteIcon && unMarkAsFavouriteIcon && textElement) {

        // Toggle visibility between the two SVGs and text content
        if (markAsFavouriteIcon.style.display === 'none') {
          markAsFavouriteIcon.style.display = 'inline';
          unMarkAsFavouriteIcon.style.display = 'none';
          textElement.textContent = 'Mark as Favourite';

          // Call function to unmark as favourite.
          UnmarkAsFavourite(fileId);
        } else {
          markAsFavouriteIcon.style.display = 'none';
          unMarkAsFavouriteIcon.style.display = 'inline';
          textElement.textContent = 'Unmark as Favourite';

          // Call function to mark as favourite.
          markAsFavourite(fileId, siteId);
        }
      }
    } catch (error) {
      console.log("This Error From toggleFavourite Function", error);
    }

  }


  // function to toggle between Favourite and UnFavourite
  // @ts-ignore
  window.toggleFavourite = async (fileId, siteId) => {

    console.log("SiteId", siteId)

    const favouriteToggle = document.getElementById(`favouriteToggle-${fileId}`);
    const markAsFavouriteIcon = favouriteToggle?.querySelector('.mark-as-favourite') as HTMLElement;
    const unMarkAsFavouriteIcon = favouriteToggle?.querySelector('.unmark-as-favourite') as HTMLElement;
    const textElement = favouriteToggle?.querySelector('.favourite-text') as HTMLElement;

    console.log("current Entity", currentEntity);
    let listToUpdate = `DMS${currentEntity}FileMaster`;

    async function markAsFavourite(fileId: any, siteId: any) {

      const siteContext = await sp.site.openWebById(siteId);
      const folderData = await siteContext.web.getFolderByServerRelativePath(currentfolderpath).files.select("Name", "Length", "ServerRelativeUrl", "UniqueId", "MajorVersion", "ListItemAllFields/Status", "ListItemAllFields/IsDeleted").expand('ListItemAllFields')();
      console.log("folderData", folderData);

      const isFavourite = true;
      const payload = {
        FileName: "",
        FileUID: fileId,
        FileVersion: "",
        FileSize: "",
        IsFavourite: isFavourite,
        CurrentUser: currentUserEmailRef.current,
        CurrentFolderPath: currentfolderpath,
        DocumentLibraryName: currentDocumentLibrary,
        FolderName: currentFolder,
        SiteName: currentEntity,
        SiteID: siteId,
        Status: "",
        FilePreviewURL: ""
      }

      folderData.forEach(async (file: any) => {
        if (file.UniqueId === fileId) {
          payload.FileName = file.Name;
          payload.FileSize = ((file.Length as unknown as number) / (1024 * 1024)).toFixed(2);
          payload.FileVersion = String(file.MajorVersion)
          payload.Status = file.ListItemAllFields.Status
          const encodedFilePath = encodeURIComponent(file.ServerRelativeUrl);
          const parentFolder = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf('/'));
          const siteUrl = window.location.origin;
          const previewUrl = `${siteUrl}${locationPath}/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          //  const previewUrl = `${siteUrl}/sites/AlRostmanispfx2/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          // const previewUrl = `${siteUrl}/sites/AlRostmani/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          // const previewUrl = `${siteUrl}/sites/SPFXDemo/${currentEntity}/${currentDocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          console.log("previewUrl", previewUrl);

          payload.FilePreviewURL = previewUrl
        }
      })
      console.log(payload);

      // Get the list by name
      const list = sp.web.lists.getByTitle(listToUpdate);

      const data = await sp.web.lists.getByTitle(listToUpdate).items
        .filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 0`)();
      console.log("Data", data);

      // Add the new item to the list
      if (data.length > 0) {
        const itemId = data[0].Id;
        console.log("items ID", itemId);
        if (!data[0].IsFavourite && currentUserEmailRef.current === data[0].CurrentUser) {

          const updatedData = await sp.web.lists.getByTitle(listToUpdate).items.getById(itemId).update({
            IsFavourite: true
          });
          console.log("Updated data", updatedData)
        }

      } else {
        const addedItem = await list.items.add(payload);
        console.log("New item added successfully:", addedItem);
      }

    }


    async function UnmarkAsFavourite(fileId: any) {


      try {

        const data = await sp.web.lists.getByTitle(listToUpdate).items
          .filter(`FileUID eq '${fileId}' and CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 0 `)();

        console.log("Data", data);
        const isFavourite = false;

        if (data.length > 0) {
          const itemId = data[0].Id;
          console.log("items ID", itemId);
          if (data[0].IsFavourite && data[0].CurrentUser === currentUserEmailRef.current) {
            const updatedData = await sp.web.lists.getByTitle(listToUpdate).items.getById(itemId).update({
              IsFavourite: isFavourite
            });

            console.log("Updated data", updatedData);
          } else {
            console.log("Can not find item relataed to current user to unmark");
          }


        } else {
          console.log("No items found with FileUID:", fileId);
        }

      } catch (error) {
        console.error("Error updating the list item:", error);
      }
    }

    try {

      if (markAsFavouriteIcon && unMarkAsFavouriteIcon && textElement) {

        // Toggle visibility between the two SVGs and text content
        if (markAsFavouriteIcon.style.display === 'none') {
          markAsFavouriteIcon.style.display = 'inline';
          unMarkAsFavouriteIcon.style.display = 'none';
          textElement.textContent = 'Mark as Favourite';

          // Call function to unmark as favourite.
          UnmarkAsFavourite(fileId);
        } else {
          markAsFavouriteIcon.style.display = 'none';
          unMarkAsFavouriteIcon.style.display = 'inline';
          textElement.textContent = 'Unmark as Favourite';

          // Call function to mark as favourite.
          markAsFavourite(fileId, siteId);
        }
      }
    } catch (error) {
      console.log("This Error From toggleFavourite Function", error);
    }

  }

  const myRequest = async (event: React.MouseEvent<HTMLButtonElement> = null, siteIdToUpdate: string = null, searchText: any = null) => {
    setdisplayuploadfileandcreatefolder(false)
    ismyrequordoclibforfilepreview = "myRequest"
    // New code to hide the create file and folder button start
    // clean Url start
    if (!cleanUrlInMyRequest) {
      const newUrl = `${window.location.origin}${window.location.pathname}`;
      window.history.pushState(null, '', newUrl)
    }
    cleanUrlInMyRequest = false;
    // end
    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }
    //End 

    setTimeout(() => {
      // alert("set timer")
      setlistorgriddata('');  // Update state to '' after a delay

      console.log(listorgriddata, "list")
    }, 100);

    const wait = document.getElementById('files-container')
    wait.classList?.remove('hidemydatacards')
    setShowMyrequButtons(true)
    setShowMyfavButtons(false)
    setMyreqormyfav('Myrequest')
    // setlistorgriddata('')
    const hidegidvewlistviewbutton = document.getElementById("hidegidvewlistviewbutton")
    if (hidegidvewlistviewbutton) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton.style.display = 'flex'

    }
    const hidegidvewlistviewbutton2 = document.getElementById("hidegidvewlistviewbutton2")
    if (hidegidvewlistviewbutton2) {
      console.log("enter here .....................")
      hidegidvewlistviewbutton2.style.display = 'none'

    }


    // console.log(listorgriddata , "listorgriddata")
    console.log("searchInput", searchText);
    console.log("siteIdToUpdate", siteIdToUpdate);

    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }





    // call this function onClick of the myRequest
    // handleShowContent(event)


    if (createFileButton2) {
      createFileButton2.style.display = 'none'
    }
    if (createFileButton) {
      createFileButton.style.display = 'none'
    }



    if (event) {
      event.preventDefault();
      event.stopPropagation();
    }

    // console.log("myFavorite Function is called");

    const container = document.getElementById("files-container");
    if (siteIdToUpdate === null) {
      container.innerHTML = "";
      // console.log("siteToUpdate")
    }

    // console.log("beforeFetchItems");
    // Fetch the list of active entity
    const FilesItems = await sp.web.lists
      .getByTitle("MasterSiteURL")
      .items.select("Title", "SiteID", "FileMasterList", "Active")
      .filter(`Active eq 'Yes'`)();

    // console.log("Active Sites List Names", FilesItems);

    FilesItems.forEach(async (fileItem, index) => {
      if (fileItem.FileMasterList !== null) {

        // console.log("FilesItesms");
        // Skip rendering if we're updating only a specific list
        if (siteIdToUpdate && fileItem.SiteID !== siteIdToUpdate) {
          return;
        }

        // console.log("SiteId", fileItem.SiteID);
        console.log("fileItem.FileMasterList", fileItem.FileMasterList);
        // const filesData = await sp.web.lists
        //   .getByTitle(`${fileItem.FileMasterList}`)
        //   .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID","CurrentFolderPath","DocumentLibraryName","SiteName","FilePreviewURL","IsDeleted")
        //   .filter(
        //     `CurrentUser eq '${currentUserEmailRef.current}'`
        //   )();
        const filesData = await sp.web.lists
          .getByTitle(`${fileItem.FileMasterList}`)
          .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "Status", "SiteID", "CurrentFolderPath", "DocumentLibraryName", "SiteName", "FilePreviewURL", "IsDeleted", "MyRequest").filter(
            `CurrentUser eq '${currentUserEmailRef.current}' and MyRequest eq 1`
          ).orderBy("Modified", false)();
        console.log("My reaquest Called");

        // console.log("enter in the myRequest------")
        console.log(fileItem.FileMasterList, "- FilesData", filesData)
        // route to different-2 sideBar

        let combineArray: any[] = [];
        // start
        routeToDiffSideBar = "myRequest";
        let filteredFileData = [];
        if (searchText !== null) {
          filteredFileData = filesData.filter((file: any) => file?.FileName?.toLowerCase().includes(searchText?.value?.toLowerCase()))
          // console.log("this is filtered data",filteredFileData)
          // New Code to show pop up when no match found start
          combineArray = [...combineArray, ...filteredFileData]
          if (combineArray.length === 0 && searchText !== null && FilesItems.length === index + 1) {
            console.log("combineArray", combineArray);
            fileNotFound(`No files match ${searchText.value}`);
          }
          // End
          // console.log("Index",index);
        } else {
          filteredFileData = filesData;
        }
        // end 

        // change the array
        filteredFileData.forEach((file) => {
          //  console.log(file.ID , "file.odata.id ")
          if (file.IsDeleted === null) {
            const card = document.createElement("div");
            const fileSizeInKB = 2048; // Example file size in KB
            const fileSizeInMB = (fileSizeInKB / 1024).toFixed(2); // Convert to MB and round to 2 decimal places

            // console.log("searchArray",searchArray);
            let fileIcon;
            const fileExtension = file.FileName?.split(".").pop().toLowerCase(); // Get the file extension
            switch (fileExtension) {
              case "doc":
              case "docx":
                fileIcon = Docicon;
                break;
              case "txt":
                fileIcon = Txticon;
                break;
              case "pdf":
                fileIcon = Pdficon;
                break;
              case "xls":
              case "xlsx":
                fileIcon = Xlsicon;
                break;
              case "zip":
                fileIcon = Zipicon;
                break;
              default:
                fileIcon = Docicon; // Default icon if no match
                break;
            }

            card.className = "card";
            card.innerHTML = ` 
    <div class="row"> 
      <div class="col-md-2 pe-0"> 
    <div class="IMGContainer">        
      <img class="filextension" src=${fileIcon} alt="${fileExtension} icon"/>
    </div>
    </div>
        <div class="col-md-10"> 
         <div class="CardTextContainer">
      <p class="p1st">${file.FileName}</p>
      <p class="p2nd"></p>
      <p class="p3rd ">${((file.FileSize as unknown as number) / (1024 * 1024)).toFixed(2)}MB</p>
      <p class="filestatus myrequestp3rd"> ${file.Status ? file.Status : ''}  </p>
      </div>

      <div class="three-dots" onclick="toggleMenu2('${file.FileUID}','${fileItem.SiteID}','${file.ID}' , '${fileItem.FileMasterList}')  ">
          <span>...</span>
      </div>
          </div> </div>
    `;

            const menu = document.createElement("div");
            // console.log(menu , "menu is here")
            menu.id = `menu-${file.FileUID}`;
            menu.className = "popup-menu";
            const showaudit = <FontAwesomeIcon style={{ color: "black" }} icon={faListSquares} />
            menu.innerHTML = `
     <ul>
    <li onclick="confirmDeleteFile('${file.FileUID}','${file.SiteID}','${false}','${fileItem.FileMasterList}')">
      <img src=${deleteIcon} alt="Delete"/> Delete
    </li>
    <li onclick="auditHistory('${file.FileUID}', '${file.SiteID}','${file?.DocumentLibraryName}','${file?.SiteName}')">
          <img src=${editIcon} alt="Edit"/>
                      Audit History
    </li>
    <li onclick="shareFile('${file.FileUID}', '${file.SiteID}','${file.CurrentFolderPath}','${file.FileName}','MyRequest','${file.FileVersion}','${file.FileSize}','${file.Status}','${file.FilePreviewURL}','${file.DocumentLibraryName}')">
      <img src=${ShareFile} alt="Share"/> Share
    </li>
     <li onclick="PreviewFile('${file.FileUID}','${file.SiteID}','${file.ID}' , '${file.FileMasterList}', '${file.FilePreviewURL}')">
         <img src=${viewIcon} alt="Preview File"/> Preview File
       </li>
       <li onclick="Download('${file.FileUID}','${file.SiteID}','${file.ID}' , '${file.FileMasterList}', '${file.FilePreviewURL}')">
         <img src=${downloadicon} alt="Download File"/> Download File
       </li>
    ${file.Status === "Rework" ? `
      <li onclick="Download('${file.FileUID}','${file.SiteID}','${file.ID}' , '${file.FileMasterList}', '${file.FilePreviewURL}')">
          <img src=${editIcon} alt="Edit File"/> Edit File
      </li>` : ''}
  </ul>
    `;



            card.appendChild(menu);
            // Change the background color and text color based on FileStatus
            const fileStatusElement = card.querySelector(".filestatus") as HTMLElement;
            switch (file.Status) {
              case "Approved":
                fileStatusElement.style.backgroundColor = "#b5e7d3";
                fileStatusElement.style.color = "#008751";
                break;
              case "Auto Approved":
                fileStatusElement.style.backgroundColor = "#b5e7d3";
                fileStatusElement.style.color = "#008751";
                fileStatusElement.style.width = "96px";

                break;
              case "Rejected":
                fileStatusElement.style.backgroundColor = "rgba(241, 85, 108, 0.1)";
                fileStatusElement.style.color = "#f1556c";
                break;
              case "Rework":
                fileStatusElement.style.backgroundColor = "#ffecc4";
                fileStatusElement.style.color = "rgba(247, 184, 75)";
                break;
              case "Pending":
                fileStatusElement.style.backgroundColor = "rgb(91 156 187 / 25%)";
                fileStatusElement.style.color = "#000b56";
                break;
              default:
                fileStatusElement.style.backgroundColor = "none";
                fileStatusElement.style.color = "none";
                break;
            }

            container.appendChild(card);
            // check file status if approved hide the delete button
            const menu1 = document.getElementById(`menu-${file.FileUID}`);
            // console.log("menu1",menu1);
            if (file.Status === "Approved" || file.Status === null) {
              const firstItem = menu1.children[0]?.children[0] as HTMLElement;
              if (firstItem && firstItem.style.display !== "none") {
                firstItem.style.display = "none";
              }
            }
          }

        });
      }
    });

  };
  // Show Error Message on file not Found start
  const fileNotFound = (fileName: any) => {
    Swal.fire(`No results found`, `${fileName}`, "warning");
  }
  // end
  const [activeComponent, setActiveComponent] = useState<string | 'MyRequest'>('');
  const [listorgriddata, setlistorgriddata] = useState<string>('');
  const handleButtonClickShow = (componentName: any) => {
    setActiveComponent(componentName); // Set the active component based on the button clicked
  };
  const handleReturnToMain = () => {
    setActiveComponent(''); // Reset to show the main component
  };


  const MyrequestshowListView = (componentName: any) => {
    const wait = document.getElementById('files-container')
    wait.classList?.add('hidemydatacards')
    setlistorgriddata('showListView');
  }

  // side text content based on click 
  // Handle button click and show the text of the clicked button
  const [selectedText, setSelectedText] = useState<string | null>(null);
  const [dynamicContent, setDynamicContent] = useState<string | null>(null);
  // Function to update the breadcrumb navigation
  const updateBreadcrumb = (path: any) => {
    console.log(path, "path")
    // For toggle the breadcrumb and selectedTextForSideBar
    const selectedTextDiv = document.getElementById('selectedText');
    const breadcrumbElement = document.getElementById("breadcrumb");

    if (breadcrumbElement) {
      // breadcrumbElement.style.position = "absolute"
      breadcrumbElement.style.width = ""      // Change width  by Amjad
      breadcrumbElement.style.top = "115px"
      breadcrumbElement.style.display = 'block';
    }
    if (selectedTextDiv) {
      selectedTextDiv.style.display = 'none';
    }



    if (breadcrumbElement) {
      breadcrumbElement.textContent = path;
    }
  };
  // Function to handle navigation and update breadcrumb
  const handleNavigation = (title: string, Devision: string, Department: string, docLibName: string = null, folderName: string = null) => {
    let path = title;
    if (Devision) {
      path += ` > ${Devision}`;
    }
    if (Department) {
      path += ` > ${Department}`;
    }
    if (docLibName) {
      path += ` > ${docLibName}`;
    }

    if (folderName) {
      path += ` > ${folderName}`;
    }

    updateBreadcrumb(path);
  };
  const handleShowContent = (event: React.MouseEvent<HTMLButtonElement>) => {
    // console.log("enter here")
    event.preventDefault();

    //toggle the breadcrumb and selectedText For SideBar
    const selectedTextDiv = document.getElementById('selectedText');
    const breadcrumbElement = document.getElementById("breadcrumb");
    breadcrumbElement.style.display = 'none';
    selectedTextDiv.style.display = 'block';


    // Find the button that was clicked
    const button = event.currentTarget;


    const spanElement = button.querySelector('.sidebarText');
    const text = spanElement?.textContent;

    if (text) {
      setSelectedText(text);

      // Update dynamic content based on the button clicked
      switch (text) {
        case 'My Request':
          setDynamicContent('Mentioned below are the documents submitted by logged in user.');
          break;
        case 'My Favourite':
          setDynamicContent('All the files and folder which is marked as Favourite.');
          break;
        case 'My Folder':
          setDynamicContent('All the folder Created by me.');
          break;
        case 'Shared with Others':
          setDynamicContent('My files shared with other users.');
          break;
        case 'Shared with me':
          setDynamicContent('File upload by other team members and shared with me.');
          break;
        case 'Recycle Bin':
          setDynamicContent('below are the documents Deleted by logged in use.');
          break;
        default:
          setDynamicContent(null);
      }
    }
  };

  const search = document.getElementById('searchinput')
  if (search) {
    search.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();  // Prevent default behavior
      }
    })
  };

  // // Manage Folder WorkFlow Action
  window.manageWorkflow = async (DocumentLibraryName: string, SiteTilte: string, SiteID: string) => {
    setShowWorkflow(true)
    console.log("enter in workflow")
    // const workflowdiv= document.getElementById('workflowdiv')
    // if(workflowdiv){
    //   workflowdiv.classList.remove('workflowdivhide')
    //   alert(workflowdiv.classList)
    //   workflowdiv.classList.add('workflowdiv')
    // }
    // setshowworkflowdiv("true")

    // setshowworkflowdiv(true)
    // console.log("Inside manageWorkFlow");

    propsForManageWorkFlow.SiteTitle = SiteTilte;
    propsForManageWorkFlow.DocumentLibraryName = DocumentLibraryName;
    propsForManageWorkFlow.SiteID = SiteID;
    // handleButtonClickShow("manageWorkFlow");

  }

  //Manage Folder Permission Action
  window.managePermissionSearch = (documentLibraryName: string, SiteTilte: string, SiteID: string, folderName: any, folderPath: any) => {
    setShowfolderpermission(true)
    // console.log(message);
    console.log("documentLibraryName", documentLibraryName)
    console.log("SiteName", SiteTilte);
    console.log("siteId", SiteID);
    console.log("folderName", folderName);
    console.log("folderPath", folderPath);

    managePermissionProps.DocumentLibraryName = documentLibraryName;
    managePermissionProps.SiteTitle = SiteTilte;
    managePermissionProps.SiteID = SiteID;
    managePermissionProps.FolderName = folderName;
    managePermissionProps.FolderPath = folderPath;
    //  handleButtonClickShow("managePermissionSearch");

    // handleButtonClickShow("managePermissionSearch");

  }

  // Edit File
  window.editFile = async (siteName: string, documentLibraryName: string) => {
    console.log("Inside the editFile");

    // Fetch the existing columns and types from the list
    const existingColumns = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "ColumnType", "ID", "IsRename").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibraryName}' and IsDocumentLibrary eq 0`)();

    console.log("existingColumns", existingColumns);

    // Create the popup container dynamically
    const popupContainer = document.createElement("div");
    popupContainer.className = "edit-popup";

    // Append to body
    document.body.appendChild(popupContainer);

    // Create close button
    const closeButton = document.createElement("span");
    closeButton.innerHTML = 'x';
    closeButton.className = 'close-button';
    closeButton.style.cursor = 'pointer';
    closeButton.style.fontSize = '24px';
    closeButton.style.position = 'absolute';
    closeButton.style.top = '10px';
    closeButton.style.right = '15px';

    // Append close button to popup
    popupContainer.appendChild(closeButton);


    // Add form elements for each existing column
    const formContent = document.createElement("div");
    formContent.innerHTML = existingColumns.map((col) => `
    <div class="form-group">
      <div class="col-md-5">
        <label>Field Name</label>
        <input type="text" class="form-control" value="${col.IsRename !== null ? col.IsRename : col.ColumnName}" disabled />
      </div>
      <div class="col-md-5">
        <label>Field Type</label>
        <input type="text" class="form-control" value="${col.ColumnType}" disabled />
      </div>
    </div>
  `).join('');


    // Add button for adding new fields
    const addFieldButton = document.createElement("a");
    addFieldButton.innerHTML = `
    <img class="bi bi-plus" src="${require("../../assets/plus.png")}" alt="add" style="width: 50px; height: 50px;" />
  `;

    // Append the content to the popup
    popupContainer.appendChild(formContent);
    popupContainer.appendChild(addFieldButton);

    // Add event listener for "+" button to add new editable fields
    addFieldButton.addEventListener('click', () => {
      const newFieldHTML = `
      <div class="form-group row">
        <div class="col-md-5">
          <label>Field Name</label>
          <input type="text" class="form-control" placeholder="Enter new field name" />
        </div>
        <div class="col-md-5">
          <label>Field Type</label>
          <select class="form-control">
            <option value="" disabled selected>Select Type</option>
            <option value="Single Line of Text">Single Line of Text</option>
            <option value="Mulitple Line of Text">Multiple Line of Text</option>
            <option value="Yes or No">Yes or No</option>
            <option value="Date & Time">Date & Time</option>
            <option value="Number">Number</option>
          </select>
        </div>
        <div class="col-md-2">
          <img class="delete-column"  src="${require("../../assets/delete.png")}" alt="add" style="width: 40px; margin-top:25px; cursor:pointer;" />
        </div>
      </div>
    `;
      formContent.insertAdjacentHTML('beforeend', newFieldHTML);
    });

    // Add save button
    const saveButton = document.createElement("button");
    saveButton.innerText = "Save";
    popupContainer.appendChild(saveButton);

    saveButton.addEventListener('click', () => {
      debugger
      // Collect data from newly added fields
      console.log(formContent, "formContent")
      const newFields = Array.from(formContent.querySelectorAll('.form-group.row')).map((group) => {
        console.log(group, "group")
        const fieldNameInput = (group.querySelector('input[type="text"]') as HTMLInputElement);
        const fieldTypeSelect = (group.querySelector('select') as HTMLSelectElement);

        return {
          columnName: fieldNameInput.value,
          columnType: fieldTypeSelect.value
        };
      });

      console.log("New Fields:", newFields);

      try {
        const payloadForPreviewFormMaster = {
          SiteName: siteName,
          DocumentLibraryName: documentLibraryName,
          IsRequired: true,
          AddorRemoveThisColumn: "Add To Library",
          IsModified: true
        }

        // existingColumns.forEach(async(column)=>{
        //       try {
        //           // Deleting the item with the provided ID from the DMSPreviewFormMaster list
        //           await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.getById(column.ID).delete();
        //           console.log(`Item with ID ${column.ID} deleted successfully from list DMSPreviewFormMaster`);
        //       } catch (error) {
        //           console.log(`Error deleting item with ID ${column.ID} from list DMSPreviewFormMaster`, error);
        //       }
        //   });

        newFields.forEach(async (column) => {
          (payloadForPreviewFormMaster as any).ColumnName = column.columnName.replace(/\s+/g, '');;
          (payloadForPreviewFormMaster as any).ColumnType = column.columnType;

          const addedItem = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.add(payloadForPreviewFormMaster);
          console.log("Item added successfully in the DMSPreviewFormField", addedItem);

        })
      } catch (error) {
        console.log("Error in adding columns in the DMSPreviewFormMaster inside the editFile onclick of the save button", error)
      }


    });

    // Add event listener to close button
    closeButton.addEventListener('click', () => {
      popupContainer.style.display = 'none';
    });

    //add event listener to  removing the field
    // Event delegation
    formContent.addEventListener('click', (event) => {
      const target = event.target as HTMLElement;
      if (target.classList.contains("delete-column")) {
        const fieldGroup = target.closest('.form-group.row');
        if (fieldGroup) {
          fieldGroup.remove();
        }
      }
    });
  };

  //  Share File
  window.shareFile = async (fileID: string, siteId: string, currentFolderPathForFile: string, fileName: string, flag: string, FileVersion: any, FileSize: any, Status: any, FilePreviewURL: any, DocumentLibraryName: any) => {
    const testidsub = await sp.site.openWebById(siteId)

    let filePath = `${currentFolderPathForFile}/${fileName}`;

    const fileServerRelativePath = testidsub.web.getFileByServerRelativePath(filePath);

    const item = await fileServerRelativePath.getItem();


    const filePermissions = await item.getCurrentUserEffectivePermissions();

    const hasFullControl = testidsub.web.hasPermissions(filePermissions, PermissionKind.ManageWeb);
    const hasEdit = testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
    const hasContribute = testidsub.web.hasPermissions(filePermissions, PermissionKind.AddListItems) && testidsub.web.hasPermissions(filePermissions, PermissionKind.EditListItems);
    const hasRead = testidsub.web.hasPermissions(filePermissions, PermissionKind.ViewListItems);

    let filePermission: string;
    if (hasFullControl) {
      filePermission = "Full Control";
    } else if (hasEdit) {
      filePermission = "Edit";
    } else if (hasContribute) {
      filePermission = "Contribute";
    } else if (hasRead) {
      filePermission = "Read";
    } else {
      filePermission = "No Access";
    }


    const parts = currentFolderPathForFile.split("/");
    const entity = parts[3];
    console.log(entity);

    const fetchUser = async (entity: any) => {
      const user0 = await sp.web.siteUsers();
      const combineUsersArray = user0.map((user) => (
        {
          id: String(user.Id),
          value: user.Title,
          email: user.Email,
        }
      ))
      console.log("Sub site users", combineUsersArray);

      return combineUsersArray;
    }

    const users = await fetchUser(entity);
    const existingPopup = document.getElementById('share-popup');
    if (existingPopup) {
      existingPopup.remove();
    }

    let selectedUsers: { id: string; value: string; email: string }[] = [];

    const popup = document.createElement("div");
    popup.id = 'share-popup';
    popup.className = "share-popup";

    let options = ''
    if (filePermission === "Full Control") {
      options = `
  <option value="Full Control">Full Control</option>
  <option value="Contribute">Contribute</option>
  <option value="Edit">Edit</option>
  <option value="Read">Read</option>
`
    } else if (filePermission === "Contribute" || filePermission === "Edit") {
      options = `
<option value="Contribute">Contribute</option>
<option value="Edit">Edit</option>
<option value="Read">Read</option>
`
    } else if (filePermission === "Read") {
      options = `
<option value="Read">Read</option>
`
    }

    popup.innerHTML = `
<div class="share-popup-content">
<div class="share-popup-header">
<h4>Share</h4>
<span class="share-close-popup" onClick="hideSharePopUp()">x</span>
</div>
<div class="share-popup-body">
<div id="share-reactSelect">
    <input type="text" id="userInput" placeholder="Add a Name, Group, or Email" style="
    width: 100%; 
    padding: 10px;
    font-size: 14px;
    border-radius: 4px;
    border: 1px solid #ccc;
  "/>
  <div id="userDropdown" class="user-dropdown" style="
    display: none;
    position: absolute;
    width: 100%;
    max-height: 150px;
    overflow-y: auto;
    background-color: white;
    border: 1px solid #ccc;
    border-radius: 4px;
    z-index: 1000;
  ">
  </div>
</div>
 <div>
  <select id="permissionSelect" style="
    margin-bottom:10px;
    width: 100%; 
    padding: 10px;
    font-size: 14px;
    border-radius: 4px;
    border: 1px solid #ccc;
    margin-top: 10px;
  ">
    <option value="" disabled selected>Permission</option>
    ${options}
  </select>
</div>
<textarea id="share-message" placeholder="Write a message..." >
</textarea>
</div>
<div class="share-popup-footer">
<button id="share-shareFileButton">Share</button>
</div>
</div>
`;

    // Append the  popup to the body
    document.body.appendChild(popup);

    // Get references to the input box and dropdown
    const userInput = document.getElementById('userInput') as HTMLInputElement;
    const userDropdown = document.getElementById('userDropdown');

    // Function to render dropdown options based on user input
    function renderDropdown(users: { id: string, value: string, email: string }[]) {
      // Clear previous options
      userDropdown.innerHTML = '';
      users.forEach(user => {
        const option = document.createElement('div');
        option.className = 'dropdown-item';
        option.style.padding = '8px';
        option.style.cursor = 'pointer';
        option.textContent = user.value;
        option.onclick = () => selectUser(user);
        userDropdown.appendChild(option);
      });
    }

    // Function to show the dropdown when the input is clicked
    userInput.addEventListener('focus', () => {
      userDropdown.style.display = 'block';

      // Display all users initially
      renderDropdown(users);
    });

    // Filter dropdown based on input value
    userInput.addEventListener('input', () => {
      const searchValue = userInput.value.toLowerCase();
      const filteredUsers = users.filter(user => user.value.toLowerCase().includes(searchValue));
      renderDropdown(filteredUsers);
    });

    // Function to select a user and display it inside the input
    function selectUser(user: { id: string, value: string, email: string }) {
      console.log("selected user", selectedUsers)
      if (!selectedUsers.some(selectedUser => selectedUser.id === user.id)) {

        selectedUsers.push(user);

        // Create a span for the selected user with a close button
        const selectedUserDiv = document.createElement('span');
        selectedUserDiv.className = 'selected-user';
        selectedUserDiv.style.display = 'inline-block';
        selectedUserDiv.style.padding = '2px 6px';
        selectedUserDiv.style.backgroundColor = '#e0e0e0';
        selectedUserDiv.style.borderRadius = '12px';
        selectedUserDiv.style.marginRight = '5px';
        selectedUserDiv.style.position = 'relative';

        selectedUserDiv.textContent = user.value;

        // Create close button for deselecting the user
        const closeButton = document.createElement('span');
        closeButton.textContent = 'x';
        closeButton.style.cursor = 'pointer';
        closeButton.style.marginLeft = '5px';
        closeButton.onclick = () => deselectUser(user.id, selectedUserDiv);
        selectedUserDiv.appendChild(closeButton);

        // Append the selected user to the input field
        userInput.parentNode!.insertBefore(selectedUserDiv, userInput);
        userInput.value = '';
      }
      userDropdown.style.display = 'none';
    }

    // Function to deselect a user
    function deselectUser(userId: string, selectedUserDiv: HTMLElement) {
      // selectedUsers = selectedUsers.filter(id => id !== userId);
      selectedUsers = selectedUsers.filter(selectedUser => selectedUser.id !== userId);
      console.log("selected user", selectedUsers);
      selectedUserDiv.remove();
    }

    // Hide the dropdown if clicked outside
    document.addEventListener('click', (event) => {
      if (!userInput.contains(event.target as Node) && !userDropdown.contains(event.target as Node)) {
        userDropdown.style.display = 'none';
      }
    });

    // Capture selected permission
    let selectedPermission = "";
    document.getElementById('permissionSelect').addEventListener('change', (event) => {
      selectedPermission = (event.target as HTMLSelectElement).value;
      console.log("Selected Permission:", selectedPermission);
    });

    // Adding event listener to the "Share" button
    document.getElementById('share-shareFileButton').addEventListener('click', async function () {
      const filePath = `${currentFolderPathForFile}/${fileName}`;
      console.log("filePath", filePath);
      // Check the Break role on the file start
      const testidsub = await sp.site.openWebById(siteId);
      const file = testidsub.web.getFileByServerRelativePath(filePath);
      const item = await file.getItem();
      const itemData = await item.select("HasUniqueRoleAssignments")();
      const breaKRole = itemData.HasUniqueRoleAssignments;
      console.log("breaKRole", breaKRole);
      if (!breaKRole) {
        // Break role inheritance, keeping current permissions
        await item.breakRoleInheritance(true);
        console.log("Inheritance broken, retaining previous permissions.");
      }

      // end

      // New Code push the data into the DMSShareWithOtherMaster Start
      try {
        const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
        const payloadForDMSShareWithOtherMaster = {
          FileName: fileName,
          FileUID: fileID,
          CurrentUser: currentUserEmailRef.current,
          CurrentFolderPath: currentFolderPathForFile,
          SiteName: entity,
          PermissionType: selectedPermission,
          ShareAt: isoDate,
          FileVersion: FileVersion,
          FileSize: FileSize,
          Status: Status,
          FilePreviewURL: FilePreviewURL,
          SiteID: siteId,
          DocumentLibraryName: DocumentLibraryName
        }
        let roleType: number;
        if (selectedPermission === "Full Control") {
          // roleType=5;
          roleType = 1073741829;
        } else if (selectedPermission === "Contribute") {
          // roleType=3;
          roleType = 1073741827;
        } else if (selectedPermission === "Edit") {
          // roleType=6;
          roleType = 1073741830;
        } else if (selectedPermission === "Read") {
          // roleType=2;
          roleType = 1073741826;
        } else {
          roleType = 0;
        }
        console.log("roletype", roleType);
        const Promiseshare = selectedUsers.map(async (user) => {
          (payloadForDMSShareWithOtherMaster as any).UserID = user.id;
          (payloadForDMSShareWithOtherMaster as any).ShareWithOthers = user.value;
          (payloadForDMSShareWithOtherMaster as any).ShareWithMe = user.email;
          const newItem = await sp.web.lists.getByTitle(`DMSShareWithOtherMaster`).items.add(payloadForDMSShareWithOtherMaster)

          //Add permission to the user in the file 
          const id = Number(user.id)
          console.log("User Id", id, "type", typeof id);
          await item.roleAssignments.add(id, roleType);
          console.log(`User ${user.email} added with role type ${selectedPermission},${roleType}.`);
          console.log("Data added successfully in the", newItem);
        })

        await Promise.all(Promiseshare)


        const existingPopup = document.getElementById('share-popup');
        if (existingPopup) {
          existingPopup.remove();
          Swal.fire({
            title: "Access Provided Successfully",
            text: "The access has been provided successfully.",
            icon: "success"
          })
        }

      } catch (error) {
        console.log("Error in adding data to the DMSShareWithOtherMaster", error);
      }

    });

  }



  // hide the share popup
  // @ts-ignore
  window.hideSharePopUp = () => {
    const popup = document.querySelector('.share-popup');

    if (popup) {
      popup.remove();
    }
  }

  //Download file popup 
  window.Download = async (path: any, siteID: any, docLibName: any, filemasterlist: any, filepreview: any) => {
    try {
      // Get the web context for the site
      console.log(siteID, "siteID")
      const siteWeb = await sp.site.openWebById(siteID)
      console.log(path, "path")
      // Fetch the file using its UniqueId (GUID)
      const file = await siteWeb.web.getFileById(path)();

      // Log the file information
      console.log("File details: ", file);

      // Get the file download URL
      // const fileUrl = file.ServerRelativeUrl;

      // Create a link element for the file download
      const link = document.createElement('a');
      link.href = file.ServerRelativeUrl;
      link.download = file.Name; // Optional: Specify the file name for download

      // Programmatically trigger the download
      link.click();

    } catch (error) {
      console.error("Error downloading file: ", error);
    }
  }
  // show the audit history popup
  // @ts-ignore
  window.auditHistory = async (fileId: string, siteId: string, DocumentLibraryName: string, SiteName: String) => {
    console.log("Audit History called", fileId, siteId);
    console.log("Audit History called", SiteName);
    console.log("Audit History called", DocumentLibraryName);


    const { web } = await sp.site.openWebById(siteId)

    // Get the list item  corresponding to the file
    const fileItem: any = await web.getFileById(fileId).expand("ListItemAllFields")();
    console.log("fileItem", fileItem.ListItemAllFields.Status);

    // fetched the columns details corresponding to the file 
    const fileColumns = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "SiteName", "DocumentLibraryName", "IsRename").filter(`SiteName eq '${SiteName}' and DocumentLibraryName eq '${DocumentLibraryName}' and IsDocumentLibrary ne 1`)();
    console.log("fileColumns", fileColumns);

    // Create an array of objects to store the columnName with there corresponding value
    const resultArrayThatContainstheColumnDetails = fileColumns.map((column) => {

      let columnName = column.ColumnName;
      const columnValue = fileItem.ListItemAllFields[columnName];
      if (column.IsRename !== null) {
        columnName = column.IsRename
      }

      return {
        label: columnName,
        value: columnValue !== undefined ? columnValue : null // Handle missing fields
      };
    });

    const objectForStatus = {
      label: "Status",
      value: fileItem.ListItemAllFields.Status || ""
    }

    resultArrayThatContainstheColumnDetails.push(objectForStatus);
    console.log("result", resultArrayThatContainstheColumnDetails);

    // get the details of approver
    const itemsFromTaskList = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
      "Log", "CurrentUser", "Remark"
      , "LogHistory", "ID"
      , "FileUID/FileUID"
      , "FileUID/SiteName"
      , "FileUID/DocumentLibraryName"
      , "FileUID/FileName"
      , "FileUID/Status"
      , "FileUID/RequestedBy"
      , "FileUID/Created"
      , "FileUID/ApproveAction"
      , "MasterApproval/ApprovalType"
      , "MasterApproval/Level"
      , "MasterApproval/DocumentLibraryName"
      , "Modified"
    )
      .expand("FileUID", "MasterApproval")
      .filter(`FileUID/FileUID eq '${fileId}'`)
      .orderBy("Modified", false)();

    console.log("itemsFromTaskList", itemsFromTaskList);


    // Mapping to the desired format
    const approverDetailsArray = itemsFromTaskList.map(task => ({
      level: `Level ${task.MasterApproval.Level}`,
      approver: task.CurrentUser,
      actionDateTime: task.Modified,
      status: task.Log || "",
      remark: task.Remark || ""
    }));

    console.log("approverDetailsArray", approverDetailsArray);

    // Generate the dynamic HTML for the detail rows
    let detailRowsHTML = "";
    resultArrayThatContainstheColumnDetails.forEach((item, index) => {
      // Start a new row every 3rd item (when index is 0, 3, 6, ...)
      if (index % 3 === 0) {
        detailRowsHTML += '<div class="detail-row">';
      }

      // Add each detail column
      detailRowsHTML += `
      <div class="detail-column">
        <div class="detail-label">${item.label}</div>
        <div class="detail-value">${item.value}</div>
      </div>
    `;

      // Close the row after 3 items (when index is 2, 5, 8, ...)
      if ((index + 1) % 3 === 0) {
        detailRowsHTML += '</div>';
      }
    });

    // If there are leftover columns (less than 3 in the last row), close the row
    if (resultArrayThatContainstheColumnDetails.length % 3 !== 0) {
      detailRowsHTML += '</div>';
    }

    // Generate the dynamic HTML for the approver details
    let approverRowsHTML = "";
    approverDetailsArray.forEach((approver) => {
      approverRowsHTML += `
     <div class="detail-row-value-approver">
       <div class="detail-value-approver">${approver.level}</div>
       <div class="detail-value-approver">${approver.approver}</div>
       <div class="detail-value-approver">${approver.actionDateTime}</div>
       <div class="detail-value-approver">${approver.status}</div>
       <div class="detail-value-approver">${approver.remark}</div>
     </div>
   `;
    });

    // Create the popup
    const popup = document.createElement("div");
    popup.className = "audit-history-popup";
    popup.innerHTML = `
  <div class="popup-content-auditHistory">
    <div class="popup-header">
      <h5>Audit History</h5>
      <span class="close-btn" onclick="hideAuditHistoryPopup()">&times;</span>
    </div>
    <div class="popup-details">
      ${detailRowsHTML}
      <div class="detail-row-approver">
        <div class="detail-label-approver">Approval Level</div>
        <div class="detail-label-approver">Approver</div>
        <div class="detail-label-approver">Action DateTime</div>
        <div class="detail-label-approver">Status</div>
        <div class="detail-label-approver">Remark</div>
      </div>
      ${approverRowsHTML}
    </div>
  </div>
  `;

    // Append to body
    document.body.appendChild(popup);
  }

  // function to hide audit history pop
  // @ts-ignore
  window.hideAuditHistoryPopup = () => {
    const popup = document.querySelector(".audit-history-popup");
    if (popup) {
      popup.remove();
    }
  }
  // start
  // Ref for MyFolder
  const myFolderButtonRef = useRef(null);
  const [triggerClick, setTriggerClick] = useState(false);

  const handleReturnToMainFromManageWorkFlow = () => {
    setShowWorkflow(false);
    setShowfolderpermission(false)
    setActiveComponent('');
    setTriggerClick(true);
  }

  useEffect(() => {
    if (triggerClick && myFolderButtonRef.current) {
      myFolderButtonRef.current.click();
      setTriggerClick(false);
    }
  }, [triggerClick]);
  useEffect(() => {
    const workflowdiv = document.getElementById('workflowdiv')
    if (workflowdiv) {
      workflowdiv.classList.add('workflowdivhide')
    }
  })

  const checkValidation = (message: any) => {
    Swal.fire(`${message}`, ``, "warning");
  }

  const showReplaceMessage = (message: any) => {
    Swal.fire(`${message}`, ``, 'success');
  }




  // This function called when we click on the edit option inside the myrequest in case of rework start
  // @ts-ignore
  window.rework = async (fileId: any, siteId: any, documentLibrary: any, siteName: any, filePath: any) => {
    console.log("Filepath", filePath);
    const { web } = await sp.site.openWebById(siteId)

    let clickedReplace = false;


    // Get the list item  corresponding to the file
    const fileItem: any = await web.getFileById(fileId).expand("ListItemAllFields")();
    console.log("fileItem", fileItem.ListItemAllFields.Status);

    // fetched the columns details corresponding to the file
    const fileColumns = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "SiteName", "DocumentLibraryName", "IsRequired", "ColumnType").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and IsDocumentLibrary ne 1`)();
    console.log("fileColumns", fileColumns);

    // Create an array of objects to store the columnName with there corresponding value
    const resultArrayThatContainstheColumnDetails = fileColumns.map((column) => {
      const columnName = column.ColumnName;
      const columnTpye = column.ColumnType;
      const columnRequired = column.IsRequired;
      const columnValue = fileItem.ListItemAllFields[columnName];

      return {
        label: columnName,
        value: columnValue !== undefined ? columnValue : null, // Handle missing fields
        type: columnTpye,
        required: columnRequired
      };
    });

    console.log("resultArrayThatContainstheColumnDetails", resultArrayThatContainstheColumnDetails)

    // Create the main container
    const mainContainer = document.createElement('div');
    mainContainer.className = 'main-containeruploadfile';
    const librarydiv = document.getElementById('files-container')
    const backButton = document.createElement('button')
    backButton.textContent = 'Close File Preview';
    const submitButton = document.createElement('button');
    const replaceButton = document.createElement('button');

    const uploadFileDiv = document.createElement('div');
    uploadFileDiv.id = "uploadFileDiv";
    uploadFileDiv.style.display = 'none'
    // input for upload file
    const uploadFileInput = document.createElement('input');
    uploadFileInput.className = "dynamic-input";
    uploadFileInput.type = "file";
    uploadFileInput.id = "fileInput";

    //  start
    const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
      const replaceButtonHide = document.getElementById('replaceButton');
      if (replaceButtonHide) {
        replaceButtonHide.style.display = 'none'
      }
      const file = event.target.files![0];
      if (file) {
        // selectedFile=file;
        uploadFile(file);
      }
    };

    const uploadFile = async (file: File) => {
      try {
        const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
        const uploadResult = await folder.files.addChunked(file.name, file);
        console.log("File uploaded successfully", uploadResult);

        // Generate the preview URL dynamically
        const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);

        previewFile(previewUrl);
      } catch (error) {
        console.error("Error uploading file:", error);
      }
    };

    const generatePreviewUrl = async (serverRelativeUrl: string) => {
      // Encode the file name and construct the preview URL
      const encodedFilePath = encodeURIComponent(serverRelativeUrl);

      // Example:
      // serverRelativeUrl = "/sites/IntranetUAT/test/DocumentLibraryInsideTest/Book.xlsx"
      const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
      const siteUrl = window.location.origin;

      const previewUrl = `${siteUrl}/sites/IntranetUAT/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
      // const previewUrl = `${siteUrl}/sites/SPFXDemo/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
      console.log("Generated Preview URL:", previewUrl);
      if (previewUrl) {
        console.log("enter herr")
        const deletebut = document.getElementById('closeCommand') as HTMLElement
        if (deletebut) {
          console.log(" here ", deletebut)
        }
      }
      return previewUrl;
    };

    const previewFile = async (previewUrl: string) => {
      try {
        console.log("Previewing file at URL:", previewUrl);
        const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
        const spinner = document.getElementById("spinner") as HTMLElement;

        // Show the spinner and hide the iframe initially
        spinner.style.display = "block";
        iframe.style.display = "none";
        iframe.src = previewUrl;

        // Add an onload event listener to the iframe
        iframe.onload = () => {
          console.log("Iframe has loaded");

          const checkAndHideButton = () => {
            try {
              const iframeDocument = iframe.contentDocument || iframe.contentWindow?.document;
              if (iframeDocument) {
                const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
                const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
                if (excelToolbar) {
                  excelToolbar.style.display = "none"
                }
                if (button) {
                  console.log("Hiding the OneUpCommandBar element");
                  button.style.display = "none";

                  // Hide the spinner and show the iframe after the button is hidden
                  spinner.style.display = "none";
                  iframe.style.display = "block";

                  // Exit the loop once the button is found and hidden
                } else {
                  console.log("OneUpCommandBar not found, rechecking...");
                }

                const helpbutton = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
                if (helpbutton) {
                  helpbutton.style.display = "none"
                }
              }
            } catch (error) {
              console.error("Error accessing iframe content:", error);
            }

            // Re-check after a short delay if the button wasn't found
            setTimeout(checkAndHideButton, 100);
          };

          // Start checking for the button
          checkAndHideButton();
        };
      } catch (error) {
        console.error("Error previewing file:", error);
      }

    };
    // end
    uploadFileInput.addEventListener('change', (event: any) =>
      handleFileChange(event))

    // Set Label For upload file
    const label = document.createElement("label");
    label.setAttribute("htmlFor", 'fileInput');
    label.textContent = 'Upload File';

    uploadFileDiv.appendChild(label);
    uploadFileDiv.appendChild(uploadFileInput);

    // Div for columns input
    const column1 = document.createElement('div');
    column1.className = 'column column1 p-3';

    // Add form to the first column
    const form = document.createElement('form');
    form.id = 'formSelector';
    const formHeading = document.createElement('h1');
    formHeading.textContent = 'Edit file';
    form.appendChild(formHeading);
    column1.appendChild(form);



    // Dynamically create input fields from the array
    resultArrayThatContainstheColumnDetails.forEach((field, index) => {

      const inputContainer = document.createElement("div");
      inputContainer.className = "input-container";
      // Create a label
      const label = document.createElement('label');
      label.textContent = field.label;
      label.setAttribute('htmlFor', `${field.label}`);
      // form.appendChild(label);

      // Add a red asterisk if the field is required
      if (field.required) {
        const asterisk = document.createElement("span");
        asterisk.textContent = " *";
        asterisk.style.color = "red";
        asterisk.style.fontWeight = "bold";
        label.appendChild(asterisk);
      }
      inputContainer.appendChild(label);

      let modifiedType = field.type.replace(/\s+/g, '').toLowerCase();

      // Create the input field based on its type
      let input: HTMLInputElement | null = null;
      if (
        modifiedType === "singlelineoftext"
        ||
        modifiedType === "multiplelineoftext"
        ||
        modifiedType === 'text'
      ) {
        input = document.createElement("input");
        input.type = "text";
      } else if (
        modifiedType === "number"
      ) {
        input = document.createElement("input");
        input.type = "number";
      } else if (
        modifiedType === "date&time"
      ) {
        input = document.createElement("input");
        input.type = "date";
      } else if (
        modifiedType === "yesorno"
      ) {
        input = document.createElement("input");
        input.type = "checkbox";
      }

      if (input) {
        input.className = "dynamic-input";
        input.value = field.value;
        input.id = field.label;
        input.name = field.label;
        input.required = field.required;
        inputContainer.appendChild(input);
        form.appendChild(inputContainer);
      }
    });
    // append the file input start
    form.appendChild(uploadFileDiv);
    // end

    // Create the second column
    const column2 = document.createElement('div');
    column2.className = 'column column2 p-3';

    // Create the spinner div
    const spinner = document.createElement('div');
    spinner.id = 'spinner';
    spinner.textContent = 'Loading...';
    spinner.style.display = 'none';



    replaceButton.type = "submit";
    replaceButton.id = "replaceButton";
    replaceButton.addEventListener('click', (event) => {
      event.preventDefault();
      console.log("replace button called");
      clickedReplace = true;
      const uploadFile = document.getElementById('uploadFileDiv')
      const iframe = document.getElementById('filePreview');
      if (uploadFile) {
        uploadFile.style.display = 'block';
      }
      if (iframe) {
        iframe.style.display = 'none';
      }
      // myRequest(null,null,null);
    })
    replaceButton.textContent = "Replace"


    // Add heading to the second column
    const column2Heading = document.createElement('h1');
    column2Heading.textContent = 'File Preview';
    column2.appendChild(column2Heading);
    column2.appendChild(replaceButton);

    // append the spinner start
    column2.appendChild(spinner);
    // end

    const previewfileframe = document.createElement('iframe')
    previewfileframe.id = 'filePreview'
    previewfileframe.style.width = '930px'
    previewfileframe.style.height = '500px'

    const segments = filePath.split('/');
    // extarct the current entity start
    const currentSubsite = segments[3];
    // end
    // Find the index of 'sites'
    const sitesIndex = segments.indexOf('sites');

    // If 'sites' is found and there are enough segments after it
    let myactualdoclib
    if (sitesIndex !== -1 && segments.length > sitesIndex + 3) {
      myactualdoclib = segments[sitesIndex + 3];
      // console.log(myactualdoclib , "myactualdoclib")
      // return segments[sitesIndex + 3];  // The document library is the 4th segment after 'sites'
    }

    // Extract the parent folder correctly
    const parentFolder = filePath.substring(0, filePath.lastIndexOf('/'));
    console.log(parentFolder, "parentFolder");

    // Correctly encode the parent folder
    const encodedParentFolder = encodeURIComponent(parentFolder);

    // Get the base site URL
    const siteUrl = window.location.origin;
    console.log(siteUrl, "siteUrl");

    const previewUrl = `${siteUrl}/sites/IntranetUAT/${currentSubsite}/${myactualdoclib}/Forms/AllItems.aspx?id=${filePath}&parent=${encodedParentFolder}`;

    if (previewUrl) {
      previewfileframe.src = previewUrl;
      column2.appendChild(previewfileframe);
    }

    // Append columns to the main container
    mainContainer.appendChild(column1);
    mainContainer.appendChild(column2);


    // Submit Button property
    submitButton.type = "submit";


    submitButton.addEventListener('click', async (event) => {
      event.preventDefault();
      console.log("submit button called");

      // Extract the last part after the last '/'
      const fileName: any = filePath.substring(filePath.lastIndexOf('/') + 1);
      // alert(fileName);
      // console.log("fileName",fileName);
      // Extract the rest of the path
      const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
      console.log("folderPath", folderPath);

      const formSelector = document.getElementById("formSelector") as HTMLFormElement;
      if (!formSelector.checkValidity()) {
        checkValidation(`Fill mandatory fields`)
        return;
      }

      // Prepare the payload for SharePoint dynamically
      const inputs = document.querySelectorAll('.dynamic-input');
      const payload: any = {};

      inputs.forEach((input) => {
        const inputElement = input as HTMLInputElement;
        const fieldName = inputElement.id;
        if (!fieldName) return; // Skip if field name is invalid

        if (inputElement.type === "checkbox") {
          // console.log("fieldName",fieldName.includes(' '));
          payload[fieldName] = inputElement.checked;
        } else if (inputElement.type !== "file") {
          if (inputElement.value === "") {
            console.log("skip");
          } else {
            // if(fieldName.includes(' '))
            // console.log("fieldName",fieldName.includes(' '));
            payload[fieldName] = inputElement.value;
          }

        }
      });

      if (clickedReplace) {
        const fileInput = document.getElementById('fileInput') as HTMLInputElement;
        const selectedFile = fileInput?.files?.[0];

        if (!selectedFile) {
          console.error("No file selected.");
          // alert("Please select the file...");
          checkValidation(`Fill mandatory fields`)
          return;
        }


        console.log("selectedFile.name", selectedFile.name)
        const fileExtensionOfSelectedFile = selectedFile.name.split('.').pop();
        const fileExtensionOfOldFile = fileName.split('.').pop();
        // for same file extension
        if (fileExtensionOfSelectedFile === fileExtensionOfOldFile) {
          // alert("Same file extension");
          const file = web.getFileByServerRelativePath(filePath);
          await file.setContentChunked(selectedFile);
          if (file.exists) {
            const fileToUpdate = await file.getItem();
            const uploadResult = await fileToUpdate.update(payload);
            console.log("uploadResult", uploadResult);
          }
          showReplaceMessage('File replaced successfully.');
          myRequest(null, null, null);
        } else {
          // alert("file extension are not same");
          const folderInWhichWeUploadTheFile = web.getFolderByServerRelativePath(folderPath);
          const uploadResult = await folderInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
          const listItem = await uploadResult.file.getItem();
          (payload as any).Status = "Pending";

          const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
          const siteUrl = window.location.origin;
          const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
          const previewUrl = `${siteUrl}/sites/IntranetUAT/${siteName}/${documentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

          await listItem.update(payload);

          const newItem = await sp.web.lists.getByTitle(`DMS${siteName}FileMaster`).items.add({
            FileName: String(uploadResult.data.Name),
            FileSize: String(uploadResult.data.Length),
            FileVersion: String(uploadResult.data.MajorVersion),
            CurrentFolderPath: String(folderPath),
            FileUID: String(uploadResult.data.UniqueId),
            CurrentUser: String(currentUserEmailRef.current),
            SiteID: String(siteId),
            Status: "Pending",
            FilePreviewURL: String(previewUrl),
            DocumentLibraryName: String(documentLibrary),
            SiteName: String(siteName),
            MyRequest: true
          });


          const AddIteminDMSFileApprovalList = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
            SiteName: String(siteName),
            DocumentLibraryName: String(documentLibrary),
            RequestedBy: String(currentUserEmailRef.current),
            FileName: String(uploadResult.data.Name),
            FileUID: String(uploadResult.data.UniqueId),
            //  FilePreviewUrl: String(previewUrl),
            Status: String('Pending'),
            FolderPath: String(folderPath),
            ApproveAction: String('Submitted'),
            ApprovedLevel: 1
          })

          // delete the file from the document library
          const deletedfile = await web.getFileById(fileId).delete();
          console.log("deletedfile", deletedfile);
          // Get the file details
          const fetchData = await sp.web.lists.getByTitle(`DMS${siteName}FileMaster`).items.select("ID", "FileName").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and CurrentUser eq '${currentUserEmailRef.current}'  and FileUID eq '${fileId}'`)();
          console.log("fetchData", fetchData);
          await sp.web.lists.getByTitle(`DMS${siteName}FileMaster`).items.getById(fetchData[0].ID).delete();
          console.log(`file has been deleted successfully.`);

          const fetchDatafromapprovalist = await sp.web.lists.getByTitle(`DMSFileApprovalList`).items.select("ID", "FileName").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and FileUID eq '${fileId}'`)();

          await sp.web.lists.getByTitle(`DMSFileApprovalList`).items.getById(fetchDatafromapprovalist[0].ID).delete();
          console.log(`file has been deleted successfully fetchDatafromapprovalist.`);

        }



      } else {
        // Without replacing the file only metadeta update
        const file = web.getFileByServerRelativePath(filePath);
        if (file.exists) {
          const fileToUpdate = await file.getItem();
          const uploadResult = await fileToUpdate.update(payload);
          console.log("uploadResult", uploadResult);
        }
        showReplaceMessage('File updated successfully.')
        myRequest(null, null, null);
      }
      // const file =sp.web.getFileByServerRelativePath(filePath);


      // myRequest(null,null,null);
    })


    submitButton.textContent = "Submit"
    form.appendChild(submitButton);

    librarydiv.innerHTML = "";
    mainContainer.appendChild(backButton)
    librarydiv.appendChild(mainContainer)



  }
  // end
  return (
    <div id="wrapper" ref={elementRef}>
      <div
        className="app-menu"
        id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>
      <div className="content-page">
        <HorizontalNavbar />
        <div className="content" style={{ marginLeft: `${!useHide ? '240px' : '80px'}`, marginTop: '0.8rem' }}>

          <div className="container-fluid  paddb">
            {activeComponent === "" ? (
              <div className=" dmsmaincontainer">
                {showWorkflow && (
                  <div id="workflowdiv">
                    {/* <ManageWorkFlow
                      OthProps={propsForManageWorkFlow}
                      onReturnToMain={handleReturnToMainFromManageWorkFlow}
                    /> */}
                    {/* To this: */}
<ManageWorkFlow
  context={props.context} // Matches WebPartContext requirement
  siteTitle={propsForManageWorkFlow.SiteTitle}
  siteUrl={props.siteUrl}
  documentLibraryName={propsForManageWorkFlow.DocumentLibraryName}
  file={null} // Replace with your file context object if available
  onClose={handleReturnToMainFromManageWorkFlow} // Your interface specifies onClose, not onReturnToMain
/>
                  </div>
                )}
                {showfolderpermission && (
                  <div id="showfolderpermission">
                    <ManageFolderPermission
                      OthProps={managePermissionProps}
                      onReturnToMain={handleReturnToMainFromManageWorkFlow}
                    />
                  </div>

                )}
                {/* Start Code Update by Amjad */}
                <div className="row">
                  <div className="col-lg-6">
                    <h4 className="page-title fw-bold mb-1 font-20">Dossier</h4>
                    <ol className="breadcrumb m-0">
                      {" "}
                      <li className="breadcrumb-item">Home</li>
                      <li className="breadcrumb-item active">Dossier</li>
                    </ol>
                  </div>


                  <div style={{ display: 'flex', justifyContent: 'end', gap: '5px' }} className="col-lg-6 newbutton">
                    <div id="hidegidvewlistviewbutton" className="view-buttons mt-2">
                      <button type="button" className="btn me-1 btngridview mt-0 grid-view active"
                        onClick={(event: any = null, siteIdToUpdate: string = null) => myRequest(event)}>
                        <a className="listviewfonticon">
                          <FontAwesomeIcon style={{ color: "black" }} icon={faTableCells} /> </a>Grid View
                      </button>
                      <button type="button" className="btn btnlistview list-view mt-0" onClick={(event: any) => MyrequestshowListView('ListViewComponent')}>
                        <a className="listviewfonticon">
                          <FontAwesomeIcon style={{ color: "black" }} icon={faListSquares} />
                        </a>
                        List View
                      </button>

                    </div>
                    {displayuploadfileandcreatefolder && (
                      <div id="createuploadfilecont" className="createuploadfilecont mt-2">
                        <button
                          className="mybutton1 mt-0"
                          id="createFileButton"
                          onClick={() => handleButtonClickShow("UploadFile")}
                        >
                          + Upload File
                        </button>

                        <button
                          className="mybutton2 mt-0"
                          id="createFileButton2"
                          onClick={() => handleButtonClickShow("CreateFolder")}
                        >
                          + Create Folder
                        </button>


                      </div>
                    )

                    }

                    {showMyfavButtons && (<div id="hidegidvewlistviewbutton2" className="view-buttons mt-2">
                      <button className="btn btngridview grid-view active"
                        onClick={(e) => myFavorite(e)}>
                        <a className="listviewfonticon">
                          <FontAwesomeIcon style={{ color: "black" }} icon={faTableCells} /> </a>Grid View
                      </button>
                      <button type="button" className="btn btnlistview list-view" onClick={(event: any) => MyrequestshowListView('ListViewComponent')}>
                        <a className="listviewfonticon">
                          <FontAwesomeIcon style={{ color: "black" }} icon={faListSquares} />
                        </a>
                        List View
                      </button>
                    </div>)
                    }
                  </div>

                </div>
                {/* End Code Update by Amjad */}
                <div className="mainsidebardms">
                  <div className="sidebardms">
                    <button
                      id="Myrequestbutton"
                      className={`sidebardmsButton ${activeButton === "MyRequest" ? "active" : ""
                        }`}
                      // onClick={() => handleClick('MyRequest')}
                      onClick={
                        (event) => {

                          myRequest(event);
                          handleShowContent(event)
                        }
                      }
                    >
                      <span className="sidebarIcon">
                        {/* <FontAwesomeIcon icon={faList} /> */}
                        <img className="sidebariconssmall" src={listicon}></img>
                      </span>
                      <span className="sidebarText">My Request</span>
                    </button>

                    <button
                      className={`sidebardmsButton ${activeButton === "MyFavourite" ? "active" : ""
                        }`}
                      onClick={(event) => {
                        myFavorite(event);
                        handleShowContent(event);
                      }}
                    >
                      <span className="sidebarIcon">
                        <img className="sidebariconssmall" src={starticon}></img>
                        {/* <FontAwesomeIcon icon={faStarRegular} /> */}
                      </span>
                      <span className="sidebarText">My Favourite</span>
                    </button>

                    <button
                      className={`sidebardmsButton ${activeButton === "MyFolder" ? "active" : ""
                        }`}
                      onClick={(event) => {
                        mycreatedfolders(event);
                        handleShowContent(event)
                      }}
                    >
                      <span className="sidebarIcon">
                        <img className="sidebariconssmall" src={foldericon}></img>
                        {/* <FontAwesomeIcon icon={faFolderRegular} /> */}
                      </span>
                      <span className="sidebarText">My Folder</span>
                    </button>

                    <button
                      className={`sidebardmsButton ${activeButton === "ShareWithOther" ? "active" : ""
                        }`}
                      onClick={(e) => {
                        ShareWithOther(e);

                        handleShowContent(e)
                      }}
                    >
                      <span className="sidebarIcon">
                        {/* <FontAwesomeIcon icon={faShareAlt} /> */}
                        <img className="sidebariconssmall" src={sharewithothericon}></img>
                      </span>
                      <span className="sidebarText">Shared with Others</span>
                    </button>

                    <button
                      onClick={(e) => { ShareWithMe(e); handleShowContent(e) }}
                      className={`sidebardmsButton ${activeButton === "ShareWithMe" ? "active" : ""
                        }`}
                    >
                      <span className="sidebarIcon">
                        <img className="sidebariconssmall" src={sharewithmeicon}></img>
                        {/* <FontAwesomeIcon icon={faShareAlt} /> */}
                      </span>
                      <span className="sidebarText">Shared with me</span>
                    </button>

                    <button
                      onClick={(e) => { Recyclebin(e); handleShowContent(e) }}
                      className={`sidebardmsButton ${activeButton === "ShareWithMe" ? "active" : ""
                        }`}
                    >
                      <span className="sidebarIcon">
                        <img className="sidebariconssmall" src={recyclebin}></img>
                        {/* <FontAwesomeIcon icon={faShareAlt} /> */}
                      </span>
                      <span className="sidebarText">Recycle Bin</span>
                    </button>
                  </div>
                  <div style={{ position: 'sticky', top: '100px' }} className="is-sticky">  <div id="folderContainer2"></div></div>
                </div>
                <div className="librarydata">
                  {showDeletepopup && (
                    <div className="popup">This is a small popup!</div>
                  )}
                  {/* Start Code Update by Amjad */}
                  <div className="row">
                    <div className="col-xl-7">
                      <div
                        id="selectedText"
                        style={{
                          display: "none",
                        }}
                      >
                        {selectedText ? (
                          <h6 className="mb-1 fw-bold text-dark header-title"
                            style={{
                              color: "black",
                              marginBottom: "0px",
                              fontSize: "16px",
                            }}
                          >
                            {selectedText}
                          </h6>
                        ) : (
                          <p className="sub-header font-14"></p>
                        )}
                        {dynamicContent && (
                          <p className="sub-header font-14" style={{ color: "#6c757d" }}>{dynamicContent}</p>
                        )}
                      </div>
                      <div
                        id="breadcrumb"
                        style={{
                          display: "none",
                        }}
                      ></div>
                    </div>
                    <div className="col-xl-5">
                      <div className="search-container position-relative">
                        <input
                          id="searchinput"
                          type="text"
                          className="search-input"
                          placeholder="Search files..."
                        />
                        <a className="searchbutton" onClick={searchFiles}>
                          <img
                            src={require("../../assets/searchicon.png")}
                            alt="Search"
                            className="search-icon"
                          />
                        </a>
                      </div>

                    </div>
                  </div>
                  {/* End Code Update by Amjad */}

                  <div className="row">
                    <DMSSearchFilter 
                         properties={['Title','Created','Created By','Modified','Modified By']} 
                         value={""} 
                         onSubmit={(property,op,value)=>{ console.log(property); runSearch(searchQueryPath,"Title:"+value ) }}  />

                  </div>
                  <DMSResultsGrid ViewMode={ResultsViewMode.Grid} results={searchDocResults} />
                  <div id='files-container' style={{ 'display': 'none' }}></div>

                </div>

                <div className="Manageworkflow">

                </div>
              </div>
            ) : (
              <div>
                {activeComponent === "UploadFile" && (
                  <UploadFile
                  sp={sp}
                    currentfolderpath={{
                      "Entity": currentEntity,
                      "Entityurl": currentEntityURL,
                      "siteID": currentsiteID,
                      "Devision": currentDevision,
                      "Department": currentDepartment,
                      "DocumentLibrary": currentDocumentLibrary,
                      "Folder": currentFolder,
                      "folderpath": currentfolderpath
                    }}
                    onReturnToMain={handleReturnToMain}
                  />
                )}
                {activeComponent === "CreateFolder" && (
                  // <CreateFolder OthProps={{
                  //   "Entity": currentEntity,
                  //   "Entityurl": currentEntityURL,
                  //   "siteID": currentsiteID,
                  //   "Devision": currentDevision,
                  //   "Department": currentDepartment,
                  //   "DocumentLibrary": currentDocumentLibrary,
                  //   "Folder": currentFolder,
                  //   "folderpath": currentfolderpath
                  // }}
                  //   onReturnToMain={handleReturnToMain} />
                  <CreateFolder 
  context={props.context} // Required by spfi().using(SPFx(context))
  OthProps={{
    "Entity": currentEntity,
    "Entityurl": currentEntityURL,
    "siteID": currentsiteID,
    "Devision": currentDevision,
    "Department": currentDepartment,
    "DocumentLibrary": currentDocumentLibrary,
    "Folder": currentFolder,
    "folderpath": currentfolderpath
  }}
  onCloseForm={handleReturnToMain} // Bound to onCloseForm to handle the cancel click
/>
                )}


              </div>
            )}
          </div>
        </div>
      </div>
    </div>



  );
};



const DMSSearch: React.FC<IEssadmsMainProps> = (props) => {
  return (
    <Provider>
      <ArgPoc props={props} />
    </Provider>
  );
};

export default DMSSearch;
