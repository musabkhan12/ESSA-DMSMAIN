// @ts-ignore
import * as React from "react";
import { useEffect , useState , useRef, useMemo } from "react";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import './uploadfilecss'
import * as XLSX from 'xlsx';
import { Web } from "@pnp/sp/webs";
import './uploadfilecss'
import Swal from 'sweetalert2';
import { AssignFrom } from "@pnp/core";
let info = require('../assets/infon.png')
let back = require('../assets/backnew.png')
let showbulkupload : any;
let IsApproval : any
let status :any;
let buttontext :any = 'Submit'
interface UploadFileProps {
  currentfolderpath: { [key: string]: string };
  onReturnToMain: () => void;
  sp: SPFI;
  // myRequest: (event: React.MouseEvent<HTMLButtonElement>) => void;
}

let previewURLN;
const submitButton=document.createElement('button');
const img = document.createElement("img");
img.src = require("../assets/submit-new1.png");
img.alt = "Create";
    // optional
  // optional
  submitButton.style.borderRadius='30px';
img.style.marginRight = "6px";
submitButton.appendChild(img);
// submitButton.textContent= buttontext
submitButton.id="submitBtn";
submitButton.type="submit";
// submitButton.style.display='none';
submitButton.disabled=true

const UploadFile: React.FC<UploadFileProps> = ({ currentfolderpath , onReturnToMain, sp  }) => {
  // const sp: SPFI = getSP();
  // let locationPath=window.location.pathname.match(/\/sites\/[^\/]+/)[0];
   let rawHash = window.location.hash.substring(1); // Remove the '#'
let sitePart = rawHash.split('/')[0]; // Get the first segment before any '/'
 
let locationPath = sitePart.startsWith('sites/')
    ? `/${sitePart}`
    : `/sites/${sitePart}`;
 
console.log(locationPath); // "/sites/AlRostmaniSpfx2"
  // check whether folder is private or public and save state

  const [showBulkUpload, setShowBulkUpload] = useState<boolean | null>(null);
 const [isFinalUploading,setIsFinalUploading] = useState(false);
  const [isChecked, setIsChecked] = useState(false);
  const [selectedIndex, setSelectedIndex] = useState(null);
  // const checkfolderprivace = async() =>{
  //   const folderItems = await sp.web.lists.getByTitle("DMSPreviewFormMaster")
  //   .items.filter(`DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' and SiteName eq '${currentfolderpath.Entity}' and IsDocumentLibrary eq 1`).select("IsApproval","IsPrivate")();
  //   console.log("folderItems",folderItems);
  //   showbulkupload = folderItems[0].IsApproval;
  //   setShowBulkUpload(folderItems[0].IsApproval)

  //   IsApproval=folderItems[0].IsApproval;

  //   console.log('currentfolderpath' , currentfolderpath)
  // }
//  const checkfolderprivace = async () => {
//     try {
//         // 1. Extract the Root Site URL (https://officeindia.sharepoint.com/sites/AlRostmaniSpfx2)
//         // This regex looks for the pattern /sites/AnyName and ignores whatever follows
//         const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
        
//         if (!rootSiteMatch) {
//             console.error("Could not determine Root Site URL from Entityurl");
//             return;
//         }

//         const rootSiteUrl = rootSiteMatch[0]; 
//         console.log("Targeting Root Site for Config List:", rootSiteUrl);

//         // 2. Initialize the Web object for the Root Site
//         const remoteWeb = Web([sp.web, rootSiteUrl]);

//         // 3. Query the list
//         const folderItems = await remoteWeb.lists
//             .getByTitle("DMSPreviewFormMaster")
//             .items
//             .filter(`DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' and SiteName eq '${currentfolderpath.Entity}' and IsDocumentLibrary eq 1`)
//             .select("IsApproval", "IsPrivate")();

//         console.log("Config Result:", folderItems);

//         if (folderItems && folderItems.length > 0) {
//             setShowBulkUpload(folderItems[0].IsApproval);
//             IsApproval = folderItems[0].IsApproval;
//             showbulkupload = folderItems[0].IsApproval;
//         }
//     } catch (error) {
//         console.error("Error in checkfolderprivace:", error);
//     }
// };
const getRootLibraryName = () => {
    // Logic: If subfolder exists, take the first part. Otherwise, use DocumentLibrary.
    if (currentfolderpath.Folder && currentfolderpath.Folder.includes('/')) {
        return currentfolderpath.Folder.split('/')[0];
    }
    return currentfolderpath.DocumentLibrary || "";
};
const checkfolderprivace = async () => {
    try {
        // 1. Extract the Root Site URL
        const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
        if (!rootSiteMatch) return;

        const rootSiteUrl = rootSiteMatch[0];

        // --- FIX START ---
        // 2. Determine the actual Library Name
        // If it's a subfolder, currentfolderpath.Folder might be "Testsrs3/Testsrs3sub"
        // We need the part before the first "/"
        const libraryName = getRootLibraryName();
        // --- FIX END ---

        const remoteWeb = Web([sp.web, rootSiteUrl]);

        // 3. Query the list using the extracted libraryName
        const folderItems = await remoteWeb.lists
            .getByTitle("DMSPreviewFormMaster")
            .items
            .filter(`DocumentLibraryName eq '${libraryName}' and SiteName eq '${currentfolderpath.Entity}' and IsDocumentLibrary eq 1`)
            .select("IsApproval", "IsPrivate")();

        if (folderItems && folderItems.length > 0) {
            setShowBulkUpload(folderItems[0].IsApproval);
            // Note: Ensure these variables are properly scoped/defined in your component
            IsApproval = folderItems[0].IsApproval;
            showbulkupload = folderItems[0].IsApproval;
        }
    } catch (error) {
        console.error("Error in checkfolderprivace:", error);
    }
};
  useEffect(() => {
  checkfolderprivace();
  }, []);


const [data, setData] = useState({
  Entity: '',
  Entityurl: '',
  siteID: '',
  Devision: '',
  Department: '',
  DocumentLibrary: '',
  Folder: '',
  folderpath: '',
});

const [state, setState] = useState({});

const currentUserEmailRef = useRef('');

const getcurrentuseremail = async()=>{
  const userdata = await sp.web.currentUser();
  currentUserEmailRef.current = userdata.Email;
  // console.log(currentUserEmailRef.current, "currentuser")
 }

useEffect(() => {
  getcurrentuseremail()
  setData({...data , ...currentfolderpath});

}, []);


console.log(data, "data"  )
console.log(data.Entity, "entity"  )
const SubsiteID = data.siteID
const currentPath =data.folderpath; 
const documentLibraryName  = data.DocumentLibrary;
console.log("documentLibraryName" , documentLibraryName)


  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  setIsUploading(true);
  const file = event.target.files![0];

  
  if (file) {
    // Directly upload the file without any validation
    uploadFile(file);
  }else{
    console.log("no file selected")
    const submitButton = document.getElementById("submitBtn") as HTMLButtonElement;
    if(submitButton){
      submitButton.disabled=true;
    }
  }
};


  // const uploadFile = async (file: File) => {
  //   try {
  //     const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
  //     const files = await folder.files();
  
  //     const originalFileName = file.name;
  //     const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
  //     const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
  //     let uniqueFileName = originalFileName;
  //     let counter = 1;
  
  //     while (files.some(f => f.Name === uniqueFileName)) {
  //       uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
  //       counter++;
  //     }
  
  //     // Encode the file name to handle special characters
  //     const encodedFileName = encodeURIComponent(uniqueFileName);
  
  //     const uploadResult = await folder.files.addChunked(encodedFileName, file);
  //     console.log("File uploaded successfully", uploadResult);
  
  //     // Generate the preview URL dynamically
  //     const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
  //     const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
  //     submitBtn.disabled = false; // Enable the button
  //     setIsUploading(false);
  //     previewFile(previewUrl,"singleUpload");
  //   } catch (error) {
  //     console.error("Error uploading file:", error);
  //   }
  // };
  
const uploadFile = async (file: File) => {
  try {
    // 1. Extract the Root Site Collection URL
    // From: https://officeindia.sharepoint.com/sites/AlRostmaniSpfx2/N27JAN
    // To: https://officeindia.sharepoint.com/sites/AlRostmaniSpfx2
    const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
    if (!rootSiteMatch) {
      console.error("Could not determine Root Site URL");
      return;
    }
    const rootSiteUrl = rootSiteMatch[0]; 

    // 2. Initialize the Web object targeting the Root Site (AlRostmaniSpfx2)
    // const rootWeb = Web([sp.web, rootSiteUrl]);
    const rootWeb = Web(rootSiteUrl).using(AssignFrom(sp.web as any));
    
    // 3. Target the Orphan library at the root site
    const folder = rootWeb.getFolderByServerRelativePath('DMSOrphanDocs');
    const files = await folder.files();

    const originalFileName = file.name;
    const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
    const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
    let uniqueFileName = originalFileName;
    let counter = 1;

    while (files.some(f => f.Name === uniqueFileName)) {
      uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
      counter++;
    }

    // 4. Upload to the Root Site's library
    const uploadResult = await folder.files.addChunked(uniqueFileName, file);
    
    // 5. Generate the Preview URL
    // We must ensure the preview path points to the Root Site where DMSOrphanDocs exists
    const serverRelativeUrl = uploadResult.data.ServerRelativeUrl;
    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteOrigin = window.location.origin;
    
    // Construct path: https://domain.sharepoint.com/sites/AlRostmaniSpfx2/DMSOrphanDocs/Forms/AllItems.aspx...
    const previewUrl = `${siteOrigin}${rootSiteMatch[0].replace(siteOrigin, "")}/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodeURIComponent(serverRelativeUrl)}&parent=${encodeURIComponent(parentFolder)}`;

    console.log("Cross-site Preview URL:", previewUrl);

    const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
    if (submitBtn) submitBtn.disabled = false;
    
    setIsUploading(false);
    previewFile(previewUrl, "singleUpload");
  } catch (error) {
    console.error("Error in cross-site uploadFile:", error);
    setIsUploading(false);
  }
};
  const generatePreviewUrl = async (serverRelativeUrl: string) => {
    // Encode the file name and construct the preview URL
    const encodedFilePath = encodeURIComponent(serverRelativeUrl);
    
    // Example: 

    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;

    //  const previewUrl = `${siteUrl}/sites/AlRostmani/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    //  const previewUrl = `${siteUrl}/sites/AlRostmanispfx2/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
      const previewUrl = `${siteUrl}${locationPath}/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    console.log("Generated Preview URL:", previewUrl);
   if(previewUrl){
    console.log("enter herr")
    const deletebut = document.getElementById('closeCommand') as HTMLElement
    if(deletebut){
      console.log(" here " , deletebut)
    }
   }
    return previewUrl;
  };


const previewFile = async (previewUrl: string,flag:string) => {
    try {
      console.log("Previewing file at URL:", previewUrl);
      const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
      const spinner = document.getElementById("spinner") as HTMLElement;
      const submitButton = document.getElementById("submitBtn") as HTMLButtonElement;
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
              if(excelToolbar){
                excelToolbar.style.display= "none"
              }
              if (button) {
                console.log("Hiding the OneUpCommandBar element");
                button.style.display = "none";
  

                spinner.style.display = "none";
                iframe.style.display = "block"; 
                const mainDiv = iframeDocument.getElementById("ModalFocusTrapZone3") as HTMLElement
                if(mainDiv){
                  mainDiv.style.background='white'
                }
                // Ensure submit button is shown only once in case of single upload
                // if (flag === "singleUpload" && submitButton && submitButton.style.display !== "block") {
                //     submitButton.style.display = "block";
                // }
                if(flag === "singleUpload" && submitButton && submitButton.disabled !== false){
                  console.log("preview for single upload")
                    submitButton.disabled = false
                }
                
              } else {
                console.log("OneUpCommandBar not found, rechecking...");
              }
              
              const helpbutton = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement; 
              if(helpbutton){
                helpbutton.style.display = "none"
              }
            }
          } catch (error) {
            console.error("Error accessing iframe content:", error);
          }
  
          // Stop rechecking once the preview is fully loaded
          if (spinner.style.display === "none") return;
          setTimeout(checkAndHideButton, 100);
        };
  

        checkAndHideButton();
      };
    } catch (error) {
      console.error("Error previewing file:", error);
    }

  };


  //  handle bulk file

  
const [uploadedFiles, setUploadedFiles] = useState<{ name: string; url: string,file:File }[]>([]); // Store uploaded files for preview
const [isUploading, setIsUploading] = useState(false);

const handlebulkFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
  setIsUploading(true);
  
  const files = event.target.files; // Get all selected files

  if (!files || files.length === 0) return; // Exit if no files are selected

  const validFiles: File[] = []; // Store valid files

  for (let i = 0; i < files.length; i++) {
    validFiles.push(files[i]); // Add all selected files without validation
  }

  // If there are valid files, proceed with upload
  if (validFiles.length > 0) {
    if (validFiles.length === 1) {
      const fileList = new DataTransfer();
      validFiles.forEach((file) => fileList.items.add(file));
      // uploadFile(validFiles[0]); // Upload single file
      bulkUploadFile(fileList.files); 
    } else {
      const fileList = new DataTransfer();
      validFiles.forEach((file) => fileList.items.add(file));
      bulkUploadFile(fileList.files); 
      // Upload multiple files
      //  bulkUploadFile(validFiles); // Upload multiple files
    }
  }

  // Clear input field to allow re-selection of the same files
  event.target.value = '';
};


// const bulkUploadFile = async (files: FileList) => {
//   console.log("Bulk uploading files:", files);
  
//   const uploadedFilesList: { name: string; url: string,file:File }[] = [];

//   for (let i = 0; i < files.length; i++) {
//     try {
//       const file = files[i];
//       console.log(`Uploading file.name: ${file.name}`);
//       console.log(`Uploading file: ${file.name}`);
//       const folder = sp.web.getFolderByServerRelativePath("DMSOrphanDocs");

//       const uploadResult = await folder.files.addChunked(file.name, file);
//       console.log(`File ${file.name} uploaded successfully`, uploadResult);

//       // Generate the preview URL dynamically
//       const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);

//       uploadedFilesList.push({ name: file.name, url: previewUrl,file:file });
//     } catch (error) {
//       console.error(`Error uploading file ${files[i].name}:`, error);
//     }
//   }

//   // setUploadedFiles(uploadedFilesList); // Update state with all uploaded files
//   setUploadedFiles((prevFiles) => [...prevFiles, ...uploadedFilesList]);
//   setIsUploading(false);
//   const submitButton = document.getElementById("submitBtn2") as HTMLButtonElement;
//   if (submitButton) submitButton.style.display = "block";
// };

const bulkUploadFile = async (files: FileList) => {
  console.log("Bulk uploading files:", files);
  setIsUploading(true); // Ensure loading state is active

  const uploadedFilesList: { name: string; url: string; file: File }[] = [];

  try {
    // 1. Extract Root Site URL (Same logic as uploadFile)
    const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
    if (!rootSiteMatch) {
      console.error("Could not determine Root Site URL");
      setIsUploading(false);
      return;
    }
    const rootSiteUrl = rootSiteMatch[0];
    const siteOrigin = window.location.origin;
    const rootSiteRelativePath = rootSiteMatch[0].replace(siteOrigin, "");

    // 2. Initialize the Root Web object
    const rootWeb = Web(rootSiteUrl).using(AssignFrom(sp.web as any));
    const folder = rootWeb.getFolderByServerRelativePath('DMSOrphanDocs');
    
    // Get existing files once to handle naming collisions for the whole batch
    const existingFiles = await folder.files();

    for (let i = 0; i < files.length; i++) {
      try {
        const file = files[i];
        
        // 3. Handle Duplicate Naming Logic
        const originalFileName = file.name;
        const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
        const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
        let uniqueFileName = originalFileName;
        let counter = 1;

        // Check against both existing files in SP and files already processed in this loop
        while (existingFiles.some(f => f.Name === uniqueFileName)) {
          uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
          counter++;
        }

        // 4. Upload to Root Site
        const uploadResult = await folder.files.addChunked(uniqueFileName, file);
        const serverRelativeUrl = uploadResult.data.ServerRelativeUrl;
        const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));

        // 5. Generate Preview URL (Matching your uploadFile logic)
        const previewUrl = `${siteOrigin}${rootSiteRelativePath}/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodeURIComponent(serverRelativeUrl)}&parent=${encodeURIComponent(parentFolder)}`;

        uploadedFilesList.push({ name: uniqueFileName, url: previewUrl, file: file });
        
        // Add to local list so next file in loop doesn't collide with this one
        existingFiles.push({ Name: uniqueFileName } as any);

      } catch (error) {
        console.error(`Error uploading file ${files[i].name}:`, error);
      }
    }

    // Update State
    setUploadedFiles((prevFiles) => [...prevFiles, ...uploadedFilesList]);
    
    const submitButton = document.getElementById("submitBtn2") as HTMLButtonElement;
    if (submitButton) submitButton.style.display = "block";

  } catch (globalError) {
    console.error("Critical error in bulk upload:", globalError);
  } finally {
    setIsUploading(false);
  }
};

const handlePreview = (previewUrl: string) => {
  previewFile(previewUrl,"bulkUpload"); // Call your preview function
};
  



  
React.useEffect(()=>{
  // const  loadFormOptions = async ()=> {
  //   try {
  //     const documentLibraryFields=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName","ColumnType","IsRequired","IsRename")
  //     .filter(
  //           `SiteName eq '${currentfolderpath.Entity}' 
  //           and DocumentLibraryName eq '${currentfolderpath.DocumentLibrary}' 
  //           and AddorRemoveThisColumn eq  'Add To Library' and IsInProgress eq 0`)();

  //     console.log("Document Library Fields",documentLibraryFields);
  //     // end

  //     const formSelector = document.getElementById("formSelector");
     
  //     const uploadFileDiv=document.createElement('div');

  //     const createElement=(fieldName:string,type:string,required:boolean,IsRename:string)=>{
  //           let fName=fieldName
  //           if(IsRename!== null){
  //             fName=IsRename
  //           }
  //           const inputContainer = document.createElement("div"); 
  //           inputContainer.className = "input-container";
             
  //           // Create and set label
  //           const label = document.createElement("label");
  //           label.setAttribute("htmlFor", fieldName);
  //           // label.textContent = fieldName;
  //           // label.textContent = fName;
  //            // Append an asterisk if the field is required 
  //             if (required) {
  //               const asterisk = document.createElement("span");
  //               asterisk.textContent = " *";
  //               asterisk.style.color = "red";
  //               label.textContent = fName;
  //               label.appendChild(asterisk); 
  //             } else {
  //               label.textContent = fName;
  //             }
  //           inputContainer.appendChild(label);
    
  //           let inputElement: HTMLInputElement | null = null;
    
  //           // Dynamically create the input field based on FieldType
  //           let modifiedType = type?.replace(/\s+/g, '').toLowerCase();
  //           console.log("modifiedType",modifiedType);

  //           if (
  //               modifiedType === "singlelineoftext"
  //               || 
  //               modifiedType === "multiplelineoftext" 
  //               || 
  //               modifiedType === 'text'
  //           ){
  //             inputElement = document.createElement("input");
  //             inputElement.type = "text";
  //           } else if (
  //             modifiedType === "number"
  //           ) {
  //             inputElement = document.createElement("input");
  //             inputElement.type = "number";
  //           } else if (
  //             modifiedType === "date&time"
  //           ) {
  //             inputElement = document.createElement("input");
  //             inputElement.type = "date";
  //           } else if (
  //             modifiedType === "yesorno"
  //           ) {
  //             inputElement = document.createElement("input");
  //             inputElement.type = "checkbox";
  //           }

  //           if (inputElement) {
  //             inputElement.className="dynamic-input";
  //             inputElement.id = fieldName;
  //             // inputElement.required = required.toLowerCase() === "yes"; 
  //             inputElement.required=required;
  //             inputContainer.appendChild(inputElement); 
  //             formSelector.appendChild(inputContainer); 
  //           }

  //           return;
  //     }


  //     // start
  //     documentLibraryFields.forEach((field)=>{
  //       createElement(field.ColumnName,field.ColumnType,field.IsRequired,field.IsRename);
  //       })
  //     // end


  //     // properties of upload file div
  //     // uploadFileDiv.className="uploadfile";
  //     uploadFileDiv.className="input-container";

  //     // input for upload file
  //     const uploadFileInput=document.createElement('input');
  //     uploadFileInput.className="dynamic-input";
  //     uploadFileInput.type="file";
  //     uploadFileInput.id="fileInput";
  //     uploadFileInput.addEventListener('change', (event:any) => handleFileChange(event))

  //     // Set Label For upload file
  //     const label = document.createElement("label");
  //     label.setAttribute("htmlFor", 'fileInput');
  //     label.textContent = 'Upload File';
      
  //     // Add a red asterisk if the field is required
  //     const asterisk = document.createElement("span");
  //     asterisk.textContent = " *";
  //     asterisk.style.color = "red"; 
  //     label.appendChild(asterisk);
  //     uploadFileDiv.appendChild(label);
  //     uploadFileDiv.appendChild(uploadFileInput);
  //     formSelector.appendChild(uploadFileDiv);

  //     // Submit Button property
      
  //     submitButton.addEventListener('click',handleSubmit)
  
  //     formSelector.appendChild(submitButton);
      
  //   } catch (error) {
  //     console.error("Error loading form options:", error);
  //   }
  //     }  

  const loadFormOptions = async () => {
  try {
    // 1. Extract the Root Site URL dynamically (e.g., https://.../sites/AlRostmaniSpfx2)
    const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
    if (!rootSiteMatch) {
      console.error("Could not determine Root Site URL for metadata fields.");
      return;
    }
    const libraryName = getRootLibraryName();
    const rootSiteUrl = rootSiteMatch[0];

    // 2. Initialize the Web object for the Root Site where the list lives
    const remoteWeb = Web([sp.web, rootSiteUrl]);

    // 3. Query the list using remoteWeb instead of sp.web
    const documentLibraryFields = await remoteWeb.lists
      .getByTitle("DMSPreviewFormMaster")
      .items.select("ColumnName", "ColumnType", "IsRequired", "IsRename")
      .filter(
        `SiteName eq '${currentfolderpath.Entity}' 
         and DocumentLibraryName eq '${libraryName}' 
         and AddorRemoveThisColumn eq 'Add To Library' and IsInProgress eq 0`
      )();

    console.log("Document Library Fields retrieved:", documentLibraryFields);

    {/* srs 19/2/26 */}
    // const formSelector = document.getElementById("formSelector");
    const formSelector = document.getElementById("dynamicMetadataContainer");
    {/* srs 19/2/26 */}
    if (!formSelector) return;

    // Clear existing dynamic content if necessary (to prevent duplicates on re-render)
    formSelector.innerHTML = "";
    const uploadFileDiv = document.createElement('div');

    const createElement = (fieldName: string, type: string, required: boolean, IsRename: string) => {
      let fName = IsRename !== null ? IsRename : fieldName;

      const inputContainer = document.createElement("div");
      inputContainer.className = "input-container";

      const label = document.createElement("label");
      label.setAttribute("htmlFor", fieldName);

      if (required) {
        const asterisk = document.createElement("span");
        asterisk.textContent = " *";
        asterisk.style.color = "red";
        label.textContent = fName;
        label.appendChild(asterisk);
      } else {
        label.textContent = fName;
      }
      inputContainer.appendChild(label);

      let inputElement: HTMLInputElement | null = null;
      let modifiedType = type?.replace(/\s+/g, '').toLowerCase();

      if (["singlelineoftext", "multiplelineoftext", "text"].includes(modifiedType)) {
        inputElement = document.createElement("input");
        inputElement.type = "text";
      } else if (modifiedType === "number") {
        inputElement = document.createElement("input");
        inputElement.type = "number";
      } else if (modifiedType === "date&time") {
        inputElement = document.createElement("input");
        inputElement.type = "date";
      } else if (modifiedType === "yesorno") {
        inputElement = document.createElement("input");
        inputElement.type = "checkbox";
      }

      if (inputElement) {
        inputElement.className = "dynamic-input";
        inputElement.id = fieldName;
        inputElement.required = required;
        inputContainer.appendChild(inputElement);
        formSelector.appendChild(inputContainer);
      }
    };

    // Render fields from the list
    documentLibraryFields.forEach((field) => {
      createElement(field.ColumnName, field.ColumnType, field.IsRequired, field.IsRename);
    });

    // Handle Upload File Field
    uploadFileDiv.className = "input-container";
    const uploadFileInput = document.createElement('input');
    uploadFileInput.className = "dynamic-input";
    uploadFileInput.type = "file";
    uploadFileInput.id = "fileInput";
    uploadFileInput.addEventListener('change', (event: any) => handleFileChange(event));

    const fileLabel = document.createElement("label");
    fileLabel.setAttribute("htmlFor", 'fileInput');
    fileLabel.textContent = 'Upload File';
    const fileAsterisk = document.createElement("span");
    fileAsterisk.textContent = " *";
    fileAsterisk.style.color = "red";
    fileLabel.appendChild(fileAsterisk);

    uploadFileDiv.appendChild(fileLabel);
    uploadFileDiv.appendChild(uploadFileInput);
    formSelector.appendChild(uploadFileDiv);

    // Re-attach Submit Button
    submitButton.addEventListener('click', handleSubmit);
    formSelector.appendChild(submitButton);

  } catch (error) {
    console.error("Error loading form options:", error);
  }
};
      loadFormOptions();
},[])
  

// const handleSubmit = async (event: any) => {
//   event.preventDefault();
//   console.log("Button clicked");

//   const formSelector = document.getElementById("formSelector") as HTMLFormElement;
//   if (!formSelector.checkValidity()) {
//       formSelector.reportValidity(); // Show validation errors
//       return;
//   }
//   const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
//   submitBtn.disabled = true;
//   submitBtn.innerText = "Submitting...";

//   const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
//   const spinner = document.getElementById("spinner") as HTMLElement;
//   spinner.style.display = "none";
//   iframe.style.display = "none";
//   setIsFinalUploading(true);
//   // Prepare the payload for SharePoint dynamically
//   const inputs = document.querySelectorAll('.dynamic-input');
//   const payload: any = {};

//   inputs.forEach((input) => {
//       const inputElement = input as HTMLInputElement;
//       const fieldName = inputElement.id;
//       if (!fieldName) return; // Skip if field name is invalid

//       if (inputElement.type === "checkbox") {
//           payload[fieldName] = inputElement.checked;
//       } else if (inputElement.type !== "file") {
//           if (inputElement.value !== "") {
//               payload[fieldName] = inputElement.value;
//           }
//       }
//   });

//   const fileInput = document.getElementById('fileInput') as HTMLInputElement;
//   const selectedFile = fileInput?.files?.[0];

//   if (!selectedFile) {
//       console.error("No file selected.");
//       return;
//   }

//   try {
//       console.log("Payload:", payload);
//       console.log("SiteID:", currentfolderpath.siteID);
//       const targetWebUrl = currentfolderpath.siteID; 
//     console.log("Targeting Web URL:", targetWebUrl);

//     // 2. Initialize the web object using the URL and cloning the context
//     // This replaces: const testidsub = await sp.site.openWebById(...)
//     const remoteWeb = Web(targetWebUrl).using(AssignFrom(sp.web as any));
//     const documentLibraryInWhichWeUploadTheFile = remoteWeb.getFolderByServerRelativePath(currentfolderpath.folderpath);

//       // const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
//       // if (!testidsub) throw new Error("Subsite not found.");

//       // const documentLibraryInWhichWeUploadTheFile = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
//       // console.log("Current Path:", documentLibraryInWhichWeUploadTheFile);

//       const files = await documentLibraryInWhichWeUploadTheFile.files();

//       const originalFileName = selectedFile.name;
//       const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
//       const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
//       let uniqueFileName = originalFileName;
//       let counter = 1;

//       while (files.some(file => file.Name === uniqueFileName)) {
//           uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
//           counter++;
//       }
//       console.log(`Unique file name generated: ${uniqueFileName}`);

//       const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(uniqueFileName, selectedFile);
//       console.log("File uploaded successfully", uploadResult.data.Name);

//       // const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
//       // submitBtn.disabled = true;
//       // submitBtn.innerText = "Submitting..."; // Optional: Change button text to indicate progress

//       const listItem = await uploadResult.file.getItem();
//       console.log("ListItems ", listItem);

//       const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
//       const siteUrl = window.location.origin;
//       const encodeSharePointURL = (url: string) => {
//         return encodeURIComponent(url)
//             .replace(/'/g, "%27")
//             .replace(/-/g, "%2D")
//             .replace(/_/g, "%5F")
//             .replace(/\./g, "%2E")
//             .replace(/!/g, "%21")
//             .replace(/\*/g, "%2A")
//             .replace(/\(/g, "%28")
//             .replace(/\)/g, "%29")
//             .replace(/~/g, "%7E")
//             .replace(/@/g, "%40")
//             .replace(/\$/g, "%24")
//             .replace(/,/g, "%2C")
//             .replace(/;/g, "%3B")
//             .replace(/:/g, "%3A")
//             .replace(/\+/g, "%2B")
//             .replace(/=/g, "%3D")
//             .replace(/\?/g, "%3F")
//             .replace(/\//g, "%2F")
//             .replace(/#/g, "%23")
//             .replace(/&/g, "%26");
//     };
    
//     // Usage
//     const encodedFilePath = encodeSharePointURL(uploadResult.data.ServerRelativeUrl);
//     const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeSharePointURL(parentFolder)}`;
    
//       // Encode the file path for the preview URL
//       // const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl).replace(/'/g, "%27");
//       // const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

//       console.log("Generated Preview URL:", previewUrl);
//       if (!listItem) throw new Error("List item not found for the uploaded file.");

//       let status = "";
//       if (IsApproval === true) {
//           status = "Pending";
//       } else if (IsApproval === false) {
//           status = "Auto Approved";
//       }
//       (payload as any).Status = status;
//       await listItem.update(payload);
//       console.log("File metadata updated successfully with:", payload);

//       const newRequestNo = await getUniqueRequestNo();
//       const newItem = await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
//           FileName: String(uploadResult.data.Name),
//           FileSize: String(uploadResult.data.Length),
//           FileVersion: String(uploadResult.data.MajorVersion),
//           CurrentFolderPath: String(currentfolderpath.folderpath),
//           FileUID: String(uploadResult.data.UniqueId),
//           CurrentUser: String(currentUserEmailRef.current),
//           SiteID: String(currentfolderpath.siteID),
//           Status: status,
//           FilePreviewURL: String(previewUrl),
//           DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
//           SiteName: String(currentfolderpath.Entity),
//           MyRequest: true,
//           Processname: 'New File Request',
//           RequestNo: newRequestNo,
       
//       });
//       console.log(newItem, "New item added FileMaster");

//       if (IsApproval === true) {
//           await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
//               SiteName: String(currentfolderpath.Entity),
//               DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
//               RequestedBy: String(currentUserEmailRef.current),
//               FileName: String(uploadResult.data.Name),
//               FileUID: String(uploadResult.data.UniqueId),
//               FilePreviewUrl: String(previewUrl),
//               Status: String('Pending'),
//               FolderPath: String(currentfolderpath.folderpath),
//               ApproveAction: String('Submitted'),
//               ApprovedLevel: 1,
//               RequestNo: newRequestNo,
//               Processname: 'New File Request',
//               CurrentLevel : 1
//           });
//       }

//       if (newItem) {
//           setIsFinalUploading(false);
//           Deletemedia();
//           setTimeout(() => {
//               location.reload();
//               onReturnToMain();
//           }, 3000);
//       }
//   } catch (error) {
//       console.error("Error during submission:", error);
//   }
// };


// const getUniqueRequestNo = async () => {
//   const counterItem = await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1)();
//   console.log("Counter Item 0", counterItem);
//   console.log("Counter Item 1", counterItem.FileCount);
//   let fileCounter = counterItem.FileCount;

//   // Increment the counter
//   fileCounter++;

//   // Generate the new RequestNo
//   const newRequestNo = `File${String(fileCounter).padStart(2, '0')}`;

//   // Update the counter in the CounterList
//   await sp.web.lists.getByTitle('DMSFileCounterList').items.getById(1).update({
//     FileCount: fileCounter
//   });

//   return newRequestNo;
// };

// const handleSubmit = async (event: any) => {
//   event.preventDefault();
  
//   const formSelector = document.getElementById("formSelector") as HTMLFormElement;
//   if (!formSelector.checkValidity()) {
//     formSelector.reportValidity();
//     return;
//   }
  
//   const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
//   submitBtn.disabled = true;
//   submitBtn.innerText = "Submitting...";

//   setIsFinalUploading(true);

//   // 1. Prepare dynamic payload
//   const inputs = document.querySelectorAll('.dynamic-input');
//   const payload: any = {};
//   inputs.forEach((input) => {
//     const inputElement = input as HTMLInputElement;
//     const fieldName = inputElement.id;
//     if (!fieldName) return;
//     if (inputElement.type === "checkbox") {
//       payload[fieldName] = inputElement.checked;
//     } else if (inputElement.type !== "file") {
//       if (inputElement.value !== "") payload[fieldName] = inputElement.value;
//     }
//   });

//   const fileInput = document.getElementById('fileInput') as HTMLInputElement;
//   const selectedFile = fileInput?.files?.[0];

//   if (!selectedFile) return;

//   try {
//     // --- SITE URL SETUP ---
//     // const targetWebUrl = currentfolderpath.siteID; // The Subsite URL
//     // srs 19/2/26 
//     const targetWebUrl = currentfolderpath.Entityurl;
//     const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
//     const rootSiteUrl = rootSiteMatch ? rootSiteMatch[0] : ""; // The Root Site Collection URL

//     // Create web objects for both Subsite (for upload) and Root (for metadata)
//     const subsiteWeb = Web(targetWebUrl).using(AssignFrom(sp.web as any));
//     const rootWeb = Web(rootSiteUrl).using(AssignFrom(sp.web as any));

//     // --- UPLOAD TO SUBSITE ---
//     const folder = subsiteWeb.getFolderByServerRelativePath(currentfolderpath.folderpath);
//     const files = await folder.files();

//     let uniqueFileName = selectedFile.name;
//     const fileExtension = uniqueFileName.substring(uniqueFileName.lastIndexOf('.'));
//     const baseFileName = uniqueFileName.substring(0, uniqueFileName.lastIndexOf('.'));
//     let counter = 1;

//     while (files.some(file => file.Name === uniqueFileName)) {
//       uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
//       counter++;
//     }

//     const uploadResult = await folder.files.addChunked(uniqueFileName, selectedFile);
//     const listItem = await uploadResult.file.getItem();

//     // --- GENERATE PREVIEW URL ---
//     const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
//     const siteUrl = window.location.origin;
//     const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl).replace(/'/g, "%27");
//     const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

//     // --- UPDATE METADATA (Subsite Library) ---
//     let status = IsApproval === true ? "Pending" : "Auto Approved";
//     payload.Status = status;
//     await listItem.update(payload);

//     // --- GET UNIQUE REQUEST NO (From Root Site) ---
//     const newRequestNo = await getUniqueRequestNo();

//     // --- ADD TO FILE MASTER (Root Site) ---
//     // Instead of sp.web.lists, we use rootWeb.lists
//     const newItem = await rootWeb.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
//       FileName: String(uploadResult.data.Name),
//       FileSize: String(uploadResult.data.Length),
//       FileVersion: String(uploadResult.data.MajorVersion),
//       CurrentFolderPath: String(currentfolderpath.folderpath),
//       FileUID: String(uploadResult.data.UniqueId),
//       CurrentUser: String(currentUserEmailRef.current),
//       SiteID: String(currentfolderpath.siteID),
//       Status: status,
//       FilePreviewURL: String(previewUrl),
//       DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
//       SiteName: String(currentfolderpath.Entity),
//       MyRequest: true,
//       Processname: 'New File Request',
//       RequestNo: newRequestNo,
//     });

//     // --- ADD TO APPROVAL LIST (Root Site) ---
//     if (IsApproval === true) {
//       await rootWeb.lists.getByTitle('DMSFileApprovalList').items.add({
//         SiteName: String(currentfolderpath.Entity),
//         DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
//         RequestedBy: String(currentUserEmailRef.current),
//         FileName: String(uploadResult.data.Name),
//         FileUID: String(uploadResult.data.UniqueId),
//         FilePreviewUrl: String(previewUrl),
//         Status: String('Pending'),
//         FolderPath: String(currentfolderpath.folderpath),
//         ApproveAction: String('Submitted'),
//         ApprovedLevel: 1,
//         RequestNo: newRequestNo,
//         Processname: 'New File Request',
//         CurrentLevel : 1
//       });
//     }

//     if (newItem) {
//       setIsFinalUploading(false);
//       Deletemedia();
//       setTimeout(() => {
//         location.reload();
//         onReturnToMain();
//       }, 3000);
//     }
//   } catch (error) {
//     console.error("Error during submission:", error);
//     setIsFinalUploading(false);
//     submitBtn.disabled = false;
//     submitBtn.innerText = "Submit";
//   }
// };

const handleSubmit = async (event: any) => {
  // Prevent default form behavior and stop event propagation 
  // to help prevent the "running many times" issue
  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }

  const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  if (!formSelector.checkValidity()) {
    formSelector.reportValidity();
    return;
  }

  const submitBtn = document.getElementById("submitBtn") as HTMLButtonElement;
  if (submitBtn.disabled) return; // Guard against multiple clicks

  submitBtn.disabled = true;
  submitBtn.innerText = "Submitting...";
  setIsFinalUploading(true);

  // 1. Prepare dynamic payload
  const inputs = document.querySelectorAll('.dynamic-input');
  const payload: any = {};
  inputs.forEach((input) => {
    const inputElement = input as HTMLInputElement;
    const fieldName = inputElement.id;
    if (!fieldName) return;
    if (inputElement.type === "checkbox") {
      payload[fieldName] = inputElement.checked;
    } else if (inputElement.type !== "file") {
      if (inputElement.value !== "") payload[fieldName] = inputElement.value;
    }
  });

  const fileInput = document.getElementById('fileInput') as HTMLInputElement;
  const selectedFile = fileInput?.files?.[0];

  if (!selectedFile) {
    setIsFinalUploading(false);
    submitBtn.disabled = false;
    submitBtn.innerText = "Submit";
    return;
  }

  try {
    // --- SITE URL SETUP ---
    const targetWebUrl = currentfolderpath.Entityurl;
    const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
    const rootSiteUrl = rootSiteMatch ? rootSiteMatch[0] : "";

    // Create web objects
    // Using a clean Web() instance and AssignFrom to avoid Proxy inheritance issues
    const subsiteWeb = Web(targetWebUrl).using(AssignFrom(sp.web as any));
    const rootWeb = Web(rootSiteUrl).using(AssignFrom(sp.web as any));

    // --- UPLOAD TO SUBSITE ---
    const folder = subsiteWeb.getFolderByServerRelativePath(currentfolderpath.folderpath);
    const files = await folder.files();

    let uniqueFileName = selectedFile.name;
    const fileExtension = uniqueFileName.substring(uniqueFileName.lastIndexOf('.'));
    const baseFileName = uniqueFileName.substring(0, uniqueFileName.lastIndexOf('.'));
    let counter = 1;

    while (files.some(file => file.Name === uniqueFileName)) {
      uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
      counter++;
    }

    const uploadResult = await folder.files.addChunked(uniqueFileName, selectedFile);

    // --- THE FIX: GET LIST ITEM BY GUID ---
    // Instead of uploadResult.file.getItem(), we use the list and filter by GUID
   // --- THE FIX: ROBUST ITEM RETRIEVAL ---
const libraryTitle = currentfolderpath.DocumentLibrary;
const fileUniqueId = uploadResult.data.UniqueId;
const fileServerRelativeUrl = uploadResult.data.ServerRelativeUrl;

let itemData = [];

// Try Lookup 1: Filter by UniqueId (Most accurate)
// Note: We use the guid'...' prefix which is the standard OData format for SharePoint
itemData = await subsiteWeb.lists.getByTitle(libraryTitle).items
  .filter(`UniqueId eq guid'${fileUniqueId}'`)
  .select("Id")();

// Fallback Lookup 2: If UniqueId fails, try FileLeafRef + Folder path (Name-based)
if (itemData.length === 0) {
  console.log("UniqueId lookup failed, trying fallback...");
  itemData = await subsiteWeb.lists.getByTitle(libraryTitle).items
    .filter(`FileRef eq '${fileServerRelativeUrl}'`)
    .select("Id")();
}

if (itemData.length === 0) {
  // If both fail, the item might not be indexed yet. Wait 1 second and try once more.
  await new Promise(resolve => setTimeout(resolve, 1500));
  itemData = await subsiteWeb.lists.getByTitle(libraryTitle).items
    .filter(`FileRef eq '${fileServerRelativeUrl}'`)
    .select("Id")();
}

if (itemData.length === 0) {
  throw new Error("SharePoint created the file but hasn't generated the list item yet. Please refresh in a moment.");
}

const itemId = itemData[0].Id;
const listItem = subsiteWeb.lists.getByTitle(libraryTitle).items.getById(itemId);

    // --- GENERATE PREVIEW URL ---
    const serverRelUrl = uploadResult.data.ServerRelativeUrl;
    const parentFolder = serverRelUrl.substring(0, serverRelUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;
    const encodedFilePath = encodeURIComponent(serverRelUrl).replace(/'/g, "%27");
    
    const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

    // --- UPDATE METADATA (Subsite Library) ---
    let statusValue = IsApproval === true ? "Pending" : "Auto Approved";
    payload.Status = statusValue;
    
    // Update the item directly via the list reference
    await listItem.update(payload);

    // --- GET UNIQUE REQUEST NO (From Root Site) ---
    const newRequestNo = await getUniqueRequestNo();

    // --- ADD TO FILE MASTER (Root Site) ---
    const newItem = await rootWeb.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add({
      FileName: String(uploadResult.data.Name),
      FileSize: String(uploadResult.data.Length),
      FileVersion: String(uploadResult.data.MajorVersion),
      CurrentFolderPath: String(currentfolderpath.folderpath),
      FileUID: String(uploadResult.data.UniqueId),
      CurrentUser: String(currentUserEmailRef.current),
      SiteID: String(currentfolderpath.siteID),
      Status: statusValue,
      FilePreviewURL: String(previewUrl),
      DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
      SiteName: String(currentfolderpath.Entity),
      MyRequest: true,
      Processname: 'New File Request',
      RequestNo: newRequestNo,
    });

    // --- ADD TO APPROVAL LIST (Root Site) ---
    if (IsApproval === true) {
      await rootWeb.lists.getByTitle('DMSFileApprovalList').items.add({
        SiteName: String(currentfolderpath.Entity),
        DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
        RequestedBy: String(currentUserEmailRef.current),
        FileName: String(uploadResult.data.Name),
        FileUID: String(uploadResult.data.UniqueId),
        FilePreviewUrl: String(previewUrl),
        Status: String('Pending'),
        FolderPath: String(currentfolderpath.folderpath),
        ApproveAction: String('Submitted'),
        ApprovedLevel: 1,
        RequestNo: newRequestNo,
        Processname: 'New File Request',
        CurrentLevel: 1
      });
    }

    if (newItem) {
      setIsFinalUploading(false);
      Deletemedia();
      setTimeout(() => {
        location.reload();
        onReturnToMain();
      }, 3000);
    }
  } catch (error) {
    console.error("Error during submission:", error);
    setIsFinalUploading(false);
    submitBtn.disabled = false;
    submitBtn.innerText = "Submit";
    
    // Alert the user to the failure
    Swal.fire({
      icon: 'error',
      title: 'Submission Failed',
      text: 'The file was uploaded, but metadata update failed. Please check the folder.'
    });
  }
};
const getUniqueRequestNo = async () => {
  try {
    // 1. Extract the Root Site URL (where the counter list lives)
    const rootSiteMatch = currentfolderpath.Entityurl.match(/.*\/sites\/[^\/]+/);
    if (!rootSiteMatch) throw new Error("Could not determine Root Site URL");
    const rootSiteUrl = rootSiteMatch[0];

    // 2. Initialize the Web object for the Root Site
    const rootWeb = Web(rootSiteUrl).using(AssignFrom(sp.web as any));

    // 3. Target the list on the ROOT web, not sp.web
    const list = rootWeb.lists.getByTitle('DMSFileCounterList');
    
    // 4. Get the counter item
    const counterItem = await list.items.getById(1)();
    let fileCounter = counterItem.FileCount;

    // Increment
    fileCounter++;

    // Generate Request No
    const newRequestNo = `File${String(fileCounter).padStart(2, '0')}`;

    // 5. Update the counter back on the ROOT site
    await list.items.getById(1).update({
      FileCount: fileCounter
    });

    return newRequestNo;
  } catch (error) {
    console.error("Error in getUniqueRequestNo:", error);
    throw error;
  }
};


const handleSubmitBulk = async (event: any) => {
  event.preventDefault();
  console.log("Bulk upload button clicked");
  const submitBtn = document.getElementById("submitBtn2") as HTMLButtonElement;
  submitBtn.disabled = true;
  submitBtn.innerText = "Submitting...";

  const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
  const spinner = document.getElementById("spinner") as HTMLElement;
  spinner.style.display = "none";
  iframe.style.display = "none";
  setIsFinalUploading(true);
  try {
    console.log("SiteID:", currentfolderpath.siteID);

    // const testidsub = await sp.site.openWebById(currentfolderpath.siteID);
    // if (!testidsub) throw new Error("Subsite not found.");

    // const documentLibrary = testidsub.web.getFolderByServerRelativePath(currentfolderpath.folderpath);
     // 1. Create the web object for the subsite
    const remoteWeb = Web(currentfolderpath.Entityurl);
    
    // 2. Assign the configuration from your main 'sp' object
    remoteWeb.using(AssignFrom(sp.web as any)); 

    // 3. Now you can use it without the "No observers" error
    const documentLibrary = remoteWeb.getFolderByServerRelativePath(currentfolderpath.folderpath);
    console.log("Current Path:", documentLibrary);

    const files = await documentLibrary.files();
    const existingFileNames = new Set(files.map(file => file.Name)); // Faster lookup
    const siteUrl = window.location.origin;

    // Function to generate a unique filename
    const generateUniqueFileName = (originalFileName: string) => {
      const fileExtension = originalFileName.substring(originalFileName.lastIndexOf('.'));
      const baseFileName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      let uniqueFileName = originalFileName;
      let counter = 1;
      
      while (existingFileNames.has(uniqueFileName)) {
        uniqueFileName = `${baseFileName}(${counter})${fileExtension}`;
        counter++;
      }
      existingFileNames.add(uniqueFileName); // Add new name to the set
      return uniqueFileName;
    };

    // Upload all files in parallel
    const uploadPromises = uploadedFiles.map(async ({ name, file }) => {
      try {
        const uniqueFileName = generateUniqueFileName(name);
        console.log(`Uploading file: ${uniqueFileName}`);

        const uploadResult = await documentLibrary.files.addChunked(uniqueFileName, file);
        console.log(`${name} uploaded successfully`);

        const listItem = await uploadResult.file.getItem();
        if (!listItem) throw new Error("List item not found for the uploaded file.");

        const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
        const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
        const previewUrl = `${siteUrl}${locationPath}/${currentfolderpath.Entity}/${currentfolderpath.DocumentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

        // Update metadata
        await listItem.update({ Status: "Auto Approved" });

        // Prepare new file entry
        // const newRequestNo = await getUniqueRequestNo(1);
        return {
          FileName: String(uploadResult.data.Name),
          FileSize: String(uploadResult.data.Length),
          FileVersion: String(uploadResult.data.MajorVersion),
          CurrentFolderPath: String(currentfolderpath.folderpath),
          FileUID: String(uploadResult.data.UniqueId),
          CurrentUser: String(currentUserEmailRef.current),
          SiteID: String(currentfolderpath.siteID),
          Status: "Auto Approved",
          FilePreviewURL: String(previewUrl),
          DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
          SiteName: String(currentfolderpath.Entity),
          MyRequest: true,
          Processname: 'New File Request',
          // RequestNo: newRequestNo,
        };
      } catch (error) {
        console.error(`Error uploading file ${name}:`, error);
        return null;
      }
    });

    // Wait for all uploads to finish
    const uploadedItems = await Promise.all(uploadPromises);
    const validItems = uploadedItems.filter(item => item !== null);
    console.log(validItems ,"validItemsvalidItems")
    console.log(typeof(validItems) , " here i am checking validItemsvalidItems")
    // srs 19/2/26
       // 1. Get the current URL hash (e.g., #AlRostmaniSpfx2/N27JAN/TESTSRS)
const urlHash = window.location.hash; 

// 2. Extract the Site Collection name (the part before the first slash)
// This removes the '#' and takes 'AlRostmaniSpfx2'
const siteCollectionName = urlHash.split('/')[0].replace('#', '');

// 3. Construct the Root Site URL
const rootSiteUrl = `${window.location.origin}/sites/${siteCollectionName}`;
console.log("Targeting List at Root Site:", rootSiteUrl);

// 4. Create a specific Web object for the Root Site
const rootWeb = Web(rootSiteUrl);
rootWeb.using(AssignFrom(sp.web as any));
    // Batch insert all successfully uploaded files into FileMaster list
    for (const item of validItems) {
      // await sp.web.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add(item);
      await rootWeb.lists.getByTitle(`DMS${currentfolderpath.Entity}FileMaster`).items.add(item);
    }

    // Cleanup and refresh UI
    setIsFinalUploading(false);
    Deletemedia();
    setTimeout(() => {
      location.reload();
      onReturnToMain();
    }, 3000);

  } catch (error) {
    console.error("Error in bulk upload:", error);
  }
};


const Deletemedia = () => {
 
  Swal.fire({
    title: "Success",
    text: "File uploaded successfully",
    icon: "success"
  });


 setTimeout(() => {
    Swal.close();
    onReturnToMain(); 
  }, 3000);

}



const handleToggle = () => {
  // alert(`isChecked before toggle: ${isChecked}`);/
   console.log("isChecked before toggle:", isChecked);
  setIsChecked((prev) => {
    const newCheckedState = !prev; // Get the updated state
    console.log(`New isChecked state: ${newCheckedState}`);
    // Get all elements with the class "input-container"
    const inputContainers = document.getElementsByClassName("input-container");

    // Hide or show input fields based on the new state
    for (let i = 0; i < inputContainers.length; i++) {
      (inputContainers[i] as HTMLElement).style.display = newCheckedState ? "none" : "block";
    }
    const getsubmitbutton = document.getElementById("submitBtn") as HTMLButtonElement;
    // getsubmitbutton.style.display = newCheckedState ? "none" : "block";
    const fileInput = document.getElementById('fileInput') as HTMLInputElement;
    const selectedFile = fileInput?.files?.[0];
    getsubmitbutton.style.display = newCheckedState ? "none" : 'block';
    getsubmitbutton.disabled= !(!newCheckedState && selectedFile)
   
    return newCheckedState; // Update the state
  });
};
const handleRemove = (fileIndex:any) => {
  event.preventDefault();
  const updatedFiles = uploadedFiles.filter((_, index) => index !== fileIndex);
  setUploadedFiles(updatedFiles); // Update state
  console.log(uploadedFiles , "uploadedFiles in remove")
};

useEffect(()=>{
  const getsubmitbuttonbulk = document.getElementById("submitBtn2") as HTMLButtonElement;
  if(getsubmitbuttonbulk){
    console.log("isChecked-isChecked",isChecked)
    getsubmitbuttonbulk.style.display= isChecked ? 'block': 'none';
    getsubmitbuttonbulk.disabled = !(isChecked && uploadedFiles.length > 0);
  }
},[isChecked,uploadedFiles]);
const breadcrumbParts = useMemo(() => {
  try{
    const parts: string[] = [];
    if (currentfolderpath?.Entity) parts.push(currentfolderpath.Entity);
    if (currentfolderpath?.DocumentLibrary) parts.push(currentfolderpath.DocumentLibrary);
    // prefer explicit Folder prop if present
    if (currentfolderpath?.Folder) {
      const f = currentfolderpath.Folder;
      if (f) parts.push(f);
    } else if (currentfolderpath?.folderpath) {
      const segs = String(currentfolderpath.folderpath).split('/').filter(s => s && s.trim() !== '');
      // try to remove any site segments and the document library itself
      const docIdx = segs.findIndex(s => s === currentfolderpath.DocumentLibrary);
      const rest = docIdx >= 0 ? segs.slice(docIdx + 1) : segs;
      parts.push(...rest);
    }
    return parts;
  }catch(e){
    return [];
  }
}, [currentfolderpath]);
    return (
      <>
      <div className="">
      <div className="card-body">


       
          <div style={{float:'right',marginTop:'-45px'}} className='' 
          onClick={()=>{location.reload() ;onReturnToMain()}}
          >   
           <span className="mb-1" data-tooltip='Back'>
           <img  src={back}></img> &nbsp;</span>
          </div>
          <div className="mt-0 pt-1 UploadFileCont">
              <div className='row'>
              <div className='col-lg-6'>
              <nav className="dms-breadcrumb" aria-label="Breadcrumb" style={{marginBottom:12}}>
              <img  className="me-1" src={info}></img> 
            {breadcrumbParts && breadcrumbParts.length > 0 ? (
              breadcrumbParts.map((seg, idx) => (
                <span key={idx} className="dms-breadcrumb-segment">
                  <span className="dms-breadcrumb-text">{seg}</span>
                  {idx < breadcrumbParts.length - 1 && (
                    <span className="dms-breadcrumb-sep">&nbsp;&gt;&nbsp;</span>
                  )}
                </span>
              ))
            ) : (
              <span className="dms-breadcrumb-text">Upload</span>
            )}
          </nav>
                      {/* <h1>File Preview</h1> */}
                      <div className="borderprev">
                        {isUploading && (
                          <>
                         
                              <div id="spinner" style={{display: "block", paddingTop:'170px', textAlign: "center"}}>
                                  <div>
                                    <img
                                      src={require("../../../CustomAsset/argloader.gif")}
                                      className="alignrightl"
                                      alt="Loading..."
                                    />                                                               
                                </div>
                                <span>Please wait, we are preparing your files for upload... </span>{" "}
                          </div>
                         
                          </>
                        )}
                        {isFinalUploading && (
                          <>
                          
                              <div id="spinner" style={{display: "block", paddingTop:'170px', textAlign: "center"}}>
                                  <div>
                                    <img
                                      src={require("../../../CustomAsset/argloader.gif")}
                                      className="alignrightl"
                                      alt="Loading..."
                                    />                                                               
                                </div>
                                <span>Uploading items... This may take a moment. </span>{" "}
                          </div>
                         
                          </>
                        )}
                    
                          <div id="spinner" style={{display: "none",paddingTop:'170px', textAlign: "center"}}>
                          <div>
                            <img
                              src={require("../../../CustomAsset/argloader.gif")}
                              className="alignrightl"
                              alt="Loading..."
                            />                                                                           
                          </div>
                          <span>Loading </span>{" "}
                          </div>
                          <iframe id="filePreview" style={{background:'transparent'}} width="100%" height="400"></iframe>
                      </div>
                      </div>
                     
                      <div className='col-lg-6'>
                          <form id='formSelector'>
                              <h1 className="font-16 fw-bold text-dark mb-0">Upload file</h1>
                              {/* <label className="switch">
                              <input type="checkbox"/>
                              <span className="slider round"></span>
                            </label> */}
                            {showBulkUpload === true && ( 
                              <p style={{color:'#6c757d'}} className="font-14"> Files uploaded to this folder require approval. After submission, your request will be reviewed, and the file will become visible only after it has been approved.
 </p>
                             )}
                              {/* srs 19/2/26 */}
                             <div id="dynamicMetadataContainer"></div> 
                             {/* srs 19/2/26 */}
                             <div>
      {showBulkUpload === false && ( // Show only if IsApproval is false
      <div style={{display:'flex', justifyContent:'space-between',alignItems:'center'}} className="mt-3 mb-3">
        <p className="mb-0 text-dark">Bulk  Upload :</p>
        <div style={{display:'flex', gap:'5px', alignItems:'center'}}>
          <label className="switch">
          <input type="checkbox" checked={isChecked} onChange={handleToggle} />
          <span className="slider round"></span>
          
        </label>
        <p className="mb-0 text-dark fw-bold">{isChecked ? "ON" : "OFF"}</p>
         </div>

        
        
      </div>
     
        
      )}
       <div>
       {/* srs 19/2/26 */}
      {/* {isChecked && ( */}
      {showBulkUpload === false && isChecked && (
        <div className="input-container mt-3">
                  {/* <label htmlFor="Uplaod bulk">Bulk upload</label> */}
                  
                  {/* <label htmlFor="bulkfile" style={{ fontWeight: "bold" }}>
                      Upload File <span style={{ color: "red" }}>*</span>
                  </label> */}
        <input type="file" name="bulkfile" id="bulkfile" multiple onChange={(e)=>handlebulkFileChange(e)}/>
        <ul className="newbulnup">
        <div className="d-flex align-items-center justify-content-between"><p className="fw-bold font-14">Selected</p>
           <span className="clearall">Clear All    <img style={{margin:'-3px 0px 0px 3px'}} src={require("../assets/delnew1.png")} /></span>
          </div>
            {uploadedFiles.map((file, index) => (
              
    <li 
      key={index}
      style={{
        backgroundColor: selectedIndex === index ? "#e0f7fa" : "transparent",
      }}
    >
        <div style={{width:'20px', textAlign:'center', fontSize:'14px', float:'left'}}>     
          {index + 1}.
        </div> 
        <div className="font-14" style={{overflow:'hidden', width:'85%', textAlign:'left', textOverflow:'ellipsis',whiteSpace:'nowrap', padding:'0px 5px',  fontWeight:'500'}}>  
          <a style={{color:'#858585'}} href="#" onClick={(e) => {
            // srs 19/2/26 
              e.preventDefault();
              handlePreview(file.url)
              setSelectedIndex(index);
            }
            }>
            {file.name}
          </a>
        </div> 
        <div>   
          <a href="" onClick={() => handleRemove(index)} >
            <img src={require("../assets/delnew.png")} className="fas fa-trash"   alt="delete" />
          </a>
      </div> 
    </li>
  ))}
</ul>
<div style={{display:'flex', justifyContent:'right'}}>
        <div style={{display:'none', marginTop:'0px'}} id="submitBtn2" className="btncolorCreate1" onClick={handleSubmitBulk}> 
        <span className="mb-1" data-tooltip='Bulk Submit'> <img src={require("../assets/submit-new1.png")}    alt="delete" /> </span> </div> 
        </div>
        </div>
      )}
    </div>
    {!isChecked ?   <h3 className="mt-2 mb-2 font-16 text-dark fw-bold">Tags</h3> : null}
    
    </div>
   
                          </form>
                      </div>

                     
              </div>
          </div>
          </div>
          </div>
      </>
    );
  }
export default UploadFile;