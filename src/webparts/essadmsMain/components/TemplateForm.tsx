import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import Swal from 'sweetalert2';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
//addhyan - 13/03/2026
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";
import { get } from '@microsoft/sp-lodash-subset';
//addhyan - 13/03/2026
 
 
 
 
type Template = any;
 
interface Props {
  selectedTemplate: Template | null;
  onSave: (data: any) => void;
  onCancel: () => void;
  sp?: any; // Optional SharePoint client
  context?: any; // Optional SPFx context
}
 
const TemplateForm: React.FC<Props> = ({ selectedTemplate, onSave, onCancel, context }) => {
  // Form fields state
  const [transmittalNumber, setTransmittalNumber] = useState('');
  const [toUser, setToUser] = useState<any>(null);
  const [company, setCompany] = useState<any>(null);
  const [classification, setClassification] = useState<any>(null);
  const [fromUser, setFromUser] = useState<any>(null);
  const [date, setDate] = useState('');
  const [reference, setReference] = useState('');
  const [description, setDescription] = useState('');
 
   const [isFullScreen, setIsFullScreen] = useState(false);
   const [uploadedItemId, setUploadedItemId] = useState<any>(null);
 
 
//addhyan - 13/03/2026
   const [selectedSite, setSelectedSite] = useState("");
   const [selectedSubSite, setSelectedSubSite] = useState("");
 
   const siteSubsiteMap: { [key: string]: string[] } = {
     AlRostmaniSpfx2: ["ALROSTMANITRANS", "AlROSTAMANITEMP", "TransmittalTest"],
     edcspfx: ["ESSATRANS", "ESSATRANS2", "TransmittalTest"],
     Intranetdemos: ["ESSADMS3", "ESSA32", "TransmittalTest"],
   };
 
//addhyan - 13/03/2026
  // Data lists state
  const [allUsers, setAllUsers] = useState<any[]>([]);
  const [companyMasterList, setCompanyMasterList] = useState<any[]>([]);
  const [classificationList, setClassificationList] = useState<any[]>([]);
  const [fileCounter, setFileCounter] = useState<number>(0);
  const [isLoading, setIsLoading] = useState(false);
  const [documentUrl, setDocumentUrl] = useState<string>('');
 
 
 
  //addhyan - 06/03/2026
  const [approvalHierachy, setApprovalHierachy] = useState<any[]>([]);
  const [fileServerRelativeUrl, setFileServerRelativeUrl] = useState<string>("");
    const [selectedfile, setselectedFile] = useState<any>(null);
    const [uploadedFileResponse, setUploadedFileResponse] = useState<any>(null);
 
 
  const [isEditorOpen, setIsEditorOpen] = useState(false);
  const [sp] = useState(() => context ? spfi().using(SPFx(context)) : null);
 
  // Refs
  const currentUserEmailRef = useRef('');
  const currentUserIDref = useRef<number>(0);
  const currentUserTitleRef = useRef('');
 
  // Initialize form data
  useEffect(() => {
    if (selectedTemplate) {
      setDescription(selectedTemplate.description || '');
      initializeFormData();
    }
  }, [selectedTemplate]);
 
  // Initialize form with user and data lists
  const initializeFormData = async () => {
    setIsLoading(true);
    try {
      // Set current date
      const today = new Date().toISOString().split('T')[0];
      setDate(today);
 
      // Fetch data if sp is provided
      if (sp) {
        
        // await fetchAllUsers(); -addhyan  //addhyan - 13/03/2026
        // await fetchDocumentTemplate();
      
      }
    } catch (error) {
      console.error('Error initializing form:', error);
      Swal.fire('Error', 'Failed to initialize form data', 'error');
    } finally {
      setIsLoading(false);
    }
  };
 
//addhyan - 13/03/2026
 
useEffect(() => {
  if (selectedSubSite) {
    resetForm();
    copyfile();
    companyclassification();
    getApprovals();
    fetchAllUsers();
  }
}, [selectedSubSite]);
 

const resetForm = () => {
  setTransmittalNumber("");
  setToUser(null);
  setCompany(null);
  setClassification(null);
  setReference("");
  setDescription("");
  setDocumentUrl("");
  setUploadedItemId(null);
  setUploadedFileResponse(null);
  setApprovalHierachy([]);
  setAllUsers([]);
};
//addhyan - 13/03/2026
 
 
   const copyfile = async () => {
  try {
 
    Swal.fire({
      title: "Preparing Document...",
      text: "Please wait.",
      allowOutsideClick: false,
      didOpen: () => Swal.showLoading(),
    });
 
    /* ---------------- SOURCE SITE ---------------- */
    const sourceWeb = Web([sp.web, `${window.location.origin}/sites/Essa`]);
 
 
 
    const files = await sourceWeb.lists
      .getByTitle("Document Template")
      .items.top(1)();
 
    if (!files.length) return;
 
    const item = await sourceWeb.lists
      .getByTitle("Document Template")
      .items.getById(files[0].ID)
      .select("FileRef", "File/Name")
      .expand("File")();
 
    const fileRef = item.FileRef;
    const fileName = item.File.Name;
 
    /* ---- GET FILE BUFFER ---- */
    const fileBuffer = await sourceWeb
      .getFileByServerRelativePath(fileRef)
      .getBuffer();
 
    /* ---- CREATE NEW NAME ---- */
    const getFormattedDateTime = () => {
      const now = new Date();
      const date = now.toISOString().split("T")[0];
      const time = now.toTimeString().split(" ")[0].replace(/:/g, "-");
      return `${date}_${time}`;
    };
 
    const baseName = fileName.replace(".docx", "");
    const newFileName = `${baseName}_${getFormattedDateTime()}.docx`;
 
    /* ---------------- DESTINATION SITE ---------------- */
    // const destSitePath = "/sites/AlRostmaniSpfx2/TransmittalTest"; - addhyan
    //addhyan - 13/03/2026
    const destSitePath = `/sites/${selectedSite}/${selectedSubSite}`;
    const destWeb = Web([sp.web, `${window.location.origin}${destSitePath}`]);
    const folderPath = `${destSitePath}/TransmittalTemplate`;
//addhyan - 13/03/2026
    /* ---- UPLOAD FILE ---- */
    const uploadResult = await destWeb
      .getFolderByServerRelativePath(folderPath)
      .files.addUsingPath(newFileName, fileBuffer, { Overwrite: true });
      setselectedFile(uploadResult);
        setUploadedFileResponse(uploadResult);
        setFileServerRelativeUrl(uploadResult.data.ServerRelativeUrl);
 
    /* ---- GET UNIQUE ID ---- */
    const uploadedFile = await uploadResult.file
      .select("UniqueId", "Name", "ServerRelativeUrl")();
 
    const uniqueId = uploadedFile.UniqueId;
    const uploadedName = uploadedFile.Name;
 
    const listItem: any = await uploadResult.file.getItem();
    console.log("Uploaded Item ID:", listItem.Id);
    setUploadedItemId(listItem.Id);
 
    
      
 
 
 
    /* ---- CREATE EDITABLE OFFICE URL ---- */
    const docUrl =
      `${window.location.origin}/:w:/r${destSitePath}/_layouts/15/Doc.aspx?sourcedoc={${uniqueId}}&file=${encodeURIComponent(uploadedName)}&action=default`;
 
      
      
      
    
    
 
    /* ---- SET PREVIEW ---- */
    setDocumentUrl(docUrl);   // iframe me show karo
    setIsEditorOpen(true);
 
    Swal.close();
 
  } catch (error) {
    console.error("Copy failed:", error);
    Swal.close();
  }
};
 
 
 
  
const companyclassification = async () => {
 
    try {
      if (!sp) return;
      
      const companyItems = await sp.web.lists.getByTitle('Transmittalcompanymaster').items();
   
      const classificationItems = await sp.web.lists.getByTitle('Transmittalcompanyclassification').items();
     
 
 
      const formattedCompany = companyItems.map((item: any) => ({
        label: item.CompanyNameShort,
        value: item.CompanyNameShort,
        id: item.ID
        
      }));  
      const formattedClassification = classificationItems.map((item: any) => ({
        label: item.classification,
        value: item.classification,
        id: item.ID
      }));
 
      setCompanyMasterList(formattedCompany);
      setClassificationList(formattedClassification);
       // Fetch counter ONCE
            const counterItem = await sp.web.lists
                .getByTitle("TransmittalFileCounterList")
                .items.getById(1)();
            
            const counter = counterItem.FileCount + 1;
            setFileCounter(counter);
            
            // Update counter in database
            await sp.web.lists
                .getByTitle("TransmittalFileCounterList")
                .items.getById(1)
                .update({ FileCount: counter });
            
            console.log("File Counter fetched:", counter);
    } catch (error) {
      console.error('Error fetching company/classification:', error);
    }
  };
 
 
//addhyan - 13/03/2026
 
  const fetchAllUsers = async () => {
 
 
 
     const userdata = await sp.web.currentUser();
     console.log("addhyan1")
    try {
      if (!sp) return;
 
      const userdata = await sp.web.currentUser();
      console.log("addhyan2")
      currentUserIDref.current = userdata.Id;
      currentUserEmailRef.current = userdata.Email;
      currentUserTitleRef.current = userdata.Title;
 
      setFromUser({
        label: userdata.Title,
        value: userdata.Email,
        id: userdata.Id
      });
 
      
      const url =`https://officeindia.sharepoint.com/sites/${selectedSite}/${selectedSubSite}/`
 
      const otherSiteSP = spfi(url).using(SPFx((sp as any)._context));
      const users = await otherSiteSP.web.siteUsers();
      console.log("userall addhyan", users)
     
 
      const formattedUsers = users
        .filter((user: any) => !user.IsHiddenInUI && user.Email)
        .map((user: any) => ({
          value: user.Title,
          label: user.Title,
          email: user.Email,
          id: user.Id
        }));
        console.log('Fetched users:', formattedUsers);
      setAllUsers(formattedUsers);
    } catch (error) {
      console.error('Error fetching users:', error);
    }
  };
 
  //addhyan - 13/03/2026
 
 
 
 
  // Generate transmittal number based on selections
  useEffect(() => {
    if (company && classification && fileCounter) {
      const year = new Date().getFullYear();
      const increment = String(fileCounter).padStart(4, '0');
      const requestNo = `${increment}/TR/${company.value}-${classification.value}/${year}`;
      setTransmittalNumber(requestNo);
    } else {
      setTransmittalNumber('');
    }
  }, [company, classification, fileCounter]);
 
  
 
 
  //addhyan - 06/03/2026
 
 
    const getApprovalTypeText = (type: number) => {
        return type === 0 ? "Approved by One Only" : "Approved by All";
    };
 
 
 
//addhyan - 13/03/2026
 
const getApprovals = async () => {
        try {
 
          const destSitePath = `/sites/${selectedSite}`;
    const destWeb = Web([sp.web, `${window.location.origin}${destSitePath}`]);
    console.log("addhyan a", destSitePath)
    console.log("addhyan b", destWeb)
           const hierarchy = await destWeb.lists
  .getByTitle("DMSFolderPermissionMaster")
  .items.select(
    "SiteName",
    "DocumentLibraryName",
    "ApprovalType",
    "Level",
    "ApprovalUser/Title",
    "ApprovalUser/EMail"
  )
  .expand("ApprovalUser")
  .filter(`SiteName eq '${selectedSubSite}' and DocumentLibraryName eq 'TransmittalTemplate'`)
  .orderBy("Level", true)
  .getAll();
 
// ✅ Use `hierarchy` (not `approvalHierachy`) for grouping
const groupedHierarchy = hierarchy.reduce((acc: any, item: any) => {
  if (!acc[item.Level]) {
    acc[item.Level] = {
      Level: item.Level,
      ApprovalType: item.ApprovalType,
      Emails: [],
    };
  }
  acc[item.Level].Emails.push(item.ApprovalUser.EMail);
  return acc;
}, {});
 
// ✅ Update state with grouped data
setApprovalHierachy(Object.values(groupedHierarchy));
 
console.log("Grouped Hierarchy:", Object.values(groupedHierarchy));
 
            console.log("Approval Hierarchy:", hierarchy);
        } catch (error) {
            console.error("Error fetching approvals:", error);
        }
    };
 
 
//addhyan - 13/03/2026
 
 
  const handleSubmit = async () => {


    if (!selectedSubSite) {
    Swal.fire({
      icon: "warning",
      title: "Select Department",
      text: "Please select a department before creating the document",
    });
    return; 
  }

  try {



    console.log(fromUser)
 
    Swal.fire({
      title: "Saving...",
      text: "Closing document editor",
      allowOutsideClick: false,
      didOpen: () => Swal.showLoading(),
    });
 
 
 
 
        const siteID = "9ebc2574-a42a-48fa-949b-f6aeac19cfa3";
 
    // STEP 1 — close word editor
    setDocumentUrl("");
    setIsEditorOpen(false);
 
    // STEP 2 — wait for SharePoint lock release
    await new Promise(resolve => setTimeout(resolve, 10000));
 
    // STEP 3 — update metadata
    // const destSitePath = "/sites/AlRostmaniSpfx2/TransmittalTest";
  //addhyan - 13/03/2026
    
    const destSitePath = `/sites/${selectedSite}/${selectedSubSite}`;    //addhyan - 13/03/2026
    const destWeb = Web([sp.web, `${window.location.origin}${destSitePath}`]);
 
          const currentUserEmail = (await destWeb.currentUser()).Email;
      
 
    await destWeb.lists
      .getByTitle("TransmittalTemplate")
      .items.getById(uploadedItemId)
      .update({
        Status: "Pending",
        TransmittalNumber: transmittalNumber,
        To: toUser?.email,
        From: fromUser?.label || "Adhyan",
        Date: date,
        Ref: reference,
       
      });
 
      const destSitePaths = `/sites/${selectedSite}`; //addhyan - 13/03/2026
    const destWebs = Web([sp.web, `${window.location.origin}${destSitePaths}`]);
 
 
 
      await destWebs.lists.getByTitle(`DMS${selectedSubSite}FileMaster`).items.add({
        FileName: String(uploadedFileResponse.data.Name),
        FileSize: String(uploadedFileResponse.data.Length),
        FileVersion: String(uploadedFileResponse.data.MajorVersion),
        CurrentFolderPath: `/sites/${selectedSite}/${selectedSubSite}/TransmittalTemplate`,
        FileUID: String(uploadedFileResponse.data.UniqueId),
        CurrentUser: String(currentUserEmail),
        SiteID: siteID,
        Status: "Pending",
        FilePreviewURL: documentUrl,
        DocumentLibraryName: "TransmittalTemplate",
        SiteName: selectedSubSite,
        MyRequest: true,
        Processname: "New File Request",
        RequestNo: `${"file"}${fileCounter}`,
      });
 
      await destWebs.lists.getByTitle("DMSFileApprovalList").items.add({
        SiteName: selectedSubSite,
        DocumentLibraryName: "TransmittalTemplate",
        RequestedBy: String(currentUserEmail),
        FileName: String(uploadedFileResponse.data.Name),
        FileUID: String(uploadedFileResponse.data.UniqueId),
        FilePreviewUrl: documentUrl,
        Status: "Pending",
        FolderPath: `/sites/${selectedSite}/${selectedSubSite}/TransmittalTemplate`,
        ApproveAction: "Submitted",
        ApprovedLevel: 1,
        RequestNo: `${"file"}${fileCounter}`,
        Processname: "New File Request",
        CurrentLevel: 1,
      });
 
 
 
 
 
 
    Swal.close();
    Swal.fire("Success", "Document Created Successfully", "success");
 
  } catch (error) {
    console.error("Error updating list item:", error);
    Swal.fire("Error", "File still locked. Please try again.", "error");
  }
};
 
  
 
  return (
 
    //addhyan - 13/03/2026
    <>
 
    {/* <div className="col-md-6 mb-3">
      <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
        Select Location
      </label>
      <select
        className="form-control"
        value={selectedSite}
        onChange={(e) => {
          setSelectedSite(e.target.value);
          setSelectedSubSite("");
          resetForm();
          setIsEditorOpen(false);
          setDocumentUrl("");
        }}
      >
        <option value="">Select Location</option>
        <option value="AlRostmaniSpfx2">AlRostmaniSpfx2</option>
        <option value="edcspfx">edcspfx</option>
        <option value="Intranetdemos">intranetdemos</option>
      </select>
    </div> */}

    {/* {selectedSite && (
      <div className="col-md-6 mb-3">
        <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
          Select Department 
        </label>
        <select
          className="form-control"
          value={selectedSubSite}
          onChange={(e) => setSelectedSubSite(e.target.value)}
        >
          <option value="">Select Department</option>
          {(siteSubsiteMap[selectedSite] || []).map((subsite) => (
            <option key={subsite} value={subsite}>
              {subsite}
            </option>
          ))}
        </select>
      </div>
    )} */}

    {/* {selectedSubSite && ( */}
      {/* {selectedSite && ( */}
      <div style={{ fontFamily: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif", padding: '20px' }}>
      <h3 className="font-16 m-0 mb-3">
        New from template: {selectedTemplate?.DocumentCategory}
      </h3>
 
        <div
     
    style={{
       
        border: "none",
       
        cursor: "pointer",  marginTop:"0px",
        fontWeight: "600",
        
        position: isFullScreen ? "fixed" : "static",
        top: isFullScreen ? "6px" : "auto",
        right: isFullScreen ? "96px" : "auto",
        zIndex: 10000,  padding:'6px',
        transition: "background-color 0.3s ease"
    }}
    
    onClick={() => setIsFullScreen(!isFullScreen)}
>
{
  isFullScreen ? (
    <span className="mb-1 mt-2" data-tooltip="Exit Full Screen">
      ✖
    </span>
  ) : (
    <span className="mb-1 mt-2" data-tooltip="Full Screen">
      ⛶
    </span>
  )
}
 
    
</div>
 
      <div className=" mb-3">
        <div className="">
          <div className="row">
            {/* Transmittal Number */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                Transmittal Number
              </label>
              <input
                type="text"
                value={transmittalNumber}
                readOnly
                className="form-control"
                style={{
                  backgroundColor: '#f3f2f1',
                  cursor: 'copy'
                }}
                onClick={(e) => e.currentTarget.select()}
              />
            </div>
 
            {/* Company Dropdown */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                Company <span style={{ color: 'red' }}>*</span>
              </label>
              <select
                className="form-control"
                value={company?.id || ''}
                onChange={(e) => {
                  const selected = companyMasterList.find((c) => c.id === parseInt(e.target.value));
                  setCompany(selected || null);
                }}
              >
                <option value="">Select Company</option>
                {companyMasterList.map((item) => (
                  <option key={item.id} value={item.id}>
                    {item.label}
                  </option>
                ))}
              </select>
            </div>
 
            {/* Classification Dropdown */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                Classification <span style={{ color: 'red' }}>*</span>
              </label>
              <select
                className="form-control"
                value={classification?.id || ''}
                onChange={(e) => {
                  const selected = classificationList.find((c) => c.id === parseInt(e.target.value));
                  setClassification(selected || null);
                }}
              >
                <option value="">Select Classification</option>
                {classificationList.map((item) => (
                  <option key={item.id} value={item.id}>
                   
                    {item.label}
                  </option>
                ))}
              </select>
            </div>
 
            {/* To User Dropdown */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                To <span style={{ color: 'red' }}>*</span>
              </label>
              <select
                className="form-control"
                value={toUser?.id || ''}
                onChange={(e) => {
                  const selected = allUsers.find((u) => u.id === parseInt(e.target.value));
                  setToUser(selected || null);
                }}
              >
                <option value="">Select User</option>
                {allUsers.map((user) => (
                  <option key={user.id} value={user.id}>
                    {user.label}
                  </option>
                ))}
              </select>
            </div>
 
            {/* From User (Read-only) */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                From <span style={{ color: 'red' }}>*</span>
              </label>
              <input
                type="text"
                className="form-control"
                value={fromUser?.label || ''}
                disabled
                style={{ backgroundColor: '#f3f2f1' }}
              />
            </div>
 
            {/* Date */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                Date <span style={{ color: 'red' }}>*</span>
              </label>
              <input
                type="date"
                className="form-control"
                value={date}
                onChange={(e) => setDate(e.target.value)}
              />
            </div>
 
            {/* Reference */}
            <div className="col-md-6 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                Reference
              </label>
              <input
                type="text"
                className="form-control"
                value={reference}
                onChange={(e) => setReference(e.target.value)}
                placeholder="Enter reference"
              />
            </div>
 
            {/* Description */}
            <div className="col-12 mb-3">
              <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
                Description
              </label>
              <textarea
                className="form-control"
                value={description}
                onChange={(e) => setDescription(e.target.value)}
                rows={3}
                placeholder="Enter description"
              />
            </div>
            <div className="col-md-6 mb-3">
  <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
    Select Location
  </label>
  <select
    className="form-control"
    value={selectedSite}
    onChange={(e) => {
      setSelectedSite(e.target.value);
      setSelectedSubSite("");
      resetForm();
      setIsEditorOpen(false);
      setDocumentUrl("");
    }}
  >
    <option value="">Select Location</option>
    <option value="AlRostmaniSpfx2">AlRostmaniSpfx2</option>
    <option value="edcspfx">edcspfx</option>
    <option value="Intranetdemos">intranetdemos</option>
  </select>
</div>

<div className="col-md-6 mb-3">
  <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
    Select Department
  </label>
  <select
    className="form-control"
    value={selectedSubSite}
    disabled={!selectedSite}
    onChange={(e) => setSelectedSubSite(e.target.value)}
  >
    <option value="">Select Department</option>

    {(siteSubsiteMap[selectedSite] || []).map((subsite) => (
      <option key={subsite} value={subsite}>
        {subsite}
      </option>
    ))}
  </select>
</div>
          </div>
        </div>
      </div>
      
 
      <div style={{
                        border: "1px solid #ddd",
                        borderRadius: "4px",
                        overflow: "hidden",
                        boxShadow: "0 2px 10px rgba(0, 0, 0, 0.1)",
                        marginBottom: "20px"
                    }}>

                      
    
{ selectedSubSite && (
  <iframe
                            id="wordEditorFrame"
                            src={documentUrl}
                            style={{
                                  width: isFullScreen ? "100vw" : "100%",
                                height: isFullScreen ? "100vh" : "70vh",
                                border: "none",
                                position: isFullScreen ? "fixed" : "relative",
                                top: isFullScreen ? 0 : "auto",
                                left: isFullScreen ? 0 : "auto",
                                zIndex: isFullScreen ? 9999 : "auto",
                                background: "#fff",
                            }}
                            title="Document Template"
                            allowFullScreen
                        ></iframe>
)}

 
      </div>
 
 { selectedSubSite && (
   <div className="card mar-90">
                    <div className="card-body">
                        <h3 className="mb-1 fw-bold text-dark header-title"
                        >
                            Approval Hierarchy
                        </h3>
                        <table style={{
                            width: "100%",
                            borderCollapse: "collapse",
                            marginTop: "10px"
                        }}>
                            <thead>
                                <tr>
                                    <th style={{
                                        backgroundColor: "#f6f9fc",
                                        color: "6f6f6f",
                                        textAlign: "left",
                                        padding: "12px",
                                        fontSize: "0.9rem"
                                    }}>
                                        User
                                    </th>
                                    <th style={{
                                         backgroundColor: "#f6f9fc",
                                         color: "6f6f6f",
                                        textAlign: "left",
                                        padding: "12px",
                                        fontSize: "0.9rem"
                                    }}>
                                        Level
                                    </th>
                                    <th style={{
                                          backgroundColor: "#f6f9fc",
                                          color: "6f6f6f",
                                        textAlign: "left",
                                        padding: "12px",
                                        fontSize: "0.9rem"
                                    }}>
                                        Approval Type
                                    </th>
                                </tr>
                            </thead>
                            <tbody>
                                {/* {
                                approvalHierachy && approvalHierachy.length > 0 ? (
                                    approvalHierachy.map((item: any) => (
                                        <tr key={item.Id} style={{ borderBottom: "1px solid #ddd" }}>
                                            <td style={{ padding: "12px", fontSize: "0.9rem" }}>{item.ApprovalUser.EMail}</td>
                                            <td style={{ padding: "12px", fontSize: "0.9rem" }}>{item.Level}</td>
                                            <td style={{ padding: "12px", fontSize: "0.9rem" }}>{getApprovalTypeText(item.ApprovalType)}</td>
                                        </tr>
                                    ))
                                )
                                : (
                                    <tr>
                                        <td colSpan={3} style={{ textAlign: "center", padding: "12px", fontSize: "0.9rem" }}>
                                            No approval hierarchy found
                                        </td>
                                    </tr>
                                )} */}
                              {
  approvalHierachy && approvalHierachy.length > 0 ? (
    approvalHierachy.map((item: any) => (
      <tr key={item.Level} style={{ borderBottom: "1px solid #ddd" }}>
        <td style={{ padding: "12px", fontSize: "0.9rem" }}>
          {item.Emails.join(", ")}
        </td>
        <td style={{ padding: "12px", fontSize: "0.9rem" }}>
         <span className="circlelevel">{item.Level}</span>
        </td>
        <td style={{ padding: "12px", fontSize: "0.9rem" }}>
        <span className="appbg">  {getApprovalTypeText(item.ApprovalType)} </span>
        </td>
      </tr>
    ))
  ) : (
    <tr>
      <td colSpan={3} style={{ textAlign: "center", padding: "12px", fontSize: "0.9rem" }}>
        No approval hierarchy found
      </td>
    </tr>
  )
}
 
 
                            </tbody>
                        </table>
                    </div>
                   
                </div>
 
  )}
          
 
      
 
      {/* Action Buttons */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', padding:"120px", gap: '8px' }}>
        <button className="btn btn-secondary" onClick={onCancel} disabled={isLoading}>
          Cancel
        </button>
        <button
         type="button"
          className="btn btn-primary"
          onClick={handleSubmit}
          disabled={isLoading}
        >
          {isLoading ? 'Loading...' : 'Create'}
        </button>
      </div>
    </div>
  
 
 
 
   
 
      </>
  );
};
 
export default TemplateForm;