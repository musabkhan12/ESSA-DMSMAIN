import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import Swal from 'sweetalert2';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/items";




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


  // Data lists state
  const [allUsers, setAllUsers] = useState<any[]>([]);
  const [companyMasterList, setCompanyMasterList] = useState<any[]>([]);
  const [classificationList, setClassificationList] = useState<any[]>([]);
  const [fileCounter, setFileCounter] = useState<number>(0);
  const [isLoading, setIsLoading] = useState(false);
  const [documentUrl, setDocumentUrl] = useState<string>('');


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
        
        await fetchAllUsers();
        // await fetchDocumentTemplate();
      
      }
    } catch (error) {
      console.error('Error initializing form:', error);
      Swal.fire('Error', 'Failed to initialize form data', 'error');
    } finally {
      setIsLoading(false);
    }
  };

   useEffect(() => {
 copyfile();
 companyclassification();
    
   }, []);




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
    const destSitePath = "/sites/AlRostmaniSpfx2";
    const destWeb = Web([sp.web, `${window.location.origin}${destSitePath}`]);
    const folderPath = `${destSitePath}/TransmittalTemplate`;

    /* ---- UPLOAD FILE ---- */
    const uploadResult = await destWeb
      .getFolderByServerRelativePath(folderPath)
      .files.addUsingPath(newFileName, fileBuffer, { Overwrite: true });

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


 

  const fetchAllUsers = async () => {



     const userdata = await sp.web.currentUser();
     
    try {
      if (!sp) return;

      const userdata = await sp.web.currentUser();
      
      currentUserIDref.current = userdata.Id;
      currentUserEmailRef.current = userdata.Email;
      currentUserTitleRef.current = userdata.Title;

      setFromUser({
        label: userdata.Title,
        value: userdata.Email,
        id: userdata.Id
      });

      const users = await sp.web.siteUsers();
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

  // Validate form


  // Handle form submission
  // const handleSubmit = async () => {
  //   // if (!validateForm()) {
  //   //   return;
  //   // }
  //   try{

  //     const destSitePath = "/sites/AlRostmaniSpfx2";
  //   const destWeb = Web([sp.web, `${window.location.origin}${destSitePath}`]);
  //   const folderPath = `${destSitePath}/TransmittalTemplate`;
  //   console.log("foler path",folderPath);

  //   await destWeb.lists
  // .getByTitle("TransmittalTemplate")   // document library name
  // .items.getById(uploadedItemId)       // file ka ID
  // .update({
  //   TransmittalNumber: transmittalNumber,
  //   To: toUser?.email,
  //    From: fromUser?.email,
  //   Date: date,
  //    Ref: reference,
  // });


   
      
    
  //   // await sp.web.lists
  //   //   .getByTitle("TransmittalTemplate") // document library name
  //   //   .items.getById(uploadedItemId)
  //   //   .update({
  //   //     TransmittalNumber: transmittalNumber,
  //   //     // Company: company?.value,
  //   //     // Classification: classification?.value,
  //   //     To: toUser?.id,
  //   //     From: fromUser?.id,
  //   //     Date: date,
  //   //     Ref: reference,
  //   //     // Description: description
  //   //   });

     

  //     console.log("List item updated with form data");
  //     console.log("List item updated with form data");
  //     console.log("List item updated with form data");

  //   // const formData = {
  //   //   transmittalNumber,
  //   //   toUser,
  //   //   company,
  //   //   classification,
  //   //   fromUser,
  //   //   date,
  //   //   reference,
  //   //   description,
  //   //   template: selectedTemplate
  //   // };

     

    
  //   // console.log('Form Data:', formData.transmittalNumber);
  //   // console.log('Form Data:', formData.toUser);

  //   // console.log('Form Data:', formData);
  //   // console.log('Form Data:', formData);
  //   // console.log('Form Data:', formData);
  //   // console.log('Form Data:', formData);
  //   // onSave(formData);
  //   }
  //   catch(error){
  //     console.error('Error updating list item:', error);
  //   }
    


  // };


  const handleSubmit = async () => {
  try {
    console.log(fromUser)

    Swal.fire({
      title: "Saving...",
      text: "Closing document editor",
      allowOutsideClick: false,
      didOpen: () => Swal.showLoading(),
    });

    // STEP 1 — close word editor
    setDocumentUrl("");
    setIsEditorOpen(false);

    // STEP 2 — wait for SharePoint lock release
    await new Promise(resolve => setTimeout(resolve, 5000));

    // STEP 3 — update metadata
    const destSitePath = "/sites/AlRostmaniSpfx2";
    const destWeb = Web([sp.web, `${window.location.origin}${destSitePath}`]);

    await destWeb.lists
      .getByTitle("TransmittalTemplate")
      .items.getById(uploadedItemId)
      .update({
        TransmittalNumber: transmittalNumber,
        To: toUser?.email,
        From: fromUser?.label || "Adhyan",
        Date: date,
        Ref: reference,
       
      });

    Swal.close();
    Swal.fire("Success", "Document Created Successfully", "success");

  } catch (error) {
    console.error("Error updating list item:", error);
    Swal.fire("Error", "File still locked. Please try again.", "error");
  }
};

  

  return (
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

      </div>

      

      {/* Action Buttons */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
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
  );
};

export default TemplateForm;
