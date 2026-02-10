import * as React from "react";
import { Modal, Button } from 'react-bootstrap';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Swal from 'sweetalert2';
import "./Modalcss.scss";
import "@pnp/sp/webs";


import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/items";
import "@pnp/sp/security";

import "@pnp/sp/lists";
import "@pnp/sp/files";
// import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/site-users/web";
import { PermissionKind } from "@pnp/sp/security";

interface ShareModalProps {
  show: boolean;
  onClose: () => void;
  file: any;
  context: WebPartContext;
  currentUserEmail: string;
  currentUserTitle: string;
  currentSiteUrl: string;
}

interface SharedUserItem {
  id: string;
  value: string;
  email: string;
}

const ShareModal: React.FC<ShareModalProps> = ({
  show,
  onClose,
  file,
  context,
  currentUserEmail,
  currentUserTitle,
  currentSiteUrl,
}) => {
  const [selectedUsers, setSelectedUsers] = React.useState<SharedUserItem[]>([]);
  const [selectedPermission, setSelectedPermission] = React.useState<string>("");
  const [selectedComment, setSelectedComment] = React.useState<string>("");
  const [userDropdownOptions, setUserDropdownOptions] = React.useState<SharedUserItem[]>([]);
  const [userInput, setUserInput] = React.useState<string>("");
  const [showDropdown, setShowDropdown] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [filePermission, setFilePermission] = React.useState<string>("");
  const userInputRef = React.useRef<HTMLInputElement>(null);
  const dropdownRef = React.useRef<HTMLDivElement>(null);

  const sp = React.useMemo(() => spfi().using(SPFx(context)), [context]);
  console.log("sppp", sp);

  // Fetch site users on modal open
  React.useEffect(() => {
    if (show && file) {
      fetchSiteUsers();
      checkFilePermission();
      console.log("File to share:", file);
      console.log("Current User Email:", currentUserEmail);
      console.log("Current Site URL:", currentSiteUrl);
    }
  }, [show, file]);


  // user fetching function addhyan - 28/01/2026

  const fetchSiteUsers = async () => {
    try {
      const users = await sp.web.siteUsers();
      const usersList = users.map((user) => ({
        id: String(user.Id),
        value: user.Title,
        email: user.Email,
      }));
      setUserDropdownOptions(usersList);

      console.log("Fetched site users:", usersList);
    } catch (error) {
      console.error("Error fetching users:", error);
      
    }
  };

  // user fetching function addhyan - 28/01/2026






  
//   const checkFilePermission = async () => {

//     // console.log("sp", sp);
//     const siteurl = file?.__siteUrl || currentSiteUrl;
//     const siteurls = siteurl.split("/").pop();
//     console.log("Addhyan kumarChecking site URL before spfi:", siteurls);
//     const siteCollectionName = siteurl.split("/sites/")[1].split("/")[0];
//     console.log("Addhyan kumarChecking site Collection Name:", siteCollectionName);
//     const sp = spfi(siteurl).using(SPFx(context));
//     console.log("Addhyan kumarChecking site URL:", siteurl);
//     // const testidsub = await sp.site.openWebById(file.SiteID)
//     // console.log("Addhyan kumarChecking site by id:", testidsub);
//     const filePath = `${file.CurrentFolderPath}/${file.FileName}`;
//       console.log("Addhyan kumarChecking permissions for file path:", filePath);
//       const fileItem = await sp.web.getFileByServerRelativePath(filePath).getItem();
//       console.log("File item fetched:", fileItem);
     

    







    
//     try {
//       const filePath = `${file.CurrentFolderPath}/${file.FileName}`;
//       console.log("Addhyan kumarChecking permissions for file path:", filePath);
//       const fileItem = await sp.web.getFileByServerRelativePath(filePath).getItem();
//       console.log("File item fetched:", fileItem);
//       const filePerms = await fileItem.getCurrentUserEffectivePermissions();
//       console.log("File permissions fetched:", filePerms);

//       const hasFullControl = sp.web.hasPermissions(filePerms, PermissionKind.ManageWeb);
//       console.log("Has Full Control:", hasFullControl);
//       const hasEdit = sp.web.hasPermissions(filePerms, PermissionKind.EditListItems);
//       const hasContribute = sp.web.hasPermissions(filePerms, PermissionKind.AddListItems) && sp.web.hasPermissions(filePerms, PermissionKind.EditListItems);
//       const hasRead = sp.web.hasPermissions(filePerms, PermissionKind.ViewListItems);

//       if (hasFullControl) {
//         setFilePermission("Full Control");
//         console.log("User has Full Control");
//       } else if (hasEdit) {
//         setFilePermission("Edit");
//         console.log("User has Edit permission");
//       } else if (hasContribute) {
//         setFilePermission("Contribute");
//         console.log("User has Contribute permission");
//       } else if (hasRead) {
//         setFilePermission("Read");
//         console.log("User has Read permission");
//       } else {
//         setFilePermission("No Access");
//         console.log("User has No Access");
//       }
//     } catch (error) {
//       console.error("Error checking file permission:", error);
//     }
//   };

const checkFilePermission = async () => {
  if (!file || !context) return;
 
  try {
    // 1. Extract the Sub-site URL from the file path
    // Path: /sites/AlRostmaniSpfx2/TestHub1/Approval Temp Lib/...
    const pathParts = file.CurrentFolderPath.split('/');
    console.log("File Path Parts:", pathParts); 
    // Based on your structure, the sub-site is the 4th element (index 3)
    // sites (1) / AlRostmaniSpfx2 (2) / TestHub1 (3)
    const subSiteUrl = `${window.location.origin}${pathParts.slice(0, 4).join('/')}`;
    console.log("Adjusted Sub-site URL:", subSiteUrl);
 
    // 2. Initialize SP using the dynamic sub-site URL
    const sp = spfi(subSiteUrl).using(SPFx(context));
 
    // 3. Construct the path (must be server-relative)
    const filePath = `${file.CurrentFolderPath}/${file.FileName}`;
    // 4. Get the Item
    // Now that 'sp' points to /TestHub1, .getItem() will succeed
    const fileItem = await sp.web.getFileByServerRelativePath(filePath).getItem();
    // 5. Check Permissions
    const filePerms = await fileItem.getCurrentUserEffectivePermissions();
 
    const hasFullControl = sp.web.hasPermissions(filePerms, PermissionKind.ManagePermissions);
    const hasEdit = sp.web.hasPermissions(filePerms, PermissionKind.EditListItems);
    const hasContribute = sp.web.hasPermissions(filePerms, PermissionKind.AddListItems);
    const hasRead = sp.web.hasPermissions(filePerms, PermissionKind.ViewListItems);
 
    let permissionLevel: string = "No Access";
    if (hasFullControl) permissionLevel = "Full Control";
   
    else if (hasEdit) permissionLevel = "Edit";

    else if (hasContribute) permissionLevel = "Contribute";
    else if (hasRead) permissionLevel = "Read";
 
    setFilePermission(permissionLevel);
    console.log("Determined file permission level:", permissionLevel);
 
  } catch (error: any) {
    console.error("Error details:", error.message);
    // If it still says "does not belong to a list", the URL in step 1 is slightly off
    setFilePermission("Error");
  }
};

  

  const handleUserInputChange = (value: string) => {
    setUserInput(value);
    if (value.trim()) {
      const filtered = userDropdownOptions.filter(user =>
        user.value.toLowerCase().includes(value.toLowerCase()) ||
        user.email.toLowerCase().includes(value.toLowerCase())
      );
      setShowDropdown(true);
    } else {
      setShowDropdown(false);
    }
  };

  const handleSelectUser = (user: SharedUserItem) => {
    if (!selectedUsers.some(u => u.id === user.id)) {
      setSelectedUsers([...selectedUsers, user]);
    }
    setUserInput("");
    setShowDropdown(false);
  };

  const handleRemoveUser = (userId: string) => {
    setSelectedUsers(selectedUsers.filter(u => u.id !== userId));
  };

  // Handle outside click to close dropdown
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node) &&
        userInputRef.current && !userInputRef.current.contains(event.target as Node)) {
        setShowDropdown(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const getPermissionOptions = (): string[] => {
    if (filePermission === "Full Control") {
      return ["Can download and edit (make any changes)", "Can download and view (cannot make changes)", "Cannot download but view (can view but cannot download)"];
    } else if (filePermission === "Contribute" || filePermission === "Edit") {
      return ["Can download and edit (make any changes)", "Can download and view (cannot make changes)", "Cannot download but view (can view but cannot download)"];
    } else if (filePermission === "Read") {
      return ["Cannot download but view (can view but cannot download)"];
    }
    return [];
  };

  const handleShare = async () => {
    if (selectedUsers.length === 0) {
      Swal.fire('Warning', 'Please select at least one user to share with', 'warning');
      return;
    }
    if (!selectedPermission) {
      Swal.fire('Warning', 'Please select a permission type', 'warning');
      return;
    }

    setLoading(true);
    try {
      const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
      const entity = file.SiteName || "Unknown";
      
      // Share file with each selected user
      for (const user of selectedUsers) {
        // Add record to DMSShareWithOtherMaster list
        const sharePayload = {
          FileName: file.FileName,
          FileUID: file.FileUID || file.UniqueId,
          CurrentUser: currentUserEmail,
          CurrentFolderPath: file.CurrentFolderPath,
          SiteName: entity,
          PermissionType: selectedPermission,
          ShareAt: isoDate,
          FileVersion: file.FileVersion || "1",
          FileSize: file.FileSize,
          Status: file.Status || "Active",
          FilePreviewURL: file.FilePreviewURL || "",
          SiteID: file.SiteID || "",
          DocumentLibraryName: file.DocumentLibraryName || "",
          UserID: user.id,
          ShareWithOthers: user.value,
          ShareWithMe: user.email,
        };




        const pathParts = file.CurrentFolderPath.split('/');
    // Based on your structure, the sub-site is the 4th element (index 3)
    // sites (1) / AlRostmaniSpfx2 (2) / TestHub1 (3)
    const subSiteUrl = `${window.location.origin}${pathParts.slice(0, 3).join('/')}`;
    console.log("Adjusted Sub-site URL:", subSiteUrl);
 
    // 2. Initialize SP using the dynamic sub-site URL
    const sp = spfi(subSiteUrl).using(SPFx(context));



        // Add to ShareWithOtherMaster list
        const newItem = await sp.web.lists.getByTitle(`DMSShareWithOtherMaster`).items.add(sharePayload);
        console.log("Share record added:", newItem);

        // Share file with user
        // const filePath = `${file.CurrentFolderPath}/${file.FileName}`;
        // const fileObj = sp.web.getFileByServerRelativePath(filePath);
        
        // If current user has full control, use role assignments
        // if (filePermission === "Full Control") {
        //   const fileItem = await fileObj.getItem();
        //   const itemData = await fileItem.select("HasUniqueRoleAssignments")();
          
        //   if (!itemData.HasUniqueRoleAssignments) {
        //     await fileItem.breakRoleInheritance(true);
        //   }
          
        //   // Map permission string to role type
        //   let roleType = 1073741826; // Default to View
        //   if (selectedPermission.toLowerCase().includes("edit")) {
        //     roleType = 1073741830; // Edit
        //   }
          
        //   await fileItem.roleAssignments.add(Number(user.id), roleType);
        // } else {
        //   // Use roleAssignments for lower permission levels
        //   const fileItem = await fileObj.getItem();
        //   let roleType = 1073741826; // Default to View
        //   if (selectedPermission.toLowerCase().includes("edit")) {
        //     roleType = 1073741830; // Edit
        //   }
        //   await fileItem.roleAssignments.add(Number(user.id), roleType);
        // }
      }

      Swal.fire('Success', `File shared successfully with ${selectedUsers.length} user(s)`, 'success');
      onClose();
    } catch (error) {
      console.error("Error sharing file:", error);
      Swal.fire('Error', 'Failed to share file. Please try again.', 'error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <Modal show={show} onHide={onClose} centered className='sharemodal' size="lg">
      <Modal.Header closeButton>
        <Modal.Title>Share: {file?.FileName}</Modal.Title>
      </Modal.Header>
      <Modal.Body>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
          {/* Column 1: Share With Input */}
          <div>
            <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
              Share With (Name, Group, or Email)
            </label>
            <div style={{ position: 'relative' }}>
              <input
                ref={userInputRef}
                type="text"
                placeholder="Enter a Name, Group, or Email"
                value={userInput}
                onChange={(e) => handleUserInputChange(e.target.value)}
                onFocus={() => setShowDropdown(userInput.trim().length > 0)}
                style={{
                  width: '100%',
                  padding: '10px',
                  fontSize: '14px',
                  borderRadius: '4px',
                  border: '1px solid #ccc',
                  boxSizing: 'border-box'
                }}
              />
              
              {/* User Dropdown */}
              {showDropdown && (
                <div
                  ref={dropdownRef}
                  style={{
                    position: 'absolute',
                    top: '100%',
                    left: 0,
                    right: 0,
                    maxHeight: '200px',
                    overflowY: 'auto',
                    backgroundColor: '#fff',
                    border: '1px solid #ccc',
                    borderRadius: '4px',
                    zIndex: 1000,
                    marginTop: '4px'
                  }}
                >
                  {userDropdownOptions
                    .filter(user =>
                      userInput.trim() === '' ||
                      user.value.toLowerCase().includes(userInput.toLowerCase()) ||
                      user.email.toLowerCase().includes(userInput.toLowerCase())
                    )
                    .map((user) => (
                      <div
                        key={user.id}
                        onClick={() => handleSelectUser(user)}
                        style={{
                          padding: '10px',
                          cursor: 'pointer',
                          borderBottom: '1px solid #f0f0f0',
                          fontSize: '14px'
                        }}
                        onMouseEnter={(e) => {
                          (e.currentTarget as HTMLElement).style.backgroundColor = '#f0f0f0';
                        }}
                        onMouseLeave={(e) => {
                          (e.currentTarget as HTMLElement).style.backgroundColor = '#fff';
                        }}
                      >
                        <div style={{ fontWeight: 500 }}>{user.value}</div>
                        <div style={{ fontSize: '12px', color: '#666' }}>{user.email}</div>
                      </div>
                    ))}
                </div>
              )}
            </div>

            {/* Selected Users Display */}
            {selectedUsers.length > 0 && (
              <div style={{ marginTop: '12px', display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                {selectedUsers.map((user) => (
                  <div
                    key={user.id}
                    style={{
                      backgroundColor: '#e0e0e0',
                      borderRadius: '12px',
                      padding: '6px 12px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px'
                    }}
                  >
                    <span>{user.value}</span>
                    <button
                      type="button"
                      onClick={() => handleRemoveUser(user.id)}
                      style={{
                        background: 'none',
                        border: 'none',
                        cursor: 'pointer',
                        fontSize: '16px',
                        color: '#666',
                        padding: 0,
                        lineHeight: '1'
                      }}
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Column 2: Permission Type Dropdown */}
          <div>
            <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
              Permission Type
            </label>
            <select
              value={selectedPermission}
              onChange={(e) => setSelectedPermission(e.target.value)}
              style={{
                width: '100%',
                padding: '10px',
                fontSize: '14px',
                borderRadius: '4px',
                border: '1px solid #ccc',
                boxSizing: 'border-box'
              }}
            >
              <option value="">Select a Permission</option>
              {getPermissionOptions().map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </select>
          </div>

          {/* Column 3: Comments */}
          <div>
            <label style={{ fontWeight: 600, marginBottom: '8px', display: 'block' }}>
              Comment (Optional)
            </label>
            <textarea
              placeholder="Add a comment (optional)"
              value={selectedComment}
              onChange={(e) => setSelectedComment(e.target.value)}
              style={{
                width: '100%',
                padding: '10px',
                fontSize: '14px',
                borderRadius: '4px',
                border: '1px solid #ccc',
                boxSizing: 'border-box',
                minHeight: '80px',
                fontFamily: 'inherit'
              }}
            />
          </div>
        </div>
      </Modal.Body>
      <Modal.Footer>
        <Button variant="secondary" onClick={onClose}>
          Cancel
        </Button>
        <Button 
          variant="primary" 
          onClick={handleShare}
          disabled={loading || selectedUsers.length === 0 || !selectedPermission}
        >
          {loading ? 'Sharing...' : 'Share'}
        </Button>
      </Modal.Footer>
    </Modal>
  );
};

export default ShareModal;
