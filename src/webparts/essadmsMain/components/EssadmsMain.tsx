import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IEssadmsMainProps } from "./IEssadmsMainProps";
import PreviewModal from "./previewfile"; // import your modal\
import MetadataModal from "./MetadataModal";
import ManageWorkflow from "./ManageWorkflow";
import DirectDownloader from "./DownloadFile";
import VersionHistoryModal from "./versionhistory"; // import version history modal
import ShareModal from "./ShareModal";
import TemplateForm from "./TemplateForm";
import Swal from 'sweetalert2'
import { Button, Modal, Dropdown, ButtonGroup } from 'react-bootstrap';
import CreateFolder from "./CreateFolder";
import Revoke from "./revoke";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
// srs 20/2/26
import { Web } from "@pnp/sp/webs";
import { AssignFrom } from "@pnp/core";
//aman changes for folder permission manage
import { PermissionKind } from "@pnp/sp/security";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
//aman changes for folder permission manage

import { useMemo } from "react";
import UploadFile from "./UploadFile"; // srs 29/1/26
// sourish 30/9/25
import "../../verticalSideBar/components/VerticalSidebar.scss";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import UserContext from "../../../GlobalContext/context";
import Provider from "../../../GlobalContext/provider";
// Ritik 20/2/26 for manage permission folder
import Select from "react-select";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { icon } from '@fortawesome/fontawesome-svg-core';
import {faRedo, faUser,faFilePdf, faCloud,  faFolder as faFolders,  faGlobeAsia, faLock, faHeart as faHearts, faShareSquare,faTrash,faFile ,faDownload,faHistory,  faShareAlt, faCalendar, faClock, faCheckCircle, faFileWord, faFileExcel,faFileImage, faFileArchive, faFileVideo,faShoppingBag, faSort } from '@fortawesome/free-solid-svg-icons';
import { faFolder, faHeart, faCalendarAlt as faCalendars, faTrashAlt, faEye,  faFileAlt,faEdit, faPaperPlane  } from '@fortawesome/free-regular-svg-icons';
// srs 6/3/26
import { faArchive } from '@fortawesome/free-solid-svg-icons';
// Addhyan 13/4/26
import ShareFileUrlModal from "./ShareFileUrlModal";
// Aman 13/4/26
import BreadcrumbSharePopup from "./BreadcrumbSharePopup";
import FolderSharePopup from "./FolderSharePopup";
import { startOfDay } from "date-fns";
 


// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// let loadfilefromnode = ''
declare global {
  interface Window {
    managePermission?: (folder: any) => void;
     // srs 10/4/26
    ManageFilePermission: (fileId: string, siteId: string, documentLibraryName: string, siteTitle: string) => Promise<void>;
    // srs 15/4/26
    rework: (fileId: any, siteId: any, documentLibrary: any, siteName: any, filePath: any) => Promise<void>;
   }
}
interface TreeNode {
  key: string;
  title: string;
  // type: "site" | "subsite" | "library" | "folder";
  type: "site" | "function" | "subsite" | "library" | "folder";
  children?: TreeNode[];
  hasChildren?: boolean;
  isExpanded?: boolean;
  siteUrl: string;
  libraryTitle?: string;
  folderPath?: string;
  parentKey?: string;
}
// ritik chnages for manage permission folder 29/04/26
interface SharedUser {
  userId: string;
 userEmail: string;  
  userTitle: string;
  roles: string[];
}

interface BreadcrumbItem {
  key: string;
  title: string;
  type: string;
  siteUrl: string;
  libraryTitle?: string;
  folderPath?: string;
}

// Ritik - 20/04/26 start

const FilePermUserSearch: React.FC<{
  users: any[];
  onSelect: (user: { id: string; title: string }) => void;
}> = ({ users, onSelect }) => {
  const [val, setVal] = React.useState('');
  const [open, setOpen] = React.useState(false);
 
  const list = (users || []).filter((u: any) =>
    val.trim() &&
    (
      (u.Title || '').toLowerCase().includes(val.toLowerCase()) ||
      (u.Email || '').toLowerCase().includes(val.toLowerCase())
    )
  );
 
  return (
    <div style={{ position: 'relative' }}>
      <input
        type="text"
        className="form-control"
        placeholder="Type name..."
        value={val}
        style={{ fontSize: '14px' }}
        onChange={(e) => { setVal(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 200)}
      />
      {open && list.length > 0 && (
        <div style={{
          position: 'absolute', top: '100%', left: 0, right: 0,
          background: 'white', border: '1px solid #d9d9d9',
          maxHeight: '180px', overflowY: 'auto',
          zIndex: 9999, boxShadow: '0 4px 12px rgba(0,0,0,0.1)'
        }}>
          {list.map((u: any) => (
            <div
              key={u.Id}
              onMouseDown={() => {
                onSelect({ id: String(u.Id), title: u.Title });
                setVal('');
                setOpen(false);
              }}
              style={{
                padding: '10px', cursor: 'pointer',
                borderBottom: '1px solid #f4f4f4'
              }}
            >
              <div style={{ fontWeight: 600, fontSize: '13px' }}>{u.Title}</div>
              <div style={{ fontSize: '11px', color: '#777' }}>{u.Email}</div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

// Ritik - 20/04/26 end



const ArgPoc = ({ context }: { context: WebPartContext }) => {
  // Template form state
  const [showTemplateForm, setShowTemplateForm] = useState(false);
  const [selectedTemplate, setSelectedTemplate] = useState<any>(null);

  const handleSaveTemplate = (data?: { title: string; description: string; template: any }) => {
    const tpl = data?.template || selectedTemplate;
    const title = data?.title || tpl?.DocumentCategory || "";
    Swal.fire('Saved', `${title} template selected`, 'success');
    setShowTemplateForm(false);
    setSelectedTemplate(null);
  };

  const handleCancelTemplate = () => {
    setShowTemplateForm(false);
    setSelectedTemplate(null);
  };
  const [sp] = useState(() => spfi().using(SPFx(context)));
  const [filesLoadedfromnode, setFilesLoadedloadedfromnode] = useState(false);

  // Aman 13/4/26
   const [breadcrumbShare, setBreadcrumbShare] = useState({
  show: false,
  url: ""
});
 
const [folderShare, setFolderShare] = useState({
  show: false,
  url: ""
});
 
const closeFolderShare = () => {
  setFolderShare({
    show: false,
    url: ""
  });
};
 // Aman 13/4/26 end
  // state variables for tree 
  const [treeData, setTreeData] = useState<TreeNode[]>([]);
  // state variables for files
  const [selectedFiles, setSelectedFiles] = useState<any[]>([]);
  // state variables for breadcrumbs
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([]);
  const [nodeMap, setNodeMap] = useState<Record<string, TreeNode>>({});
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  // state variables for active view
  const [activeView, setActiveView] = useState<string>("");
  // state variables for file upload
  const [showUploadPanel, setShowUploadPanel] = useState(false);
  // state variables for file upload
  const [selectedUploadFiles, setSelectedUploadFiles] = useState<File[]>([]);
  // state variables for file upload progress
  const [uploadProgress, setUploadProgress] = useState<number>(0);
  // state variables for current folder path 
  const [currentFolderPath, setCurrentFolderPath] = useState<string>("");

    const [navItems, setNavItems] = useState<any[]>([]);
  // state variables for current site URL
  const [currentSiteUrl, setCurrentSiteUrl] = useState<string>("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const menuButtonRefs = useRef<(HTMLButtonElement | null)[]>([]); // ritik 18/03/26
  // preview file state
  const [previewFileUrls, setPreviewFileUrls] = useState<string[]>([]);
  // modal state
  const [menuOpenIdx, setMenuOpenIdx] = useState<number | null>(null);
const [currentPage, setCurrentPage] = useState(1);
  // state for selected node
  const [selectedCurrentNode, setSelectedCurrentNode] = useState<TreeNode | null>(null);
  const [showAuditModal, setShowAuditModal] = useState(false);
  const [showShareModal, setShowShareModal] = useState(false);
  const [shareFile, setShareFile] = useState<any>(null);
  // this was by om revoke user permission
  //revoke state
  const [showRevokeModal, setShowRevokeModal] = useState(false);
  const [selectedFolderForRevoke, setSelectedFolderForRevoke] = useState<any>(null);
  const [revoke, setAcessFile] = useState<any>(null);
  const [sharedUsers, setSharedUsers] = useState<SharedUser[]>([]);

  // version history modal state
  const [showVersionModal, setShowVersionModal] = useState(false);
  // modal file state
  const [modalFile, setModalFile] = useState<any>(null);
  // create folder modal state
  const [activeComponent, setActiveComponent] = useState(false);


  // Aman 20/4/26 start
const isCreateAllowed = React.useMemo(() => {
  return !!selectedCurrentNode && selectedCurrentNode.type !== "site";
}, [selectedCurrentNode]);
 
// Aman 20/4/26
const isUploadAllowed = React.useMemo(() => {
  return (
    !!selectedCurrentNode &&
    (selectedCurrentNode.type === "library" ||
     selectedCurrentNode.type === "folder")
  );
}, [selectedCurrentNode]);


// aman - 20/04/26 end
  // preview file state
  const [previewFile, setPreviewFile] = useState<any>(null);
  // preview modal state
  const [showPreviewModal, setShowPreviewModal] = useState(false);

  const [previewKey, setPreviewKey] = useState(0);

  //Metadata state Ritik 29/01/2026
  const [showMetadataModal, setShowMetadataModal] = useState(false);
  const [selectedFileForMetadata, setSelectedFileForMetadata] = useState<any>(null);

  // srs 20/2/26
  const [showRenameMetadataModal, setShowRenameMetadataModal] = useState(false);
  const [selectedFileForRenameetadata, setSelectedFileForRenameMetadata] = useState<any>(null);
  const [existingColumns, setExistingColumns] = useState<any[]>([]);
const [loadingColumns, setLoadingColumns] = useState(false);

  //sourish 21/8/25
  const [renameModalOpen, setRenameModalOpen] = useState(false);
  const [renameValue, setRenameValue] = useState("");

  // ===== File Rename specific state =====
  const [renameFileModalOpen, setRenameFileModalOpen] = useState(false);
  const [renameFileValue, setRenameFileValue] = useState("");

  // ===== Manage Workflow (ROOT FOLDER ONLY) =====
  const [showManageWorkflow, setShowManageWorkflow] = useState<boolean>(false);
  const [selectedWorkflowFolder, setSelectedWorkflowFolder] = useState<any>(null);
  //  setting direct file download state
  const [directDownloadFile, setDirectDownloadFile] = React.useState<any | null>(null);

  // ====== AUDIT HISTORY STATES (added) ======
  const [auditVersions, setAuditVersions] = useState<any>({ Metadata: {}, Versions: [] });
  const [auditLoading, setAuditLoading] = useState<boolean>(false);
  //  manag folder permission state
  const [showPermissionModal, setShowPermissionModal] = useState(false);
  const [showManagePermissionModal, setShowManagePermissionModal] = useState(false);
  const [selectedFolder, setSelectedFolder] = useState<any | null>(null);
  const [permissionUsers, setPermissionUsers] = useState<{ principalId: number; principalTitle: string; roles: string[] }[]>([]);
  const [newUser, setNewUser] = useState<string>("");
  const [newPermission, setNewPermission] = useState<string>("Read");
  const [mpHasUnique, setMpHasUnique] = useState<boolean>(false);
  const [mpCanManage, setMpCanManage] = useState<boolean>(false);
  const [mpLoading, setMpLoading] = useState<boolean>(false);
  const [mpError, setMpError] = useState<string>("");
  const [mpSiteUsers, setMpSiteUsers] = useState<{ id: number; title: string; email: string; loginName: string }[]>([]);
//20/04/26 Added state for filter in folder hirarchy in list view 
  const [columnSearch, setColumnSearch] = useState<{name: string; size: string; library: string; status: string}>({
    name: '', size: '', library: '', status: ''
  });
  
// Ritik 20/2/26 for manage permission folder
  const [rowsForPermission, setRowsForPermission] = useState<{id: number; selectedUserForPermission: any[]; selectedPermission: any}[]>([
  { id: 0, selectedUserForPermission: [], selectedPermission: "" }
]);
const [folderPrivacyTableData, setFolderPrivacyTableData] = useState<any[]>([]);
const [folderPrivacyUsers, setFolderPrivacyUsers] = useState<any[]>([]);
const [togglePermission, setTogglePermission] = useState<string | undefined>();
const [fpIsPrivateId, setFpIsPrivateId] = useState<number | null>(null);
const [fpIsPrivateLibId, setFpIsPrivateLibId] = useState<number | null>(null);
const [fpPathState, setFpPathState] = useState<string>('');
const [fpErrors, setFpErrors] = useState<{[key: number]: any}>({});
const fpCurrentUserEmail = useRef<string>('');
const [fpCurrentPage, setFpCurrentPage] = useState(1);
const fpItemsPerPage = 10;
// Ritik 20/2/26

const [viewCounts, setViewCounts] = useState<Record<string, number>>({});

// Sorting states
const [sortColumn, setSortColumn] = useState<string>("");
const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");

// add useState for share file url modal - Addhyan 13/4/26 start
  const [showShareUrlModal, setShowShareUrlModal] = useState(false);
  const [shareUrlFile, setShareUrlFile] = useState<any>(null);


  // ritik - 20/04/26 start
const [showFilePermModal, setShowFilePermModal] = useState(false);
const [filePermState, setFilePermState] = useState<any>(null);
  // ritik - 20/04/26 end
  const isDeepLinkPreview = useRef<Boolean>(false);

  useEffect(() => {
  const params = new URLSearchParams(window.location.search);
  if (params.get("essaPreview") !== "1") return;
 
  const serverRelativeUrl = params.get("essaFileUrl") || "";
  const fileName          = params.get("essaFileName") || "";
  const siteUrl           = params.get("essaSiteUrl") || "";
  const filePreviewURL    = params.get("essaFilePreviewURL") || "";
  const siteName          = params.get("essaSiteName") || "";
  const folderPath        = params.get("essaFolderPath") || "";
 
  if (!serverRelativeUrl && !filePreviewURL) {
    console.warn("[DeepLink] essaPreview=1 found but no file URL — ignoring.");
    return;
  }
 
  const reconstructedFile: any = {
    ServerRelativeUrl: serverRelativeUrl,
    FileName:          fileName,
    Name:              fileName,
    __siteUrl:         siteUrl,
    SiteName:          siteName,
    CurrentFolderPath: folderPath,
    FilePreviewURL:    filePreviewURL || null,
  };
 
  console.log("[DeepLink] Auto-opening PreviewModal for:", reconstructedFile);
 
  // ── Mark that we are in deep-link preview mode ──────────────────────
  // loadViewData will check this ref and skip resetting showPreviewModal
  isDeepLinkPreview.current = true;
 
  // Open preview after component fully mounts + loadRootSites completes
  // NOTE: No replaceState here — initialize() needs the URL intact to detect deep link
  // URL will be cleaned after preview is confirmed open
  setTimeout(() => {
    setPreviewFile(reconstructedFile);
    setShowPreviewModal(true);
    console.log("[DeepLink] PreviewModal opened ✅");
 
    // Clean URL AFTER preview is open (so refresh doesn't re-trigger)
    try {
      window.history.replaceState(null, "", window.location.pathname + window.location.hash);
    } catch (_) { /* non-critical */ }
 
    // After 5 seconds, release the guard so normal navigation works again
    // setTimeout(() => {
    //   isDeepLinkPreview.current = false;
    // }, 5000);
  }, 1500); // wait longer to ensure loadRootSites + initialize() both complete first
 
}, []); // run once on mount
  // Addhyan 13/4/26 end 
 
   // add useState for share file url modal - addhyan /10/04/26 start
  // ✅ Load all view counts on initial page load
// srs 12/3/26 correct code if count mismatch in archival issue in current code data duplicate
//   const getUniqueFiles = (files: any[]) => {
//   const seen = new Set();
//   return files.filter((item: any) => {
//     const uid = item.FileUID || item.UniqueId || item.Id;
//     if (!uid || seen.has(uid)) return false;
//     seen.add(uid);
//     return true;
//   });
// };
// useEffect(() => {
//   const loadAllCounts = async () => {
//     const [myReq, myFav, myFol, sharedMe, sharedOther, recycle, archivedFiles] = await Promise.all([
//       loadViewData("MyRequest"),
//       loadViewData("MyFavourite"),
//       loadViewData("MyFolders"),
//       loadViewData("SharedWithMe"),
//       loadViewData("SharedWithOthers"),
//       loadViewData("RecycleBin"),
//       // srs 6/3/26 
//       loadViewData("ArchivedFiles"),
//     ]);
 
//     setViewCounts({
//       "My request": myReq.length,
//       "My favourite": myFav.length,
//       "My Folders": myFol.length,
//       "Share with me": sharedMe.length,
//       "Share with other": sharedOther.length,
//       "Recycle bin": recycle.length,
//       //srs 6/3/26
//       "Archived Files": archivedFiles.length,
//       // srs 12/3/26 correct code if count mismatch in archival issue in current code data duplicate if same file name present with same fileUID
//       // "My request": getUniqueFiles(myReq).length,
//       // "My favourite": getUniqueFiles(myFav).length,
//       // "My Folders": getUniqueFiles(myFol).length,
//       // "Share with me": getUniqueFiles(sharedMe).length,
//       // "Share with other": getUniqueFiles(sharedOther).length,
//       // "Recycle bin": getUniqueFiles(recycle).length,
//       // "Archived Files": getUniqueFiles(archivedFiles).length,
//     });
//   };
 
//   loadAllCounts();
// }, []); // ✅ Empty array = runs once on mount

// ritik - 10/04/26 start

const refreshViewCount = async (viewType: ViewType) => {
  const viewKey = viewType === "MyRequest" ? "My request"
    : viewType === "MyFavourite" ? "My favourite"
    : viewType === "MyFolders" ? "My Folders"
    : viewType === "SharedWithMe" ? "Share with me"
    : viewType === "SharedWithOthers" ? "Share with other"
    : viewType === "ArchivedFiles" ? "Archived Files"
    : "Recycle bin";
 
  try {
    const files = await loadViewData(viewType);
    const seen = new Set();
    const unique = files.filter((item: any) => {
      const uid = item.FileUID || item.UniqueId || item.Id;
      if (seen.has(uid)) return false;
      seen.add(uid);
      return true;
    });
    // Har view ka count aate hi set ho jata hai - user ko gradually dikhega
    setViewCounts(prev => ({ ...prev, [viewKey]: unique.length }));
  } catch (e) {
    console.error(`Count load failed for ${viewType}:`, e);
  }
};
 
useEffect(() => {
  // Sequential load - throttle safe, counts gradually appear karte hain
  const loadAllCountsSequentially = async () => {
    await refreshViewCount("MyRequest");
    await refreshViewCount("MyFavourite");
    await refreshViewCount("MyFolders");
    await refreshViewCount("SharedWithMe");
    await refreshViewCount("SharedWithOthers");
    await refreshViewCount("RecycleBin");
    await refreshViewCount("ArchivedFiles");
  };
  loadAllCountsSequentially();
}, []); // mount pe ek baar

// ritik - 10/04/26 end

//  type ViewType = "MyRequest" | "MyFavourite" | "MyFolders" | "SharedWithMe" | "SharedWithOthers" | "RecycleBin";
 // srs 6/3/26
  type ViewType = "MyRequest" | "MyFavourite" | "MyFolders" | "SharedWithMe" | "SharedWithOthers" | "RecycleBin" | "ArchivedFiles";
// Add type definition type ViewType = "MyRequest" | "MyFavourite" | "MyFolders" | "SharedWithMe" | "SharedWithOthers" | "RecycleBin";
  
  //Privew file tittle and path state declare Ritik 29/01/2026
  //Aman Manage Permission Folder 
  // simple typeahead states for Manage Permission user picker (no external libs)
  const [mpUserSuggestions, setMpUserSuggestions] = useState<{ id: number; title: string; email: string; loginName: string }[]>([]);
  const [showUserSuggestions, setShowUserSuggestions] = useState<boolean>(false);
  const [newUserDisplay, setNewUserDisplay] = useState<string>("");
  const userSuggestRef = useRef<HTMLDivElement | null>(null);

  // hide suggestions on outside click
  useEffect(() => {
    const handler = (ev: any) => {
      if (userSuggestRef && userSuggestRef.current && !userSuggestRef.current.contains(ev.target)) {
        setShowUserSuggestions(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  // 1. Add useEffect to close menu when switching sections/views
useEffect(() => {
  setMenuOpenIdx(null);
}, [activeView, currentPage, selectedFiles]);
 
// close modals on navigation (Ritik 17/03/2026)
useEffect(() => {
  setShowPermissionModal(false);
  setShowManageWorkflow(false);
  setShowMetadataModal(false);
  setShowRenameMetadataModal(false);
}, [activeView, selectedCurrentNode]);
// ✅ END close all modal when clicking outside the modal
 
// aman 20/04/26 - start 

useEffect(() => {
  setActiveComponent(false);
  setShowUploadPanel(false);
}, [selectedCurrentNode]);

// aman 20/04/26 - end  



// 2. Add click outside handler
useEffect(() => {
  const handleClickOutside = (event: MouseEvent) => {
    const target = event.target as HTMLElement;
    // Check if click is outside three-dot menu
    const isInsideMenu = target.closest('.dotbutton') || 
                         target.closest('.dotbutton2') || 
                         target.closest('.internalbutton');
    if (!isInsideMenu && menuOpenIdx !== null) {
      setMenuOpenIdx(null);
    }
  };
 
  document.addEventListener('mousedown', handleClickOutside);
  return () => document.removeEventListener('mousedown', handleClickOutside);
}, [menuOpenIdx]);

  {/* sourish 30/9/25 */ }

  const { useHide }: any = React.useContext(UserContext);
  // sourish 3/10/25
  const [activeLayout, setActiveLayout] = useState<'grid' | 'list'>('grid');
  // Pagination state

  // const pageSize = 12;

  // abhay change for seach in all tabs
  useEffect(() => {
    setCurrentPage(1); // reset to page 1 whenever selectedFiles changes
  }, [selectedFiles, activeView, activeLayout, currentFolderPath]);


  //previous working code before add search on all tabs 
  //   let location: string = "";
  //   const paginatedFiles = useMemo(() => {
  //   const start = (currentPage - 1) * pageSize;
  //   const end = start + pageSize;
  //   return selectedFiles.slice(start, end);
  // }, [selectedFiles, currentPage]);

  //--------------undodelete function starts---------------
  // const handleUndoDelete = async (file: any) => {
  //   let undoFailed = false;
  //     try {
  //       const siteUrl: string = `${file.__siteUrl}/${file.SiteName}`; // here we are trying to get our subsite context
  //       const listTitle: string = file.__fileMasterList || "FileMaster";
  //       const id: number = file.Id || file.ID;
  //       console.log("FILE TO DELETE:", file);
  //       console.log("Undo Restore Triggered for");
  //       console.log("Site URL ", siteUrl);
  //       console.log("list title", listTitle);
  //       console.log("item id", id);
  //       debugger;
  //       try {
  //     const spnew = spfi(siteUrl).using(SPFx(context));
  //     console.log("file.CurrentFolderPath", file.CurrentFolderPath);
  //       const fileItem = await spnew.web.getFileByServerRelativePath(`${file.CurrentFolderPath}/${file.FileName}`).getItem();
  //     let payload:any={
  //       IsDeleted:null
  //     }
  //       const itemData = await fileItem.update(payload)
  //     console.log("column updated successfully",itemData);
  //      } catch (err) {
  //     console.error("Error updating file in document library:", err);
  //     Swal.fire({
  //   title: 'Error!',
  //   text: 'File restore failed. Please try again.' + err,
  //   icon: 'error',
  //   confirmButtonText: 'Cool'
  //    })
  //       }

  //       try {
  //         const siteurl = `${file.__siteUrl}`  // here we are trying to get our site collection context
  //       const siteSP = spfi(siteurl).using(SPFx(context));
  //       await siteSP.web.lists.getByTitle(listTitle).items.getById(id).update({ IsDeleted: null });

  //         const refreshed = await loadViewData("RecycleBin");
  //       if (refreshed) {
  //         setSelectedFiles([...refreshed]); // spread → force re-render
  //       }

  //       } catch (err) {
  //   console.error("Error updating file metadata:", err);
  //     Swal.fire({
  //   title: 'Error!',
  //   text: 'File restore failed. Please try again.' + err,
  //   icon: 'error',
  //   confirmButtonText: 'Cool'
  //    })
  //       }


  //     } catch (e) {
  //       console.error("Undo (restore) failed:", e);
  //     }
  //   }

  // Undo delete function revised by Aman 
  const handleUndoDelete = async (file: any) => {
    let undoFailed = false;

    try {
      const siteUrl: string = `${file.__siteUrl}/${file.SiteName}`;
      const listTitle: string = file.__fileMasterList || "FileMaster";
      const id: number = file.Id || file.ID;

      console.log("Undo Restore Triggered for:", file);

      //  Restore in Document Library
      try {
        const spnew = spfi(siteUrl).using(SPFx(context));
        const fileItem = await spnew.web
          .getFileByServerRelativePath(`${file.CurrentFolderPath}/${file.FileName}`)
          .getItem();

        await fileItem.update({ IsDeleted: null });
        console.log("Document library restored");
      } catch (err) {
        console.error("Error restoring file in document library:", err);
        undoFailed = true;
      }

      //  Restore in FileMaster list
      try {
        const siteSP = spfi(file.__siteUrl).using(SPFx(context));
        await siteSP.web.lists
          .getByTitle(listTitle)
          .items.getById(id)
          .update({ IsDeleted: null });

        const refreshed = await loadViewData("RecycleBin");
        if (refreshed) {
          setSelectedFiles([...refreshed]);
        }

        console.log("FileMaster restored");
      } catch (err) {
        console.error("Error restoring file metadata:", err);
        undoFailed = true;
      }

      // Show error ONLY if undo actually failed
      if (undoFailed) {
        Swal.fire({
          title: "Error!",
          text: "File restore failed. Please try again.",
          icon: "error",
          confirmButtonText: "OK"
        });
      }

    } catch (e) {
      console.error("Undo (restore) failed:", e);
      Swal.fire({
        title: "Error!",
        text: "File restore failed. Please try again.",
        icon: "error",
        confirmButtonText: "OK"
      });
    }
  };

// srs 6/3/26
//   const handleUndoArchive = async (file: any) => {
//     let itemId = file.Id || file.ID;
//   try {
//     // 1. Target the specific item by ID
//     await sp.web.lists
//       .getByTitle("EssaArchivallist")
//       .items
//       .getById(itemId)
//       .update({
//         Isarchive: false
//       });

//     console.log(`Item ${itemId} has been successfully restored.`);
    
//     // 2. Trigger a refresh of your UI/state here if necessary
//     // e.g., fetchArchiveItems();

//   } catch (error) {
//     console.error("Error updating Isarchive status:", error);
//     alert("Failed to restore the item. Please try again.");
//   }
// };

// srs 31/3/26

const handleUndoArchive = async (file: any) => {
  // Get the ID regardless of casing
  const itemId = file.Id || file.ID;

  try {
    // 1. Update the SharePoint List Item
    await sp.web.lists
      .getByTitle("EssaArchivallist")
      .items
      .getById(itemId)
      .update({
        Isarchive: false
      });

    // 2. Show Success Popup
    Swal.fire({
      title: 'Success!',
      text: 'Item has been unarchived successfully.',
      icon: 'success',
      confirmButtonColor: '#3085d6',
      confirmButtonText: 'OK'
    }).then((result) => {
      /* 3. Reload the page only after the "OK" button is clicked */
      if (result.isConfirmed) {
        window.location.reload();
      }
    });

  } catch (error) {
    console.error("Error updating Isarchive status:", error);
    
    // Show Error Popup if the update fails
    Swal.fire({
      title: 'Error!',
      text: 'Failed to restore the item. Please try again.',
      icon: 'error',
      confirmButtonText: 'Close'
    });
  }
};

  //-----------------Audit History function -----------------

  const fetchAuditHistory = async (file: any) => {
    setAuditLoading(true);
    try {
      const siteUrl: string = `${file.__siteUrl}/${file.SiteName}`;
      const listTitle: string = file.__fileMasterList || "FileMaster";
      const id: number = file.Id || file.ID;

      const spnew = spfi(siteUrl).using(SPFx(context));
      const versions = await spnew.web.lists.getByTitle(listTitle).items.getById(id).versions();
      setAuditVersions({ Metadata: {}, Versions: versions });
    } catch (err) {
      console.error("Error fetching audit history:", err);
      Swal.fire({
        title: 'Error!',
        text: 'Failed to fetch audit history. Please try again.' + err,
        icon: 'error',
        confirmButtonText: 'Cool'
      });
    } finally {
      setAuditLoading(false);
    }
  };

  //-----------------Audit History function ends---------------

  const handleRenameFile = (file: any) => {
    setRenameFileValue(file.FileName);
    setModalFile(file);
    setRenameFileModalOpen(true);
  };

  // ----- Rename File Submit Handler -----
  const submitRenameFile = async () => {
    if (!modalFile || !renameFileValue) return;

    try {
      // 1. Rename in document library
      const siteUrl = `${modalFile.__siteUrl}/${modalFile.SiteName}`;
      const spSubsite = spfi(siteUrl).using(SPFx(context));

      const oldPath =
        `${modalFile.CurrentFolderPath}/${modalFile.FileName}`;

      const newPath =
        `${modalFile.CurrentFolderPath}/${renameFileValue}`;

      await spSubsite.web
        .getFileByServerRelativePath(oldPath)
        .moveByPath(newPath, true);

      // 2. Update ONLY FileName in FileMaster

      const rootSP = spfi(modalFile.__siteUrl).using(SPFx(context));
      await rootSP.web.lists
        .getByTitle(modalFile.__fileMasterList)
        .items.getById(modalFile.Id)
        .update({
          FileName: renameFileValue,
        });

      // 3. Refresh list
      const refreshed = await loadViewData("MyRequest");
      if (refreshed) {
        setSelectedFiles([...refreshed]);
      }

      // 4. Reset preview state (important)
      setShowPreviewModal(false);
      setPreviewFile(null);
      setPreviewKey(prev => prev + 1);

      // 5. Close modal
      setRenameFileModalOpen(false);
      setRenameFileValue("");
      setModalFile(null);

      Swal.fire("Success", "File renamed successfully", "success");
    } catch (err) {
      console.error("Rename File failed:", err);
      Swal.fire("Error", "Rename File failed", "error");
    }
  };



  /** ----------------- helpers (no logic change, just safer operations + logs) ----------------- */

  const findChildByTitleCI = (parent: TreeNode | undefined, title: string) => {
    if (!parent || !parent.children) return undefined;
    const t = (title || "").trim().toLowerCase();
    return parent.children.find((c) => (c.title || "").trim().toLowerCase() === t);
  };

  const ensureExpanded = async (node: TreeNode) => {
    if (!node) return;
    if (!node.isExpanded) {
      console.log("[ensureExpanded] Expanding node:", { key: node.key, title: node.title, type: node.type });
      await toggleNode(node);
    } else {
      console.log("[ensureExpanded] Already expanded:", { key: node.key, title: node.title, type: node.type });
    }
  };

  /** ---------------------------------- effect: initial load ---------------------------------- */
  const [pendingPath, setPendingPath] = useState<string[] | null>(null);



  // useEffect(() => {
    // const initialize = async () => {
    //   console.log("[initialize] Start");
    //   await loadRootSites();

    //   // Check URL hash on initial load
    //   const rawHash = window.location.hash.substring(1);
    //   const hash = decodeURIComponent(rawHash || "");
    //   console.log("[initialize] Hash:", rawHash, "decoded:", hash);

    //   if (hash) {
    //     setPendingPath(hash.split("/"));
    //   } else {
    //     handleViewButtonClick("My request");
    //   }
    //   setIsInitialLoad(false);
    //   console.log("[initialize] Complete");
    // };

    // initialize();


    // Addhyan 13/4/26
    useEffect(() => {
  const initialize = async () => {
    console.log("[initialize] Start");
    await loadRootSites();
 
    // ── Check karo kya yeh deep link se open hua hai ──
    const params = new URLSearchParams(window.location.search);
    const isDeepLink = params.get("essaPreview") === "1";
 
    // Check URL hash
    const rawHash = window.location.hash.substring(1);
    const hash = decodeURIComponent(rawHash || "");
    console.log("[initialize] Hash:", rawHash, "decoded:", hash, "isDeepLink:", isDeepLink);
 
    if (isDeepLink) {
      // Deep link hai — "My request" load mat karo, preview handle karega
      console.log("[initialize] Deep link detected — skipping handleViewButtonClick");
      // (preview wala useEffect neeche handle karega)
    } else if (hash) {
      setPendingPath(hash.split("/"));
    } else {
      handleViewButtonClick("My request");
    }
 
    setIsInitialLoad(false);
    console.log("[initialize] Complete");
  };
 
  initialize();

    // Listen for browser hash changes
    const handleHashChange = () => {
      const rawHash = window.location.hash.substring(1);
      const hash = decodeURIComponent(rawHash || "");
      console.log("[hashchange] New hash:", rawHash, "decoded:", hash);
      if (hash) {
        setPendingPath(hash.split("/"));
      } else {
        handleViewButtonClick("My request");
      }
    };

    window.addEventListener("hashchange", handleHashChange);
    return () => window.removeEventListener("hashchange", handleHashChange);
  }, []);

  useEffect(() => {
    if (pendingPath && treeData.length) {
      console.log("[pendingPath effect] Navigating to:", pendingPath);
      navigateToPath(pendingPath);
      setPendingPath(null); // prevent repeat
    }
  }, [pendingPath, treeData]);

  // aman changes for folder permission manage 
  // useEffect(() => {
  //   window.managePermission = (folder: any) => {
  //     setSelectedFolder(folder);
  //     setShowPermissionModal(true);
  //   };
  //   return () => {
  //     if (window.managePermission) window.managePermission = undefined;
  //   };
  //   // eslint-disable-next-line react-hooks/exhaustive-deps
  // }, []);

// Ritik 20/2/26 for manage permission folder
    useEffect(() => {
    window.managePermission = (folder: any) => {
      fpOpenPermission(folder);
    };
    return () => {
      if (window.managePermission) window.managePermission = undefined;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  //addhyan chanes on three dot should hide on click anywhere outside oif screen
  // three dot menu 

  //   useEffect(() => {
  //   if (menuOpenIdx !== null) {
  //     const handleClickOutside = (event: MouseEvent) => {
  //       // Only close if the click is outside any menu
  //       const menus = document.querySelectorAll('.three-dot-menu');
  //       let clickedInside = false;
  //       menus.forEach(menu => {
  //         if (menu.contains(event.target as Node)) {
  //           clickedInside = true;
  //         }
  //       });
  //       if (!clickedInside) {
  //         setMenuOpenIdx(null);
  //       }
  //     };
  //     document.addEventListener('mousedown', handleClickOutside);
  //     return () => {
  //       document.removeEventListener('mousedown', handleClickOutside);
  //     };
  //   }
  // }, [menuOpenIdx]);



  const loadRootSites = async () => {
    try {
      console.log("[loadRootSites] Fetching master lists...");
      const [siteList, subsiteList] = await Promise.all([
        // sp.web.lists.getByTitle("MasterSiteCollection").items.select("Id", "Title", "SiteURL").top(5000)(),
        // srs 18/3/26
        sp.web.lists.getByTitle("MasterSiteCollection").items
    .select("Id", "Title", "SiteURL", "IsActive")
    .filter("IsActive eq 1")
    .top(5000)(),
        // sp.web.lists.getByTitle("MasterSiteURL").items.select("Id", "Description", "Title", "SiteURL", "Active", "SiteID", "FileMasterList", "Function").top(5000)(),
      sp.web.lists.getByTitle("MasterSiteURL").items
  .select("Id", "Description", "Title", "SiteURL", "Active", "SiteID", "FileMasterList", "Function")
  .filter("Active eq 'Yes'")
  .top(5000)(),
      ]);

      debugger; 
    
    console.log("MasterSiteCollection items:", siteList);
    console.log("MasterSiteURL items:", subsiteList);
      const nodes: TreeNode[] = siteList.map((site: any) => {
        const siteUrl = site.SiteURL?.trim();
        const siteSubs = subsiteList.filter((s: any) => {
          const subUrl = s.SiteURL?.trim().toLowerCase();
          const parentUrl = siteUrl?.toLowerCase();
          if (!subUrl || !parentUrl) return false;
          return subUrl === parentUrl || subUrl.startsWith(parentUrl + "/");
        });

        // Group subsites by function, but only include subsites that have a function
        const subsitesWithFunctions = siteSubs.filter(sub => sub.Function);
        const subsitesWithoutFunctions = siteSubs.filter(sub => !sub.Function);

        const groupedByFunction = subsitesWithFunctions.reduce((acc: any, sub: any) => {
          const func = sub.Function;
          if (!acc[func]) acc[func] = [];
          acc[func].push(sub);
          return acc;
        }, {});

        const functionNodes: TreeNode[] = Object.entries(groupedByFunction).map(([funcName, subs]: [string, any[]]) => ({
          key: `func-${site.Id}-${funcName}`,
          title: funcName,
          type: "function",
          siteUrl: siteUrl,
          hasChildren: true,
          parentKey: `site-${site.Id}`,
          children: (subs as any[]).map((sub: any) => ({
            key: `subsite-${sub.Id}`,
            title: sub.Title,
            type: "subsite",
            siteUrl: sub.SiteURL?.trim(),
            hasChildren: true,
            parentKey: `func-${site.Id}-${funcName}`,
          })),
        }));

        // Create the site node with both function nodes and direct subsites
        const siteNode: TreeNode = {
          key: `site-${site.Id}`,
          title: site.Title,
          type: "site",
          siteUrl: siteUrl,
          hasChildren: functionNodes.length > 0 || subsitesWithoutFunctions.length > 0,
          children: [
            ...functionNodes,
            ...subsitesWithoutFunctions.map((sub: any) => ({
              key: `subsite-${sub.Id}`,
              title: sub.Title,
              type: "subsite" as TreeNode["type"], // Add type guard here
              siteUrl: sub.SiteURL?.trim(),
              hasChildren: true,
              parentKey: `site-${site.Id}`,
            }))
          ],
        };

        return siteNode;
      });

      console.log("[loadRootSites] Root sites loaded:", nodes);
      setTreeData(nodes);
      buildNodeMap(nodes);
    } catch (error) {
      console.error("[loadRootSites] Error:", error);
    }
  };
  /** -------------------------------- build node map (same) ----------------------------------- */
  const buildNodeMap = (nodes: TreeNode[], map: Record<string, TreeNode> = {}) => {
    nodes.forEach((node) => {
      map[node.key] = node;
      if (node.children) {
        buildNodeMap(node.children, map);
      }
    });
    console.log("[buildNodeMap] map keys:", Object.keys(map).length);
    setNodeMap(map);
    return map;
  };

  const getNodePath = (nodeKey: string): BreadcrumbItem[] => {
    const path: BreadcrumbItem[] = [];
    let currentNode = nodeMap[nodeKey];
    while (currentNode) {
      path.unshift({
        key: currentNode.key,
        title: currentNode.title,
        type: currentNode.type,
        siteUrl: currentNode.siteUrl,
        libraryTitle: currentNode.libraryTitle,
        folderPath: currentNode.folderPath,
      });
      if (currentNode.parentKey) {
        currentNode = nodeMap[currentNode.parentKey];
      } else {
        break;
      }
    }
    console.log("[getNodePath]", path.map((p) => `${p.type}:${p.title}`).join(" / "));
    return path;
  };

  const updateUrl = (node: TreeNode) => {
    const path = getNodePath(node.key)
      .map((item) => encodeURIComponent(item.title))
      .join("/");
    console.log("[updateUrl] ->", `#${path}`);
    window.history.pushState(null, "", `#${path}`);
  };
  const navigateToPath = async (pathTitles: string[]) => {
    console.log("[navigateToPath] raw segments:", pathTitles);
    if (pathTitles.length === 0 || !treeData.length) return;

    const decodedTitles = pathTitles.map((t) => decodeURIComponent(t)).filter((t) => t.trim() !== "");
    let currentNode = treeData.find((n) => (n.title || "").toLowerCase() === decodedTitles[0].toLowerCase()) || undefined;

    if (!currentNode) return;

    await ensureExpanded(currentNode);
    const pathNodes: TreeNode[] = [currentNode];

    for (let i = 1; i < decodedTitles.length; i++) {
      const seg = decodedTitles[i];
      await ensureExpanded(currentNode);

      let nextNode = findChildByTitleCI(currentNode, seg);

      // If we can't find the next node, check if we're at a site/subsite and look for "Documents"
      if (!nextNode && (currentNode.type === "site" || currentNode.type === "subsite")) {
        const docsNode = findChildByTitleCI(currentNode, "Documents");
        if (docsNode) {
          await ensureExpanded(docsNode);
          nextNode = findChildByTitleCI(docsNode, seg);
          if (nextNode) {
            currentNode = nextNode;
            pathNodes.push(currentNode);
            continue;
          } else {
            currentNode = docsNode;
            pathNodes.push(currentNode);
            i--; // Retry the same segment with the documents node as current
            continue;
          }
        }
      }

      if (!nextNode && currentNode.type === "library") {
        await ensureExpanded(currentNode);
        nextNode = findChildByTitleCI(currentNode, seg);
      }

      if (!nextNode) break;

      currentNode = nextNode;
      pathNodes.push(currentNode);
    }

    console.log("[navigateToPath] Final path:", pathNodes.map(n => n.title));

    // Update breadcrumbs based on the full path we navigated
    const breadcrumbPath = pathNodes.map(node => ({
      key: node.key,
      title: node.title,
      type: node.type,
      siteUrl: node.siteUrl,
      libraryTitle: node.libraryTitle,
      folderPath: node.folderPath
    }));

    setBreadcrumbs(breadcrumbPath);

    // If we ended on a library or folder, load its files
    if (currentNode.type === "library" || currentNode.type === "folder") {
      // Aman 16/4/26
      setSelectedCurrentNode(currentNode);
      // Aman 16/4/26 end
      await loadFilesForNode(currentNode);
      setBreadcrumbs(breadcrumbPath); //ritik 09/04/26 added after load file node set the breadcrumb
    }
    // Aman 16/4/26 - if we ended on site/subsite, we should still set selectedCurrentNode to update the UI context, but we won't load files until user clicks the library
    if (currentNode.type === "site" || currentNode.type === "subsite") {
  await handleNodeClick(currentNode);
}
  };
  /** --------------------------------------- clicking ----------------------------------------- */
  // const handleNodeClick = async (node: TreeNode) => {
  //   console.log("[handleNodeClick] node clicked:", { key: node.key, title: node.title, type: node.type, siteUrl: node.siteUrl });
  //   location = node.siteUrl;
  //   updateUrl(node);
  //   setActiveView("");

  //   // Case 1: Library or folder → load files
  //   if (node.type === "library" || node.type === "folder") {
  //     await loadFilesForNode(node);
  //     return;
  //   }

  //   // Case 2: Site or subsite → load its libraries first, then prefer 'Documents'
  //   if (node.type === "site" || node.type === "subsite") {
  //     console.log("[handleNodeClick] Site/Subsite: ensuring libraries are expanded");
  //     await ensureExpanded(node); // expands and fetches libraries

  //     // Try to find the 'Documents' library in the existing children (no transient nodes)
  //     const documentsNode = findChildByTitleCI(node, "Documents") || (node.children && node.children[0]);
  //     if (documentsNode) {
  //       console.log("[handleNodeClick] Will load files for library:", documentsNode.title);
  //       await loadFilesForNode(documentsNode);
  //     } else {
  //       console.log("[handleNodeClick] No libraries found under this site/subsite.");
  //     }
  //   }
  // };
  const handleNodeClick = async (node: TreeNode) => {
    setShowUploadPanel(false);    // Upload File Button Hide From Site collection and Subsites click by Aman
    setActiveComponent(false); // addhyan 16/4/26
    // this below lines will close preview modal  (addhyan work on this )
    setShowPreviewModal(false); // <-- Close preview modal
    setPreviewFile(null);       // <-- Clear preview file
    // hide template form when clicking folder/site nodes
    setShowTemplateForm(false);

    // close all modal when clicking on folder/site nodes (Ritik 17/03/2026)
    setRenameModalOpen(false);
    setShowAuditModal(false);
setShowVersionModal(false);
setShowShareModal(false);
// ✅ END close all modal when clicking on folder/site nodes

    // this below lines will close preview modal (addhyan work on this )
    console.log("[handleNodeClick] node clicked:", { key: node.key, title: node.title, type: node.type, siteUrl: node.siteUrl });
    location = node.siteUrl;
    updateUrl(node);
    setActiveView("");
    setSelectedCurrentNode(node);
    // Always update breadcrumbs first
    const breadcrumbPath = getNodePath(node.key);
    setBreadcrumbs(breadcrumbPath);

    if (node.type === "library" || node.type === "folder") {
      await loadFilesForNode(node);
      return;
    }

    if (node.type === "site" || node.type === "subsite") {
      await ensureExpanded(node);
      setSelectedFiles([]);    //Addhyan 14/01/2026 Clear files when clicking on site/subsite - show only tree structure
      // i have commented this below code it was previous working code but it was auto selecting document library 
      // console.log("[handleNodeClick] Site/Subsite: ensuring libraries are expanded");
      // await ensureExpanded(node);
      // const documentsNode = findChildByTitleCI(node, "Documents") || (node.children && node.children[0]);
      // if (documentsNode) {
      //   console.log("[handleNodeClick] Will load files for library:", documentsNode.title);
      //   await loadFilesForNode(documentsNode);
      // } else {
      //   console.log("[handleNodeClick] No libraries found under this site/subsite.");
      // }

    }
  };
  /** ------------------------------------ load files (same) ----------------------------------- */
  const loadFilesForNode = async (node: TreeNode) => {
    try {
      setFilesLoadedloadedfromnode(true);

      console.log("[loadFilesForNode] Start:", { key: node.key, title: node.title, type: node.type, siteUrl: node.siteUrl });
      const siteSP = spfi(node.siteUrl).using(SPFx(context));
      let folderPath = "";

      if (node.type === "library") {
        folderPath = `${node.libraryTitle}`;
        setCurrentFolderPath(node.libraryTitle || "");
      } else if (node.type === "folder") {
        folderPath = `${node.folderPath}`;
        setCurrentFolderPath(node.folderPath || "");
      }

      setCurrentSiteUrl(node.siteUrl);

      if (folderPath) {
        const serverRel = `/sites/${node.siteUrl.split("/sites/")[1]}/${folderPath}`;
        console.log("[loadFilesForNode] Fetching files from:", serverRel);
        const files = await siteSP.web.getFolderByServerRelativePath(serverRel).files();
        console.log("[loadFilesForNode] Files fetched:", files);
        console.log("[loadFilesForNode] Files loaded:", files?.length || 0);
        // setSelectedFiles(files);

        const baseSiteUrl = node.siteUrl.split("/sites/")[0] + "/sites/";
        const siteCollection = node.siteUrl.split("/sites/")[1].split("/")[0];
        const entityName = node.siteUrl.split("/").pop();

        const spRoot = spfi(`${baseSiteUrl}${siteCollection}`).using(SPFx(context));

      //  const fmItems  = await spRoot.web.lists
      //     .getByTitle(`DMS${entityName}FileMaster`)
      //     .items.select("FileName", "IsDeleted")
      //     .top(5000)();

      //   const deletedSet = new Set(
      //     fmItems
      //       .filter(i => i.IsDeleted !== null && i.IsDeleted !== undefined)
      //       .map(i => i.FileName)
      //   );

      //   const visibleFiles = (files || []).filter(
      //     (f: any) => !deletedSet.has(f.Name)
      //   );

      //   setSelectedFiles(visibleFiles);

//       const fmItems = await spRoot.web.lists
//   .getByTitle(`DMS${entityName}FileMaster`)
//   .items.select("FileName", "IsDeleted", "Status")
//   .top(5000)();
 
// // 2. Build the "Allowed" Set
// const allowedSet = new Set(
//   fmItems
//     .filter(i => {
//       // Must NOT be deleted
//       const isNotDeleted = !i.IsDeleted;
//       // Must be Approved, Auto Approved, or Null/Empty
//       const hasValidStatus = i.Status === "Approved" ||
//                              i.Status === "Auto Approved" ||
//                              i.Status === null ||
//                              i.Status === "";
 
//       return isNotDeleted && hasValidStatus;
//     })
//     .map(i => i.FileName)
// );
 
// // 3. Filter your folder files based on the allowedSet
// const visibleFiles = (files || []).filter(
//   (f: any) => allowedSet.has(f.Name)
// );
 
// setSelectedFiles(visibleFiles);
 
 // 1. Fetch the Master List data
const currentFolder = node.title;
const fmItems = await spRoot.web.lists
  .getByTitle(`DMS${entityName}FileMaster`)
  .items.select("ID", "FileName", "IsDeleted", "Status") // Included ID and Status
  .filter(`DocumentLibraryName eq '${currentFolder}'`)
  .top(5000)();

// 2. Create a Map for quick lookup (Key: FileName, Value: {ID, Status})
const masterDataMap = new Map();

fmItems.forEach(i => {
  // Logic: Only add to the map if it meets your requirements
  const hasValidStatus = i.Status === "Approved" || 
                         i.Status === "Auto Approved" || 
                         i.Status === null || 
                         i.Status === "";
                         
  const isNotDeleted = !i.IsDeleted; // Adjust if IsDeleted is "Yes"/1

  if (hasValidStatus && isNotDeleted) {
    masterDataMap.set(i.FileName, {
      id: i.ID,
      status: i.Status
    });
  }
});

// 3. Filter and Enrich the files array
const visibleFiles = (files || [])
  .filter((f: any) => masterDataMap.has(f.Name)) // Only keep if in the "allowed" map
  .map((f: any) => {
    const extraData = masterDataMap.get(f.Name);
    return {

          // ritik chnage - 29/04/26 start

    //   ...f,             // Keep all original file properties (ServerRelativeUrl, etc.)
    //   ID: extraData.id, // Inject the List ID
    //   Status: extraData.status, // Inject the Status
    //   // srs 10/4/26
    //   // --- ADD THESE LINES TO FIX PERMISSIONS ---
    //   FileUID: f.UniqueId,             // Maps library GUID to the expected property
    //   SiteID: node.siteUrl,            // Passes the current subsite URL
    //   DocumentLibraryName: node.libraryTitle,
    //   SiteName: entityName             // Used for Admin Group naming logic
    // };



    ...f,
      ID: extraData.id,
      Status: extraData.status,
      FileUID: f.UniqueId,
      SiteID: node.siteUrl,
      __siteUrl: node.siteUrl.split("/sites/")[0] + "/sites/" + node.siteUrl.split("/sites/")[1]?.split("/")[0],
      CurrentFolderPath: serverRel.replace(`/${f.Name}`, ""),
      DocumentLibraryName: node.libraryTitle,
      SiteName: entityName,
      FileName: f.Name,
      // FilePreviewURL: `${window.location.origin}${f.ServerRelativeUrl}`,
      FilePreviewURL: (() => {
        const serverRel = f.ServerRelativeUrl || "";
        const parentFolder = serverRel.substring(0, serverRel.lastIndexOf("/"));
        return `${window.location.origin}${parentFolder}/Forms/AllItems.aspx?id=${encodeURIComponent(serverRel)}&parent=${encodeURIComponent(parentFolder)}`;
      })(),
      // ritik chnage - 29/04/26 end 
    };
  });

  // ritik 10/04/26 start

  // setSelectedFiles(visibleFiles); Ritik added 10/04/26 when someone upload the file it will apairing on first page
const sortedFiles = [...visibleFiles].sort((a: any, b: any) => {
  const dateA = new Date(a.TimeCreated || a.TimeLastModified || 0).getTime();
  const dateB = new Date(b.TimeCreated || b.TimeLastModified || 0).getTime();
  return dateB - dateA; // latest first
});
setSelectedFiles(sortedFiles);
console.log("[loadFilesForNode] Visible Files with ID and Status:", visibleFiles);

// setBreadcrumbs(getNodePath(node.key)); ritik 09/04/26
        const nodePath = getNodePath(node.key);
if (nodePath && nodePath.length > 0) {
  setBreadcrumbs(nodePath);
}

// ritik 10/04/26 end 

setSelectedFiles(visibleFiles);
console.log("[loadFilesForNode] Visible Files with ID and Status:", visibleFiles);

        console.log("[loadFilesForNode] Selected files set:", selectedFiles);
        setBreadcrumbs(getNodePath(node.key));

        console.log("Breadcrumbs updated:", breadcrumbs);
        console.log("Breadcrumbs updated 2:" + JSON.stringify(getNodePath(node.key)))
      } else {
        console.log("[loadFilesForNode] No folderPath computed for node; skipping files fetch.");
      }
    } catch (error) {
      console.error("[loadFilesForNode] Error:", error);
      setSelectedFiles([]);
    }
  };

  const toggleNode = async (node: TreeNode) => {
    if (node.type === "function") {
      console.log("[toggleNode] Expanding function:", node.title);
      node.isExpanded = !node.isExpanded;
      setTreeData([...treeData]);
      buildNodeMap(treeData);
      return;
    }

    if (!node.isExpanded && node.hasChildren) {
      const siteSP = spfi(node.siteUrl).using(SPFx(context));
      console.log("[toggleNode] Expanding:", { title: node.title, type: node.type, site: node.siteUrl });
      try {
        console.log("[toggleNode] Fetching node type:", node.type);
        console.log("[toggleNode] Fetching node type:", node.type, "for node:", node.title);

        if (node.type === "subsite") {
          // i am adding this here set currrent site url because when we click on subsite it will set current site url for upload file / or create fodler it pass as props
          setCurrentSiteUrl(node.siteUrl);
          console.log("[toggleNode] Loading libraries for subsite:", node.title);
          // const libs = await siteSP.web.lists.filter("BaseTemplate eq 101 and Hidden eq false").select("Title")();
           // srs 31/3/26 filter out Documents and Site Assets 
          const libs = await siteSP.web.lists
    .filter("BaseTemplate eq 101 and Hidden eq false and Title ne 'Documents' and Title ne 'Site Assets'")
    .select("Title")();
          node.children = libs.map((lib: any) => ({
            key: `lib-${node.key}-${lib.Title}`,
            title: lib.Title,
            type: "library",
            siteUrl: node.siteUrl,
            libraryTitle: lib.Title,
            hasChildren: true,
            parentKey: node.key,
          }));
          console.log("[toggleNode] Libraries loaded:", node.children?.map((c) => c.title));
        }

        if (node.type === "library") {

          const serverRel = `/sites/${node.siteUrl.split("/sites/")[1]}/${node.libraryTitle}`;
          console.log("[toggleNode] Loading folders for library:", node.title, "path:", serverRel);
          const folders = await siteSP.web.getFolderByServerRelativePath(serverRel).folders();
          node.children = folders
            .filter((f: any) => !f.Name.startsWith("Forms"))
            .map((f: any) => ({
              key: `folder-${node.key}-${f.Name}`,
              title: f.Name,
              type: "folder",
              siteUrl: node.siteUrl,
              libraryTitle: node.libraryTitle,
              folderPath: `${node.libraryTitle}/${f.Name}`,
              hasChildren: true,
              parentKey: node.key,
            }));
          console.log("[toggleNode] Folders loaded:", node.children?.map((c) => c.title));
        }

        if (node.type === "folder") {
          const serverRel = `/sites/${node.siteUrl.split("/sites/")[1]}/${node.folderPath}`;
          console.log("[toggleNode] Loading subfolders for folder:", node.title, "path:", serverRel);
          const subFolders = await siteSP.web.getFolderByServerRelativePath(serverRel).folders();
          node.children = subFolders
            .filter((f: any) => !f.Name.startsWith("Forms"))
            .map((f: any) => ({
              key: `subfolder-${node.key}-${f.Name}`,
              title: f.Name,
              type: "folder",
              siteUrl: node.siteUrl,
              libraryTitle: node.libraryTitle,
              folderPath: `${node.folderPath}/${f.Name}`,
              hasChildren: true,
              parentKey: node.key,
            }));
          console.log("[toggleNode] Subfolders loaded:", node.children?.map((c) => c.title));
        }

        node.isExpanded = true;
        setTreeData([...treeData]);
        buildNodeMap(treeData);
      } catch (error) {
        console.error("[toggleNode] Error expanding node:", error);
      }
    } else {
      node.isExpanded = !node.isExpanded;
      console.log("[toggleNode] Toggled to", node.isExpanded ? "expanded" : "collapsed", "for:", node.title);
      setTreeData([...treeData]);
    }
  };
  /** -------------------------------- breadcrumb click (same) --------------------------------- */
  const handleBreadcrumbClick = async (item: BreadcrumbItem) => {
    setShowUploadPanel(false);                  // Upload File Button Hide From Site collection and Subsites click by Aman
   setActiveComponent(false); // addhyan 16/4/26
    console.log("[breadcrumb] Clicked:", item);
    if (item.type === "view") {
      handleViewButtonClick(item.title);
    } else {
      const node = nodeMap[item.key];
      if (node) {
        await expandPathToNode(node.key);

        if (node.type === "subsite") {
          // aman 20/04/26 start 
           setSelectedCurrentNode(node);
 
          // aman 20/04/26 end  

          // Load libraries instead of files
          await toggleNode(node); // Expands and loads children (libraries)
          setSelectedFiles([]); // Clear file list so only sidebar shows folders
          setBreadcrumbs(getNodePath(node.key));
        } else {
          await handleNodeClick(node);
        }
      } else {
        console.warn("[breadcrumb] Node not found in nodeMap for key:", item.key);
      }
    }
  };

  // Aman 13/4/26
const handleBreadcrumbShareClick = () => {
  const url = window.location.href;
 
  setBreadcrumbShare({
    show: true,
    url: url
  });
};
 
const closeBreadcrumbShare = () => {
  setBreadcrumbShare(prev => ({
    ...prev,
    show: false
  }));
};

const handleFolderShareClick = (file: any) => {
  try {
    if (!file || !file.__siteUrl) return;
 
    const baseUrl = "https://officeindia.sharepoint.com/sites/ESSA/SitePages/ESSADMS.aspx";
 
    let folderPath = file.FolderPath || "";
 
    if (typeof folderPath !== "string") return;
 
    if (folderPath.includes("/sites/")) {
      folderPath = folderPath.split("/sites/")[1];
    }
 
    if (folderPath.startsWith("/")) {
      folderPath = folderPath.substring(1);
    }
 
    const finalUrl = `${baseUrl}#${encodeURI(folderPath)}`;
 
    // ✅ POPUP OPEN
    setFolderShare({
      show: true,
      url: finalUrl
    });
 
  } catch (error) {
    console.error("Folder Share Error:", error);
  }
};
// End Aman 13/4/26

  const expandPathToNode = async (nodeKey: string) => {
    console.log("[expandPathToNode] for key:", nodeKey);
    let currentNode = nodeMap[nodeKey];
    const nodesToExpand: TreeNode[] = [];

    while (currentNode) {
      nodesToExpand.unshift(currentNode);
      if (currentNode.parentKey) {
        currentNode = nodeMap[currentNode.parentKey];
      } else {
        break;
      }
    }

    for (const node of nodesToExpand) {
      if (!node.isExpanded) {
        await toggleNode(node);
      }
    }
  };

  /** ------------------------------------ quick views (same) ---------------------------------- */
  const handleViewButtonClick = async (viewName: string) => {
    setFilesLoadedloadedfromnode(false);
    setActiveComponent(false);   // srs 29/1/26
    setShowUploadPanel(false);   // srs 29/1/26
    setMenuOpenIdx(null);
    setSelectedCurrentNode(null);
    // hide template form when switching views
    setShowTemplateForm(false);

    // close all modal when switching views (Ritik 17/03/2026)
    setShowAuditModal(false);
setShowVersionModal(false);
setShowShareModal(false);
setRenameModalOpen(false);
// ✅ END close all modal when switching views
    console.log("[view] button clicked:", viewName);
    setActiveView(viewName);
    setBreadcrumbs([
      {
        key: viewName.toLowerCase().replace(/\s+/g, "-"),
        title: viewName,
        type: "view",
        siteUrl: "",
      },
    ]);

    // Clear any URL hash when switching to a view
    window.history.pushState(null, "", window.location.pathname);

//     try {
//     let files: any[] = [];
//     let viewType: ViewType;
//     switch (viewName) {
//       case "My request":
//         viewType = "MyRequest";
//         break;
//       case "My favourite":
//         viewType = "MyFavourite";
//         break;
//       case "My Folders":
//         viewType = "MyFolders";
//         break;
//       case "Share with me":
//         viewType = "SharedWithMe";
//         break;
//       case "Share with other":
//         viewType = "SharedWithOthers";
//         break;
//       case "Recycle bin":
//         viewType = "RecycleBin";
//         break;
//       default:
//         viewType = "MyRequest";
//     }
//     // ritik 19/02/2026 (remove duplicate files from view data i have added FileUID, UniqueId, Id in view data reason is when we share file with other user it will show duplicate files in view data)
//     files = await loadViewData(viewType);
//     const seen = new Set();
//     const uniqueFiles = files.filter((item: any) => {
//       const uid = item.FileUID || item.UniqueId || item.Id;
//       if (seen.has(uid)) return false;
//       seen.add(uid);
//       return true;
//       });
// updateViewCount(viewType, uniqueFiles.length);
// setSelectedFiles(uniqueFiles);
//   // end here
//     files = await loadViewData(viewType);
//     updateViewCount(viewType, files.length);
//     console.log(`[view] Loaded ${viewName} data:`, files);
//     setSelectedFiles(files);
//     console.log("[loadFilesForNode] Selected files set:", selectedFiles);
//     setCurrentFolderPath("");
//     setCurrentSiteUrl("");
//   } // ritik 24/2/26 for handle view button click
try {
  let viewType: ViewType;
  switch (viewName) {
    case "My request":       viewType = "MyRequest"; break;
    case "My favourite":     viewType = "MyFavourite"; break;
    case "My Folders":       viewType = "MyFolders"; break;
    case "Share with me":    viewType = "SharedWithMe"; break;
    case "Share with other": viewType = "SharedWithOthers"; break;
    case "Recycle bin":      viewType = "RecycleBin"; break;
    // srs 6/3/26
    case "Archived Files":   viewType = "ArchivedFiles"; break;
    default:                 viewType = "MyRequest";
  }
 
  const files = await loadViewData(viewType);
 
  const seen = new Set();
  const uniqueFiles = files.filter((item: any) => {
    const uid = item.FileUID || item.UniqueId || item.Id;
    if (seen.has(uid)) return false;
    seen.add(uid);
    return true;
  });
  // srs 12/3/26 correct code if count mismatch in archival issue in current code data duplicate
//  const uniqueFiles = getUniqueFiles(files);
  console.log(`[view] Loaded ${viewName} data:`, uniqueFiles);
  updateViewCount(viewType, uniqueFiles.length);
  setSelectedFiles(uniqueFiles);
  setCurrentFolderPath("");
  setCurrentSiteUrl("");
    refreshViewCount(viewType); //Ritik 10/04/26
 
} catch (error) {
      console.error(`[view] Error loading ${viewName} data:`, error);
      setSelectedFiles([]);
    }
  };

  //Ritik 28/01/2026
  // 🔹 Safe field reader (case / internal name safe)
  const readField = (obj: any, keys: string[]): string => {
    if (!obj) return "";
    const map: any = {};
    Object.keys(obj).forEach(k => (map[k.toLowerCase()] = k));

    for (const k of keys) {
      const real = map[k.toLowerCase()];
      if (real) return obj[real] || "";
    }
    return "";
  };
const loadViewData = async (viewType: ViewType): Promise<any[]> => {
    // setShowPreviewModal(false)
    // Addhyan 13/4/26
    if (!isDeepLinkPreview.current) setShowPreviewModal(false);
    // sourish 20/8/25 previous wworking code 
    // if (viewType === "MyFolders") {
    //   try {
    //     const meEmail =
    //       (context.pageContext as any)?.user?.email ||
    //       (context.pageContext as any)?.user?.loginName ||
    //       "";

    //     const masterSites = await sp.web.lists
    //       .getByTitle("MasterSiteCollection")
    //       .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite")
    //       .top(5000)();

    //     const allMyFoldersData: any[] = [];

    //     for (const ms of masterSites) {
    //       try {
    //         // Always target "DMSFolderMaster" list from SiteURL
    //         const siteSP = spfi(ms.SiteURL).using(SPFx(context));
    //         const listItems = await siteSP.web.lists
    //           .getByTitle("DMSFolderMaster")
    //           .items.select(
    //             "Id",
    //             "Title",
    //             "FolderPath",
    //             "SiteTitle",
    //             "DocumentLibraryName",
    //             "FolderName",
    //             "CurrentUser",
    //             "Modified",
    //             "Author/Id",
    //             "Author/Title",
    //             "Author/EMail"
    //           )
    //           .expand("Author") // 🔹 expand ModifiedBy lookup
    //           .top(5000)();

    //         const normalizedItems = listItems.map((it: any) => ({
    //           ...it,
    //           ParentFolder: it.ParentFolder || it.ParentFolderId || "", // normalize
    //         }));

    //         // 🔹 Filter by CurrentUser column
    //         const filtered = (listItems || []).filter(
    //           (it: any) =>
    //             (it.Author.EMail || "").toLowerCase() === meEmail.toLowerCase()
    //         );

    //         allMyFoldersData.push(
    //           ...filtered.map((it: any) => ({
    //             ...it,
    //             __source: "MyFolders",
    //             __siteUrl: ms.SiteURL,
    //             __listName: "DMSFolderMaster",
    //             __CreatedBy: it.Author?.Title || it.Author?.EMail || "",
    //           }))
    //         );
    //       } catch (err) {
    //         console.error(
    //           `Error fetching list DMSFolderMaster from ${ms.SiteURL}:`,
    //           err
    //         );
    //       }
    //     }

    //     return allMyFoldersData;
    //   } catch (e) {
    //     console.error("Error fetching MasterSiteCollection:", e);
    //     return [];
    //   }
    // }
    //updated code by aman after manage folder permission
    if (viewType === "MyFolders") {
      try {
        const meEmail =
          (context.pageContext as any)?.user?.email ||
          (context.pageContext as any)?.user?.loginName ||
          "";


        // const masterSites = await sp.web.lists
        //   .getByTitle("MasterSiteCollection")
        //   .items.select("Id", "Title", "SiteURL")
        //   .top(5000)();

           // srs 18/3/26
              const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL", "IsActive").filter("IsActive eq 1")
          .top(5000)();

        // i commented this because it was fetching only records from spfx2site Dmsfodlermaster list
        // const rootSiteUrl = "https://officeindia.sharepoint.com/sites/AlRostmaniSpfx2";

        // const scopedSites = (masterSites || []).filter((ms: any) =>
        //   (ms.SiteURL || "").toLowerCase().startsWith(rootSiteUrl.toLowerCase())
        // );

        const allMyFoldersData: any[] = [];

        // for (const ms of scopedSites) {
        for (const ms of masterSites) {
          try {
            const siteSP = spfi(ms.SiteURL).using(SPFx(context));
            const listItems = await siteSP.web.lists
              .getByTitle("DMSFolderMaster")
              .items.select("*", "Author/Id", "Author/Title", "Author/EMail")
              .expand("Author")
              .top(5000)();

            let __webId = "";
            try {
              const webMeta: any = await siteSP.web.select("Id")();
              __webId = webMeta?.Id || "";
            } catch { }

            const mine = (listItems || []).filter(
              (it: any) => ((it?.Author?.EMail || "") as string).toLowerCase() === meEmail.toLowerCase()
            );

            const readBool = (obj: any, keys: string[]): boolean => {
              if (!obj) return false;
              const lowerMap: Record<string, string> = {};
              Object.keys(obj).forEach((k) => (lowerMap[k.toLowerCase()] = k));
              for (const k of keys) {
                const real = lowerMap[k.toLowerCase()];
                if (real !== undefined) return !!obj[real];
              }
              return false;
            };

            const normalized = mine.map((it: any) => {
              const isLibrary = readBool(it, ["IsLibrary", "Library", "islibrary", "islib"]);
              const isFolder = readBool(it, ["IsFolder", "Folder", "isfolder", "issubfolder"]);
              const isPrivate = readBool(it, ["IsPrivate", "Private", "isprivate"]);
              const isPublic = readBool(it, ["IsPublic", "Public", "ispublic"]);
              return {
                ...it,
                 DisplayName:
    it.FolderName ||
    it.Title ||
    it.DocumentLibraryName || "",
                __source: "MyFolders",
                __siteUrl: ms.SiteURL,
                __listName: "DMSFolderMaster",
                __CreatedBy: it?.Author?.Title || it?.Author?.EMail || "",
                __webId,
                __flags: { isLibrary, isFolder, isPrivate, isPublic },
                __scopeLabel: isLibrary ? "Root folder" : isFolder ? "Sub folder" : "",
                __visibilityLabel: isPrivate ? "Private" : isPublic ? "Public" : "",
              };
            });

            // Remove system Forms & dedupe, keep roots and folders
            // const cleaned = normalized.filter((row: any) => {
            //   const isLibrary = row.__flags?.isLibrary === true;
            //   const name = row.FolderName || row.Title || "";
            //   const path = row.FolderPath || row.folderpath || row.ServerRelativeUrl || "";
            //   const lowerPath = (path || "").toLowerCase().replace(/\/+$/, "");
            //   const isForms = /\/forms(\/|$)/i.test(lowerPath);
            //   return !!name && !!path && !isForms;
            // });

            const cleaned = normalized.filter((row: any) => {

  const isLibrary = row.__flags?.isLibrary === true;
 

  const name =
    row.FolderName || row.Title ||
    row.Title ||
    row.DocumentLibraryName || "";

  const path =
    row.FolderPath ||
    row.folderpath ||
    row.ServerRelativeUrl || "";

  const lowerPath = (path || "").toLowerCase().replace(/\/+$/, "");
  const isForms = /\/forms(\/|$)/i.test(lowerPath);

  // ⭐ Library ko force allow karo
  if (isLibrary) return !!path && !isForms;

  return !!name && !!path && !isForms;
});


            const unique: any[] = [];
            const seen = new Set<string>();
            for (const r of cleaned) {
              const key = (r.FolderPath || r.folderpath || r.ServerRelativeUrl || String(r.ID || r.Id || "")).toLowerCase();
              if (!key || seen.has(key)) continue;
              seen.add(key);
              unique.push(r);
            }

            allMyFoldersData.push(...unique);
          } catch (err) {
            console.error(`Error fetching list DMSFolderMaster from ${ms.SiteURL}:`, err);
          }
        }

        return allMyFoldersData;
      } catch (e) {
        console.error("Error fetching MasterSiteCollection:", e);
        return [];
      }
    }

    // sourish 19/8/25
    if (viewType === "SharedWithOthers") {
      try {
        const meEmail =
          (context.pageContext as any)?.user?.email ||
          (context.pageContext as any)?.user?.loginName ||
          "";

        // const masterSites = await sp.web.lists
        //   .getByTitle("MasterSiteCollection")
        //   .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite")
        //   .top(5000)();
             // srs 18/3/26
              const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite", "IsActive").filter("IsActive eq 1")
          .top(5000)();

        const allSharedOtherData: any[] = [];

        for (const ms of masterSites) {
          // 🔹 Extract list name from SharewithOtherMeMasterSite URL
          let listName = "";
          if (ms.SharewithOtherMeMasterSite) {
            const parts = ms.SharewithOtherMeMasterSite.split("/Lists/");
            if (parts.length > 1) {
              listName = parts[1].split("/")[0]; // e.g. DMSShareWithOtherMaster
            }
          }

          if (listName) {
            try {
              // 🔹 Connect to the site & fetch items
              const siteSP = spfi(ms.SiteURL).using(SPFx(context));
              const listItems = await siteSP.web.lists
                .getByTitle(listName)
                .items.top(5000)();
              // 🔹 Filter by CurrentUser
              //Ritik 28/01/2026
              const filtered = (listItems || []).filter((it: any) => {
                const owner = readField(it, ["CurrentUser", "Owner", "CreatedBy"]);
                const sharedTo = readField(it, ["ShareWithMe", "SharedTo", "SharedWith"]);
                const isDeleted = it.IsDeleted !== null && it.IsDeleted !== undefined;

                return (
                  owner.toLowerCase() === meEmail.toLowerCase() &&
                  sharedTo &&
                  !sharedTo.toLowerCase().includes(meEmail.toLowerCase()) &&
                  !isDeleted &&
                  it.IsFavourite !== true
                );
              });

              //debugging console table Ritik 28/01/2026
              console.table(
                filtered.map(f => ({
                  File: f.FileName || f.Title,
                  Owner: readField(f, ["CurrentUser", "Owner", "CreatedBy"]),
                  SharedTo: readField(f, ["ShareWithMe", "SharedTo", "SharedWith"])
                }))
              );

              allSharedOtherData.push(
                ...filtered.map((it: any) => ({
                  ...it,
                  __source: "SharedWithOthers",
                  __siteUrl: ms.SiteURL,
                  __listName: listName,
                }))
              );
            } catch (err) {
              console.error(
                `Error fetching list ${listName} from ${ms.SiteURL}:`,
                err
              );
            }
          }
        }

        return allSharedOtherData;
      } catch (e) {
        console.error("Error fetching MasterSiteCollection:", e);
        return [];
      }
    }

    // sourish 19/8/25
    if (viewType === "SharedWithMe") {
      try {
        const meEmail =
          (context.pageContext as any)?.user?.email ||
          (context.pageContext as any)?.user?.loginName ||
          "";

        // const masterSites = await sp.web.lists
        //   .getByTitle("MasterSiteCollection")
        //   .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite")
        //   .top(5000)();

             // srs 18/3/26
              const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite", "IsActive").filter("IsActive eq 1")
          .top(5000)();

        const allSharedData: any[] = [];

        for (const ms of masterSites) {
          // 🔹 Extract list name from SharewithOtherMeMasterSite URL
          let listName = "";
          if (ms.SharewithOtherMeMasterSite) {
            const parts = ms.SharewithOtherMeMasterSite.split("/Lists/");
            if (parts.length > 1) {
              listName = parts[1].split("/")[0]; // e.g. DMSShareWithOtherMaster
            }
          }

          if (listName) {
            try {
              // 🔹 Connect to the site & fetch items
              const siteSP = spfi(ms.SiteURL).using(SPFx(context));
              const listItems = await siteSP.web.lists
                .getByTitle(listName)
                .items.top(5000)();

              // 🔹 Filter by CurrentUser
              // const filtered = (listItems || []).filter(
              //   (it: any) =>
              //     (it.CurrentUser || "").toLowerCase() === meEmail.toLowerCase()
              // );



              // 🔹 Shared WITH ME = kisi aur ne mujhe share kiya - Ritik Today 29/01/2026
              const filtered = (listItems || []).filter((it: any) => {
                const owner = readField(it, ["CurrentUser", "Owner", "CreatedBy"]);
                const sharedTo = readField(it, ["ShareWithMe", "SharedTo", "SharedWith"]);

                return (
                  sharedTo.toLowerCase().includes(meEmail.toLowerCase()) && // 👈 mujhe mila
                  owner.toLowerCase() !== meEmail.toLowerCase()             // 👈 maine create/share nahi ki
                );
              });



              //Debugging console table Ritik 28/01/2026
              console.table(
                filtered.map(f => ({
                  File: f.FileName || f.Title,
                  Owner: readField(f, ["CurrentUser", "Owner", "CreatedBy"]),
                  SharedTo: readField(f, ["ShareWithMe", "SharedTo", "SharedWith"])
                }))
              );





              allSharedData.push(
                ...filtered.map((it: any) => ({
                  ...it,
                  __source: "SharedWithMe",
                  __siteUrl: ms.SiteURL,
                  __listName: listName,
                }))
              );
            } catch (err) {
              console.error(
                `Error fetching list ${listName} from ${ms.SiteURL}:`,
                err
              );
            }
          }
        }

        return allSharedData;
      } catch (e) {
        console.error("Error fetching MasterSiteCollection:", e);
        return [];
      }
    }
    if (viewType === "MyRequest" || viewType === "MyFavourite" || viewType === "RecycleBin") {

      // 🔹 Step 1: Get all site + list mappings from master config
      // const configItems = await sp.web.lists
      //   .getByTitle("Allsitesfilemaster")
      //   .items.select("FileMaster", "sitecollection")
      //   .top(5000)();

       const configItems = await sp.web.lists
        .getByTitle("Allsitesfilemaster")
        .items.select("FileMaster", "sitecollection").filter("Isactive eq 1")
        .top(5000)();
        console.log("configItems" ,configItems)

      const groupedBySite = configItems.reduce((acc, config) => {
        (acc[config.sitecollection] ||= []).push(config);
        return acc;
      }, {} as Record<string, any[]>);

      const allData: any[] = [];
      const currentUrl = new URL(context.pageContext.web.absoluteUrl);

      // Current user info (for filtering RecycleBin items)
      const meName = context.pageContext.user.displayName || "";
      const meEmail =
        (context.pageContext as any)?.user?.email ||
        (context.pageContext as any)?.user?.loginName ||
        "";

      // 🔹 Step 2: Process each site collection
      for (const [siteCollection, siteConfigs] of Object.entries(groupedBySite)) {
        try {
          const siteUrl = `${currentUrl.protocol}//${currentUrl.hostname}/sites/${siteCollection}`;
          const siteSP = spfi(siteUrl).using(SPFx(context));

          // Process each FileMaster list inside this site collection
          const siteData = await Promise.all(
            (siteConfigs as any[]).map(async (config: any) => {
              try {
                console.log(`Fetching from ${config.FileMaster} in ${siteCollection}`);
                const items = await siteSP.web.lists
                  .getByTitle(config.FileMaster)
                  .items.top(1000)();

                // if (viewType === "MyFavourite") {
                //   // 🔹 Filter favourites // ritik 29/01/2026
                //   return items.filter((item: any) =>
                //     item.IsFavourite === true &&
                //     item.IsDeleted === null &&          // 🔴
                //     item.MyRequest !== true &&          // 🔴
                //     !item.ShareWithMe                   // 🔴
                //   );

                // }
                if (viewType === "MyFavourite") {
  return (items || [])
    .filter((item: any) => {
 
      const owner =
        (item.CurrentUser ||
          item.Owner ||
          item.CreatedBy ||
          "").toLowerCase();
 
      return (
        item.IsFavourite === true &&
        item.IsDeleted === null &&
        !item.ShareWithMe &&
        (
          owner === meEmail.toLowerCase() ||
          owner === meName.toLowerCase()
        )
      );
    })
    .map((it: any) => ({
      ...it,
      __siteUrl: siteUrl,
      __fileMasterList: config.FileMaster,
    }));
}

                if (viewType === "RecycleBin") {
                  // 🔹 Only items deleted by current user
                  return (items || [])
                    .filter((it: any) => {
                      const cu = it?.CurrentUser || "";
                      const isDel = it?.IsDeleted;
                      return isDel !== null && isDel !== undefined && (cu === meName || cu === meEmail);
                    })
                    .map((it: any) => ({
                      ...it,
                      __siteUrl: siteUrl,
                      __fileMasterList: config.FileMaster,
                      __libraryTitle: (it as any).DocumentLibraryName || (it as any).LibraryName || (it as any).DocumentLibrary || (it as any).DocLibName || "",
                      __folderPath: (it as any).FolderPath || (it as any).FilePath || (it as any).RelativeFolderPath || "",
                    }));
                }

                // 🔹 Default: MyRequest (no filter, just tag for traceability)
                //Ritik 29/01/2026 added isdeleted filter

                // 🔹 MyRequest ONLY
                // if (viewType === "MyRequest") {
                //   return (items || [])
                //     .filter((it: any) =>
                //       it.MyRequest === true &&                 // 👈 sirf request wali
                //       it.IsDeleted === null &&                 // 👈 recycle bin nahi
                //       it.IsFavourite !== true &&               // 👈 favourite nahi
                //       !it.ShareWithMe                          // 👈 shared nahi
                //     )
                //     .map((it: any) => ({
                //       ...it,
                //       __siteUrl: siteUrl,
                //       __fileMasterList: config.FileMaster,
                //     }));
                // }

                // Aman 24/2/26
                 if (viewType === "MyRequest") {
                  return (items || [])
    .filter((item: any) => {
 
      const owner =
        (item.CurrentUser ||
          item.Owner ||
          item.CreatedBy ||
          "").toLowerCase();
 
      return (
        item.MyRequest === true &&
        item.IsDeleted === null &&
        item.IsFavourite !== true &&
        !item.ShareWithMe &&
        (
          owner === meEmail.toLowerCase() ||
          owner === meName.toLowerCase()
        )
      );
    })
    .map((it: any) => ({
      ...it,
      __siteUrl: siteUrl,
      __fileMasterList: config.FileMaster,
    }));
                }

                // 🔹 fallback (kuch bhi nahi)
                return [];







              } catch (error) {
                console.error(`Error fetching from ${config.FileMaster} in ${siteCollection}:`, error);
                return [];
              }
            })
          );

          allData.push(...siteData.flat());
        } catch (error) {
          console.error(`Error processing site collection ${siteCollection}:`, error);
        }
      }

  return allData;
    }

    // srs 6/3/26 
   if (viewType === "ArchivedFiles") {
  try {
    // 1. Get current user's email for filtering
     const meName = context.pageContext.user.displayName || "";
     const meEmail = context.pageContext.user.email;

    if (!meEmail) {
      console.warn("Could not determine current user email.");
    }

    // 2. Fetch items from the central EssaArchivallist list
    // Note: We use the list Title "EssaArchivallist" 
    const archiveItems = await sp.web.lists
      .getByTitle("EssaArchivallist")
      .items
      .select(
        "Id", "Title", "filename", "fileuid", "filepreviewurl", 
        "sitenamae", "sitecollectionname", "documentlibraryName", 
        "folderNames", "folderPath", "Isarchive", "Etag_azure", 
        "Blob_path", "Modified", "Author/Title", "Author/EMail"
      )
      .expand("Author") // To access "Created By" email
      .top(5000)();

    // 3. Filter the data
    // We filter where the current user is the creator and Isarchive is true
    const filteredData = (archiveItems || []).filter((item: any) => {
      const creatorEmail = item.Author.EMail || "";
      const isArchived = item.Isarchive === true;

      return (
        creatorEmail.toLowerCase() === meEmail.toLowerCase() &&
        isArchived
      );
    });

    // 4. Map to your standard object format
    return filteredData.map((item: any) => ({
      ...item,
      FileName: item.filename || item.Title,
      FileUID: item.fileuid,
      SiteName: item.sitenamae,
      LibraryName: item.documentlibraryName,
      __source: "ArchivedFiles"
    }));

  } catch (error) {
    console.error("Error fetching items from EssaArchivallist:", error);
    return [];
  }
}
    return [];
  };
 
  // Helper function to update view counts
  const updateViewCount = (viewType: ViewType, count: number) => {
    const viewKey = viewType === "MyRequest" ? "My request" 
                  : viewType === "MyFavourite" ? "My favourite"
                  : viewType === "MyFolders" ? "My Folders"
                  : viewType === "SharedWithMe" ? "Share with me"
                  : viewType === "SharedWithOthers" ? "Share with other"
                  // srs 6/3/26
                  : viewType === "ArchivedFiles" ? "Archived Files"
                  : "Recycle bin";
 
    setViewCounts(prev => ({...prev, [viewKey]: count}));
  };
  // #region Toggle IsFav List
  // const toggleFavourite = async (file: any) => {
  //   try {
  //     // Build site URL dynamically
  //     const currentUrl = new URL(context.pageContext.web.absoluteUrl);
  //     const siteUrl = `${currentUrl.protocol}//${currentUrl.hostname}/sites/${file.SiteCollection}`;
  //     const siteSP = spfi(siteUrl).using(SPFx(context));

  //     // Update SharePoint item
  //     await siteSP.web.lists
  //       .getByTitle(file.ListName)
  //       .items.getById(file.Id)
  //       .update({
  //         IsFavourite: !file.IsFavourite, // flip value
  //       });

  //     console.log(
  //       `Updated favourite for ${
  //         file.Title || file.FileName
  //       } → ${!file.IsFavourite}`
  //     );

  //     // Refresh current view (example: MyFavourite)
  //     await loadViewData("MyFavourite");
  //     alert("Done!");
  //     // Close modal if you want
  //     // setIsModalOpen(null);
  //   } catch (error) {
  //     console.error("Error toggling favourite:", error);
  //   }
  // };

  // const toggleFavourite = async (file: any
  // ) => {


  //   let isFav = false;

  //   try {

  //     console.log("Toggling favourite for file:", file);
  //     const siteurl = file.FilePreviewURL;
  //     console.log("siteurl", siteurl)

  //     const fileUniqueId = file.FileUID;
  //     console.log("File UniqueId:", fileUniqueId);


  //     const meName = context.pageContext.user.displayName || "";
  //     const meEmail =
  //       (context.pageContext as any)?.user?.email ||
  //       (context.pageContext as any)?.user?.loginName ||
  //       "";





  //     const baseSiteUrl = siteurl.split("/sites/")[0] + "/sites/";
  //     const siteCollectionName = siteurl.split("/sites/")[1].split("/")[0];


  //     // const subSiteName = siteurl.split("/").pop(); // TestHub1
  //     const subSiteName = siteurl.split("/sites/")[1].split("/")[1];


  //     const siteSP = spfi(`${baseSiteUrl}${siteCollectionName}`).using(
  //       SPFx(context)
  //     );

  //     const listName = `DMS${subSiteName}FileMaster`;





  //     console.log("BaseSiteUrl:", baseSiteUrl);
  //     console.log("Site Collection:", siteCollectionName);
  //     console.log("Sub Site:", subSiteName);




  //     const items = await siteSP.web.lists
  //       .getByTitle(listName)
  //       .items
  //       .select("Id", "FileUID", "IsFavourite", "CurrentUser")
  //       .filter(
  //         `FileUID eq '${fileUniqueId}'`
  //       ).top(1)();

  //     console.log("items Name:", items);

  //     console.log("Matched Items:", items);
  //     if (items.length === 0) {
  //       console.log("❌ No matching record found for this user & file");
  //       return;
  //     }
  //     const item = items[0];
  //     const newFavouriteValue = !item.IsFavourite;

  //     console.log("Old IsFavourite:", item.IsFavourite);
  //     console.log("New IsFavourite:", newFavouriteValue);

  //     await siteSP.web.lists
  //       .getByTitle(listName)
  //       .items
  //       .getById(item.Id)
  //       .update({
  //         IsFavourite: newFavouriteValue
  //       });

  //     console.log("✅ Favourite status updated successfully");
  //     const refreshed = await loadViewData("MyFavourite");
  //     if (refreshed) {
  //       setSelectedFiles([...refreshed]);
  //     }
  //     console.log("remove favorate marke  for:", file);
  //     return newFavouriteValue;


  //   } catch (e) {

  //   }
  //   // try {





  //   //   // // Build site URL dynamically
  //   //   // const currentUrl = new URL(context.pageContext.web.absoluteUrl);
  //   //   // const siteUrl = `${currentUrl.protocol}//${currentUrl.hostname}/sites/${file.SiteCollection}`;
  //   //   // const siteSP = spfi(siteUrl).using(SPFx(context));

  //   //   // console.log("Addhyan  kumar ->>>[toggleFavourite] siteUrl:", siteUrl);


  //   //   // // Update SharePoint item
  //   //   // await siteSP.web.lists
  //   //   //   .getByTitle(file.ListName)
  //   //   //   .items.getById(file.Id)
  //   //   //   .update({
  //   //   //     IsFavourite: !file.IsFavourite, // flip value
  //   //   //   });

  //   //   //   console.log("Addhyan   ->>>[toggleFavourite] siteUrl:", siteUrl);

  //   //   // console.log(
  //   //   //   `Updated favourite for ${
  //   //   //     file.Title || file.FileName
  //   //   //   } → ${!file.IsFavourite}`
  //   //   // );

  //   //   // // Refresh current view (example: MyFavourite)
  //   //   // await loadViewData("MyFavourite");
  //   //   // alert("Done!");

  //   // } catch (error) {



  //   //   console.error("Error toggling favourite:", error);
  //   // }
  // };


  // Addhyan 19/2/26
   const toggleFavourite = async (file: any,  siteUrl: string,
  context: any) => {
 
 
 
 
 
    console.log("File Object:", file);
 
    const siteUrls = file.SiteID || file.SiteID || siteUrl;
    console.log("Determined Site URL:", siteUrls);
 
 
 
 
 
 
  const fileUniqueId = file.FileUID;
 
  const meEmail =
    (context.pageContext as any)?.user?.email ||
    (context.pageContext as any)?.user?.loginName ||
    "";
 
  const baseSiteUrl = siteUrls.split("/sites/")[0] + "/sites/";
  const siteCollectionName = siteUrls.split("/sites/")[1].split("/")[0];
  const subSiteName = siteUrls.split("/").pop();
 
  const siteSP = spfi(`${baseSiteUrl}${siteCollectionName}`).using(
    SPFx(context)
  );
 
  const listName = `DMS${subSiteName}FileMaster`;
 
  // ✅ Important Fix: User-wise filter
  const items = await siteSP.web.lists
    .getByTitle(listName)
    .items
    .select("Id", "FileUID", "IsFavourite", "CurrentUser", "MyRequest", )
    .filter(
      `FileUID eq '${fileUniqueId}' and CurrentUser eq '${meEmail}' and MyRequest eq 0`
    )();
 
  console.log("Matched Items:", items);
 
  // 🔥 CASE 1: Record not found → Add new
  if (items.length === 0) {
 
    console.log("No record found → Adding new item");
 
   
 
    console.log("✅ New favourite added");
    const refreshed = await loadViewData("MyFavourite");
if (refreshed) {
  setSelectedFiles([...refreshed]);
} // Refresh the page to reflect changes
``
    return true;
  }
 
  // 🔥 CASE 2: Record exists → Toggle
  const item = items[0];
  const newFavouriteValue = !item.IsFavourite;
 
  await siteSP.web.lists
    .getByTitle(listName)
    .items
    .getById(item.Id)
    .update({
      IsFavourite: newFavouriteValue,
      // MyRequest: false
    });
 
  console.log("✅ Favourite status updated:", newFavouriteValue);
 
    //  window.location.reload(); // Refresh the page to reflect changes
     const refreshed = await loadViewData("MyFavourite");
if (refreshed) {
  setSelectedFiles([...refreshed]);
}
     
  return newFavouriteValue;
 
 
   
   
  };
  /** ----------------------------------- uploads (same) --------------------------------------- */
  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const files = Array.from(e.target.files);
      setSelectedUploadFiles(files);

      try {
        const getSiteCollectionUrl = (fullUrl: any): any => {
          const pattern = /^(https?:\/\/[^\/]+\/sites\/[^\/]+)/i;
          const match = fullUrl.match(pattern);
          return match ? match[1] : fullUrl;
        }
        console.log(
          "[upload] getSiteCollectionUrl Uploading files to preview location...",
          getSiteCollectionUrl(currentSiteUrl)
        );
        const siteSP = spfi(getSiteCollectionUrl(currentSiteUrl)).using(SPFx(context));
        const previewUrls: string[] = [];
        console.log("[upload] currentSiteUrl:", currentSiteUrl);
        console.log("[upload] siteSP:", siteSP);
        console.log(
          "[upload] preview folder path:",
          `/sites/${currentSiteUrl.split("/sites/")[1]}/DMSOrphanDocs`
        );

        for (const file of files) {
          const result = await siteSP.web
            .getFolderByServerRelativePath(`/sites/AlRostmaniSpfx2/DMSOrphanDocs`)
            .files.addChunked(file.name, file);

          previewUrls.push(result.data.ServerRelativeUrl);
        }

        setPreviewFileUrls(previewUrls);
      } catch (error) {
        console.error("[upload] Error uploading preview files:", error);
      }
    }
  };

  const triggerFileInput = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  // const uploadFiles = async () => {
  //   if (!currentSiteUrl || !currentFolderPath || selectedUploadFiles.length === 0) return;

  //   try {
  //     const siteSP = spfi(currentSiteUrl).using(SPFx(context));

  //     for (let i = 0; i < selectedUploadFiles.length; i++) {
  //       const file = selectedUploadFiles[i];
  //       setUploadProgress(((i + 1) / selectedUploadFiles.length) * 100);

  //       const uplaodfile = await siteSP.web
  //         .getFolderByServerRelativePath(
  //           `/sites/${currentSiteUrl.split("/sites/")[1]}/${currentFolderPath}`
  //         )
  //         .files.addChunked(file.name, file);
  //       console.log("[upload] File uploaded:", uplaodfile.data, JSON.stringify(uplaodfile.data));

  //       const encodeSharePointPath = (path: string): string => {
  //         return encodeURIComponent(path)
  //           .replace(/'/g, "%27")
  //           .replace(/\(/g, "%28")
  //           .replace(/\)/g, "%29")
  //           .replace(/\*/g, "%2A")
  //           .replace(/!/g, "%21")
  //           .replace(/#/g, "%23")
  //           .replace(/\$/g, "%24")
  //           .replace(/&/g, "%26")
  //           .replace(/\+/g, "%2B")
  //           .replace(/,/g, "%2C")
  //           .replace(/;/g, "%3B")
  //           .replace(/=/g, "%3D")
  //           .replace(/\?/g, "%3F")
  //           .replace(/\[/g, "%5B")
  //           .replace(/\]/g, "%5D")
  //           .replace(/_/g, "%5F")
  //           .replace(/\./g, "%2E")
  //           .replace(/-/g, "%2D");
  //       };

  //       const getSharePointPreviewUrl = (siteUrl: string, serverRelativePath: string): string => {
  //         const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf("/"));
  //         const encodedFilePath = encodeSharePointPath(serverRelativePath);
  //         const encodedParentPath = encodeSharePointPath(parentFolder);

  //         return `${siteUrl}${parentFolder.replace(/ /g, "%20")}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodedParentPath}`;
  //       };

  //       const siteUrl = "https://officeindia.sharepoint.com";
  //       const filePath = uplaodfile.data.ServerRelativeUrl;
  //       const previewUrl = getSharePointPreviewUrl(siteUrl, filePath);
  //       console.log("[upload] Preview URL:", previewUrl);
  //     }

  //     // Refresh the file list
  //     if (breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== "view") {
  //       const lastNode = nodeMap[breadcrumbs[breadcrumbs.length - 1].key];
  //       if (lastNode) {
  //         await loadFilesForNode(lastNode);
  //       }
  //     }

  //     setSelectedUploadFiles([]);
  //     setPreviewFileUrls([]);
  //     setShowUploadPanel(false);
  //     setUploadProgress(0);
  //   } catch (error) {
  //     console.error("[upload] Error uploading files:", error);
  //   }
  // }; hide it updated with new function by ritik 




  //Helper function to get root site URL // 19/01/2026 //Ritik

  const getRootSiteUrl = (url: string): string => {
    const match = url.match(/^(https?:\/\/[^\/]+\/sites\/[^\/]+)/i);
    if (!match) {
      throw new Error("Invalid SharePoint site URL");
    }
    return match[1];
  };

  const formatFileSizeForMeta = (bytes: number): string => {
    if (!bytes || bytes === 0) return "0 KB";

    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(2)} KB`;
    }

    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  };

  //Doc Lib upload function and metadata save to root site FileMaster list // 19/01/2026 // Ritik
  const uploadFiles = async () => {
    if (!currentSiteUrl || !currentFolderPath || selectedUploadFiles.length === 0) {
      return;
    }

    try {
      /* ===============================
         STEP 1: Contexts
         =============================== */

      const spSubsite = spfi(currentSiteUrl).using(SPFx(context)); // upload
      const rootSiteUrl = getRootSiteUrl(currentSiteUrl);
      const spRoot = spfi(rootSiteUrl).using(SPFx(context)); // metadata

      /* ===============================
         STEP 2: Resolve FileMaster List Name Dynamically
         =============================== */

      // Extract subsite name from currentSiteUrl
      // Example: https://officeindia.sharepoint.com/sites/edcspfx/Test%20Entity

      const urlParts = currentSiteUrl.split('/sites/')[1]?.split('/');
      const rootSiteName = urlParts?.[0]; // e.g., "edcspfx", "Intranetdemos", "AlRostmaniSpfx2"

      // Decode URL-encoded subsite name (Test%20Entity → Test Entity)
      let subsiteName = urlParts?.[1];
      if (subsiteName) {
        subsiteName = decodeURIComponent(subsiteName);
      }

      if (!subsiteName) {
        Swal.fire({
          icon: "error",
          title: "Invalid Context",
          text: "Cannot determine subsite. Please ensure you're uploading to a subsite document library.",
        });
        return;
      }

      // Construct FileMaster list name using naming convention
      // Pattern: DMS{SubsiteName}FileMaster
      // **KEEP SPACES** - SharePoint list names can have spaces
      const fileMasterListName = `DMS${subsiteName}FileMaster`;

      console.log("🔍 Derived FileMaster List:", fileMasterListName);
      console.log("📍 Root Site:", rootSiteUrl);
      console.log("📍 Subsite:", currentSiteUrl);
      console.log("📍 Subsite Name:", subsiteName);

      // Validate that the list exists on root site
      try {
        await spRoot.web.lists
          .getByTitle(fileMasterListName)
          .select("Title")();
        console.log("✅ FileMaster list found:", fileMasterListName);
      } catch (error) {
        console.error("❌ FileMaster list not found:", fileMasterListName, error);
        Swal.fire({
          icon: "error",
          title: "Configuration Error",
          html: `FileMaster list "<strong>${fileMasterListName}</strong>" not found on root site "${rootSiteName}".<br><br>Expected list: <strong>${fileMasterListName}</strong><br>Subsite: <strong>${subsiteName}</strong><br><br>Please contact administrator.`,
        });
        return;
      }

      /* ===============================
         STEP 3: Upload Loop (SUBSITE)
         =============================== */

      for (let i = 0; i < selectedUploadFiles.length; i++) {
        const file = selectedUploadFiles[i];
        setUploadProgress(((i + 1) / selectedUploadFiles.length) * 100);

        const siteRelativePath = currentSiteUrl.replace(window.location.origin, "");
        const serverRelativePath =
          siteRelativePath +
          (currentFolderPath.startsWith("/") ? "" : "/") +
          currentFolderPath;

        console.log("📤 Uploading to:", serverRelativePath);

        const uploadResult = await spSubsite.web
          .getFolderByServerRelativePath(serverRelativePath)
          .files.addChunked(file.name, file);

        const uploadedFile = uploadResult.data;
        console.log("✅ File uploaded:", uploadedFile.Name);

        /* ===============================
           STEP 4: Dynamic metadata (ROOT)
           =============================== */

        const fields = await spRoot.web.lists
          .getByTitle(fileMasterListName)
          .fields.select("InternalName")();

        const allowed = fields.map(f => f.InternalName);



        // 🔹 STEP: Build full site-relative folder path (NO hard-code)
        const siteRelativeUrl = currentSiteUrl.replace(window.location.origin, "");

        const fullFolderPath =
          siteRelativeUrl +
          (currentFolderPath.startsWith("/") ? "" : "/") +
          currentFolderPath;



        // ✅ ALWAYS use serverRelativePath / fullFolderPath to extract folder
        let folderName = "";

        const normalizedFullPath = fullFolderPath.replace(/^\/+/, "");
        const pathParts = normalizedFullPath.split("/");

        // example:
        // sites / AlRostmaniSpfx2 / TestHub1 / Approval Temp Lib / Test8
        // index 0   1               2          3                4

        if (pathParts.length > 4) {
          folderName = pathParts[pathParts.length - 1];
        }






        const rawPayload: Record<string, any> = {
          Title: uploadedFile.Name,
          FileName: uploadedFile.Name,
          FileUID: uploadedFile.UniqueId,

          FileVersion: "1",
          FileSize: formatFileSizeForMeta(file.size),

          IsFavourite: false,     // Yes/No → boolean
          MyRequest: true,       // Yes/No → boolean
          IsDeleted: null,       // ✅ DateTime column FIX

          CurrentUser: context.pageContext.user.email,

          DocumentLibraryName: selectedCurrentNode?.libraryTitle || "",
          FolderName: folderName,
          CurrentFolderPath: fullFolderPath,

          ShareWithMe: "",
          ShareWithOthers: "",

          Status: "Pending",

          SiteID: context.pageContext.site.id.toString(),
          SiteName: subsiteName,

          FilePreviewURL: `${window.location.origin}${uploadedFile.ServerRelativeUrl}`,

          Processname: "New File Request",
          RequestNo: uploadedFile.Name.split(".")[0],
        };





        const safePayload: Record<string, any> = {};
        for (const key in rawPayload) {
          if (allowed.includes(key)) {
            safePayload[key] = rawPayload[key];
          }
        }

        console.log("💾 Saving metadata to:", fileMasterListName);
        console.log("📝 Payload:", safePayload);

        await spRoot.web.lists
          .getByTitle(fileMasterListName)
          .items.add(safePayload);

        console.log("✅ Metadata saved for:", uploadedFile.Name);
      }

      /* ===============================
         STEP 5: UI reset
         =============================== */

      setSelectedUploadFiles([]);
      setPreviewFileUrls([]);
      setShowUploadPanel(false);
      setUploadProgress(0);

      Swal.fire({
        icon: "success",
        title: "Upload Successful",
        text: `File(s) uploaded to ${subsiteName} and metadata saved to ${fileMasterListName}.`,
      });

      // Refresh file list
      if (selectedCurrentNode) {
        await loadFilesForNode(selectedCurrentNode);
      }

    } catch (err) {
      console.error("❌ uploadFiles failed:", err);
      Swal.fire({
        icon: "error",
        title: "Upload Failed",
        text: err?.message || "Please check console for details.",
      });
    }
  };

  /** --------------------------------------- render ------------------------------------------- */
  const renderTree = (nodes: TreeNode[]) => (
    <ul style={{ listStyleType: "none", paddingLeft: "5px" }}>
      {nodes.map((node) => (
        <li key={node.key}>
          <div style={{ display: "flex", alignItems: "center" }}>
            {node.hasChildren && (
              <button type="button"
                onClick={() => toggleNode(node)}
                className="arrowdesign"
              >
                {node.isExpanded ? "−" : "+"}
              </button>
            )}
            <span
              id="location"
              onClick={() => handleNodeClick(node)}
              style={{
                cursor: "pointer",
                fontWeight: breadcrumbs.some((b) => b.key === node.key) ? "bold" : "normal",
                color: breadcrumbs.some((b) => b.key === node.key) ? "#fa901d" : "#000  ",
                padding: "2px 5px",
                borderRadius: "3px",
                fontSize: "14px",
              }}
            >
              {node.title}
            </span>
          </div>
          {node.isExpanded && node.children && renderTree(node.children)}
        </li>
      ))}
    </ul>
  );


  /* ----------------- AUDIT HISTORY FUNCTION------------------ */
  // const handleAuditHistory = async (file: any) => {
  //   try {
  //     setAuditLoading(true);
  //     setModalFile(file);
  //     setShowAuditModal(true);

  //     /*---------------build siteUrl for handleAuditHistory---------------*/
  //     let siteUrl: string = (file && (file.__siteUrl || file.SiteUrl || file.SiteURL)) || currentSiteUrl || context.pageContext?.site?.absoluteUrl;
  //     if (file && file.SiteName) {
  //       siteUrl = `${siteUrl.replace(/\/$/, '')}/${file.SiteName.replace(/^\//, '')}`;
  //     }

  //     const spnew = spfi(siteUrl).using(SPFx(context));

  //     console.log('Audit history (File)');
  //     console.log('fileUID', file?.FileUID || file?.fileUID || file?.UniqueId || file?.Id || file?.ID);
  //     console.log('siteUrl', siteUrl);

  //     // determine server relative path robustly
  //     const srCandidates = [
  //       file?.ServerRelativeUrl,
  //       file?.ServerRelativeUrlDecoded,
  //       (file?.ServerRelativePath && file.ServerRelativePath.DecodedUrl) ? file.ServerRelativePath.DecodedUrl : null,
  //       file?.FileRef,
  //       file?.FileLeafRef ? (file.FileRef ? file.FileRef : null) : null,
  //       (file?.CurrentFolderPath && file?.FileName) ? `${file.CurrentFolderPath}/${file.FileName}` : null,
  //       file?.ServerRelativePath
  //     ].filter(x => x !== undefined && x !== null);

  //     let serverRel: any = null;
  //     if (srCandidates.length > 0) {
  //       serverRel = srCandidates.find(s => !!s) || null;
  //       if (serverRel && typeof serverRel === 'string' && !serverRel.startsWith('/')) {
  //         serverRel = '/' + serverRel;
  //       }
  //     }

  //     console.log('serverRel candidate', serverRel);

  //     let fileItem: any = null;
  //     let listItem: any = null;
  //     let versions: any[] = [];

  //     if (serverRel) {
  //       try {
  //         fileItem = await spnew.web.getFileByServerRelativePath(serverRel).getItem();
  //         console.log('fileItem from getFileByServerRelativePath', fileItem);
  //       } catch (e) {
  //         console.warn('getFileByServerRelativePath.getItem() failed', e);
  //       }

  //       try {
  //         const fileObj = spnew.web.getFileByServerRelativePath(serverRel);
  //         const vers = await fileObj.versions();
  //         versions = (vers || []).map((v: any) => ({
  //           VersionLabel: v?.VersionLabel || v?.Version || '',
  //           Created: v?.Created ? new Date(v.Created).toLocaleString() : '',
  //           ModifiedByName: v?.CreatedBy?.Title || v?.CreatedBy?.Name || '',
  //           ModifiedByEmail: v?.CreatedBy?.Email || v?.CreatedBy?.LoginName || '',
  //           SizeDisplay: (v?.Size ? (v.Size / 1024) : v?.Length ? (v.Length / 1024) : 0).toFixed(2) + ' KB'
  //         }));
  //       } catch (e) {
  //         console.warn('versions() fetch failed', e);
  //       }
  //     }

  //     // if fileItem not found, try to get list item from FileMaster list using Id
  //     const listTitle = file.__fileMasterList || file.FileMaster || 'FileMaster';
  //     const id = file.Id || file.ID;
  //     if ((!fileItem || (Object.keys(fileItem || {}).length === 0)) && listTitle && id) {
  //       try {
  //         listItem = await spnew.web.lists.getByTitle(listTitle).items.getById(id).select('*', 'Editor/Title', 'Editor/Email').expand('Editor')();
  //         console.log('list item from FileMaster list', listItem);
  //       } catch (e) {
  //         console.warn('fetch FileMaster list item failed', e);
  //       }
  //     }

  //     // Compose metadata using fileItem, listItem, or the file object
  //     const src = fileItem || listItem || file || {};
  //     console.log('source for metadata', src);
  //     const metadata: any = {
  //       Title: src?.Title || file?.Title || file?.FileName || '',
  //       FileName: src?.FileLeafRef || file?.FileName || '',
  //       FileUID: file?.FileUID || file?.UniqueId || file?.GUID || '',
  //       Status: src?.Status || file?.Status || '',
  //       IsDeleted: (src?.IsDeleted ?? file?.IsDeleted ?? listItem?.IsDeleted) ? 'Yes' : 'No',
  //       Modified: src?.Modified ? new Date(src.Modified).toLocaleString() : (file?.Modified ? new Date(file.Modified).toLocaleString() : ''),
  //       ModifiedBy: (src?.Editor?.Title || src?.Editor?.Name || file?.ModifiedBy || '')
  //     };
  //     console.log('composed metadata', metadata);
  //     // If no versions found, try to synthesize single current entry
  //     if ((!versions || versions.length === 0) && (fileItem || listItem)) {
  //       const s = fileItem || listItem;
  //       versions = [{
  //         VersionLabel: s?.OData__UIVersionString || s?.VersionLabel || '1.0',
  //         Created: s?.Modified ? new Date(s.Modified).toLocaleString() : '',
  //         ModifiedByName: s?.Editor?.Title || '',
  //         ModifiedByEmail: s?.Editor?.Email || '',
  //         SizeDisplay: s?.Length ? (s.Length / 1024).toFixed(2) + ' KB' : ''
  //       }];
  //     }

  //     setAuditVersions({ Metadata: metadata, Versions: versions });
  //   } catch (err) {
  //     console.error('Audit history fetch error:', err);
  //     setAuditVersions({ Metadata: {}, Versions: [] });
  //   } finally {
  //     setAuditLoading(false);
  //   }
  // }




  // Aman 19/2/26

  
const handleAuditHistory = async (file: any) => {
  console.log("Initiating audit history for file:", file);



  try {
    setAuditLoading(true);
    setModalFile(file);
    setShowAuditModal(true);
 
    /* ---------- Path & Subsite Resolution (LOGIC FROM WORKING VERSION HISTORY) ---------- */
    const folderPath = file.CurrentFolderPath || file.ServerRelativeUrl || "";
    const fileName = file.FileName || file.Name;
 
    let serverRelativePath = "";
    if (file.ServerRelativeUrl) {
      serverRelativePath = file.ServerRelativeUrl;
    } else {
      serverRelativePath = folderPath.endsWith("/")
        ? `${folderPath}${fileName}`
        : `${folderPath}/${fileName}`;
    }
    const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
    console.log("Tenant URL:", tenantUrl);
    const parts = folderPath.split("/").filter(Boolean);
    console.log("Folder path parts:", parts);
    const subsitePath2 = "/" + parts.slice(0, 2).join("/");
    console.log("Subsite path 2:", subsitePath2);
    const subsitePath = "/" + parts.slice(0, 3).join("/");

    console.log("Derived subsite path:", subsitePath);
    const subsiteUrl = `${tenantUrl}${subsitePath}`;
    console.log("Calculated subsite URL for audit:", subsiteUrl);
    const subsiteUrl2 = `${tenantUrl}${subsitePath2}`;
    console.log("Calculated subsite URL 2 for audit:", subsiteUrl2);
 
    // Initialize SP using the calculated subsite URL
    const siteSP = spfi(subsiteUrl).using(SPFx(context));

     const siteSP2 = spfi(subsiteUrl2).using(SPFx(context));
    // const approvalItems = await siteSP2.web.lists
    //   .getByTitle("DMSFileApprovalTaskList")
    //   .items
    //   .select("CurrentUser", "Remark", "ApprovalType", "LogHistory", "Log", "FileUID")
    //   .filter(`FileUID eq '${file.FileUID}'`)
    //   .getAll();
    //   console.log("Approval items for file:", approvalItems);

  //   const approvalItems = await siteSP2.web.lists
  // .getByTitle("DMSFileApprovalTaskList")
  // .items
  // .select(
  //   "CurrentUser",
  //   "Remark",
  //   "ApprovalType",
  //   "LogHistory",
  //   "Log",
  //   "FileUID/Id",
  //   "FileUID/Title"
  // )
  // .expand("FileUID")
  // .filter(`FileUID/Id eq '${file.FileUID}'`) // 👈 important change
  // .getAll();
  // console.log("Approval items for file:", approvalItems);


  const approvalItems = await siteSP2.web.lists
      .getByTitle("DMSFileApprovalTaskList")
      .items
      .select(
        "CurrentUser",
        "Remark",
        // "ApprovalType",
        "LogHistory",
        "Log",
        "FileUID/FileUID",   // 👈 inner field
        "MasterApproval/Level",
         "FileUID/Status",
      )
      .expand("FileUID", "MasterApproval")    // 👈 MUST for lookup
      .filter(`FileUID/FileUID eq '${file.FileUID || file.UniqueId}'`) // 👈 main fix
      .getAll();

      console.log("Approval items for file:", approvalItems);



 
    let fetchedFileItem: any = null;
    let fetchedListItem: any = null;
    let fileProps: any = null;
    let versions: any[] = [];

    /* ---------- Fetch Data (Same logic as VersionHistory) ---------- */
    if (serverRelativePath) {
      try {
        const fileItemObj = siteSP.web.getFileByServerRelativePath(serverRelativePath);
 
        // 1. Fetch File Props
        fileProps = await fileItemObj
          .select(
            "TimeLastModified",
            "Length",
            "Name",
            "ServerRelativeUrl",
            "UIVersionLabel",
            "ModifiedBy/Title",
            "ModifiedBy/UserPrincipalName"
          )
          .expand("ModifiedBy")();
 
        // 2. Fetch Version Data
        const versionData = await fileItemObj.versions
          .select("ID,VersionLabel,Created,Size,Url,CreatedBy/Title,CreatedBy/UserPrincipalName")
          .expand("CreatedBy")();
        // 3. Metadata (detailedItem)
        const itemQuery = await fileItemObj.getItem();
        const detailedItem = await itemQuery
          .select("*", "Editor/Title")
          .expand("Editor")();
 
        fetchedFileItem = { ...fileProps, ...detailedItem };
 
        // 4. Combine and Sort Versions (Matching your working logic)
        const allVersions = [
          ...versionData.map((v: any) => ({
            VersionLabel: v.VersionLabel,
            Created: new Date(v.Created).toLocaleString(),
            ModifiedByName: v.CreatedBy?.Title || "-",
            SizeDisplay: (v.Size / 1024).toFixed(2) + " KB"
          })),
          {
            VersionLabel: fileProps.UIVersionLabel,
            Created: new Date(fileProps.TimeLastModified).toLocaleString(),
            ModifiedByName: fileProps.ModifiedBy?.Title || "-",
            SizeDisplay: (fileProps.Length / 1024).toFixed(2) + " KB"
          }
        ];
 
        // Sort descending to show latest first in Audit
        versions = allVersions.reverse();
 
      } catch (e) {
        console.warn("Audit: Data fetch failed", e);
      }
    }
 
    /* ---------- Fallback for Metadata List ---------- */
    const listTitle = file.__fileMasterList || file.FileMaster || "FileMaster";
    const id = file.Id || file.ID;
 
    if (!fetchedFileItem && listTitle && id) {
      try {
        fetchedListItem = await siteSP.web.lists
          .getByTitle(listTitle)
          .items.getById(id)
          .select("*", "Editor/Title")
          .expand("Editor")();
      } catch (e) {
        console.warn("Audit: List fallback failed", e);
      }
    }


    
 
    /* ---------- Final Metadata Mapping ---------- */
    const src = fetchedFileItem || fetchedListItem || file || {};
    // const metadata: any = {
    //   Title: src?.Title || src?.Name || file?.Title || file?.FileName || file?.Name || "",
    //   FileName: src?.FileLeafRef || src?.Name || file?.FileName || file?.Name || "",
    //   // FileUID: file?.FileUID || file?.UniqueId || file?.GUID || "",  -- addhyan 09/04/2026
    //   Status: src?.Status || file?.Status || "",
    //   // IsDeleted: (src?.IsDeleted ?? file?.IsDeleted ?? fetchedListItem?.IsDeleted) ? "Yes" : "No", -- addhyan 09/04/2026
    //   Modified: src?.Modified ? new Date(src.Modified).toLocaleString() :
    //             src?.TimeLastModified ? new Date(src.TimeLastModified).toLocaleString() :
    //             fileProps?.TimeLastModified ? new Date(fileProps.TimeLastModified).toLocaleString() :
    //             file?.Modified ? new Date(file.Modified).toLocaleString() : "-",
    //   ModifiedBy: src?.Editor?.Title || src?.ModifiedBy?.Title || fileProps?.ModifiedBy?.Title || file?.ModifiedBy || "-",
    //   ...src // include all other available fields for display
    // };

//     const metadata: any = Object.fromEntries(
//   Object.entries(src).filter(([key, value]) => {
    
//     // ❌ Skip odata fields
//     if (key.startsWith("odata")) return false;

//     // ❌ Skip system fields
//     const systemFields = [
//       "ID", "Id", "GUID",
//       "Created", "Modified",
//       "AuthorId", "EditorId",
//       "Attachments",
//       "OData__CopySource",
//       "CheckoutUserId",
//       "FileSystemObjectType",
//       "OData__UIVersionString",
//       "ContentTypeId",
//       "ServerRedirectedEmbedUri",
//       "OData__ColorTag",
//       "ComplianceAssetId",
//       "MediaServiceOCR",
//       "ServerRedirectedEmbedUrl",
//       "Length", "ServerRelativeUrl",
//       "TimeLastModified", "UIVersionLabel"
//     ];

//     if (systemFields.includes(key)) return false;

//     // ❌ Skip objects (Editor, ModifiedBy etc)
//     if (typeof value === "object" && value !== null) return false;

//     // ❌ Skip navigation links
//     if (key.includes("@odata")) return false;

//     // ✅ Keep rest (custom columns)
//     return true;
//   })
// );


const systemFields = [
      "ID", "Id", "GUID",
      "Created", "Modified",
      "AuthorId", "EditorId",
      "Attachments",
      "OData__CopySource",
      "CheckoutUserId",
      "FileSystemObjectType",
      "OData__UIVersionString",
      "ContentTypeId",
      "ServerRedirectedEmbedUri",
      "OData__ColorTag",
      "ComplianceAssetId",
      "MediaServiceOCR",
      "ServerRedirectedEmbedUrl",
      "Length", "ServerRelativeUrl",
      "TimeLastModified", "UIVersionLabel"
    ];

const customFields = Object.fromEntries(
  Object.entries(src).filter(([key, value]) => {

    if (key.startsWith("odata")) return false;
    if (key.includes("@odata")) return false;
    if (systemFields.includes(key)) return false;
    if (typeof value === "object" && value !== null) return false;
    if (
      value === null ||
      value === undefined ||
      value === "" ||
      value === "-" 
      // (typeof value === "string" && value.trim() === "")
    ) return false;

    return true;
  })
);

// STEP 2: final metadata (ADD THIS 👇)
const metadata: any = {
  ...customFields,
// aman comment this to make data fomrat dd/mm/yyyy
  // Modified:
  //   src?.Modified
  //     ? new Date(src.Modified).toLocaleString()
  //     : src?.TimeLastModified
  //     ? new Date(src.TimeLastModified).toLocaleString()
  //     : fileProps?.TimeLastModified
  //     ? new Date(fileProps.TimeLastModified).toLocaleString()
  //     : file?.Modified
      // ? new Date(file.Modified).toLocaleString() 
      
      // : "-",
      // aman comment this to make data fomrat dd/mm/yyyy
Modified: (() => {
    const rawDate = src?.Modified || src?.TimeLastModified || fileProps?.TimeLastModified || file?.Modified;
    if (!rawDate || rawDate === "-") return "-";
   
    const d = new Date(rawDate);
    // Format to dd/mm/yyyy
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
    const year = d.getFullYear();
   
    return `${day}/${month}/${year}`;
  })(),
  ModifiedBy:
    src?.Editor?.Title ||
    src?.ModifiedBy?.Title ||
    fileProps?.ModifiedBy?.Title ||
    file?.ModifiedBy ||
    "-"
};
 
    setAuditVersions({
      Metadata: metadata,
      Versions: versions,
       ApprovalData: approvalItems 
    });
 
  } catch (err) {
    console.error("Audit history error:", err);
    setAuditVersions({ Metadata: {}, Versions: [], ApprovalData: [] });
  } finally {
    setAuditLoading(false);
  }
};
 
  /* ----------------- end handleAuditHistory ------------------ */
  // srs 10/4/26

  // ritik 20/04/26 - this code replace by ritik for manage permission function (start)
// window.ManageFilePermission = async (
//     fileId: string,
//     siteUrl: string, 
//     documentLibraryName: string,
//     siteTitle: string
// ) => {
//     try {
//         const web = Web(siteUrl).using(AssignFrom(sp.web));

//         // 1. Fetch File and Item
//         const file = web.getFileById(fileId);
//         const item = await file.getItem();
        
//         // 2. Fresh check for Inheritance
//         const itemQuery = item.select("HasUniqueRoleAssignments", "Id");
//         itemQuery.query.set("v", Date.now().toString());
//         const itemInfo: any = await itemQuery();
        
//         const isUnique = itemInfo.HasUniqueRoleAssignments === true;

//         // 3. Fetch assignments - Cast to any[] to avoid the 'Member' property error
//         const currentAssignments: any[] = await item.roleAssignments
//             .expand("Member", "RoleDefinitionBindings")();

//         const allUsers = await web.siteUsers.select("Id", "Title", "Email")();
//         const filteredUsers = allUsers.filter(u => u.Email && !u.Title.includes("System"));

//         let selectedUsers: Array<{ id: string, title: string }> = [];

//         // 4. Management Dialog
//         const { value: formValues, isDenied: revertClicked } = await Swal.fire({
//             title: 'Manage File Permission',
//             width: '600px',
//             padding: '1rem',
//             backdrop: false,
//             showCloseButton: true,
//             customClass: { popup: 'sharepoint-style-shadow' },
//             html: `
//                 <style>
//                     .sharepoint-style-shadow { box-shadow: 0 0 20px rgba(0,0,0,0.2) !important; border: 1px solid #ddd !important; }
//                     .user-option:hover { background-color: #f3f2f1 !important; color: #0078d4; }
//                     .status-indicator { 
//                         margin-bottom: 15px; margin-top:8px; text-align:left; padding: 10px; font-size: 12px; border-radius: 4px; 
//                         border: 1px solid ${isUnique ? '#fbc7c7' : '#c7ebc7'}; 
//                         background: ${isUnique ? '#fff4f4' : '#f3fbf3'}; 
//                         color: ${isUnique ? '#d13438' : '#107c10'}; 
//                     }
//                     .selected-tags-container { display: flex; flex-wrap: wrap; gap: 5px; margin-bottom: 10px; min-height: 35px; border: 1px solid #ddd; padding: 5px; border-radius: 4px; background: #faf9f8; }
//                     .user-tag { background: #0078d4; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; display: flex; align-items: center; gap: 5px; }
//                     .remove-tag { cursor: pointer; font-weight: bold; }
//                     .user-management-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 12px; border: 1px solid #eee; }
//                     .user-management-table th { background: #f3f2f1; padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
//                     .user-management-table td { padding: 6px 8px; text-align:left; border-bottom: 1px solid #eee; }
//                     .delete-user-btn { color: #d13438; cursor: pointer; font-size: 18px; background: none; border: none; font-weight: bold; }
//                 </style>

//                 <div>
//                     <div class="status-indicator">
//                         <strong>Current Status:</strong> ${isUnique ? '⚠️ Unique Permissions' : '✅ Inheriting Permissions'}
//                     </div>
//                     <label style="font-weight: 600; display: block; margin-bottom: 5px; text-align:left;">Add New Users:</label>
//                     <div id="selected-users-tags" class="selected-tags-container">
//                         <span style="color: #999; font-size: 12px;">Search and click users...</span>
//                     </div>
//                     <div style="position: relative;">
//                         <input type="text" id="user-search-input" class="swal2-input" placeholder="Type name..." style="width: 100%; margin: 0; font-size: 14px;">
//                         <div id="user-dropdown-list" style="display: none; position: absolute; top: 100%; left: 0; right: 0; background: white; border: 1px solid #d9d9d9; max-height: 180px; overflow-y: auto; z-index: 9999; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
//                             ${filteredUsers.map(u => `
//                                 <div class="user-option" data-id="${u.Id}" data-title="${u.Title}" style="padding: 10px; cursor: pointer; border-bottom: 1px solid #f4f4f4;">
//                                     <div style="font-weight: 600; text-align:left; font-size: 13px; color:#000">${u.Title}</div>
//                                     <div style="font-size: 11px; color: #777; text-align:left;">${u.Email}</div>
//                                 </div>
//                             `).join('')}
//                         </div>
//                     </div>
//                     <label style="font-weight: 600; display: block; margin: 20px 0 5px 0; text-align:left;">Permission Level:</label>
//                     <select id="swal-permission-level" class="swal2-select" style="width: 100%; margin: 0; font-size: 14px; margin-bottom: 20px;">
//                         <option value="Full Control">Full Control</option>
//                         <option value="Edit">Edit</option>
//                         <option value="Contribute">Contribute</option>
//                         <option value="Read" selected>Read</option>
//                     </select>

//                     ${isUnique ? `
//                     <hr style="border: 0; border-top: 1px solid #eee; margin: 15px 0;">
//                     <label style="font-weight: 600; display: block; margin-bottom: 5px; text-align:left">Existing Access:</label>
//                     <div style="max-height: 150px; overflow-y: auto;">
//                         <table class="user-management-table">
//                             <thead>
//                                 <tr><th>User</th><th>Permission</th><th style="text-align:center">Action</th></tr>
//                             </thead>
//                             <tbody>
//                                 ${currentAssignments
//                                     .filter(a => (a.Member?.Title || "").trim() !== "DMSSuper_Admin")
//                                     .map(a => `
//                                     <tr>
//                                         <td style="text-align:left">
//                                             <div style="font-weight:600">${a.Member?.Title || 'Unknown'}</div>
//                                             <div style="font-size:10px; color:#666">${a.Member?.Email || ''}</div>
//                                         </td>
//                                         <td style="text-align:left">${a.RoleDefinitionBindings.map((r: any) => r.Name).join(', ')}</td>
//                                         <td style="text-align:center">
//                                             <button class="delete-user-btn" data-pid="${a.PrincipalId}">&times;</button>
//                                         </td>
//                                     </tr>`).join('')}
//                             </tbody>
//                         </table>
//                     </div>
//                     ` : ''}
//                 </div>
//             `,
//             showCancelButton: true,
//             confirmButtonText: 'Grant Access',
//             confirmButtonColor: '#0078d4',
//             showDenyButton: !!isUnique,
//             denyButtonText: 'Restore Inheritance',
//             denyButtonColor: '#6e7881',
//             didOpen: () => {
//                 const input = document.getElementById('user-search-input') as HTMLInputElement;
//                 const list = document.getElementById('user-dropdown-list') as HTMLDivElement;
//                 const tagsContainer = document.getElementById('selected-users-tags') as HTMLDivElement;
//                 const options = list.querySelectorAll('.user-option');

//                 document.querySelectorAll('.delete-user-btn').forEach(btn => {
//                     btn.addEventListener('click', async (e) => {
//                         const pid = (e.currentTarget as HTMLElement).getAttribute('data-pid');
//                         Swal.close(); 
//                         const confirm = await Swal.fire({ title: 'Remove User?', icon: 'warning', showCancelButton: true });
//                         if (confirm.isConfirmed) {
//                             await item.roleAssignments.getById(parseInt(pid!)).delete();
//                             window.ManageFilePermission(fileId, siteUrl, documentLibraryName, siteTitle);
//                         } else {
//                             window.ManageFilePermission(fileId, siteUrl, documentLibraryName, siteTitle);
//                         }
//                     });
//                 });

//                 const renderTags = () => {
//                     if (selectedUsers.length === 0) {
//                         tagsContainer.innerHTML = '<span style="color: #999; font-size: 12px;">Search and click users...</span>';
//                         return;
//                     }
//                     tagsContainer.innerHTML = selectedUsers.map(u => `
//                         <span class="user-tag">${u.title} <span class="remove-tag" data-id="${u.id}">&times;</span></span>
//                     `).join('');
//                     tagsContainer.querySelectorAll('.remove-tag').forEach(btn => {
//                         btn.addEventListener('click', (e) => {
//                             const id = (e.target as HTMLElement).getAttribute('data-id');
//                             selectedUsers = selectedUsers.filter(user => user.id !== id);
//                             renderTags();
//                         });
//                     });
//                 };

//                 input.addEventListener('input', () => {
//                     const val = input.value.toLowerCase();
//                     list.style.display = val ? 'block' : 'none';
//                     options.forEach((opt: any) => {
//                         const text = opt.innerText.toLowerCase();
//                         opt.style.display = text.includes(val) ? 'block' : 'none';
//                     });
//                 });

//                 options.forEach((opt: any) => {
//                     opt.addEventListener('click', () => {
//                         const id = opt.getAttribute('data-id');
//                         const title = opt.getAttribute('data-title');
//                         if (!selectedUsers.find(u => u.id === id)) {
//                             selectedUsers.push({ id, title });
//                             renderTags();
//                         }
//                         input.value = "";
//                         list.style.display = 'none';
//                     });
//                 });
//             },
//             preConfirm: () => {
//                 const permission = (document.getElementById('swal-permission-level') as HTMLSelectElement).value;
//                 if (selectedUsers.length === 0) {
//                     Swal.showValidationMessage('Please select at least one user');
//                     return false;
//                 }
//                 return { users: selectedUsers, permission };
//             }
//         });

//         // 5. ACTION: RESTORE INHERITANCE
//         if (revertClicked) {
//             Swal.fire({ title: 'Restoring...', didOpen: () => Swal.showLoading()});
//             await item.resetRoleInheritance();
//             window.ManageFilePermission(fileId, siteUrl, documentLibraryName, siteTitle);
//             return;
//         }

//         // 6. ACTION: GRANT ACCESS (WITH CLEAN BREAK LOGIC)
//         if (formValues) {
//             Swal.fire({ title: 'Updating...', allowOutsideClick: false, didOpen: () => Swal.showLoading() });

//             const adminGroupName = `${siteTitle}_Admin`.trim();
//             const superAdminName = "DMSSuper_Admin";

//             if (!isUnique) {
//                 await item.breakRoleInheritance(true); 

//                 // FIX: Cast assignments to any[] to allow accessing Member property
//                 const assignments: any[] = await item.roleAssignments.expand("Member").select("PrincipalId", "Member/Title")();

//                 for (const assignment of assignments) {
//                     const title = (assignment.Member?.Title || "").trim();
//                     if (title !== superAdminName && title !== adminGroupName) {
//                         try {
//                             await item.roleAssignments.getById(assignment.PrincipalId).delete();
//                         } catch (e) {
//                             console.warn(`Cleanup: Could not remove ${title}`);
//                         }
//                     }
//                 }
//             }

//             const roleDef = await web.roleDefinitions.getByName(formValues.permission)();
//             for (const user of formValues.users) {
//                 await item.roleAssignments.add(parseInt(user.id), roleDef.Id);
//             }

//             Swal.fire({ icon: 'success', title: 'Permissions Set', timer: 1500, showConfirmButton: false }).then(() => {
//                 window.ManageFilePermission(fileId, siteUrl, documentLibraryName, siteTitle);
//             });
//         }

//     } catch (error) {
//         console.error("Permission Logic Error:", error);
//         Swal.fire({ icon: "error", title: "Operation Failed", text: error.message });
//     }
// };

window.ManageFilePermission = async (
    fileId: string,
    siteUrl: string,
    documentLibraryName: string,
    siteTitle: string
  ) => {
    try {
      const web = Web(siteUrl).using(AssignFrom(sp.web));
      const file = web.getFileById(fileId);
      const item = await file.getItem();
  
      const itemQuery = item.select("HasUniqueRoleAssignments", "Id");
      itemQuery.query.set("v", Date.now().toString());
      const itemInfo: any = await itemQuery();
      const isUnique = itemInfo.HasUniqueRoleAssignments === true;
  
      const currentAssignments: any[] = await item.roleAssignments
        .expand("Member", "RoleDefinitionBindings")();
  
      const allUsers = await web.siteUsers.select("Id", "Title", "Email")();
      const filteredUsers = allUsers.filter(
        (u: any) => u.Email && !u.Title.includes("System")
      );
  
      setFilePermState({
        fileId,
        siteUrl,
        documentLibraryName,
        siteTitle,
        isUnique,
        currentAssignments,
        filteredUsers,
        selectedUsers: [] as Array<{id: string; title: string}>,
        permission: "Read",
      });
      setShowFilePermModal(true);
  
    } catch (error: any) {
      console.error("Permission Logic Error:", error);
      Swal.fire({
        icon: "error",
        title: "Operation Failed",
        text: error.message,
      });
    }
  };

// ritik 20/04/26 end 


// srs 15/4/26

 const checkValidation=(message:any)=>{
  Swal.fire(`${message}`,``, "warning");
}

const showReplaceMessage=(message:any)=>{
  Swal.fire(`${message}`,``, 'success');
}

   const currentUserEmailRef = useRef('');
  const currentUserIDref = useRef<number>(0);
  const currentUserTitleRef = useRef('');
    useEffect(() => {
     getcurrentuseremail()
     
}, []);

 const getcurrentuseremail = async()=>{
  const userProfile = await sp.profiles.myProperties();
  console.log(userProfile , "userProfile")
  console.log(userProfile.Title , "userProfile userProfile.Title")
  const userdata = await sp.web.currentUser();

  console.log(userdata , "user data edc")
  console.log(userdata.Id , "user data edc")
  currentUserIDref.current = userdata.Id;
  currentUserEmailRef.current = userdata.Email;
  currentUserTitleRef.current = userdata.Title;
  // console.log(currentUserEmailRef.current, "currentuser")
 }

window.rework=async(fileId:any,siteId:any,documentLibrary:any,siteName:any,filePath:any)=>{
// FIX 1: Use a unique name for segments to avoid redeclaration error
  // Define it here so it's recognized in this "block"
  const pathSegments = filePath.split('/'); 
  const locationPath = `/${pathSegments[1]}/${pathSegments[2]}`; 
  const masterSiteUrl = window.location.origin + locationPath;

  // Use sp.web here to satisfy the Timeline requirement
  const subsiteSp = spfi(siteId).using(AssignFrom(sp.web));
  const web = subsiteSp.web;

  const masterSp = spfi(masterSiteUrl).using(AssignFrom(sp.web));
  const masterWeb = masterSp.web;
  
  let clickedReplace=false;

  
  // Get the list item  corresponding to the file
  const fileItem:any = await web.getFileById(fileId).expand("ListItemAllFields")();
  console.log("fileItem",fileItem.ListItemAllFields.Status);
  
 
 // fetched the columns details corresponding to the file 
 const fileColumns =await masterWeb.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName","SiteName","DocumentLibraryName","IsRequired","ColumnType").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and IsDocumentLibrary ne 1`)();
 console.log("fileColumns",fileColumns);

 // Create an array of objects to store the columnName with there corresponding value
 const resultArrayThatContainstheColumnDetails = fileColumns.map((column) => {
 const columnName = column.ColumnName;
 const columnTpye=column.ColumnType;
 const columnRequired=column.IsRequired;
 const columnValue = fileItem.ListItemAllFields[columnName];

   return {
     label: columnName,
     value: columnValue !== undefined ? columnValue : null, // Handle missing fields
     type:columnTpye,
     required:columnRequired
   };
 });

  console.log("resultArrayThatContainstheColumnDetails",resultArrayThatContainstheColumnDetails)

  // Create the main container
  const mainContainer = document.createElement('div');
  mainContainer.className = 'main-containeruploadfile';
  const librarydiv= document.getElementById('filelistcontainer')
  const backButton = document.createElement('button')
  backButton.textContent = 'Close File Preview';
  backButton.className = 'btn btn-secondary me-2 mt-2'; 
  
  backButton.addEventListener('click', () => {
    // Since you used innerHTML = "" to wipe the container, 
    // a reload is the cleanest way to restore the React state and the file list.
    window.location.reload();
    
    // ALTERNATIVE: If you don't want a full reload, you can try calling your view handler:
    // handleViewButtonClick(activeView || "My request");
  });
  const submitButton=document.createElement('button');
  const replaceButton=document.createElement('button');

  const uploadFileDiv=document.createElement('div');
  uploadFileDiv.id="uploadFileDiv";
  uploadFileDiv.style.display='none'
   // input for upload file
   const uploadFileInput=document.createElement('input');
   uploadFileInput.className="dynamic-input";
   uploadFileInput.type="file";
   uploadFileInput.id="fileInput";

  //  start
   const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const replaceButtonHide=document.getElementById('replaceButton');
    if(replaceButtonHide){
      replaceButtonHide.style.display='none'
    }
    const file = event.target.files![0];
    if (file) {
      // selectedFile=file;
      uploadFile(file);
    }
  };

  // const uploadFile = async (file: File) => {
  //   try {
  //     const folder = sp.web.getFolderByServerRelativePath('DMSOrphanDocs');
  //     const uploadResult = await folder.files.addChunked(file.name, file);
  //     console.log("File uploaded successfully", uploadResult);

  //     // Generate the preview URL dynamically
  //     const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
      
  //     previewFile(previewUrl);
  //   } catch (error) {
  //     console.error("Error uploading file:", error);
  //   }
  // };


  const uploadFile = async (file: File) => {
    try {
      // 1. USE masterWeb (the instance targeting your Master Site Collection)
      // 2. USE the full server-relative path (e.g., /sites/AlRostmaniSpfx2/DMSOrphanDocs)
      const folderPath = `${locationPath}/DMSOrphanDocs`.replace(/\/\/+/g, '/');
      
      console.log("Uploading to folder:", folderPath);
      
      const folder = masterWeb.getFolderByServerRelativePath(folderPath);
      const uploadResult = await folder.files.addChunked(file.name, file);
      
      console.log("File uploaded successfully to Master Site", uploadResult);

      // Generate the preview URL dynamically
      const previewUrl = await generatePreviewUrl(uploadResult.data.ServerRelativeUrl);
      
      previewFile(previewUrl);
    } catch (error) {
      console.error("Error uploading file to cross-site location:", error);
      // If 'masterWeb' fails, it might be because 'DMSOrphanDocs' 
      // doesn't exist on the Master site. Check if it's on the subsite instead:
      // const folder = web.getFolderByServerRelativePath(...);
    }
  };
  const generatePreviewUrl = async (serverRelativeUrl: string) => {
    // Encode the file name and construct the preview URL
    const encodedFilePath = encodeURIComponent(serverRelativeUrl);
    
    // Example: 
    // serverRelativeUrl = "/sites/AlRostmani/test/DocumentLibraryInsideTest/Book.xlsx"
    const parentFolder = serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'));
    const siteUrl = window.location.origin;

    // const previewUrl = `${siteUrl}/sites/AlRostmani/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    const previewUrl = `${siteUrl}${locationPath}/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
    // const previewUrl = `${siteUrl}/sites/SPFXDemo/DMSOrphanDocs/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
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

  const previewFile = async (previewUrl: string) => {
    try {
      console.log("Previewing file at URL:", previewUrl);
      const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
      const spinner = document.getElementById("spinner") as HTMLElement;
   
      const librarydiv = document.getElementById('filelistcontainer');
      const ribbonDiv = document.querySelector('.col-lg-10.newbutton.tool') as HTMLElement;
     
       // 🔹 REMOVE EXISTING CLOSE BUTTON IF IT EXISTS
  const existingCloseButton = document.getElementById('closePreviewRibbon');
  if (existingCloseButton && ribbonDiv && ribbonDiv.contains(existingCloseButton)) {
    ribbonDiv.removeChild(existingCloseButton);
  }
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
  
                // Hide the spinner and show the iframe after the button is hidden
                spinner.style.display = "none";
                iframe.style.display = "block"; 

               // Exit the loop once the button is found and hidden
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
   uploadFileInput.addEventListener('change', (event:any) =>
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
    asterisk.style.color="red";
    asterisk.style.fontWeight="bold";
    label.appendChild(asterisk);
}
inputContainer.appendChild(label);

  // let modifiedType = field.type.replace(/\s+/g, '').toLowerCase();

  let modifiedType = (field.type || "").replace(/\s+/g, '').toLowerCase();

  // Create the input field based on its type
  let input: HTMLInputElement | null = null;
  if (
    modifiedType === "singlelineoftext"
    || 
    modifiedType === "multiplelineoftext" 
    || 
    modifiedType === 'text'
){
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
  input.className="dynamic-input";
  input.value = field.value;
  input.id = field.label;
  input.name = field.label;
  input.required=field.required;
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

  

  replaceButton.type="submit";
  replaceButton.id="replaceButton";
  replaceButton.addEventListener('click',(event)=>{
    event.preventDefault();
    console.log("replace button called");
    clickedReplace=true;
    const uploadFile=document.getElementById('uploadFileDiv')
    const iframe = document.getElementById('filePreview');
    if(uploadFile){
      uploadFile.style.display='block';
    }
    if(iframe){
      iframe.style.display='none';
    }
    // myRequest(null,null,null);
   })
   replaceButton.textContent="Replace"
  

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
  
  // const previewUrl = `${siteUrl}/sites/AlRostmani/${currentSubsite}/${myactualdoclib}/Forms/AllItems.aspx?id=${filePath}&parent=${encodedParentFolder}`;
  const previewUrl = `${siteUrl}${locationPath}/${currentSubsite}/${myactualdoclib}/Forms/AllItems.aspx?id=${filePath}&parent=${encodedParentFolder}`;

  if(previewUrl){
    previewfileframe.src = previewUrl;
    column2.appendChild(previewfileframe);
  }

  // Append columns to the main container
  mainContainer.appendChild(column1);
  mainContainer.appendChild(column2);


   // Submit Button property
   submitButton.type="submit";
  //  submitButton.addEventListener('click',async(event)=>{
  //   event.preventDefault();
  //   console.log("submit button called");

  //   // Extract the last part after the last '/'
  //   const fileName:any = filePath.substring(filePath.lastIndexOf('/') + 1);
  //   alert(fileName);
  //   // console.log("fileName",fileName);
  //   // Extract the rest of the path
  //   const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
  //   // console.log("folderPath",folderPath);

  //   const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  //   if (!formSelector.checkValidity()) {
  //       checkValidation(`Fill mandatory fields`)
  //       return;
  //   }

  // // Prepare the payload for SharePoint dynamically
  // const inputs = document.querySelectorAll('.dynamic-input');
  // const payload: any = {};

  // inputs.forEach((input) => {
  //     const inputElement = input as HTMLInputElement;
  //     const fieldName = inputElement.id;
  //     if (!fieldName) return; // Skip if field name is invalid

  //     if (inputElement.type === "checkbox") {
  //         // console.log("fieldName",fieldName.includes(' '));
  //         payload[fieldName] = inputElement.checked;
  //     } else if (inputElement.type !== "file") {
  //         if(inputElement.value === ""){
  //            console.log("skip");
  //         }else{
  //           // if(fieldName.includes(' '))
  //           // console.log("fieldName",fieldName.includes(' '));
  //           payload[fieldName] = inputElement.value;
  //         }
          
  //     }
  //   });
   
  //   if(clickedReplace){
  //     const fileInput = document.getElementById('fileInput') as HTMLInputElement;
  //     const selectedFile = fileInput?.files?.[0]; 

  //     if (!selectedFile) {
  //         console.error("No file selected.");
  //         // alert("Please select the file...");
  //         checkValidation(`Fill mandatory fields`)
  //         return;
  //     }

  //     const documentLibraryInWhichWeUploadTheFile = web.getFolderByServerRelativePath(folderPath);
  //     // const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile,true);
  //     const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.update(
  //       fileName,
  //       selectedFile,
  //       null,
  //       true
  //     );
  
  //     const listItem = await uploadResult.file.getItem();
  //     const result =await listItem.update(payload);
  //     console.log("fileupdated ",result);
  //   }else{

  //   }
  //   const file =sp.web.getFileByServerRelativePath(filePath);
   
    
  //   // myRequest(null,null,null);
  //  })
  // submitButton.addEventListener('click', async (event) => {
  //   event.preventDefault();
  //   console.log("submit button called");
   
  
  //   const fileName = filePath.substring(filePath.lastIndexOf('/') + 1);
  //   console.log("fileName", fileName);
  
   
  //   const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
  //   console.log("folderPath", folderPath);
   
  //   const formSelector = document.getElementById("formSelector") as HTMLFormElement;
  //   if (!formSelector.checkValidity()) {
  //     checkValidation('Fill mandatory fields');
  //     return;
  //   }
   
  //   // Prepare the payload for SharePoint dynamically
  //   const inputs = document.querySelectorAll('.dynamic-input');
  //   const payload:any = {};
   
  //   inputs.forEach((input) => {
  //     const inputElement = input as HTMLInputElement;
  //     const fieldName:any = inputElement.id;
  //     if (!fieldName) return;
   
  //     if (inputElement.type === "checkbox") {
  //       payload[fieldName] = inputElement.checked;
  //     } else if (inputElement.type !== "file") {
  //       if (inputElement.value !== "") {
  //         payload[fieldName] = inputElement.value;
  //       }
  //     }
  //   });
   
  //   let selectedFile;
  //   if (clickedReplace) {
  //     const fileInput= document.getElementById('fileInput') as HTMLInputElement;
  //     selectedFile = fileInput?.files?.[0];
   
  //     if (!selectedFile) {
  //       console.error("No file selected.");
  //       checkValidation('Fill mandatory fields');
  //       return;
  //     }
   
  //     try {
  //       const documentLibraryInWhichWeUploadTheFile = sp.web.getFolderByServerRelativePath(folderPath);
  
  //       const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(
  //         fileName,
  //         selectedFile,
  //         null,
  //         true 
  //       );
   
   
  //       const listItem = await uploadResult.file.getItem();
  //       const result = await listItem.update(payload);
  //       console.log("file updated", result);
  //     } catch (error) {
  //       console.error("Error replacing file:", error);
  //     }
  //   }
  // });

  submitButton.addEventListener('click',async(event)=>{
      event.preventDefault();
      console.log("submit button called");
  
      // Extract the last part after the last '/'
      const fileName:any = filePath.substring(filePath.lastIndexOf('/') + 1);
      // alert(fileName);
      // console.log("fileName",fileName);
      // Extract the rest of the path
      const folderPath = filePath.substring(0, filePath.lastIndexOf('/'));
      console.log("folderPath",folderPath);
  
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
            if(inputElement.value === ""){
               console.log("skip");
            }else{
              // if(fieldName.includes(' '))
              // console.log("fieldName",fieldName.includes(' '));
              payload[fieldName] = inputElement.value;
            }
            
        }
      });
     
    if(clickedReplace){
        const fileInput = document.getElementById('fileInput') as HTMLInputElement;
        const selectedFile = fileInput?.files?.[0]; 
  
        if (!selectedFile) {
            console.error("No file selected.");
            // alert("Please select the file...");
            checkValidation(`Fill mandatory fields`)
            return;
        }
  
        // const documentLibraryInWhichWeUploadTheFile = web.getFolderByServerRelativePath(folderPath);
        // // const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile,true);
        // const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(
        //   fileName,
        //   selectedFile,
        //   null,
        //   true
        // );
    
        // const listItem = await uploadResult.file.getItem();
        // const result =await listItem.update(payload);
        // console.log("fileupdated ",result);
         
        // Get the target file
        // const file =sp.web.getFileByServerRelativePath(filePath);
 
        // // Check out the file
        // await file.checkout();
 
        // // Upload the new file
        // const uploadedFile = await file.getParentFolder().files.add(
        //   `${selectedFile.name.replace(/\.[^.]+$/, '.docx')}`,
        //   selectedFile
        // );
 
        // // Update the properties of the new file
        // await uploadedFile.item.update(payload);
 
        // // Check in the file
        // await uploadedFile.checkin();
 
        // console.log('File replaced successfully!');
        // Get the target file
        console.log("selectedFile.name",selectedFile.name) 
        const fileExtensionOfSelectedFile = selectedFile.name.split('.').pop();
        const fileExtensionOfOldFile =fileName.split('.').pop();
        // for same file extension
        // if(fileExtensionOfSelectedFile === fileExtensionOfOldFile){
        //     // alert("Same file extension");
        //     const file = web.getFileByServerRelativePath(filePath);
        //       await file.setContentChunked(selectedFile);
        //       if (file.exists) {
        //         const fileToUpdate = await file.getItem();
        //         const uploadResult = await fileToUpdate.update(payload);
        //         console.log("uploadResult",uploadResult);
        //       }
        //     showReplaceMessage('File replaced successfully.');
        //     myRequest(null,null,null);
        // }else{
          // alert("file extension are not same");
          const folderInWhichWeUploadTheFile=web.getFolderByServerRelativePath(folderPath);
          const uploadResult = await folderInWhichWeUploadTheFile.files.addChunked(selectedFile.name, selectedFile);
          const listItem = await uploadResult.file.getItem();
          (payload as any).Status="Pending";

          const parentFolder = uploadResult.data.ServerRelativeUrl.substring(0, uploadResult.data.ServerRelativeUrl.lastIndexOf('/'));
          const siteUrl = window.location.origin;
          const encodedFilePath = encodeURIComponent(uploadResult.data.ServerRelativeUrl);
          // const previewUrl = `${siteUrl}/sites/AlRostmani/${siteName}/${documentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;
          const previewUrl = `${siteUrl}${locationPath}/${siteName}/${documentLibrary}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodeURIComponent(parentFolder)}`;

          await listItem.update(payload);
          // Get the file details
          const fetchData=await masterWeb.lists.getByTitle(`DMS${siteName}FileMaster`).items.select("ID","FileName","RequestNo").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and CurrentUser eq '${currentUserEmailRef.current}'  and FileUID eq '${fileId}'`)();
          console.log("fetchData",fetchData);

          const newItem = await masterWeb.lists.getByTitle(`DMS${siteName}FileMaster`).items.add({
            FileName: String(uploadResult.data.Name),
            FileSize: String(uploadResult.data.Length),
            FileVersion: String(uploadResult.data.MajorVersion),
            CurrentFolderPath: String(folderPath),
            FileUID: String(uploadResult.data.UniqueId),
            CurrentUser: String(currentUserEmailRef.current),
            SiteID: String(siteId),
            Status: "Pending",
            FilePreviewURL : String(previewUrl),
            DocumentLibraryName:String(documentLibrary),
            SiteName : String(siteName),
            MyRequest:true,
            RequestNo:String(fetchData[0].RequestNo)
        });
        

      //   const AddIteminDMSFileApprovalList = await sp.web.lists.getByTitle('DMSFileApprovalList').items.add({
      //     SiteName : String(siteName),  
      //      DocumentLibraryName : String(documentLibrary),
      //      RequestedBy  : String(currentUserEmailRef.current),
      //      FileName: String(uploadResult.data.Name),
      //      FileUID: String(uploadResult.data.UniqueId),
      //     //  FilePreviewUrl: String(previewUrl),
      //      Status: String('Pending'),
      //      FolderPath : String(folderPath),
      //      ApproveAction : String('Submitted'),
      //      ApprovedLevel : 1
      // })

      // delete the file from the document library
      const deletedfile =  await web.getFileById(fileId).delete();
      console.log("deletedfile",deletedfile);
      
      await masterWeb.lists.getByTitle(`DMS${siteName}FileMaster`).items.getById(fetchData[0].ID).delete();
      console.log(`file has been deleted successfully.`);

      const fetchDatafromapprovalist=await masterWeb.lists.getByTitle(`DMSFileApprovalList`).items.select("ID","FileName").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and FileUID eq '${fileId}'`)();
      
      await masterWeb.lists.getByTitle('DMSFileApprovalList').items.getById(fetchDatafromapprovalist[0].ID).update({
        FileName:String(uploadResult.data.Name),
        FileUID:String(uploadResult.data.UniqueId),
        Status: String('Pending'),
        FilePreviewUrl:String(previewUrl),
        ApproveAction:String('Submitted'),
        CurrentLevel:1
        
      });
      // await masterWeb.lists.getByTitle(`DMSFileApprovalList`).items.getById(fetchDatafromapprovalist[0].ID).delete();
      console.log(`file has been deleted updated fetchDatafromapprovalist.`);

      // Fetch data from DMSSharewithother 
      const fetchDataFromDMSShareWithOtherMaster=await masterWeb.lists.getByTitle(`DMSShareWithOtherMaster`).items.select("ID","FileName").filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${documentLibrary}' and FileUID eq '${fileId}'`)();

      // delete the record if that file is share 
      if(fetchDataFromDMSShareWithOtherMaster.length > 0){
        fetchDataFromDMSShareWithOtherMaster.forEach(async(shareRecord)=>{
          try {
            await masterWeb.lists.getByTitle(`DMSShareWithOtherMaster`).items.getById(shareRecord.ID).delete();
            console.log(`file has been deleted successfully.`);
  
          } catch (error) {
            console.log("Error in deleting the share record from DMSShareWithOthersMaster",error);
          }
        })
      }
      
      showReplaceMessage('File replaced successfully.');
    // }
        
          // console.log('File replaced successfully!');
          
          // if (file.exists) {
          //   const fileToUpdate = await file.getItem();
          //   const uploadResult = await fileToUpdate.update({
          //     // FileLeafRef: fileToUpdate.,
          //     File: selectedFile,
          //   });
          // } else {

        // const fileInfo = await file.select("ServerRelativeUrl")();
        // const parentFolderUrl = fileInfo.ServerRelativeUrl.substring(0, fileInfo.ServerRelativeUrl.lastIndexOf("/"));
        // console.log("Parent Folder URL:", parentFolderUrl)y

        // const parentFolder = web.getFolderByServerRelativePath(parentFolderUrl);

        //   // Upload the new file
        //   const uploadedFile = await parentFolder.files.addUsingPath(
        //     `${parentFolderUrl}/${fileName.replace(/\.[^.]+$/, '.docx')}`,
        //     selectedFile,
        //     { Overwrite: true }
        //   );

        //   // Update the properties of the new file
        //   const uploadedFileItem = await uploadedFile.file.getItem();
        //   await uploadedFileItem.update(payload);


         
        //   // const uploadResult = await documentLibraryInWhichWeUploadTheFile.files.addChunked(
        //   //   fileName,
        //   //   selectedFile,
        //   //   null,
        //   //   true
        //   // );
        // }

        // Check out the file
        // await file.checkout();
        // Check if the file path is valid
        // try {
        //   await file.checkout();
        // } catch (error) {
        //   console.error("Error checking out the file:", error);
        //   throw new Error(`Invalid file path: ${filePath}`);
        // }

        // Get the parent folder's URL
        // const fileInfo = await file.select('ServerRelativeUrl')();
        // const parentFolderUrl = fileInfo.ServerRelativeUrl.substring(0, fileInfo.ServerRelativeUrl.lastIndexOf('/'));
        // Get folder URL
        // const fileInfo = await file.select("ServerRelativeUrl")();
        // const parentFolderUrl = fileInfo.ServerRelativeUrl.substring(0, fileInfo.ServerRelativeUrl.lastIndexOf("/"));
        // console.log("Parent Folder URL:", parentFolderUrl)

       
        // Get the parent folder
        // const parentFolder = web.getFolderByServerRelativePath(parentFolderUrl);
        // Convert the file content to ArrayBuffer
        // const fileContent = await selectedFile.arrayBuffer();

        // Upload the new file
        // const uploadedFile = await parentFolder.files.addUsingPath(
        //   `${parentFolderUrl}/${fileName.replace(/\.[^.]+$/, '.docx')}`,
        //   selectedFile,
        //   { Overwrite: true }
        // );

        // Update the properties of the new file
        // const uploadedFileItem = await uploadedFile.file.getItem();
        // await uploadedFileItem.update(payload);

        // Check in the file
        // await uploadedFile.file.checkin('Checked in by script', 1); // 1 for Major version
 
    }else{
      // Without replacing the file only metadeta update
      const file = web.getFileByServerRelativePath(filePath);
      if (file.exists) {
        const fileToUpdate = await file.getItem();
        const uploadResult = await fileToUpdate.update(payload);
        console.log("uploadResult",uploadResult);
      }
      showReplaceMessage('File updated successfully.')
    }
    // const file =sp.web.getFileByServerRelativePath(filePath);
   
   
    // myRequest(null,null,null);
   })


 submitButton.textContent="Submit"
 form.appendChild(submitButton);

librarydiv.innerHTML = "";
mainContainer.appendChild(backButton)
librarydiv.appendChild(mainContainer)



}
  // sourish 20/8/25
  // const deleteFolder = async (file: any) => {
  //   try {
  //     if (!file?.FolderPath || !file?.__siteUrl) {
  //       console.error("Missing folder path or site URL");
  //       return;
  //     }

  //     // Delete Folder

  //     // Build correct web URL for the subsite
  //     const fullWebUrl = `${file.__siteUrl}/${encodeURIComponent(file.SiteTitle)}`;
  //     const sp = spfi(fullWebUrl).using(SPFx(context));

  //     // FolderPath is already server-relative, so just use it directly
  //     const finalPath = file.FolderPath;

  //     console.log("Deleting from:", fullWebUrl);
  //     console.log("Final server-relative path:", finalPath);

  //     await sp.web.getFolderByServerRelativePath(finalPath).delete();


  //     // Delete corresponding item from list in site collection root
  //     const spRoot = spfi(file.__siteUrl).using(SPFx(context));

  //     console.log("Deleting list item from site:", file.__siteUrl, " List: DMSFolderMaster, ID:", file.ID);

  //     await spRoot.web.lists.getByTitle("DMSFolderMaster").items.getById(file.ID).delete();

  //     const refreshed = await loadViewData("MyFolders");
  //     setSelectedFiles(refreshed);
  //     setActiveView("My Folders");
  //     setBreadcrumbs([
  //       { key: "my-folders", title: "My Folders", type: "view", siteUrl: "" },
  //     ]);
  //     setCurrentPage(1);


  //   } catch (err) {
  //     console.error("Error deleting folder:", err);
  //   }
  // };

  // Addhyan 16/4/26
  const deleteFolder = async (file: any) => {
  try {
    if (!file?.FolderPath || !file?.__siteUrl) {
      console.error("Missing folder path or site URL");
      return;
    }
    const result = await Swal.fire({
      title: "Are you sure?",
      text: "want to delete this folder?",
      icon: "warning",
      showCancelButton: true,
      confirmButtonText: "Yes, delete it!",
      cancelButtonText: "Cancel",
    });
 
    if (!result.isConfirmed) return;
 
    const fullWebUrl = `${file.__siteUrl}/${encodeURIComponent(file.SiteTitle)}`;
    const sp = spfi(fullWebUrl).using(SPFx(context));
    const finalPath = file.FolderPath;
    console.log("Deleting from:", fullWebUrl);
    console.log("Final server-relative path:", finalPath);
    const spRoot = spfi(file.__siteUrl).using(SPFx(context));
    const pathParts = file.FolderPath.split("/");
    const libraryName = pathParts[3]; // Migarationtest
    const dynamicListName = `DMS${libraryName}FileMaster`;
    console.log("Dynamic List:", dynamicListName);
    const items = await spRoot.web.lists
  .getByTitle(dynamicListName)
  .items
  .filter(`startswith(CurrentFolderPath,'${file.FolderPath}')`)
  .select("Id")();   
    console.log("Metadata items found:", items.length);
    for (let item of items) {
      await spRoot.web.lists
        .getByTitle(dynamicListName)
        .items.getById(item.Id)
        .delete();
    }
    await sp.web.getFolderByServerRelativePath(finalPath).delete();
    console.log("Deleting list item from site:", file.__siteUrl, " List: DMSFolderMaster, ID:", file.ID);
    await spRoot.web.lists
      .getByTitle("DMSFolderMaster")
      .items.getById(file.ID)
      .delete();
 
      Swal.fire({
      title: "Deleted!",
      text: "Folder successfully deleted.",
      icon: "success",
    });
 
    const refreshed = await loadViewData("MyFolders");
    setSelectedFiles(refreshed);
    setActiveView("My Folders");
    setBreadcrumbs([
      { key: "my-folders", title: "My Folders", type: "view", siteUrl: "" },
    ]);
    setCurrentPage(1);
 
  } catch (err) {
    console.error("Error deleting folder:", err);
    Swal.fire({
      title: "Error!",
      text: "Failed to delete folder.",
      icon: "error",
    });
  }
};

  // sourish 21/8/25
  const renameFolder = async () => {
    try {
      if (!modalFile?.FolderPath || !modalFile?.__siteUrl) {
        console.error("Missing folder path or site URL");
        return;
      }

      const fullWebUrl = `${modalFile.__siteUrl}/${encodeURIComponent(modalFile.SiteTitle)}`;
      const sp = spfi(fullWebUrl).using(SPFx(context));

      const folder = sp.web.getFolderByServerRelativePath(modalFile.FolderPath);

      // Build new path with the renamed folder name
      const parentPath = modalFile.FolderPath.substring(0, modalFile.FolderPath.lastIndexOf("/"));
      const newPath = `${parentPath}/${renameValue}`;

      console.log("Renaming folder:", modalFile.FolderPath, " → ", newPath);

      // ✅ Use moveByPath instead of moveTo
      await folder.moveByPath(newPath);

      // Update item in DMSFolderMaster list
      const spRoot = spfi(modalFile.__siteUrl).using(SPFx(context));
      await spRoot.web.lists.getByTitle("DMSFolderMaster").items.getById(modalFile.ID).update({
        FolderName: renameValue,
        FolderPath: newPath,
      });

      // Refresh view
      const refreshed = await loadViewData("MyFolders");
      setSelectedFiles(refreshed);
      setActiveView("My Folders");
      setBreadcrumbs([{ key: "my-folders", title: "My Folders", type: "view", siteUrl: "" }]);
      setCurrentPage(1);

      // Close modal
      setRenameModalOpen(false);
      setModalFile(null);
      setRenameValue("");

    } catch (err) {
      console.error("Error renaming folder:", err);
    }
  };

  // aman code manage folder permission
  // === Manage Permission helpers ===
  const deriveFolderContext = (folder: any) => {
    const rootSite = folder?.__siteUrl || context.pageContext.web.absoluteUrl;
    const siteTitle = folder?.SiteTitle || folder?.SiteName || "";
    const webUrl = siteTitle ? `${rootSite}/${encodeURIComponent(siteTitle)}` : rootSite;
    const serverRel =
      folder?.ServerRelativeUrl ||
      folder?.FolderPath ||
      folder?.folderpath ||
      "";
    const deriveLibrary = (sr: string): string => {
      const parts = (sr || "").split("/").filter(Boolean);
      const idxSites = parts.indexOf("sites");
      if (idxSites >= 0 && parts.length > idxSites + 2) return parts[idxSites + 2];
      return parts[2] || "";
    };
    const documentLibraryName =
      folder?.DocumentLibraryName ||
      folder?.DocumentLibrary ||
      folder?.LibraryName ||
      deriveLibrary(serverRel);
    return { webUrl, serverRel, documentLibraryName };
  };

  const getFolderApi = (webUrl: string, serverRel: string) => {
    const w: any = spfi(webUrl).using(SPFx(context)).web;
    if ((w as any).getFolderByServerRelativeUrl) {
      return (w as any).getFolderByServerRelativeUrl(serverRel);
    }
    return w.getFolderByServerRelativePath(serverRel);
  };

  const loadAndOpenManagePermission = async () => {
    if (!selectedFolder) return;
    setMpLoading(true);
    setMpError("");
    try {
      const { webUrl, serverRel } = deriveFolderContext(selectedFolder);
      const web = spfi(webUrl).using(SPFx(context));
      const folderApi = getFolderApi(webUrl, serverRel);
      const item = await folderApi.getItem();
      const itemInfo: any = await item.select("HasUniqueRoleAssignments")();
      setMpHasUnique(!!itemInfo?.HasUniqueRoleAssignments);
      let can = false;
      try {
        can = await web.web.currentUserHasPermissions(PermissionKind.ManagePermissions);
      } catch { }
      setMpCanManage(!!can);
      let ras: any[] = [];
      try {
        ras = await item.roleAssignments.expand("Member", "RoleDefinitionBindings")();
      } catch (e) {
        console.warn("[ManagePermission] roleAssignments read failed:", e);
        ras = [];
      }
      const mapped = (ras || []).map((ra: any) => ({
        principalId: ra?.Member?.Id,
        principalTitle: ra?.Member?.Title,
        roles: (ra?.RoleDefinitionBindings || []).map((r: any) => r?.Name).filter(Boolean),
      }));
      setPermissionUsers(mapped);

      // People picker users
      try {
        const rawUsers: ISiteUserInfo[] = await web.web.siteUsers();
        const mappedUsers = (rawUsers || [])
          .map((u) => ({
            id: u.Id || 0,
            title: u.Title || u.LoginName || "",
            email: u.Email || "",
            loginName: u.LoginName || u.Email || "",
          }))
          .filter((u) => !!u.loginName);
        setMpSiteUsers(mappedUsers);
      } catch (e) {
        console.warn("[ManagePermission] siteUsers fetch failed:", e);
        setMpSiteUsers([]);
      }

      setShowManagePermissionModal(true);
    } catch (e: any) {
      setMpError(e?.message || String(e));
    } finally {
      setMpLoading(false);
    }
  };

  const addUserToSelectedFolder = async () => {
    if (!selectedFolder || !newUser || !newPermission) return;
    setMpLoading(true);
    setMpError("");
    try {
      const { webUrl, serverRel } = deriveFolderContext(selectedFolder);
      const web = spfi(webUrl).using(SPFx(context));
      const folderApi = getFolderApi(webUrl, serverRel);
      const item = await folderApi.getItem();

      if (!mpHasUnique) {
        try {
          await item.breakRoleInheritance(true, false);
          setMpHasUnique(true);
        } catch (e) {
          console.warn("[ManagePermission] breakRoleInheritance failed:", e);
        }
      }

      let userId: number | null = null;
      try {
        const ensured = await web.web.ensureUser(newUser.trim());
        userId = ensured?.data?.Id;
      } catch (e) {
        console.error("[ManagePermission] ensureUser failed:", e);
        throw e;
      }
      if (!userId) throw new Error("Could not resolve user/group.");

      const roleDef = await web.web.roleDefinitions.getByName(newPermission)();
      await item.roleAssignments.add(userId, roleDef.Id);

      const picked = mpSiteUsers.find((u) => u.loginName === newUser.trim());
      const displayTitle = picked?.title || newUser.trim();

      setPermissionUsers((prev) => {
        const idx = prev.findIndex((p) => p.principalId === userId);
        if (idx >= 0) {
          const setRoles = new Set([...(prev[idx].roles || []), newPermission]);
          const updated = [...prev];
          updated[idx] = { ...prev[idx], roles: Array.from(setRoles) };
          return updated;
        }
        return [...prev, { principalId: userId, principalTitle: displayTitle, roles: [newPermission] }];
      });

      setNewUser("");
      setNewPermission("Read");
    } catch (e: any) {
      setMpError(e?.message || String(e));
    } finally {
      setMpLoading(false);
    }
  };

  const removeUserFromSelectedFolder = async (principalId: number) => {
    if (!selectedFolder || !principalId) return;
    setMpLoading(true);
    setMpError("");
    try {
      const { webUrl, serverRel } = deriveFolderContext(selectedFolder);
      const item = await getFolderApi(webUrl, serverRel).getItem();
      await item.roleAssignments.getById(principalId).delete();
      setPermissionUsers((prev) => prev.filter((p) => p.principalId !== principalId));
    } catch (e: any) {
      setMpError(e?.message || String(e));
    } finally {
      setMpLoading(false);
    }
  };

  //abhay delete file from folder
  // const deleteFileFolder = async (file: any, siteUrl: string, context: any) => {
  // try {
  //   const siteSP = spfi(siteUrl).using(SPFx(context));

  //   if (!file?.ServerRelativeUrl) {
  //     console.error("[deleteFile] No ServerRelativeUrl found on file:", file);
  //     return;
  //   }

  //   console.log("[deleteFile] Deleting:", file.ServerRelativeUrl);

  //   await siteSP.web
  //     .getFileByServerRelativePath(file.ServerRelativeUrl)
  //     .delete();
  //     // UI update just after deletion of files from Document libraries fix by aman 15/01/26
  //      if (selectedCurrentNode) {
  //     await loadFilesForNode(selectedCurrentNode);
  //     }


  //   console.log("[deleteFile] File deleted successfully:", file.Name);
  //   alert("File Deleted");

  //   // Optionally, refresh your list after delete
  //   // await loadFilesForNode(currentNode);
  // } catch (err) {
  //   console.error("[deleteFile] Error deleting file:", err);
  // } 
  // }
useEffect(() => {
  getdocumentcategory();
}, []);

//addhyan 24/2/26
 const ArchivedFile = async (file: any, siteUrl: string, context: any) => {
    try {
      Swal.fire({
        title: 'Archiving File',
        html: 'Please wait while we archive this file...',
        allowOutsideClick: false,
        didOpen: async () => {
          Swal.showLoading();
          try {
            console.log("[ArchivedFile] file:", file);
            const isoDate = new Date().toISOString();
            const targetLibraryName = "ArchivedDocuments";
            const listName  = "EssaArchivallist"
 
 
            const siteSP = spfi(siteUrl).using(SPFx(context));
            const fileItem = await siteSP.web
              .getFileByServerRelativePath(file.ServerRelativeUrl)
              .getItem();
 
            // Get root web and destination path
            const rootWeb = await sp.site.getRootWeb();
            const rootSiteData = await rootWeb.select("ServerRelativeUrl")();
            const rootSiteRelativeUrl = rootSiteData.ServerRelativeUrl;
           
            const destinationFolderPath = `${rootSiteRelativeUrl.replace(/\/$/, "")}/${targetLibraryName}`;
           
            const filedetails = file.ServerRelativeUrl.split("/");
            const siteCollection = filedetails[2];
const subSite = filedetails[3];
const documentLibrary = filedetails[4];
 
const folderPath = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf("/") + 1);
 
const segments = file.ServerRelativeUrl.split("/").filter(Boolean);
const folderName = segments[segments.length - 2];
console.log("folderName:", folderName);
 
 
 
console.log("siteCollection:", siteCollection);
console.log("subSite:", subSite);
console.log("documentLibrary:", documentLibrary);
console.log("folderPath:", folderPath);
 
const finalFolderName =
  folderName !== documentLibrary ? folderName : "";
 
            const filedestination  = `${rootSiteRelativeUrl.replace(/\/$/, "")}/${listName}`;
             await sp.web.lists.getByTitle("EssaArchivallist").items.add({
              filename: file.Name,
              fileuid: file.UniqueId,
              sitenamae : subSite,
              sitecollectionname : siteCollection,
              documentlibraryName: documentLibrary,
              folderPath: folderPath,
              folderNames: finalFolderName,
             
 
 
       
      });
 
           
 
           
 
           
 
 
           
            console.log("Archiving file:", file.Name);
            console.log("From path:", file.ServerRelativeUrl);
            console.log("To path:", destinationFolderPath);
            console.log("List url :", filedestination);
 
            // Create destination file path with original filename
            const destinationFilePath = `${destinationFolderPath}/${file.Name}`;
 
            // Copy the file to ArchivedDocuments library
            const sourceFile = siteSP.web.getFileByServerRelativePath(file.ServerRelativeUrl);
            await sourceFile.copyByPath(destinationFilePath, true, false);
 
            // Update the file item metadata to mark as deleted
            await fileItem.update({
              IsDeleted: isoDate,
            });
 
            console.log("File successfully archived:", file.Name);
 
 
             const destFile = await rootWeb.getFileByServerRelativePath(destinationFilePath);
        const destItem = await destFile.getItem();
   
    await destItem.update({
      Status: "Complete"
    });
 
            // Remove from selected files
            setSelectedFiles((prev: any[]) =>
              prev.filter(
                (f) => f.ServerRelativeUrl !== file.ServerRelativeUrl
              )
            );
 
            Swal.fire({
              icon: 'success',
              title: 'Success!',
              text: `File '${file.Name}' has been archived successfully.`,
              timer: 2000
            });
 
          } catch (err) {
            console.error("Archive failed:", err);
            Swal.fire({
              icon: 'error',
              title: 'Error',
              text: `Failed to archive file: ${err instanceof Error ? err.message : 'Unknown error'}`
            });
          }
        }
      });
    } catch (err) {
      console.error("Archive function error:", err);
      Swal.fire({
        icon: 'error',
        title: 'Error',
        text: `An error occurred: ${err instanceof Error ? err.message : 'Unknown error'}`
      });
    }
  };
 
  const getdocumentcategory = async ()=>{

  const arr:any = [];
  const navItem = await sp.web.lists.getByTitle("TemplateDocumentCategory").items.select("DocumentCategory").getAll();

  // alert("navItem" + JSON.stringify(navItem))

  console.log("Addhyan navItem", navItem);
    setNavItems(navItem);

  // setNavItems(navItem)
}

// Ritik 20/2/26 for manage permission folder
// ===== ManageFolderPermission Inline Logic =====
const fpMapRoleAssignments = (roleAssignments: any[]) => {
  return (roleAssignments || [])
    .filter((ra: any) => {
      const member = ra?.Member;
      if (!member) return false;
      if (member.IsHiddenInUI) return false;
      const title = (member.Title || "").toLowerCase();
      if (title.includes("system account")) return false;
      if (title.includes("everyone")) return false;
      if (title.includes("all users")) return false;
      if (title.includes("limited access")) return false;
      return true;
    })
    .flatMap((ra: any) => {
      // ✅ Har permission ke liye alag row
      const permissions = (ra.RoleDefinitionBindings || [])
        .map((r: any) => r.Name)
        .filter((name: string) => name !== "Limited Access"); // ✅ Sirf Limited Access hatao, baaki sab dikho
 
      return permissions.map((permName: string) => ({
        itemId: `${ra.Member.Id}_${permName}`, // unique key per permission
        userId: ra.Member.Id,
        value: ra.Member.Title || ra.Member.LoginName || "",
        label: ra.Member.Title || "",
        Permission: permName, // ✅ Sirf ek permission per row
      }));
    })
    .filter((item: any) => item.Permission !== "");
};
 
// ✅ SHARED FETCH ROLE ASSIGNMENTS - teeno jagah same logic
const fpFetchRoleAssignments = async (spSubsite: any, folderName: string, folderPath: string, docLibName: string) => {
  try {
    if (folderName !== "null") {
      const folderItem = await spSubsite.web
        .getFolderByServerRelativePath(folderPath)
        .getItem();
      return await folderItem.roleAssignments
        .expand("Member", "RoleDefinitionBindings")();
    } else {
      const lib = spSubsite.web.lists.getByTitle(docLibName);
      return await lib.roleAssignments
        .expand("Member", "RoleDefinitionBindings")();
    }
  } catch (e) {
    console.warn("Role assignments fetch failed", e);
    return [];
  }
};
 
const fpLoadDefaultValue = async (flag: string | null, folder: any) => {
  try {
    const siteTitle = folder.SiteTitle || folder.Title || "";
    const folderName = folder.FolderName || "null";
    const folderPath = folder.FolderPath || folder.folderpath || "";
    const docLibName = folder.DocumentLibraryName || "";
 
    const spRoot = spfi(folder.__siteUrl).using(SPFx(context));
 
    // ===== STEP 1: IsPrivate check =====
    let isPrivate = false;
 
    try {
      if (folderName !== "null") {
        const folderItems = await spRoot.web.lists
          .getByTitle("DMSFolderMaster")
          .items
          .select("Id", "IsPrivate", "FolderPath")
          .filter(`FolderPath eq '${folderPath}'`)
          .top(1)();
 
        if (folderItems && folderItems.length > 0) {
          isPrivate = !!folderItems[0].IsPrivate;
          setFpIsPrivateId(folderItems[0].Id);
        }
      } else {

          // rohit - 20/04/26 Start 
    //     const libItems = await spRoot.web.lists
    //       .getByTitle("DMSPreviewFormMaster")
    //       .items
    //       .select("Id", "IsPrivate", "DocumentLibraryName")
    //       .filter(`DocumentLibraryName eq '${docLibName}' and IsDocumentLibrary eq 1`)
    //       .top(1)();
 
    //     if (libItems && libItems.length > 0) {
    //       isPrivate = !!libItems[0].IsPrivate;
    //       setFpIsPrivateLibId(libItems[0].Id);
    //     }
    //   }
    // } catch (e) {
    //   console.warn("IsPrivate check failed", e);
    // }


    const libItems = await spRoot.web.lists
    .getByTitle("DMSFolderMaster")        // Rohit change list for DMSFolderMaster for library level permission check
    .items
    .select("Id", "IsPrivate", "DocumentLibraryName")
    // Adjust filter as needed – e.g., if IsLibrary column identifies root folders
    .filter(`DocumentLibraryName eq '${docLibName}' and IsLibrary eq 1`)
    .top(1)();
 
  if (libItems && libItems.length > 0) {
    isPrivate = !!libItems[0].IsPrivate;
    setFpIsPrivateLibId(libItems[0].Id);
        }
      }
    } catch (e) {
      console.warn("IsPrivate check failed", e);
    }


    // rohit - 20/04/26 end 
 
    // ===== STEP 2: Public hai to "No" set karo =====
    if (!isPrivate && flag !== "force") {
      setTogglePermission("No");
     // ✅ Pura hierarchical path build karo
let path = siteTitle;
if (docLibName) path += ` > ${docLibName}`;
 
// FolderPath se saare nested folders nikalo
// FolderPath example: /sites/Intranetdemos/ESSADMS3/DocumentLibraryName/ParentFolder/SubFolder/CurrentFolder
if (folderName !== "null" && folderPath) {
  const spSiteUrl = folder.__siteUrl || selectedFolder?.__siteUrl || "";
  const siteRelativePath = `${spSiteUrl}/${siteTitle}`.replace(/https?:\/\/[^/]+/, "");
  
  // FolderPath se site-relative prefix hatao
  let relativePath = folderPath;
  if (relativePath.startsWith(siteRelativePath)) {
    relativePath = relativePath.slice(siteRelativePath.length);
  }
  
  // "/" se split karo aur docLibName ke baad ke saare parts lo
  const parts = relativePath.split("/").filter(Boolean);
  const libIndex = parts.findIndex(
    (p: string) => p.toLowerCase() === docLibName.toLowerCase()
  );
  const folderParts = libIndex >= 0 ? parts.slice(libIndex + 1) : parts;
  
  // Har folder level " > " se join karo
  if (folderParts.length > 0) {
    path += ` > ${folderParts.join(" > ")}`;
  }
}
 
setFpPathState(path);
      return;
    }
 
    // ===== STEP 3: Role assignments load karo =====
    const spSubsite = spfi(`${folder.__siteUrl}/${siteTitle}`).using(SPFx(context));
    const roleAssignments = await fpFetchRoleAssignments(spSubsite, folderName, folderPath, docLibName);
 
    // ✅ Shared mapping function use karo
    const mapped = fpMapRoleAssignments(roleAssignments);
    setFolderPrivacyTableData(mapped);
    setTogglePermission("Yes");
 
    // ✅ Pura hierarchical path build karo
let path = siteTitle;
if (docLibName) path += ` > ${docLibName}`;
 
// FolderPath se saare nested folders nikalo
// FolderPath example: /sites/Intranetdemos/ESSADMS3/DocumentLibraryName/ParentFolder/SubFolder/CurrentFolder
if (folderName !== "null" && folderPath) {
  const spSiteUrl = folder.__siteUrl || selectedFolder?.__siteUrl || "";
  const siteRelativePath = `${spSiteUrl}/${siteTitle}`.replace(/https?:\/\/[^/]+/, "");
  
  // FolderPath se site-relative prefix hatao
  let relativePath = folderPath;
  if (relativePath.startsWith(siteRelativePath)) {
    relativePath = relativePath.slice(siteRelativePath.length);
  }
  
  // "/" se split karo aur docLibName ke baad ke saare parts lo
  const parts = relativePath.split("/").filter(Boolean);
  const libIndex = parts.findIndex(
    (p: string) => p.toLowerCase() === docLibName.toLowerCase()
  );
  const folderParts = libIndex >= 0 ? parts.slice(libIndex + 1) : parts;
  
  // Har folder level " > " se join karo
  if (folderParts.length > 0) {
    path += ` > ${folderParts.join(" > ")}`;
  }
}
 
setFpPathState(path);
 
  } catch (e) {
    console.error("fpLoadDefaultValue error", e);
    setTogglePermission("Yes");
  }
};
 
const fpFetchUsers = async (folder: any) => {
  try {
    const spRoot = spfi(folder.__siteUrl).using(SPFx(context));
    const allUsers = await spRoot.web.siteUsers();
 
    console.table(allUsers.map((u: any) => ({
      Id: u.Id,
      Title: u.Title,
      Email: u.Email,
      LoginName: u.LoginName,
      Source: folder.__siteUrl
    })));
 
    setFolderPrivacyUsers(
      (allUsers || [])
        .filter((u: any) => u.Email && !u.IsHiddenInUI)
        .map((u: any) => ({
          userId: u.Id,
          value: u.Title,
          label: `${u.Title} (${u.Email})`,
          email: u.Email,
        }))
    );
  } catch (e) {
    console.error("fpFetchUsers error", e);
    setFolderPrivacyUsers([]);
  }
};
 
const fpOpenPermission = async (folder: any) => {
  setSelectedFolder(folder);
  setTogglePermission(undefined);
  setRowsForPermission([{ id: 0, selectedUserForPermission: [], selectedPermission: "" }]);
  setFolderPrivacyTableData([]);
  setFpErrors({});
  setFpCurrentPage(1);
  setFpIsPrivateId(null);
  setFpIsPrivateLibId(null);
 
  try {
    const userdata = await spfi(folder.__siteUrl).using(SPFx(context)).web.currentUser();
    fpCurrentUserEmail.current = userdata.Email;
  } catch (e) { console.warn("currentUser fetch failed", e); }
 
  await fpLoadDefaultValue(null, folder);
  await fpFetchUsers(folder);
  setShowPermissionModal(true);
};
 
const fpSetPrivate = async () => {
  if (!selectedFolder) return;
  try {
    const siteTitle = selectedFolder.SiteTitle || selectedFolder.Title || "";
    const folderName = selectedFolder.FolderName || "null";
    const folderPath = selectedFolder.FolderPath || selectedFolder.folderpath || "";
    const docLibName = selectedFolder.DocumentLibraryName || "";
 
    const spRoot = spfi(selectedFolder.__siteUrl).using(SPFx(context));
    const spSubsite = spfi(`${selectedFolder.__siteUrl}/${siteTitle}`).using(SPFx(context));
 
    const loggedInUserEmail = context.pageContext.user.email;
    console.log("=== LOGGED IN USER ===", loggedInUserEmail);
 
    if (folderName === "null") {
      if (fpIsPrivateLibId !== null) {
        await spRoot.web.lists.getByTitle("DMSFolderMaster")          // rohit 20/04/26 ("DmsformpreviewMaster") change list for DMSFolderMaster for library level permission check
          .items.getById(fpIsPrivateLibId).update({ IsPrivate: true });
      }
    } else {
      if (fpIsPrivateId !== null) {
        await spRoot.web.lists.getByTitle("DMSFolderMaster")
          .items.getById(fpIsPrivateId).update({ IsPrivate: true });
      }
    }
 
    if (folderName !== "null") {
      const folderItem = await spSubsite.web
        .getFolderByServerRelativePath(folderPath).getItem();
      const data = await folderItem.select("HasUniqueRoleAssignments")();
      if (!data.HasUniqueRoleAssignments) {
        await folderItem.breakRoleInheritance(true);
      }
      const ensured = await spSubsite.web.ensureUser(loggedInUserEmail);
      const fullControlDef = await spSubsite.web.roleDefinitions.getByName("Full Control")();
      await folderItem.roleAssignments.add(ensured.data.Id, fullControlDef.Id);
    } else {
      const lib = spSubsite.web.lists.getByTitle(docLibName);
      const data = await lib.select("HasUniqueRoleAssignments")();
      if (!(data as any).HasUniqueRoleAssignments) {
        await lib.breakRoleInheritance(true);
      }
      const ensured = await spSubsite.web.ensureUser(loggedInUserEmail);
      const fullControlDef = await spSubsite.web.roleDefinitions.getByName("Full Control")();
      await lib.roleAssignments.add(ensured.data.Id, fullControlDef.Id);
    }
 

    // Rohit 20/04/26 start
//     await fpLoadDefaultValue("force", selectedFolder);
 
//   } catch (e) {
//     console.error("fpSetPrivate error", e);
//     Swal.fire('Error', 'Failed to set private permission.', 'error');
//   }
// };

await fpLoadDefaultValue("force", selectedFolder);
    // Rohit 20/04/2026 FIX: Update selectedFiles so the card badge reflects "Private" immediately
    const folderKey = selectedFolder.FolderPath || selectedFolder.folderpath || selectedFolder.DocumentLibraryName || "";
    setSelectedFiles((prev: any[]) =>
      prev.map((f: any) => {
        const fKey = f.FolderPath || f.folderpath || f.DocumentLibraryName || "";
        if (fKey && fKey === folderKey) {
          return { ...f, IsPrivate: true };
        }
        return f;
      })
    );
 
  } catch (e) {
    console.error("fpSetPrivate error", e);
    Swal.fire('Error', 'Failed to set private permission.', 'error');
  }
};

// Rohit 20/04/26 end
 
const fpValidate = (): boolean => {
  let isValid = true;
  const newErrors: any = {};
  rowsForPermission.forEach((row) => {
    if (!row.selectedUserForPermission?.length) {
      newErrors[row.id] = { ...newErrors[row.id], userSelect: "Please select at least one user." };
      isValid = false;
    }
    if (!row.selectedPermission) {
      newErrors[row.id] = { ...newErrors[row.id], permissionSelect: "Please select a permission." };
      isValid = false;
    }
  });
  setFpErrors(newErrors);
  return isValid;
};
 
const fpHandleCreate = async () => {
  if (!fpValidate() || !selectedFolder) return;
  try {
    const siteTitle = selectedFolder.SiteTitle || selectedFolder.Title || "";
    const folderName = selectedFolder.FolderName || "null";
    const folderPath = selectedFolder.FolderPath || selectedFolder.folderpath || "";
    const docLibName = selectedFolder.DocumentLibraryName || "";
 
    const spSubsite = spfi(`${selectedFolder.__siteUrl}/${siteTitle}`).using(SPFx(context));
 
    let securableObject: any;
    if (folderName !== "null") {
      const folderItem = await spSubsite.web
        .getFolderByServerRelativePath(folderPath)
        .getItem();
      const data = await folderItem.select("HasUniqueRoleAssignments")();
      if (!data.HasUniqueRoleAssignments) {
        await folderItem.breakRoleInheritance(true);
      }
      securableObject = folderItem;
    } else {
      const lib = spSubsite.web.lists.getByTitle(docLibName);
      const data = await lib.select("HasUniqueRoleAssignments")();
      if (!(data as any).HasUniqueRoleAssignments) {
        await lib.breakRoleInheritance(true);
      }
      securableObject = lib;
    }
 
    // ✅ STEP 1: Pehle saare duplicate check karo - koi bhi duplicate mila to rok do
    for (const row of rowsForPermission) {
      for (const user of row.selectedUserForPermission) {
        const ensured = await spSubsite.web.ensureUser(user.email || user.value);
        const subsiteUserId = ensured.data.Id;
 
        // ✅ folderPrivacyTableData me har row ek permission hai (flatMap ki wajah se)
        const duplicate = folderPrivacyTableData.find(
          (item: any) => item.userId === subsiteUserId && item.Permission === row.selectedPermission.value
        );
 
        if (duplicate) {
          setFpErrors((prev: any) => ({
            ...prev,
            [row.id]: {
              ...prev[row.id],
              userSelect: `"${user.value}" already has "${row.selectedPermission.value}" permission`,
            }
          }));
          return; // ✅ Poora function rok do
        }
      }
    }
 
    // ✅ STEP 2: Koi duplicate nahi - sab add karo
    let anyAdded = false;
    for (const row of rowsForPermission) {
      for (const user of row.selectedUserForPermission) {
        try {
          const ensured = await spSubsite.web.ensureUser(user.email || user.value);
          const subsiteUserId = ensured.data.Id;
          const roleDef = await spSubsite.web.roleDefinitions
            .getByName(row.selectedPermission.value)();
          await securableObject.roleAssignments.add(subsiteUserId, roleDef.Id);
          anyAdded = true;
        } catch (e) {
          console.warn(`Permission add failed for ${user.value}`, e);
        }
      }
    }
 
    if (!anyAdded) return;
 
    // ✅ STEP 3: Backend se fresh data lo aur shared mapping use karo
    const roleAssignments = await fpFetchRoleAssignments(spSubsite, folderName, folderPath, docLibName);
    const mapped = fpMapRoleAssignments(roleAssignments);
 
    setFolderPrivacyTableData(mapped);
    setRowsForPermission([{ id: 0, selectedUserForPermission: [], selectedPermission: "" }]);
    setFpErrors({});
 
    Swal.fire("Success", "Permissions added successfully.", "success");
 
  } catch (e) {
    Swal.fire("Error", "An error occurred.", "error");
  }
};
 
const fpHandleDeleteUser = async (userId: number, itemId: number, permission: string) => {
  const result = await Swal.fire({
    title: "Are you sure?",
    icon: "warning",
    showCancelButton: true,
    confirmButtonColor: "#3085d6",
    cancelButtonColor: "#d33",
    confirmButtonText: "Yes, remove it!",
  });
  if (!result.isConfirmed || !selectedFolder) return;
 
  try {
    const siteTitle = selectedFolder.SiteTitle || selectedFolder.Title || "";
    const folderName = selectedFolder.FolderName || "null";
    const folderPath = selectedFolder.FolderPath || selectedFolder.folderpath || "";
    const docLibName = selectedFolder.DocumentLibraryName || "";
 
    const spSubsite = spfi(`${selectedFolder.__siteUrl}/${siteTitle}`).using(SPFx(context));
 
    let securableObject: any;
    if (folderName !== "null") {
      securableObject = await spSubsite.web
        .getFolderByServerRelativePath(folderPath)
        .getItem();
    } else {
      securableObject = spSubsite.web.lists.getByTitle(docLibName);
    }
 
    // ✅ Sirf ek permission remove karo jo delete button pe click hua
    try {
      const roleDef = await spSubsite.web.roleDefinitions.getByName(permission)();
      await securableObject.roleAssignments.remove(userId, roleDef.Id);
    } catch (e) {
      console.warn(`Role remove failed for "${permission}"`, e);
    }
 
    // ✅ Backend se fresh data lo aur shared mapping use karo
    const roleAssignments = await fpFetchRoleAssignments(spSubsite, folderName, folderPath, docLibName);
    const mapped = fpMapRoleAssignments(roleAssignments);
 
    setFolderPrivacyTableData(mapped);
    Swal.fire({ title: "Removed!", text: "Permission removed successfully.", icon: "success" });
 
  } catch (e) {
    Swal.fire("Error", "An error occurred while removing permission.", "error");
  }
};
 
// ===== ManageFolderPermission Inline Logic END =====
// Ritik 20/2/26 for manage permission folder

  const deleteFileFolder = async (file: any, siteUrl: string, context: any) => {
    try {
      console.log("[deleteFileFolder] file:", file);
      console.log("[deleteFileFolder] siteUrl:", siteUrl);

      const isoDate = new Date().toISOString();


      const siteSP = spfi(siteUrl).using(SPFx(context));
      const fileItem = await siteSP.web
        .getFileByServerRelativePath(file.ServerRelativeUrl)
        .getItem();

      await fileItem.update({
        IsDeleted: isoDate,
      });


      const baseSiteUrl = siteUrl.split("/sites/")[0] + "/sites/";
      const siteCollection = siteUrl.split("/sites/")[1].split("/")[0]; // AlRostmaniSpfx2
      const entityName = siteUrl.split("/").pop(); // Rohittest

      console.log("[deleteFileFolder] entityName:", entityName);

      const spRoot = spfi(`${baseSiteUrl}${siteCollection}`).using(SPFx(context));


      const fmItems = await spRoot.web.lists
        .getByTitle(`DMS${entityName}FileMaster`)
        .items.filter(`FileName eq '${file.Name}' and IsDeleted eq null`)();

      if (!fmItems.length) {
        console.warn("FileMaster item not found:", file.Name);
        return;
      }

      await spRoot.web.lists
        .getByTitle(`DMS${entityName}FileMaster`)
        .items.getById(fmItems[0].Id)
        .update({
          IsDeleted: isoDate,
        });

      setSelectedFiles((prev: any[]) =>
        prev.filter(
          (f) => f.ServerRelativeUrl !== file.ServerRelativeUrl
        )
      );

      console.log("File soft deleted and FileMaster updated");
    } catch (err) {
      console.error("Delete failed:", err);
    }
  };





  // toggle favourite in my folder
  //   const toggleFavouriteDoc = async (
  //   file: any,
  //   siteUrl: string,
  //   context: any
  // ) => {
  //   try {
  //     const siteSP = spfi(siteUrl).using(SPFx(context));

  //     if (!file?.ServerRelativeUrl) {
  //       console.error(
  //         "[toggleFavourite] No ServerRelativeUrl found on file:",
  //         file
  //       );
  //       return;
  //     }

  //     // Get the list item backing this file
  //     const item = await siteSP.web
  //       .getFileByServerRelativePath(file.ServerRelativeUrl)
  //       .getItem();

  //     // Read current favourite value
  //     const currentItem = await item.select("Id", "IsFavourite")();
  //     const currentFav = currentItem?.IsFavourite || false;

  //     console.log(`[toggleFavourite] Current favourite: ${currentFav}`);

  //     // Toggle
  //     await item.update({
  //       IsFavourite: !currentFav,
  //     });

  //     console.log(
  //       `[toggleFavourite] File ${file.Name} is now ${
  //         !currentFav ? "marked as favourite" : "unmarked as favourite"
  //       }`
  //     );

  //     alert(
  //       ` File ${file.Name} is now ${
  //         !currentFav ? "marked as favourite" : "unmarked as favourite"
  //       }`
  //     );

  //     // Optional: return new status
  //     return !currentFav;
  //   } catch (err) {
  //     console.error("[toggleFavourite] Error toggling favourite:", err);
  //   }
  // };


  // const toggleFavouriteDoc = async (





  //   file: any,
  //   siteUrl: string,
  //   context: any
  // ) => {
  //   //       console.log("  Addhyan [toggleFavourite] Toggling favourite for:");
  //   //       console.log("this is file value:",    file);
  //   //       console.log("this is siteUrl value:",    siteUrl);
  //   //       console.log("this is context value:",    context);

  //   //       const siteurl = siteUrl;
  //   //       const urlsss = siteUrl.split("/sites/")[1].split("/")[0];

  //   //       console.log("this is urlsss value:",    urlsss);
  //   //       const baseSiteUrl = siteUrl.split("/sites/")[0] + "/sites/";
  //   //       console.log("this is baseSiteUrl value:",    baseSiteUrl);

  //   //       const test = siteurl.split("/").pop();
  //   //       console.log("this is test value:",    test);
  //   // //addhyan fav
  //   //        const meName = context.pageContext.user.displayName || "";
  //   //     const meEmail =
  //   //       (context.pageContext as any)?.user?.email ||
  //   //       (context.pageContext as any)?.user?.loginName ||
  //   //       "";


  //   //       console.log("Addhyan   ->>>[toggleFavourite] meName:", meName);
  //   //       console.log("Addhyan   ->>>[toggleFavourite] meEmail:", meEmail);      

  //   //        const siteSP = spfi(`${baseSiteUrl}${urlsss}`).using(SPFx(context));
  //   //       const datacheck  = await siteSP.web.lists.getByTitle(`DMS${test}FileMaster`).items.select("Id", "Title", " IsFavourite").top(5000)();
  //   //       console.log("Addhyan   ->>>[toggleFavourite] datacheck:", datacheck);

  //   console.log("🔁 Toggle Favourite Started");
  //   console.log("File Object:", file);

  //   const fileUniqueId = file.UniqueId;
  //   console.log("File UniqueId:", fileUniqueId);

  //   const meName = context.pageContext.user.displayName || "";
  //   const meEmail =
  //     (context.pageContext as any)?.user?.email ||
  //     (context.pageContext as any)?.user?.loginName ||
  //     "";

  //   console.log("User Name:", meName);
  //   console.log("User Email:", meEmail);

  //   const baseSiteUrl = siteUrl.split("/sites/")[0] + "/sites/";
  //   const siteCollectionName = siteUrl.split("/sites/")[1].split("/")[0];
  //   const subSiteName = siteUrl.split("/").pop(); // TestHub1

  //   console.log("BaseSiteUrl:", baseSiteUrl);
  //   console.log("Site Collection:", siteCollectionName);
  //   console.log("Sub Site:", subSiteName);

  //   const siteSP = spfi(`${baseSiteUrl}${siteCollectionName}`).using(
  //     SPFx(context)
  //   );

  //   const listName = `DMS${subSiteName}FileMaster`;

  //   const items = await siteSP.web.lists
  //     .getByTitle(listName)
  //     .items
  //     .select("Id", "FileUID", "IsFavourite", "CurrentUser")
  //     .filter(
  //       `FileUID eq '${fileUniqueId}'`
  //     )
  //     .top(1)();

  //   console.log("Matched Items:", items);
  //   if (items.length === 0) {
  //     console.log("❌ No matching record found for this user & file");
  //     return;
  //   }

  //   const item = items[0];
  //   const newFavouriteValue = !item.IsFavourite;

  //   console.log("Old IsFavourite:", item.IsFavourite);
  //   console.log("New IsFavourite:", newFavouriteValue);

  //   await siteSP.web.lists
  //     .getByTitle(listName)
  //     .items
  //     .getById(item.Id)
  //     .update({
  //       IsFavourite: newFavouriteValue
  //     });

  //   console.log("✅ Favourite status updated successfully");


  //   return newFavouriteValue;






















  //   try {
  //     // const siteSP = spfi(siteUrl).using(SPFx(context));


  //     // if (!file?.ServerRelativeUrl) {
  //     //   console.error(
  //     //     "[toggleFavourite] No ServerRelativeUrl found on file:",
  //     //     file
  //     //   );
  //     //   return;
  //     // }

  //     // // Get the list item backing this file
  //     // const item = await siteSP.web
  //     //   .getFileByServerRelativePath(file.ServerRelativeUrl)
  //     //   .getItem();

  //     // // Read current favourite value
  //     // const currentItem = await item.select("Id", "IsFavourite")();
  //     // const currentFav = currentItem?.IsFavourite || false;

  //     // console.log(`[toggleFavourite] Current favourite: ${currentFav}`);

  //     // // Toggle
  //     // await item.update({
  //     //   IsFavourite: !currentFav,
  //     // });

  //     // console.log(
  //     //   `[toggleFavourite] File ${file.Name} is now ${
  //     //     !currentFav ? "marked as favourite" : "unmarked as favourite"
  //     //   }`
  //     // );

  //     // alert(
  //     //   ` File ${file.Name} is now ${
  //     //     !currentFav ? "marked as favourite" : "unmarked as favourite"
  //     //   }`
  //     // );

  //     // // Optional: return new status
  //     // return !currentFav;
  //   } catch (err) {
  //     console.error("[toggleFavourite] Error toggling favourite:", err);
  //   }
  // };


  // Addhyan 19/2/26

  const toggleFavouriteDoc = async (
  file: any,
  siteUrl: string,
  context: any
) => {
  console.log("siteUrl:", siteUrl);
    console.log("File Object:", file);
  console.log("🔁 Toggle Favourite Started");
  const lastPart = siteUrl.split("/").pop();
  const siteUrls = file.__siteUrl || context.pageContext.site.absoluteUrl;
  console.log("Derived site URL for file:", siteUrls);
 
 
 
 
  const fileUniqueId = file.UniqueId;
  const fileUid = file.FileUID;
  console.log("File UniqueId:", fileUniqueId);
  console.log( " FileUID:", fileUid);
 
  const meEmail =
    (context.pageContext as any)?.user?.email ||
    (context.pageContext as any)?.user?.loginName ||
    "";
 
  const baseSiteUrl = siteUrl.split("/sites/")[0] + "/sites/";
  const siteCollectionName = siteUrl.split("/sites/")[1].split("/")[0];
  const subSiteName = siteUrl.split("/").pop();
 
  const siteSP = spfi(`${baseSiteUrl}${siteCollectionName}`).using(
    SPFx(context)
  );
 
  const listName = `DMS${subSiteName}FileMaster`;
 
 
 
const existingItems = await siteSP.web.lists
  .getByTitle(listName)
  .items
  .select("*") // 🔥 sab columns le aayega
  .filter(`FileUID eq '${fileUniqueId}'`)();
 
  console.log("All matching items for FileUID:", existingItems);
  const fullMetaData = existingItems[0];
 
 
 
 
  // ✅ Important Fix: User-wise filter
  const items = await siteSP.web.lists
    .getByTitle(listName)
    .items

    // ritik 29/04/26 - start
    // .select("Id", "FileUID", "IsFavourite", "CurrentUser", "MyRequest", )
    // .filter(
    //   `FileUID eq '${fileUniqueId}' and CurrentUser eq '${meEmail}' and MyRequest eq 0`
    // )();
 


    .select("Id", "FileUID", "IsFavourite", "CurrentUser", "MyRequest") // Ritik 27/04/26 removed and MyRequest eq 0
    .filter(
      `FileUID eq '${fileUniqueId}' and CurrentUser eq '${meEmail}'` // Ritik 27/04/26 removed and MyRequest eq 0
    )();

    // ritik 29/04/26 - start

  console.log("Matched Items:", items);
 
  // 🔥 CASE 1: Record not found → Add new
  if (items.length === 0) {
 
    console.log("No record found → Adding new item");
 
    const payload = {
      FileUID: fileUniqueId,
      FileName: file.Name,
      FileSize: file.Length,
      DocumentLibraryName : fullMetaData.DocumentLibraryName ,
      CurrentFolderPath : fullMetaData.CurrentFolderPath,
      SiteID: siteUrl,
      SiteName: lastPart,
      IsFavourite: true,
      CurrentUser: meEmail,
      Processname: fullMetaData.Processname,
      Status: fullMetaData.Status,
      MyRequest: false,
      FilePreviewURL: fullMetaData.FilePreviewURL,
      RequestNo: `DMS-${fileUniqueId}`
    };
 
    await siteSP.web.lists
      .getByTitle(listName)
      .items
      .add(payload);
 
    console.log("✅ New favourite added");

    // addhyan
  // window.location.reload(); // Refresh the page to reflect changes  
  Swal.fire({
        title: 'Added to Favourites!',
        text: 'The file has been successfully added to your favourites.',
        icon: 'success',
        confirmButtonText: 'ok'
      });
    return true;
  }
 
  // 🔥 CASE 2: Record exists → Toggle
  const item = items[0];
  const newFavouriteValue = !item.IsFavourite;
 
  await siteSP.web.lists
    .getByTitle(listName)
    .items
    .getById(item.Id)
    .update({
      IsFavourite: newFavouriteValue,
      MyRequest: false
    });
 
  console.log("✅ Favourite status updated:", newFavouriteValue);
  // addhyan
//  window.location.reload(); // Refresh the page to reflect changes
  // const refreshed = await loadViewData("MyFavourite");
// if (refreshed) {
//   setSelectedFiles([...refreshed]);
// }
Swal.fire({
        title: 'Added to Favourites!',
        text: 'The file has been successfully added to your favourites.',
        icon: 'success',
        confirmButtonText: 'ok'
      });
  return newFavouriteValue;
 
};
  // abhay change for seach in all tabs
  // #region SEARCH TAB 
  // Abhay 10/10/25 merger 

  const [searchTerm, setSearchTerm] = useState("");
  //ritik 19/01/2026 for search input box value
  const [searchInput, setSearchInput] = useState("");





  // const filteredFiles = !searchTerm.trim()
  //   ? selectedFiles
  //   : selectedFiles.filter((file: any) => {
  //     const term = searchTerm.toLowerCase();

  //     if (activeView === "My Folders") {
  //       return (
  //         (file.FolderName || "").toLowerCase().includes(term) ||
  //         (file.SiteTitle || "").toLowerCase().includes(term) ||
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term)

  //       );
  //     }
  //     else if (activeView === "SharedWithMe") {
  //       return (
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.ShareWithMe || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term)
  //       );
  //     }
  //     else if (activeView === "SharedWithOthers") {
  //       return (
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.ShareWithOther || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term)
  //       );
  //     }
  //     else if (activeView === "MyRequest") {
  //       return (
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.RequestStatus || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term)
  //       );
  //     }
  //     else if (activeView === "MyFavourite") {
  //       return (
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term)
  //       );
  //     }
  //     else if (activeView === "RecycleBin") {
  //       return (
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.DeletedBy || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term)
  //       );
  //     }
  //     else {
  //       // Default case
  //       return (
  //         (file.FileName || "").toLowerCase().includes(term) ||
  //         (file.DocumentLibraryName || "").toLowerCase().includes(term) ||
  //         (file.Status || "").toLowerCase().includes(term)
  //       );
  //     }
  //   });

  // const pageSize = 12;
  // let location: string = "";
  // const paginatedFiles = useMemo(() => {
  //   const start = (currentPage - 1) * pageSize;
  //   const end = start + pageSize;
  //   return filteredFiles.slice(start, end);
  // }, [filteredFiles, currentPage]);

  // useEffect(() => {
  //   setCurrentPage(1);
  // }, [searchTerm]);

  // // 🧩 ✅ Clear search when section changes
  // useEffect(() => {
  //   setSearchTerm("");
  //   setCurrentPage(1);
  // }, [activeView]);




  //Handle browse view search //Ritik 14/01/2025

  const filteredFiles = !searchTerm.trim()
    ? selectedFiles
    : selectedFiles.filter((file: any) => {
      const term = searchTerm.toLowerCase();

      // Handle Browse View (Folder Hierarchy)
      if (!activeView || activeView === "browse") {
        return (
          (file.Name || "").toLowerCase().includes(term) ||
          (file.FileLeafRef || "").toLowerCase().includes(term) ||
          (file.ServerRelativeUrl || "").toLowerCase().includes(term)
        );
      }

      if (activeView === "My Folders") {
        return (
          (file.FolderName || "").toLowerCase().includes(term) ||
          (file.SiteTitle || "").toLowerCase().includes(term) ||
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term)
        );
      }
      else if (activeView === "SharedWithMe") {
        return (
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.ShareWithMe || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term)
        );
      }
      else if (activeView === "SharedWithOthers") {
        return (
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.ShareWithOther || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term)
        );
      }
      else if (activeView === "MyRequest") {
        return (
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.RequestStatus || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term)
        );
      }
      else if (activeView === "MyFavourite") {
        return (
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term)
        );
      }
      else if (activeView === "RecycleBin") {
        return (
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.DeletedBy || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term)
        );
      }
      else {
        // Default case
        return (
          (file.FileName || "").toLowerCase().includes(term) ||
          (file.DocumentLibraryName || "").toLowerCase().includes(term) ||
          (file.Status || "").toLowerCase().includes(term)
        );
      }
    });
  
  const [pageSize, setPageSize] = useState(12); //Rohit 23/04/2026
  let location: string = "";
  const paginatedFiles = useMemo(() => {
    // Apply sorting to filteredFiles
    let sortedFiles = [...filteredFiles];
    if (sortColumn) {
      sortedFiles.sort((a, b) => {
        let aValue: any, bValue: any;

        switch (sortColumn) {
          case "name":
            aValue = (a.FileName || a.Name || a.FolderName || "").toLowerCase();
            bValue = (b.FileName || b.Name || b.FolderName || "").toLowerCase();
            break;
          case "size":
            aValue = parseFloat(a.Length || a.FileSize || "0") || 0;
            bValue = parseFloat(b.Length || b.FileSize || "0") || 0;
            break;
          case "library":
            aValue = (a.DocumentLibraryName || "").toLowerCase();
            bValue = (b.DocumentLibraryName || "").toLowerCase();
            break;
          case "createdDate":
            aValue = new Date(a.TimeCreated || a.Created || 0).getTime();
            bValue = new Date(b.TimeCreated || b.Created || 0).getTime();
            break;
          case "status":
            aValue = (a.Status || "").toLowerCase();
            bValue = (b.Status || "").toLowerCase();
            break;
          default:
            return 0;
        }

        if (aValue < bValue) return sortDirection === "asc" ? -1 : 1;
        if (aValue > bValue) return sortDirection === "asc" ? 1 : -1;
        return 0;
      });
    }

    // Apply pagination
    const start = (currentPage - 1) * pageSize;
    const end = start + pageSize;
    return sortedFiles.slice(start, end);
  }, [filteredFiles, currentPage, sortColumn, sortDirection,pageSize]);
  //Rohit 22/04/2026 ----------------start
  const getQuickViewRenderKey = (file: any, idx: number) => {
    return [
      file.FileUID,
      file.UniqueId,
      file.__siteUrl,
      file.__fileMasterList,
      file.Id ?? file.ID ?? idx,
    ]
      .filter((value) => value !== undefined && value !== null && value !== "")
      .join("__");
  };
  //Rohit 22/04/2026 ------------------end
  //Rohit 22/04/2026------------------start
  
  // AFTER:
const columnFilteredFiles = useMemo(() => {
  const hasColumnFilter = columnSearch.name || columnSearch.size || columnSearch.library || columnSearch.status;
  
  // If column filters active, search across ALL filtered files (not just current page)
  const sourceFiles = hasColumnFilter ? filteredFiles : paginatedFiles;
  
  const filtered = sourceFiles.filter((file: any) => {
    const nameMatch = !columnSearch.name || 
      (file.FileName || file.Name || file.FolderName || "").toLowerCase().includes(columnSearch.name.toLowerCase());
    const sizeMatch = !columnSearch.size || 
      (file.FileSize || (file.Length ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB` : "")).toLowerCase().includes(columnSearch.size.toLowerCase());
    const libMatch = !columnSearch.library || 
      (file.DocumentLibraryName || (file.TimeCreated ? new Date(file.TimeCreated).toLocaleDateString() : "")).toLowerCase().includes(columnSearch.library.toLowerCase());
    const statusMatch = !columnSearch.status || 
      (file.Status || "").toLowerCase().includes(columnSearch.status.toLowerCase());
    return nameMatch && sizeMatch && libMatch && statusMatch;
  });

  // If column filters active, apply sorting + pagination here
  if (hasColumnFilter) {
    let sorted = [...filtered];
    if (sortColumn) {
      sorted.sort((a, b) => {
        let aValue: any, bValue: any;
        switch (sortColumn) {
          case "name":    aValue = (a.FileName || a.Name || a.FolderName || "").toLowerCase(); bValue = (b.FileName || b.Name || b.FolderName || "").toLowerCase(); break;
          case "size":    aValue = parseFloat(a.Length || a.FileSize || "0") || 0; bValue = parseFloat(b.Length || b.FileSize || "0") || 0; break;
          case "library": aValue = (a.DocumentLibraryName || "").toLowerCase(); bValue = (b.DocumentLibraryName || "").toLowerCase(); break;
          case "createdDate": aValue = new Date(a.TimeCreated || a.Created || 0).getTime(); bValue = new Date(b.TimeCreated || b.Created || 0).getTime(); break;
          case "status":  aValue = (a.Status || "").toLowerCase(); bValue = (b.Status || "").toLowerCase(); break;
          default: return 0;
        }
        if (aValue < bValue) return sortDirection === "asc" ? -1 : 1;
        if (aValue > bValue) return sortDirection === "asc" ? 1 : -1;
        return 0;
      });
    }
    const start = (currentPage - 1) * pageSize;
    return sorted.slice(start, start + pageSize);
  }

  return filtered;
}, [filteredFiles, paginatedFiles, columnSearch, sortColumn, sortDirection, currentPage, pageSize]);

  const paginationTotalCount = useMemo(() => {
  const hasColumnFilter = columnSearch.name || columnSearch.size || columnSearch.library || columnSearch.status;
  if (!hasColumnFilter) return filteredFiles.length;
  // Count total matches across all pages when column filter is active
  return filteredFiles.filter((file: any) => {
    const nameMatch = !columnSearch.name || (file.FileName || file.Name || file.FolderName || "").toLowerCase().includes(columnSearch.name.toLowerCase());
    const sizeMatch = !columnSearch.size || (file.FileSize || (file.Length ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB` : "")).toLowerCase().includes(columnSearch.size.toLowerCase());
    const libMatch  = !columnSearch.library || (file.DocumentLibraryName || "").toLowerCase().includes(columnSearch.library.toLowerCase());
    const statusMatch = !columnSearch.status || (file.Status || "").toLowerCase().includes(columnSearch.status.toLowerCase());
    return nameMatch && sizeMatch && libMatch && statusMatch;
  }).length;
}, [filteredFiles, columnSearch]);

  //Rohit 22/04/2026------------------end
  

  // Handle sorting
  const handleSort = (column: string) => {
    if (sortColumn === column) {
      setSortDirection(sortDirection === "asc" ? "desc" : "asc");
    } else {
      setSortColumn(column);
      setSortDirection("asc");
    }
    setCurrentPage(1); // Reset to first page when sorting
  };

  useEffect(() => {
    setCurrentPage(1);
    setColumnSearch({ name: '', size: '', library: '', status: '' });  //Rohit 22/04/2026
  }, [searchTerm]);

  useEffect(() => {            //Rohit 23/04/2026
  setCurrentPage(1);
}, [pageSize]);

  // 🧩 ✅ Clear search when section changes
  useEffect(() => {
    setSearchTerm("");
    setSearchInput("");
    setCurrentPage(1);
    setColumnSearch({ name: '', size: '', library: '', status: '' }); // column search reset
  }, [activeView]);
  const extensionColors: any = {
    doc: "#105abe", // Blue
    docx: "#105abe",
    txt: "#bebfc1", // Gray (Text Files)
    pdf: "#fe0100", // Red (PDFs)
    xls: "#257952", // Green (Excel)
    xlsx: "#257952",
    zip: "#fcc41e", // Yellow (Archives)

    // 🎬 Video Files  
    mp4: "#ff5733", // Orange-Red  
    avi: "#ff5733",
    mkv: "#ff5733",
    mov: "#ff5733",
    wmv: "#ff5733",
    flv: "#ff5733",

    // 🎵 Audio Files  
    mp3: "#4caf50", // Green  
    wav: "#4caf50",
    flac: "#4caf50",
    aac: "#4caf50",
    ogg: "#4caf50",

    // 🖼️ Image Files  
    jpg: "#ff9800", // Orange  
    jpeg: "#ff9800",
    png: "#00bcd4", // Cyan  
    gif: "#9c27b0", // Purple  
    svg: "#673ab7", // Dark Purple  
    webp: "#009688", // Teal  

    default: "#17a2b8", // Teal (Unknown Files)
  };

  // helper function to get actual file name // Ritik today 
  const getActualFileName = (file: any) => {
    if (file.FileName) return file.FileName;

    if (file.Name) {
      // "/{name}" remove karega
      return file.Name.split("/")[0];
    }

    return "";
  };


const createFileExtensionHtml = (FileName: any, bgColor?: string, isListView: boolean = false) => {
  const fileExtension = FileName?.split(".")?.pop()?.toLowerCase() || "file";
  // Mapping icons dynamically to your extensions
  let selectedIcon = faFile;
  let iconColor = extensionColors.default;
  switch (fileExtension) {
    case "pdf": 
      selectedIcon = faFilePdf; 
      iconColor = extensionColors.pdf;
      break;
    case "doc":
    case "docx": 
      selectedIcon = faFileWord; 
      iconColor = extensionColors.docx;
      break;
    case "xls":
    case "xlsx": 
      selectedIcon = faFileExcel; 
      iconColor = extensionColors.xlsx;
      break;
    case "png":
    case "jpg":
    case "jpeg":
    case "gif":
    case "svg":
    case "webp": 
      selectedIcon = faFileImage; 
      iconColor = extensionColors.jpg;
      break;
    case "mp4":
    case "mov":
    case "avi": 
      selectedIcon = faFileVideo; 
      iconColor = extensionColors.mp4;
      break;
    case "zip":
    case "rar": 
      selectedIcon = faFileArchive; 
      iconColor = extensionColors.zip;
      break;
    case "txt": 
      selectedIcon = faFileAlt; 
      iconColor = extensionColors.txt;
      break;
    default: 
      selectedIcon = faFile;
      iconColor = extensionColors.default;
  }
  const iconData = icon(selectedIcon).html[0];
  const finalColor = bgColor || iconColor;
  // ✅ LIST VIEW: Icon only (no background, no text)
  if (isListView) {
    return `<div class="newicon1" style="
      
        height: 18px; 
        display: flex; 
        align-items: center; 
        fill: ${finalColor};
        color: ${finalColor};
      ">
        ${iconData}
</div>`;
  }
  // ✅ GRID VIEW: Badge with background + icon + text (default behavior)
  return ` <div class="file-extension-badge" style="
      background-color: ${finalColor};
      color: white;
      padding: 6px 15px 5px 15px;;
      border-radius: 30px;
      font-weight: bold;
      display: inline-flex;
      align-items: center;
      gap: 0px;
      font-size: 11px;">
<div style="width: 14px; margin-top:-2px; height: 14px; display: flex; align-items: center; fill: currentColor;">
          ${iconData}
</div>
<span>${fileExtension.toUpperCase()}</span>
</div>`;
}





  //Preview Header Resolver (Quick View vs Folder Hierarchy) Ritik 29/01/2026
  // 🔹 Preview Header Resolver (Quick View vs Folder Hierarchy)
  const getPreviewHeader = (file: any, activeView: string) => {
    const fileName =
      file.FileName ||
      file.Name ||
      file.Title ||
      (file.FilePreviewURL ? file.FilePreviewURL.split("/").pop() : "") ||
      "Untitled file";

    // 🔹 Quick Views → ONLY file name
    const quickViews = [
      "My request",
      "My favourite",
      "Share with me",
      "Share with other",
      "Recycle bin",
    ];

    if (quickViews.includes(activeView)) {
      return {
        name: fileName,
        path: "",
      };
    }

    // 🔹 Folder Hierarchy → name + full path
    let path = "";

    if (file.CurrentFolderPath) {
      path = file.CurrentFolderPath;
    } else if (file.__folderPath) {
      path = file.__folderPath;
    } else if (file.ServerRelativeUrl) {
      path = file.ServerRelativeUrl.substring(
        0,
        file.ServerRelativeUrl.lastIndexOf("/")
      );
    }

    return {
      name: fileName,
      path,
    };
  };



// srs 20/2/26
useEffect(() => {
  if (showRenameMetadataModal && selectedFileForRenameetadata) {
    const fetchColumns = async () => {
      setLoadingColumns(true);
      try {
        const masterSiteUrl = selectedFileForRenameetadata.__siteUrl;
        const siteName = selectedFileForRenameetadata.SiteTitle;
        const libraryName = selectedFileForRenameetadata.DocumentLibraryName;

        // 1. Create the new web object
        // 2. Use AssignFrom(sp.web) to copy the SPFx context/observers
        const remoteWeb = Web(masterSiteUrl).using(AssignFrom(sp.web));

        const columns = await remoteWeb.lists
          .getByTitle("DMSPreviewFormMaster")
          .items
          .select("ColumnName", "ColumnType", "ID", "IsRename")
          .filter(`SiteName eq '${siteName}' and DocumentLibraryName eq '${libraryName}' and IsDocumentLibrary eq 0`)();
        
        setExistingColumns(columns);
      } catch (error) {
        console.error("Error fetching columns:", error);
      } finally {
        setLoadingColumns(false);
      }
    };
    fetchColumns();
  }
}, [showRenameMetadataModal, selectedFileForRenameetadata]);

const handleSaveRename = async () => {
  try {
    const masterSiteUrl = selectedFileForRenameetadata.__siteUrl;
    
    // Connect the remote web to your current PnPjs configuration
    const remoteWeb = Web(masterSiteUrl).using(AssignFrom(sp.web));
    const list = remoteWeb.lists.getByTitle("DMSPreviewFormMaster");
    // aman 25/3/26
    setLoadingColumns(true);
    const updatePromises = existingColumns.map(item => 
      list.items.getById(item.ID).update({ 
        IsRename: item.IsRename 
      })
    );

    await Promise.all(updatePromises);
    
    Swal.fire('Success', 'Columns updated successfully', 'success');
    setShowRenameMetadataModal(false);
    // aman 25/3/26
    setExistingColumns([]);
    if(selectedCurrentNode){
      await loadFilesForNode(selectedCurrentNode);
    } 
  } catch (error) {
    console.error("Error updating columns:", error);
    Swal.fire('Error', 'Failed to update columns', 'error');
  }
  // Aman 25/3/26
  finally {
    setLoadingColumns(false);
  }
 
};

  return (
    <div
      id="maincontainer"
      style={{
        display: "flex",
        height: "calc(100vh - 50px)",
        marginTop: "10px",
      }}
    >
      {/* sourish 30/9/25 */}
      <div
        className="app-menu"
        id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>
      {/* sourish 30/9/25 */}
      <div className="content-page">
        <HorizontalNavbar _context={sp} />
        <div className="content" style={{ marginLeft: `${!useHide ? '240px' : '80px'}`, marginTop: '0.8rem' }}>
          {/* 
      <HorizontalNavbar _context={sp}/>
      <div className="content" style={{marginLeft: `${!useHide ? '240px' : '80px'}`,marginTop:'0.8rem'}}> */}
          {/* Left Panel with Quick Views and Folder Hierarchy */}
          <div className="content-page">
            <div className="" style={{ 
    background: '#fff', 
    borderBottom: '1px solid #eee', 
    marginBottom: '10px', 
    padding: '8px 10px', 
    minHeight: '65px', 
    display: 'flex',
    alignItems: 'center'
}}>
  <div className="row w-100 align-items-center" style={{ margin: 0 }}>
    <div className="col-md-3">
 <h2 className="page-title fw-bold m-0 mb-0 pt-5 mt-0 font-20"> Quick Views</h2>
    </div>
    
        
    <div className="col-md-9 d-flex align-items-center gap-4 justify-content-end" style={{ padding: 0 }}>
      
      {/* Group 1: Create Group */}
      <div className="d-flex flex-column align-items-center" style={{ borderRight: '1px solid #eee', paddingRight: '20px' }}>

        <div className="d-flex gap-3">
          <button 
            type="button"

                      // aman -20/04/26 start 


                              onClick={() => {
                if (!isCreateAllowed) return;
                setActiveComponent(true);
                setShowUploadPanel(false);
              }}
              disabled={!isCreateAllowed}
              style={{
                background: 'none',
                marginTop:'0px',
                border: 'none',
                padding: 0,
                opacity: !isCreateAllowed ? 0.4 : 1,
                cursor: !isCreateAllowed ? 'not-allowed' : 'pointer'
              }}

          //   onClick={() => setActiveComponent(true)}
          //   disabled={!selectedCurrentNode || selectedCurrentNode.type === "site"}
          //   style={{ background: 'none',   marginTop:'0px',  border: 'none', padding: 0, opacity: (!selectedCurrentNode || selectedCurrentNode.type === "site") ? 0.4 : 1, cursor: (!selectedCurrentNode || selectedCurrentNode.type === "site") ? 'not-allowed' : 'pointer' }}
          // >
       
            >
            <span className="mb-1 mt-2" data-tooltip="Create Folder">
            <img src={require("../assets/createnew.png")} alt="Create" /></span>
          </button>

          <button 
            type="button"
                onClick={() => {
                  if (
                    !selectedCurrentNode ||
                    !(selectedCurrentNode.type === "library" || selectedCurrentNode.type === "folder")
                  ) return;
                
                  setShowUploadPanel(true);
                  setActiveComponent(false);
                }}
                disabled={!isUploadAllowed}
                style={{
                  background: 'none',
                  marginTop:'0px',
                  border: 'none',
                  padding: 0,
                  opacity: !isUploadAllowed ? 0.4 : 1,
                  cursor: !isUploadAllowed ? 'not-allowed' : 'pointer'
                }}
            // onClick={() => {setShowUploadPanel(true); setActiveComponent(false)}}
            // disabled={!selectedCurrentNode || !(selectedCurrentNode.type === "library" || selectedCurrentNode.type === "folder")}
            // style={{ background: 'none',  marginTop:'0px', border: 'none', padding: 0, opacity: (!selectedCurrentNode || !(selectedCurrentNode.type === "library" || selectedCurrentNode.type === "folder")) ? 0.4 : 1, cursor: (!selectedCurrentNode || !(selectedCurrentNode.type === "library" || selectedCurrentNode.type === "folder")) ? 'not-allowed' : 'pointer' }}
          
          >
             <span className="mb-1 mt-2" data-tooltip="Upload File">
            <img src={require("../assets/uploafnew.png")} alt="Upload" /></span>
          </button>
        </div>
                <span style={{ fontSize: '14px',  color: '#333', marginTop: '5px' }}>Create</span>
      </div>
   {/* // aman -20/04/26 end  */}
      {/* srs 31/3/26 comment Ask Ai  */}
      {/* Group 2: ASK AI */}
      {/* <div className="d-flex flex-column align-items-center" style={{ borderRight: '1px solid #eee', paddingRight: '20px' }}>
       
        <div className="d-flex">
          <button 
            type="button"
            onClick={() => { console.log("Ask AI clicked"); }}
            style={{ background: 'none', border: 'none', padding: 0,  marginTop:'0px', cursor: 'pointer' }}
          >
            <span className="mb-1 mt-2" data-tooltip="Ask AI">
            <img src={require("../assets/newr.png")} alt="Ask AI" /></span>
          </button>
        </div>
         <span style={{ fontSize: '14px', color: '#333', marginTop: '5px' }}>Ask AI</span>
      </div> */}

      {/* Group 3: NEW */}
     <div className="d-flex flex-column align-items-center" style={{ borderRight: '1px solid #eee', paddingRight: '20px' }}>
      
        <div className="d-flex gap-3">
          <button 
            type="button"
            onClick={() => { console.log("New Request clicked"); }}
            style={{ background: 'none', border: 'none', padding: 0,  marginTop:'0px', cursor: 'pointer' }}
            title="New Request"
          >   <span className="mb-1 mt-2" data-tooltip="New Request">
            <img src={require("../assets/newr.png")} alt="New Request" />
            </span>
          </button>
          <Dropdown as={ButtonGroup} style={{ marginTop: '0px' }}>
  <Dropdown.Toggle
    style={{ padding: '0px 0px', background:'none' }}
   
    className="mt-0"  id="dropdown-template"
  >
   
      <span className="mb-1 mt-2" data-tooltip="Select Template">
      <img
        src={require("../assets/newt.png")}
        alt="Select Template"
      
      />
    </span>
  </Dropdown.Toggle>

    <Dropdown.Menu className="dropdown-menu-start newtheme font-14">
    {navItems.map((item, index) => (
      <Dropdown.Item
        key={index}
        onClick={() => {
          setSelectedTemplate(item);
          setShowTemplateForm(true);
        }}
      >
        {item.DocumentCategory}
      </Dropdown.Item>
    ))}
  </Dropdown.Menu>
</Dropdown>
          
        </div>
        

          <span style={{ fontSize: '14px',  color: '#333',  marginTop: '5px' }}>New</span>
      </div>

      {/* Group 4: ACTION */}
      <div className="d-flex flex-column align-items-center" style={{ borderRight: '1px solid #eee', paddingRight: '20px' }}>
      
        <div className="d-flex gap-3">
          <button 
            type="button"
            onClick={() => { console.log("Share clicked"); }}
            style={{ background: 'none', border: 'none', padding: 0, marginTop:'0px', cursor: 'pointer' }}
            title="Share"
          ><span className="mb-1 mt-2" data-tooltip="Share">
            <img src={require("../assets/listiconshare.png")} alt="Share"  />
            </span>
          </button>

          <button 
            type="button"
            onClick={() => { console.log("Delete clicked"); }}
            style={{ background: 'none', border: 'none', padding: 0, marginTop:'0px', cursor: 'pointer' }}
            title="Delete"
          ><span className="mb-1 mt-2" data-tooltip="Delete">
            <img src={require("../assets/listicond.png")} alt="Delete" />
            </span>
          </button>
        </div>
          <span style={{ fontSize: '14px',color: '#333', marginTop: '5px' }}>Action</span>
      </div>

        <div className="d-flex flex-column align-items-center">
     
        <div className="">
          <button 
            type="button"
            onClick={() => setActiveLayout('grid')}
            title="Grid View"
            style={{ 
              background: activeLayout === 'grid' ? '#fff' : 'transparent', 
              border: 'none', 
              boxShadow: activeLayout === 'grid' ? '0 0px 0px rgba(0,0,0,0.1)' : 'none',
              marginTop:'0px',
             
              cursor: 'pointer'
            }}
          ><span className="mb-1 mt-2" data-tooltip="Grid View">
            <img src={require("../assets/gridview.png")} alt="Grid" style={{  opacity: activeLayout === 'grid' ? 1 : 0.6 }} />
        </span>  </button>
          <button 
            type="button"
            onClick={() => setActiveLayout('list')}
            title="List View"
            style={{ 
              background: activeLayout === 'list' ? '#fff' : 'transparent', 
              border: 'none', 
              boxShadow: activeLayout === 'list' ? '0 0px 0px rgba(0,0,0,0.1)' : 'none',
          marginTop:'0px',
              
              cursor: 'pointer'
            }}
          >
            <span className="mb-1 mt-2" data-tooltip="List View">
            <img src={require("../assets/listview.png")} alt="List" style={{  opacity: activeLayout === 'list' ? 1 : 0.6 }} />
     </span>     </button>
        </div>
           <span style={{ fontSize: '14px',color: '#333', marginTop: '-1px' }}>View</span>
      </div>

    </div>

    {/* View Switcher (Right Side) - Updated with Icons and "View" Label */}
   
  </div>
</div>
             
          {/* <div  className="main-content-new">
            <div className="inbox-leftbar"> */}
            {/* Ritik 13/4/26 */}
                 <div className="main-content-new" >
                  <div className="inbox-leftbar" style={{overflowY:'auto', height:'100%'}}>
              {/* Quick Views Panel */}
              <div
                id="buttonpanel"
                style={{
                


                  flexShrink: 0,
                }}
              >
            
              {[
                  { originalKey: "My request", label: "My Requests", icon: faCloud },
                  { originalKey: "My favourite", label: "My Favourite", icon: faHeart },
                  { originalKey: "My Folders", label: "My Folders", icon: faFolder },
                  { originalKey: "Share with me", label: "Shared with Me", icon: faShareAlt },
                  { originalKey: "Share with other", label: "Shared with Other", icon: faPaperPlane},
                  { originalKey: "Recycle bin", label: "Recycle Bin", icon: faTrashAlt },
                  // srs 6/3/26
                  { originalKey: "Archived Files", label: "Archived Files", icon: faArchive },
                ].map((item,view) => (
                  <button type="button" className="buttonaligntext"
                    key={item.originalKey}
                    onClick={() => handleViewButtonClick(item.originalKey)}
                    style={{
                      width: "100%",
                   backgroundColor: activeView === item.originalKey ? "#2c9942" : "#fff ",
                      color: activeView === item.originalKey ? "white" : "#2c3e50",
                      border: "0px solid #ccc",
                      borderRadius: "8px",
                      cursor: "pointer",
                  
                      transition: "all 0.2s",
                    }}
                  >
                   {item.icon && <FontAwesomeIcon icon={item.icon} style={{ marginRight: "8px" }} />}
                    {item.label} <span className="doc-count"> {viewCounts[item.originalKey] || 0}</span>
                  </button>
                ))}
              </div>

              {/* Folder Hierarchy Panel */}
              <div
                id="folderhierarchycontainer"
                style={{
                  flexGrow: 1,
                  padding: "0px 0px",
                  overflow: "auto",

                }}
              >
                <h2
                  style={{
                    fontSize: "16px",
                    fontWeight: "600",
                   
                    color: "#333",
                         paddingTop: "10px",
                    paddingBottom: "5px",
                    borderTop: "1px solid #eee",
                  }}
                >
                  Folder Hierarchy
                </h2>
                {treeData.length > 0 ? (
                  renderTree(treeData)
                ) : (
                  <p>Loading folder structure...</p>
                )}
              </div>
            </div>

            {/* File List Panel  11 */}
            <div
  id="filelistcontainer"
  className="inbox-rightbar"
  // style={{ position: 'relative', overflow: 'hidden' }}>
  // Ritik 13/4/26
    style={{ position: 'relative',  flex: 1 }}>  
              {showTemplateForm && (
                <TemplateForm
                  selectedTemplate={selectedTemplate}
                  onSave={(data) => handleSaveTemplate(data)}
                  onCancel={handleCancelTemplate}
                  context={context}
                />
              )}
              <div style={{ display: showTemplateForm ? 'none' : 'block' }}>
              {/* ===== FILE LIST / UI ===== */}
              {/* srs 29/1/26 */}
             {/* --- RIBBON TOOLBAR START: DMSMAIN STYLE --- */}
{/* --- ICON-BASED RIBBON: NO DUPLICATION & NO IMPORT ERROR --- */}
{/* --- ICON-BASED RIBBON: DMSMAIN STYLE WITH EXACT PATH --- */}

{/* --- RIBBON TOOLBAR END --- */}

                  <h2
                    style={{
                      fontSize: "16px",
                      fontWeight: "600",
                      marginBottom: "0px",
                      paddingTop:'18px',
                      paddingLeft:'10px',
                      color: "#333",
                      paddingBottom: "4px",
                      borderBottom: "0px solid #eee",
                      paddingRight: "100px", // Make space for upload button
                    }}
                  >
                    {breadcrumbs.length > 0 ? (
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          flexWrap: "wrap",
                        }}
                      >
                        {breadcrumbs.map((item, index) => (
                          <React.Fragment key={item.key}>
                            {index > 0 && (
                              <span
                                style={{
                                  margin: "0 8px",
                                  color: "#999",
                                  fontSize: "14px",
                                }}
                              >
                                ›
                              </span>
                            )}
                            <span
                              className="breadcrumb-item mb-0 font-18 fw-bold text-dark header-title"
                              onClick={() => handleBreadcrumbClick(item)}
                              style={{
                                cursor: "pointer",
                                color: index === breadcrumbs.length - 1 ? "#333" : "#0066cc",
                                fontWeight: index === breadcrumbs.length - 1 ? "600" : "normal",
                                fontSize: "14px",
                                padding: "2px 5px",
                                borderRadius: "3px",
                              }}
                            >
                              {
  item.title === "My request"
    ? "My Requests"
    : item.title === "My favourite"
    ? "My Favourite"
    : item.title === "Share with me"
    ? "Shared with Me"
    : item.title === "Share with other"
    ? "Shared with Other"
    : item.title === "Recycle bin"
    ? "Recycle Bin"
    : item.title
}
                              
                            </span>
                          </React.Fragment>
                        ))}
                         {/* Aman 13/4/26 */}
                          {breadcrumbs.length > 2 && selectedCurrentNode && (
    <FontAwesomeIcon
      icon={faShareAlt}
      title="Share URL Path"
      style={{
        marginLeft: "8px",
        cursor: "pointer",
        fontSize: "13px",
        color: "#666",
      }}
     onClick={() => {
  handleBreadcrumbShareClick();
}}
    />
  )}
                      </div>
                    ) : (
                      <div style={{ color: "#666", fontSize: "14px" }}>Select a folder to view files</div>
                    )}
                  </h2>

 {activeComponent ? (
    <div className="create-folder-section" style={{ padding: "8px 0px", background: "#fff", borderRadius: "8px" }}>
         <div style={{ marginTop: "-45px", float:'right', width:'100%',textAlign:"right" }}>
        <button style={{background:'#fff'}} className="me-2 mt-0 " id="CreateFolderInsideSharePoint">
          <span className="mb-1 mt-2" data-tooltip="Return">
 <img src={require('../assets/backn.png')} alt="Create" onClick={() => setActiveComponent(false)}
/>
            </span></button>
        
 
      </div>
     
      {/* --- Path Display --- */}
      <div style={{ margin: "0px 17px 17px 17px", padding: "10px", gap:'10px', backgroundColor: "#f9f9f9", display:'flex', justifyContent:'left', alignItems:'center', borderLeft: "4px solid #0078d4" }}>
        <h4 className="font-16 text-dark fw-bold m-0 mb-0">This Folder will create under :</h4>
        <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap", fontSize: "14px", color: "#333" }}>
          {breadcrumbs.length > 0 ? (
            breadcrumbs.map((item, index) => (
              <React.Fragment key={item.key}>
                {index > 0 && <span style={{ margin: "0 8px", color: "#999" }}>›</span>}
                <span style={{ fontWeight: index === breadcrumbs.length - 1 ? "600" : "normal" }}>
                  {item.title}
                </span>
              </React.Fragment>
            ))
          ) : (
            <span>Root</span>
          )}
        </div>
      </div>

      {/* --- Create Folder Component --- */}
      <CreateFolder 
        OthProps={{
          "Entity": `${currentSiteUrl.split("/sites/")[1]?.split("/")[1] || ""}`,
          "Entityurl": currentSiteUrl,
          "siteID": currentSiteUrl,
          "Devision": "",
          "Department": "",
          "DocumentLibrary": selectedCurrentNode?.type === "site" || selectedCurrentNode?.type === "subsite" ? "" : selectedCurrentNode?.title,
          "Folder": currentFolderPath,
          "folderpath": `/sites/${currentSiteUrl.split("/sites/")[1] || ""}/${currentFolderPath}`,
          "IsFolderDeligationUser": "false",
        }}
        context={context}
        onCloseForm={() => setActiveComponent(false)} // aman 20/04/26 added onCloseForm to reset activeComponent state when folder creation is done or cancelled
      />

      {/* --- Back Button --- */}
  
    </div>
  ) : showUploadPanel ? (
    /* --- SECTION 2: UPLOAD FILE VIEW (Hides everything else) --- */
    <UploadFile
    // srs 30/1/26
    sp={sp}
      currentfolderpath={{
        "Entity": `${currentSiteUrl.split("/sites/")[1]?.split("/")[1] || ""}`,
        "Entityurl": currentSiteUrl,
        // "siteID": currentSiteUrl,
        // srs 19/2/26
        "siteID": context.pageContext.site.id.toString(),
        "Devision": "",
        "Department": "",
        "DocumentLibrary": selectedCurrentNode?.type === "site" || selectedCurrentNode?.type === "subsite" ? "" : selectedCurrentNode?.title,
        "Folder": currentFolderPath,
        "folderpath": `/sites/${currentSiteUrl.split("/sites/")[1] || ""}/${currentFolderPath}`,
        "IsFolderDeligationUser": "false",
      }}
      onReturnToMain={() => setShowUploadPanel(false)} 
    />
  ) : (
    <>
      {/* srs 29/1/26 */}
              {!showPreviewModal && (
                <>

                  {/* Upload File Button - Only shown when in a folder/library */}


                  {/* sourish 3/10/25 */}
                  {/* srs 29/1/26 moved the whole part on top*/}

                  {/* Upload Panel */}
                   {/* srs 29/1/26 moved the whole part on top*/}
                  {/* {showUploadPanel && (
                    <div
                      style={{
                        marginBottom: "20px",
                        padding: "15px",
                        backgroundColor: "#f5f5f5",
                        borderRadius: "4px",
                        border: "1px solid #ddd",
                      }}
                    >
                      <h3 style={{ fontSize: "14px", marginBottom: "10px" }}>Upload Files</h3>
                      <input type="file" ref={fileInputRef} onChange={handleFileSelect} multiple style={{ display: "none" }} />
                      <button type="button"
                        onClick={triggerFileInput}
                        style={{
                          padding: "8px 15px",
                          backgroundColor: "#0078d4",
                          color: "white",
                          border: "none",
                          borderRadius: "4px",
                          cursor: "pointer",
                          marginRight: "10px",
                        }}
                      >
                        Choose Files
                      </button>

                      {selectedUploadFiles.length > 0 && (
                        <div style={{ marginTop: "10px" }}>
                          <h4 style={{ fontSize: "13px", marginBottom: "5px" }}>Selected Files:</h4>
                          <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
                            {selectedUploadFiles.map((file, index) => (
                              <li
                                key={index}
                                style={{
                                  padding: "5px 0",
                                  fontSize: "13px",
                                  display: "flex",
                                  alignItems: "center",
                                }}
                              >
                                <span style={{ marginRight: "10px" }}>📄</span>
                                <div>
                                  <div>{file.name}</div>
                                  {previewFileUrls[index] && (
                                    <div
                                      style={{
                                        fontSize: "11px",
                                        color: "#666",
                                        marginTop: "2px",
                                      }}
                                    >
                                      Preview:{" "}
                                      <a href={previewFileUrls[index]} target="_blank" rel="noopener noreferrer">
                                        View File
                                      </a>
                                    </div>
                                  )}
                                </div>
                              </li>
                            ))}
                          </ul>
                          <button type="button"
                            onClick={uploadFiles}
                            style={{
                              padding: "8px 15px",
                              backgroundColor: "#107c10",
                              color: "white",
                              border: "none",
                              borderRadius: "4px",
                              cursor: "pointer",
                              marginTop: "10px",
                            }}
                          >
                            Upload Files to Destination
                          </button>
                          {uploadProgress > 0 && (
                            <div style={{ marginTop: "10px" }}>
                              <div
                                style={{
                                  width: "100%",
                                  backgroundColor: "#e0e0e0",
                                  borderRadius: "4px",
                                  height: "20px",
                                }}
                              >
                                <div
                                  style={{
                                    width: `${uploadProgress}%`,
                                    backgroundColor: "#0078d4",
                                    height: "100%",
                                    borderRadius: "4px",
                                    transition: "width 0.3s",
                                  }}
                                />
                              </div>
                              <div
                                style={{
                                  textAlign: "center",
                                  marginTop: "5px",
                                  fontSize: "12px",
                                }}
                              >
                                {Math.round(uploadProgress)}% Complete
                              </div>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  )} */}


                  {/* File List - edited by ritik - 13/01/2026 - fixed Whenever we searched, the page would disappear if the file was not found.*/}
                  {!showUploadPanel && (
                    activeView === "browse" ? (
                      <>
                        {/* Search bar for Browse View */}
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            width: "260px",
                            border: "1px solid #d1d1d1",
                            borderRadius: "18px",
                            padding: "6px 10px",
                            backgroundColor: "#fff",
                            marginBottom: "12px",
                          }}
                        >
                          {/* 🔍 search icon  ritik 19/01/2026*/}
                          <span
                            onClick={() => {
                              setSearchTerm(searchInput.trim()); // ✅ search yahan trigger hoga
                              setColumnSearch({ name: '', size: '', library: '', status: '' });  //Rohit 22/04/2026
                              setCurrentPage(1);
                            }}
                            style={{
                              cursor: "pointer",
                              color: "#666",
                              fontSize: "14px",
                              marginRight: "6px",
                            }}
                          >
                            🔍
                          </span>

                          {/* ritik 19/01/2026 - search input box */}
                          <input
                            type="text"
                            placeholder="Search"
                            value={searchInput}
                            onChange={(e) => setSearchInput(e.target.value)}
                            style={{
                              border: "none",
                              outline: "none",
                              flex: 1,
                              fontSize: "14px",
                              background: "transparent",
                            }}
                          />


                          {/* ✖ clear button fix Ritik 19/01/2026 */}


                          {searchInput && (
                            <span
                              onClick={() => {
                                setSearchInput("");
                                setSearchTerm("");
                                setCurrentPage(1);
                              }}
                              style={{
                                cursor: "pointer",
                                color: "#999",
                                fontSize: "14px",
                                marginLeft: "6px",
                              }}
                            >
                              ✖
                            </span>
                          )}

                        </div>



                        {selectedFiles.length > 0 ? (
                          <>
                            {(() => {
                              // Browse View Filter
                              const browseFiltered = selectedFiles.filter((file) => {
                                if (!searchTerm.trim()) return true;
                                const term = searchTerm.toLowerCase();
                                return (
                                  (file.Name || "").toLowerCase().includes(term) ||
                                  (file.FileLeafRef || "").toLowerCase().includes(term) ||
                                  (file.ServerRelativeUrl || "").toLowerCase().includes(term)
                                );
                              });

                              const start = (currentPage - 1) * pageSize;
                              const end = start + pageSize;
                              const paginatedBrowseFiles = browseFiltered.slice(start, end);

                              return browseFiltered.length > 0 ? (
                                <>
                                  <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
                                    {paginatedBrowseFiles.map((file, idx) => (
                                      <li
                                        key={file.Name || file.UniqueId || idx}
                                        style={{
                                          padding: "8px 0",
                                          borderBottom: "1px solid #eee",
                                          display: "flex",
                                          alignItems: "center",
                                        }}
                                      >
                                        <div
                                          style={{ marginRight: "8px" }}
                                          dangerouslySetInnerHTML={{
                                            __html: createFileExtensionHtml(
                                              getActualFileName(file)
                                            ),
                                          }}
                                        />
                                        <span style={{ fontSize: "14px" }}>
                                          {file.Name || file.FileLeafRef}
                                        </span>
                                      </li>
                                    ))}
                                  </ul>

                                  {/* //rohit 23/04/2026 ---start */}
                                  {browseFiltered.length > pageSize && (
  <div
    style={{
      margin: "10px 0",
      display: "flex",
      alignItems: "center",
      justifyContent: "end",
      gap: "6px",
      flexWrap: "wrap",
    }}
  >
    <button
      type="button"
      onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
      disabled={currentPage === 1}
      style={{
        marginRight: "4px",
        padding: "5px 12px",
        marginTop: "0px",
        borderRadius: "4px",
        border: "1px solid #ccc",
        background: currentPage === 1 ? "#eee" : "#1d4ed8",
        color: currentPage === 1 ? "#666" : "#fff",
        cursor: currentPage === 1 ? "not-allowed" : "pointer",
      }}
    >
      Previous
    </button>

    {/* Page numbers (existing pill logic) */}
    {/* ... keep the existing page number generation code ... */}

    <button
      type="button"
      onClick={() =>
        setCurrentPage((p) =>
          p < Math.ceil(browseFiltered.length / pageSize) ? p + 1 : p
        )
      }
      disabled={currentPage === Math.ceil(browseFiltered.length / pageSize)}
      style={{
        marginLeft: "4px",
        padding: "5px 12px",
        marginTop: "0px",
        borderRadius: "4px",
        border: "1px solid #ccc",
        background: currentPage === Math.ceil(browseFiltered.length / pageSize) ? "#eee" : "#1d4ed8",
        color: currentPage === Math.ceil(browseFiltered.length / pageSize) ? "#666" : "#fff",
        cursor: currentPage === Math.ceil(browseFiltered.length / pageSize) ? "not-allowed" : "pointer",
      }}
    >
      Next
    </button>

    <select
      value={pageSize}
      onChange={(e) => {
        setPageSize(Number(e.target.value));
        setCurrentPage(1);
      }}
      style={{
        padding: "4px 8px",
        borderRadius: "4px",
        border: "1px solid #ccc",
        fontSize: "13px",
        cursor: "pointer",
        marginLeft: "4px",
        height: "30px",
      }}
    >
      {[10, 20, 40, 80, 100].map((n) => (
        <option key={n} value={n}></option>
      ))}
    </select>
      {/* //Rohit 23/04/2026 ---remove the total numbers*/}
    {/* <span style={{ fontSize: "13px", color: "#555", marginLeft: "4px" }}>
      {browseFiltered.length} records
    </span> */}   
  </div>
)}
                                  {/* //rohit 23/04/2026 ---end */}
  <div
    style={{
      margin: "10px 0",
      display: "flex",
      alignItems: "center",
      justifyContent: "end",
      gap: "6px",
      flexWrap: "wrap",
    }}
  >
    {/* Records per page dropdown - Rohit 23/04/2026 */}

    <button
      type="button"
      onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
      disabled={currentPage === 1}
      style={{
        marginRight: "4px",
        padding: "5px 12px", marginTop: '0px',
        borderRadius: "4px",
        border: "1px solid #ccc",
        background: currentPage === 1 ? "#eee" : "#1d4ed8",
        cursor: currentPage === 1 ? "not-allowed" : "pointer",
      }}
    >
      Prev
    </button>
    

    {/* Page number pills */}
    {(() => {
      const total = Math.ceil(browseFiltered.length / pageSize) || 1;
      const pages: any[] = [];
      let start = Math.max(1, currentPage - 2);
      let end = Math.min(total, start + 4);
      if (end - start < 4) start = Math.max(1, end - 4);

      if (start > 1) {
        pages.push(
          <button key="p1" type="button"
            onClick={() => setCurrentPage(1)}
            style={{ padding: "4px 9px", marginTop: "0px", borderRadius: "4px", border: "1px solid #ccc", background: "#fff", color: "#333", cursor: "pointer", fontSize: "13px" }}>
            1
          </button>
        );
        if (start > 2) pages.push(<span key="e1" style={{ padding: "0 2px", fontSize: "13px", color: "#555" }}>…</span>);
      }

      for (let p = start; p <= end; p++) {
        pages.push(
          <button key={p} type="button"
            onClick={() => setCurrentPage(p)}
            style={{
              padding: "4px 9px", marginTop: "0px",
              borderRadius: "4px",
              border: "1px solid #ccc",
              background: currentPage === p ? "#1d4ed8" : "#fff",
              color: currentPage === p ? "#fff" : "#333",
              cursor: "pointer",
              fontWeight: currentPage === p ? "600" : "400",
              fontSize: "13px",
            }}>
            {p}
          </button>
        );
      }

      if (end < total) {
        if (end < total - 1) pages.push(<span key="e2" style={{ padding: "0 2px", fontSize: "13px", color: "#555" }}>…</span>);
        pages.push(
          <button key={`last${total}`} type="button"
            onClick={() => setCurrentPage(total)}
            style={{ padding: "4px 9px", marginTop: "0px", borderRadius: "4px", border: "1px solid #ccc", background: "#fff", color: "#1d4ed8", cursor: "pointer", fontSize: "13px" }}>
            {total}
          </button>
        );
      }
      return pages;
    })()}

    <button
      type="button"
      onClick={(e) => {
        e.preventDefault();
        setCurrentPage((p) =>
          p < Math.ceil(browseFiltered.length / pageSize) ? p + 1 : p
        );
      }}
      disabled={currentPage === Math.ceil(browseFiltered.length / pageSize)}
      style={{
        marginLeft: "4px",
        padding: "5px 12px", marginTop: '0px',
        borderRadius: "4px",
        border: "1px solid #ccc",
        background: currentPage === Math.ceil(browseFiltered.length / pageSize) ? "#eee" : "#1d4ed8",
        cursor: currentPage === Math.ceil(browseFiltered.length / pageSize) ? "not-allowed" : "pointer",
      }}
    >
      Next
    </button>
    <select
      value={pageSize}
      onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(1); }}
      style={{
        padding: "4px 8px",
        borderRadius: "4px",
        border: "1px solid #ccc",
        fontSize: "13px",
        cursor: "pointer",
        marginRight: "4px",
        height: "30px",
      }}
    >
      {[10, 20, 40, 80, 100].map((n) => (
        <option key={n} value={n}>{n}</option>
      ))}
    </select>
  </div>


                                  {/* //rohit 23/04/2026 -----end */}
                                </>
                              ) : (
                                <div
                                  style={{
                                    color: "#666",
                                    fontSize: "14px",
                                    padding: "20px",
                                    textAlign: "center",
                                    marginTop: "20px",
                                  }}
                                >
                                  {searchTerm.trim() !== ""
                                    ? `No files found matching "${searchTerm}"`
                                    : <div style={{display:'flex', alignItems:'center', justifyContent:'center', flexDirection:'column'}}> <img src={require('../assets/no-filef.png')}></img>  <p className="mt-2 font-14 fw-bold text-dark">No files found in this view</p> </div>}
                                </div>
                              );
                            })()}
                          </>
                        ) : (
                          <div
                            style={{
                              color: "#666",
                              fontSize: "14px",
                              padding: "20px",
                              textAlign: "center",
                              marginTop: "20px",
                            }}
                          >
                            No files in this folder
                          </div>
                        )}
                      </>
                    ) : null
                  )}

                  {/* Search bar for Quick Views */}
                  {!showUploadPanel && activeView !== "browse" && (
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "flex-end",
                        marginBottom: "12px",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          width: "300px",
                          border: "1px solid #dcdcdc",
                          borderRadius: "22px",
                          padding: "7px 12px",
                          position:'absolute',top:'8px',
                          backgroundColor: "#fff", right:'12px'
                        }}
                      >
                        <input
                          type="text"
                          placeholder="Search files..."
                          value={searchInput} className="searc_global"
                          onChange={(e) => setSearchInput(e.target.value)}
                          style={{
                            border: "none",
                            outline: "none",
                            flex: 1,
                            fontSize: "14px",
                            background: "transparent",
                          }}
                        />

                        {/* ✖ clear */}
                        {searchInput && (
                          <span
                            onClick={() => {
                              setSearchInput("");
                              setSearchTerm("");
                              setCurrentPage(1);
                            }}
                            style={{
                              cursor: "pointer",
                              color: "#999",
                              marginRight: "8px",
                              fontSize: "14px",
                            }}
                          >
                            ✖
                          </span>
                        )}

                        {/* 🔍 search icon */}
                        <span
                          onClick={() => {
                            setSearchTerm(searchInput.trim()); // ✅ SEARCH TRIGGER
                            setColumnSearch({ name: '', size: '', library: '', status: '' });  // Rohit 22/04/2026
                            setCurrentPage(1);
                          }}
                          style={{
                            color: "#666",
                            fontSize: "15px",
                            cursor: "pointer",
                          }}
                        >
                          🔍
                        </span>
                      </div>
                    </div>
                  )}

                  {/* Rohit 22/04/2026 ------------------start */}
                  {/* Quick Views Grid/List */}
                  {!showUploadPanel && activeView !== "browse" && (
                    <>
                      {paginatedFiles.length > 0 ? (
                        <>
                          {/* Grid View */}
                          {activeLayout === "grid" && (
                            <div className="layoutdesign">
                              {filesLoadedfromnode && paginatedFiles.length > 0 ? (
                                <>
                                
                                {paginatedFiles.map((file, idx) => (
                                  <div
                                    key={getQuickViewRenderKey(file, idx)}
                                    className="carddesign"
                                  >
                                    {/* Rohit 22/04/2026 --------------end*/}
                                    {/* File Content */}


                                    {/* File Extension Badge */}
                                    <div
                                      dangerouslySetInnerHTML={{
                                        __html: createFileExtensionHtml(getActualFileName(file)),
                                      }}
                                    />

                                    {/* File Name */}
                                    <div className="text-dark fw-bold1 font-16 onelinetrim ">
                                      {getActualFileName(file)}
                                    </div>







                                    {/* <div className="text-muted font-12">
                                      {file.Length
                                        ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB`
                                        : ""}
                                    </div> */}
                                    <div className="text-muted font-12">
  {file.Length
    ? (() => {
        const sizeInBytes = parseInt(file.Length);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
</div>
                                    <div className="text-muted font-12">
                                      {file.TimeCreated
                                        ? new Date(file.TimeCreated).toLocaleDateString("en-GB")
                                        : ""}
                                    </div>

                                    {/* Three-dot menu button */}
                                    <div className="dotbutton">
                                      <button className="action-btn2" type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          setMenuOpenIdx(menuOpenIdx === idx ? null : idx);
                                        }}
                                       
                                      >
                                        ⋮
                                      </button>

                                      {/* Context-specific menu */}
                                      {menuOpenIdx === idx && (
                                        <div
                                          style={{
                                            position: "absolute",
                                            right: 0,
                                            top: "28px",
                                            background: "#fff",
                                            border: "1px solid #ddd",
                                            borderRadius: "4px",
                                            boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
                                            zIndex: 10,
                                            minWidth: "160px",
                                          }}
                                        >
                                          <ul className="internalbutton">
                                            {/* Common actions for all views */}
                                            <li>
                                              <button type="button" className="newbuttontext"

                                                onClick={() => {
                                                  setPreviewFile(file);
                                                  setShowPreviewModal(true);
                                                }}
                                              >
                                                <span> <FontAwesomeIcon icon={faEye} /> </span> Preview File
                                              </button>
                                            </li>
{/* // Ritik 16/4/26 */}
                                              <li>
  <button type="button" className="newbuttontext"
    onClick={() => {
      setMenuOpenIdx(null);
      setShareFile(file);
      setShowShareModal(true);
    }}
  >
    <span><FontAwesomeIcon icon={faShareAlt} /></span> Share
  </button>
</li>
                                            {/* // addhyan 24/2/26 */}
                                            <li>
                                              <button type="button"
                                                className="newbuttontext"
                                                onClick={async () => {
                                                  await ArchivedFile(
                                                    file,
                                                    currentSiteUrl,
                                                    context
                                                  );
                                                  setMenuOpenIdx(null);
                                                }}
 
                                              >
                                                <span><FontAwesomeIcon icon={faTrashAlt} /></span> Archive
                                              </button>
                                            </li>
                                            <li>
                                              <button type="button"
                                                className="newbuttontext"
                                                onClick={() => {
                                                  // eslint-disable-next-line @typescript-eslint/no-floating-promises, @typescript-eslint/no-use-before-define
                                                  handleAuditHistory(file);
                                                  setMenuOpenIdx(null);
                                                }}
                                              >
                                                <span> <FontAwesomeIcon icon={faFileAlt} /> </span> Audit History
                                              </button>
                                            </li>
                                             {/* srs 10/4/26 */}
                                             <li>
  <button 
    type="button"
    className="newbuttontext"
    onClick={() => {
      // Extract the required fields from your file object
      // Note: Ensure your 'file' object has these properties (adjust names if necessary)
      // 1. Get the domain (https://officeindia.sharepoint.com)
      const origin = window.location.origin;

      // 2. Construct the URL by finding the path up to the SiteName
      // This ensures even if you are deep in a folder, the URL stops at the subsite level

      // aman - 29/04/26 start
      // const pathUntilSite = file.CurrentFolderPath.split(file.SiteName)[0];
      const pathRef = file.CurrentFolderPath || file.ServerRelativeUrl || "";
      const pathUntilSite = pathRef.split(file.SiteName)[0];
      const finalSiteUrl = `${origin}${pathUntilSite}${file.SiteName}`;

      console.log("Target Subsite URL:", finalSiteUrl);
      window.ManageFilePermission(
        //  file.FileUID,            // fileId
         file.FileUID || file.UniqueId,
 
        finalSiteUrl,             // siteId (Your data shows this is the URL)
        // file.DocumentLibraryName, // documentLibraryName

        file.DocumentLibraryName || file.__libraryTitle || file.LibraryName,

        // aman - 29/04/26 end 
 
        file.SiteName            // siteTitle (Used for the Admin Group naming)
      );
      setMenuOpenIdx(null);
    }}
  >
    <span> <FontAwesomeIcon icon={faFileAlt} /> </span> Manage File Permission
  </button>
</li>

                                            <li>
                                              <button type="button"
                                                className="newbuttontext"
                                                onClick={() => {
                                                  setModalFile(file);

                                                  setShowVersionModal(true);
                                                  setMenuOpenIdx(null);
                                                }}
                                              >
                                                <span> <FontAwesomeIcon icon={faHistory} /> </span> Version History
                                              </button>
                                            </li>
                                            <li>
                                              <button type="button"
                                                className="newbuttontext"
                                                onClick={async () => {
                                                  await deleteFileFolder(
                                                    file,
                                                    currentSiteUrl,
                                                    context
                                                  );
                                                  setMenuOpenIdx(null);
                                                }}

                                              >
                                                <span> <FontAwesomeIcon icon={faTrashAlt} /> </span> Delete File
                                              </button>
                                            </li>
                                              {/* Addhyan  */}
                                            <li><button type="button" onClick={() => {
                                                    setMenuOpenIdx(null);
                                                    setShareUrlFile(file);
                                                    setShowShareUrlModal(true);
                                                  }}><FontAwesomeIcon icon={faShareAlt} /> Share file url</button></li>
                                            <li>
                                              <button type="button"
                                                className="newbuttontext"
                                                onClick={async () => {
                                                  const newFav = await toggleFavouriteDoc(
                                                    file,
                                                    currentSiteUrl,
                                                    context
                                                  );
                                                  setMenuOpenIdx(null);

                                                  // Update local state to reflect new favourite status
                                                  setSelectedFiles((prev) =>
                                                    prev.map((f) =>
                                                      f.UniqueId === file.UniqueId
                                                        ? { ...f, IsFavourite: newFav }
                                                        : f
                                                    )
                                                  );
                                                }}
                                              >
                                                <span> <FontAwesomeIcon icon={faHearts} />  </span>{" "}
                                                {file.IsFavourite
                                                  ? "Unmark Favourite"
                                                  : "Mark as Favourite"}
                                              </button>
                                            </li>

                                          </ul>
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                  // Rohit 22/04/2026 --------------start
                                ))}
                                  
                                </>
                              ) :
                                // this is my request and my favourite my recucle bin files ternary operator
                                (
                                  <>
                                    
                                  {paginatedFiles.map((file, idx) => (
                                    <div
                                      key={getQuickViewRenderKey(file, idx)}
                                      className="carddesign"

                                    >
                                      {/* Rohit 22/04/2026 end ---------*/}
                                      {/* Card Content */}
                                      {
                                        activeView === "My request" ? (
                                         <>
                                            <div className="d-flex align-items-center gap-2">
                                               <div dangerouslySetInnerHTML={{
                                              __html: createFileExtensionHtml(getActualFileName(file))
                                            }} />
                                            {/* <div className="text-muted font-12"> <FontAwesomeIcon icon={faShoppingBag}  className="me-1" /> {file.FileSize}</div> */}
                                            <div className="text-muted font-12">
  <FontAwesomeIcon icon={faShoppingBag} className="me-1" />

  {file.FileSize
    ? (() => {
        const sizeInBytes = parseInt(file.FileSize);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
</div>
                                            </div>
                                           
 
 
                                            <div className="font-14 fw-bold1 mb-1 mt-1 onelinetrim">{file.FileName}</div>
                                           <div className="d-flex align-items-center justify-content-start gap-3 mb-1">  <div className="text-muted font-12"><FontAwesomeIcon icon={faFolder} />  {file.DocumentLibraryName}</div>
                                            </div>
 
                                            <div style={{borderTop:'1px solid #f1f5f9', paddingTop:'10px'}} className="d-flex align-items-center justify-content-between">
                                            {/* <div className="text-muted font-12 gap-1"> <FontAwesomeIcon icon={faCalendar}  className="me-1"/>
    {new Date(file.Created).toLocaleDateString("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
  })}
</div> */}
<div className="text-muted font-12 gap-1">
  <FontAwesomeIcon icon={faCalendar} className="me-1" />
  
  {new Date(file.Created).toLocaleDateString("en-GB")}
</div>
                                            {file.Status=== "Pending" && (
                                              <div className=" font-12 status_new" style={{ backgroundColor: '#fffbeb', color:'#b45309', border:'1px solid #fde68a' }}><FontAwesomeIcon icon={faClock} /> {file.Status}</div>  
                                              )
                                              
                                            }
                                            {file.Status=== "Auto Approved" && (
                                              <div className="text-success font-12 status_new" style={{ backgroundColor: '#f0fdf4', color:'#15803d', border:'1px solid #bbf7d0' }}><FontAwesomeIcon icon={faCheckCircle} /> {file.Status}</div>  
                                              )
                                              
                                            }
                                            {file.Status=== "Rejected" && (
                                              <div className="text-danger font-12 status_new" style={{ backgroundColor: '#FDECEA', color:'#b45309', border:'1px solid #fde68a' }}><FontAwesomeIcon icon={faRedo} /> {file.Status}</div>  
                                              )
                                              
                                            }
                                             {file.Status=== "Rework" && (
                                              <div className="text-danger font-12 status_new" style={{ backgroundColor: '#fef2f2', color:'#b91c1c', border:'1px solid #fecaca' }}><FontAwesomeIcon icon={faRedo} /> {file.Status}</div>  
                                              )
                                              
                                            }
                                            {file.Status=== "Approved" && (
                                              <div className="text-success font-12 status_new" style={{ backgroundColor: '#f0fdf4', color:'#15803d', border:'1px solid #bbf7d0' }}><FontAwesomeIcon icon={faCheckCircle} /> {file.Status}</div>  
                                              )
                                              
                                            }  </div>
                                            {/* <div className="text-muted font-12">{file.Status}</div> */}
                                          </>
                                        ) : activeView === "My favourite" ? (
                                          <>
                                            <div dangerouslySetInnerHTML={{
                                              __html: createFileExtensionHtml(getActualFileName(file))
                                            }} />


                                            <div style={{ fontWeight: 600, fontSize: "15px" }} className="onelinetrim">{file.FileName}</div>
                                            {/* <div className="text-muted font-12">{file.FileSize}</div> */}
                                            <div className="text-muted font-12">
  <FontAwesomeIcon icon={faShoppingBag} className="me-1" />

  {file.FileSize
    ? (() => {
        const sizeInBytes = parseInt(file.FileSize);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
</div>
                                          </>
                                        ) : activeView === "My Folders" ? (
                                         <>
                                         
                                          {file.IsLibrary ? <div style={{ fontWeight: 600, fontSize: "15px", display:'flex', alignItems:'center',gap:'10px' }}>
                                           <span className="folder-icon"><FontAwesomeIcon icon={faFolders} /> </span> <div> {file.DisplayName} <div style={{fontWeight:'400'}} className="font-12 text-muted  mt-0"> 
                                            <FontAwesomeIcon icon={faCalendars} /> {new Date(file.Created).toLocaleDateString("en-GB")}</div></div></div> : <div style={{ fontWeight: 600, fontSize: "15px", display:'flex', alignItems:'center',gap:'10px'  }}> <span className="folder-icon"><FontAwesomeIcon icon={faFolders} /> </span> <div>  {file.FolderName}   <div style={{fontWeight:'400'}} className="font-12 text-muted  mt-0"> 
                                              <FontAwesomeIcon icon={faCalendars} /> {new Date(file.Created).toLocaleDateString("en-GB")}</div></div> </div>}
                                            {/* <div style={{ fontWeight: 600, fontSize: "20px" }}>📁{file.FolderName}</div> */}
                                            {/* <div style={{ fontWeight: 600, fontSize: "20px" }}>📁{file.documentLibraryName}</div> */}
                                        
 
                                        
                                            <div  className="text-muted font-14 mt-1 mb-1">{file.SiteTitle}</div>
                                             <div style={{ fontSize: "14px", display:'flex', alignItems:'center',gap:'10px', borderTop:'1px solid #f1f5f9', paddingTop:'10px',  justifyContent:'space-between' }}>  <div> <b>Type: </b> {file.IsFolder ? "subFolder" : "Root Folder"}</div>
                                            <div className={file.IsPrivate ? "text-danger badge1" : "text-success badge1"} style={{
    backgroundColor: file.IsPrivate ? "#FDECEA" : "#E6F4EA" }} ><FontAwesomeIcon
                                                        icon={file.IsPrivate ? faLock : faGlobeAsia}
                                                         className="me-1"
                                                />{file.IsPrivate ? "Private" : "Public"}</div> </div> 
                                         
                                           
 
                                          </>

                                        ) : activeView === "Share with me" ? (
                                          <>
                                            <div dangerouslySetInnerHTML={{
                                              __html: createFileExtensionHtml(getActualFileName(file))
                                            }} />

                                            <div style={{ fontWeight: 600, fontSize: "15px" }} className="onelinetrim">{file.FileName}</div>
                                           <div className="text-muted font-12">
  <FontAwesomeIcon icon={faShoppingBag} className="me-1" />

  {file.FileSize
    ? (() => {
        const sizeInBytes = parseInt(file.FileSize);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
</div>
                                          </>
                                        ) : activeView === "Share with other" ? (
                                          <>
                                            <div dangerouslySetInnerHTML={{
                                              __html: createFileExtensionHtml(getActualFileName(file))
                                            }} />
                                            <div style={{ fontWeight: 600, fontSize: "15px" }} className="onelinetrim">{file.FileName}</div>
                                            {/* <div className="text-muted font-12">{file.FileSize}</div> */}
                                            <div className="text-muted font-12">
  <FontAwesomeIcon icon={faShoppingBag} className="me-1" />

  {file.FileSize
    ? (() => {
        const sizeInBytes = parseInt(file.FileSize);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
</div>
                                          </>
                                        ) : activeView === "Recycle bin" ? (
                                          <>
                                            <div dangerouslySetInnerHTML={{
                                              __html: createFileExtensionHtml(getActualFileName(file))
                                            }} />
                                            <div style={{ fontWeight: 600, fontSize: "15px" }} className="onelinetrim">{file.FileName}</div>
                                            <div className="text-muted font-12">
  <FontAwesomeIcon icon={faShoppingBag} className="me-1" />

  {file.FileSize
    ? (() => {
        const sizeInBytes = parseInt(file.FileSize);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
</div>
                                          </>
                                        )
                                        // srs 6/3/26 
                                        : activeView === "Archived Files" ? (
                                          <>
                                            <div dangerouslySetInnerHTML={{
                                              __html: createFileExtensionHtml(getActualFileName(file))
                                            }} />
                                            <div style={{ fontWeight: 600, fontSize: "15px" }} className="onelinetrim">{file.FileName}</div>
                                            <div className="text-muted font-12">{file.FileSize}</div>
                                           
                                          </>
                                        ) : null
                                      }

                                      {/* Three-dot Menu */}
                                      <div className="dotbutton">
                                        <button className="" type="button"
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setMenuOpenIdx(menuOpenIdx === idx ? null : idx);
                                          }}
                                      
                                        >
                                          ⋮
                                        </button>

                                        {menuOpenIdx === idx && (
                                          <div

                                            className="three-dot-menu"
                                            style={{
                                              position: "absolute",
                                              right: 0,
                                              top: "30px",
                                              background: "#fff",
                                              border: "1px solid #ddd",
                                              borderRadius: "4px",
                                              boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
                                              zIndex: 10,
                                              minWidth: "180px",
                                            }}
                                          >
                                            <ul className="internalbutton">

                                              {/* Menu for My Request */}
                                              {activeView === "My request" && (
                                                <>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setPreviewFile(file);
                                                    setShowPreviewModal(true);
                                                  }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
                                                  <li><button onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    // eslint-disable-next-line @typescript-eslint/no-floating-promises, @typescript-eslint/no-use-before-define
                                                    handleAuditHistory(file);
                                                    setMenuOpenIdx(null);
                                                  }}><FontAwesomeIcon icon={faFileAlt} /> Audit History</button></li>
                                                   {/* srs 10/4/26 */}
                                                   <li><button onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    // eslint-disable-next-line @typescript-eslint/no-floating-promises, @typescript-eslint/no-use-before-define
                                                    // 1. Get the domain (https://officeindia.sharepoint.com)
      const origin = window.location.origin;

      // 2. Construct the URL by finding the path up to the SiteName
      // This ensures even if you are deep in a folder, the URL stops at the subsite level
      const pathUntilSite = file.CurrentFolderPath.split(file.SiteName)[0];
      const finalSiteUrl = `${origin}${pathUntilSite}${file.SiteName}`;

      console.log("Target Subsite URL:", finalSiteUrl);
                                                     window.ManageFilePermission(
          file.FileUID,            // fileId
        finalSiteUrl,             // siteId (Your data shows this is the URL)
        file.DocumentLibraryName, // documentLibraryName
        file.SiteName            // siteTitle (Used for the Admin Group naming)
      );
                                                    setMenuOpenIdx(null);
                                                  }}><FontAwesomeIcon icon={faFileAlt} /> Manage File Permission</button></li>
                                                  {/* srs 15/4/26 */}
                                                  {file.Status === "Rework" && (
  <li>
  <button 
    type="button"
    className="newbuttontext"
    onClick={() => {
      // Extract the required fields from your file object
      // Note: Ensure your 'file' object has these properties (adjust names if necessary)
       const constructedPath = `${file.CurrentFolderPath || file.ServerRelativeUrl}/${file.FileName || file.Name}`;
      window.rework(
      file.FileUID,          // 1. GUID (e.g. "afaec963...")
  file.SiteID,           // 2. URL (e.g. "https://officeindia...")
  file.DocumentLibraryName, 
  file.SiteName, 
  constructedPath
      );
      setMenuOpenIdx(null);
    }}
  >
    <span> <FontAwesomeIcon icon={faFileAlt} /> </span> Edit File
  </button>
</li>
)}
                                                  <li><button type="button" onClick={() => {
                                                    setMenuOpenIdx(null);
                                                    setShareFile(file);
                                                    setShowShareModal(true);
                                                  }}><FontAwesomeIcon icon={faShareAlt} /> Share</button></li>

                                                  {/* // add Button for share file url modal - Addhyan 13/4/26 start */}
                                                  <li><button type="button" onClick={() => {
                                                    setMenuOpenIdx(null);
                                                    setShareUrlFile(file);
                                                    setShowShareUrlModal(true);
                                                  }}><FontAwesomeIcon icon={faShareAlt} /> Share file url</button></li>
 
                                                   {/* // add button for share file url modal - Addhyan 13/4/26 end */}
                                                  <li><button type="button" onClick={() => {


                                                    setDirectDownloadFile(file); // new state for direct downloader
                                                    // ritik 
                                                    setTimeout(() => setDirectDownloadFile(null), 3000);
                                                    setMenuOpenIdx(null);
                                                  }}><FontAwesomeIcon icon={faDownload} /> Download</button></li>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setModalFile(file);
                                                    setShowVersionModal(true);
                                                    setMenuOpenIdx(null);
                                                  }}><FontAwesomeIcon icon={faHistory} /> Version History</button></li>

                                                  {/* Addhyan 19/2/26 */}
{/* 
                                                  <li>
                                                    <button
                                                      type="button"
                                                      onClick={() => {
                                                        setMenuOpenIdx(null);
                                                        handleRenameFile(file);
                                                      }}
                                                    >
                                                      <FontAwesomeIcon icon={faEdit} /> Rename File
                                                    </button>
                                                  </li> */}


                                                </>
                                              )}

                                              {/* Menu for My Favourite */}
                                              {activeView === "My favourite" && (
                                                <>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setPreviewFile(file);
                                                    setShowPreviewModal(true);
                                                  }}><FontAwesomeIcon icon={faEye} />  Preview File</button></li>
                                                  <li><button type="button" onClick={() => {
                                                    setMenuOpenIdx(null);
                                                    setShareFile(file);
                                                    setShowShareModal(true);
                                                  }}><FontAwesomeIcon icon={faShareAlt} /> Share</button></li>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    // handleUnmarkFavourite(file);
                                                    // eslint-disable-next-line @typescript-eslint/no-floating-promises
                                                    // Addhyan 19/2/26
                                                    // toggleFavourite(file);
                                                    toggleFavourite(file, currentSiteUrl, context)

                                                    setMenuOpenIdx(null);
                                                  }}><FontAwesomeIcon icon={faHearts} /> Unmark as Favourite</button></li>
                                                </>
                                              )}

                                              {/* Menu for My Folders */}
                                              {activeView === "My Folders" && (
                                                <>
                                                  {
                                                    !file.IsFolder && (
                                                      <>
                                                        {/* <li><button type="button" onClick={() => { }}
                                                        >Manage WorkFlow </button></li> */}

                                                        <li>
                                                          <button
                                                            type="button"
                                                            onClick={() => {
                                                              setSelectedWorkflowFolder({
                                                                SiteID: file.SiteID,
                                                                SiteTitle: file.SiteTitle,
                                                                SiteUrl: file.SiteUrl,              // ✅ MUST
                                                                DocumentLibraryName: file.DocumentLibraryName,
                                                                file: file,                         // Add file object
                                                              });
                                                              setShowManageWorkflow(true);
                                                            }}
                                                          >
                                                          <img src={require("../assets/managework.svg")} alt="userlock" />    Manage WorkFlow
                                                          </button>
                                                        </li>

                                                        {/* Metadata Modal Trigger Ritik 29/01/2026 */}

                                                        <li><button type="button" onClick={() => {
                                                          setMenuOpenIdx(null);
                                                          setSelectedFileForMetadata(file);
                                                          setShowMetadataModal(true);
                                                        }}>  <img src={require("../assets/metadata.svg")} alt="userlock" /> Add MetaData</button></li>

                                                         {/* srs 20/2/26 */}
                                                         <li><button type="button" onClick={() => {
                                                          setMenuOpenIdx(null);
                                                          setSelectedFileForRenameMetadata(file);
                                                          setShowRenameMetadataModal(true);
                                                        }}>  <img src={require("../assets/Rename-Meta-Data.svg")} alt="userlock" /> Rename Meta Data</button></li>

                                                      </>
                                                    )



                                                  }








                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setModalFile(file);
                                                    setMenuOpenIdx(null);
                                                    // eslint-disable-next-line @typescript-eslint/no-floating-promises, @typescript-eslint/no-use-before-define
                                                    deleteFolder(file); // sourish 20/8/25
                                                  }}
                                                  ><img src={require("../assets/deletefile.svg")} alt="deletefile" /> Delete Folder</button></li>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setModalFile(file);
                                                    setRenameValue(file?.FolderName || ""); // prefill with old name
                                                    setRenameModalOpen(true);
                                                    setMenuOpenIdx(null);
                                                  }}> <img src={require("../assets/rename.svg")} alt="rename" /> Rename Folder</button></li>
                                                  <li>
                                                    <button type="button" onClick={() => {
                                                      const pageBefore = currentPage;
                                                      setMenuOpenIdx(null);
                                                      window.managePermission?.(file);
                                                      setCurrentPage(pageBefore);
                                                    }}>
                                                      <img src={require("../assets/userlock.svg")} alt="userlock" />  Manage Permission
                                                    </button>
                                                  </li>
                                                  {/* Aman 13/4/26 */}
                                                   <li>
  <button
    type="button"
    onClick={() => {
      setMenuOpenIdx(null);
      handleFolderShareClick(file);
    }}
  >
    <FontAwesomeIcon icon={faShareAlt} /> Share Folder Path
  </button>
</li>
                                                </>
                                              )}

                                              {/* Menu for Share with me */}
                                              {activeView === "Share with me" && (
                                                <>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setPreviewFile(file);
                                                    setShowPreviewModal(true);
                                                  }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)

                                                    setDirectDownloadFile(file); // new state for direct downloader
                                                    // ritik 
                                                    setTimeout(() => setDirectDownloadFile(null), 3000);
                                                    setMenuOpenIdx(null);
                                                  }}><FontAwesomeIcon icon={faDownload} /> Download File</button></li>
                                                </>
                                              )}

                                              {/* Menu for Share with other */}
                                              {activeView === "Share with other" && (
                                                <>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setPreviewFile(file);
                                                    setShowPreviewModal(true);
                                                  }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
                                                  <li><Button
                                                    variant="warning"
                                                    onClick={() => {
                                                      setShowRevokeModal(true);
                                                      setAcessFile(file); // file = your folder/file object
                                                    }}
                                                  >
                                                     <img src={require("../assets/revokeacc.svg")} alt="revoke" /> Revoke Access
                                                  </Button></li>
                                                </>
                                              )}

                                              {/* Menu for Recycle Bin */}
                                              {activeView === "Recycle bin" && (
                                                <>
                                                  <li><button type="button" onClick={() => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    setPreviewFile(file);
                                                    setShowPreviewModal(true);
                                                  }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
                                                  <li><button type="button" onClick={async () => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    await handleUndoDelete(file); setMenuOpenIdx(null);
                                                  }}
                                                  ><img src={require("../assets/undo.svg")} alt="rename" /> Undo (Restore)</button></li>
                                                </>
                                              )}
                                                
                                                {/* // srs 6/3/26 */}
                                                {activeView === "Archived Files" && (
                                                <>
                                                  <li><button type="button" onClick={async () => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    await handleUndoArchive(file); setMenuOpenIdx(null);
                                                  }}
                                                  ><img src={require("../assets/undo.svg")} alt="rename" /> Unarchive</button></li>
                                                </>
                                              )}
                                            </ul>
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  )
                                  //Rohit 22/04/2026 --------------start
                                  )}
                                    
                                  </>

                                )}
                                {/* Rohit 22/04/2026 end */}
                            </div>
                          )}

                          {activeLayout === "list" && (
                            <div>
                              <table className="mtablenew">
                              <thead>
  <tr>
    <th style={{ minWidth: '50px', maxWidth: '50px' }}>S.No</th>
    <th style={{ minWidth: '250px', maxWidth: '250px', cursor: 'pointer' }}
      onClick={() => handleSort("name")}>
      Name <FontAwesomeIcon icon={faSort} />
    </th>
    <th style={{ minWidth: '80px', maxWidth: '80px', cursor: 'pointer' }}
      onClick={() => handleSort("size")}>
      Size <FontAwesomeIcon icon={faSort} />
    </th>
    <th style={{ minWidth: '150px', maxWidth: '150px', cursor: 'pointer' }}
      onClick={() => handleSort(paginatedFiles.some((f) => f.DocumentLibraryName) ? "library" : "createdDate")}>
      {paginatedFiles.some((f) => f.DocumentLibraryName)
        ? <>Library <FontAwesomeIcon icon={faSort} /></>
        : <>Created Date <FontAwesomeIcon icon={faSort} /></>}
    </th>
    {paginatedFiles.some((f) => f.Status && f.Status.trim() !== "") && (
      <th style={{ minWidth: '118px', maxWidth: '118px', cursor: 'pointer' }}
        onClick={() => handleSort("status")}>
        Status <FontAwesomeIcon icon={faSort} />
      </th>
    )}
    <th style={{ textAlign: "center", minWidth: '80px', maxWidth: '80px' }}>Action</th>
  </tr>
  {/* ✅ Search Row */}
  <tr>
    <th style={{ minWidth: '50px', maxWidth: '50px' }}></th>
    <th style={{ minWidth: '250px', maxWidth: '250px' }}>
      <input
        type="text"
        placeholder="Search name..."
        value={columnSearch.name}
        onChange={(e) => setColumnSearch(prev => ({ ...prev, name: e.target.value }))}
        onClick={(e) => e.stopPropagation()}
        style={{ width: '100%', padding: '4px 6px', fontSize: 12, border: '1px solid #ddd', borderRadius: 4, boxSizing: 'border-box' }}
      />
    </th>
    <th style={{ minWidth: '80px', maxWidth: '80px' }}>
      <input
        type="text"
        placeholder="Search..."
        value={columnSearch.size}
        onChange={(e) => setColumnSearch(prev => ({ ...prev, size: e.target.value }))}
        onClick={(e) => e.stopPropagation()}
        style={{ width: '100%', padding: '4px 6px', fontSize: 12, border: '1px solid #ddd', borderRadius: 4, boxSizing: 'border-box' }}
      />
    </th>
    <th style={{ minWidth: '150px', maxWidth: '150px' }}>
      <input
        type="text"
        placeholder="Search..."
        value={columnSearch.library}
        onChange={(e) => setColumnSearch(prev => ({ ...prev, library: e.target.value }))}
        onClick={(e) => e.stopPropagation()}
        style={{ width: '100%', padding: '4px 6px', fontSize: 12, border: '1px solid #ddd', borderRadius: 4, boxSizing: 'border-box' }}
      />
    </th>
    {paginatedFiles.some((f) => f.Status && f.Status.trim() !== "") && (
      <th style={{ minWidth: '118px', maxWidth: '118px' }}>
        <input
          type="text"
          placeholder="Search status..."
          value={columnSearch.status}
          onChange={(e) => setColumnSearch(prev => ({ ...prev, status: e.target.value }))}
          onClick={(e) => e.stopPropagation()}
          style={{ width: '100%', padding: '4px 6px', fontSize: 12, border: '1px solid #ddd', borderRadius: 4, boxSizing: 'border-box' }}
        />
      </th>
    )}
    <th style={{ minWidth: '80px', maxWidth: '80px' }}></th>
  </tr>
</thead>

<tbody>
{/* Rohit 22/04/2026 ----start */}
{columnFilteredFiles.map((file, idx) => (
                                    <tr key={getQuickViewRenderKey(file, idx)}>
{/* //Rohit 22/04/2026  ----end */}
                                      <td style={{ minWidth: '50px', maxWidth: '50px', }}> <span className="indexdesign">{(currentPage - 1) * pageSize + idx + 1}</span>  </td>
                                      {/* //Rohit 22/04/2026  end ------ */}
                                      {/* <td style={{minWidth:'250px',maxWidth:'250px',}}>{file.FileName}</td> */}
                                      {/* <td style={{ minWidth: '250px', maxWidth: '250px', }}>{file.FileName || file.Name || file.FolderName}</td> */}
                                      {/* <td >{file.FileSize}</td> */}
                                     
<td style={{ minWidth: '250px', maxWidth: '250px' }}>
  <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
    <div
      dangerouslySetInnerHTML={{
        __html: createFileExtensionHtml(
          file.FileName || file.Name || file.FolderName || "",
          undefined,  // bgColor parameter (optional)
          true        // ✅ isListView = true
        ),
      }}
    />
    {/* <span>{file.FileName || file.Name || file.FolderName}</span> */}
    {/* // addhyan 17/3/26*/}
    <span>{file.DisplayName ||file.FileName || file.Name || file.FolderName}</span>
  </div>
</td>
                                      <td style={{minWidth: '80px', maxWidth: '80px'}}>
                                        {file.FileSize
    ? (() => {
        const sizeInBytes = parseInt(file.FileSize);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
                                      </td>
                                     
                                      {/* <td  style={{minWidth: '150px', maxWidth: '150px'}}>
                                        {file.Length
    ? (() => {
        const sizeInBytes = parseInt(file.Length);
        const sizeInKB = sizeInBytes / 1024;
        const sizeInMB = sizeInBytes / (1024 * 1024);

        return sizeInMB < 1
          ? `${Math.floor(sizeInKB)} KB`
          : `${sizeInMB.toFixed(2)} MB`;
      })()
    : ""}
                                      </td> */}
                                       <td style={{minWidth: '150px', maxWidth: '150px'}}>{file.DocumentLibraryName}</td>
                                      {/* <td style={{minWidth:'80px',maxWidth:'80px',}}>{file.Status}</td> */}
                                     {paginatedFiles.some((f) => f.Status && f.Status.trim() !== "") && (
                                        <td style={{ minWidth: '118px', maxWidth: '118px', }}>
                                          {file.Status === "Pending" && (
                                            <div className="font-12 text-center" style={{ backgroundColor: '#fffbeb', color:'#b45309', border:'1px solid #fde68a', padding: '4px 8px', borderRadius: '30px', whiteSpace: 'nowrap' }}><FontAwesomeIcon icon={faClock} /> {file.Status}</div>
                                          )}
                                          {file.Status === "Auto Approved" && (
                                            <div className="text-success text-center font-12" style={{backgroundColor: '#f0fdf4', color:'#15803d', border:'1px solid #bbf7d0', padding: '4px 8px', borderRadius: '30px', whiteSpace: 'nowrap' }}><FontAwesomeIcon icon={faCheckCircle} /> {file.Status}</div>
                                          )}
                                          {file.Status === "Rejected" && (
                                            <div className="text-danger text-center font-12" style={{ backgroundColor: 'rgb(253, 236, 234)', border:'1px solid rgb(253, 230, 138)', padding: '4px 8px', borderRadius: '30px', whiteSpace: 'nowrap' }}><FontAwesomeIcon icon={faRedo} /> {file.Status}</div>
                                          )}
                                          {file.Status === "Rework" && (
                                            <div className="text-danger text-center font-12" style={{ backgroundColor: '#fef2f2',color:'#b91c1c', border:'1px solid #fecaca', padding: '4px 8px', borderRadius: '30px', whiteSpace: 'nowrap' }}><FontAwesomeIcon icon={faRedo} /> {file.Status}</div>
                                          )}
                                          {file.Status === "Approved" && (
                                            <div className="text-success text-center font-12" style={{ backgroundColor: '#f0fdf4', color:'#15803d', border:'1px solid #bbf7d0', padding: '4px 8px', borderRadius: '30px', whiteSpace: 'nowrap' }}><FontAwesomeIcon icon={faCheckCircle} /> {file.Status}</div>
                                          )}
                                          {!file.Status || (file.Status !== "Pending" && file.Status !== "Auto Approved" && file.Status !== "Rejected" && file.Status !== "Rework" && file.Status !== "Approved") && (
                                            <div className="text-muted text-center font-12">{file.Status || ""}</div>
                                          )}
                                        </td>
                                      )}
                                      <td style={{ textAlign: "center",minWidth: '80px', maxWidth: '80px' }}>
                                        <button type="button" className="dotbutton2"
                                        ref={(el) => menuButtonRefs.current[idx] = el} // ritik 18/03/26
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setMenuOpenIdx(menuOpenIdx === idx ? null : idx);
                                          }}

                                        >
                                          ⋮
                                        </button>

                                        {/* {menuOpenIdx === idx && (
                                          <div
                                            style={{
                                              position: "absolute",
                                              right: 0,
                                              top: "20px",
                                              background: "#fff",
                                              border: "1px solid #ddd",
                                              borderRadius: "4px",
                                              boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
                                              zIndex: 10,
                                              minWidth: "140px",
                                            }}
                                          > */}
                                          {/* ritik 18/03/26 for position of menu button */}
                                          {menuOpenIdx === idx && (() => {
  const btn = menuButtonRefs.current[idx];
  const rect = btn?.getBoundingClientRect();
  return (
  <div
    style={{
      position: "fixed",
      top: rect ? rect.bottom - 50 : "auto",
      right: rect ? window.innerWidth - rect.right : -0,
      background: "#fff",
      border: "1px solid #ddd",
      borderRadius: "4px",
      boxShadow: "0 2px 8px rgba(0,0,0,0.15)",
      zIndex: 99999,
      minWidth: "140px",
    }}
  >
                                            <ul className="internalbutton">
  {/* --- SYNCED DYNAMIC LOGIC FOR ALL MODULES --- */}
  {filesLoadedfromnode && 
   activeView !== "My request" && 
   activeView !== "My favourite" && 
   activeView !== "My Folders" && 
   activeView !== "Share with me" && 
   activeView !== "Share with other" && 
   activeView !== "Recycle bin" ? (
    <>
      {/* Browsing Menu Logic (Map 1 Sync) */}
      <li><button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
      <li><button type="button" className="newbuttontext" onClick={async () => {await ArchivedFile(file,currentSiteUrl,context);setMenuOpenIdx(null);}}><span><FontAwesomeIcon icon={faTrashAlt} /></span> Archive</button></li>
           {/* // Ritik 31/3/26 */}
      <li><button type="button" onClick={() => { setShareFile(file); setShowShareModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faShareAlt} /> Share</button></li>
      <li><button type="button" onClick={() => { handleAuditHistory(file); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faFileAlt} /> Audit History</button></li>
       {/* srs 10/4/26 */}
      <li><button type="button" onClick={() => { 
        // 1. Get the domain (https://officeindia.sharepoint.com)
      const origin = window.location.origin;

      // 2. Construct the URL by finding the path up to the SiteName
      // This ensures even if you are deep in a folder, the URL stops at the subsite level
      // const pathUntilSite = file.CurrentFolderPath.split(file.SiteName)[0];
      // aman - 29/04/26 start
      const pathRef = file.CurrentFolderPath || file.ServerRelativeUrl || "";
      const pathUntilSite = pathRef.split(file.SiteName)[0];
      const finalSiteUrl = `${origin}${pathUntilSite}${file.SiteName}`;

      console.log("Target Subsite URL:", finalSiteUrl);
       window.ManageFilePermission(
          // file.FileUID,            // fileId
           file.FileUID || file.UniqueId,
        finalSiteUrl,             // siteId (Your data shows this is the URL)
        // file.DocumentLibraryName, // documentLibraryName
        file.DocumentLibraryName || file.__libraryTitle || file.LibraryName,
        // aman - 29/04/26 end 
 
        file.SiteName            // siteTitle (Used for the Admin Group naming)
      ); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faFileAlt} /> Manage File Permission</button></li>
      <li><button type="button" onClick={() => {
                                                    setMenuOpenIdx(null);
                                                    setShareUrlFile(file);
                                                    setShowShareUrlModal(true);
                                                  }}><FontAwesomeIcon icon={faShareAlt} /> Share file url</button></li>
      <li><button type="button" onClick={() => { setModalFile(file); setShowVersionModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faHistory} /> Version History</button></li>
      <li><button type="button" onClick={async () => { await deleteFileFolder(file, currentSiteUrl, context); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faTrashAlt} /> Delete File</button></li>
      <li>
        <button type="button" onClick={async () => { 
          const newFav = await toggleFavouriteDoc(file, currentSiteUrl, context); 
          setSelectedFiles((prev) => prev.map((f) => f.UniqueId === file.UniqueId ? { ...f, IsFavourite: newFav } : f));
          setMenuOpenIdx(null);
        }}><FontAwesomeIcon icon={faHearts} /> {file.IsFavourite ? "Unmark Favourite" : "Mark as Favourite"}</button>
      </li>
    </>
  ) : (
    <>
      {/* Tab-Specific Menu Logic (Map 2 Sync) */}
      {activeView === "My request" && (
        <>
          <li><button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
          <li><button type="button" onClick={() => { handleAuditHistory(file); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faFileAlt} />  Audit History</button></li>
          {/* srs 10/4/26 */}
          <li><button type="button" onClick={() => { 
            // 1. Get the domain (https://officeindia.sharepoint.com)
      const origin = window.location.origin;

      // 2. Construct the URL by finding the path up to the SiteName
      // This ensures even if you are deep in a folder, the URL stops at the subsite level
      const pathUntilSite = file.CurrentFolderPath.split(file.SiteName)[0];
      const finalSiteUrl = `${origin}${pathUntilSite}${file.SiteName}`;

      console.log("Target Subsite URL:", finalSiteUrl);
       window.ManageFilePermission(
          file.FileUID,            // fileId
        finalSiteUrl,             // siteId (Your data shows this is the URL)
        file.DocumentLibraryName, // documentLibraryName
        file.SiteName            // siteTitle (Used for the Admin Group naming)
      ); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faFileAlt} />  Manage File Permission</button></li>
      {/* srs 15/4/26 */}
      {file.Status === "Rework" && (
  <li>
  <button 
    type="button"
    className="newbuttontext"
    onClick={() => {
      // Extract the required fields from your file object
      // Note: Ensure your 'file' object has these properties (adjust names if necessary)
       const constructedPath = `${file.CurrentFolderPath || file.ServerRelativeUrl}/${file.FileName || file.Name}`;
      window.rework(
      file.FileUID,          // 1. GUID (e.g. "afaec963...")
  file.SiteID,           // 2. URL (e.g. "https://officeindia...")
  file.DocumentLibraryName, 
  file.SiteName, 
  constructedPath
      );
      setMenuOpenIdx(null);
    }}
  >
    <span> <FontAwesomeIcon icon={faFileAlt} /> </span>Edit File
  </button>
</li>
)}
          <li><button type="button" onClick={() => { setShareFile(file); setShowShareModal(true); setMenuOpenIdx(null); }}> <FontAwesomeIcon icon={faShareAlt} />  Share</button></li>
          {/* <li><button type="button" onClick={() => { setDirectDownloadFile(file); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faDownload} />  Download</button></li> */}
          {/* ririk  */}
          <li><button type="button" onClick={() => { setDirectDownloadFile(file); setMenuOpenIdx(null); setTimeout(() => setDirectDownloadFile(null), 3000);}}><FontAwesomeIcon icon={faDownload} />  Download</button></li>
          <li><button type="button" onClick={() => { setModalFile(file); setShowVersionModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faHistory} /> Version History</button></li>
          <li><button type="button" onClick={() => {
                                                    setMenuOpenIdx(null);
                                                    setShareUrlFile(file);
                                                    setShowShareUrlModal(true);
                                                  }}><FontAwesomeIcon icon={faShareAlt} /> Share file url</button></li>
          {/* Addhyan 19/2/26 */}
          {/* <li><button type="button" onClick={() => { handleRenameFile(file); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEdit} />  Rename File</button></li> */}
        </>
      )}

      {activeView === "My favourite" && (
        <>
          <li><button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEye} />  Preview File</button></li>
          <li><button type="button" onClick={() => { setShareFile(file); setShowShareModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faShareAlt} /> Share</button></li>
          {/* <li><button type="button" onClick={() => { toggleFavourite(file); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faHearts} /> Unmark as Favourite</button></li> */}
          {/* Addhyan 19/2/26 */}
          <li><button type="button" onClick={() => { toggleFavourite(file, currentSiteUrl, context); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faHearts} /> Unmark as Favourite</button></li>
        </>
      )}

      {activeView === "My Folders" && (
        <>
          {!file.IsFolder && (
            <>
              <li><button type="button" onClick={() => { setSelectedWorkflowFolder({ SiteID: file.SiteID, SiteTitle: file.SiteTitle, SiteUrl: file.SiteUrl, DocumentLibraryName: file.DocumentLibraryName, file: file }); setShowManageWorkflow(true); setMenuOpenIdx(null); }}><img src={require("../assets/managework.svg")} alt="userlock" /> Manage WorkFlow</button></li>
              <li><button type="button" onClick={() => { setSelectedFileForMetadata(file); setShowMetadataModal(true); setMenuOpenIdx(null); }}> <img src={require("../assets/metadata.svg")} alt="userlock" /> Add MetaData</button></li>
              {/* srs 20/2/26 */}
              <li><button type="button" onClick={() => {
                                                          setMenuOpenIdx(null);
                                                          setSelectedFileForRenameMetadata(file);
                                                          setShowRenameMetadataModal(true);
                                                        }}>  <img src={require("../assets/Rename-Meta-Data.svg")} alt="userlock" /> Rename Meta Data</button></li>
            </>
          )}
          <li><button type="button" onClick={() => { setModalFile(file); deleteFolder(file); setMenuOpenIdx(null); }}><img src={require("../assets/deletefile.svg")} alt="deletefile" /> Delete Folder</button></li>
          <li><button type="button" onClick={() => { setModalFile(file); setRenameValue(file?.FolderName || ""); setRenameModalOpen(true); setMenuOpenIdx(null); }}><img src={require("../assets/rename.svg")} alt="rename" /> Rename Folder</button></li>
          <li><button type="button" onClick={() => { window.managePermission?.(file); setMenuOpenIdx(null); }}><img src={require("../assets/userlock.svg")} alt="userlock" /> Manage Permission</button></li>
          {/* Aman 13/4/26 */}
           <li>
  <button
    type="button"
    onClick={() => {
      setMenuOpenIdx(null);
      handleFolderShareClick(file);
    }}
  >
    <FontAwesomeIcon icon={faShareAlt} /> Share Folder Path
  </button>
</li>
        </>
      )}

      {activeView === "Share with me" && (
        <>
          <li><button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
          <li><button type="button" onClick={() => { setDirectDownloadFile(file); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faDownload} /> Download File</button></li>
        </>
      )}

      {activeView === "Share with other" && (
        <>
          <li><button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
          <li><Button variant="warning" size="sm" onClick={() => { setAcessFile(file); setShowRevokeModal(true); setMenuOpenIdx(null); }}><img src={require("../assets/undo.svg")} alt="rename" /> Revoke Access</Button></li>
        </>
      )}

      {activeView === "Recycle bin" && (
        <>
          <li><button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}><FontAwesomeIcon icon={faEye} /> Preview File</button></li>
          <li><button type="button" onClick={async () => { await handleUndoDelete(file); setMenuOpenIdx(null); }}><img src={require("../assets/undo.svg")} alt="rename" />  Undo (Restore)</button></li>
        </>
      )}
    </>
  )}
  {/* addhyan - 17/03/26 */}
  {activeView === "Archived Files" && (
                                                <>
                                                  <li><button type="button" onClick={async () => {
                                                    // this is to hide li options in every tab (addhyan)
                                                    setMenuOpenIdx(null);
                                                    // this is to hide li options in every tab (addhyan)
                                                    await handleUndoArchive(file); setMenuOpenIdx(null);
                                                  }}
                                                  ><img src={require("../assets/undo.svg")} alt="rename" /> Unarchive</button></li>
                                                </>
                                              )}
 
                                              {/* addhyan - 17/03/26 */}
</ul>
                                          </div>
                                          );
                                })()}
                                      </td>
                                    </tr>
                                  ))}
{/* Rohit 22/04/2026 end */}
                                </tbody>

                              </table>
                            </div>
                          )}
                        </>
                      ) : (
                        <div
                          style={{
                            color: "#666",
                            fontSize: "14px",
                            padding: "20px",
                            textAlign: "center",
                            marginTop: "20px",
                          }}
                        >
                          {searchTerm.trim() !== ""
                            ? `No records found matching "${searchTerm}"`
                             : <div style={{display:'flex',  marginTop:'40px', alignItems:'center', justifyContent:'center', flexDirection:'column'}}> <img src={require('../assets/no-filef.png')}></img>  <p className="mt-2 font-14 fw-bold text-dark">No files found in this view</p> </div>}
                        </div>
                      )}
                      {/* create Folder modal */}
                      <Modal show={activeComponent} onHide={() => setActiveComponent(false)} className='filemodal' size="lg" centered container={() => document.getElementById('filelistcontainer')} backdrop={true}>
                        <Modal.Header closeButton>
                          <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-1'>This Folder will create under: <h2
                            style={{
                              fontSize: "16px",
                              fontWeight: "600",
                              marginBottom: "15px",
                              color: "#333",
                              paddingBottom: "5px",
                              borderBottom: "1px solid #eee",
                              paddingRight: "100px", // Make space for upload button
                            }}
                          >
                            {breadcrumbs.length > 0 ? (
                              <div
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  flexWrap: "wrap",
                                }}
                              >
                                {breadcrumbs.map((item, index) => (
                                  <React.Fragment key={item.key}>
                                    {index > 0 && (
                                      <span
                                        style={{
                                          margin: "0 8px",
                                          color: "#999",
                                          fontSize: "14px",
                                        }}
                                      >
                                        ›
                                      </span>
                                    )}
                                    <span
                                      className="breadcrumb-item"
                                      // onClick={() => handleBreadcrumbClick(item)}
                                      style={{
                                        cursor: "pointer",
                                        color:
                                          "#333",
                                        // color:
                                        // index === breadcrumbs.length - 1 ? "#333" : "#0066cc",
                                        fontWeight:
                                          index === breadcrumbs.length - 1 ? "600" : "normal",
                                        fontSize: "14px",
                                        padding: "2px 5px",
                                        borderRadius: "3px",
                                      }}
                                    >
                                      {item.title}
                                    </span>
                                  </React.Fragment>
                                ))}
                              </div>
                            ) : (
                              <div style={{ color: "#666", fontSize: "14px" }}>
                                Select a folder to view files
                              </div>
                            )}
                          </h2></h4>
                            {/* <p className='text-muted font-14 mb-0 fw-400'>Below are the attachment details for Memorandum
              </p> */}

                          </Modal.Title>


                        </Modal.Header>
                        <Modal.Body className="" id="style-5">

                          <>

                            {activeComponent && (
                              <CreateFolder
                                OthProps={{
                                  "Entity": `${currentSiteUrl.split("/sites/")[1].split("/")[1]}`,
                                  "Entityurl": currentSiteUrl,
                                  "siteID": currentSiteUrl,
                                  "Devision": "",
                                  "Department": "",
                                  // "DocumentLibrary": currentFolderPath,
                                  "DocumentLibrary": selectedCurrentNode.type === "site" || selectedCurrentNode.type === "subsite" ? "" : selectedCurrentNode.title,
                                  "Folder": currentFolderPath,
                                  "folderpath": `/sites/${currentSiteUrl.split("/sites/")[1]}/${currentFolderPath}`,
                                  "IsFolderDeligationUser": "false",
                                }}
                                context={context}
                                onCloseForm={() => setActiveComponent(false)} // aman 20/04/26 to close modal after folder creation
                              // onReturnToMain={handleReturnToMain}
                              />
                            )}

                          </>
                          {/* </>
              )
              } */}

                        </Modal.Body>

                      </Modal>
                      {/* // sourish 21/8/25 */}
                      {renameModalOpen && (
  <div
    style={{
      position: "absolute",  // ✅ fixed se absolute
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      background: "rgba(0,0,0,0.5)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 9999,
    }}
  >
                          <div
                            style={{
                              background: "#fff",
                              padding: "20px",
                              borderRadius: "10px",
                              width: "320px",
                              boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
                              display: "flex",
                              flexDirection: "column",
                              gap: "12px",
                            }}
                          >
                            <h3 style={{ margin: 0 }}>Rename Folder</h3>
                            <input
                              type="text"
                              value={renameValue}
                              onChange={(e) => setRenameValue(e.target.value)}
                              style={{
                                width: "100%",
                                padding: "8px",
                                borderRadius: "6px",
                                border: "1px solid #ccc",
                              }}
                            />
                            <div
                              style={{
                                display: "flex",
                                justifyContent: "flex-end",
                                gap: "10px",
                                marginTop: "10px",
                              }}
                            >
                              <button
                                style={{
                                  
                                  padding: "6px 12px",
                                  borderRadius: "6px",
                                  border: "1px solid #ccc",
                                  background: "red",
                                  cursor: "pointer",
                                }}
                                onClick={() => setRenameModalOpen(false)}
                              >
                                Cancel
                              </button>
                              {/* <button
                                style={{
                                  padding: "6px 12px",
                                  borderRadius: "6px",
                                  border: "none",
                                  background: "#0078d4",
                                  color: "#fff",
                                  cursor: "pointer",
                                }}
                                onClick={renameFolder}
                              >
                                Rename
                              </button> */}

                              <button
  style={{
    padding: "6px 12px",
    borderRadius: "6px",
    border: "none",
    background: "transparent",
    cursor: "pointer",
  }}
  onClick={renameFolder}
>
<img src={require("../assets/submit-new1.png")} alt="Rename" style={{ height: "24px" }} />
</button>
                            </div>
                          </div>
                        </div>
                      )}


                      {/* srs 20/2/26 */}
{showRenameMetadataModal && (
  <div
  onClick={(e) => { if (e.target === e.currentTarget) setShowRenameMetadataModal(false); }} // Ritik 17/03/2026 close modal
    style={{
      position: "absolute", // Changed from fixed to absolute to match Rename Folder
      top: 0, 
      left: 0, 
      right: 0, 
      bottom: 0,
      background: "rgba(0,0,0,0.5)",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 9999,
    }}
  >
    <div
      className="blur-background"
      style={{
        background: "#fff",
        padding: "30px",
        borderRadius: "20px",
        width: "450px", // You can reduce this to "320px" if you want the exact same width as the folder modal
        boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
        position: "relative",
      }}
    >
      {/* Red Close Button */}
      <button
        onClick={() => setShowRenameMetadataModal(false)}
        style={{
          position: "absolute",
          top: "15px",
          right: "15px",
          width: "30px",
          height: "30px",
          borderRadius: "50%",
          border: "2px solid #ec1c24",
          background: "#fff",
          color: "#ec1c24",
          fontSize: "20px",
          fontWeight: "bold",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          cursor: "pointer",
          padding: "0",
          lineHeight: "1"
        }}
      >
        &times;
      </button>

      <h3 style={{ 
        marginTop: 0, 
        textAlign: "center", 
        fontSize: "22px", 
        fontWeight: "500",
        color: "#333",
        borderBottom: "1px solid #eee",
        paddingBottom: "15px"
      }}>
        Rename Meta Columns
      </h3>

      <div style={{ maxHeight: "300px", overflowY: "auto", marginTop: "15px", paddingRight: "5px" }}>
       {loadingColumns ? (
            <div style={{ textAlign: "center", padding: "40px", color: "#666" }}>
                Processing metadata...
            </div>
        ) : (
          existingColumns.map((column, index) => (
            <div key={column.ID} style={{ marginBottom: "20px" }}>
              <label style={{ 
                display: "block", 
                fontWeight: "500", 
                marginBottom: "8px", 
                fontSize: "14px",
                color: "#555" 
              }}>
                Column {index + 1} ({column.ColumnName})
              </label>
              <input
                type="text"
                value={column.IsRename || column.ColumnName || ""}
                onChange={(e) => {
                  const updated = [...existingColumns];
                  updated[index].IsRename = e.target.value;
                  setExistingColumns(updated);
                }}
                style={{
                  width: "100%",
                  padding: "12px",
                  fontSize: "14px",
                  border: "1px solid #ddd",
                  borderRadius: "8px",
                  boxSizing: "border-box"
                }}
              />
            </div>
          ))
        )}
      </div>

      <div style={{ 
        display: "flex", 
        justifyContent: "center", 
        gap: "15px", 
        marginTop: "30px" 
      }}>
        <button type="button"
          onClick={() => setShowRenameMetadataModal(false)}
          style={{
            background: "#cccccc",
            color: "#fff",
            border: "none",
            padding: "10px 30px",
            borderRadius: "6px",
            fontSize: "16px",
            fontWeight: "500",
            cursor: "pointer",
            minWidth: "120px"
          }}
        >
          Cancel
        </button>
        <button type="button"
          onClick={handleSaveRename}
          style={{
            background: "#28a745",
            color: "#fff",
            border: "none",
            padding: "10px 30px",
            borderRadius: "6px",
            fontSize: "16px",
            fontWeight: "500",
            cursor: "pointer",
            minWidth: "120px"
          }}
        >
          Save
        </button>
      </div>
    </div>
  </div>
)}
                      {/* download component */}
                      <DirectDownloader
                        file={directDownloadFile}
                        context={context}
                        // trigger={true}
                        trigger={!!directDownloadFile}  // Ritik added 23/2/26
                      />




                      {/*revoke modal*/}
                      <Revoke
                                  show={showRevokeModal}
                        selectedFolder={revoke}
                        context={context}
                        onClose={() => setShowRevokeModal(false)}
                        onRevoke={(userId: string) => {
                          console.log("Revoked user ID:", userId);
                          // alert(`Access revoked for user ID: ${userId}`);
                          
                        }}
                         onSuccess={() => handleViewButtonClick("Share with other")} // ritik add 29/04/26
                      />

                      {/* version history */}
                      <VersionHistoryModal
                        show={showVersionModal}
                        file={modalFile}
                        context={context}
                        onClose={() => setShowVersionModal(false)}
                      />

                      {/* share modal */}
                      <ShareModal
                                  show={showShareModal}
                        
                        file={shareFile}
                        context={context}
                        onClose={() => {
                          setShowShareModal(false);
                          setShareFile(null);
                        }}
                        currentUserEmail={(context.pageContext as any)?.user?.email || ""}
                        currentUserTitle={context.pageContext.user.displayName || ""}
                        currentSiteUrl={currentSiteUrl}
                      />





{/* // add call for share file url modal - Addhyan 13/4/26 start */}
                      <ShareFileUrlModal
                        show={showShareUrlModal}
                        file={shareUrlFile}
                        context={context}
                        currentSiteUrl={currentSiteUrl}
                        onClose={() => {
                          setShowShareUrlModal(false);
                          setShareUrlFile(null);
                        }}
                      />




                      {/* // ritik 20/04/26 - start  */}


                      {showFilePermModal && filePermState && (
  <Modal
    show={showFilePermModal}
    onHide={() => setShowFilePermModal(false)}
    size="lg"
    container={() => document.getElementById('filelistcontainer')}
    backdrop={true}
    centered
  >
    <Modal.Header closeButton>
      <Modal.Title>Manage File Permission</Modal.Title>
    </Modal.Header>
 
    <Modal.Body style={{ maxHeight: '60vh', overflowY: 'auto' }}>
 
      {/* Status */}
      <div style={{
        marginBottom: '15px',
        marginTop: '8px',
        padding: '10px',
        fontSize: '12px',
        borderRadius: '4px',
        border: `1px solid ${filePermState.isUnique ? '#fbc7c7' : '#c7ebc7'}`,
        background: filePermState.isUnique ? '#fff4f4' : '#f3fbf3',
        color: filePermState.isUnique ? '#d13438' : '#107c10',
      }}>
        <strong>Current Status:</strong>{' '}
        {filePermState.isUnique ? '⚠️ Unique Permissions' : '✅ Inheriting Permissions'}
      </div>
 
      {/* Add New Users label */}
      <label style={{ fontWeight: 600, display: 'block', marginBottom: '5px' }}>
        Add New Users:
      </label>
 
      {/* Selected user tags */}
      <div style={{
        display: 'flex', flexWrap: 'wrap', gap: '5px',
        marginBottom: '10px', minHeight: '35px',
        border: '1px solid #ddd', padding: '5px',
        borderRadius: '4px', background: '#faf9f8'
      }}>
        {filePermState.selectedUsers.length === 0
          ? <span style={{ color: '#999', fontSize: '12px' }}>Search and click users...</span>
          : filePermState.selectedUsers.map((u: any) => (
              <span key={u.id} style={{
                background: '#0078d4', color: 'white',
                padding: '2px 8px', borderRadius: '12px',
                fontSize: '12px', display: 'flex',
                alignItems: 'center', gap: '5px'
              }}>
                {u.title}
                <span
                  style={{ cursor: 'pointer', fontWeight: 'bold' }}
                  onClick={() => setFilePermState((prev: any) => ({
                    ...prev,
                    selectedUsers: prev.selectedUsers.filter((su: any) => su.id !== u.id)
                  }))}
                >
                  &times;
                </span>
              </span>
            ))
        }
      </div>
 
      {/* User search */}
      <FilePermUserSearch
        users={filePermState.filteredUsers}
        onSelect={(user: any) => {
          setFilePermState((prev: any) => {
            if (prev.selectedUsers.find((u: any) => u.id === user.id)) return prev;
            return { ...prev, selectedUsers: [...prev.selectedUsers, user] };
          });
        }}
      />
 
      {/* Permission dropdown */}
      <label style={{ fontWeight: 600, display: 'block', margin: '20px 0 5px 0' }}>
        Permission Level:
      </label>
      <select
        className="form-select"
        value={filePermState.permission}
        onChange={(e) => setFilePermState((prev: any) => ({
          ...prev, permission: e.target.value
        }))}
        style={{ width: '100%', marginBottom: '20px', fontSize: '14px', padding: '6px' }}
      >
        <option value="Full Control">Full Control</option>
        <option value="Edit">Edit</option>
        <option value="Contribute">Contribute</option>
        <option value="Read">Read</option>
      </select>
 
      {/* Existing access table — only if unique */}
      {filePermState.isUnique && (
        <>
          <hr style={{ border: 0, borderTop: '1px solid #eee', margin: '15px 0' }} />
          <label style={{ fontWeight: 600, display: 'block', marginBottom: '5px' }}>
            Existing Access:
          </label>
          <div style={{ maxHeight: '150px', overflowY: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', border: '1px solid #eee' }}>
              <thead>
                <tr>
                  <th style={{ background: '#f3f2f1', padding: '8px', textAlign: 'left', borderBottom: '1px solid #ddd' }}>User</th>
                  <th style={{ background: '#f3f2f1', padding: '8px', textAlign: 'left', borderBottom: '1px solid #ddd' }}>Permission</th>
                  <th style={{ background: '#f3f2f1', padding: '8px', textAlign: 'center', borderBottom: '1px solid #ddd' }}>Action</th>
                </tr>
              </thead>
              <tbody>
                {filePermState.currentAssignments
                  .filter((a: any) => (a.Member?.Title || '').trim() !== 'DMSSuper_Admin')
                  .map((a: any) => (
                    <tr key={a.PrincipalId}>
                      <td style={{ padding: '6px 8px', borderBottom: '1px solid #eee' }}>
                        <div style={{ fontWeight: 600 }}>{a.Member?.Title || 'Unknown'}</div>
                        <div style={{ fontSize: '10px', color: '#666' }}>{a.Member?.Email || ''}</div>
                      </td>
                      <td style={{ padding: '6px 8px', borderBottom: '1px solid #eee' }}>
                        {a.RoleDefinitionBindings.map((r: any) => r.Name).join(', ')}
                      </td>
                      <td style={{ padding: '6px 8px', textAlign: 'center', borderBottom: '1px solid #eee' }}>
                        <button
                          type="button"
                          style={{
                            color: '#d13438', cursor: 'pointer',
                            fontSize: '18px', background: 'none',
                            border: 'none', fontWeight: 'bold'
                          }}
                          onClick={async () => {
                            try {
                              const web2 = Web(filePermState.siteUrl).using(AssignFrom(sp.web));
                              const item2 = await web2.getFileById(filePermState.fileId).getItem();
                              await item2.roleAssignments.getById(a.PrincipalId).delete();
                              // Refresh
                              const newAssignments: any[] = await item2.roleAssignments
                                .expand("Member", "RoleDefinitionBindings")();
                              setFilePermState((prev: any) => ({
                                ...prev,
                                currentAssignments: newAssignments
                              }));
                            } catch (err: any) {
                              Swal.fire({ icon: 'error', title: 'Error', text: err.message });
                            }
                          }}
                        >
                          &times;
                        </button>
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </>
      )}
 
    </Modal.Body>
 
    <Modal.Footer>
      {/* Restore Inheritance button */}
      {filePermState.isUnique && (
        <Button
          variant="secondary"
          className="mt-0"
          onClick={async () => {
            try {
              const web3 = Web(filePermState.siteUrl).using(AssignFrom(sp.web));
              const item3 = await web3.getFileById(filePermState.fileId).getItem();
              await item3.resetRoleInheritance();
              setShowFilePermModal(false);
              // Reopen to refresh
              window.ManageFilePermission(
                filePermState.fileId,
                filePermState.siteUrl,
                filePermState.documentLibraryName,
                filePermState.siteTitle
              );
            } catch (err: any) {
              Swal.fire({ icon: 'error', title: 'Error', text: err.message });
            }
          }}
        >
          Restore Inheritance
        </Button>
      )}
 
      <Button
        variant="secondary"
        className="mt-0"
        onClick={() => setShowFilePermModal(false)}
      >
        Cancel
      </Button>
 
      {/* Grant Access button */}
      <Button
        variant="primary"
        className="mt-0"
        onClick={async () => {
          if (filePermState.selectedUsers.length === 0) {
            Swal.fire({ icon: 'warning', title: 'Please select at least one user' });
            return;
          }
          try {
            const web4 = Web(filePermState.siteUrl).using(AssignFrom(sp.web));
            const item4 = await web4.getFileById(filePermState.fileId).getItem();
 
            const adminGroupName = `${filePermState.siteTitle}_Admin`.trim();
            const superAdminName = "DMSSuper_Admin";
 
            if (!filePermState.isUnique) {
              await item4.breakRoleInheritance(true);
              const assignments: any[] = await item4.roleAssignments
                .expand("Member")
                .select("PrincipalId", "Member/Title")();
              for (const assignment of assignments) {
                const title = (assignment.Member?.Title || '').trim();
                if (title !== superAdminName && title !== adminGroupName) {
                  try {
                    await item4.roleAssignments.getById(assignment.PrincipalId).delete();
                  } catch (e) {
                    console.warn(`Cleanup: Could not remove ${title}`);
                  }
                }
              }
            }
 
            const roleDef = await web4.roleDefinitions
              .getByName(filePermState.permission)();
            for (const user of filePermState.selectedUsers) {
              await item4.roleAssignments.add(parseInt(user.id), roleDef.Id);
            }
 
            setShowFilePermModal(false);
            Swal.fire({
              icon: 'success',
              title: 'Permissions Set',
              timer: 1500,
              showConfirmButton: false,
            }).then(() => {
              window.ManageFilePermission(
                filePermState.fileId,
                filePermState.siteUrl,
                filePermState.documentLibraryName,
                filePermState.siteTitle
              );
            });
 
          } catch (err: any) {
            console.error("Grant Access Error:", err);
            Swal.fire({ icon: 'error', title: 'Operation Failed', text: err.message });
          }
        }}
      >
        Grant Access
      </Button>
    </Modal.Footer>
  </Modal>
)}


{/* ritik - 20/04/26 - end  */}
{/* // add call for share file url modal - Addhyan 13/4/26 start */}

{/* Aman 13/4/26 */}
 <BreadcrumbSharePopup
  show={breadcrumbShare.show}
  url={breadcrumbShare.url}
  onClose={closeBreadcrumbShare}
  context={context}
/>
 
<FolderSharePopup
  show={folderShare.show}
  url={folderShare.url}
  onClose={closeFolderShare}
  context={context}
/>


                      {/* aman code manage folder permission */}
                      {/* === Step 1 Popup === */}
                      {/* {showPermissionModal && (
                        <Modal show={showPermissionModal} onHide={() => setShowPermissionModal(false)} container={() => document.getElementById('filelistcontainer')}>
                    
                          <Modal.Header closeButton>
                            <Modal.Title>Set Permission</Modal.Title>
                          </Modal.Header>
                          <Modal.Body>
                            <div style={{ marginBottom: 10 }}>
                              Set permissions for:{" "}
                              <strong>{selectedFolder?.FolderName || selectedFolder?.Title || selectedFolder?.Name || "Selected Folder"}</strong>
                            </div>
                            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                              <button
                                onClick={() => setShowPermissionModal(false)}
                                style={{ padding: "6px 10px", border: "1px solid #ccc", borderRadius: 4, background: "#fff" }}
                              >
                                Cancel
                              </button>
                              <button
                                onClick={async () => {
                                  const pageBefore = currentPage;
                                  setShowPermissionModal(false);
                                  await loadAndOpenManagePermission();
                                  setCurrentPage(pageBefore);
                                }}
                                style={{ padding: "6px 10px", border: "1px solid #0b66c3", borderRadius: 4, background: "#0b66c3", color: "#fff" }}
                              >
                                Set Permission
                              </button>
                            </div>
                          </Modal.Body>
                        </Modal>
                      )} */}

                      {/* // Ritik 20/2/26 for manage permission folder */}
                      {/* model code for folder permission manage added by Ritik 20/02/2026 */}
                     {/* ===== MANAGE FOLDER PERMISSION MODAL ===== */}
{showPermissionModal && selectedFolder && (
  <Modal
    show={showPermissionModal}
    onHide={() => { setShowPermissionModal(false); setTogglePermission(undefined); }}
    size="xl"
    container={() => document.getElementById('filelistcontainer')}
    backdrop={true}
    centered
  >
    <Modal.Header closeButton>
      <Modal.Title>Manage Permission</Modal.Title>
    </Modal.Header>
    {/* <Modal.Body style={{ padding: '15px' }}> */}
    {/* //Ritik */}
    <Modal.Body style={{ padding: '15px', maxHeight: '60vh', overflow: 'auto' }}>
 
      {!togglePermission && (
        <div className="text-center p-4">Loading...</div>
      )}
 
      {togglePermission === "No" && (
        <div>
          <h6 style={{ marginBottom: '20px' }}>
            This folder is public. Would you like to make it private?
          </h6>
          <div style={{ display: "flex", gap: "10px" }}>
             {/* //Rohit 27/04/2026  ----start */}
            {/* <button type="button" className="btn btn-primary" onClick={fpSetPrivate}>
              Set Permission
            </button> */}
           
            <button type="button" className="btn btn-primary" onClick={async () => {
        // Show the permission UI first — add users before making private
        await fpFetchUsers(selectedFolder);
        setFolderPrivacyTableData([]);
        setTogglePermission("PendingPrivate"); // new intermediate state
      }}>
        Set Permission
      </button>
      {/* //Rohit 27/04/2026  ----end */}
            <button type="button" className="btn btn-secondary" onClick={() => { setShowPermissionModal(false); setTogglePermission(undefined); }}>
              Cancel
            </button>
          </div>
        </div>
      )}
 
      {/* {togglePermission === "Yes" && ( //Rohit 29/04/2026 */}
      {(togglePermission === "Yes" || togglePermission === "PendingPrivate") && (  //Rohit 29/04/2026
 
        <div>
          <div className="row mb-2">
            <div className="col-sm-8">
              {/* <h5 style={{ fontWeight: '600', color: '#4c4c4c' }}>Manage Permission</h5> */}
              <div className="font-12 text-muted">{fpPathState}</div>
            </div>
            <div className="col-sm-4 d-flex justify-content-end">
              <a href="#" onClick={(e) => {
                e.preventDefault();
                const newId = rowsForPermission.length ? rowsForPermission[rowsForPermission.length - 1].id + 1 : 0;
                setRowsForPermission(prev => [...prev, { id: newId, selectedUserForPermission: [], selectedPermission: "" }]);
              }}>
                {/* <img src={require("../assets/addnew.png")} alt="add" style={{ height: '30px' }} /> */}
              </a>
            </div>
          </div>
          <hr />
 
          {rowsForPermission.map((row) => (
            <div key={row.id} className="row mb-2 align-items-center">
              <div className="col-md-6">
                <Select
                  isMulti
                  options={folderPrivacyUsers}
                  value={row.selectedUserForPermission}
                  onChange={(sel: any) => setRowsForPermission(prev => prev.map(r => r.id === row.id ? { ...r, selectedUserForPermission: sel } : r))}
                  placeholder="Enter names or email addresses..."
                />
                {fpErrors[row.id]?.userSelect && <span className="text-danger font-12">{fpErrors[row.id].userSelect}</span>}
                {fpErrors[row.id]?.duplicate && <span className="text-danger font-12">{fpErrors[row.id].duplicate}</span>}
              </div>
              <div className="col-md-4">
                <Select
                  options={[
                    { value: "Full Control", label: "Full Control" },
                    { value: "Contribute", label: "Contribute" },
                    { value: "Edit", label: "Edit" },
                    { value: "Read", label: "Read" },
                      // ritik 23/2/26 for restricted view permission
                    { value: "Restricted View", label: "Restricted View" },
                    { value: "View", label: "View" },
                  ]}
                  value={row.selectedPermission || null}
                  onChange={(sel: any) => setRowsForPermission(prev => prev.map(r => r.id === row.id ? { ...r, selectedPermission: sel } : r))}
                  placeholder="Select Permission..."
                />
                {fpErrors[row.id]?.permissionSelect && <span className="text-danger font-12">{fpErrors[row.id].permissionSelect}</span>}
              </div>
              <div className="col-md-2">
                {rowsForPermission.length > 1 && (
                  <a href="#" onClick={(e) => { e.preventDefault(); setRowsForPermission(prev => prev.filter(r => r.id !== row.id)); }}>
                    {/* <img src={require("../assets/delemodal.png")} alt="delete" style={{ height: '30px' }} /> */}
                  </a>
                )}
              </div>
            </div>
          ))}
 
          <table className="mtbalenew mt-3">
            <thead>
              <tr>
                <th style={{ minWidth: '55px',maxWidth:'55px' }}>S.No.</th>
                <th>User/Groups</th>
                <th style={{ minWidth: '100px',maxWidth:'100px' }}>Permission</th>
                <th style={{ minWidth: '75px',maxWidth: '75px' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {folderPrivacyTableData
                .slice((fpCurrentPage - 1) * fpItemsPerPage, fpCurrentPage * fpItemsPerPage)
                .map((item: any, idx: number) => (
                  <tr key={item.itemId}>
                    {/* <td><span className="indexdesign">{idx + 1}</span></td> */}
                    {/* // ritik */}
                    <td style={{ minWidth: '55px',maxWidth:'55px' }}><span className="indexdesign">{(fpCurrentPage - 1) * fpItemsPerPage + idx + 1}</span></td>
                    <td >{item.value}</td>
                    <td style={{ minWidth: '100px',maxWidth:'100px' }}>{item.Permission}</td>
                    <td style={{ minWidth: '75px',maxWidth: '75px' }}>
                      {/* Ritik 20/02/2026 changes for folder permission manage add trash icon */}
  <button
    type="button"
    title="Remove Permission"
    onClick={() => fpHandleDeleteUser(item.userId, item.itemId, item.Permission)}
    style={{
      background: 'none',
      border: 'none',
      cursor: 'pointer',
      color: '#dc3545',
      padding: '0px', marginTop:'0px'
    }}
  >
    <FontAwesomeIcon icon={faTrash} size="sm" />
  </button>
</td>
{/* end of Ritik 20/02/2026 changes for folder permission manage add trash icon */}
                  </tr>
                ))}
            </tbody>
          </table>
 
          <div className="row mt-2">
            <div className="col-md-8">
              {Math.ceil(folderPrivacyTableData.length / fpItemsPerPage) > 1 && (
                <nav>
                  <ul className="pagination">
                    <li className={`page-item ${fpCurrentPage === 1 ? 'disabled' : ''}`}>
                      <a className="page-link" onClick={() => setFpCurrentPage(p => Math.max(p - 1, 1))}>«</a>
                    </li>
                    {Array.from({ length: Math.ceil(folderPrivacyTableData.length / fpItemsPerPage) }, (_, i) => i + 1).map(p => (
                      <li key={p} className={`page-item ${fpCurrentPage === p ? 'active' : ''}`}>
                        <a className="page-link" onClick={() => setFpCurrentPage(p)}>{p}</a>
                      </li>
                    ))}
                    <li className={`page-item ${fpCurrentPage === Math.ceil(folderPrivacyTableData.length / fpItemsPerPage) ? 'disabled' : ''}`}>
                      <a className="page-link" onClick={() => setFpCurrentPage(p => p + 1)}>»</a>
                    </li>
                  </ul>
                </nav>
              )}
            </div>
            <div className="col-md-4 d-flex justify-content-end gap-2">

              {/* //Rohit 27/04/2026 - start */}
              <button type="button" className="btncolorCreate1 me-2" onClick={async () => {
  if (togglePermission === ("PendingPrivate" as any)) {
    // Validate first — only make private if validation passes
    const isValid = fpValidate();
    if (!isValid) return; // Stop here, show errors, do NOT make private
    await fpSetPrivate();
    await fpHandleCreate();
  } else {
    await fpHandleCreate();
  }
}}>
  {/* //Rohit 27/04/2026 - end */}
                <span className="mb-1 mt-2" data-tooltip="Submit">
                  <img src={require("../assets/submit-new.png")} alt="submit" />
                </span>
              </button>
              <button type="button" className="btncolorCreate1" onClick={() => { setShowPermissionModal(false); setTogglePermission(undefined); }}>
                <span className="mb-1 mt-2" data-tooltip="Cancel">
                  {/* <img src="" alt="cancel" style={{color:"red", fontWeight:"bold"}}> X</img> */}
                </span>
              </button>
            </div>
          </div>
        </div>
      )}
 
    </Modal.Body>
  </Modal>
)}
{/* end of model code for folder permission manage added by Ritik 20/02/2026 */}
{/* // Ritik 20/2/26 for manage permission folder */}

                      {/* === Step 2 Popup (with People Picker) === */}
                      {showManagePermissionModal && (
                        <Modal show={showManagePermissionModal} onHide={() => setShowManagePermissionModal(false)} size="lg"  container={() => document.getElementById('filelistcontainer')} backdrop={true} centered>
                          <Modal.Header closeButton>
                            <Modal.Title>Manage Permission</Modal.Title>
                          </Modal.Header>
                          <Modal.Body>
                            {(() => {
                              const ctx = selectedFolder ? deriveFolderContext(selectedFolder) : { webUrl: "", serverRel: "", documentLibraryName: "" };
                              return (
                                <>
                                  <div style={{ marginBottom: 10, fontSize: 13, color: "#555" }}>
                                    <div><strong>Folder:</strong> {selectedFolder?.FolderName || selectedFolder?.Title || selectedFolder?.Name || ctx.serverRel}</div>
                                    <div><strong>Library:</strong> {ctx.documentLibraryName}</div>
                                    <div><strong>Web:</strong> {ctx.webUrl}</div>
                                  </div>

                                  {mpLoading ? (
                                    <div>Loading permissions...</div>
                                  ) : (
                                    <>
                                      {mpError && (
                                        <div style={{ color: "#b00020", marginBottom: 10, whiteSpace: "pre-wrap" }}>{mpError}</div>
                                      )}

                                      <div style={{ marginBottom: 10 }}>
                                        <strong>Inheritance:</strong>{" "}
                                        {mpHasUnique ? (
                                          <span style={{ color: "#b26a00" }}>This folder has unique permissions</span>
                                        ) : (
                                          <span style={{ color: "#2e7d32" }}>Inheriting from parent</span>
                                        )}
                                      </div>

                                      <div style={{ marginBottom: 10 }}>
                                        <strong>Your ability:</strong>{" "}
                                        {mpCanManage ? (
                                          <span style={{ color: "#2e7d32" }}>You can manage permissions</span>
                                        ) : (
                                          <span style={{ color: "#b00020" }}>You cannot manage permissions on this folder</span>
                                        )}
                                      </div>

                                      <div
                                        style={{
                                          borderTop: "1px solid #eee",
                                          paddingTop: 10,
                                          marginTop: 10,
                                          display: "flex",
                                          gap: 8,
                                          alignItems: "center",
                                          flexWrap: "wrap",
                                        }}
                                      >
                                        {/* <select
              value={newUser}
              onChange={(e) => setNewUser(e.target.value)}
              style={{ padding: "6px 8px", border: "1px solid #ccc", borderRadius: 4, minWidth: 320 }}
              disabled={!mpCanManage}
              >
              <option value="" disabled>
              Select a user…
              </option>
              {mpSiteUsers.map((u) => (
              <option key={u.id} value={u.loginName}>
              {u.title} {u.email ? `(${u.email})` : ""}
              </option>
              ))}
              </select> */}

                                        {/* Typeahead user picker (replaces static dropdown) */}
                                        <div ref={userSuggestRef} style={{ position: "relative", minWidth: 320 }}>
                                          <input
                                            type="text"
                                            placeholder="Type user name or email..."
                                            value={newUserDisplay || (mpSiteUsers.find(u => u.loginName === newUser)?.title || "")}
                                            onChange={(e) => {
                                              const val = e.target.value || "";
                                              setNewUserDisplay(val);
                                              setShowUserSuggestions(true);
                                              const q = val.trim().toLowerCase();
                                              const filtered = (mpSiteUsers || []).filter((u) => {
                                                return (
                                                  (u.title || "").toLowerCase().includes(q) ||
                                                  (u.email || "").toLowerCase().includes(q) ||
                                                  (u.loginName || "").toLowerCase().includes(q)
                                                );
                                              }).slice(0, 50);
                                              setMpUserSuggestions(filtered);
                                              // Clear previously selected loginName (we'll set login on selection)
                                              setNewUser("");
                                            }}
                                            onFocus={() => {
                                              setShowUserSuggestions(true);
                                              setMpUserSuggestions((mpSiteUsers || []).slice(0, 50));
                                            }}
                                            disabled={!mpCanManage}
                                            style={{ padding: "6px 8px", border: "1px solid #ccc", borderRadius: 4, minWidth: 320 }}
                                          />
                                          {showUserSuggestions && mpUserSuggestions && mpUserSuggestions.length > 0 && (
                                            <ul style={{ position: "absolute", top: "100%", left: 0, right: 0, zIndex: 9999, background: "#fff", border: "1px solid #ccc", borderRadius: 4, maxHeight: 200, overflowY: "auto", listStyle: "none", margin: 0, padding: 0 }}>
                                              {mpUserSuggestions.map((u) => (
                                                <li
                                                  key={u.id}
                                                  onMouseDown={(e) => { e.preventDefault(); }}
                                                  onClick={() => {
                                                    setNewUser(u.loginName);
                                                    setNewUserDisplay(`${u.title}${u.email ? ` (${u.email})` : ""}`);
                                                    setShowUserSuggestions(false);
                                                  }}
                                                  style={{ padding: "8px", cursor: "pointer", borderBottom: "1px solid #eee" }}
                                                >
                                                  <div style={{ fontSize: 13 }}>{u.title}</div>
                                                  <div style={{ fontSize: 12, opacity: 0.7 }}>{u.email || u.loginName}</div>
                                                </li>
                                              ))}
                                            </ul>
                                          )}
                                        </div>

                                        <select
                                          value={newPermission}
                                          onChange={(e) => setNewPermission(e.target.value)}
                                          style={{ padding: "6px 8px", border: "1px solid #ccc", borderRadius: 4 }}
                                          disabled={!mpCanManage}
                                        >
                                          <option value="Read">Read</option>
                                          <option value="Edit">Edit</option>
                                          <option value="Full Control">Full Control</option>
                                        </select>
                                        <button
                                          onClick={addUserToSelectedFolder}
                                          style={{ padding: "6px 10px", border: "1px solid #107c10", borderRadius: 4, background: "#107c10", color: "#fff" }}
                                          disabled={!mpCanManage || !newUser.trim()}
                                        >
                                          + Add User
                                        </button>
                                      </div>

                                      <div style={{ margin: "10px 0" }}>
                                        <strong>Current access:</strong>
                                        {permissionUsers?.length ? (
                                          <div style={{ marginTop: 8 }}>
                                            <div style={{ display: "grid", gridTemplateColumns: "70px 1fr 280px 120px", gap: 8, padding: "8px 12px", background: "#f6f7f8", border: "1px solid #e1e4e8", borderRadius: 6, fontWeight: 600 }}>
                                              <div>S.No.</div>
                                              <div>User/Groups</div>
                                              <div>Permission</div>
                                              <div>Action</div>
                                            </div>
                                            {permissionUsers.map((a: any, i: number) => (
                                              <div key={a.principalId} style={{ display: "grid", gridTemplateColumns: "70px 1fr 280px 120px", gap: 8, padding: "8px 12px", border: "1px solid #e1e4e8", borderTop: "none" }}>
                                                <div>{i + 1}</div>
                                                <div title={a.principalTitle}>{a.principalTitle}</div>
                                                <div>{a.roles && a.roles.length ? a.roles.join(", ") : "—"}</div>
                                                <div>
                                                  <button
                                                    style={{ padding: "4px 8px", border: "1px solid #ccc", borderRadius: 4, background: "#fff", cursor: "pointer" }}
                                                    onClick={() => removeUserFromSelectedFolder(a.principalId)}
                                                    disabled={!mpCanManage}
                                                  >
                                                    Remove
                                                  </button>
                                                </div>
                                              </div>
                                            ))}
                                          </div>
                                        ) : (
                                          <div style={{ fontSize: 13, color: "#666" }}>No direct role assignments found.</div>
                                        )}
                                      </div>
                                    </>
                                  )}
                                </>
                              );
                            })()}
                          </Modal.Body>
                        </Modal>
                      )}
                      {/* aman code manage folder permission */}


                      {/* Pagination controls - edited by Rohit 23/04/2026 */}
                      {filteredFiles.length > 0 && (
  <div
    style={{
      margin: "10px 14px 10px 0",
      display: "flex",
      alignItems: "center",
      justifyContent: "end",
      gap: "6px",
      flexWrap: "wrap",
    }}
  >
    {/* Records per page dropdown - Rohit 23/04/2026 */}
    
    <button
      type="button"
      onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
      disabled={currentPage === 1}
      style={{
        padding: "5px 12px", marginTop: '0px',
        borderRadius: "4px",
        border: "1px solid #ccc",
        background: currentPage === 1 ? "#eee" : "#1d4ed8",
        cursor: currentPage === 1 ? "not-allowed" : "pointer",
      }}
    >
      Prev
    </button>
    
    {/* //Rohit 23/04/2026 ---remove total count  */}
    {/* <span style={{ fontSize: "13px", color: "#555", marginLeft: "4px" }}>
  {paginationTotalCount} records
</span> */}


    {/* Page number pills */}
    {(() => {
      const total = Math.ceil(paginationTotalCount / pageSize) || 1;
      const pages: any[] = [];
      let start = Math.max(1, currentPage - 2);
      let end = Math.min(total, start + 4);
      if (end - start < 4) start = Math.max(1, end - 4);

      if (start > 1) {
        pages.push(
          <button key="p1" type="button"
            onClick={() => setCurrentPage(1)}
            style={{ padding: "4px 9px", marginTop: "0px", borderRadius: "4px", border: "1px solid #ccc", background: "#fff", color: "#333", cursor: "pointer", fontSize: "13px" }}>
            1
          </button>
        );
        if (start > 2) pages.push(<span key="e1" style={{ padding: "0 2px", fontSize: "13px", color: "#555" }}>…</span>);
      }

      for (let p = start; p <= end; p++) {
        pages.push(
          <button key={p} type="button"
            onClick={() => setCurrentPage(p)}
            style={{
              padding: "4px 9px", marginTop: "0px",
              borderRadius: "4px",
              border: "1px solid #ccc",
              background: currentPage === p ? "#1d4ed8" : "#fff",
              color: currentPage === p ? "#fff" : "#333",
              cursor: "pointer",
              fontWeight: currentPage === p ? "600" : "400",
              fontSize: "13px",
            }}>
            {p}
          </button>
        );
      }

      if (end < total) {
        if (end < total - 1) pages.push(<span key="e2" style={{ padding: "0 2px", fontSize: "13px", color: "#555" }}>…</span>);
        pages.push(
          <button key={`last${total}`} type="button"
            onClick={() => setCurrentPage(total)}
            style={{ padding: "4px 9px", marginTop: "0px", borderRadius: "4px", border: "1px solid #ccc", background: "#fff", color: "#1d4ed8", cursor: "pointer", fontSize: "13px" }}>
            {total}
          </button>
        );
      }
      return pages;
    })()}

    <button
      type="button"
      onClick={(e) => {
        e.preventDefault();
        setCurrentPage((p) =>
          p < Math.ceil(paginationTotalCount / pageSize) ? p + 1 : p
        );
      }}
      disabled={currentPage === Math.ceil(paginationTotalCount / pageSize)}
      style={{
        padding: "5px 12px", marginTop: '0px',
        borderRadius: "4px",
        border: "1px solid #ccc",
        background: currentPage === Math.ceil(paginationTotalCount / pageSize) ? "#eee" : "#1d4ed8",
        cursor: currentPage === Math.ceil(paginationTotalCount / pageSize) ? "not-allowed" : "pointer",
      }}
    >
      Next
    </button>
    <select
      value={pageSize}
      onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(1); }}
      style={{
        padding: "4px 8px",
        borderRadius: "4px",
        border: "1px solid #ccc",
        fontSize: "13px",
        cursor: "pointer",
        marginRight: "4px",
        height: "30px",
      }}
    >
      {[10, 20, 40, 80, 100].map((n) => (
        <option key={n} value={n}>{n}</option>
      ))}
    </select>

      {/* //Rohit 23/04/2026 total number of record showen */}
    {/* <span style={{ fontSize: "13px", color: "#555", marginLeft: "4px" }}>
      {paginationTotalCount} records
    </span> */}
  </div>
)}
{/* //Rohit 23/04/2026------end*/}
                    </>
                  )
                  }

                </>
              )}
          {/* srs 29/1/26 */}
    </>
  )}
    {/* srs 29/1/26 */}

              {/* //update by ritik 29/01/2026 for preview file title and path display */}


              {showPreviewModal && (
                <PreviewModal
                  show={true}
                  fileUrl={previewFile}
                  fileName={previewFile?.FileName || previewFile?.Name}
                  filePath={
                    activeView
                      ? undefined
                      : (breadcrumbs.map(b => b.title).join(' › '))
                  }
                  viewName={
                    activeView
                      ? breadcrumbs[breadcrumbs.length - 1]?.title
                      : undefined
                  }
                  onClose={() => {
                    setShowPreviewModal(false);
                    setPreviewFile(null);
                    setDirectDownloadFile(null); // Ritik added 23/2/26
                    isDeepLinkPreview.current = false; // addhyan 16/4/26
                  }}
                />
              )}

              {/* meta data model Render Ritik 29/01/2026 */}

              {/* Metadata Modal */}
              {showMetadataModal && (
                <MetadataModal
                    show={showMetadataModal}
        
                  file={selectedFileForMetadata}
                  context={context}
                  onClose={() => {
                    setShowMetadataModal(false);
                    setSelectedFileForMetadata(null);
                  }}
                />
              )}

              {showManageWorkflow && selectedWorkflowFolder && (
                <ManageWorkflow
                  context={context}
                  siteUrl={selectedWorkflowFolder.SiteUrl}   // ✅ FIX
                  siteTitle={selectedWorkflowFolder.SiteTitle}
                  documentLibraryName={selectedWorkflowFolder.DocumentLibraryName}
                  file={selectedWorkflowFolder.file}
                  onClose={() => {
                    setShowManageWorkflow(false);
                    setSelectedWorkflowFolder(null);
                  }}
                />
              )}


            </div>
            </div></div></div>
          {/* ---------------------------Audit History Modal ---------------------*/}
         {showAuditModal && (
  <Modal 
  show={showAuditModal} 
  onHide={() => setShowAuditModal(false)} 
  size="xl"
  container={() => document.getElementById('filelistcontainer')}
  backdrop={true}
>
    <Modal.Header closeButton>
      <Modal.Title>Audit History</Modal.Title>
    </Modal.Header>
    <Modal.Body style={{ maxHeight: '70vh', padding:'0px', overflowY: 'auto' }}>
      {auditLoading ? (
        <div style={{ textAlign: 'center', padding: '20px' }}>Loading...</div>
      ) : (
        <div>
          {/* Metadata Section */}
          <div style={{ marginBottom: '10px' }}>
            <div style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(4, 1fr)',
              gap: '15px',
              padding: '20px',
              backgroundColor: '#fff',
              borderRadius: '0px',
              border: '0px solid #dee2e6'
            }}>
              {Object.entries(auditVersions.Metadata || {}).filter(([key]) => key !== 'Title') .map(([key, value]) => (
                <div key={key}>
                  <div style={{
                    fontSize: '14px',
                    color: '#4a4a4a',
                    marginBottom: '0px',
                    fontWeight: '600'
                  }}>
                    {key}:
                  </div>
                  <div style={{
                    fontSize: '14px',
                    fontWeight: '500',
                    color: '#9f9f9f',
                    wordBreak: 'break-word'
                  }}>
                    {key === 'IsDeleted' ? (
                      <span style={{
                        padding: '2px 8px',
                        borderRadius: '4px',
                        backgroundColor: value === 'No' ? '#d4edda' : '#f8d7da',
                        color: value === 'No' ? '#155724' : '#721c24',
                        fontSize: '13px'
                      }}>
                        {String(value)}
                      </span>
                    ) : key === 'Status' ? (
                      <span style={{
                        padding: '2px 8px',
                        borderRadius: '4px',
                        backgroundColor:
                          value === 'Pending' ? '#fff3cd' :
                          value === 'Approved' ? '#d4edda' :
                          value === 'Rejected' ? '#f8d7da' : '#e2e3e5',
                        color:
                          value === 'Pending' ? '#856404' :
                          value === 'Approved' ? '#155724' :
                          value === 'Rejected' ? '#721c24' : '#383d41',
                        fontSize: '13px'
                      }}>
                        {String(value)}
                      </span>
                    ) : (
                      String(value || '-')
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>
 
          {/* Approval Section */}
          <div style={{padding:'0px 18px 12px 18px'}}>
            <h6 style={{
              fontSize: '16px',
              fontWeight: '600',
              marginBottom: '5px',
              paddingBottom: '8px',
              borderBottom: '0px solid #0078d4'
            }}>
              Approval Details 
            </h6>
            {auditVersions.Versions && auditVersions.Versions.length ? (
              <div style={{ overflowX: 'auto' }}>
                <table style={{
                  width: '100%',
                  borderCollapse: 'collapse',
                  border: '1px solid #dee2e6'
                }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f5f7fa', color: '#000' }}>
                      <th style={{
                        padding: '10px',
                        textAlign: 'left',
                        fontSize: '13px',
                        fontWeight: '600',
                        borderBottom: '0px solid #005a9e'
                      }}>
                        Approval Level
                      </th>
                      <th style={{
                        padding: '10px',
                        textAlign: 'left',
                        fontSize: '13px',
                        fontWeight: '600',
                        borderBottom: '0px solid #005a9e'
                      }}>
                        Approver
                      </th>
                      <th style={{
                        padding: '10px',
                        textAlign: 'left',
                        fontSize: '13px',
                        fontWeight: '600',
                        borderBottom: '0px solid #005a9e'
                      }}>
                        Action Date
                      </th>
                      {/* <th style={{
                        padding: '10px',
                        textAlign: 'left',
                        fontSize: '13px',
                        fontWeight: '600',
                        borderBottom: '0px solid #005a9e'
                      }}>
                        Email
                      </th> */}
                      <th style={{
                        padding: '10px',
                        textAlign: 'right',
                        fontSize: '13px',
                        fontWeight: '600',
                        borderBottom: '0px solid #005a9e'
                      }}>
                       Status
                      </th>
                      <th style={{
                        padding: '10px',
                        textAlign: 'right',
                        fontSize: '13px',
                        fontWeight: '600',
                        borderBottom: '0px solid #005a9e'
                      }}>
                       Remark
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {auditVersions.ApprovalData.map((v: any, i: number) => (
                      <tr key={i} style={{
                        borderBottom: '1px solid #dee2e6',
                        backgroundColor: i % 2 === 0 ? '#fff' : '#f5f7fa'
                      }}>
                        <td style={{
                          padding: '10px',
                          fontSize: '13px',
                          fontWeight: '600',
                          textAlign: 'left'
                        }}>
                          Level {v.MasterApproval?.Level || "-"}
                          
                        </td>
                        <td style={{ padding: '10px', fontSize: '13px' }}>
                          {v.CurrentUser || "-"}
                         
                        </td>
                        <td style={{ padding: '10px', fontSize: '13px' }}>
                          {/* {v.LogHistory || "-"} */}
                          {v.LogHistory
  ? new Date(v.LogHistory).toLocaleDateString("en-GB")
  : "-"}
                        </td>
                        {/* <td style={{
                          padding: '10px',
                          fontSize: '13px',
                          color: '#6c757d'
                        }}>
                          {v.ModifiedByEmail || '-'}
                        </td> */}
                        <td style={{
                          padding: '10px',
                          fontSize: '13px',
                          textAlign: 'right'
                        }}>
                          {v.FileUID?.Status || "-"}
                        </td>
                        <td style={{
                          padding: '10px',
                          fontSize: '13px',
                          textAlign: 'right'
                        }}>
                          {v.Remark || "-"}
                          
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div style={{
                padding: '20px',
                backgroundColor: '#d1ecf1',
                border: '0px solid #bee5eb',
                borderRadius: '4px',
                color: '#0c5460',
                fontSize: '14px'
              }}>
                No Approval available
              </div>
            )}
          </div>
        </div>
      )}
    </Modal.Body>
    {/* <Modal.Footer>
      <Button variant="secondary" onClick={() => setShowAuditModal(false)}>
        Close
      </Button>
    </Modal.Footer> */}
  </Modal>
)}
          {/* ================= RENAME FILE MODAL – ADD EXACTLY HERE ================= */}
          <Modal
            show={renameFileModalOpen}
            onHide={() => setRenameFileModalOpen(false)}
            container={() => document.getElementById('filelistcontainer')}
            backdrop={true}
            centered
          >
            <Modal.Header closeButton>
              <Modal.Title>Rename File</Modal.Title>
            </Modal.Header>

            <Modal.Body>
              <input
                type="text" style={{padding:'6px 10px'}}
                className="form-control"
                value={renameFileValue}
                onChange={(e) => setRenameFileValue(e.target.value)}
                autoFocus
              />
            </Modal.Body>

            <Modal.Footer>
              {/* <Button
                variant="secondary"
                onClick={() => setRenameFileModalOpen(false)}
              >
                Cancel
              </Button> */}
              {/* <Button
                variant="primary"
                onClick={submitRenameFile}
                disabled={!renameFileValue?.trim()}
              >
                Rename
              </Button> */}
<Button
  variant="link"
  onClick={submitRenameFile}
  disabled={!renameFileValue?.trim()}
  style={{ padding: 0, border: 'none', width:'50px', height:'50px', borderRadius:'50%',display:'flex', background:'#fff',boxShadow:'rgba(0, 0, 0, 0.1) 0px 2px 8px',alignItems:'center',justifyContent:'center' }}
>
<img 
    src={require("../assets/submit-new1.png")} 
    alt="Rename" 
    style={{ 
      height: "32px",
      opacity: !renameFileValue?.trim() ? 0.5 : 1 
    }} 
  />
</Button>

            </Modal.Footer>
          </Modal>
          {/* ================= RENAME FILE MODAL END ================= */}



          {/* sourish 30/9/25 two div added*/}
        </div>
      </div>
    </div>
  );
};

const DMSMain: React.FC<IEssadmsMainProps> = (props) => {
  return (
    <Provider>
      <ArgPoc context={props.context} />;
    </Provider>
  )
};

export default DMSMain;
