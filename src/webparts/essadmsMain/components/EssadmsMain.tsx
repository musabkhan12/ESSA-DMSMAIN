import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IEssadmsMainProps } from "./IEssadmsMainProps";
import PreviewModal from "./previewfile"; // import your modal\
import DirectDownloader from "./DownloadFile";
import VersionHistoryModal from "./versionhistory"; // import version history modal
import Swal from 'sweetalert2'
import {  Button, Modal } from 'react-bootstrap';   
import CreateFolder from "./CreateFolder";
import Revoke from "./revoke";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
//aman changes for folder permission manage
import { PermissionKind } from "@pnp/sp/security";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
//aman changes for folder permission manage
import { useMemo } from "react";

// sourish 30/9/25
import "../../verticalSideBar/components/VerticalSidebar.scss";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import UserContext from "../../../GlobalContext/context";
import Provider from "../../../GlobalContext/provider";

// import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// let loadfilefromnode = ''
declare global {
  interface Window {
    managePermission?: (folder: any) => void;
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
// this was by om revoke user permission
interface SharedUser {
  userId: string;
  permission: string;
}

interface BreadcrumbItem {
  key: string;
  title: string;
  type: string;
  siteUrl: string;
  libraryTitle?: string;
  folderPath?: string;
}

const ArgPoc = ({ context }: { context: WebPartContext }) => {
  const [sp] = useState(() => spfi().using(SPFx(context)));
  const [filesLoadedfromnode, setFilesLoadedloadedfromnode] = useState(false);

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
  // state variables for current site URL
  const [currentSiteUrl, setCurrentSiteUrl] = useState<string>("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  // preview file state
  const [previewFileUrls, setPreviewFileUrls] = useState<string[]>([]);
  // modal state
  const [menuOpenIdx, setMenuOpenIdx] = useState<number | null>(null);

  // state for selected node
  const [selectedCurrentNode, setSelectedCurrentNode] = useState<TreeNode | null>(null);
  const [showAuditModal, setShowAuditModal] = useState(false);
   const [showShareModal, setShowShareModal] = useState(false);
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
  // preview file state
  const [previewFile, setPreviewFile] = useState<any>(null);
  // preview modal state
  const [showPreviewModal, setShowPreviewModal] = useState(false);
  
    //sourish 21/8/25
  const [renameModalOpen, setRenameModalOpen] = useState(false);
  const [renameValue, setRenameValue] = useState("");

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

  {/* sourish 30/9/25 */}
  const { useHide }: any = React.useContext(UserContext);
// sourish 3/10/25
  const [activeLayout, setActiveLayout] = useState<'grid' | 'list'>('grid');
// Pagination state
  const [currentPage, setCurrentPage] = useState(1);
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
const handleUndoDelete = async (file: any) => {
    try {
      const siteUrl: string = `${file.__siteUrl}/${file.SiteName}`; // here we are trying to get our subsite context
      const listTitle: string = file.__fileMasterList || "FileMaster";
      const id: number = file.Id || file.ID;
      console.log("FILE TO DELETE:", file);
      console.log("Undo Restore Triggered for");
      console.log("Site URL ", siteUrl);
      console.log("list title", listTitle);
      console.log("item id", id);
      debugger;
      try {
    const spnew = spfi(siteUrl).using(SPFx(context));
    console.log("file.CurrentFolderPath", file.CurrentFolderPath);
      const fileItem = await spnew.web.getFileByServerRelativePath(`${file.CurrentFolderPath}/${file.FileName}`).getItem();
    let payload:any={
      IsDeleted:null
    }
      const itemData = await fileItem.update(payload)
    console.log("column updated successfully",itemData);
     } catch (err) {
    console.error("Error updating file in document library:", err);
    Swal.fire({
  title: 'Error!',
  text: 'File restore failed. Please try again.' + err,
  icon: 'error',
  confirmButtonText: 'Cool'
   })
      }
      
      try {
        const siteurl = `${file.__siteUrl}`  // here we are trying to get our site collection context
      const siteSP = spfi(siteurl).using(SPFx(context));
      await siteSP.web.lists.getByTitle(listTitle).items.getById(id).update({ IsDeleted: null });

        const refreshed = await loadViewData("RecycleBin");
      if (refreshed) {
        setSelectedFiles([...refreshed]); // spread → force re-render
      }
    
      } catch (err) {
  console.error("Error updating file metadata:", err);
    Swal.fire({
  title: 'Error!',
  text: 'File restore failed. Please try again.' + err,
  icon: 'error',
  confirmButtonText: 'Cool'
   })
      }
    

    } catch (e) {
      console.error("Undo (restore) failed:", e);
    }
  }

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
 


useEffect(() => {
  const initialize = async () => {
    console.log("[initialize] Start");
    await loadRootSites();

    // Check URL hash on initial load
    const rawHash = window.location.hash.substring(1);
    const hash = decodeURIComponent(rawHash || "");
    console.log("[initialize] Hash:", rawHash, "decoded:", hash);

    if (hash) {
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
  useEffect(() => {
    window.managePermission = (folder: any) => {
      setSelectedFolder(folder);
      setShowPermissionModal(true);
    };
    return () => {
      if (window.managePermission) window.managePermission = undefined;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

   //addhyan chanes on three dot should hide on click anywhere outside oif screen
  // three dot menu 

    useEffect(() => {
    if (menuOpenIdx !== null) {
      const handleClickOutside = (event: MouseEvent) => {
        // Only close if the click is outside any menu
        const menus = document.querySelectorAll('.three-dot-menu');
        let clickedInside = false;
        menus.forEach(menu => {
          if (menu.contains(event.target as Node)) {
            clickedInside = true;
          }
        });
        if (!clickedInside) {
          setMenuOpenIdx(null);
        }
      };
      document.addEventListener('mousedown', handleClickOutside);
      return () => {
        document.removeEventListener('mousedown', handleClickOutside);
      };
    }
  }, [menuOpenIdx]);



const loadRootSites = async () => {
  try {
    console.log("[loadRootSites] Fetching master lists...");
    const [siteList, subsiteList] = await Promise.all([
      sp.web.lists.getByTitle("MasterSiteCollection").items.select("Id", "Title", "SiteURL").top(5000)(),
      sp.web.lists.getByTitle("MasterSiteURL").items.select("Id", "Description", "Title", "SiteURL", "Active", "SiteID", "FileMasterList", "Function").top(5000)(),
    ]);

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
    await loadFilesForNode(currentNode);
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

  // this below lines will close preview modal  (addhyan work on this )
   setShowPreviewModal(false); // <-- Close preview modal
  setPreviewFile(null);       // <-- Clear preview file
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
        setSelectedFiles(files);
        
        console.log("[loadFilesForNode] Selected files set:", selectedFiles);
        setBreadcrumbs(getNodePath(node.key));

        console.log("Breadcrumbs updated:", breadcrumbs);
        console.log("Breadcrumbs updated 2:" +JSON.stringify(getNodePath(node.key)) )
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
        const libs = await siteSP.web.lists.filter("BaseTemplate eq 101 and Hidden eq false").select("Title")();
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
    console.log("[breadcrumb] Clicked:", item);
    if (item.type === "view") {
      handleViewButtonClick(item.title);
    } else {
      const node = nodeMap[item.key];
      if (node) {
        await expandPathToNode(node.key);

        if (node.type === "subsite") {
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

    try {
      let files: any[] = [];
      switch (viewName) {
        case "My request":
          files = await loadViewData("MyRequest");
          break;
        case "My favourite":
          files = await loadViewData("MyFavourite");
          break;
          case "My Folders":
          files = await loadViewData("MyFolders");
          break;
        case "Share with me":
          files = await loadViewData("SharedWithMe");
          break;
        case "Share with other":
          files = await loadViewData("SharedWithOthers");
          break;
        case "Recycle bin":
          files = await loadViewData("RecycleBin");
          break;
      }
      console.log(`[view] Loaded ${viewName} data:`, files);
      setSelectedFiles(files);
          console.log("[loadFilesForNode] Selected files set:", selectedFiles);
      setCurrentFolderPath("");
      setCurrentSiteUrl("");
    } catch (error) {
      console.error(`[view] Error loading ${viewName} data:`, error);
      setSelectedFiles([]);
    }
  };

  const loadViewData = async (viewType: string): Promise<any[]> => {
    setShowPreviewModal(false)
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

        
        const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL")
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
            } catch {}

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
            const cleaned = normalized.filter((row: any) => {
              const name = row.FolderName || row.Title || "";
              const path = row.FolderPath || row.folderpath || row.ServerRelativeUrl || "";
              const lowerPath = (path || "").toLowerCase().replace(/\/+$/, "");
              const isForms = /\/forms(\/|$)/i.test(lowerPath);
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

        const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite")
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

              const filtered = (listItems || []).filter(
                (it: any) =>
                  (it.CurrentUser || "").toLowerCase() === meEmail.toLowerCase() &&
                  (it.ShareWithMe || "").toLowerCase() !== meEmail.toLowerCase()
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

        const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite")
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
              const filtered = (listItems || []).filter(
                (it: any) =>
                  (it.CurrentUser || "").toLowerCase() === meEmail.toLowerCase()
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
    const configItems = await sp.web.lists
      .getByTitle("Allsitesfilemaster")
      .items.select("FileMaster", "sitecollection")
      .top(5000)();

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

              if (viewType === "MyFavourite") {
                // 🔹 Filter favourites
                return items.filter((item: any) => item.IsFavourite === true);
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
              return (items || []).map((it: any) => ({
                ...it,
                __siteUrl: siteUrl,
                __fileMasterList: config.FileMaster,
              }));
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
    return [];
  };
 // #region Toggle IsFav List
  const toggleFavourite = async (file: any) => {
    try {
      // Build site URL dynamically
      const currentUrl = new URL(context.pageContext.web.absoluteUrl);
      const siteUrl = `${currentUrl.protocol}//${currentUrl.hostname}/sites/${file.SiteCollection}`;
      const siteSP = spfi(siteUrl).using(SPFx(context));
 
      // Update SharePoint item
      await siteSP.web.lists
        .getByTitle(file.ListName)
        .items.getById(file.Id)
        .update({
          IsFavourite: !file.IsFavourite, // flip value
        });
 
      console.log(
        `Updated favourite for ${
          file.Title || file.FileName
        } → ${!file.IsFavourite}`
      );
 
      // Refresh current view (example: MyFavourite)
      await loadViewData("MyFavourite");
      alert("Done!");
      // Close modal if you want
      // setIsModalOpen(null);
    } catch (error) {
      console.error("Error toggling favourite:", error);
    }
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

  const uploadFiles = async () => {
    if (!currentSiteUrl || !currentFolderPath || selectedUploadFiles.length === 0) return;

    try {
      const siteSP = spfi(currentSiteUrl).using(SPFx(context));

      for (let i = 0; i < selectedUploadFiles.length; i++) {
        const file = selectedUploadFiles[i];
        setUploadProgress(((i + 1) / selectedUploadFiles.length) * 100);

        const uplaodfile = await siteSP.web
          .getFolderByServerRelativePath(
            `/sites/${currentSiteUrl.split("/sites/")[1]}/${currentFolderPath}`
          )
          .files.addChunked(file.name, file);
        console.log("[upload] File uploaded:", uplaodfile.data, JSON.stringify(uplaodfile.data));

        const encodeSharePointPath = (path: string): string => {
          return encodeURIComponent(path)
            .replace(/'/g, "%27")
            .replace(/\(/g, "%28")
            .replace(/\)/g, "%29")
            .replace(/\*/g, "%2A")
            .replace(/!/g, "%21")
            .replace(/#/g, "%23")
            .replace(/\$/g, "%24")
            .replace(/&/g, "%26")
            .replace(/\+/g, "%2B")
            .replace(/,/g, "%2C")
            .replace(/;/g, "%3B")
            .replace(/=/g, "%3D")
            .replace(/\?/g, "%3F")
            .replace(/\[/g, "%5B")
            .replace(/\]/g, "%5D")
            .replace(/_/g, "%5F")
            .replace(/\./g, "%2E")
            .replace(/-/g, "%2D");
        };

        const getSharePointPreviewUrl = (siteUrl: string, serverRelativePath: string): string => {
          const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf("/"));
          const encodedFilePath = encodeSharePointPath(serverRelativePath);
          const encodedParentPath = encodeSharePointPath(parentFolder);

          return `${siteUrl}${parentFolder.replace(/ /g, "%20")}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodedParentPath}`;
        };

        const siteUrl = "https://officeindia.sharepoint.com";
        const filePath = uplaodfile.data.ServerRelativeUrl;
        const previewUrl = getSharePointPreviewUrl(siteUrl, filePath);
        console.log("[upload] Preview URL:", previewUrl);
      }

      // Refresh the file list
      if (breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== "view") {
        const lastNode = nodeMap[breadcrumbs[breadcrumbs.length - 1].key];
        if (lastNode) {
          await loadFilesForNode(lastNode);
        }
      }

      setSelectedUploadFiles([]);
      setPreviewFileUrls([]);
      setShowUploadPanel(false);
      setUploadProgress(0);
    } catch (error) {
      console.error("[upload] Error uploading files:", error);
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
                backgroundColor: breadcrumbs.some((b) => b.key === node.key) ? "#f0f0f0" : "transparent",
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
const handleAuditHistory = async (file: any) => {
    try {
      setAuditLoading(true);
      setModalFile(file);
      setShowAuditModal(true);

      /*---------------build siteUrl for handleAuditHistory---------------*/
      let siteUrl: string = (file && (file.__siteUrl || file.SiteUrl || file.SiteURL)) || currentSiteUrl || context.pageContext?.site?.absoluteUrl;
      if (file && file.SiteName) {
        siteUrl = `${siteUrl.replace(/\/$/, '')}/${file.SiteName.replace(/^\//, '')}`;
      }

      const spnew = spfi(siteUrl).using(SPFx(context));

      console.log('Audit history (File)');
      console.log('fileUID', file?.FileUID || file?.fileUID || file?.UniqueId || file?.Id || file?.ID);
      console.log('siteUrl', siteUrl);

      // determine server relative path robustly
      const srCandidates = [
        file?.ServerRelativeUrl,
        file?.ServerRelativeUrlDecoded,
        (file?.ServerRelativePath && file.ServerRelativePath.DecodedUrl) ? file.ServerRelativePath.DecodedUrl : null,
        file?.FileRef,
        file?.FileLeafRef ? (file.FileRef ? file.FileRef : null) : null,
        (file?.CurrentFolderPath && file?.FileName) ? `${file.CurrentFolderPath}/${file.FileName}` : null,
        file?.ServerRelativePath
      ].filter(x => x !== undefined && x !== null);

      let serverRel: any = null;
      if (srCandidates.length > 0) {
        serverRel = srCandidates.find(s => !!s) || null;
        if (serverRel && typeof serverRel === 'string' && !serverRel.startsWith('/')) {
          serverRel = '/' + serverRel;
        }
      }

      console.log('serverRel candidate', serverRel);

      let fileItem: any = null;
      let listItem: any = null;
      let versions: any[] = [];

      if (serverRel) {
        try {
          fileItem = await spnew.web.getFileByServerRelativePath(serverRel).getItem();
          console.log('fileItem from getFileByServerRelativePath', fileItem);
        } catch (e) {
          console.warn('getFileByServerRelativePath.getItem() failed', e);
        }

        try {
          const fileObj = spnew.web.getFileByServerRelativePath(serverRel);
          const vers = await fileObj.versions();
          versions = (vers || []).map((v: any) => ({
            VersionLabel: v?.VersionLabel || v?.Version || '',
            Created: v?.Created ? new Date(v.Created).toLocaleString() : '',
            ModifiedByName: v?.CreatedBy?.Title || v?.CreatedBy?.Name || '',
            ModifiedByEmail: v?.CreatedBy?.Email || v?.CreatedBy?.LoginName || '',
            SizeDisplay: (v?.Size ? (v.Size/1024) : v?.Length ? (v.Length/1024) : 0).toFixed(2) + ' KB'
          }));
        } catch (e) {
          console.warn('versions() fetch failed', e);
        }
      }

      // if fileItem not found, try to get list item from FileMaster list using Id
      const listTitle = file.__fileMasterList || file.FileMaster || 'FileMaster';
      const id = file.Id || file.ID;
      if ((!fileItem || (Object.keys(fileItem || {}).length === 0)) && listTitle && id) {
        try {
          listItem = await spnew.web.lists.getByTitle(listTitle).items.getById(id).select('*', 'Editor/Title', 'Editor/Email').expand('Editor')();
          console.log('list item from FileMaster list', listItem);
        } catch (e) {
          console.warn('fetch FileMaster list item failed', e);
        }
      }

      // Compose metadata using fileItem, listItem, or the file object
      const src = fileItem || listItem || file || {};
      console.log('source for metadata', src);
      const metadata: any = {
        Title: src?.Title || file?.Title || file?.FileName || '',
        FileName: src?.FileLeafRef || file?.FileName || '',
        FileUID: file?.FileUID || file?.UniqueId || file?.GUID || '',
        Status: src?.Status || file?.Status || '',
        IsDeleted: (src?.IsDeleted ?? file?.IsDeleted ?? listItem?.IsDeleted) ? 'Yes' : 'No',
        Modified: src?.Modified ? new Date(src.Modified).toLocaleString() : (file?.Modified ? new Date(file.Modified).toLocaleString() : ''),
        ModifiedBy: (src?.Editor?.Title || src?.Editor?.Name || file?.ModifiedBy || '')
      };
       console.log('composed metadata', metadata);
      // If no versions found, try to synthesize single current entry
      if ((!versions || versions.length === 0) && (fileItem || listItem)) {
        const s = fileItem || listItem;
        versions = [{
          VersionLabel: s?.OData__UIVersionString || s?.VersionLabel || '1.0',
          Created: s?.Modified ? new Date(s.Modified).toLocaleString() : '',
          ModifiedByName: s?.Editor?.Title || '',
          ModifiedByEmail: s?.Editor?.Email || '',
          SizeDisplay: s?.Length ? (s.Length/1024).toFixed(2) + ' KB' : ''
        }];
      }

      setAuditVersions({ Metadata: metadata, Versions: versions });
    } catch (err) {
      console.error('Audit history fetch error:', err);
      setAuditVersions({ Metadata: {}, Versions: [] });
    } finally {
      setAuditLoading(false);
    }
  }
  /* ----------------- end handleAuditHistory ------------------ */
  

    // sourish 20/8/25
    const deleteFolder = async (file: any) => {
      try {
        if (!file?.FolderPath || !file?.__siteUrl) {
          console.error("Missing folder path or site URL");
          return;
        }
  
        // Delete Folder
  
        // Build correct web URL for the subsite
        const fullWebUrl = `${file.__siteUrl}/${encodeURIComponent(file.SiteTitle)}`;
        const sp = spfi(fullWebUrl).using(SPFx(context));
  
        // FolderPath is already server-relative, so just use it directly
        const finalPath = file.FolderPath;
  
        console.log("Deleting from:", fullWebUrl);
        console.log("Final server-relative path:", finalPath);
  
        await sp.web.getFolderByServerRelativePath(finalPath).delete();
  
  
        // Delete corresponding item from list in site collection root
        const spRoot = spfi(file.__siteUrl).using(SPFx(context));
  
        console.log("Deleting list item from site:", file.__siteUrl, " List: DMSFolderMaster, ID:", file.ID);
  
        await spRoot.web.lists.getByTitle("DMSFolderMaster").items.getById(file.ID).delete();
  
        const refreshed = await loadViewData("MyFolders");
        setSelectedFiles(refreshed);
        setActiveView("My Folders");
        setBreadcrumbs([
          { key: "my-folders", title: "My Folders", type: "view", siteUrl: "" },
        ]);
        setCurrentPage(1);
  
  
      } catch (err) {
        console.error("Error deleting folder:", err);
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
      } catch {}
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
    const deleteFileFolder = async (file: any, siteUrl: string, context: any) => {
    try {
      const siteSP = spfi(siteUrl).using(SPFx(context));
 
      if (!file?.ServerRelativeUrl) {
        console.error("[deleteFile] No ServerRelativeUrl found on file:", file);
        return;
      }
 
      console.log("[deleteFile] Deleting:", file.ServerRelativeUrl);
 
      await siteSP.web
        .getFileByServerRelativePath(file.ServerRelativeUrl)
        .delete();
 
      console.log("[deleteFile] File deleted successfully:", file.Name);
      alert("File Deleted");
 
      // Optionally, refresh your list after delete
      // await loadFilesForNode(currentNode);
    } catch (err) {
      console.error("[deleteFile] Error deleting file:", err);
    } 
    }



    // toggle favourite in my folder
    const toggleFavouriteDoc = async (
    file: any,
    siteUrl: string,
    context: any
  ) => {
    try {
      const siteSP = spfi(siteUrl).using(SPFx(context));
 
      if (!file?.ServerRelativeUrl) {
        console.error(
          "[toggleFavourite] No ServerRelativeUrl found on file:",
          file
        );
        return;
      }
 
      // Get the list item backing this file
      const item = await siteSP.web
        .getFileByServerRelativePath(file.ServerRelativeUrl)
        .getItem();
 
      // Read current favourite value
      const currentItem = await item.select("Id", "IsFavourite")();
      const currentFav = currentItem?.IsFavourite || false;
 
      console.log(`[toggleFavourite] Current favourite: ${currentFav}`);
 
      // Toggle
      await item.update({
        IsFavourite: !currentFav,
      });
 
      console.log(
        `[toggleFavourite] File ${file.Name} is now ${
          !currentFav ? "marked as favourite" : "unmarked as favourite"
        }`
      );
 
      alert(
        ` File ${file.Name} is now ${
          !currentFav ? "marked as favourite" : "unmarked as favourite"
        }`
      );
 
      // Optional: return new status
      return !currentFav;
    } catch (err) {
      console.error("[toggleFavourite] Error toggling favourite:", err);
    }
  };

  // abhay change for seach in all tabs
 // #region SEARCH TAB 
  // Abhay 10/10/25 merger 

const [searchTerm, setSearchTerm] = useState("");

const filteredFiles = !searchTerm.trim()
  ? selectedFiles
  : selectedFiles.filter((file: any) => {
      const term = searchTerm.toLowerCase();

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

  const pageSize = 12;
  let location: string = "";
  const paginatedFiles = useMemo(() => {
  const start = (currentPage - 1) * pageSize;
  const end = start + pageSize;
  return filteredFiles.slice(start, end);
  }, [filteredFiles, currentPage]);
  
useEffect(() => {
  setCurrentPage(1);
}, [searchTerm]);

// 🧩 ✅ Clear search when section changes
useEffect(() => {
  setSearchTerm("");
  setCurrentPage(1);
}, [activeView]);















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
      <HorizontalNavbar _context={sp}/>
      <div className="content" style={{marginLeft: `${!useHide ? '240px' : '80px'}`,marginTop:'0.8rem'}}>
     {/* <div className="content-page">
      <HorizontalNavbar _context={sp}/>
      <div className="content" style={{marginLeft: `${!useHide ? '240px' : '80px'}`,marginTop:'0.8rem'}}> */}
      {/* Left Panel with Quick Views and Folder Hierarchy */}
      <div style={{display:'flex'}}>
      <div className="inbox-leftbar">
        {/* Quick Views Panel */}
        <div
          id="buttonpanel"
          style={{
            padding: "0px 15px",
         
          
            flexShrink: 0,
          }}
        >
          <h2 className="page-title fw-bold mb-3 pt-5 mt-0 font-20"> Quick Views</h2>
          {[
            "My request",
            "My favourite",
            "My Folders",
            "Share with me",
            "Share with other",
            "Recycle bin",
          ].map((view) => (
            <button type="button"
              key={view}
              onClick={() => handleViewButtonClick(view)}
              style={{
                display: "block",
                width: "100%",
                padding: "8px 10px",
                marginBottom: "8px",
                textAlign: "left",
                backgroundColor: activeView === view ? "#0078d4" : "#f7fbfc ",
                color: activeView === view ? "white" : "#6c757d",
                border:"0px solid #ccc",
                borderRadius: "30px",
                cursor: "pointer",
                fontSize: "14px", 
                transition: "all 0.2s",
              }}
            >
              {view}
            </button>
          ))}
        </div>

        {/* Folder Hierarchy Panel */}
        <div
          id="folderhierarchycontainer"
          style={{
            flexGrow: 1,
            padding: "15px",
            overflow: "auto",
           
          }}
        >
          <h2
            style={{
              fontSize: "16px",
              fontWeight: "600",
              marginBottom: "15px",
              color: "#333",
              paddingBottom: "5px",
              borderBottom: "1px solid #eee",
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

      {/* File List Panel */}
      <div
        id="filelistcontainer"
        className="inbox-rightbar">
        {/* Upload File Button - Only shown when in a folder/library */}
       

{/* sourish 3/10/25 */}
         <div className="newalignbutton">
         {
        breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== "view" && (
            <div>
               <button type="button"
            className="mybutton2 mt-0 me-1"
            id="CreateFolder"
            onClick={() => setActiveComponent(true)}
          >
            + Create Folder
          </button>
          <button type="button"   className="mybutton2 mt-0"
            onClick={() => setShowUploadPanel(!showUploadPanel)}
           >
            Upload File
          </button>
 
            </div>
          
        )}
        
    <button type="button"
      onClick={() => setActiveLayout('grid')}
      style={{
        background: activeLayout === 'grid' ? '#0078d4' : '#ffffff',
        color: activeLayout === 'grid' ? 'white' : '#6c757d',
        border: '0px solid #ddd',
        borderRadius: 4,
        padding: '6px 15px',
        fontWeight: 500,
        fontSize: "14px",
      }}
    >
      Grid View
    </button>
    <button type="button"
      onClick={() => setActiveLayout('list')}
      style={{
        background: activeLayout === 'list' ? '#0078d4' : '#ffffff',
        color: activeLayout === 'list' ? 'white' : '#333',
        border: '0px solid #ddd',
        borderRadius: 4,
        padding: '6px 15px',
        fontWeight: 500,
        fontSize: "14px",
      }}
    >
      List View
    </button>
    
  </div>

        <h2
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
                    {item.title}
                  </span>
                </React.Fragment>
              ))}
            </div>
          ) : (
            <div style={{ color: "#666", fontSize: "14px" }}>Select a folder to view files</div>
          )}
        </h2>

        {/* Upload Panel */}
        {showUploadPanel && (
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
                      ></div>
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
        )}

        {/* File List */}
   {!showUploadPanel && (
  activeView === "browse" ? (
    // 📂 BROWSE VIEW (no pagination, list style)
    selectedFiles.length > 0 ? (
      <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
        {selectedFiles.map((file) => (
          <li
            key={file.Name}
            style={{
              padding: "8px 0",
              borderBottom: "1px solid #eee",
              display: "flex",
              alignItems: "center",
            }}
          >
            <span style={{ marginRight: "8px" }}>📄</span>
            <span style={{ fontSize: "14px" }}>{file.Name}</span>
            <span
              style={{
                fontSize: "14px",
                marginLeft: "10px",
                color: "#666",
              }}
            >
              {currentFolderPath}
            </span>
          </li>
        ))}
      </ul>
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
        {breadcrumbs.length > 0
          ? "No files in this folder"
          : "Please select a folder from the hierarchy to view files"}
      </div>
    )
     ) : null
)}
 
 
        {/* --- CONDITIONAL RENDERING FOR PREVIEW --- */}
  {showPreviewModal ? (
    <PreviewModal
      show={true}
        fileUrl={previewFile}
        onClose={() => setShowPreviewModal(false)}
    />
    
  ) :

  (

    // 📄 ALL OTHER VIEWS (grid + pagination) 
    <>
 
      {/* this is my documnet library files ternary oprator */}
      {
      paginatedFiles.length > 0 ? 
      (
                        <>
                          


  {/* abhay change for seach in all tabs */}

 <input
  type="text"
  placeholder="Search..."
  value={searchTerm}
  onChange={(e) => setSearchTerm(e.target.value)}
  style={{
    padding: "8px 12px",
    marginBottom: "10px",
    width: "100%",
    borderRadius: "6px",
    border: "1px solid #ccc",
  }}
/>






        {activeLayout === "grid" && (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
            gap: "18px",
            marginBottom: "20px",
          }}
        >
          {filesLoadedfromnode  && paginatedFiles.length > 0 ? (
            paginatedFiles.map((file, idx) => (
              <div
              key={file.Id || file.FileUID || idx}
             className="carddesign"
            >
              {/* File Content */}
              <div style={{ fontWeight: 600, fontSize: "15px" }}>
                {`${file.Name}/{name}`  || "Unnamed File"}
              </div>
              <div  className="text-muted font-12">
                {file.Length
                  ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB`
                  : ""}
              </div>
              <div  className="text-muted font-12">
                {file.TimeCreated
                  ? new Date(file.TimeCreated).toLocaleDateString()
                  : ""}
              </div>

              {/* Three-dot menu button */}
              <div className="dotbutton">
                <button type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    setMenuOpenIdx(menuOpenIdx === idx ? null : idx);
                  }}
                  style={{
                    background: "none",
                    border: "none",
                    fontSize: "20px",
                    cursor: "pointer",
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
                          <span>👁️</span> Preview File
                        </button>
                      </li>
                       <li>
                            <button type="button"
                             className="newbuttontext"
                              onClick={() => {
                                handleAuditHistory(file); setMenuOpenIdx(null);
                              }}
                            >
                              <span>📝</span> Audit History
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
                              <span>🕰️</span> Version History
                            </button>
                          </li>
                          <li>
                            <button type="button"
                             className="newbuttontext"
                            onClick={() => {
                                           
 
                                            setDirectDownloadFile(file); // new state for direct downloader
                                            setMenuOpenIdx(null);
                                          }}
                            >
                              <span>🕰️</span> Download File
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
                              <span>🕰️</span> Delete File
                            </button>
                          </li>
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
                                      <span>⭐</span>{" "}
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
            ))
          ) :
          // this is my request and my favourite my recucle bin files ternary operator
           (
               paginatedFiles.map((file, idx) => (
  <div
    key={file.Id || file.FileUID || idx}
    className="carddesign"
   
  >
    {/* Card Content */}
    {
         activeView === "My request" ? (   
      <>
        <div style={{ fontWeight: 600, fontSize: "15px" }}>{file.FileName}</div>
        <div  className="text-muted font-12">{file.FileSize}</div>
        <div  className="text-muted font-12">{file.DocumentLibraryName}</div>
        <div  className="text-muted font-12">{file.Status}</div>    
      </>
      ) : activeView === "My favourite" ? (
        <>
          <div style={{ fontWeight: 600, fontSize: "15px" }}>{file.FileName}</div>
          <div  className="text-muted font-12">{file.FileSize}</div>
        </>
      ) : activeView === "My Folders" ? (
        <>
          <div style={{ fontWeight: 600, fontSize: "15px" }}>{file.FolderName}</div>
          <div  className="text-muted font-12">{file.SiteTitle}</div>
        </>
      ) : activeView === "Share with me" ? (
        <>
          <div style={{ fontWeight: 600, fontSize: "15px" }}>{file.FileName}</div>
          <div  className="text-muted font-12">{file.FileSize}</div>
        </>
      ) : activeView === "Share with other" ? (
        <>
          <div style={{ fontWeight: 600, fontSize: "15px" }}>{file.FileName}</div>
          <div  className="text-muted font-12">{file.FileSize}</div>
        </>
      ) : activeView === "Recycle bin" ? (
        <>
          <div style={{ fontWeight: 600, fontSize: "15px" }}>{file.FileName}</div>
          <div  className="text-muted font-12">{file.FileSize}</div>
        </>
      ) : null
    }

    {/* Three-dot Menu */}
    <div className="dotbutton">
      <button type="button"
        onClick={(e) => {
          e.stopPropagation();
          setMenuOpenIdx(menuOpenIdx === idx ? null : idx);
        }}
        style={{
          background: "none",
          border: "none",
          fontSize: "20px",
          cursor: "pointer",
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
            
            {/* Menu for My Request */}
            {activeView === "My request" && (
              <>
                <li><button type="button"  onClick={() => {
                  // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
              setPreviewFile(file);
              setShowPreviewModal(true);
            }}>👁️ Preview File</button></li>
                <li><button  onClick={() => {
                                  // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                                handleAuditHistory(file); setMenuOpenIdx(null);
                              }}>📝 Audit History</button></li>
                <li><button type="button">↗️ Share</button></li>
                <li><button type="button" onClick={() => {
                                           
   // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                                            setDirectDownloadFile(file); // new state for direct downloader
                                            setMenuOpenIdx(null);
                                          }}>⬇️ Download</button></li>
                <li><button type="button" onClick={() => {
                    // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                                 setModalFile(file);
                                        setShowVersionModal(true);
                                        setMenuOpenIdx(null);
                              }}>🕰️ Version History</button></li>
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
            }}>👁️ Preview File</button></li>
                <li><button type="button">↗️ Share</button></li>
                <li><button type="button" onClick={() => {
                    // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                                          // handleUnmarkFavourite(file);
                                          // eslint-disable-next-line @typescript-eslint/no-floating-promises
                                          toggleFavourite(file);
 
                                          setMenuOpenIdx(null);
                                        }}>⭐ Unmark as Favourite</button></li>
              </>
            )}

            {/* Menu for My Folders */}
            {activeView === "My Folders" && (
              <>
                <li><button type="button" onClick={() => {
                    // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                                                setModalFile(file);
                                                setMenuOpenIdx(null);
                                                deleteFolder(file);  // sourish 20/8/25
                                              }}
                        >🗑️ Delete Folder</button></li>
                <li><button type="button" onClick={() => {
                    // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                                                setModalFile(file);
                                                setRenameValue(file?.FolderName || ""); // prefill with old name
                                                setRenameModalOpen(true);
                                                setMenuOpenIdx(null);
                                              }}>✏️ Rename Folder</button></li>
                                                 <li>
                                    <button type="button" onClick={() => {
                                      const pageBefore = currentPage;
                                      setMenuOpenIdx(null);
                                      window.managePermission?.(file);
                                      setCurrentPage(pageBefore);
                                    }}>
                                      🔒 Manage Permission
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
            }}>👁️ Preview File</button></li>
                <li><button type="button" onClick={() => {
                                             // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
 
                                            setDirectDownloadFile(file); // new state for direct downloader
                                            setMenuOpenIdx(null);
                                          }}>⬇️ Download File</button></li>
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
            }}>👁️ Preview File</button></li>
                <li><Button
  variant="warning"
  onClick={() => {
    setShowRevokeModal(true);
    setAcessFile(file); // file = your folder/file object
  }}
>
  Revoke Access
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
            }}>👁️ Preview File</button></li>
                <li><button type="button" onClick={async () => {
                    // this is to hide li options in every tab (addhyan)
                  setMenuOpenIdx(null);
                  // this is to hide li options in every tab (addhyan)
                  await handleUndoDelete(file); setMenuOpenIdx(null); }}
                  >↩️ Undo (Restore)</button></li>
              </>
            )}
          </ul>
        </div>
      )}
    </div>
  </div>
)

          )
          
        )}
    
        </div>
        )}

 {activeLayout === "list" && (
        <div>
           <table className="mtablenew">
         <thead>
  <tr >
    <th style={{minWidth:'50px',maxWidth:'50px',}}>S.No</th>
    <th style={{minWidth:'250px',maxWidth:'250px',}}>Name</th>
    {/* <th >Size</th> */}
    <th>Size</th>
    {/* <th>Library</th> */}
    <th>
            {paginatedFiles.some((f) => f.DocumentLibraryName)
              ? "Library"
              : "Created Date"}
          </th>
      {paginatedFiles.some((f) => f.Status && f.Status.trim() !== "") && (
            <th style={{ minWidth:'80px',maxWidth:'80px',}}>Status</th>
          )}
    {/* <th style={{minWidth:'80px',maxWidth:'80px',}}>Status</th> */}
    <th style={{textAlign:"center"}}>Action</th> {/* New column */}
  </tr>
</thead>
 
        <tbody>
  {paginatedFiles.map((file, idx) => (
    <tr key={file.Id || idx}>
      <td style={{minWidth:'50px',maxWidth:'50px',}}>{file.SNo || idx + 1}</td>
      {/* <td style={{minWidth:'250px',maxWidth:'250px',}}>{file.FileName}</td> */}
      <td style={{minWidth:'250px',maxWidth:'250px',}}>{file.FileName || file.Name || file.FolderName}</td>
      {/* <td >{file.FileSize}</td> */}
        <td>
  {file.FileSize
    ? file.FileSize
    : file.Length
    ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB`
    : "-"}
</td>
      {/* <td >{file.DocumentLibraryName}</td> */}
         <td>
  {file.DocumentLibraryName
    ? file.DocumentLibraryName
    : file.TimeCreated
    ? new Date(file.TimeCreated).toLocaleDateString()
    : "-"}
</td>
      {/* <td style={{minWidth:'80px',maxWidth:'80px',}}>{file.Status}</td> */}
      {paginatedFiles.some((f) => f.Status && f.Status.trim() !== "") && (
      <td style={{minWidth:'80px',maxWidth:'80px',}}>{file.Status || ""}</td>
       )}
      <td style={{textAlign:"center"}}>
        <button type="button" className="dotbutton2"
          onClick={(e) => {
            e.stopPropagation();
            setMenuOpenIdx(menuOpenIdx === idx ? null : idx);
          }}
         
        >
          ⋮
        </button>
 
        {menuOpenIdx === idx && (
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
          >
            <ul  className="internalbutton">
              <li>
                <button type="button" onClick={() => { setPreviewFile(file); setShowPreviewModal(true); setMenuOpenIdx(null); }}>
                  👁️ Preview
                </button>
              </li>
              <li>
                <button type="button" onClick={() => { setDirectDownloadFile(file); setMenuOpenIdx(null); }}>
                  ⬇️ Download
                </button>
              </li>
              <li>
                <button type="button" onClick={async () => { await deleteFileFolder(file, currentSiteUrl, context); setMenuOpenIdx(null); }}>
                  🗑️ Delete
                </button>
              </li>
            </ul>
          </div>
        )}
      </td>
    </tr>
  ))}
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
          No files found in 
        </div>
      )}
      {/* create Folder modal */}
        <Modal show={activeComponent} onHide={() => setActiveComponent(false)} className='filemodal'>
                <Modal.Header closeButton>
                  <Modal.Title > <h4 className='font-16 text-dark fw-bold mb-1'>This Folder will create under:  <h2
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
                                //   index === breadcrumbs.length - 1 ? "#333" : "#0066cc",
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
                      position: "fixed",
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
                            background: "#f5f5f5",
                            cursor: "pointer",
                          }}
                          onClick={() => setRenameModalOpen(false)}
                        >
                          Cancel
                        </button>
                        <button
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
                        </button>
                      </div>
                    </div>
                  </div>
                )}
                {/* download component */}
                <DirectDownloader
                file={directDownloadFile}
                context={context}
                trigger={true}
              />
 

      {/* preview file modal */}
 <PreviewModal
    show={showPreviewModal}
    fileUrl={previewFile}
    onClose={() => setShowPreviewModal(false)}
  />

{/*revoke modal*/ }
 <Revoke
  show={showRevokeModal}
  selectedFolder={revoke}
  context={context}
  onClose={() => setShowRevokeModal(false)}
  onRevoke={(userId: string) => {
    console.log("Revoked user ID:", userId);
    alert(`Access revoked for user ID: ${userId}`);
  }}
/>

     {/* version history  */}
   <VersionHistoryModal
                show={showVersionModal}
                file={modalFile}
                context={context}
                onClose={() => setShowVersionModal(false)}
              />

              {/* aman code manage folder permission */}
    {/* === Step 1 Popup === */}
{showPermissionModal && (
  <Modal show={showPermissionModal} onHide={() => setShowPermissionModal(false)}>
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
)}

{/* === Step 2 Popup (with People Picker) === */}
{showManagePermissionModal && (
  <Modal show={showManagePermissionModal} onHide={() => setShowManagePermissionModal(false)} size="lg">
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
                        setMpUserSuggestions((mpSiteUsers || []).slice(0,50));
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


      {/* Pagination controls */}
                  {
                    
                    // selectedFiles
                    
                    
                       filteredFiles
                    
                    .length > 0 && (
        <div
          style={{
            margin: "10px 0",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <button
            onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
            disabled={currentPage === 1}
            style={{
              marginRight: "10px",
              padding: "5px 12px",
              borderRadius: "4px",
              border: "1px solid #ccc",
              background: currentPage === 1 ? "#eee" : "#fff",
              cursor: currentPage === 1 ? "not-allowed" : "pointer",
            }}
          >
            Prev
          </button>
          <span style={{ margin: "0 10px" }}>
                          {/* Page {currentPage} of {Math.ceil(selectedFiles.length / pageSize)} */}
                            Page {currentPage} of {Math.ceil(filteredFiles.length / pageSize) || 1}

          </span>
          <button
           onClick={(e) => {
            e.preventDefault(); // Add this line
            setCurrentPage((p) =>
              p < Math.ceil(selectedFiles.length / pageSize) ? p + 1 : p
            );
          }}
            // onClick={() =>
            //   setCurrentPage((p) =>
            //     p < Math.ceil(selectedFiles.length / pageSize) ? p + 1 : p
            //   )
            // }
            disabled={currentPage === Math.ceil(selectedFiles.length / pageSize)}
            style={{
              marginLeft: "10px",
              padding: "5px 12px",
              borderRadius: "4px",
              border: "1px solid #ccc",
              background:
                currentPage === Math.ceil(selectedFiles.length / pageSize)
                  ? "#eee"
                  : "#fff",
              cursor:
                currentPage === Math.ceil(selectedFiles.length / pageSize)
                  ? "not-allowed"
                  : "pointer",
            }}
          >
            Next
          </button>
        </div>
      )}
    </>
  )
}
     


      </div></div>
       {/* ---------------------------Audit History Modal ---------------------*/}
    {showAuditModal && (
      <Modal show={showAuditModal} onHide={() => setShowAuditModal(false)} size="lg">
        <Modal.Header closeButton>
          <Modal.Title>Audit History</Modal.Title>
        </Modal.Header>
        <Modal.Body>
          {auditLoading ? (
            <div>Loading...</div>
          ) : (
            <div>
              <h5>Metadata</h5>
              <ul>
                {Object.entries(auditVersions.Metadata || {}).map(([key, value]) => (
                  <li key={key}><strong>{key}:</strong> {String(value)}</li>
                ))}
              </ul>

              <h5>Versions</h5>
              {auditVersions.Versions && auditVersions.Versions.length ? (
                <ul>
                  {auditVersions.Versions.map((v: any, i: number) => (
                    <li key={i}>{JSON.stringify(v)}</li>
                  ))}
                </ul>
              ) : (
                <div>No versions available</div>
              )}
            </div>
          )}
        </Modal.Body>
      </Modal>
    )}
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