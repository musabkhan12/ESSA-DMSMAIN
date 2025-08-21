
import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IEssadmsMainProps } from "./IEssadmsMainProps";
import PreviewModal from "./previewfile"; // import your modal
import VersionHistoryModal from "./versionhistory"; // import version history modal
import Swal from 'sweetalert2'
import { Modal } from 'react-bootstrap';
import CreateFolder from "./CreateFolder";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
import { useMemo } from "react";
// let loadfilefromnode = ''
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

  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 12;
  let location: string = "";
  const paginatedFiles = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    const end = start + pageSize;
    return selectedFiles.slice(start, end);
  }, [selectedFiles, currentPage]);

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
        let payload: any = {
          IsDeleted: null
        }
        const itemData = await fileItem.update(payload)
        console.log("column updated successfully", itemData);
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
  //--------------undodelete function ends---------------


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
        // sourish 20/8/25
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
    // sourish 20/8/25
    if (viewType === "MyFolders") {
      try {
        const meEmail =
          (context.pageContext as any)?.user?.email ||
          (context.pageContext as any)?.user?.loginName ||
          "";

        const masterSites = await sp.web.lists
          .getByTitle("MasterSiteCollection")
          .items.select("Id", "Title", "SiteURL", "SharewithOtherMeMasterSite")
          .top(5000)();

        const allMyFoldersData: any[] = [];

        for (const ms of masterSites) {
          try {
            // Always target "DMSFolderMaster" list from SiteURL
            const siteSP = spfi(ms.SiteURL).using(SPFx(context));
            const listItems = await siteSP.web.lists
              .getByTitle("DMSFolderMaster")
              .items.select(
                "Id",
                "Title",
                "FolderPath",
                "SiteTitle",
                "DocumentLibraryName",
                "FolderName",
                "CurrentUser",
                "Modified",
                "Author/Id",
                "Author/Title",
                "Author/EMail"
              )
              .expand("Author") // 🔹 expand ModifiedBy lookup
              .top(5000)();

            const normalizedItems = listItems.map((it: any) => ({
              ...it,
              ParentFolder: it.ParentFolder || it.ParentFolderId || "", // normalize
            }));

            // 🔹 Filter by CurrentUser column
            const filtered = (listItems || []).filter(
              (it: any) =>
                (it.Author.EMail || "").toLowerCase() === meEmail.toLowerCase()
            );

            allMyFoldersData.push(
              ...filtered.map((it: any) => ({
                ...it,
                __source: "MyFolders",
                __siteUrl: ms.SiteURL,
                __listName: "DMSFolderMaster",
                __CreatedBy: it.Author?.Title || it.Author?.EMail || "",
              }))
            );
          } catch (err) {
            console.error(
              `Error fetching list DMSFolderMaster from ${ms.SiteURL}:`,
              err
            );
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
    <ul style={{ listStyleType: "none", paddingLeft: "20px" }}>
      {nodes.map((node) => (
        <li key={node.key}>
          <div style={{ display: "flex", alignItems: "center" }}>
            {node.hasChildren && (
              <button
                onClick={() => toggleNode(node)}
                style={{
                  marginRight: "5px",
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                  fontSize: "16px",
                  width: "20px",
                  height: "20px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                }}
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


  return (
    <div
      id="maincontainer"
      style={{
        display: "flex",
        height: "calc(100vh - 50px)",
        marginTop: "10px",
      }}
    >
      {/* Left Panel with Quick Views and Folder Hierarchy */}
      <div style={{ width: "30%", display: "flex", flexDirection: "column" }}>
        {/* Quick Views Panel */}
        <div
          id="buttonpanel"
          style={{
            padding: "15px",
            backgroundColor: "#f0f0f0",
            borderBottom: "1px solid #ddd",
            flexShrink: 0,
          }}
        >
          <h2
            style={{
              fontSize: "16px",
              fontWeight: "600",
              marginBottom: "15px",
              color: "#333",
            }}
          >
            Quick Views
          </h2>
          {[
            "My request",
            "My favourite",
            "My Folders",  // sourish 20/8/25
            "Share with me",
            "Share with other",
            "Recycle bin",
          ].map((view) => (
            <button
              key={view}
              onClick={() => handleViewButtonClick(view)}
              style={{
                display: "block",
                width: "100%",
                padding: "8px 10px",
                marginBottom: "8px",
                textAlign: "left",
                backgroundColor: activeView === view ? "#0078d4" : "#f0f0f0",
                color: activeView === view ? "white" : "#333",
                border: "1px solid #ddd",
                borderRadius: "4px",
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
            backgroundColor: "#f9f9f9",
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
        style={{
          width: "70%",
          padding: "15px",
          overflow: "auto",
          backgroundColor: "#fff",
          position: "relative",
        }}
      >
        {/* Upload File Button - Only shown when in a folder/library */}
        {
          breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== "view" && (
            <div>
              <button
                className="mybutton2 mt-0"
                id="CreateFolder"
                onClick={() => setActiveComponent(true)}
              >
                + Create Folder
              </button>
              <button
                onClick={() => setShowUploadPanel(!showUploadPanel)}
                style={{
                  position: "absolute",
                  top: "15px",
                  right: "15px",
                  padding: "8px 15px",
                  backgroundColor: "#0078d4",
                  color: "white",
                  border: "none",
                  borderRadius: "4px",
                  cursor: "pointer",
                  fontSize: "14px",
                }}
              >
                Upload File
              </button>
            </div>

          )}

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
                    className="breadcrumb-item"
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
            <button
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
                <button
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
          ) :

            (

              // 📄 ALL OTHER VIEWS (grid + pagination) 
              <>

                {/* this is my documnet library files ternary oprator */}
                {
                  paginatedFiles.length > 0 ? (
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
                        gap: "18px",
                        marginBottom: "20px",
                      }}
                    >
                      {filesLoadedfromnode && paginatedFiles.length > 0 ? (
                        paginatedFiles.map((file, idx) => (
                          <div
                            key={file.Id || file.FileUID || idx}
                            style={{
                              background: "#f8f8f8",
                              border: "1px solid #e0e0e0",
                              borderRadius: "8px",
                              padding: "16px",
                              minHeight: "120px",
                              display: "flex",
                              flexDirection: "column",
                              justifyContent: "space-between",
                              position: "relative",
                            }}
                          >
                            {/* File Content */}
                            <div style={{ fontWeight: 600, fontSize: "15px" }}>
                              {`${file.Name}/{name}` || "Unnamed File"}
                            </div>
                            <div style={{ fontSize: "12px", color: "#666" }}>
                              {file.Length
                                ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB`
                                : ""}
                            </div>
                            <div style={{ fontSize: "12px", color: "#666" }}>
                              {file.TimeCreated
                                ? new Date(file.TimeCreated).toLocaleDateString()
                                : ""}
                            </div>

                            {/* Three-dot menu button */}
                            <div style={{ position: "absolute", top: 10, right: 10 }}>
                              <button
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
                                  <ul style={{ margin: 0, padding: 0, listStyle: "none" }}>
                                    {/* Common actions for all views */}
                                    <li>
                                      <button
                                        style={{
                                          width: "100%",
                                          padding: "8px 12px",
                                          background: "none",
                                          border: "none",
                                          textAlign: "left",
                                          cursor: "pointer",
                                          fontSize: "14px",
                                          display: "flex",
                                          alignItems: "center",
                                          gap: "8px",
                                        }}
                                        onClick={() => {
                                          setPreviewFile(file);
                                          setShowPreviewModal(true);
                                        }}
                                      >
                                        <span>👁️</span> Preview File
                                      </button>
                                    </li>

                                    {/* My Request specific actions */}
                                    {activeView === "My request" && (
                                      <>
                                        <li>
                                          <button
                                            style={{
                                              width: "100%",
                                              padding: "8px 12px",
                                              background: "none",
                                              border: "none",
                                              textAlign: "left",
                                              cursor: "pointer",
                                              fontSize: "14px",
                                              display: "flex",
                                              alignItems: "center",
                                              gap: "8px",
                                            }}
                                            onClick={() => {
                                              setModalFile(file);
                                              setShowAuditModal(true);
                                              setMenuOpenIdx(null);
                                            }}
                                          >
                                            <span>📝</span> Audit History
                                          </button>
                                        </li>
                                        <li>
                                          <button
                                            style={{
                                              width: "100%",
                                              padding: "8px 12px",
                                              background: "none",
                                              border: "none",
                                              textAlign: "left",
                                              cursor: "pointer",
                                              fontSize: "14px",
                                              display: "flex",
                                              alignItems: "center",
                                              gap: "8px",
                                            }}
                                            onClick={() => {
                                              setModalFile(file);
                                              setShowShareModal(true);
                                              setMenuOpenIdx(null);
                                            }}
                                          >
                                            <span>↗️</span> Share
                                          </button>
                                        </li>
                                        <li>
                                          <button
                                            style={{
                                              width: "100%",
                                              padding: "8px 12px",
                                              background: "none",
                                              border: "none",
                                              textAlign: "left",
                                              cursor: "pointer",
                                              fontSize: "14px",
                                              display: "flex",
                                              alignItems: "center",
                                              gap: "8px",
                                            }}
                                            onClick={() => {
                                              setMenuOpenIdx(null);
                                              // handleDownloadFile(file);
                                            }}
                                          >
                                            <span>⬇️</span> Download
                                          </button>
                                        </li>
                                        <li>
                                          <button
                                            style={{
                                              width: "100%",
                                              padding: "8px 12px",
                                              background: "none",
                                              border: "none",
                                              textAlign: "left",
                                              cursor: "pointer",
                                              fontSize: "14px",
                                              display: "flex",
                                              alignItems: "center",
                                              gap: "8px",
                                            }}
                                            onClick={() => {
                                              setModalFile(file);
                                              setShowVersionModal(true);
                                              setMenuOpenIdx(null);
                                            }}
                                          >
                                            <span>🕰️</span> Version History
                                          </button>
                                        </li>
                                      </>
                                    )}

                                    {/* My Favourite specific action */}
                                    {activeView === "My favourite" && (
                                      <li>
                                        <button
                                          style={{
                                            width: "100%",
                                            padding: "8px 12px",
                                            background: "none",
                                            border: "none",
                                            textAlign: "left",
                                            cursor: "pointer",
                                            fontSize: "14px",
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "8px",
                                          }}
                                          onClick={() => {
                                            // handleUnmarkFavourite(file);

                                            setMenuOpenIdx(null);
                                          }}
                                        >
                                          <span>⭐</span> Unmark as Favourite
                                        </button>
                                      </li>

                                    )}

                                    {/* Recycle Bin specific action */}
                                    {activeView === "Recycle bin" && (
                                      <li>
                                        <button
                                          style={{
                                            width: "100%",
                                            padding: "8px 12px",
                                            background: "none",
                                            border: "none",
                                            textAlign: "left",
                                            cursor: "pointer",
                                            fontSize: "14px",
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "8px",
                                          }}
                                          onClick={async () => { await handleUndoDelete(file); setMenuOpenIdx(null); }}
                                        >
                                          <span>↩️</span> Undo (Restore)
                                        </button>
                                      </li>
                                    )}
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
                              style={{
                                background: "#f8f8f8",
                                border: "1px solid #e0e0e0",
                                borderRadius: "8px",
                                padding: "16px",
                                minHeight: "120px",
                                display: "flex",
                                flexDirection: "column",
                                justifyContent: "space-between",
                                position: "relative",
                              }}
                            >
                              {/* File Content */}
                              {/* <div style={{ fontWeight: 600, fontSize: "15px" }}>
                                {`${file.FileName}` || "Unnamed File"}
                              </div>
                              <div style={{ fontSize: "12px", color: "#666" }}>
                                {file.FileSize
                                  ? `${(parseInt(file.FileSize) / (1024 * 1024)).toFixed(2)} MB`
                                  : ""}
                              </div>
                              <div style={{ fontSize: "12px", color: "#666" }}>
                                {file.TimeCreated
                                  ? new Date(file.TimeCreated).toLocaleDateString()
                                  : ""}
                              </div> */}
                              {/* // Sourish 22/8/25 */}
                              {activeView === "Share with me" || activeView === "Share with other" ? (
                                <>
                                  <div style={{ fontWeight: 600, fontSize: "15px" }}>
                                    {file.FileName ? file.FileName : "Unnamed File"}
                                  </div>
                                  <div style={{ fontSize: "12px", color: "#666" }}>
                                    {file.FileSize
                                      ? `${(parseInt(file.FileSize) / (1024 * 1024)).toFixed(2)} MB`
                                      : ""}
                                  </div>
                                  <div style={{ fontSize: "12px", color: "#666" }}>
                                    {file.TimeCreated
                                      ? new Date(file.TimeCreated).toLocaleDateString()
                                      : ""}
                                  </div>
                                </>
                              ) : activeView === "My Folders" ? (
                                <>
                                  <div style={{ fontWeight: 600, fontSize: "15px" }}>
                                    {file.FolderName ? file.FolderName : "Unnamed File"}
                                  </div>
                                  <div style={{ fontSize: "12px", color: "#666" }}>
                                    {file.FileSize
                                      ? `${(parseInt(file.FileSize) / (1024 * 1024)).toFixed(2)} MB`
                                      : file.Author.Title || ""}
                                  </div>
                                  <div style={{ fontSize: "12px", color: "#666" }}>
                                    {file.TimeCreated
                                      ? new Date(file.TimeCreated).toLocaleDateString()
                                      : ""}
                                  </div>
                                </>
                              ) : (
                                <>
                                  {/* File Content */}
                                  <div style={{ fontWeight: 600, fontSize: "15px" }}>
                                    {`${file.FileName}` || "Unnamed File"}
                                  </div>
                                  <div style={{ fontSize: "12px", color: "#666" }}>
                                    {file.FileSize
                                      ? `${(parseInt(file.FileSize) / (1024 * 1024)).toFixed(2)} MB`
                                      : ""}
                                  </div>
                                  <div style={{ fontSize: "12px", color: "#666" }}>
                                    {file.TimeCreated
                                      ? new Date(file.TimeCreated).toLocaleDateString()
                                      : ""}
                                  </div>
                                </>
                              )}



                              {/* Three-dot menu button */}
                              <div style={{ position: "absolute", top: 10, right: 10 }}>
                                <button
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
                                    <ul style={{ margin: 0, padding: 0, listStyle: "none" }}>
                                      {/* Common actions for all views */}
                                      <li>
                                        <button
                                          style={{
                                            width: "100%",
                                            padding: "8px 12px",
                                            background: "none",
                                            border: "none",
                                            textAlign: "left",
                                            cursor: "pointer",
                                            fontSize: "14px",
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "8px",
                                          }}
                                          onClick={() => {
                                            setPreviewFile(file);
                                            setShowPreviewModal(true);
                                          }}
                                        >
                                          <span>👁️</span> Preview File
                                        </button>
                                      </li>

                                      {/* My Request specific actions */}
                                      {activeView === "My request" && (
                                        <>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setModalFile(file);
                                                alert("Audit History is not implemented yet");
                                                setShowAuditModal(true);
                                                setMenuOpenIdx(null);
                                              }}
                                            >
                                              <span>📝</span> Audit History
                                            </button>
                                          </li>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setModalFile(file);
                                                setShowShareModal(true);
                                                setMenuOpenIdx(null);
                                              }}
                                            >
                                              <span>↗️</span> Share
                                            </button>
                                          </li>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setMenuOpenIdx(null);
                                                alert("downloaad is not implemented yet");
                                                // handleDownloadFile(file);
                                              }}
                                            >
                                              <span>⬇️</span> Download
                                            </button>
                                          </li>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setModalFile(file);

                                                setShowVersionModal(true);
                                                setMenuOpenIdx(null);
                                              }}
                                            >
                                              <span>🕰️</span> Version History
                                            </button>
                                          </li>
                                        </>
                                      )}

                                      {/* My Favourite specific action */}
                                      {activeView === "My favourite" && (
                                        <li>
                                          <button
                                            style={{
                                              width: "100%",
                                              padding: "8px 12px",
                                              background: "none",
                                              border: "none",
                                              textAlign: "left",
                                              cursor: "pointer",
                                              fontSize: "14px",
                                              display: "flex",
                                              alignItems: "center",
                                              gap: "8px",
                                            }}
                                            onClick={() => {
                                              // handleUnmarkFavourite(file);

                                              setMenuOpenIdx(null);
                                            }}
                                          >
                                            <span>⭐</span> Unmark as Favourite
                                          </button>
                                        </li>

                                      )}

                                      {/* Recycle Bin specific action */}
                                      {activeView === "Recycle bin" && (
                                        <li>
                                          <button
                                            style={{
                                              width: "100%",
                                              padding: "8px 12px",
                                              background: "none",
                                              border: "none",
                                              textAlign: "left",
                                              cursor: "pointer",
                                              fontSize: "14px",
                                              display: "flex",
                                              alignItems: "center",
                                              gap: "8px",
                                            }}
                                            onClick={async () => { await handleUndoDelete(file); setMenuOpenIdx(null); }}
                                          >
                                            <span>↩️</span> Undo (Restore)
                                          </button>
                                        </li>
                                      )}

                                      {/* sourish 20/8/25 */}
                                      {/* My Folders specific action */}
                                      {activeView === "My Folders" && (
                                        <>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setModalFile(file);
                                                setMenuOpenIdx(null);
                                                deleteFolder(file);  // sourish 20/8/25
                                              }}
                                            >
                                              <span>📝</span> Delete Folder
                                            </button>
                                          </li>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              // onClick={() => {
                                              //   setModalFile(file);
                                              //   setMenuOpenIdx(null);
                                              //   // RenameFolder(file); //sourish 21/8/25
                                              // }}
                                              onClick={() => {
                                                setModalFile(file);
                                                setRenameValue(file?.FolderName || ""); // prefill with old name
                                                setRenameModalOpen(true);
                                                setMenuOpenIdx(null);
                                              }}

                                            >
                                              <span>↗️</span> Rename Folder
                                            </button>
                                          </li>
                                        </>
                                      )}

                                      {/* sourish 21/8/25 */}
                                      {activeView === "Share with other" && (
                                        <>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setModalFile(file);
                                                setMenuOpenIdx(null);
                                              }}
                                            >
                                              <span>📝</span> Revoke Access
                                            </button>
                                          </li>
                                        </>
                                      )}
                                      {/* sourish 21/8/25 */}
                                      {activeView === "Share with me" && (
                                        <>
                                          <li>
                                            <button
                                              style={{
                                                width: "100%",
                                                padding: "8px 12px",
                                                background: "none",
                                                border: "none",
                                                textAlign: "left",
                                                cursor: "pointer",
                                                fontSize: "14px",
                                                display: "flex",
                                                alignItems: "center",
                                                gap: "8px",
                                              }}
                                              onClick={() => {
                                                setModalFile(file);
                                                setMenuOpenIdx(null);
                                              }}
                                            >
                                              <span>📝</span> Download File
                                            </button>
                                          </li>
                                        </>
                                      )}
                                    </ul>
                                  </div>
                                )}
                              </div>
                            </div>
                          ))
                        )}


                      {/* previous code */}
                      {/* {paginatedFiles.map((file, idx) => (
            <div
              key={file.Id || file.FileUID || idx}
              style={{
                background: "#f8f8f8",
                border: "1px solid #e0e0e0",
                borderRadius: "8px",
                padding: "16px",
                minHeight: "120px",
                display: "flex",
                flexDirection: "column",
                justifyContent: "space-between",
                position: "relative",
              }}
            >
           
              <div style={{ fontWeight: 600, fontSize: "15px" }}>
                {`${file.Name}/{name}`  || "Unnamed File"}
              </div>
              <div style={{ fontSize: "12px", color: "#666" }}>
                {file.Length
                  ? `${(parseInt(file.Length) / (1024 * 1024)).toFixed(2)} MB`
                  : ""}
              </div>
              <div style={{ fontSize: "12px", color: "#666" }}>
                {file.TimeCreated
                  ? new Date(file.TimeCreated).toLocaleDateString()
                  : ""}
              </div>

              <div style={{ position: "absolute", top: 10, right: 10 }}>
                <button
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
                    <ul style={{ margin: 0, padding: 0, listStyle: "none" }}>
                    
                      <li>
                        <button
                          style={{
                            width: "100%",
                            padding: "8px 12px",
                            background: "none",
                            border: "none",
                            textAlign: "left",
                            cursor: "pointer",
                            fontSize: "14px",
                            display: "flex",
                            alignItems: "center",
                            gap: "8px",
                          }}
                        onClick={() => {
              setPreviewFile(file);
              setShowPreviewModal(true);
            }}
                        >
                          <span>👁️</span> Preview File
                        </button>
                      </li>

                      {activeView === "My request" && (
                        <>
                          <li>
                            <button
                              style={{
                                width: "100%",
                                padding: "8px 12px",
                                background: "none",
                                border: "none",
                                textAlign: "left",
                                cursor: "pointer",
                                fontSize: "14px",
                                display: "flex",
                                alignItems: "center",
                                gap: "8px",
                              }}
                              onClick={() => {
                                setModalFile(file);
                                setShowAuditModal(true);
                                setMenuOpenIdx(null);
                              }}
                            >
                              <span>📝</span> Audit History
                            </button>
                          </li>
                          <li>
                            <button
                              style={{
                                width: "100%",
                                padding: "8px 12px",
                                background: "none",
                                border: "none",
                                textAlign: "left",
                                cursor: "pointer",
                                fontSize: "14px",
                                display: "flex",
                                alignItems: "center",
                                gap: "8px",
                              }}
                              onClick={() => {
                                setModalFile(file);
                                setShowShareModal(true);
                                setMenuOpenIdx(null);
                              }}
                            >
                              <span>↗️</span> Share
                            </button>
                          </li>
                          <li>
                            <button
                              style={{
                                width: "100%",
                                padding: "8px 12px",
                                background: "none",
                                border: "none",
                                textAlign: "left",
                                cursor: "pointer",
                                fontSize: "14px",
                                display: "flex",
                                alignItems: "center",
                                gap: "8px",
                              }}
                              onClick={() => {
                                setMenuOpenIdx(null);
                                // handleDownloadFile(file);
                              }}
                            >
                              <span>⬇️</span> Download
                            </button>
                          </li>
                          <li>
                            <button
                              style={{
                                width: "100%",
                                padding: "8px 12px",
                                background: "none",
                                border: "none",
                                textAlign: "left",
                                cursor: "pointer",
                                fontSize: "14px",
                                display: "flex",
                                alignItems: "center",
                                gap: "8px",
                              }}
                              onClick={() => {
                                 setModalFile(file);
                                        setShowVersionModal(true);
                                        setMenuOpenIdx(null);
                              }}
                            >
                              <span>🕰️</span> Version History
                            </button>
                          </li>
                        </>
                      )}

                     
                      {activeView === "My favourite" && (
                        <li>
                          <button
                            style={{
                              width: "100%",
                              padding: "8px 12px",
                              background: "none",
                              border: "none",
                              textAlign: "left",
                              cursor: "pointer",
                              fontSize: "14px",
                              display: "flex",
                              alignItems: "center",
                              gap: "8px",
                            }}
                            onClick={() => {
                              // handleUnmarkFavourite(file);
                              setMenuOpenIdx(null);
                            }}
                          >
                            <span>⭐</span> Unmark as Favourite
                          </button>
                        </li>

                      )}

                      {activeView === "Recycle bin" && (
                        <li>
                          <button
                            style={{
                              width: "100%",
                              padding: "8px 12px",
                              background: "none",
                              border: "none",
                              textAlign: "left",
                              cursor: "pointer",
                              fontSize: "14px",
                              display: "flex",
                              alignItems: "center",
                              gap: "8px",
                            }}
                            onClick={async () => { await handleUndoDelete(file); setMenuOpenIdx(null); }}
                          >
                            <span>↩️</span> Undo (Restore)
                          </button>
                        </li>
                      )}
                    </ul>
                  </div>
                )}
              </div>
            </div>
          ))} */}


                    </div>
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
                      No files found in {activeView}
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

                {/* preview file modal */}
                <PreviewModal
                  show={showPreviewModal}
                  fileUrl={previewFile}
                  onClose={() => setShowPreviewModal(false)}
                />
                {/* version history  */}
                <VersionHistoryModal
                  show={showVersionModal}
                  file={modalFile}
                  context={context}
                  onClose={() => setShowVersionModal(false)}
                />
                {/* Pagination controls */}
                {selectedFiles.length > 0 && (
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
                      Page {currentPage} of {Math.ceil(selectedFiles.length / pageSize)}
                    </span>
                    <button
                      onClick={() =>
                        setCurrentPage((p) =>
                          p < Math.ceil(selectedFiles.length / pageSize) ? p + 1 : p
                        )
                      }
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

        )}

      </div>
    </div>
  );
};

const DMSMain: React.FC<IEssadmsMainProps> = (props) => {
  return <ArgPoc context={props.context} />;
};

export default DMSMain;
