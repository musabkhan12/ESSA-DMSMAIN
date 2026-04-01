

import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IEssadmsMainProps } from "./IEssadmsMainProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";

interface TreeNode {
  key: string;
  title: string;
  type: "site" | "subsite" | "library" | "folder";
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
  const [treeData, setTreeData] = useState<TreeNode[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<any[]>([]);
  const [breadcrumbs, setBreadcrumbs] = useState<BreadcrumbItem[]>([]);
  const [nodeMap, setNodeMap] = useState<Record<string, TreeNode>>({});
  const [isInitialLoad, setIsInitialLoad] = useState(true);
  const [activeView, setActiveView] = useState<string>("");
  const [showUploadPanel, setShowUploadPanel] = useState(false);
  const [selectedUploadFiles, setSelectedUploadFiles] = useState<File[]>([]);
  const [uploadProgress, setUploadProgress] = useState<number>(0);
  const [currentFolderPath, setCurrentFolderPath] = useState<string>("");
  const [currentSiteUrl, setCurrentSiteUrl] = useState<string>("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [previewFileUrls, setPreviewFileUrls] = useState<string[]>([]);
  let location: string = "";

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

// useEffect(() => {
//   const initialize = async () => {
//     console.log("[initialize] Start");
//     await loadRootSites();

//     const rawHash = window.location.hash.substring(1);
//     const hash = decodeURIComponent(rawHash || "");
//     console.log("[initialize] Hash:", rawHash, "decoded:", hash);

//     if (hash) {
//       setPendingPath(hash.split("/"));
//     } else {
//       handleViewButtonClick("My request");
//     }
//     setIsInitialLoad(false);
//     console.log("[initialize] Complete");
//   };

//   initialize();

//   const handleHashChange = () => {
//     const rawHash = window.location.hash.substring(1);
//     const hash = decodeURIComponent(rawHash || "");
//     console.log("[hashchange] New hash:", rawHash, "decoded:", hash);
//     if (hash) {
//       setPendingPath(hash.split("/"));
//     } else {
//       handleViewButtonClick("My request");
//     }
//   };

//   window.addEventListener("hashchange", handleHashChange);
//   return () => window.removeEventListener("hashchange", handleHashChange);
// }, []);

// useEffect(() => {
//   if (pendingPath && treeData.length) {
//     navigateToPath(pendingPath);
//     setPendingPath(null);
//   }
// }, [pendingPath, treeData]);

  /** ------------------------------ load root sites/subsites ---------------------------------- */
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
        sp.web.lists
          .getByTitle("MasterSiteURL")
          .items.select("Id", "Description", "Title", "SiteURL", "Active", "SiteID", "FileMasterList")
          .top(5000)(),
      ]);

      const nodes: TreeNode[] = siteList.map((site: any) => {
        const siteUrl = site.SiteURL?.trim();
        const siteSubs = subsiteList.filter((s: any) => {
          const subUrl = s.SiteURL?.trim().toLowerCase();
          const parentUrl = siteUrl?.toLowerCase();
          if (!subUrl || !parentUrl) return false;
          // exact match or direct subsite
          return subUrl === parentUrl || subUrl.startsWith(parentUrl + "/");
        });

        return {
          key: `site-${site.Id}`,
          title: site.Title,
          type: "site",
          siteUrl: siteUrl,
          hasChildren: siteSubs.length > 0,
          children: siteSubs.map((sub: any) => ({
            key: `subsite-${sub.Id}`,
            title: sub.Title,
            type: "subsite",
            siteUrl: sub.SiteURL?.trim(),
            hasChildren: true,
            parentKey: `site-${site.Id}`,
          })),
        };
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

  /** ------------------------ smarter deep-link navigation (key change) ------------------------ */
  // const navigateToPath = async (pathTitles: string[]) => {
  //   loadRootSites();
  //   console.log("[navigateToPath] raw segments:", pathTitles);
  //   console.log("[navigateToPath] treeData:", treeData  ,JSON.stringify(treeData));
  //   if (pathTitles.length === 0 || !treeData.length) {
  //     console.log("[navigateToPath] Abort: no segments or tree not ready");
  //     return;
  //   }

  //   const decodedTitles = pathTitles.map((t) => decodeURIComponent(t)).filter((title) => title.trim() !== "");
  //   console.log("[navigateToPath] decoded segments:", decodedTitles);
  //   if (decodedTitles.length === 0) return;

  //   // 1) Find top-level site by title (CI)
  //   let currentNode =
  //     treeData.find((node) => (node.title || "").toLowerCase() === decodedTitles[0].toLowerCase()) || undefined;

  //   if (!currentNode) {
  //     console.warn("[navigateToPath] Top-level site not found for:", decodedTitles[0]);
  //     return;
  //   }

  //   await ensureExpanded(currentNode);

  //   // 2) Walk remaining segments
  //   for (let i = 1; i < decodedTitles.length; i++) {
  //     const seg = decodedTitles[i];
  //     console.log(`[navigateToPath] Seeking segment ${i}:`, seg, "under", currentNode.title, "type:", currentNode.type);

  //     // Make sure we have children loaded
  //     await ensureExpanded(currentNode);

  //     // Try exact child match first (CI)
  //     let nextNode = findChildByTitleCI(currentNode, seg);

  //     // If not found and we're at site/subsite, try auto-step into "Documents" library
  //     if (!nextNode && (currentNode.type === "site" || currentNode.type === "subsite")) {
  //       const docsNode = findChildByTitleCI(currentNode, "Documents");
  //       if (docsNode) {
  //         console.log(
  //           `[navigateToPath] No direct child '${seg}' under ${currentNode.type} '${currentNode.title}'. ` +
  //             "Auto-selecting 'Documents' library and continuing..."
  //         );
  //         await ensureExpanded(docsNode);

  //         // After entering Documents, try to resolve the same segment as a folder name under Documents
  //         nextNode = findChildByTitleCI(docsNode, seg);
  //         if (nextNode) {
  //           console.log("[navigateToPath] Found segment under 'Documents':", seg);
  //           currentNode = nextNode;
  //           continue;
  //         } else {
  //           // Maybe this seg is the library name itself but not loaded yet (rare), fall back to normal search after ensuring docs expanded
  //           currentNode = docsNode;
  //           // loop iteration will try seg again on next continue
  //           i--; // re-process same seg inside 'Documents'
  //           continue;
  //         }
  //       } else {
  //         console.log("[navigateToPath] 'Documents' library not present under current node.");
  //       }
  //     }

  //     if (!nextNode) {
  //       // If current is 'library' but folders not loaded yet, expand and retry
  //       if (currentNode.type === "library") {
  //         console.log("[navigateToPath] Not found. Ensuring library folders are loaded, then retrying:", seg);
  //         await ensureExpanded(currentNode);
  //         nextNode = findChildByTitleCI(currentNode, seg);
  //       }
  //     }

  //     if (!nextNode) {
  //       console.warn("[navigateToPath] Could not resolve path segment:", seg, "Stopping at:", currentNode.title);
  //       break;
  //     }

  //     currentNode = nextNode;
  //     console.log("[navigateToPath] Stepped into:", currentNode.title, "type:", currentNode.type);
  //   }

  //   console.log("[navigateToPath] Final target node:", { key: currentNode.key, title: currentNode.title, type: currentNode.type });
  //   await handleNodeClick(currentNode);
  // };
const navigateToPath = async (pathTitles: string[]) => {
  console.log("[navigateToPath] raw segments:", pathTitles);
  if (pathTitles.length === 0 || !treeData.length) return;

  const decodedTitles = pathTitles
    .map((t) => decodeURIComponent(t))
    .filter((t) => t.trim() !== "");

  let currentNode =
    treeData.find((n) => (n.title || "").toLowerCase() === decodedTitles[0].toLowerCase()) ||
    undefined;
  if (!currentNode) return;

  await ensureExpanded(currentNode); // Wait for children

  // Walk the path
  for (let i = 1; i < decodedTitles.length; i++) {
    const seg = decodedTitles[i];
    await ensureExpanded(currentNode);
    let nextNode = findChildByTitleCI(currentNode, seg);

    // Documents auto-step logic...
    if (!nextNode && (currentNode.type === "site" || currentNode.type === "subsite")) {
      const docsNode = findChildByTitleCI(currentNode, "Documents");
      if (docsNode) {
        await ensureExpanded(docsNode);
        nextNode = findChildByTitleCI(docsNode, seg);
        if (nextNode) {
          currentNode = nextNode;
          continue;
        } else {
          currentNode = docsNode;
          i--;
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
  }

  console.log("[navigateToPath] Final:", currentNode);

  // Force same logic as clicking
  console.log("[navigateToPath] Triggering click for node:", currentNode);
  handleNodeClick(currentNode);
};

  /** --------------------------------------- clicking ----------------------------------------- */
  const handleNodeClick = async (node: TreeNode) => {
    console.log("[handleNodeClick] node clicked:", { key: node.key, title: node.title, type: node.type, siteUrl: node.siteUrl });
    location = node.siteUrl;
    updateUrl(node);
    setActiveView("");

    // Case 1: Library or folder → load files
    if (node.type === "library" || node.type === "folder") {
      await loadFilesForNode(node);
      return;
    }

    // Case 2: Site or subsite → load its libraries first, then prefer 'Documents'
    if (node.type === "site" || node.type === "subsite") {
      console.log("[handleNodeClick] Site/Subsite: ensuring libraries are expanded");
      await ensureExpanded(node); // expands and fetches libraries

      // Try to find the 'Documents' library in the existing children (no transient nodes)
      const documentsNode = findChildByTitleCI(node, "Documents") || (node.children && node.children[0]);
      if (documentsNode) {
        console.log("[handleNodeClick] Will load files for library:", documentsNode.title);
        await loadFilesForNode(documentsNode);
      } else {
        console.log("[handleNodeClick] No libraries found under this site/subsite.");
      }
    }
  };

  /** ------------------------------------ load files (same) ----------------------------------- */
  const loadFilesForNode = async (node: TreeNode) => {
    try {
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
        console.log("[loadFilesForNode] Files loaded:", files?.length || 0);
        setSelectedFiles(files);
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

  /** ------------------------------------ expand/toggle (same) -------------------------------- */
  const toggleNode = async (node: TreeNode) => {
    if (!node.isExpanded && node.hasChildren) {
      const siteSP = spfi(node.siteUrl).using(SPFx(context));
      console.log("[toggleNode] Expanding:", { title: node.title, type: node.type, site: node.siteUrl });

      try {
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
      setCurrentFolderPath("");
      setCurrentSiteUrl("");
    } catch (error) {
      console.error(`[view] Error loading ${viewName} data:`, error);
      setSelectedFiles([]);
    }
  };

  const loadViewData = async (viewType: string): Promise<any[]> => {
    if (viewType === "MyRequest") {
      const configItems = await sp.web.lists
        .getByTitle("Allsitesfilemaster")
        .items.select("FileMaster", "sitecollection")
        .top(5000)();

      console.log("[loadViewData] configItems", configItems);

      const groupedBySite = configItems.reduce((acc: any, config: any) => {
        if (!acc[config.sitecollection]) {
          acc[config.sitecollection] = [];
        }
        acc[config.sitecollection].push(config);
        return acc;
      }, {} as Record<string, typeof configItems>);

      const allData: any[] = [];

      for (const [siteCollection, siteConfigs] of Object.entries(groupedBySite)) {
        try {
          console.log(`[loadViewData] Processing site collection: ${siteCollection}`);

          const currentUrl = new URL(context.pageContext.web.absoluteUrl);
          const siteUrl = `${currentUrl.protocol}//${currentUrl.hostname}/sites/${siteCollection}`;
          console.log(`[loadViewData] Target site URL: ${siteUrl}`);

          const siteSP = spfi(siteUrl).using(SPFx(context));

          const siteData = await Promise.all(
            (siteConfigs as any[]).map(async (config: any) => {
              try {
                console.log(`[loadViewData] Fetching from ${config.FileMaster} in ${siteCollection}`);
                return await siteSP.web.lists.getByTitle(config.FileMaster).items.top(1000)();
              } catch (error) {
                console.error(
                  `[loadViewData] Error fetching from ${config.FileMaster} in ${siteCollection}:`,
                  error
                );
                return [];
              }
            })
          );

          allData.push(...siteData.flat());
        } catch (error) {
          console.error(`[loadViewData] Error processing site collection ${siteCollection}:`, error);
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
        {breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== "view" && (
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
        {!showUploadPanel &&
          (activeView === "browse" ? (
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
                    <span style={{ fontSize: "14px" }}>{file.Name}hj</span>
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
          ) : selectedFiles.length > 0 ? (
            <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
              {selectedFiles.map((file) => (
                <li
                  key={file.Id}
                  style={{
                    padding: "8px 0",
                    borderBottom: "1px solid #eee",
                    display: "flex",
                    alignItems: "center",
                  }}
                >
                  <span style={{ marginRight: "8px" }}>📄</span>
                  <span style={{ fontSize: "14px" }}>{file.Title || file.FileName}</span>
                  <span style={{ fontSize: "14px" }}>
                    {file.Name || file.FileLeafRef || file.Title || file.File?.Name || "Unnamed File"}
                  </span>
                  <span
                    style={{
                      fontSize: "14px",
                      marginLeft: "10px",
                      color: "#666",
                    }}
                  >
                    {file.Modified ? new Date(file.Modified).toLocaleDateString() : ""}
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
              No files found in {activeView}
            </div>
          ))}
      </div>
    </div>
  );
};

const DMSMain: React.FC<IEssadmsMainProps> = (props) => {
  return <ArgPoc context={props.context} />;
};

export default DMSMain;
