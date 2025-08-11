// import * as React from "react";
// import { useState, useEffect } from "react";
// import { spfi, SPFI, SPFx } from "@pnp/sp";
// import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { IHelloWorldProps } from "../components/IHelloWorldProps";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/folders";
// import "@pnp/sp/files";

// interface TreeNode {
//   key: string;
//   title: string;
//   type: "site" | "subsite" | "library" | "folder";
//   children?: TreeNode[];
//   hasChildren?: boolean;
//   isExpanded?: boolean;
//   siteUrl: string;
//   libraryTitle?: string;
//   folderPath?: string;
// }


// const ArgPoc = ({ context }: { context: WebPartContext }) => {
//   const [sp] = useState(() => spfi().using(SPFx(context)));
//   const [treeData, setTreeData] = useState<TreeNode[]>([]);

//   useEffect(() => {
//     loadRootSites();
//   }, []);

//   const loadRootSites = async () => {
//     const siteList = await sp.web.lists.getByTitle("MasterSiteCollection").items.getAll();
//     const subsiteList = await sp.web.lists.getByTitle("MasterSiteURL").items.getAll();

//     const nodes: TreeNode[] = siteList.map(site => {
//       const siteUrl = site.SiteURL?.trim();
//       const siteSubs = subsiteList.filter(s => s.SiteURL?.startsWith(siteUrl));

//       return {
//         key: `site-${site.Id}`,
//         title: site.Title,
//         type: "site",
//         siteUrl: siteUrl,
//         hasChildren: siteSubs.length > 0,
//         children: siteSubs.map(sub => ({
//           key: `subsite-${sub.Id}`,
//           title: sub.Title,
//           type: "subsite",
//           siteUrl: sub.SiteURL?.trim(),
//           hasChildren: true,
//         }))
//       };
//     });
//     setTreeData(nodes);
//   };

//   const toggleNode = async (node: TreeNode, parent?: TreeNode) => {
//     if (!node.isExpanded && node.hasChildren) {
//       const siteSP = spfi(node.siteUrl).using(SPFx(context));

//       if (node.type === "subsite") {
//         const libs = await siteSP.web.lists.filter("BaseTemplate eq 101 and Hidden eq false").select("Title")();
//         node.children = libs.map(lib => ({
//           key: `lib-${node.key}-${lib.Title}`,
//           title: lib.Title,
//           type: "library",
//           siteUrl: node.siteUrl,
//           libraryTitle: lib.Title,
//           hasChildren: true
//         }));
//       }

//       if (node.type === "library") {
//         // const folders = await siteSP.web.getFileByServerRelativePath(`/sites/${node.siteUrl.split("/sites/")[1]}/${node.libraryTitle}`).folders();
//            const folders = await siteSP.web.getFolderByServerRelativePath(`/sites/${node.siteUrl.split("/sites/")[1]}/${node.libraryTitle}`).folders();
//         node.children = folders
//           .filter((f:any) => !f.Name.startsWith("Forms"))
//           .map((f:any) => {
//   console.log("Folder Path:", `${node.libraryTitle}/${f.Name}`);
//   console.log("Folder Path f:", `${JSON.stringify(f)}`);
//   return {
//     key: `folder-${f.Name}`,
//     title: f.Name,
//     type: "folder",
//     siteUrl: node.siteUrl,
//     libraryTitle: node.libraryTitle,
//     folderPath: `${node.libraryTitle}/${f.Name}`,
//     hasChildren: true,
//   };
// });
//           // .map((f:any) => ({
//           //   key: `folder-${f.Name}`,
//           //   title: f.Name,
//           //   type: "folder",
//           //   siteUrl: node.siteUrl,
//           //   libraryTitle: node.libraryTitle,
//           //   folderPath: `${node.libraryTitle}/${f.Name}`,
//           //   hasChildren: true,
            
//           // }
        
//       }

//       if (node.type === "folder") {
//         const subFolders = await siteSP.web.getFolderByServerRelativePath(`/sites/${node.siteUrl.split("/sites/")[1]}/${node.folderPath}`).folders();
//         node.children = subFolders
//           .filter((f:any) => !f.Name.startsWith("Forms"))
//           .map((f:any) => ({
//             key: `subfolder-${f.Name}`,
//             title: f.Name,
//             type: "folder",
//             siteUrl: node.siteUrl,
//             libraryTitle: node.libraryTitle,
//             folderPath: `${node.folderPath}/${f.Name}`,
//             hasChildren: true
//           }));
//       }

//       node.isExpanded = true;
//       setTreeData([...treeData]);
//     } else {
//       node.isExpanded = !node.isExpanded;
//       setTreeData([...treeData]);
//     }
//   };

//   const renderTree = (nodes: TreeNode[]) => (
//     <ul className="ml-4">
//       {nodes.map(node => (
//         <li key={node.key}>
//           {node.hasChildren && (
//             <button onClick={() => toggleNode(node)}>
//               {node.isExpanded ? "−" : "+"}
//             </button>
//           )}
//           <span className="ml-2">{node.title}</span>
//           {node.isExpanded && node.children && renderTree(node.children)}
//         </li>
//       ))}
//     </ul>
//   );

//   return (
//     <div className="p-4">
//       <h2 className="text-lg font-semibold mb-2">Folder Hierarchy</h2>
//       {renderTree(treeData)}
//     </div>
//   );
// };

// const DMSMain: React.FC<IHelloWorldProps> = (props) => {
//   return (
//     // <Provider>
//       <ArgPoc context={props.context} />

//     // </Provider>
//   );
// };

// export default DMSMain;
import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IEssadmsMainProps } from './IEssadmsMainProps';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/site-users/web";
// import { Web } from '@pnp/sp/webs';
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
  useEffect(() => {
    const initialize = async () => {
      await loadRootSites();
      
      // Check URL hash on initial load
      const hash = decodeURIComponent(window.location.hash.substring(1));
      if (hash) {
        await navigateToPath(hash.split('/'));
      } else {
        // Default to "My request" view when URL is empty
        handleViewButtonClick("My request");
      }
    };

    initialize();

    const handleHashChange = () => {
      const hash = decodeURIComponent(window.location.hash.substring(1));
      if (hash) {
        navigateToPath(hash.split('/'));
      } else {
        handleViewButtonClick("My request");
      }
    };

    window.addEventListener('hashchange', handleHashChange);
    return () => window.removeEventListener('hashchange', handleHashChange);
  }, []);

  const loadRootSites = async () => {
    try {
      const [siteList, subsiteList] = await Promise.all([
        sp.web.lists.getByTitle("MasterSiteCollection").items.select("Id", "Title", "SiteURL").top(5000)(),
        sp.web.lists.getByTitle("MasterSiteURL").items.select("Description","Title","SiteURL","Active","SiteID","FileMasterList").top(5000)()
      ]);

      const nodes: TreeNode[] = siteList.map(site => {
        const siteUrl = site.SiteURL?.trim();
        const siteSubs = subsiteList.filter(s => s.SiteURL?.startsWith(siteUrl));
        return {
          key: `site-${site.Id}`,
          title: site.Title,
          type: "site",
          siteUrl: siteUrl,
          hasChildren: siteSubs.length > 0,
          children: siteSubs.map(sub => ({
            key: `subsite-${sub.Id}`,
            title: sub.Title,
            type: "subsite",
            siteUrl: sub.SiteURL?.trim(),
            hasChildren: true,
            parentKey: `site-${site.Id}`
          }))
        };
      });
      console.log("Root sites loaded:", nodes);
      setTreeData(nodes);
      buildNodeMap(nodes);
    } catch (error) {
      console.error("Error loading sites:", error);
    }
  };

  const buildNodeMap = (nodes: TreeNode[], map: Record<string, TreeNode> = {}) => {
    nodes.forEach(node => {
      map[node.key] = node;
      if (node.children) {
        buildNodeMap(node.children, map);
      }
    });
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
        folderPath: currentNode.folderPath
      });
      if (currentNode.parentKey) {
        currentNode = nodeMap[currentNode.parentKey];
      } else {
        break;
      }
    }
    return path;
  };

  const updateUrl = (node: TreeNode) => {
    const path = getNodePath(node.key).map(item => encodeURIComponent(item.title)).join('/');
    window.history.pushState(null, '', `#${path}`);
  };

  const navigateToPath = async (pathTitles: string[]) => {
    if (pathTitles.length === 0 || !treeData.length) return;
    
    const decodedTitles = pathTitles.map(decodeURIComponent).filter(title => title.trim() !== '');
    if (decodedTitles.length === 0) return;

    let currentNode = treeData.find(node => node.title === decodedTitles[0]);
    if (!currentNode) return;

    if (!currentNode.isExpanded) {
      await toggleNode(currentNode);
    }

    for (let i = 1; i < decodedTitles.length; i++) {
      const nextTitle = decodedTitles[i];
      const nextNode = currentNode.children?.find(child => child.title === nextTitle);
      if (nextNode) {
        if (!nextNode.isExpanded) {
          await toggleNode(nextNode);
        }
        currentNode = nextNode;
      } else {
        break;
      }
    }

    await handleNodeClick(currentNode);
  };

  const handleNodeClick = async (node: TreeNode) => {
    console.log(node, "node clicked")
   location = node.siteUrl
    updateUrl(node);
    setActiveView("");
    if (node.type === "library" || node.type === "folder") {
      await loadFilesForNode(node);
    }
  };

  const loadFilesForNode = async (node: TreeNode) => {
    try {
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
        const files = await siteSP.web.getFolderByServerRelativePath(
          `/sites/${node.siteUrl.split("/sites/")[1]}/${folderPath}`
        ).files();
        
        setSelectedFiles(files);
        setBreadcrumbs(getNodePath(node.key));
      }
    } catch (error) {
      console.error("Error fetching files:", error);
      setSelectedFiles([]);
    }
  };

  const toggleNode = async (node: TreeNode) => {
    if (!node.isExpanded && node.hasChildren) {
      const siteSP = spfi(node.siteUrl).using(SPFx(context));
      console.log("Expanding node:", node.title, "at site:", node.siteUrl);
      console.log("siteSP:", siteSP);
      try {
        if (node.type === "subsite") {
          const libs = await siteSP.web.lists
            .filter("BaseTemplate eq 101 and Hidden eq false")
            .select("Title")();
          
          node.children = libs.map(lib => ({
            key: `lib-${node.key}-${lib.Title}`,
            title: lib.Title,
            type: "library",
            siteUrl: node.siteUrl,
            libraryTitle: lib.Title,
            hasChildren: true,
            parentKey: node.key
          }));
        }

        if (node.type === "library") {
          const folders = await siteSP.web
            .getFolderByServerRelativePath(`/sites/${node.siteUrl.split("/sites/")[1]}/${node.libraryTitle}`)
            .folders();
          
          node.children = folders
            .filter((f: any) => !f.Name.startsWith("Forms"))
            .map((f: any) => ({
              key: `folder-${f.Name}`,
              title: f.Name,
              type: "folder",
              siteUrl: node.siteUrl,
              libraryTitle: node.libraryTitle,
              folderPath: `${node.libraryTitle}/${f.Name}`,
              hasChildren: true,
              parentKey: node.key
            }));
        }

        if (node.type === "folder") {
          const subFolders = await siteSP.web
            .getFolderByServerRelativePath(`/sites/${node.siteUrl.split("/sites/")[1]}/${node.folderPath}`)
            .folders();
          
          node.children = subFolders
            .filter((f: any) => !f.Name.startsWith("Forms"))
            .map((f: any) => ({
              key: `subfolder-${f.Name}`,
              title: f.Name,
              type: "folder",
              siteUrl: node.siteUrl,
              libraryTitle: node.libraryTitle,
              folderPath: `${node.folderPath}/${f.Name}`,
              hasChildren: true,
              parentKey: node.key
            }));
        }

        node.isExpanded = true;
        setTreeData([...treeData]);
        buildNodeMap(treeData);
      } catch (error) {
        console.error("Error expanding node:", error);
      }
    } else {
      node.isExpanded = !node.isExpanded;
      setTreeData([...treeData]);
    }
  };

  const handleBreadcrumbClick = async (item: BreadcrumbItem) => {
    if (item.type === 'view') {
      handleViewButtonClick(item.title);
    } else {
      const node = nodeMap[item.key];
      if (node) {
        await expandPathToNode(node.key);
        await handleNodeClick(node);
      }
    }
  };

  const expandPathToNode = async (nodeKey: string) => {
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

  const handleViewButtonClick = async (viewName: string) => {
    console.log("View button clicked:", viewName);
    setActiveView(viewName);
    setBreadcrumbs([{
      key: viewName.toLowerCase().replace(/\s+/g, '-'),
      title: viewName,
      type: 'view',
      siteUrl: ''
    }]);
    
    // Clear any URL hash when switching to a view
    window.history.pushState(null, '', window.location.pathname);
    
    try {
      let files = [];
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
      console.log(`Loaded ${viewName} data:`, files);
      setSelectedFiles(files);
      setCurrentFolderPath("");
      setCurrentSiteUrl("");
    } catch (error) {
      console.error(`Error loading ${viewName} data:`, error);
      setSelectedFiles([]);
    }
  };

const loadViewData = async (viewType: string): Promise<any[]> => {
  if (viewType === "MyRequest") {
    // 1. Get configuration from Allsitesfilemaster list
    const configItems = await sp.web.lists.getByTitle("Allsitesfilemaster")
      .items
      .select('FileMaster', 'sitecollection')
      .top(5000)();
    
    console.log("configItems", configItems);



    // Group config items by sitecollection
    const groupedBySite = configItems.reduce((acc, config) => {
      if (!acc[config.sitecollection]) {
        acc[config.sitecollection] = [];
      }
      acc[config.sitecollection].push(config);
      return acc;
    }, {} as Record<string, typeof configItems>);

    const allData = [];
    
    // 2. Process each site collection
    for (const [siteCollection, siteConfigs] of Object.entries(groupedBySite)) {
      try {
        console.log(`Processing site collection: ${siteCollection}`);
        
        // Create proper URL for the target site collection
        const currentUrl = new URL(context.pageContext.web.absoluteUrl);
        const siteUrl = `${currentUrl.protocol}//${currentUrl.hostname}/sites/${siteCollection}`;
        console.log(`Target site URL: ${siteUrl}`);
        console.log(`Target currentUrl: ${currentUrl}`);
        // Create new SP client for the target site
        const siteSP = spfi(siteUrl).using(SPFx(context));
        
        // Process all FileMaster lists for this site collection
        const siteData = await Promise.all(
          (siteConfigs as any[]).map(async (config:any) => {
            try {
              console.log(`Fetching from ${config.FileMaster} in ${siteCollection}`);
              return await siteSP.web.lists.getByTitle(config.FileMaster)
                .items
                .top(1000)();
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

  // File Upload Functions
const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
  if (e.target.files && e.target.files.length > 0) {
    const files = Array.from(e.target.files);
    setSelectedUploadFiles(files);
    
    try {
      // @ts-ignore
      function getSiteCollectionUrl(fullUrl: any): any {
  const pattern = /^(https?:\/\/[^\/]+\/sites\/[^\/]+)/i;
  const match = fullUrl.match(pattern);
  return match ? match[1] : fullUrl;
        }
        console.log("getSiteCollectionUrl Uploading files to preview location..." , getSiteCollectionUrl(currentSiteUrl));
      const siteSP = spfi(getSiteCollectionUrl(currentSiteUrl)).using(SPFx(context));
      const previewUrls: string[] = [];
        console.log("currentSiteUrl Uploading files to preview location..." , currentSiteUrl);
        console.log("siteSP Uploading files to preview location..." , siteSP);
        console.log("currentFolderPath Uploading files to preview location..." , `/sites/${currentSiteUrl.split("/sites/")[1]}/DMSOrphanDocs`);
      // Upload each file to preview location immediately
      for (const file of files) {
        const result = await siteSP.web.getFolderByServerRelativePath(
          `/sites/AlRostmaniSpfx2/DMSOrphanDocs`
        ).files.addChunked(file.name, file);
        
        previewUrls.push(result.data.ServerRelativeUrl);
      }
      
      setPreviewFileUrls(previewUrls);
    } catch (error) {
      console.error("Error uploading preview files:", error);
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
      
      // Upload to actual destination
      const uplaodfile = await siteSP.web.getFolderByServerRelativePath(
        `/sites/${currentSiteUrl.split("/sites/")[1]}/${currentFolderPath}`
      ).files.addChunked(file.name, file);
      console.log("File uploaded:", uplaodfile.data , JSON.stringify(uplaodfile.data));
       

   /**
 * Encodes a SharePoint file/folder path for URL usage (matches SharePoint's encoding rules)
 * @param path The path to encode (e.g., `/sites/siteName/DocLib/Folder/File (1).txt`)
 * @returns Encoded path (e.g., `%2Fsites%2FsiteName%2FDocLib%2FFolder%2FFile%20%281%29%2Etxt`)
 */
const encodeSharePointPath = (path: string): string => {
  return encodeURIComponent(path)
    .replace(/'/g, "%27")       // Single quote
    .replace(/\(/g, "%28")      // Opening parenthesis
    .replace(/\)/g, "%29")      // Closing parenthesis
    .replace(/\*/g, "%2A")      // Asterisk
    .replace(/!/g, "%21")       // Exclamation mark
    .replace(/#/g, "%23")       // Hash
    .replace(/\$/g, "%24")      // Dollar sign
    .replace(/&/g, "%26")       // Ampersand
    .replace(/\+/g, "%2B")      // Plus sign
    .replace(/,/g, "%2C")       // Comma
    .replace(/;/g, "%3B")       // Semicolon
    .replace(/=/g, "%3D")       // Equals sign
    .replace(/\?/g, "%3F")      // Question mark
    .replace(/\[/g, "%5B")      // Opening square bracket
    .replace(/\]/g, "%5D")      // Closing square bracket
    .replace(/_/g, "%5F")       // Underscore
    .replace(/\./g, "%2E")      // Period
    .replace(/-/g, "%2D");     // Hyphen
};

/**
 * Generates a SharePoint preview URL for a file
 * @param siteUrl Base site URL (e.g., `https://tenant.sharepoint.com/sites/siteName`)
 * @param serverRelativePath File's server-relative path (e.g., `/sites/siteName/DocLib/Folder/File.docx`)
 * @returns Full preview URL (e.g., `https://tenant.sharepoint.com/sites/siteName/DocLib/Folder/Forms/AllItems.aspx?...`)
 */
const getSharePointPreviewUrl = (siteUrl: string, serverRelativePath: string): string => {
  const parentFolder = serverRelativePath.substring(0, serverRelativePath.lastIndexOf("/"));
  const encodedFilePath = encodeSharePointPath(serverRelativePath);
  const encodedParentPath = encodeSharePointPath(parentFolder);
  
  return `${siteUrl}${parentFolder.replace(/ /g, "%20")}/Forms/AllItems.aspx?id=${encodedFilePath}&parent=${encodedParentPath}`;
};

// Example Usage
const siteUrl = "https://officeindia.sharepoint.com";
const filePath = uplaodfile.data.ServerRelativeUrl;
const previewUrl = getSharePointPreviewUrl(siteUrl, filePath);
console.log(previewUrl , "Preview URL for uploaded file:");   
      // FileName: String(uploadResult.data.Name),
            // FileSize: String(uploadResult.data.Length),
            // FileVersion: String(uploadResult.data.MajorVersion),
            // CurrentFolderPath: String(currentfolderpath.folderpath),
            // FileUID: String(uploadResult.data.UniqueId),
            // CurrentUser: String(currentUserEmailRef.current),
            // SiteID: String(currentfolderpath.siteID),
            // Status: "Auto Approved",
            // FilePreviewURL: String(previewUrl),
            // DocumentLibraryName: String(currentfolderpath.DocumentLibrary),
            // SiteName: String(currentfolderpath.Entity),
            // MyRequest: true,
            // Processname: 'New File Request',
    }
    
    // Refresh the file list
    if (breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== 'view') {
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
    console.error("Error uploading files:", error);
  }
};

  const renderTree = (nodes: TreeNode[]) => (
    <ul style={{ listStyleType: "none", paddingLeft: "20px" }}>
      {nodes.map(node => (
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
                  justifyContent: "center"
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
                fontWeight: breadcrumbs.some(b => b.key === node.key) ? "bold" : "normal",
                backgroundColor: breadcrumbs.some(b => b.key === node.key) ? "#f0f0f0" : "transparent",
                padding: "2px 5px",
                borderRadius: "3px",
                fontSize: "14px"
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
    <div id="maincontainer" style={{ display: "flex", height: "calc(100vh - 50px)", marginTop: "10px" }}>
      {/* Left Panel with Quick Views and Folder Hierarchy */}
      <div style={{ width: "30%", display: "flex", flexDirection: "column" }}>
        {/* Quick Views Panel */}
        <div id="buttonpanel" style={{ 
          padding: "15px", 
          backgroundColor: "#f0f0f0", 
          borderBottom: "1px solid #ddd",
          flexShrink: 0
        }}>
          <h2 style={{ fontSize: "16px", fontWeight: "600", marginBottom: "15px", color: "#333" }}>
            Quick Views
          </h2>
          {["My request", "My favourite", "Share with me", "Share with other", "Recycle bin"].map((view) => (
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
                transition: "all 0.2s"
              }}
            >
              {view}
            </button>
          ))}
        </div>

        {/* Folder Hierarchy Panel */}
        <div id="folderhierarchycontainer" style={{ 
          flexGrow: 1,
          padding: "15px",
          overflow: "auto",
          backgroundColor: "#f9f9f9"
        }}>
          <h2 style={{ 
            fontSize: "16px", 
            fontWeight: "600", 
            marginBottom: "15px", 
            color: "#333", 
            paddingBottom: "5px", 
            borderBottom: "1px solid #eee" 
          }}>
            Folder Hierarchy
          </h2>
          {treeData.length > 0 ? renderTree(treeData) : <p>Loading folder structure...</p>}
        </div>
      </div>

      {/* File List Panel */}
      <div id="filelistcontainer" style={{ 
        width: "70%", 
        padding: "15px", 
        overflow: "auto", 
        backgroundColor: "#fff",
        position: "relative"
      }}>
        {/* Upload File Button - Only shown when in a folder/library */}
        {breadcrumbs.length > 0 && breadcrumbs[breadcrumbs.length - 1].type !== 'view' && (
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
              fontSize: "14px"
            }}
          >
            Upload File
          </button>
        )}

        <h2 style={{ 
          fontSize: "16px", 
          fontWeight: "600", 
          marginBottom: "15px", 
          color: "#333", 
          paddingBottom: "5px", 
          borderBottom: "1px solid #eee",
          paddingRight: "100px" // Make space for upload button
        }}>
          {breadcrumbs.length > 0 ? (
            <div style={{ display: "flex", alignItems: "center", flexWrap: "wrap" }}>
              {breadcrumbs.map((item, index) => (
                <React.Fragment key={item.key}>
                  {index > 0 && <span style={{ margin: "0 8px", color: "#999", fontSize: "14px" }}>›</span>}
                  <span
                    className="breadcrumb-item"
                    onClick={() => handleBreadcrumbClick(item)}
                    style={{
                      cursor: "pointer",
                      color: index === breadcrumbs.length - 1 ? "#333" : "#0066cc",
                      fontWeight: index === breadcrumbs.length - 1 ? "600" : "normal",
                      fontSize: "14px",
                      padding: "2px 5px",
                      borderRadius: "3px"
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
  <div style={{ 
    marginBottom: "20px",
    padding: "15px",
    backgroundColor: "#f5f5f5",
    borderRadius: "4px",
    border: "1px solid #ddd"
  }}>
    <h3 style={{ fontSize: "14px", marginBottom: "10px" }}>Upload Files</h3>
    <input
      type="file"
      ref={fileInputRef}
      onChange={handleFileSelect}
      multiple
      style={{ display: "none" }}
    />
    <button
      onClick={triggerFileInput}
      style={{
        padding: "8px 15px",
        backgroundColor: "#0078d4",
        color: "white",
        border: "none",
        borderRadius: "4px",
        cursor: "pointer",
        marginRight: "10px"
      }}
    >
      Choose Files
    </button>
    
    {selectedUploadFiles.length > 0 && (
      <div style={{ marginTop: "10px" }}>
        <h4 style={{ fontSize: "13px", marginBottom: "5px" }}>Selected Files:</h4>
        <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
          {selectedUploadFiles.map((file, index) => (
            <li key={index} style={{ padding: "5px 0", fontSize: "13px", display: "flex", alignItems: "center" }}>
              <span style={{ marginRight: "10px" }}>📄</span>
              <div>
                <div>{file.name}</div>
                {previewFileUrls[index] && (
                  <div style={{ fontSize: "11px", color: "#666", marginTop: "2px" }}>
                    Preview: <a href={previewFileUrls[index]} target="_blank" rel="noopener noreferrer">View File</a>
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
            marginTop: "10px"
          }}
        >
          Upload Files to Destination
        </button>
        {uploadProgress > 0 && (
          <div style={{ marginTop: "10px" }}>
            <div style={{ 
              width: "100%", 
              backgroundColor: "#e0e0e0", 
              borderRadius: "4px",
              height: "20px"
            }}>
              <div style={{ 
                width: `${uploadProgress}%`, 
                backgroundColor: "#0078d4", 
                height: "100%",
                borderRadius: "4px",
                transition: "width 0.3s"
              }}></div>
            </div>
            <div style={{ textAlign: "center", marginTop: "5px", fontSize: "12px" }}>
              {Math.round(uploadProgress)}% Complete
            </div>
          </div>
        )}
      </div>
    )}
  </div>
)}

        {/* File List */}

        {/* {selectedFiles.length > 0 ? (
          <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
            {selectedFiles.map(file => (
              <li key={file.Name} style={{ 
                padding: "8px 0", 
                borderBottom: "1px solid #eee", 
                display: "flex", 
                alignItems: "center" 
              }}>
                <span style={{ marginRight: "8px" }}>📄</span>
                <span style={{ fontSize: "14px" }}>{file.FileName}</span>
                <span style={{ fontSize: "14px" }}>{file.CurrentFolderPath}</span>
              </li>
            ))}
          </ul>
        ) : (
          <div style={{ 
            color: "#666", 
            fontSize: "14px", 
            padding: "20px", 
            textAlign: "center", 
            marginTop: "20px" 
          }}>
            {breadcrumbs.length > 0 ? "No files in this folder" : "Please select a folder from the hierarchy to view files"}
          </div>
        )} */}
{/* Main render section */}
{activeView === 'browse' ? (
  // Display files from folder/library browsing
  selectedFiles.length > 0 ? (
    <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
      {selectedFiles.map(file => (
        <li key={file.Name} style={{ 
          padding: "8px 0", 
          borderBottom: "1px solid #eee", 
          display: "flex", 
          alignItems: "center" 
        }}>
          <span style={{ marginRight: "8px" }}>📄</span>
          <span style={{ fontSize: "14px" }}>{file.Name}</span>
          <span style={{ fontSize: "14px", marginLeft: "10px", color: "#666" }}>
            {currentFolderPath}
          </span>
        </li>
      ))}
    </ul>
  ) : (
    <div style={{ 
      color: "#666", 
      fontSize: "14px", 
      padding: "20px", 
      textAlign: "center", 
      marginTop: "20px" 
    }}>
      {breadcrumbs.length > 0 ? "No files in this folder" : "Please select a folder from the hierarchy to view files"}
    </div>
  )
) : (
  // Display files from view buttons (My request, etc.)
  selectedFiles.length > 0 ? (
    <ul style={{ listStyleType: "none", paddingLeft: "0" }}>
      {selectedFiles.map(file => (
        <li key={file.Id} style={{ 
          padding: "8px 0", 
          borderBottom: "1px solid #eee", 
          display: "flex", 
          alignItems: "center" 
        }}>
          <span style={{ marginRight: "8px" }}>📄</span>
          <span style={{ fontSize: "14px" }}>{file.Title || file.FileName}</span>
          <span style={{ fontSize: "14px", marginLeft: "10px", color: "#666" }}>
            {file.Modified ? new Date(file.Modified).toLocaleDateString() : ''}
          </span>
        </li>
      ))}
    </ul>
  ) : (
    <div style={{ 
      color: "#666", 
      fontSize: "14px", 
      padding: "20px", 
      textAlign: "center", 
      marginTop: "20px" 
    }}>
      No files found in {activeView}
    </div>
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