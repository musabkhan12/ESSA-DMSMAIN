// Preview section handle by Addhyan code change on 20/2/26 
import * as React from "react";
import { useEffect, useRef } from "react";
 
interface PreviewModalProps {
  show: boolean;
  fileUrl?: any;
  fileName?: string;
  filePath?: string;
  viewName?: string;        // ✅ New prop for quick view name
  onClose: () => void;
}
 
const PreviewModal: React.FC<PreviewModalProps> = ({
  show,
  fileUrl,
  fileName,
  filePath,
  viewName,
  onClose
}) => {
 
  console.log("PreviewModal Props:", {  show, fileUrl, fileName, filePath, viewName });
 
  if (!show || !fileUrl) return null;
 
  const previewFrameRef = useRef<HTMLIFrameElement>(null);
 
  // ✅ FIX FOR SVG/IMAGE DOWNLOAD: Added ?Web=1 here to handle non-office files in Hierarchy
  const previewUrl = fileUrl.FilePreviewURL
    ? fileUrl.FilePreviewURL
    : `https://officeindia.sharepoint.com${fileUrl.ServerRelativeUrl}?Web=1`;
   
  // Extension check (Safe for Hierarchy)
  const safeFileName = fileName || fileUrl?.FileName || fileUrl?.Name || "";
  const fileExtension = safeFileName.split('.').pop().toLowerCase();
  console.log("File Extension:", fileExtension);
 
  const officeExtensions = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'csv'];
  console.log("Office Extensions:", officeExtensions);
 
  const newpreviewUrl = fileUrl.FilePreviewURL;
  console.log("New Preview URL:", newpreviewUrl);
 
  let previewUrlS = "";
 
  // Editable Logic (My Requests, Folders etc.)
  if (newpreviewUrl && typeof newpreviewUrl === 'string' && newpreviewUrl.includes("?")) {
    const url = newpreviewUrl;
    const params = new URLSearchParams(url.split("?")[1]);
    const encodedId = params.get("id");
    const decodedId = decodeURIComponent(encodedId || "");
   
    previewUrlS = `${fileUrl.__siteUrl}/${fileUrl.SiteName}/_layouts/15/Doc.aspx?sourcedoc=${decodedId}`;
  } else {
    // FIX FOR DOWNLOAD ISSUE: Added ?Web=1 to force browser preview
    previewUrlS = fileUrl.FilePreviewURL
      ? fileUrl.FilePreviewURL
      : `https://officeindia.sharepoint.com${fileUrl.ServerRelativeUrl}?Web=1`;
  }
 
  console.log("Previews URL:", previewUrlS);
 
  const checkAndHideButton = () => {
    let iframeDocument: Document | null = null;
    try {
      const iframeEl = previewFrameRef.current;
      if (!iframeEl) return;
      iframeDocument = iframeEl.contentDocument || iframeEl.contentWindow?.document;
      if (iframeDocument) {
        const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
        const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
        if (button) button.style.display = "none";
        if (excelToolbar) excelToolbar.style.display = "none";
        iframeEl.style.display = "block";
      }
    } catch (error) {
      console.error("Error accessing iframe content:", error);
      if (previewFrameRef.current) previewFrameRef.current.style.display = "block";
    }
    if (!iframeDocument || !iframeDocument.getElementById("OneUpCommandBar")) {
      setTimeout(checkAndHideButton, 100);
    }
  };
 
  useEffect(() => {
    checkAndHideButton();
  }, [fileUrl, previewUrlS]);
 
  return (
    <div style={{ width: "100%", height: "100%" }}>
        <button style={{background:'#fff', position:'absolute', top:'13px', right:'0px'}} className="me-2 mt-0 " id="CreateFolderInsideSharePoint">
           <span className="mb-1 mt-2" data-tooltip="Back">
             <img src={require('../assets/backn.png')} alt="Back" onClick={onClose} />
           </span>
        </button>
 
      <div style={{
        padding: "0px 177px 10px 17px",
        borderBottom: "0px solid #ddd",
        marginBottom: "0px"
      }}>
        {(fileName || fileUrl?.FileName) && (
          <div style={{ fontWeight: 600, fontSize: "16px", marginBottom: "4px", color: "#333" }}>
             {fileName || fileUrl.FileName}
          </div>
        )}
        {/* {!viewName && (filePath || fileUrl?.CurrentFolderPath) && (
          <div style={{ fontSize: "12px", color: "#666", fontWeight: 400 }}>
            📁 {filePath || fileUrl.CurrentFolderPath}
          </div>
        )} */}
      </div>
 
      <iframe
        ref={previewFrameRef}
        src={officeExtensions.includes(fileExtension) ? previewUrlS : previewUrl}
        width="100%"
        height="800"
        style={{ border: "none", padding:'0px 15px 15px 15px' }}
      />
    </div>
  );
};
 
export default PreviewModal;
 