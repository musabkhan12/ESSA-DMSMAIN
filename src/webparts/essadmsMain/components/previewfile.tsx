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

  const previewUrl = fileUrl.FilePreviewURL
    ? fileUrl.FilePreviewURL
    : `https://officeindia.sharepoint.com${fileUrl.ServerRelativeUrl}`;
   
  // here find extension 
  const fileExtension = fileUrl.FileName.split('.').pop().toLowerCase();
  console.log("File Extension:", fileExtension);




  const officeExtensions = ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'csv'];
  console.log("Office Extensions:", officeExtensions);










  const newpreviewUrl = fileUrl.FilePreviewURL;
  console.log("New Preview URL:", newpreviewUrl);
  const url = newpreviewUrl;
  const params = new URLSearchParams(url.split("?")[1]);
  const encodedId = params.get("id");
  const decodedId = decodeURIComponent(encodedId || "");
  console.log(decodedId);

const previewUrlS =
  `${fileUrl.__siteUrl}/${fileUrl.SiteName}/_layouts/15/Doc.aspx?sourcedoc=${decodedId}` ;

 
  

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
  }, [fileUrl]);

  return (
    <div style={{ width: "100%", height: "100%" }}>
      <button onClick={onClose} style={{ marginBottom: "10px" }}>Close</button>

      {/* ✅ HEADER SECTION */}
      <div style={{
        padding: "10px 0",
        borderBottom: "1px solid #ddd",
        marginBottom: "10px"
      }}>
        {/* Show Quick View Name if present */}
        {viewName && (
          <div style={{
            fontWeight: 600,
            fontSize: "16px",
            marginBottom: "4px",
            color: "#0078d4"
          }}>
            📋 {viewName}
          </div>
        )}

        {/* Show file name */}
        {(fileName || fileUrl?.FileName) && (
          <div style={{
            fontWeight: 600,
            fontSize: "16px",
            marginBottom: "4px",
            color: "#333"
          }}>
            📄 {fileName || fileUrl.FileName}
          </div>
        )}

        {/* Show folder path ONLY if not a quick view */}
        {!viewName && (filePath || fileUrl?.CurrentFolderPath) && (
          <div style={{
            fontSize: "12px",
            color: "#666",
            fontWeight: 400
          }}>
            📁 {filePath || fileUrl.CurrentFolderPath}
          </div>
        )}
      </div>

      <iframe
        ref={previewFrameRef}
        src={officeExtensions.includes(fileExtension) ? previewUrlS : previewUrl}
        width="100%"
        height="600px"
        style={{ border: "none" }}
      />
    </div>
  );
};

export default PreviewModal;
