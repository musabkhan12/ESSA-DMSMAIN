// old code musaib
// import * as React from "react";

// interface PreviewModalProps {
//   show: boolean;
//   fileUrl?: any;
//   onClose: () => void;
// }

// const PreviewModal: React.FC<PreviewModalProps> = ({ show, fileUrl, onClose }) => {
//   if (!show) return null;

//   return (
//     <div className="modal-overlay">
//       <div className="modal-content">
//         <button onClick={onClose}>Close</button>

    
//           <iframe src={fileUrl.FilePreviewURL} width="100%" height="600px" />
      
//       </div>
//     </div>
//   );
// };

// export default PreviewModal;


import * as React from "react";
import { useEffect, useRef } from "react";
 
interface PreviewModalProps {
  show: boolean;
  fileUrl?: any;
  onClose: () => void;
}
 
const PreviewModal: React.FC<PreviewModalProps> = ({ show, fileUrl, onClose }) => {
  if (!show || !fileUrl) return null;
 
  const previewFrameRef = useRef<HTMLIFrameElement>(null);
 
  const previewUrl = fileUrl.FilePreviewURL
    ? fileUrl.FilePreviewURL
    : `https://officeindia.sharepoint.com${fileUrl.ServerRelativeUrl}`;
 
  // Function to hide toolbar inside iframe
  const checkAndHideButton = () => {
    let iframeDocument: Document | null = null;
 
    try {
      const iframeEl = previewFrameRef.current;
      if (!iframeEl) return;
 
      iframeDocument = iframeEl.contentDocument || iframeEl.contentWindow?.document;
 
      if (iframeDocument) {
        // Main toolbar
        const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
        const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
 
        if (button) button.style.display = "none";
        if (excelToolbar) excelToolbar.style.display = "none";
 
        // Show iframe only after hiding elements
        iframeEl.style.display = "block";
      }
    } catch (error) {
      console.error("Error accessing iframe content:", error);
      if (previewFrameRef.current) previewFrameRef.current.style.display = "block";
    }
 
    // Continue checking if elements are not yet hidden
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
      <iframe
        ref={previewFrameRef}
        src={`${previewUrl}?web=0&embed=true&toolbar=0`}
        width="100%"
        height="600px"
        style={{ border: "none", display: "none" }} // initially hide iframe
      />
    </div>
  );
};
 
export default PreviewModal;