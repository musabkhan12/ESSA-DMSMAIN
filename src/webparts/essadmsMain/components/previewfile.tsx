import * as React from "react";

interface PreviewModalProps {
  show: boolean;
  fileUrl?: any;
  onClose: () => void;
}

const PreviewModal: React.FC<PreviewModalProps> = ({ show, fileUrl, onClose }) => {
  if (!show) return null;

  return (
    <div className="modal-overlay">
      <div className="modal-content">
        <button onClick={onClose}>Close</button>

    
          <iframe src={fileUrl.FilePreviewURL} width="100%" height="600px" />
      
      </div>
    </div>
  );
};

export default PreviewModal;
