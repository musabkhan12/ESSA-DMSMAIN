import * as React from "react";
import { Modal, Button } from "react-bootstrap";
import { useEffect } from "react";

interface Props {
  show: boolean;
  onClose: () => void;
  url: string;
}

const BreadcrumbSharePopup: React.FC<Props> = ({ show, onClose, url }) => {

  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      const modal = document.querySelector(".modal-dialog");

      if (modal && !modal.contains(event.target as Node)) {
        onClose();
      }
    };

    if (show) {
      document.addEventListener("mousedown", handleOutsideClick);
    }

    return () => {
      document.removeEventListener("mousedown", handleOutsideClick);
    };
  }, [show, onClose]);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(url);
      console.log("Copied:", url);
      onClose();
    } catch (err) {
      console.error("Copy failed", err);
    }
  };

  return (
    <Modal
      show={show}
      onHide={onClose}
      centered
      size="sm"
      backdrop={true}
      container={() => document.getElementById('filelistcontainer')}
    >
      <Modal.Header closeButton>
        <Modal.Title>Share URL</Modal.Title>
      </Modal.Header>

      <Modal.Body>
        <div
          style={{
            wordBreak: "break-all",
            fontSize: "13px",
            background: "#f5f5f5",
            padding: "8px",
            borderRadius: "4px"
          }}
        >
          {url}
        </div>
      </Modal.Body>

      <Modal.Footer>
        <Button onClick={handleCopy}>
          Copy URL
        </Button>
      </Modal.Footer>
    </Modal>
  );
};

export default BreadcrumbSharePopup;