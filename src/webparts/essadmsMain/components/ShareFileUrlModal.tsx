import * as React from "react";
import { Modal, Button } from 'react-bootstrap';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Form } from "react-bootstrap";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import Swal from "sweetalert2";
import Select from "react-select";

interface ShareFileUrlModalProps {
  show: boolean;
  onClose: () => void;
  file: any;
  currentSiteUrl: string;
  context: WebPartContext;
}

type UserOption = {
  value: string;
  label: string;
};

const ShareFileUrlModal: React.FC<ShareFileUrlModalProps> = ({ show, onClose, file, currentSiteUrl, context }) => {
  const [shareUrl, setShareUrl] = React.useState<string>("");
  const [copied, setCopied] = React.useState<boolean>(false);
 
 
  const [users, setUsers] = React.useState<any[]>([]);
const [selectedUsers, setSelectedUsers] = React.useState<UserOption[]>([]);
 
 
 
 
 
const userOptions = users.map((u) => ({
  value: u.Email,
  label: `${u.Title} (${u.Email})`
}));
 
 
 
 
React.useEffect(() => {
  const fetchUsers = async () => {
    try {
      if (!context) return;
 
      const sp = spfi().using(SPFx(context));
 
      const sites = await sp.web.lists
        .getByTitle("MasterSiteCollection")
        .items.select("Title", "SiteURL")();
        sites.push({
  Title: "ESSA",
  SiteURL: context.pageContext.web.absoluteUrl
});
      let allUsers: any[] = [];
 
      for (const site of sites) {
        if (!site.SiteURL) continue;
 
        try {
          const siteSp = spfi(site.SiteURL).using(SPFx(context));
          const siteUsers = await siteSp.web.siteUsers();
 
          const filteredUsers = siteUsers.filter(
  (u: any) =>
    u.Email &&
    u.PrincipalType === 1
);
 
allUsers = [...allUsers, ...filteredUsers];
        } catch (err) {
          console.error("User fetch error:", site.SiteURL);
        }
      }
 
      const uniqueUsers = Array.from(
  new Map(
    allUsers
      .filter((u) => u.Email)
      .map((u) => [u.Email.toLowerCase(), u])
  ).values()
);
 
      setUsers(uniqueUsers);
 
    } catch (error) {
      console.error("User fetch failed:", error);
    }
  };
 
  if (show) {
    setSelectedUsers([]);
    fetchUsers();
  }
}, [show, context]);
 
 
 
// aman tak ka code

 
  // ─────────────────────────────────────────────
  const buildDeepLinkUrl = (file: any): string => {
    console.log("Building deep link with file:", file);
    if (!file) return "";

    try {
      // Base: current page URL without any previous essaPreview params
    //   const base = window.location.href.split("?")[0];
      const base = "https://officeindia.sharepoint.com/sites/ESSA/SitePages/ESSADMS.aspx";
      console.log("Base URL for deep link:", base);
      const params = new URLSearchParams();
      console.log("Initial URLSearchParams:", params.toString());
      params.set("essaPreview", "1");

      // ServerRelativeUrl — needed by PreviewModal to build the iframe src
      const serverRelativeUrl =
        file.ServerRelativeUrl ||
        (file.CurrentFolderPath && file.FileName
          ? `${file.CurrentFolderPath.replace(/\/$/, "")}/${file.FileName}`
          : "");

      if (serverRelativeUrl) {
        params.set("essaFileUrl", serverRelativeUrl);
      }

      // File name
      const fileName = file.FileName || file.Name || file.DisplayName || "";
      if (fileName) params.set("essaFileName", fileName);

      // // Site URL (for subsite-aware PnP calls in PreviewModal)
      // const siteUrl = file.__siteUrl || currentSiteUrl || "";
      // if (siteUrl) params.set("essaSiteUrl", siteUrl);

      // // Optional: pass the full FilePreviewURL so PreviewModal can use it directly
      // const previewUrl = typeof file.FilePreviewURL === "string" ? file.FilePreviewURL : "";
      // if (previewUrl) params.set("essaFilePreviewURL", previewUrl);

      // // Optional: SiteName (used in Doc.aspx path construction in previewfile.tsx)
      // if (file.SiteName) params.set("essaSiteName", file.SiteName);

      // // Optional: CurrentFolderPath (used as fallback in PreviewModal)
      // if (file.CurrentFolderPath) params.set("essaFolderPath", file.CurrentFolderPath);

      return `${base}?${params.toString()}`;
    } catch (err) {
      console.error("buildDeepLinkUrl failed:", err);
      return "";
    }
 };

  const handleSendMail = async () => {
  try {
    if (!selectedUsers || selectedUsers.length === 0) {
      alert("Please select at least one user");
      return;
    }
 
    const sp = spfi().using(SPFx(context));
 
    // current user (Share By)
    const currentUser = await sp.web.currentUser();
 
    //selected users  get IDs
    const userIds: number[] = [];
 
    for (const user of selectedUsers) {
      const ensuredUser = await sp.web.ensureUser(user.value);
      userIds.push(ensuredUser.data.Id);
    }
 
    // create item in list
    const payload = {
      
      Url: shareUrl,
      SharebyId: currentUser.Id,
      SharetoId: { results: userIds }
    };
    console.log("ShareFileUrl payload:", payload);

    try {
      await sp.web.lists.getByTitle("shareFileUrlData").items.add(payload);
    } catch (innerError: any) {
      const message = innerError?.message || "";
      console.warn("ShareFileUrl add failed with payload object, retrying with raw array:", message);

      if (message.includes("StartObject") || message.includes("StartArray")) {
        const fallbackPayload = {
          ...payload,
          SharetoId: userIds
        };
        console.log("Retry payload for ShareFileUrl:", fallbackPayload);
        await sp.web.lists.getByTitle("shareFileUrlData").items.add(fallbackPayload);
      } else {
        throw innerError;
      }
    }

    Swal.fire({
      icon: "success",
      title: "Mail request sent successfully",
      showConfirmButton: false,
      timer: 1500
    });
    onClose();
 
  } catch (error) {
    console.error("Error sending mail:", error);
    alert("Something went wrong");
  }
};

  React.useEffect(() => {
    if (!show || !file) {
      setShareUrl("");
      setCopied(false);
      return;
    }

    setShareUrl(buildDeepLinkUrl(file));
    setCopied(false);
  }, [show, file]);

  const handleCopy = async () => {
    if (!shareUrl) return;
    try {
      await navigator.clipboard.writeText(shareUrl);
      setCopied(true);
    } catch (error) {
      console.error("Copy failed", error);
      setCopied(false);
    }
  };

  const fileName = file?.FileName || file?.Name || file?.DisplayName || "File";

  return (
    <Modal
      show={show}
      onHide={onClose}
      centered
      className='sharemodal'
      size="lg"
      container={() => document.getElementById('filelistcontainer')}
      backdrop={true}
    >
      <Modal.Header closeButton>
        <Modal.Title>Share File URL</Modal.Title>
      </Modal.Header>
      <Modal.Body>
        <div style={{ display: 'flex', flexDirection: 'column', gap: '18px' }}>
          <div>
            <div style={{ fontSize: '14px', fontWeight: 600, marginBottom: '8px' }}>File</div>
            <div style={{ fontSize: '16px', fontWeight: 600 }}>{fileName}</div>
          </div>
 
          <Form.Group style={{ marginTop: "10px" }}>
  <Form.Label>Select User</Form.Label>
 
  <Select
    isMulti
    options={userOptions}
    value={selectedUsers}
    onChange={(selected: any) => setSelectedUsers(Array.isArray(selected) ? selected : [])}
    placeholder="Select users..."
    isSearchable={true}
  />
</Form.Group>

          <div>
            <div style={{ fontSize: '14px', fontWeight: 600, marginBottom: '8px' }}>Shareable link</div>
            <div style={{ display: 'flex', gap: '10px', alignItems: 'stretch', flexWrap: 'wrap' }}>
              <input
                type="text"
                readOnly
                value={shareUrl}
                placeholder="No URL available for this file"
                style={{
                  flex: 1,
                  minWidth: '220px',
                  padding: '10px 12px',
                  borderRadius: '6px',
                  border: '1px solid #ccc',
                  backgroundColor: '#f7f9fc',
                  fontSize: '14px',
                }}
              />
              <Button
                variant="primary"
                onClick={handleCopy}
                disabled={!shareUrl}
                style={{ minWidth: '140px' }}
              >
                {copied ? '✅ Copied!' : 'Copy link'}
              </Button>
            </div>
            {!shareUrl && (
              <div style={{ marginTop: '12px', color: '#666', fontSize: '13px' }}>
                This file does not have an available shareable link yet.
              </div>
            )}
          </div>

          <div style={{ color: '#444', fontSize: '13px', lineHeight: '1.6' }}>
            🔗 When the recipient opens this link, the file will automatically open in the preview panel.
            If the file is stored in a private library, they may still need permission to access it.
          </div>
        </div>
      </Modal.Body>
      <Modal.Footer>
        <Button variant="secondary" onClick={onClose}>Close</Button>
        <Button variant="primary" onClick={handleSendMail}>Send Mail</Button>
      </Modal.Footer>
    </Modal>
  );
};

export default ShareFileUrlModal;
