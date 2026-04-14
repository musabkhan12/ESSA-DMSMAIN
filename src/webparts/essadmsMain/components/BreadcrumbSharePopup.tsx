import * as React from "react";
import { Modal, Button } from "react-bootstrap";
import { useEffect } from "react";
// Aman 14/4/26
import { useState } from "react";
import { Form } from "react-bootstrap";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import Select from "react-select";
import Swal from "sweetalert2";

interface Props {
  show: boolean;
  onClose: () => void;
  url: string;
  context: any;
}

const BreadcrumbSharePopup: React.FC<Props> = ({ show, onClose, url, context }) => {

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

  // Aman 14/4/26
  const [users, setUsers] = useState<any[]>([]);
const [selectedUsers, setSelectedUsers] = useState<any[]>([]);

const userOptions = users.map((u) => ({
  value: u.Email,
  label: `${u.Title} (${u.Email})`
}));

// Aman 14/4/26
useEffect(() => {
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
            (u: any) => u.Email && u.PrincipalType === 1
          );

          allUsers = [...allUsers, ...filteredUsers];
        } catch {}
      }

      const uniqueUsers = Array.from(
        new Map(
          allUsers
            .filter((u) => u.Email)
            .map((u) => [u.Email.toLowerCase(), u])
        ).values()
      );

      setUsers(uniqueUsers);

    } catch {}
  };

  if (show) {
    setSelectedUsers([]);
    fetchUsers();
  }
}, [show, context]);

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(url);
      console.log("Copied:", url);
      onClose();
    } catch (err) {
      console.error("Copy failed", err);
    }
  };

  // Aman 14/4/26

  const handleSendMail = async () => {
  try {
    if (!selectedUsers || selectedUsers.length === 0) {
      Swal.fire({
  icon: "warning",
  title: "No user selected",
  text: "Please select at least one user",
  showConfirmButton: false,
  timer: 1500
});
      return;
    }

    const sp = spfi().using(SPFx(context));

    // current user (Share by)
    const currentUser = await sp.web.currentUser();

    // selected users → IDs
    const userIds: number[] = [];

    for (const user of selectedUsers) {
      const ensuredUser = await sp.web.ensureUser(user.value);
      userIds.push(ensuredUser.data.Id);
    }

    // insert into list
    await sp.web.lists.getByTitle("ShareFolderurl").items.add({
      Url: url,
      SharebyId: currentUser.Id,
      SharetoId: userIds
    });

    Swal.fire({
  icon: "success",
  title: "Mail request sent successfully",
  showConfirmButton: false,
  timer: 1500
}).then(() => {
  onClose();
});

  } catch (error) {
    console.error("Error sending mail:", error);
    alert("Something went wrong");
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
        <Form.Group style={{ marginTop: "10px" }}>
  <Form.Label>Select User</Form.Label>

  <Select
    isMulti
    options={userOptions}
    value={selectedUsers}
    onChange={(selected: any) => setSelectedUsers(selected || [])}
    placeholder="Select users..."
    isSearchable={true}
    closeMenuOnSelect={true}
    hideSelectedOptions={false}
    menuPlacement="auto"
    menuPortalTarget={document.body}
    styles={{
      menuPortal: (base: any) => ({ ...base, zIndex: 9999 })
    }}
  />
</Form.Group>
      </Modal.Body>

      <Modal.Footer>
        <Button onClick={handleCopy}>
          Copy URL
        </Button>
          {/* Aman 14/4/26 */}
         <Button variant="primary" onClick={handleSendMail}>
    Send Mail
  </Button>
      </Modal.Footer>
    </Modal>
  );
};

export default BreadcrumbSharePopup;