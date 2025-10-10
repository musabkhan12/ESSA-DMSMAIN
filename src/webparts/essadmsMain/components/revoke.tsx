
// //revoke access by Om

// import React, { useState, useEffect } from "react";
// import { Modal, Button, Spinner } from "react-bootstrap";
// import { spfi, SPFx } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/site-users/web";
// import "@pnp/sp/security";

// import type { WebPartContext } from "@microsoft/sp-webpart-base";

// interface SharedUser {
//   userId: string;        // principalId as string
//   userTitle: string;     // user display name
//   roles: string[];       // assigned permission roles
// }

// interface RevokeProps {
//   show: boolean;
//   selectedFolder: any;
//   context: WebPartContext;
//   onClose: () => void;
//   onRevoke?: (userId: string) => void; // <-- change this
// }


// // Determine if item is file/folder and construct server relative URL
// const deriveItemContext = (item: any, context: WebPartContext) => {
//   const rootSite = item?.__siteUrl || context.pageContext.web.absoluteUrl;
//   const siteTitle = item?.SiteTitle || item?.SiteName || "";
//   const webUrl = siteTitle ? `${rootSite}/${encodeURIComponent(siteTitle)}` : rootSite;
//   const serverRel = item?.FileName
//     ? `${item.CurrentFolderPath}/${item.FileName}`
//     : item?.FolderPath || item?.ServerRelativeUrl || "";
//   const isFile = !!item.FileName;
//   return { webUrl, serverRel, isFile };
// };

// // Get PnPjs API for file or folder item (return awaited IItem)
// const getItemApi = async (
//   webUrl: string,
//   serverRel: string,
//   context: WebPartContext,
//   isFile: boolean
// ) => {
//   const web = spfi(webUrl).using(SPFx(context)).web;
//   if (isFile) {
//     return await web.getFileByServerRelativePath(serverRel).getItem();
//   } else {
//     return await web.getFolderByServerRelativePath(serverRel).getItem();
//   }
// };

// const Revoke: React.FC<RevokeProps> = ({ show, selectedFolder, context, onClose, onRevoke }) => {
//   const [users, setUsers] = useState<SharedUser[]>([]);
//   const [loading, setLoading] = useState(false);
//   const [error, setError] = useState<string>("");

//   // Fetch permissions dynamically
//   useEffect(() => {
//     if (!show || !selectedFolder) {
//       setUsers([]);
//       return;
//     }

//     const fetchPermissions = async () => {
//       setLoading(true);
//       setError("");
//       try {
//         const { webUrl, serverRel, isFile } = deriveItemContext(selectedFolder, context);

//         // ✅ wait for IItem
//         const itemApi = await getItemApi(webUrl, serverRel, context, isFile);

//         // Expand role assignments
//         const ras: any[] = await itemApi.roleAssignments.expand("Member", "RoleDefinitionBindings")();

//         const mappedUsers: SharedUser[] = ras.map((ra: any) => ({
//           userId: String(ra?.Member?.Id),
//           userTitle: ra?.Member?.Title || "",
//           roles: (ra?.RoleDefinitionBindings || []).map((r: any) => r?.Name).filter(Boolean),
//         }));

//         setUsers(mappedUsers);
//       } catch (e) {
//         setError("Failed to load permissions.");
//         setUsers([]);
//         console.error("Error loading role assignments:", e);
//       } finally {
//         setLoading(false);
//       }
//     };

//     fetchPermissions();
//   }, [show, selectedFolder, context]);

//   // Revoke access from SharePoint
//   const revokeAccess = async (userId: string) => {
//     if (!window.confirm("Are you sure you want to revoke access for this user?")) return;

//     setLoading(true);
//     try {
//       const { webUrl, serverRel, isFile } = deriveItemContext(selectedFolder, context);

//       // ✅ wait for IItem
//       const itemApi = await getItemApi(webUrl, serverRel, context, isFile);

//       // Remove role assignment
//       await itemApi.roleAssignments.getById(parseInt(userId)).delete();

//       // Update local state
//       setUsers((prev) => prev.filter((u) => u.userId !== userId));
// if (onRevoke) onRevoke(userId);  // optional callback
//     } catch (e) {
//       console.error("Failed to revoke access:", e);
//       alert("Failed to revoke access. Check console for details.");
//     } finally {
//       setLoading(false);
//     }
//   };

//   return (
//     <Modal show={show} onHide={onClose} centered size="lg">
//       <Modal.Header closeButton>
//         <Modal.Title>Shared Users & Permissions</Modal.Title>
//       </Modal.Header>
//       <Modal.Body>
//         {loading ? (
//           <div>
//             <Spinner animation="border" size="sm" /> Loading...
//           </div>
//         ) : error ? (
//           <p style={{ color: "red" }}>{error}</p>
//         ) : users.length > 0 ? (
//           <table style={{ width: "100%", borderCollapse: "collapse" }}>
//             <thead>
//               <tr style={{ borderBottom: "1px solid #ccc" }}>
//                 <th style={{ textAlign: "left", padding: "8px" }}>User</th>
//                 <th style={{ textAlign: "left", padding: "8px" }}>Permissions</th>
//                 <th style={{ textAlign: "center", padding: "8px" }}>Action</th>
//               </tr>
//             </thead>
//             <tbody>
//               {users.map((user) => (
//                 <tr key={user.userId} style={{ borderBottom: "1px solid #eee" }}>
//                   <td style={{ padding: "8px" }}>{user.userTitle || user.userId}</td>
//                   <td style={{ padding: "8px" }}>{user.roles.join(", ")}</td>
//                   <td style={{ padding: "8px", textAlign: "center" }}>
//                     <Button variant="danger" size="sm" onClick={() => revokeAccess(user.userId)}>
//                       Revoke
//                     </Button>
//                   </td>
//                 </tr>
//               ))}
//             </tbody>
//           </table>
//         ) : (
//           <p>No users have access.</p>
//         )}
//       </Modal.Body>
//       <Modal.Footer>
//         <Button variant="secondary" onClick={onClose}>
//           Close
//         </Button>
//       </Modal.Footer>
//     </Modal>
//   );
// };

// export default Revoke;









//revoke access by Om

import React, { useState, useEffect } from "react";
import { Modal, Button, Spinner } from "react-bootstrap";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/security";

import type { WebPartContext } from "@microsoft/sp-webpart-base";

interface SharedUser {
  userId: string;        // principalId as string
  userTitle: string;     // user display name
  roles: string[];       // assigned permission roles
}

interface RevokeProps {
  show: boolean;
  selectedFolder: any;
  context: WebPartContext;
  onClose: () => void;
  onRevoke?: (userId: string) => void; // <-- change this
}


// Determine if item is file/folder and construct server relative URL
const deriveItemContext = (item: any, context: WebPartContext) => {
  const rootSite = item?.__siteUrl || context.pageContext.web.absoluteUrl;
  const siteTitle = item?.SiteTitle || item?.SiteName || "";
  const webUrl = siteTitle ? `${rootSite}/${encodeURIComponent(siteTitle)}` : rootSite;
  const serverRel = item?.FileName
    ? `${item.CurrentFolderPath}/${item.FileName}`
    : item?.FolderPath || item?.ServerRelativeUrl || "";
  const isFile = !!item.FileName;
  return { webUrl, serverRel, isFile };
};

// Get PnPjs API for file or folder item (return awaited IItem)
const getItemApi = async (
  webUrl: string,
  serverRel: string,
  context: WebPartContext,
  isFile: boolean
) => {
  const web = spfi(webUrl).using(SPFx(context)).web;
  if (isFile) {
    return await web.getFileByServerRelativePath(serverRel).getItem();
  } else {
    return await web.getFolderByServerRelativePath(serverRel).getItem();
  }
};

const Revoke: React.FC<RevokeProps> = ({ show, selectedFolder, context, onClose, onRevoke }) => {
  const [users, setUsers] = useState<SharedUser[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>("");

  // Fetch permissions dynamically
  useEffect(() => {
    if (!show || !selectedFolder) {
      setUsers([]);
      return;
    }

    const fetchPermissions = async () => {
      setLoading(true);
      setError("");
      try {
        const { webUrl, serverRel, isFile } = deriveItemContext(selectedFolder, context);

        // ✅ wait for IItem
        const itemApi = await getItemApi(webUrl, serverRel, context, isFile);

        // Expand role assignments
        const ras: any[] = await itemApi.roleAssignments.expand("Member", "RoleDefinitionBindings")();

        const mappedUsers: SharedUser[] = ras.map((ra: any) => ({
          userId: String(ra?.Member?.Id),
          userTitle: ra?.Member?.Title || "",
          roles: (ra?.RoleDefinitionBindings || []).map((r: any) => r?.Name).filter(Boolean),
        }));

        setUsers(mappedUsers);
      } catch (e) {
        setError("Failed to load permissions.");
        setUsers([]);
        console.error("Error loading role assignments:", e);
      } finally {
        setLoading(false);
      }
    };

    fetchPermissions();
  }, [show, selectedFolder, context]);

  // Revoke access from SharePoint
  const revokeAccess = async (userId: string) => {
    if (!window.confirm("Are you sure you want to revoke access for this user?")) return;

    setLoading(true);
    try {
      const { webUrl, serverRel, isFile } = deriveItemContext(selectedFolder, context);

      // ✅ wait for IItem
      const itemApi = await getItemApi(webUrl, serverRel, context, isFile);

      // 🔑 ensure unique permissions (break inheritance if needed)
      await itemApi.breakRoleInheritance(true, false);

      // Remove role assignment
      await itemApi.roleAssignments.getById(parseInt(userId)).delete();

      // Update local state
      setUsers((prev) => prev.filter((u) => u.userId !== userId));
      if (onRevoke) onRevoke(userId);  // optional callback
    } catch (e) {
      console.error("Failed to revoke access:", e);
      alert("Failed to revoke access. Check console for details.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <Modal show={show} onHide={onClose} centered size="lg">
      <Modal.Header closeButton>
        <Modal.Title>Shared Users & Permissions</Modal.Title>
      </Modal.Header>
      <Modal.Body>
        {loading ? (
          <div>
            <Spinner animation="border" size="sm" /> Loading...
          </div>
        ) : error ? (
          <p style={{ color: "red" }}>{error}</p>
        ) : users.length > 0 ? (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ borderBottom: "1px solid #ccc" }}>
                <th style={{ textAlign: "left", padding: "8px" }}>User</th>
                <th style={{ textAlign: "left", padding: "8px" }}>Permissions</th>
                <th style={{ textAlign: "center", padding: "8px" }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {users.map((user) => (
                <tr key={user.userId} style={{ borderBottom: "1px solid #eee" }}>
                  <td style={{ padding: "8px" }}>{user.userTitle || user.userId}</td>
                  <td style={{ padding: "8px" }}>{user.roles.join(", ")}</td>
                  <td style={{ padding: "8px", textAlign: "center" }}>
                    <Button variant="danger" size="sm" onClick={() => revokeAccess(user.userId)}>
                      Revoke
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p>No users have access.</p>
        )}
      </Modal.Body>
      <Modal.Footer>
        <Button variant="secondary" onClick={onClose}>
          Close
        </Button>
      </Modal.Footer>
    </Modal>
  );
};

export default Revoke;
