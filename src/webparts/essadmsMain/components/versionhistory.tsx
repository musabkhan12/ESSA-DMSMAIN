import * as React from "react";
<<<<<<< HEAD
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/files";
import "@pnp/sp/webs";


=======
import { Modal, Button, Spinner } from "react-bootstrap";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/files";
import "@pnp/sp/webs";
>>>>>>> ESSAMUSAIBUPDATED2
import "@pnp/sp/site-users/web";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

interface VersionHistoryModalProps {
  show: boolean;
  onClose: () => void;
  file: any;
  context: WebPartContext;
}

const VersionHistoryModal: React.FC<VersionHistoryModalProps> = ({
  show,
  onClose,
  file,
  context,
}) => {
  const [versions, setVersions] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
<<<<<<< HEAD
   console.log("File for version history:", file);
  React.useEffect(() => {
    let canceled = false;
  //  this was previous workign code 
//     if (show && file) {
//       (async () => {
//         setLoading(true);
//         setError(null);

//         // try {
//         //   const folderPath = file.CurrentFolderPath || "";
//         //   const fileName = file.FileName;
//         //   const serverRelativePath =
//         //     folderPath.endsWith("/")
//         //       ? `${folderPath}${fileName}`
//         //       : `${folderPath}/${fileName}`;

//         //   const parts = folderPath.split("/").filter(Boolean);
//         //   const subsitePath = "/" + parts.slice(0, 3).join("/");
//         //   const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
//         //   const subsiteUrl = `${tenantUrl}${subsitePath}`;

//         //   const siteSP = spfi(subsiteUrl).using(SPFx(context));
//         //   const fileItem = siteSP.web.getFileByServerRelativePath(serverRelativePath);

//         //   const fileProps: any = await fileItem
//         //     .select(
//         //       "TimeLastModified",
//         //       "Length",
//         //       "Name",
//         //       "UIVersionLabel",
//         //       "ModifiedBy/Title",
//         //       "ModifiedBy/UserPrincipalName"
//         //     )
//         //     .expand("ModifiedBy")();

//         //   const versionData = await fileItem.versions
//         //     .select("ID,VersionLabel,Created,Size, Url,CreatedBy/Title,CreatedBy/UserPrincipalName")
//         //     .expand("CreatedBy")();
//         //   console.log("Version data:", versionData);
//         //   const allVersions = [
//         //     ...versionData.map((v) => ({
//         //       ID: v.ID, // store version ID for download
//         //       VersionLabel: v.VersionLabel,
//         //       Created: new Date(v.Created),
//         //       Size: v.Size,
//         //       CreatedBy: v.CreatedBy,
//         //     })),
//         //     {
//         //       ID: "current", // special ID for latest version
//         //       VersionLabel: fileProps.UIVersionLabel,
//         //       Created: new Date(fileProps.TimeLastModified),
//         //       Size: fileProps.Length,
//         //       CreatedBy: fileProps.ModifiedBy,
//         //     },
//         //   ];

//         //   allVersions.sort((a, b) => a.Created.getTime() - b.Created.getTime());
//         //    console.log("Version history loaded:", allVersions);
//         //   if (!canceled) setVersions(allVersions);
//         // } catch (err: any) {
//         //   if (!canceled) {
//         //     setVersions([]);
//         //     setError("Could not load version history.");
//         //     console.error("Error loading version history:", err);
//         //   }
//         // } finally {
//         //   if (!canceled) setLoading(false);
//         // }
//       try {
//   const folderPath = file.CurrentFolderPath || "";
//   const fileName = file.FileName;
//   const serverRelativePath = folderPath.endsWith("/")
//     ? `${folderPath}${fileName}`
//     : `${folderPath}/${fileName}`;

//   // Build full site URL from context
//   const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
//   const parts = folderPath.split("/").filter(Boolean);
//   const subsitePath = "/" + parts.slice(0, 3).join("/");
//   const subsiteUrl = `${tenantUrl}${subsitePath}`;

//   const siteSP = spfi(subsiteUrl).using(SPFx(context));
//   const fileItem = siteSP.web.getFileByServerRelativePath(serverRelativePath);

//   // Get latest file properties
//   const fileProps: any = await fileItem
//     .select(
//       "TimeLastModified",
//       "Length",
//       "Name",
//       "ServerRelativeUrl",
//       "UIVersionLabel",
//       "ModifiedBy/Title",
//       "ModifiedBy/UserPrincipalName"
//     )
//     .expand("ModifiedBy")();

//   // Get version history
//   const versionData = await fileItem.versions
//     .select("ID,VersionLabel,Created,Size,Url,CreatedBy/Title,CreatedBy/UserPrincipalName")
//     .expand("CreatedBy")();

//   // Base site URL for constructing download links
//   const webUrl = context.pageContext.web.absoluteUrl;

//   const allVersions = [
//     ...versionData.map((v: any) => ({
//       ID: v.ID,
//       VersionLabel: v.VersionLabel,
//       Created: new Date(v.Created),
//       Size: v.Size,
//       CreatedBy: v.CreatedBy,
//       // Build direct download link for this version
//       DownloadUrl: `${webUrl}/${v.Url}`,
//     })),
//     {
//       ID: "current",
//       VersionLabel: fileProps.UIVersionLabel,
//       Created: new Date(fileProps.TimeLastModified),
//       Size: fileProps.Length,
//       CreatedBy: fileProps.ModifiedBy,
//       // Current version download link
//       DownloadUrl: `${webUrl}${fileProps.ServerRelativeUrl}`,
//     },
//   ];

//   // Sort by date
//   allVersions.sort((a, b) => a.Created.getTime() - b.Created.getTime());

//   console.log("Version history loaded:", allVersions);

//   if (!canceled) setVersions(allVersions);

// } catch (err: any) {
//   if (!canceled) {
//     setVersions([]);
//     setError("Could not load version history.");
//     console.error("Error loading version history:", err);
//   }
// } finally {
//   if (!canceled) setLoading(false);
// }

//       })();
//     }
if (show && file) {
=======

  console.log("File for version history:", file);

  React.useEffect(() => {
    let canceled = false;

    if (show && file) {
>>>>>>> ESSAMUSAIBUPDATED2
      (async () => {
        setLoading(true);
        setError(null);
        try {
<<<<<<< HEAD
          const folderPath = file.CurrentFolderPath || file.ServerRelativeUrl || "";  //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
          const fileName = file.FileName || file.Name;
          //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
=======
          const folderPath =
            file.CurrentFolderPath || file.ServerRelativeUrl || "";
          const fileName = file.FileName || file.Name;

>>>>>>> ESSAMUSAIBUPDATED2
          let serverRelativePath = "";
          if (file.ServerRelativeUrl) {
            serverRelativePath = file.ServerRelativeUrl;
          } else {
            serverRelativePath = folderPath.endsWith("/")
              ? `${folderPath}${fileName}`
              : `${folderPath}/${fileName}`;
          }
<<<<<<< HEAD
 
 
          // Build full site URL from context
          const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
          const parts = folderPath.split("/").filter(Boolean);
          const subsitePath = "/" + parts.slice(0, 3).join("/");
          const subsiteUrl = `${tenantUrl}${subsitePath}`;
 
          const siteSP = spfi(subsiteUrl).using(SPFx(context));
          const fileItem = siteSP.web.getFileByServerRelativePath(serverRelativePath);
 
          // Get latest file properties
=======

          const tenantUrl =
            context.pageContext.web.absoluteUrl.split("/sites/")[0];
          const parts = folderPath.split("/").filter(Boolean);
          const subsitePath = "/" + parts.slice(0, 3).join("/");
          const subsiteUrl = `${tenantUrl}${subsitePath}`;

          const siteSP = spfi(subsiteUrl).using(SPFx(context));
          const fileItem =
            siteSP.web.getFileByServerRelativePath(serverRelativePath);

>>>>>>> ESSAMUSAIBUPDATED2
          const fileProps: any = await fileItem
            .select(
              "TimeLastModified",
              "Length",
              "Name",
              "ServerRelativeUrl",
              "UIVersionLabel",
              "ModifiedBy/Title",
              "ModifiedBy/UserPrincipalName"
            )
            .expand("ModifiedBy")();
<<<<<<< HEAD
 
          // Get version history
          const versionData = await fileItem.versions
            .select("ID,VersionLabel,Created,Size,Url,CreatedBy/Title,CreatedBy/UserPrincipalName")
            .expand("CreatedBy")();
 
          // Base site URL for constructing download links
          const webUrl = context.pageContext.web.absoluteUrl;
 
=======

          const versionData = await fileItem.versions
            .select(
              "ID,VersionLabel,Created,Size,Url,CreatedBy/Title,CreatedBy/UserPrincipalName"
            )
            .expand("CreatedBy")();

          const webUrl = context.pageContext.web.absoluteUrl;

>>>>>>> ESSAMUSAIBUPDATED2
          const allVersions = [
            ...versionData.map((v: any) => ({
              ID: v.ID,
              VersionLabel: v.VersionLabel,
              Created: new Date(v.Created),
              Size: v.Size,
              CreatedBy: v.CreatedBy,
<<<<<<< HEAD
              // Build direct download link for this version
=======
>>>>>>> ESSAMUSAIBUPDATED2
              DownloadUrl: `${webUrl}/${v.Url}`,
            })),
            {
              ID: "current",
              VersionLabel: fileProps.UIVersionLabel,
              Created: new Date(fileProps.TimeLastModified),
              Size: fileProps.Length,
              CreatedBy: fileProps.ModifiedBy,
<<<<<<< HEAD
              // Current version download link
              DownloadUrl: `${webUrl}${fileProps.ServerRelativeUrl}`,
            },
          ];
 
          // Sort by date
          allVersions.sort((a, b) => a.Created.getTime() - b.Created.getTime());
 
          console.log("Version history loaded:", allVersions);
 
          if (!canceled) setVersions(allVersions);
 
=======
              DownloadUrl: `${webUrl}${fileProps.ServerRelativeUrl}`,
            },
          ];

          allVersions.sort(
            (a, b) => a.Created.getTime() - b.Created.getTime()
          );

          if (!canceled) setVersions(allVersions);
>>>>>>> ESSAMUSAIBUPDATED2
        } catch (err: any) {
          if (!canceled) {
            setVersions([]);
            setError("Could not load version history.");
            console.error("Error loading version history:", err);
          }
        } finally {
          if (!canceled) setLoading(false);
        }
<<<<<<< HEAD
 
      })();
    }
=======
      })();
    }

>>>>>>> ESSAMUSAIBUPDATED2
    return () => {
      canceled = true;
    };
  }, [show, file, context]);

<<<<<<< HEAD
  if (!show) return null;

  const formatDate = (date: Date) => date.toLocaleString("en-US");

  //  Function to download version
// inside your component

// this was woprking
// const downloadVersion = async (file: any, version: any) => {
//   try {
//     const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
//     const parts = file.CurrentFolderPath.split("/").filter(Boolean);
//     const subsitePath = "/" + parts.slice(0, 3).join("/");
//     const subsiteUrl = `${tenantUrl}${subsitePath}`;

//     const siteSP = spfi(subsiteUrl).using(SPFx(context));
//     const fileItem = siteSP.web.getFileByServerRelativePath(
//       `${file.CurrentFolderPath}/${file.FileName}`
//     );

//     let blob: Blob;

//     if (version.ID === "current") {
//       // current version → just get the file itself
//       blob = await fileItem.getBlob();
//     } else {
//       // historical version → use version id
//       const versionItem = fileItem.versions.getById(version.ID);
//       blob = await versionItem.getBlob();
//     }

//     if (!(blob instanceof Blob)) {
//       throw new Error("Did not receive a Blob, check permissions or file path.");
//     }

//     const url = URL.createObjectURL(blob);
//     const link = document.createElement("a");
//     link.href = url;
//     link.download = file.FileName; // or `${file.FileName} (v${version.VersionLabel})`
//     document.body.appendChild(link);
//     link.click();
//     link.remove();
//     URL.revokeObjectURL(url);

//     console.log(`Downloaded version ${version.VersionLabel}`);
//   } catch (error) {
//     console.error("Download failed:", error);
//   }
// };


 const downloadVersion = async (file: any, version: any) => {
    try {
      const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
 
      const folderPath = file.CurrentFolderPath || file.ServerRelativeUrl || "";
      const fileName = file.FileName || file.Name;
      //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
=======
  const formatDate = (date: Date) => date.toLocaleString("en-US");

  const downloadVersion = async (file: any, version: any) => {
    try {
      const tenantUrl =
        context.pageContext.web.absoluteUrl.split("/sites/")[0];

      const folderPath =
        file.CurrentFolderPath || file.ServerRelativeUrl || "";
      const fileName = file.FileName || file.Name;

>>>>>>> ESSAMUSAIBUPDATED2
      let parts;
      let serverRelativePath = "";
      if (file.ServerRelativeUrl) {
        parts = file.ServerRelativeUrl.split("/").filter(Boolean);
        serverRelativePath = file.ServerRelativeUrl;
      } else {
        parts = file.CurrentFolderPath.split("/").filter(Boolean);
        serverRelativePath = `${file.CurrentFolderPath}/${file.FileName}`;
      }
<<<<<<< HEAD
 
      //  const parts = file.CurrentFolderPath.split("/").filter(Boolean);
      const subsitePath = "/" + parts.slice(0, 3).join("/");
      const subsiteUrl = `${tenantUrl}${subsitePath}`;
 
      const siteSP = spfi(subsiteUrl).using(SPFx(context));
      // const fileItem = siteSP.web.getFileByServerRelativePath(
      //   `${file.CurrentFolderPath}/${file.FileName}`
      // );
      //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
      const fileItem = siteSP.web.getFileByServerRelativePath(serverRelativePath);
 
      let blob: Blob;
 
      if (version.ID === "current") {
        // current version → just get the file itself
        blob = await fileItem.getBlob();
      } else {
        // historical version → use version id
        const versionItem = fileItem.versions.getById(version.ID);
        blob = await versionItem.getBlob();
      }
 
      if (!(blob instanceof Blob)) {
        throw new Error("Did not receive a Blob, check permissions or file path.");
      }
 
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      //sourish 25/8/25 file.Name to handle vesrion from nodes
      link.download = file.FileName || file.Name; // or `${file.FileName} (v${version.VersionLabel})`
=======

      const subsitePath = "/" + parts.slice(0, 3).join("/");
      const subsiteUrl = `${tenantUrl}${subsitePath}`;

      const siteSP = spfi(subsiteUrl).using(SPFx(context));
      const fileItem =
        siteSP.web.getFileByServerRelativePath(serverRelativePath);

      let blob: Blob;

      if (version.ID === "current") {
        blob = await fileItem.getBlob();
      } else {
        const versionItem = fileItem.versions.getById(version.ID);
        blob = await versionItem.getBlob();
      }

      if (!(blob instanceof Blob)) {
        throw new Error(
          "Did not receive a Blob, check permissions or file path."
        );
      }

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = file.FileName || file.Name;
>>>>>>> ESSAMUSAIBUPDATED2
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
<<<<<<< HEAD
 
=======

>>>>>>> ESSAMUSAIBUPDATED2
      console.log(`Downloaded version ${version.VersionLabel}`);
    } catch (error) {
      console.error("Download failed:", error);
    }
  };
<<<<<<< HEAD
  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-labelledby="version-history-title"
      tabIndex={-1}
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        width: "100vw",
        height: "100vh",
        backgroundColor: "rgba(0,0,0,0.5)",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        zIndex: 9999,
      }}
    >
      <div
        style={{
          background: "white",
          padding: "20px",
          borderRadius: "8px",
          width: "750px",
          maxHeight: "80vh",
          overflowY: "auto",
          boxShadow: "0 4px 8px rgba(0,0,0,0.2)",
        }}
      >
        <h2 id="version-history-title" style={{ marginBottom: "15px" }}>
          Version History - {file.FileName || file.Name || "Unknown File"}
        </h2>

        {loading ? (
          <div>Loading versions...</div>
=======

  return (
    <Modal
      show={show}
      onHide={onClose}
      container={() => document.getElementById('filelistcontainer')}
      backdrop="static"
      size="lg"
      centered
    >
      <Modal.Header closeButton>
        <Modal.Title>
          Version History - {file?.FileName || file?.Name || "Unknown File"}
        </Modal.Title>
      </Modal.Header>
    
      <Modal.Body>
        {loading ? (
          <div>
            <Spinner animation="border" size="sm" /> Loading versions...
          </div>
>>>>>>> ESSAMUSAIBUPDATED2
        ) : error ? (
          <div style={{ color: "red", marginBottom: "10px" }}>{error}</div>
        ) : versions.length > 0 ? (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #ddd" }}>
<<<<<<< HEAD
                <th style={{ padding: "8px" }}>Version</th>
                <th style={{ padding: "8px" }}>Modified</th>
                <th style={{ padding: "8px" }}>Modified By</th>
                <th style={{ padding: "8px" }}>Size</th>
                <th style={{ padding: "8px" }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {versions.map(v => (
                <tr key={v.ID}>
                  <td style={{ padding: "8px" }}>{v.VersionLabel}</td>
                  <td style={{ padding: "8px" }}>{formatDate(v.Created)}</td>
=======
                <th style={{ padding: "8px",background:'rgb(245, 247, 250)',color:'#000' }}>Version</th>
                <th style={{ padding: "8px",background:'rgb(245, 247, 250)',color:'#000' }}>Modified</th>
                <th style={{ padding: "8px",background:'rgb(245, 247, 250)',color:'#000' }}>Modified By</th>
                <th style={{ padding: "8px",background:'rgb(245, 247, 250)',color:'#000' }}>Size</th>
                <th style={{ padding: "8px",background:'rgb(245, 247, 250)',color:'#000' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {versions.map((v) => (
                <tr key={v.ID}>
                  <td style={{ padding: "8px" }}>{v.VersionLabel}</td>
                  <td style={{ padding: "8px" }}>
                    {formatDate(v.Created)}
                  </td>
>>>>>>> ESSAMUSAIBUPDATED2
                  <td style={{ padding: "8px" }}>
                    {v.CreatedBy?.Title || "-"}
                    <br />
                    <small>{v.CreatedBy?.UserPrincipalName}</small>
                  </td>
                  <td style={{ padding: "8px" }}>
<<<<<<< HEAD
                    {v.Size > 0 ? `${(v.Size / 1024).toFixed(2)} KB` : "-"}
                  </td>
                  <td style={{ padding: "8px" }}>
                    <button
                      onClick={() => downloadVersion(file, v)}
                      style={{
                        padding: "5px 10px",
                        backgroundColor: "#0078d4",
                        color: "white",
                        border: "none",
                        borderRadius: "4px",
                        cursor: "pointer",
                        fontSize: "13px",
                      }}
                    >
                      Download
                    </button>
=======
                    {v.Size > 0
                      ? `${(v.Size / 1024).toFixed(2)} KB`
                      : "-"}
                  </td>
                  <td style={{ padding: "8px" }}>
                    <Button
                      size="sm" className="mt-0"
                      onClick={() => downloadVersion(file, v)}
                    >
                      Download
                    </Button>
>>>>>>> ESSAMUSAIBUPDATED2
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div>No version history found.</div>
        )}
<<<<<<< HEAD

        <div style={{ marginTop: "20px", textAlign: "right" }}>
          <button
            onClick={onClose}
            style={{
              padding: "10px 18px",
              backgroundColor: "#0078d4",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              fontWeight: "600",
            }}
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

export default VersionHistoryModal;
=======
      </Modal.Body>

      {/* <Modal.Footer>
        <Button variant="secondary" onClick={onClose}>
          Close
        </Button>
      </Modal.Footer> */}
    </Modal>
  );
};

export default VersionHistoryModal;
>>>>>>> ESSAMUSAIBUPDATED2
