import * as React from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/files";
import "@pnp/sp/webs";
 
 
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
   console.log("File for version history:", file);
  React.useEffect(() => {
    let canceled = false;
if (show && file) {
      (async () => {
        setLoading(true);
        setError(null);
        try {
          const folderPath = file.CurrentFolderPath || file.ServerRelativeUrl || "";  //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
          const fileName = file.FileName || file.Name;
          //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
          let serverRelativePath = "";
          if (file.ServerRelativeUrl) {
            serverRelativePath = file.ServerRelativeUrl;
          } else {
            serverRelativePath = folderPath.endsWith("/")
              ? `${folderPath}${fileName}`
              : `${folderPath}/${fileName}`;
          }
 
 
          // Build full site URL from context
          const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
          const parts = folderPath.split("/").filter(Boolean);
          const subsitePath = "/" + parts.slice(0, 3).join("/");
          const subsiteUrl = `${tenantUrl}${subsitePath}`;
 
          const siteSP = spfi(subsiteUrl).using(SPFx(context));
          const fileItem = siteSP.web.getFileByServerRelativePath(serverRelativePath);
 
          // Get latest file properties
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
 
          // Get version history
          const versionData = await fileItem.versions
            .select("ID,VersionLabel,Created,Size,Url,CreatedBy/Title,CreatedBy/UserPrincipalName")
            .expand("CreatedBy")();
 
          // Base site URL for constructing download links
          const webUrl = context.pageContext.web.absoluteUrl;
 
          const allVersions = [
            ...versionData.map((v: any) => ({
              ID: v.ID,
              VersionLabel: v.VersionLabel,
              Created: new Date(v.Created),
              Size: v.Size,
              CreatedBy: v.CreatedBy,
              // Build direct download link for this version
              DownloadUrl: `${webUrl}/${v.Url}`,
            })),
            {
              ID: "current",
              VersionLabel: fileProps.UIVersionLabel,
              Created: new Date(fileProps.TimeLastModified),
              Size: fileProps.Length,
              CreatedBy: fileProps.ModifiedBy,
              // Current version download link
              DownloadUrl: `${webUrl}${fileProps.ServerRelativeUrl}`,
            },
          ];
 
          // Sort by date
          allVersions.sort((a, b) => a.Created.getTime() - b.Created.getTime());
 
          console.log("Version history loaded:", allVersions);
 
          if (!canceled) setVersions(allVersions);
 
        } catch (err: any) {
          if (!canceled) {
            setVersions([]);
            setError("Could not load version history.");
            console.error("Error loading version history:", err);
          }
        } finally {
          if (!canceled) setLoading(false);
        }
 
      })();
    }
    return () => {
      canceled = true;
    };
  }, [show, file, context]);
 
  if (!show) return null;
 
  const formatDate = (date: Date) => date.toLocaleString("en-US");
 
const downloadVersion = async (file: any, version: any) => {
    try {
      const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
 
      const folderPath = file.CurrentFolderPath || file.ServerRelativeUrl || "";
      const fileName = file.FileName || file.Name;
      //sourish 25/8/25 file.ServerRelativeUrl to handle vesrion from nodes
      let parts;
      let serverRelativePath = "";
      if (file.ServerRelativeUrl) {
        parts = file.ServerRelativeUrl.split("/").filter(Boolean);
        serverRelativePath = file.ServerRelativeUrl;
      } else {
        parts = file.CurrentFolderPath.split("/").filter(Boolean);
        serverRelativePath = `${file.CurrentFolderPath}/${file.FileName}`;
      }
 
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
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
 
      console.log(`Downloaded version ${version.VersionLabel}`);
    } catch (error) {
      console.error("Download failed:", error);
    }
  };
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
          position: "relative",
        }}
      >
        {/* Close button upar right corner me */}
        <button
          onClick={onClose}
          style={{
            position: "absolute",
            top: "15px",
            right: "15px",
            width: "32px",
            height: "32px",
            borderRadius: "50%",
            border: "none",
            backgroundColor: "#f0f0f0",
            color: "#333",
            fontSize: "20px",
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            transition: "background-color 0.2s",
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.backgroundColor = "#e0e0e0";
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.backgroundColor = "#f0f0f0";
          }}
          aria-label="Close"
        >
          ×
        </button>
 
        <h2 id="version-history-title" style={{ marginBottom: "15px", paddingRight: "40px" }}>
          Version History - {file.FileName || file.Name || "Unknown File"}
        </h2>
 
        {loading ? (
          <div>Loading versions...</div>
        ) : error ? (
          <div style={{ color: "red", marginBottom: "10px" }}>{error}</div>
        ) : versions.length > 0 ? (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #ddd" }}>
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
                  <td style={{ padding: "8px" }}>
                    {v.CreatedBy?.Title || "-"}
                    <br />
                    <small>{v.CreatedBy?.UserPrincipalName}</small>
                  </td>
                  <td style={{ padding: "8px" }}>
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
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div>No version history found.</div>
        )}
      </div>
    </div>
  );
};
 
export default VersionHistoryModal;