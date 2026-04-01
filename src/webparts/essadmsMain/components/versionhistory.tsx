import * as React from "react";
import { Modal, Button, Spinner } from "react-bootstrap";
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
          const folderPath =
            file.CurrentFolderPath || file.ServerRelativeUrl || "";
          const fileName = file.FileName || file.Name;

          let serverRelativePath = "";
          if (file.ServerRelativeUrl) {
            serverRelativePath = file.ServerRelativeUrl;
          } else {
            serverRelativePath = folderPath.endsWith("/")
              ? `${folderPath}${fileName}`
              : `${folderPath}/${fileName}`;
          }

          const tenantUrl =
            context.pageContext.web.absoluteUrl.split("/sites/")[0];
          const parts = folderPath.split("/").filter(Boolean);
          const subsitePath = "/" + parts.slice(0, 3).join("/");
          const subsiteUrl = `${tenantUrl}${subsitePath}`;

          const siteSP = spfi(subsiteUrl).using(SPFx(context));
          const fileItem =
            siteSP.web.getFileByServerRelativePath(serverRelativePath);

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

          const versionData = await fileItem.versions
            .select(
              "ID,VersionLabel,Created,Size,Url,CreatedBy/Title,CreatedBy/UserPrincipalName"
            )
            .expand("CreatedBy")();

          const webUrl = context.pageContext.web.absoluteUrl;

          const allVersions = [
            ...versionData.map((v: any) => ({
              ID: v.ID,
              VersionLabel: v.VersionLabel,
              Created: new Date(v.Created),
              Size: v.Size,
              CreatedBy: v.CreatedBy,
              DownloadUrl: `${webUrl}/${v.Url}`,
            })),
            {
              ID: "current",
              VersionLabel: fileProps.UIVersionLabel,
              Created: new Date(fileProps.TimeLastModified),
              Size: fileProps.Length,
              CreatedBy: fileProps.ModifiedBy,
              DownloadUrl: `${webUrl}${fileProps.ServerRelativeUrl}`,
            },
          ];

          allVersions.sort(
            (a, b) => a.Created.getTime() - b.Created.getTime()
          );

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

  const formatDate = (date: Date) => date.toLocaleString("en-US");

  const downloadVersion = async (file: any, version: any) => {
    try {
      const tenantUrl =
        context.pageContext.web.absoluteUrl.split("/sites/")[0];

      const folderPath =
        file.CurrentFolderPath || file.ServerRelativeUrl || "";
      const fileName = file.FileName || file.Name;

      let parts;
      let serverRelativePath = "";
      if (file.ServerRelativeUrl) {
        parts = file.ServerRelativeUrl.split("/").filter(Boolean);
        serverRelativePath = file.ServerRelativeUrl;
      } else {
        parts = file.CurrentFolderPath.split("/").filter(Boolean);
        serverRelativePath = `${file.CurrentFolderPath}/${file.FileName}`;
      }

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
    <Modal
      show={show}
      onHide={onClose}
      container={() => document.getElementById('filelistcontainer')}
      // backdrop="static"
      // ritik
      backdrop={true}
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
        ) : error ? (
          <div style={{ color: "red", marginBottom: "10px" }}>{error}</div>
        ) : versions.length > 0 ? (
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #ddd" }}>
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
                  <td style={{ padding: "8px" }}>
                    {v.CreatedBy?.Title || "-"}
                    <br />
                    <small>{v.CreatedBy?.UserPrincipalName}</small>
                  </td>
                  <td style={{ padding: "8px" }}>
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
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <div>No version history found.</div>
        )}
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