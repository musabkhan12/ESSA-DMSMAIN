import * as React from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/files";
import "@pnp/sp/webs";
import type { WebPartContext } from "@microsoft/sp-webpart-base";

interface DirectDownloadProps {
  file: any;
  context: WebPartContext;
  trigger?: boolean; // optional → if you want to auto-download on mount
}

const DirectDownloader: React.FC<DirectDownloadProps> = ({ file, context, trigger }) => {
  React.useEffect(() => {
    if (trigger && file) {
      // eslint-disable-next-line @typescript-eslint/no-floating-promises, @typescript-eslint/no-use-before-define
      downloadFile(file);
    }
  }, [trigger, file]);

  const downloadFile = async (file: any) => {
    try {
      const tenantUrl = context.pageContext.web.absoluteUrl.split("/sites/")[0];
      const parts = file.CurrentFolderPath.split("/").filter(Boolean);
      const subsitePath = "/" + parts.slice(0, 3).join("/");
      const subsiteUrl = `${tenantUrl}${subsitePath}`;

      const siteSP = spfi(subsiteUrl).using(SPFx(context));
      const fileItem = siteSP.web.getFileByServerRelativePath(
        `${file.CurrentFolderPath}/${file.FileName}`
      );

      const blob: Blob = await fileItem.getBlob();

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = file.FileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      console.log(`Directly downloaded file: ${file.FileName}`);
    } catch (error) {
      console.error("Direct download failed:", error);
    }
  };

  return null; // no modal, silent component
};

export default DirectDownloader;
