import * as React from "react";
import { useState, useEffect } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import type { IEssadmsMainProps } from "./IEssadmsMainProps";

interface PermissionResult {
  level: string;
  object: string;
  permission: string;
}
declare var SP: any;
const ArgPoc = ({ context }: { context: WebPartContext }) => {
  const [results, setResults] = useState<PermissionResult[]>([]);
  const userEmail = "user@domain.com"; // Replace with target user

  useEffect(() => {
    // Ensure SP.js is loaded
    SP.SOD.executeFunc("sp.js", "SP.ClientContext", () => {
      checkPermissions();
    });
  }, []);

  const checkPermissions = () => {
    const ctx = SP.ClientContext.get_current();
    const web = ctx.get_web();
    const user = web.ensureUser(userEmail);

    ctx.load(user);
    ctx.executeQueryAsync(
      () => {
        // ✅ Site Permissions
        const sitePerms = web.getUserEffectivePermissions(user.get_loginName());
        if (sitePerms.has(SP.PermissionKind.manageWeb)) {
          addResult("Site", window.location.href, "Manage Web");
        }

        // ✅ Libraries
        const lists = web.get_lists();
        ctx.load(lists, "Include(Title, BaseType)");
        ctx.executeQueryAsync(
          () => {
            const listEnumerator = lists.getEnumerator();
            while (listEnumerator.moveNext()) {
              const list = listEnumerator.get_current();
              if (list.get_baseType() === SP.BaseType.documentLibrary) {
                const libPerms = list.getUserEffectivePermissions(
                  user.get_loginName()
                );
                if (libPerms.has(SP.PermissionKind.viewListItems)) {
                  addResult("Library", list.get_title(), "View Items");
                }

                // ✅ Items inside library
                const caml = new SP.CamlQuery();
                const items = list.getItems(caml);
                ctx.load(items);
                ctx.executeQueryAsync(
                  () => {
                    const itemEnum = items.getEnumerator();
                    while (itemEnum.moveNext()) {
                      const item = itemEnum.get_current();
                      const itemPerms = item.getUserEffectivePermissions(
                        user.get_loginName()
                      );
                      if (itemPerms.has(SP.PermissionKind.viewListItems)) {
                        addResult(
                          "Item",
                          item.get_item("FileLeafRef"),
                          "View Item"
                        );
                      }
                    }
                  },
                  (s: any, e: any) =>
                    console.error("Error loading items:", e.get_message())
                );
              }
            }
          },
          (s:any, e:any) => console.error("Error loading lists:", e.get_message())
        );
      },
      (s:any, e:any) => console.error("Error loading user:", e.get_message())
    );
  };

  const addResult = (level: string, object: string, permission: string) => {
    setResults((prev) => [...prev, { level, object, permission }]);
  };

  return (
    <div>
      <h2>User Permission Report</h2>
      <table style={{ borderCollapse: "collapse", width: "100%" }}>
        <thead>
          <tr>
            <th style={{ border: "1px solid #ccc", padding: "8px" }}>Level</th>
            <th style={{ border: "1px solid #ccc", padding: "8px" }}>Object</th>
            <th style={{ border: "1px solid #ccc", padding: "8px" }}>
              Permission
            </th>
          </tr>
        </thead>
        <tbody>
          {results.map((res, idx) => (
            <tr key={idx}>
              <td style={{ border: "1px solid #ccc", padding: "8px" }}>
                {res.level}
              </td>
              <td style={{ border: "1px solid #ccc", padding: "8px" }}>
                {res.object}
              </td>
              <td style={{ border: "1px solid #ccc", padding: "8px" }}>
                {res.permission}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

const Useraccessreport: React.FC<IEssadmsMainProps> = (props) => {
  return <ArgPoc context={props.context} />;
};

export default Useraccessreport;
