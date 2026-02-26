// import * as React from "react";
// import { useEffect, useRef, useState } from "react";
// import { spfi, SPFx } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/site-users/web";
// import Select from "react-select";
// import Swal from "sweetalert2";
// import { WebPartContext } from "@microsoft/sp-webpart-base";

// interface Props {
//   context: WebPartContext;
//   siteTitle: string;
//   siteUrl: string;                 // child site url
//   documentLibraryName: string;
//   onClose: () => void;
// }

// interface Row {
//   id: number;
//   selectionType: "All" | "One";
//   approvedUserList: any[];
// }

// const ESSA_SITE_URL = "https://officeindia.sharepoint.com/sites/ESSA";

// const ManageWorkflow: React.FC<Props> = ({
//   context,
//   siteTitle,
//   siteUrl,
//   documentLibraryName,
//   onClose
// }) => {

//   // 🔹 ESSA (Users)
//   const spRoot = React.useMemo(
//     () => spfi(ESSA_SITE_URL).using(SPFx(context)),
//     [context]
//   );

//   // 🔹 Child Site (Lists)
//   const spChild = React.useMemo(
//     () => spfi(siteUrl).using(SPFx(context)),
//     [context, siteUrl]
//   );

//   const currentUserEmail = useRef<string>("");

//   const [users, setUsers] = useState<any[]>([]);
//   const [rows, setRows] = useState<Row[]>([
//     { id: 0, selectionType: "One", approvedUserList: [] }
//   ]);
//   const [toggleApprover, setToggleApprover] = useState<"Yes" | "No">("No");

//   const approvalMasterItemId = useRef<number>(0);
//   const isUpdate = useRef<boolean>(false);





 



//   // ---------------- INIT ----------------
//   useEffect(() => {
//     if (!siteUrl) return;

//     (async () => {
//       try {
//         const me = await spChild.web.currentUser();
//         currentUserEmail.current = me.Email;

//         await loadUsersFromESSA();
//         await loadInitialState();
//       } catch (e) {
//         console.error(e);
//         Swal.fire("Error", "Error loading workflow", "error");
//       }
//     })();
//   }, [siteUrl]);

//   useEffect(() => {
//     loadUsersFromESSA();

//   }, []);

//   // ---------------- USERS FROM ESSA ----------------
//   const loadUsersFromESSA = async () => {
//     const siteUsers = await spRoot.web.siteUsers();

//     setUsers(
//       siteUsers
//         .filter(u => u.Email)
//         .map(u => ({
//           label: u.Title,
//           value: u.Title,
//           userId: u.Id
//         }))
//     );
//   };








//   // ---------------- LOAD WORKFLOW (CHILD SITE) ----------------
//   const loadInitialState = async () => {
//     const master = await spChild.web.lists
//       .getByTitle("DMSPreviewFormMaster")
//       .items
//       .filter(
//         `SiteName eq '${siteTitle}'
//          and DocumentLibraryName eq '${documentLibraryName}'
//          and IsDocumentLibrary eq 1`
//       )();

//     if (!master.length) {
//       setToggleApprover("No");
//       return;
//     }

//     approvalMasterItemId.current = master[0].Id;
//     isUpdate.current = master[0].IsApproval;

//     if (!isUpdate.current) {
//       setToggleApprover("No");
//       return;
//     }

//     setToggleApprover("Yes");

//     const existing = await spChild.web.lists
//       .getByTitle("DMSFolderPermissionMaster")
//       .items
//       .select("ApprovalType", "Level", "ApprovalUser/Title", "ApprovalUser/Id")
//       .expand("ApprovalUser")
//       .filter(
//         `CurrentUser eq '${currentUserEmail.current}'
//          and SiteName eq '${siteTitle}'
//          and DocumentLibraryName eq '${documentLibraryName}'`
//       )();

//     const grouped: any = {};

//     existing.forEach(i => {
//       if (!grouped[i.Level]) {
//         grouped[i.Level] = {
//           id: i.Level - 1,
//           selectionType: i.ApprovalType === 1 ? "All" : "One",
//           approvedUserList: []
//         };
//       }

//       grouped[i.Level].approvedUserList.push({
//         label: i.ApprovalUser.Title,
//         value: i.ApprovalUser.Title,
//         userId: i.ApprovalUser.Id
//       });
//     });

//     setRows(Object.values(grouped));
//   };

//   // ---------------- SAVE ----------------
//   const handleSave = async () => {
//     try {
//       if (rows.some(r => r.approvedUserList.length === 0)) {
//         Swal.fire("Validation", "Select at least one approver per level", "warning");
//         return;
//       }

//       if (!isUpdate.current) {
//         await spChild.web.lists
//           .getByTitle("DMSPreviewFormMaster")
//           .items.getById(approvalMasterItemId.current)
//           .update({ IsApproval: true });
//       }

//       const oldItems = await spChild.web.lists
//         .getByTitle("DMSFolderPermissionMaster")
//         .items
//         .filter(
//           `CurrentUser eq '${currentUserEmail.current}'
//            and SiteName eq '${siteTitle}'
//            and DocumentLibraryName eq '${documentLibraryName}'`
//         )();

//       for (const item of oldItems) {
//         await spChild.web.lists
//           .getByTitle("DMSFolderPermissionMaster")
//           .items.getById(item.Id)
//           .delete();
//       }

//       for (const row of rows) {
//         for (const user of row.approvedUserList) {
//           await spChild.web.lists
//             .getByTitle("DMSFolderPermissionMaster")
//             .items.add({
//               SiteName: siteTitle,
//               DocumentLibraryName: documentLibraryName,
//               CurrentUser: currentUserEmail.current,
//               ApprovalUserId: user.userId,
//               ApprovalType: row.selectionType === "All" ? 1 : 0,
//               Level: row.id + 1
//             });
//         }
//       }

//       Swal.fire("Success", "Workflow saved successfully", "success");
//       onClose();
//     } catch (e) {
//       console.error(e);
//       Swal.fire("Error", "Error saving workflow", "error");
//     }
//   };

//   // ---------------- UI ----------------
//   return (
//     <div className="modal show d-block" style={{ background: "rgba(0,0,0,.4)" }}>
//       <div className="modal-dialog modal-xl">
//         <div className="modal-content p-4">

//           {toggleApprover === "No" ? (
//             <>
//               <h5>This document library does not have an approver.</h5>
//               <button className="btn btn-primary" onClick={() => setToggleApprover("Yes")}>
//                 Set Approver
//               </button>
//               <button className="btn btn-secondary ms-2" onClick={onClose}>
//                 Cancel
//               </button>
//             </>
//           ) : (
//             <>
//               <h5>Approval Hierarchy</h5>

//               {rows.map(row => (
//                 <div key={row.id} className="d-flex gap-2 mb-2">
//                   <input disabled value={`Level ${row.id + 1}`} className="form-control w-25" />

//                   <Select
//                     isMulti
//                     options={users}
//                     value={row.approvedUserList}
//                     onChange={(v: any) =>
//                       setRows(r =>
//                         r.map(x => x.id === row.id ? { ...x, approvedUserList: v } : x)
//                       )
//                     }
//                     className="flex-grow-1"
//                   />

//                   <div>
//                     <label>
//                       <input
//                         type="radio"
//                         checked={row.selectionType === "One"}
//                         onChange={() =>
//                           setRows(r =>
//                             r.map(x => x.id === row.id ? { ...x, selectionType: "One" } : x)
//                           )
//                         }
//                       /> One
//                     </label>
//                     <label className="ms-2">
//                       <input
//                         type="radio"
//                         checked={row.selectionType === "All"}
//                         onChange={() =>
//                           setRows(r =>
//                             r.map(x => x.id === row.id ? { ...x, selectionType: "All" } : x)
//                           )
//                         }
//                       /> All
//                     </label>
//                   </div>
//                 </div>
//               ))}

//               <button
//                 className="btn btn-link"
//                 onClick={() =>
//                   setRows(r => [...r, { id: r.length, selectionType: "One", approvedUserList: [] }])
//                 }
//               >
//                 + Add Level
//               </button>

//               <div className="text-end mt-3">
//                 <button className="btn btn-success" onClick={handleSave}>
//                   Save
//                 </button>
//                 <button className="btn btn-danger ms-2" onClick={onClose}>
//                   Cancel
//                 </button>
//               </div>
//             </>
//           )}
//         </div>
//       </div>
//     </div>
//   );
// };

// export default ManageWorkflow;



import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/site-users/web";
import Select from "react-select";
import Swal from "sweetalert2";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface Props {
  context: WebPartContext;
  siteTitle: string;
  siteUrl: string;
  documentLibraryName: string;
  file:any;
  
  onClose: () => void;

}

interface Row {
  id: number;
  selectionType: "All" | "One";
  approvedUserList: any[];
  file?:any;
}

const ESSA_SITE_URL = "https://officeindia.sharepoint.com/sites/ESSA";

const ManageWorkflow: React.FC<Props> = ({
  context,
  siteTitle,
  siteUrl,
  documentLibraryName,
  file,
  onClose

}) => {
  console.log("files" , file);
  const siteUrlforfile = file.__siteUrl;

  console.log("siteUrlforfile" , siteUrlforfile);
  

  const spfilepath = React.useMemo(() => spfi(siteUrlforfile).using(SPFx(context)), [context, siteUrlforfile]);
  




  const spRoot = React.useMemo(() => spfi(ESSA_SITE_URL).using(SPFx(context)), [context]);


  const spChild = React.useMemo(() => spfi(siteUrl).using(SPFx(context)), [context, siteUrl]);

  const currentUserEmail = useRef("");
  const approvalMasterItemId = useRef(0);
  const isUpdate = useRef(false);

  const [users, setUsers] = useState<any[]>([]);
  const [rows, setRows] = useState<Row[]>([{ id: 0, selectionType: "One", approvedUserList: [] }]);
  const [toggleApprover, setToggleApprover] = useState<"Yes" | "No">("No");

  // INIT
  useEffect(() => {
    (async () => {
      const me = await spChild.web.currentUser();
      currentUserEmail.current = me.Email;

      await loadUsers();
      await loadWorkflow();
    })();
  }, []);

  // USERS
  const loadUsers = async () => {
    const u = await spRoot.web.siteUsers();
    setUsers(
      u.filter(x => x.Email).map(x => ({
        label: x.Title,
        value: x.Title,
        userId: x.Id
      }))
    );
  };

  // LOAD EXISTING APPROVAL
  const loadWorkflow = async () => {

    const master = await spfilepath.web.lists
      .getByTitle("DMSPreviewFormMaster")
      .items.filter(
        `SiteName eq '${siteTitle}' and DocumentLibraryName eq '${documentLibraryName}'`
      )();

      // srs 19/2/26
      approvalMasterItemId.current = master[0].Id;
      isUpdate.current = true;
       if (!master.length || !master[0].IsApproval) return;

    // approvalMasterItemId.current = master[0].Id;
    // isUpdate.current = true;
    setToggleApprover("Yes");

    const items = await spfilepath.web.lists
      .getByTitle("DMSFolderPermissionMaster")
      .items
      .select("ApprovalType","Level","ApprovalUser/Title","ApprovalUser/Id")
      .expand("ApprovalUser")
      .filter(
        `SiteName eq '${siteTitle}' and DocumentLibraryName eq '${documentLibraryName}'`
      )();

    const grouped:any = {};

    items.forEach(i=>{
      if(!grouped[i.Level]){
        grouped[i.Level]={
          id:i.Level-1,
          selectionType:i.ApprovalType===1?"All":"One",
          approvedUserList:[]
        }
      }

      grouped[i.Level].approvedUserList.push({
        label:i.ApprovalUser.Title,
        value:i.ApprovalUser.Title,
        userId:i.ApprovalUser.Id
      });
    });

    setRows(Object.values(grouped));
  };

  // SAVE
  // const handleSave = async () => {

  //   if(rows.some(r=>r.approvedUserList.length===0)){
  //     Swal.fire("Validation","Select approver for each level","warning");
  //     return;
  //   }

  //   if(!isUpdate.current){
  //     await spfilepath.web.lists
  //       .getByTitle("DMSPreviewFormMaster")
  //       .items.getById(approvalMasterItemId.current)
  //       .update({IsApproval:true});
  //   }

  //   const old = await spfilepath.web.lists
  //     .getByTitle("DMSFolderPermissionMaster")
  //     .items.filter(
  //       `SiteName eq '${siteTitle}' and DocumentLibraryName eq '${documentLibraryName}'`
  //     )();

  //   for(const o of old){
  //     await spfilepath.web.lists.getByTitle("DMSFolderPermissionMaster").items.getById(o.Id).delete();
  //   }

  //   for(const row of rows){
  //     for(const u of row.approvedUserList){
  //       await spfilepath.web.lists.getByTitle("DMSFolderPermissionMaster").items.add({
  //         SiteName:siteTitle,
  //         DocumentLibraryName:documentLibraryName,
  //         CurrentUser:currentUserEmail.current,
  //         ApprovalUserId:u.userId,
  //         ApprovalType:row.selectionType==="All"?1:0,
  //         Level:row.id+1
  //       });
  //     }
  //   }

  //   Swal.fire("Success","Approval saved","success");
  //   onClose();
  // };

  // srs 19/2/26 
const handleSave = async () => {
  if(rows.some(r => r.approvedUserList.length === 0)){
    Swal.fire("Validation", "Select approver for each level", "warning");
    return;
  }

  try {
    // FIX: Handle the Master Record
    if (isUpdate.current && approvalMasterItemId.current !== 0) {
      await spfilepath.web.lists
        .getByTitle("DMSPreviewFormMaster")
        .items.getById(approvalMasterItemId.current)
        .update({ IsApproval: true });
    } else {
      // If it doesn't exist, you should probably ADD it or fetch the ID first
      // Example of adding if it's missing:
      /*
      await spfilepath.web.lists.getByTitle("DMSPreviewFormMaster").items.add({
          SiteName: siteTitle,
          DocumentLibraryName: documentLibraryName,
          IsApproval: true
      });
      */
    }

    // Delete old permissions (This part is fine, but ensure the filter is specific)
    const old = await spfilepath.web.lists
      .getByTitle("DMSFolderPermissionMaster")
      .items.filter(`SiteName eq '${siteTitle}' and DocumentLibraryName eq '${documentLibraryName}'`)();

    // Use a batch or Promise.all for faster deletion
    for(const o of old){
      await spfilepath.web.lists.getByTitle("DMSFolderPermissionMaster").items.getById(o.Id).delete();
    }

    // Add new levels
    for(const row of rows){
      for(const u of row.approvedUserList){
        await spfilepath.web.lists.getByTitle("DMSFolderPermissionMaster").items.add({
          SiteName: siteTitle,
          DocumentLibraryName: documentLibraryName,
          CurrentUser: currentUserEmail.current,
          ApprovalUserId: u.userId,
          ApprovalType: row.selectionType === "All" ? 1 : 0,
          Level: row.id + 1
        });
      }
    }

    Swal.fire("Success", "Approval saved", "success");
    onClose();
  } catch (error) {
    console.error(error);
    Swal.fire("Error", "Failed to save workflow", "error");
  }
};
  // UI
  return (
    <div className="modal show d-block" style={{background:"rgba(0,0,0,.2)"}}>
      <div className="modal-dialog modal-xl">
        <div className="modal-content p-4">

          {toggleApprover==="No"?(
            <>
              <h5>No approver configured</h5>
              <button style={{width:'130px', margin:'auto'}} className="btn btn-primary" onClick={()=>setToggleApprover("Yes")}>Add Approval</button>
            </>
          ):(
           <>
           <div className="d-flex align-items-center justify-content-between border-bottom">
  <h4 className="modal-title m-0 mb-2">Approval Hierarchy</h4>
    <button
    type="button"
    className="btn btn-link mt-0 p-0 mb-2"
    onClick={() =>
      setRows(r => [
        ...r,
        { id: r.length, selectionType: "One", approvedUserList: [] }
      ])
    }
  >
    
   <img 
    src={require("../assets/addbn.png")} 
    alt="Rename" />
  </button>
 </div>
  <div className="table-responsive">
    <table className="table table-bordered align-middle">
      <thead className="table-light">
        <tr>
          <th style={{ width: "15%" }}>Level</th>
          <th style={{ width: "45%" }}>Approver</th>
          <th style={{ width: "20%" }}>Type</th>
          <th style={{ width: "20%" }} className="text-center">Action</th>
        </tr>
      </thead>
 
      <tbody>
        {rows.map((row, index) => (
          <tr key={row.id}>
 
            {/* Level */}
            <td>
              <input
                disabled
                value={`Level ${index + 1}`}
                className="form-control"
              />
            </td>
 
            {/* Approver */}
            <td>
              <Select
                isMulti
                options={users}
                value={row.approvedUserList}
                onChange={(v: any) =>
                  setRows(r =>
                    r.map(x =>
                      x.id === row.id
                        ? { ...x, approvedUserList: v }
                        : x
                    )
                  )
                }
              />
            </td>
 
            {/* Type */}
            <td>
              <div className="d-flex gap-3">
                <div className="d-flex align-items-center gap-2">
                  <input
                    type="radio"
                    name={`selection-${row.id}`}
                    checked={row.selectionType === "One"}
                    onChange={() =>
                      setRows(r =>
                        r.map(x =>
                          x.id === row.id
                            ? { ...x, selectionType: "One" }
                            : x
                        )
                      )
                    }
                  />{" "}
                  One
                </div>
 
                   <div className="d-flex align-items-center gap-2">
                  <input
                    type="radio"
                    name={`selection-${row.id}`}
                    checked={row.selectionType === "All"}
                    onChange={() =>
                      setRows(r =>
                        r.map(x =>
                          x.id === row.id
                            ? { ...x, selectionType: "All" }
                            : x
                        )
                      )
                    }
                  />{" "}
                  All
                </div>
              </div>
            </td>
 
            {/* Delete */}
            <td className="text-center">
              <button style={{background:'transparent', border:'none'}}
                className="btn btn-sm btn-danger mt-0"
                onClick={() =>
                  setRows(r => r.filter(x => x.id !== row.id))
                }
              >
               <img 
    src={require("../assets/deletn.png")} 
    alt="Rename" />
              </button>
            </td>
 
          </tr>
        ))}
 
        {rows.length === 0 && (
          <tr>
            <td colSpan={4} className="text-center text-muted">
              No approval level added.
            </td>
          </tr>
        )}
      </tbody>
    </table>
  </div>
 
  {/* Add Level Button */}

 
  {/* Footer Buttons */}
  <div className="text-end">
   <button type="button"
    style={{
      background: 'transparent',
      border: 'none',
      padding: 0
    }}
    onClick={handleSave}
  > <div className="newicon">
    <img
      src={require("../assets/submit-new1.png")}
      alt="Save"
      style={{ height: "32px" }}
    /> </div>
                  </button>
    <button style={{background:'transparent', border:'none',   padding: '0'}} className="btn btn-secondary ms-2" onClick={onClose}>
 <div className="newicon">
 <img
      src={require("../assets/popupcross.png")}
      alt="Save"
      style={{ height: "32px" }}
    /></div>
 
    </button>
  </div>
</>
          )}

        </div>
      </div>
    </div>
  );
};

export default ManageWorkflow;

