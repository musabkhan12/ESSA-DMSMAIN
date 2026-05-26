import React, { useEffect } from 'react'
import { useRef, useState } from "react";
// import { getSP } from "../loc/pnpjsConfig";
import { getSP } from "../../advancedSearch/loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import Select from "react-select";
import styles from './Form.module.scss';
import Swal from "sweetalert2";
interface ManageFolderPermissionProps {
    OthProps: { [key: string]: string };
    onReturnToMain: () => void;
  }

let IsUpdate:boolean;
let IsPrivateColumnId:number;
let folderName:any;
let IsPrivateColumnIdForDocumentLibrary:number
let division='';
let department='';
let path='';
const ManageFolderPermission : React.FC<ManageFolderPermissionProps> = ({
    OthProps,
    onReturnToMain,
  }) =>{
 
    console.log(OthProps, "oth props");
    const sp: SPFI = getSP();

 

    const currentUserEmailRef = useRef('');
    const getcurrentuseremail = async()=>{
        const userdata = await sp.web.currentUser();
        currentUserEmailRef.current = userdata.Email;
        defaultValue(null);
    }

    // New Code
    const [rowsForPermission, setRowsForPermission] = React.useState<
    { id: number; selectedUserForPermission: string[]; selectedPermission:"" }[]
    >([{ id: 0, selectedUserForPermission: [],selectedPermission:"" }]);
    console.log("rowsForPermission",rowsForPermission);
    // End
    // creating state variable to store the table data
    const [tableData,setTableData]= React.useState<any[]>([])
    const [users, setUsers] = React.useState<any[]>([]);
    console.log("Users Array", users);

    const [toggelPermission,setTogglePermission]=React.useState<string>();
    console.log("toggelPermission",toggelPermission);

    const [errorsForPermissionSelection, setErrorsForPermissionSelection] = useState<{ [key: number]: { userSelect?: string, permissionSelect?: string; duplicate?: string } }>({});
      const validatePermissionsSelect = () => {
        let isValid = true;
        const newErrors: { [key: number]: { userSelect?: string, permissionSelect?: string; duplicate?: string } } = {};
      
        rowsForPermission.forEach((row:any) => {
          if (!row.selectedUserForPermission || row.selectedUserForPermission.length === 0) {
            newErrors[row.id] = { ...newErrors[row.id], userSelect: 'Please select at least one user.' };
            isValid = false;
          }
          if (!row.selectedPermission) {
            newErrors[row.id] = { ...newErrors[row.id], permissionSelect: 'Please select a permission.' };
            isValid = false;
          }
        // **Collect all duplicate users**
        let duplicateUsers: string[] = [];
          row.selectedUserForPermission.forEach((user: any) => {
            const isDuplicate = tableData.some(
                (item) => item.userId === user.userId && item.Permission === row.selectedPermission.value
            );
            if (isDuplicate) {
              duplicateUsers.push(user.value);
          }
        });

        // If there are duplicates, set a single error message
        if (duplicateUsers.length > 0) {
            newErrors[row.id] = {
              ...newErrors[row.id],
                duplicate: `${duplicateUsers.join(", ")} already have "${row.selectedPermission.value}" permission.`,
              };
                isValid = false;
        }
        });
      
        setErrorsForPermissionSelection(newErrors);
        return isValid;
      };
      console.log("newErrors after added error message",errorsForPermissionSelection)

    const handlesetTogglePermission=()=>{
        setTogglePermission("Yes");
    }

    // const [selectedPermission,setSelectedPermission]=useState([]);
    // const [defaultUser,setDefaultUser]=useState<{
    //     userId:number,
    //     value: String,
    //     label: String,
    //   }[]>([]);

    // console.log("Default user array",defaultUser)
  
    const permissionArray:any[]=[
      {value:"Full Control",label:"Full Control"},
      {value:"Contribute",label:"Contribute"},
      {value:"Edit",label:"Edit"},
      {value:"Read",label:"Read"},
      {value:"View",label:"View"}
    ];
  
    const handleUserSelectForPermission=(selectedUser:any,Id:any)=>{
        console.log("Selected user for permission",selectedUser);
        const newRows = rowsForPermission.map((row) =>
          row.id === Id ? { ...row, selectedUserForPermission: selectedUser } : row
        );
        setRowsForPermission(newRows);
    }
  
    const handlePermissionSelect=(selectedPermission:any,Id:any)=>{
        console.log("Selected Permission",selectedPermission)
        const newRows = rowsForPermission.map((row) =>
          row.id === Id ? { ...row, selectedPermission: selectedPermission } : row
        );
        setRowsForPermission(newRows);

    }

    const handleAddRow = (
      event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
    ) => {
      event.preventDefault();
      const newId = rowsForPermission.length ? rowsForPermission[rowsForPermission.length - 1].id + 1 : 0;
      setRowsForPermission([
        ...rowsForPermission,
        { id: newId, selectedUserForPermission: [], selectedPermission:"" },
      ]);
    };

    const handleRemoveRow = (
      id: number,
      event: React.MouseEvent<HTMLAnchorElement, MouseEvent>
    ) => {
      event.preventDefault();
      setRowsForPermission(rowsForPermission.filter((row) => row.id !== id));
  };

    // Fetch users from SharePoint
  React.useEffect(() => {
    getcurrentuseremail();
    console.log(currentUserEmailRef.current ,"my current id")
    const fetchUsers = async () => {
      try {
        if(OthProps.externalFolder === "true"){
          const user0 = await sp.web.siteUsers();
          const user1 = await sp.web.siteGroups();
          const groupsArray=user1.map((user)=>(
            {
            PrincipalType:user.PrincipalType,
            userId:user.Id,
            value: user.Title,
            label: user.Title,
            email: user.Title,
            }
          ))
          const combineUsersArray=user0.map((user)=>(
                {
                userId:user.Id,
                value: user.Title,
                label: user.Title,
                email: user.Email,
            }
          ))
          let resultArray =[...combineUsersArray, ...groupsArray];
          console.log("resultArray --->",resultArray)
          setUsers(resultArray);
        }else{
        //   const siteContext = await sp.site.openWebById(OthProps.SiteID);
        //   const user0 = await siteContext.web.siteUsers();
        //   const combineUsersArray=user0.map((user)=>(
        //       {
        //       userId:user.Id,
        //       value: user.Title,
        //       label: user.Title,
        //       email: user.Email,
        //   }
        // ))
        // setUsers(combineUsersArray);
        // console.log("Sub site users",combineUsersArray);
        // fetch the data from site Gropus
        const [
          users,
          users1,
          users2,
          users3,
          users4,
          users5,
          users6,
          users7
        ] = await Promise.all([
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Read`).users(),
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Initiator`).users(),
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Contribute`).users(),
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Admin`).users(),
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_View`).users(),
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_AllUsers`).users(),
          sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Approval`).users(),
          sp.web.siteGroups.getByName(`DMSSuper_Admin`).users(),
        ]);

        const combineArray = [
          ...(users || []),
          ...(users1 || []),
          ...(users2 || []),
          ...(users3 || []),
          ...(users4 || []),
          ...(users5 || []),
          ...(users6 || []),
          ...(users7 || []),
        ];
        setUsers(
          combineArray.map((user) => ( 
          {
            userId:user.Id,
            value: user.Title,
            label: user.Title,
            email: user.Email,
          }
        ))
        );
        console.log("combineArray", combineArray);
        }
        // start
        // const siteContext = await sp.site.openWebById(OthProps.SiteID);
        // const user0 = await siteContext.web.siteUsers();
        // const combineUsersArray=user0.map((user)=>(
        //       {
        //       userId:user.Id,
        //       value: user.Title,
        //       label: user.Title,
        //       email: user.Email,
        //   }
        // ))
        // setUsers(combineUsersArray);
        // console.log("Sub site users",combineUsersArray);
        // const user0 = await sp.web.siteUsers();
        // const [
        //   users,
        //   users1,
        //   users2,
        //   users3,
        //   users4,
        // ] = await Promise.all([
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Read`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Initiator`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Contribute`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_Admin`).users(),
        //   sp.web.siteGroups.getByName(`${OthProps.SiteTitle}_View`).users(),
        // ]);
        // console.log(users, "users ", users1,users2,users3,users4);
        // const combineArray = [
        //   ...(users || []),
        //   ...(users1 || []),
        //   ...(users2 || []),
        //   ...(users3 || []),
        //   ...(users4 || []),
        // ];
        // setUsers(
        //   combineArray.map((user) => ( 
        //   {
        //     userId:user.Id,
        //     value: user.Title,
        //     label: user.Title,
        //     email: user.Email,
        //   }
        // ))
        // );
        // console.log("combineArray", combineArray);
        // end
      } catch (error) {
        console.error("Error fetching site users:", error);
      }
    };

    fetchUsers();
  }, []);

  // interface User {
  //   userId: number;
  //   value: string;
  //   label: string;
  //   Permission: string;
  // }
  
  // interface GroupedPermissions {
  //   id: number;
  //   selectedUserForPermission: User[];  
  //   selectedPermission: string;
  // }
// fetch the initial data from the  DMSFolderPrivacy list
    const defaultValue=async(flag:string=null)=>{
        try {
          // alert(`OthProps.FolderName is ${OthProps.FolderName}`)
          if(OthProps.FolderName === "null" ){
            folderName=null;
            // alert(`folder id undefied or null ${OthProps.FolderName}`)
          }else{
            folderName=`'${OthProps.FolderName}'`
          }
            // fetch the IsPrivate from DMSPreviewFormMaster to chcek the folder is private or not
            // const permissionDetails=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("IsPrivate","Id").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and IsDocumentLibrary eq 1`)();
            // console.log("approverDetails",permissionDetails);
            // IsUpdate=permissionDetails[0].IsPrivate;
            // IsPrivateColumnId=permissionDetails[0].Id;
            
              //   try {
              //     // if(files.FolderPath){

              //     interface IMember {
              //       PrincipalType: number;
              //       Title:String;
              //       Id:number 
              //     }
              //     interface IRoleAssignmentInfo {
              //       Member?: IMember; 
              //       roleDefinitionBindings?:IRoleAssignmentInfo
              //     }
 
              //      alert(`encodedFolderPath : ${encodedFolderPath}`)
              //      const {web} =await sp.site.openWebById(`b74b305b-6893-4f7e-ab77-d25bce5cfb04`)
           
              //      const itemId = folderItem.Id;
              //      alert(`itemId of folder is : ${itemId}`)
              
              //      const roleAssignments:IRoleAssignmentInfo[] = await web.lists.getByTitle(`CMSD1`).items.getById(itemId).roleAssignments();
              //      console.log("roleAssignments",roleAssignments);
              //      alert(`roleAssignments on folder is: ${roleAssignments}`)

              //      roleAssignments.forEach(async (roleAssignment) => {
              //       const member = await roleAssignment.Member;
              //       const roles = await roleAssignment.roleDefinitionBindings;
              //       console.log("User/Group: ", member.Title);
              //       console.log("Permissions: ", roles.map((role:any) => role.Name).join(", "));
              //   });
              //     // }
              //  } catch (error) {
                
              //  }

              // interface IMember {
              //     PrincipalType: number;
              //     Title: string;
              //     Id: number;
              // }
              
              // interface IRoleDefinitionBinding {
              //     Name: string;
              //     Id: number;
              // }
              
              // interface IRoleAssignmentInfo {
              //     Member?: IMember;
              //     RoleDefinitionBindings?: IRoleDefinitionBinding[]; // Correctly typed as an array
              // }
              
          
              //     try {
              //       
              //         alert(`encodedFolderPath : ${encodedFolderPath}`);
              
              //         const { web } = await sp.site.openWebById(`b74b305b-6893-4f7e-ab77-d25bce5cfb04`);
              //      
              //         const itemId = folderItem.Id;
              //         alert(`itemId of folder is : ${itemId}`);
              
              //         // Fetch role assignments with expanded properties
              //         const roleAssignments: IRoleAssignmentInfo[] = await web.lists
              //             .getByTitle(`CMSD1`)
              //             .items.getById(itemId)
              //             .roleAssignments.expand("Member", "RoleDefinitionBindings")(); 
              
              //         console.log("roleAssignments", roleAssignments);
              //         alert(`roleAssignments on folder is: ${JSON.stringify(roleAssignments)}`);
              
              //         // Iterate over role assignments
              //         roleAssignments.forEach((roleAssignment) => {
              //             const member = roleAssignment.Member;
              //             const roles = roleAssignment.RoleDefinitionBindings;
              
              //             if (member) {
              //                 console.log("User/Group: ", member.Title);
              //             }
              
              //             if (roles && roles.length > 0) {
              //                 console.log(
              //                     "Permissions: ",
              //                     roles.map((role) => role.Name).join(", ")
              //                 );
              //             }
              //         });
              //     } catch (error) {
              //         console.error("Error fetching folder permissions:", error);
              //     }

            console.log("Flag",flag)
            if(flag=== null){
            if(OthProps.FolderName === "null"){
              // fetch the IsPrivate from DMSPreviewFormMaster to chcek the folder is private or not
              const permissionDetails=await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("IsPrivate","Id").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and IsDocumentLibrary eq 1`)();
              console.log("approverDetails",permissionDetails);
              // IsUpdate=permissionDetails[0].IsPrivate;
              IsPrivateColumnIdForDocumentLibrary=permissionDetails[0].Id;
            }
            
            // alert(`here is filter data  sitetitle ${OthProps.SiteTitle} , OthProps.DocumentLibraryName${OthProps.DocumentLibraryName} , folderName ${folderName}`)
            // const permissionDetails=await sp.web.lists.getByTitle("DMSFolderMaster").items.select("*").filter(`SiteTitle eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and FolderName eq ${folderName}`)();
            const permissionDetails=await sp.web.lists.getByTitle("DMSFolderMaster").items.getById(Number(OthProps.FolderID))();
            console.log("permissionDetails1",permissionDetails);
            // IsUpdate=permissionDetails[0]?.IsPrivate;
            IsUpdate=permissionDetails?.IsPrivate;
            // IsPrivateColumnId=permissionDetails[0].Id;
            IsPrivateColumnId=permissionDetails.Id;
            // division=permissionDetails[0].Devision;
            division=permissionDetails.Devision;
            // department=permissionDetails[0].Department
            department=permissionDetails.Department
            
            // Bread crumb start
            path = OthProps.SiteTitle;
            if(division) {
              path += ` > ${division}`;
            }
            if(department) {
              path += ` > ${department}`;
            }
            if (OthProps.DocumentLibraryName) {
              path += ` > ${OthProps.DocumentLibraryName}`;
            }
          
            if (OthProps.FolderName !== 'null') {
              path += ` > ${OthProps.FolderName}`;
            }
            // End
           
              if(!IsUpdate){
                setTogglePermission("No");
              }else{
                setTogglePermission("Yes");
              }
            }
            
            // const fetchData=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("User","UserID","UserPermission","FolderName").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and CurrentUser eq '${currentUserEmailRef.current}' and FolderName eq ${null}`)();
            // alert(`folderName is :${folderName}`)
            const fetchData=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("User","UserID","UserPermission","FolderName","Id").filter(`SiteName eq '${OthProps.SiteTitle}' and DocumentLibraryName eq '${OthProps.DocumentLibraryName}' and FolderName eq ${folderName} and FolderID eq ${Number(OthProps.FolderID)}`)();
            console.log("Fetch data",fetchData);    
            

            // Initialize array to store the default users  
            const arrayToStoreDefaultUser = fetchData.map((user) => ({
                itemId:user.Id,
                userId: user.UserID,
                value: user.User,
                label: user.User,
                Permission:user.UserPermission
            }));

            // filter the data.
            const filteredData = arrayToStoreDefaultUser.filter(item => item.userId !== null && item.value !== null && item.label !== null);

            // const permission=filteredData[0].Permission;
            // const objectPermission={value:permission,label:permission};
            // console.log("Permission Object",objectPermission);
            // setSelectedPermission([objectPermission]);
            
            console.log("filteredData",filteredData);
            setTableData(filteredData)
            // New Code
            // Helper function to generate a random ID
            // const generateRandomId = (): number => Math.floor(Math.random() * 100000);
            // const grouped: { [key: string]: any } = {};
            // filteredData.forEach((user)=>{
            //       const { Permission } = user;
            //       // If the permission group doesn't exist, create it with a random id
            //       if (!grouped[Permission]) {
            //           grouped[Permission] = {
            //           id: generateRandomId(),
            //           selectedUserForPermission: [],
            //           selectedPermission: {value:Permission,label:Permission}
            //       };
            //     }

            //     // Add the user to the correct permission group
            //     grouped[Permission].selectedUserForPermission.push(user);
            // })
            // console.log("grouped",grouped)
            // const permissionsArray = Object.keys(grouped).map(key => grouped[key]);
            // console.log("permissionsArray",permissionsArray);
            // setRowsForPermission(permissionsArray)
            // End
            // setDefaultUser(filteredData);
            console.log("Fetch data from DMSFolderPrivacy",fetchData);
            

            // This code get all the users and groups on the document library or folders 
            // try {            
            //   const { web } = await sp.site.openWebById(`${OthProps.SiteID}`);
              
            //   let roleAssignments: any[];
            //   if(OthProps.FolderName  !== "null"){
            //     const folderItem = await web.getFolderByServerRelativePath(`${OthProps.FolderPath}`).listItemAllFields();
            //     const itemId = folderItem.Id;
            //     // alert(`itemId of folder is : ${itemId}`);
             
             
            //    roleAssignments= await web.lists.getByTitle(`${OthProps.DocumentLibraryName}`).items.getById(itemId).roleAssignments.expand('Member', 'RoleDefinitionBindings')();
            //     console.log("roleAssignments", roleAssignments);
            //   }else{
            //       // Fetch document library permissions
            //       roleAssignments = await web.lists
            //             .getByTitle(`${OthProps.DocumentLibraryName}`)
            //             .roleAssignments.expand("Member", "RoleDefinitionBindings")();

            //       console.log("Role Assignments for Document Library:", roleAssignments);
            //   }
            //   // const folderItem = await web.getFolderByServerRelativePath(`${OthProps.FolderPath}`).listItemAllFields();
            //   // const itemId = folderItem.Id;
            //   // // alert(`itemId of folder is : ${itemId}`);
             
             
            //   // const roleAssignments:any = await web.lists.getByTitle(`${OthProps.DocumentLibraryName}`).items.getById(itemId).roleAssignments.expand('Member', 'RoleDefinitionBindings')();
            //   // console.log("roleAssignments", roleAssignments);
             
             
            //   // for (const roleAssignment of roleAssignments) {
            //   //     const member = roleAssignment.Member;
            //   //     const roles = roleAssignment.RoleDefinitionBindings;
            //   //     console.log(`User/Group:Id", ${member.Id}`)
            //   //     console.log(`User/Group: ${member.Title}`);
            //   //     console.log("Permissions: ", roles.map((role:any) => role.Name).join(", "));
            //   // }
    
            //    // Initialize array to store the default users  
              
            //   const arrayToStoreDefaultUser = roleAssignments.map((roleAssignment:any) => ({
            //     userId: roleAssignment.Member.Id,
            //     value: roleAssignment.Member.Title,
            //     label:roleAssignment.Member.Title,
            //     Permission:roleAssignment.RoleDefinitionBindings.map((role:any) => role.Name).join(", "),
            //     IsGroupOrIsUser:roleAssignment.Member.PrincipalType
            //     }));
            //   console.log("arrayToStoreDefaultUser direct",arrayToStoreDefaultUser);
    
            //   // New Code
            //     // Helper function to generate a random ID
            //     const generateRandomId = (): number => Math.floor(Math.random() * 100000);
            //     const grouped: { [key: string]: any } = {};
            //     arrayToStoreDefaultUser.forEach((user:any)=>{
            //           const { Permission } = user;
            //           // If the permission group doesn't exist, create it with a random id
            //           if (!grouped[Permission]) {
            //               grouped[Permission] = {
            //               id: generateRandomId(),
            //               selectedUserForPermission: [],
            //               selectedPermission: {value:Permission,label:Permission}
            //           };
            //         }
    
            //         // Add the user to the correct permission group
            //         grouped[Permission].selectedUserForPermission.push(user);
            //     })
    
            //     const permissionsArray = Object.keys(grouped).map(key => grouped[key]);
            //     setRowsForPermission(permissionsArray)
    
            // } catch (error) {
            //   console.error("Error getting folder permissions:", error);
            // }

        } catch (error) {
            console.log("Erroe fetching data from DMSFolderPrivacy",error);
        }
    }
    

  const handleCreate=async()=>{
    console.log("create called");
    console.log("rowsForPermission",rowsForPermission);
    // console.log("selected User Array",defaultUser);
    // console.log("selected permission",selectedPermission);
    if(!validatePermissionsSelect()){
      return
    }
     
    // this code will make public folder to private and change all public checks to private
    // if(!IsUpdate){
    //   const getDataFromFolderPrivacy=await sp.web.lists.getByTitle("DMSFolderPrivacy").items.select("*").filter(`FolderID eq ${Number(OthProps.FolderID)} and User eq ${null} and UserID eq ${null}`)();
    //   console.log("getDataFromFolderPrivacy",getDataFromFolderPrivacy);
    //    // if(getDataFromFolderPrivacy.length > 0){
    //   //   try {
    //   //    await sp.web.lists.getByTitle("DMSFolderPrivacy").items.getById(getDataFromFolderPrivacy[0].ID).update({
    //   //       PublicFolderPermission:false,
    //   //     })
    //   //     console.log("PublicFolderPermission updated successfully");
    //   //   } catch (error) {
    //   //     console.log("error in PublicFolderPermission updated",error);
    //   //   }
    //   // }
    //   if(getDataFromFolderPrivacy.length > 0){
    //     alert("getDataFromFolderPrivacy[0].PublicFolderPermission " + getDataFromFolderPrivacy[0].PublicFolderPermission )
    //     if(getDataFromFolderPrivacy[0].PublicFolderPermission === true){
    //       alert("Please update the folder privacy first")
    //       try {
    //         await sp.web.lists.getByTitle("DMSFolderPrivacy").items.getById(getDataFromFolderPrivacy[0].ID).update({
    //            PublicFolderPermission:false,
    //          })
    //          console.log("PublicFolderPermission updated successfully");
    //        } catch (error) {
    //          console.log("error in PublicFolderPermission updated",error);
    //        }
    //     }
      
    //   }
    //     try {
    //       await sp.web.lists.getByTitle('DMSFolderMaster').items.getById(Number(OthProps.FolderID)).update({
    //         IsPrivate:true
    //       })
    //       console.log("successfully updated the IsPrivate column");
    //     } catch (error) {
    //       console.log("error in updating IsPrivate column",error);
    //     }

    //     if(OthProps.FolderName === "null"){
          
    //       try {
    //         await sp.web.lists.getByTitle('DMSPreviewFormMaster').items.getById(IsPrivateColumnIdForDocumentLibrary).update({
    //           IsPrivate:true
    //         })
    //         console.log("successfully updated the IsPrivate column of the DMSPreviewformmaster list")
    //       } catch (error) {
    //         console.log("error in updated the IsPrivate column of the DMSPreviewformmaster list")
    //       }
    //     }
      
    // }

    try {
        
        const payloadForDMSFolderPrivacy={
            SiteName:OthProps.SiteTitle,
            DocumentLibraryName:OthProps.DocumentLibraryName,
            CurrentUser:currentUserEmailRef.current,
            IsModified:true,
            FolderID:Number(OthProps.FolderID)
            // UserPermission:selectedPermission[0].value
        }

        // Add the folder name if its folder
        if(OthProps.FolderName !== "null"){
          (payloadForDMSFolderPrivacy as any).FolderName=OthProps.FolderName;
        }
        rowsForPermission.forEach((row:any)=>{
            try {
                
                // (payloadForDMSFolderPrivacy as any).User=user.value;
                // (payloadForDMSFolderPrivacy as any).UserID=user.userId;
                (payloadForDMSFolderPrivacy as any).UserPermission=row.selectedPermission.value;
                row.selectedUserForPermission.forEach(async(user:any)=>{
                  (payloadForDMSFolderPrivacy as any).User=user.value;
                  (payloadForDMSFolderPrivacy as any).UserID=user.userId;
                  console.log("Payload for DMSFolderPrivacy",payloadForDMSFolderPrivacy);
                  const addedItem = await sp.web.lists.getByTitle("DMSFolderPrivacy").items.add(payloadForDMSFolderPrivacy);
                  console.log("Item added to the list DMSFolderPrivacy after selected value ",addedItem);
                })
                
            } catch (error) {
                console.log("error adding data to the DMSFolderPrivacy",error);
            }
            
        })

        // Add permission directly to the library/folder 
        try {
          const { web } = await sp.site.openWebById(`${OthProps.SiteID}`);
          let securableObject: any;

          if (OthProps.FolderName !== "null") {
              // For folder
              // alert('add permission to folder')
              const folder =await web.getFolderByServerRelativePath(`${OthProps.FolderPath}`).getItem();
              // securableObject=await folder.listItemAllFields.select("RoleAssignments").expand("RoleAssignments")();
              securableObject=folder;
              const itemData = await folder.select("HasUniqueRoleAssignments")();
              const breaKRole=itemData.HasUniqueRoleAssignments;
              if (!breaKRole) {
                // Break role inheritance, keeping current permissions
                await folder.breakRoleInheritance(true);
                console.log("Inheritance broken, retaining previous permissions.");
              }
              console.log("securableObject",securableObject);
          } else {
            // alert('add permission to doclib')
              // For document library
              securableObject =await web.lists.getByTitle(`${OthProps.DocumentLibraryName}`);
              console.log("securableObject",securableObject);
              // Break inheritance if needed (optional)
              const hasUniquePermissions = await securableObject.hasUniqueRoleAssignments;
              if (!hasUniquePermissions) {
                  await securableObject.breakRoleInheritance(true); // First `true` copies permissions, second `true` clears unique assignments
              }
          }
          // t222_Admin
          // const roleDefinition = await web.roleDefinitions.getByName('Edit')();
          // const roleDefinitionId = roleDefinition.Id;
          // const group = await sp.web.siteGroups.getByName('t222_Admin')();
          // await securableObject.roleAssignments.add(group.Id, roleDefinitionId);
          // Iterate through filteredArray and add role assignments
          rowsForPermission.forEach(async(row:any)=>{
              try {
                        // Get the role definition for the specified role
                        const roleDefinition = await web.roleDefinitions.getByName(row.selectedPermission.value)();
                        const roleDefinitionId = roleDefinition.Id;
                        console.log("roleDefinition",roleDefinition);
                        console.log("roleDefinitionId",roleDefinitionId);
                        row.selectedUserForPermission.forEach(async(userOrGroup:any)=>{
                          const principalId = userOrGroup.userId;
                          await securableObject.roleAssignments.add(principalId, roleDefinitionId);
                          console.log(`Adding ${userOrGroup.value} (${principalId}) with ${row.selectedPermission.value} permissions`);
                        })
              }catch (error){
                    console.log("Error in Adding permission to the document/folder directly inside for loop",error)
              }
            })
        } catch (error) {
          console.log("Error in Adding permission to the document/folder directly",error)
        }

     
        defaultValue("Create")
        setRowsForPermission([{ id: 0, selectedUserForPermission: [],selectedPermission:"" }]);
        Swal.fire('Added','Users Added Successfully','success');
    } catch (error) {
        console.log("error creating data inside the handle create function",error);
    }

  //   try {
  //     const { web } = await sp.site.openWebById(`${OthProps.SiteID}`);

  //     let securableObject: any;

  //     if (OthProps.FolderName !== "null") {
  //         // For folder
  //         const folder =await web.getFolderByServerRelativePath(`${OthProps.FolderPath}`);
  //         securableObject=folder.listItemAllFields();
  //         console.log("securableObject",securableObject);
  //     } else {
  //         // For document library
  //         securableObject =await web.lists.getByTitle(`${OthProps.DocumentLibraryName}`);
  //     }
      
  //     // Break inheritance if needed (optional)
  //     const hasUniquePermissions = await securableObject.hasUniqueRoleAssignments;
  //     if (!hasUniquePermissions) {
  //         await securableObject.breakRoleInheritance(true); // First `true` copies permissions, second `true` clears unique assignments
  //     }

      

  //     // Iterate through filteredArray and add role assignments
  //     rowsForPermission.forEach(async(row:any)=>{
  //               try {
  //                   // Get the role definition for the specified role
  //                   const roleDefinition = await web.roleDefinitions.getByName(row.selectedPermission.value)();
  //                   const roleDefinitionId = roleDefinition.Id;
  //                   console.log("roleDefinition",roleDefinition);
  //                   console.log("roleDefinitionId",roleDefinitionId);
  //                   row.selectedUserForPermission.forEach(async(userOrGroup:any)=>{
  //                     const principalId = userOrGroup.userId;
  //                     await securableObject.roleAssignments.add(principalId, roleDefinitionId);
  //                     console.log(`Adding ${userOrGroup.value} (${principalId}) with ${row.selectedPermission.value} permissions`);
  //                   })
                    
  //               } catch (error) {
  //                   console.log("error adding data to the folder/document library",error);
  //               }
                
  //           })

  //     // console.log("Permissions added successfully!");
  // } catch (error) {
  //     console.error("Error adding permissions:", error);
  // }

  

  }

  const handleDeleteUser=async(userId:any,itemId:number,permission:string)=>{
    console.log("userId",userId);
    console.log("itemId",itemId);
    // remove the user from list 
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, delete it!"
    }).then(async(result) => {
      if (result.isConfirmed) {
        try {
          await sp.web.lists.getByTitle(`DMSFolderPrivacy`).items.getById(itemId).delete();
          console.log(`file has been deleted successfully.`);
          defaultValue("Delete")
        } catch (error) {
          console.log("error in deleting the user from the dmsfolderprivacy",error)
        }
        // remove the user directly from library/folder
        try {
          const { web } = await sp.site.openWebById(`${OthProps.SiteID}`);
          let securableObject: any;
    
          if (OthProps.FolderName !== "null") {
              // For folder
              const folder =await web.getFolderByServerRelativePath(`${OthProps.FolderPath}`).getItem();
              securableObject=folder;
              const itemData = await folder.select("HasUniqueRoleAssignments")();
              const breaKRole=itemData.HasUniqueRoleAssignments;
              if (!breaKRole) {
                // Break role inheritance, keeping current permissions
                await folder.breakRoleInheritance(true);
                console.log("Inheritance broken, retaining previous permissions.");
              }
              console.log("securableObject",securableObject);
          } else {
              // For document library
              securableObject =await web.lists.getByTitle(`${OthProps.DocumentLibraryName}`);
              console.log("securableObject",securableObject);
              // Break inheritance if needed (optional)
              const hasUniquePermissions = await securableObject.hasUniqueRoleAssignments;
              if (!hasUniquePermissions) {
                  await securableObject.breakRoleInheritance(true); // First `true` copies permissions, second `true` clears unique assignments
              }
          }
          // Get the role definition for the specified role
          const roleDefinition = await web.roleDefinitions.getByName(permission)();
          const roleDefinitionId = roleDefinition.Id;
          await securableObject.roleAssignments.remove(userId,roleDefinitionId);
        } catch (error) {
          console.log("Error in deleting permission to the document/folder directly",error)
        }
        Swal.fire({
          title: "Deleted!",
          text: "User removed Successfully.",
          icon: "success"
        });
      }
    });
    // try {
    //   await sp.web.lists.getByTitle(`DMSFolderPrivacy`).items.getById(itemId).delete();
    //   console.log(`file has been deleted successfully.`);
    //   defaultValue("Delete")
    // } catch (error) {
    //   console.log("error in deleting the user from the dmsfolderprivacy",error)
    // }
    // // remove the user directly from library/folder
    // try {
    //   const { web } = await sp.site.openWebById(`${OthProps.SiteID}`);
    //   let securableObject: any;

    //   if (OthProps.FolderName !== "null") {
    //       // For folder
    //       const folder =await web.getFolderByServerRelativePath(`${OthProps.FolderPath}`).getItem();
    //       securableObject=folder;
    //       const itemData = await folder.select("HasUniqueRoleAssignments")();
    //       const breaKRole=itemData.HasUniqueRoleAssignments;
    //       if (!breaKRole) {
    //         // Break role inheritance, keeping current permissions
    //         await folder.breakRoleInheritance(true);
    //         console.log("Inheritance broken, retaining previous permissions.");
    //       }
    //       console.log("securableObject",securableObject);
    //   } else {
    //       // For document library
    //       securableObject =await web.lists.getByTitle(`${OthProps.DocumentLibraryName}`);
    //       console.log("securableObject",securableObject);
    //       // Break inheritance if needed (optional)
    //       const hasUniquePermissions = await securableObject.hasUniqueRoleAssignments;
    //       if (!hasUniquePermissions) {
    //           await securableObject.breakRoleInheritance(true); // First `true` copies permissions, second `true` clears unique assignments
    //       }
    //   }
    //   // Get the role definition for the specified role
    //   const roleDefinition = await web.roleDefinitions.getByName(permission)();
    //   const roleDefinitionId = roleDefinition.Id;
    //   await securableObject.roleAssignments.remove(userId,roleDefinitionId);
    // } catch (error) {
    //   console.log("Error in deleting permission to the document/folder directly",error)
    // }
  }

   // Code for filter and search start
   const [filters, setFilters] = React.useState({
    SNo: '',
    Title : '',
    // Title: '',
    CurrentUser: '',
    Modified: '',
    Status: '',

    SubmittedDate: ''
  });
  const [sortConfig, setSortConfig] = React.useState({ key: '', direction: 'ascending' });

  const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    setFilters({
      ...filters,
      [field]: e.target.value,
    });
    console.log(filters , "filters filters")
  };

  const handleSortChange = (key: string) => {
    let direction = 'ascending';
    if (sortConfig.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };
 
  const applyFiltersAndSorting = (data: any[]) => {
    const filteredData = data.filter((item, index) => {
      return (
        (filters.SNo === '' || String(index + 1).includes(filters.SNo)) &&
        (filters.Title === '' || 
          (item.Title && item.Title.toLowerCase().includes(filters.Title.toLowerCase()))) &&
        (filters.CurrentUser === '' || 
          (item.Author.Title && item.Author.Title.toLowerCase().includes(filters.CurrentUser.toLowerCase()))) &&
        (filters.Modified === '' || 
          (item.Editor.Title && item.Editor.Title.toLowerCase().includes(filters.Modified.toLowerCase()))) &&
        (filters.SubmittedDate === '' || 
          (item.Status && item.Status.toLowerCase().includes(filters.SubmittedDate.toLowerCase())))
      );
    });
  
    const naturalSort = (a: any, b: any) => {
      return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
    };
  
    const sortedData = filteredData.sort((a, b) => {
      if (sortConfig.key === 'SNo') {
        const aIndex = data.indexOf(a);
        const bIndex = data.indexOf(b);
        return sortConfig.direction === 'ascending' ? aIndex - bIndex : bIndex - aIndex;
      } else if (sortConfig.key) {
        const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : '';
        const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : '';
        return sortConfig.direction === 'ascending' ? naturalSort(aValue, bValue) : naturalSort(bValue, aValue);
      }
      return 0;
    });
  
    return sortedData;
  };
  
  const filteredDivisionData=applyFiltersAndSorting(tableData);
  // end

  // Add pagination start
  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;
  const totalPages = Math.ceil(filteredDivisionData.length / itemsPerPage);
  
  const handlePageChange = (pageNumber: any) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };
  
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = filteredDivisionData.slice(startIndex, endIndex);

  interface PaginationProps{
    currentPage: number;
    totalPages: any;
    handlePageChange: any;
  }
  const Pagination = ( { currentPage, totalPages, handlePageChange }: PaginationProps) => {
    const pageLimit = 5; // Number of visible page items
  
    // Determine the start and end page based on the current page and total pages
    const startPage = Math.max(1, currentPage - Math.floor(pageLimit / 2));
    const endPage = Math.min(totalPages, startPage + pageLimit - 1);
  
    // Adjust start page if it's too close to the end
    const adjustedStartPage = Math.max(1, Math.min(startPage, totalPages - pageLimit + 1));
  
    // Create an array for the visible page numbers
    const visiblePages = Array.from(
      { length: Math.min(pageLimit, totalPages) },
      (_, index) => adjustedStartPage + index
    );
  
    return (
      <nav className="pagination-container">
        <ul className="pagination">
          {/* Previous Button */}
          <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
            <a
              className="page-link PreviousPage"
              onClick={() => handlePageChange(currentPage - 1)}
              aria-label="Previous"
            >
              «
            </a>
          </li>
  
          {/* Render visible page numbers */}
          {visiblePages.map((pageNumber) => (
            <li
              key={pageNumber}
              className={`page-item ${currentPage === pageNumber ? 'active' : ''}`}
            >
              <a className="page-link" onClick={() => handlePageChange(pageNumber)}>
                {pageNumber}
              </a>
            </li>
          ))}
  
          {/* Next Button */}
          <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
            <a
              className="page-link NextPage"
              onClick={() => handlePageChange(currentPage + 1)}
              aria-label="Next"
            >
              »
            </a>
          </li>
        </ul>
      </nav>
    );
  };
  // End


  return (
    <div className="container mt-0 second">
            <div className="modal show d-block" tabIndex={-1}>
                    <div className="modal-dialog">
                        <div className="modal-content" style={{
                                    width:"140%",
                                    padding:"0px"
                                    
                        }}>


                    {toggelPermission === "Yes" ? 
                    (
                      <div>           
                            
                            <div className="" style={{
                              height: "auto",
                              width:"auto",
                          }}
                          >
                          
                         
                          <div  className='row'>
                            <div className='col-sm-8 w90'>
                            <h5 className="mb-3 " style={{
                              display:"block", margin:'inherit'
                             
                          }}>
                              Manage Permission
                              
                          </h5>
                          <div  className='font-12 text-muted'>
                            {path}
                          </div>

                            </div>
                            <div className='col-sm-4'>
                            <div  className="mb-0">
                            <div style={{height:'20px'}} className="col-12 d-flex justify-content-end">
                              <a onClick={handleAddRow}>
                                <img className="bi bi-plus newl" src={require("../assets/plus.png")} alt="add" style={{ width: "50px", position:'relative', height: "50px" , top:'0px'}} />
                              </a>
                            </div>
                          </div>
                            </div>
                            <div style={{borderBottom:'1px solid #ccc', marginBottom:'15px', height:'15px', float:'left', width:'100%',  paddingBottom:'10px'}}></div>
                          </div>
                        
                          {rowsForPermission.map((row)=>(
                          <div style={{clear:'both'}} className="row  approvalheirarcystyle" key={row.id}>
                                  <div className="col-12 col-md-6 mb-2">
                                      <Select
                                          value={row.selectedUserForPermission}
                                          isMulti
                                          options={users}
                                          onChange={(selected: any) =>
                                          handleUserSelectForPermission(selected,row.id)
                                          }
                                          placeholder="Enter names or email addresses..."
                                          noOptionsMessage={() => "No User Found..."}
                                      />
                                      {errorsForPermissionSelection[row.id]?.userSelect && (
                                    <span className="text-danger">{errorsForPermissionSelection[row.id].userSelect}</span>
                                  )}
                                  
                                    {errorsForPermissionSelection[row.id]?.duplicate && (
                                      <span className="text-danger">{errorsForPermissionSelection[row.id].duplicate}</span>
                                    )}
                                  </div>
                                  <div className="col-12 col-md-4 mb-2" style={{
                               
                                  }}>
                                      <Select
                                          value={row.selectedPermission || null}
                                          options={permissionArray}
                                          onChange={(selected: any) =>
                                          handlePermissionSelect(selected,row.id)
                                          }
                                          placeholder="Select Permission..."
                                          noOptionsMessage={() => "No Such Permission Find"}
                                      />
                                      {errorsForPermissionSelection[row.id]?.permissionSelect && (
                                    <span className="text-danger">{errorsForPermissionSelection[row.id].permissionSelect}</span>
                                  )}
                                  </div>
                                  {/* {row.id === 0 ? null : ( */}
                                    <div className="col-12 mb-2 col-md-2 d-flex align-items-end">
                                      <a onClick={(e) => handleRemoveRow(row.id, e)} style={{ width: "50px",    height: "50px", cursor: "pointer" }}>
                                        <img className="fas fa-trash" src={require("../assets/del.png")} alt="delete" />
                                      </a>
                                    </div>
                                  {/* )} */}
                          </div>
                          ))}
                      {/* Tabular view */}
                      <div style={{
                         
                        }}>

                       
                        <div>
                        <table className='mtbalenew'>

                            <thead>
                            <tr>
                                <th style={{minWidth:'55px', maxWidth:'55px'}}>S.No.</th>
                                <th>User/Groups</th>
                                {/* <th className={styles.tabledept}>Email</th> */}
                                <th >Permisson</th>
                                <th style={{minWidth:'75px', maxWidth:'75px'}}>Action</th>
                            </tr>
                            </thead>
                            <tbody>
                            {currentData.map((item:any, index:any) => (
                                <React.Fragment key={item.Id}>
                                <tr >
                                    <td style={{minWidth:'55px', maxWidth:'55px'}} >
                                  <span style={{marginLeft:'20px'}}  className='indexdesign'>{index + 1}</span>   
                                    </td>
                                    <td >
                                    {item.value || ''}
                                    </td>
                                    {/* <td className={styles.tabledept}>
                                    {item?.email || ''}
                                    </td> */}
                                    <td >
                                    {item.Permission || ''}
                                    </td>
                                    <td style={{minWidth:'75px', maxWidth:'75px'}}>
                                        <img
                                            className={styles.deleteicon}
                                            src={require("../assets/del.png")}
                                            alt="Delete"
                                            style={{
                                              height:"25px"
                                            }}
                                            onClick={(event)=>{
                                                handleDeleteUser(item.userId,item.itemId,item.Permission)
                                            }}
                                        />
                                     </td>
                                </tr>
                                </React.Fragment>
                            ))}
                        </tbody>
                        </table>
                        <div className='row'>
                          <div className='col-md-8 newpag'>
                          <Pagination
                          currentPage={currentPage}
                          totalPages={totalPages}
                          handlePageChange={handlePageChange}
  
                        />

                            </div>
                            <div className='col-md-4'>
                            <div style={{textAlign:'right'}} className="text-right pb-0">
                                    <button type="button" className="btn btn-primary me-2" 
                                    onClick={handleCreate}
                                    >
                                    Create
                                    </button>
                                    <button type="button" className="btn btn-secondary" 
                                    //   onClick={toggleModal}
                                    onClick={onReturnToMain}
                                    >
                                        Cancel{" "}
                                    </button>
                            </div>

                              </div>

                          </div>
                        
                       
                        </div>
                      </div>
                          </div>   
                        {/* <div className="card cardborder marginleftcard" style={{
                                height: "auto",
                                width:"auto",
                            }}>
                            <h5 className="mb-3 " style={{
                                display:"block",
                                width:"200px"
                            }}>
                                <strong>Manage Permission</strong>
                            </h5>
                            <div className="row mb-3 approvalheirarcystyle">
                                    <div className="col-12 col-md-6">
                                        <Select
                                            value={defaultUser}
                                            isMulti
                                            options={users}
                                            onChange={(selected: any) =>
                                            handleUserSelectForPermission(selected)
                                            }
                                            placeholder="Enter names or email addresses..."
                                            noOptionsMessage={() => "No User Found..."}
                                        />
                                    </div>
                                    <div className="col-12 col-md-6" style={{
                                    width:"auto"
                                    }}>
                                        <Select
                                            value={selectedPermission}
                                            options={permissionArray}
                                            onChange={(selected: any) =>
                                            handlePermissionSelect(selected)
                                            }
                                            placeholder="Select Permission"
                                            noOptionsMessage={() => "No Such Permission Find"}
                                        />
                                    </div>

                            </div>
                            </div> */}
                           
                      </div>
                    )
                    :null
                    }             
                    {toggelPermission === "No"  ?
                    (
                    <div>
                       <h6>"This document library is public. Would you like to Make it private?"</h6>
                        <div className='setApprover' style={{
                          display:"flex",
                          marginLeft:"280px",
                          gap:"10px"
                        }}>
                        <button type="button" className="btn btn-primary" onClick={handlesetTogglePermission}>Set Permission</button>
                        <button type="button" className="btn btn-secondary" 
                          onClick={onReturnToMain}
                        >
                          Cancel{" "}
                        </button>
                      </div>
                  </div>
                )
                : null
                }
                        </div>
                    </div>
            </div>
</div>
  )
}

export default ManageFolderPermission