import Swal from 'sweetalert2';
export const getAnncouncement = async (_sp) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.select("*,AnnouncementandNewsTypeMaster/Id,AnnouncementandNewsTypeMaster/TypeMaster,Category/Id,Category/Category,Author/ID,Author/Title").expand("AnnouncementandNewsTypeMaster,Category,Author").filter(`AnnouncementandNewsTypeMaster/TypeMaster eq '${str}'`).getAll()
    .then((res) => {
      console.log(res);


      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}
export const getAnncouncementNewsCategory = async (_sp) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGAnnouncementsandNewsCategory").items.getAll()
    .then((res) => {
      console.log(res);


      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}

export const uploadFileToLibrary = async (file, sp, docLib) => {
  debugger
  let arrFIleData = [];
  let fileSize = 0
  try {
    const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file,

      // const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(
      // file.name,
      // file,
      (progress, data) => {
        console.log(progress, data);
        fileSize = progress.fileSize
      },
      true
    );
    if (result.data != undefined) {
      const item = await sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("*", "ID", "AuthorId", "Modified")
      console.log(item.Id, 'itemitem');
      let arr = {
        ID: item.Id,
        Createdby: item.AuthorId,
        Modified: item.Modified,
        fileUrl: result.data.ServerRelativeUrl,
        fileSize: fileSize,
        fileType: file.type,
        fileName: file.name,
      }
      arrFIleData.push(arr)
      console.log(arrFIleData);

      return arrFIleData;
    }

  } catch (error) {
    console.error("Error uploading file:", error);
    return null; // Or handle error differently
  }
};

// export const uploadFile = async (file, sp, docLib, siteUrl) => {
//   var arr = {};
//   sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
//     console.log(`progress`, data);
//   }, true).then(async result => {
//     console.log(result, 'result')
//   })
//   //  const siteUrl=  await getUrl(sp)
//   const img = {
//     "type": "thumbnail",
//     "fileName": file.name,
//     "serverUrl": siteUrl,
//     "fieldName": "AnnouncementandNewsBannerImage",
//     "serverRelativeUrl": '/SiteAssets/' + file.name,
//     // "thumbnailRenderer":siteUrl+'/Shared%20Documents/Lists/a6432673-c28a-4d23-b736-7cc02ca46a53/' + file.name
//   };
//   arr = img
//   return arr;
// };

export const uploadFile = async (file, sp, docLib, siteUrl) => {
  let arr = {};
  debugger
  const uploadResult = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
    console.log(`progress`, data);
  }, true);

  const fileUrl = uploadResult.data.ServerRelativeUrl;

  const imgMetadata = {
    "__metadata": { "type": "SP.FieldUrlValue" },
    "Description": file.name,
    "Url": `${siteUrl}${fileUrl}`
  };

  // await sp.web.lists.getByTitle(docLib).items.getById(uploadResult.data.UniqueId).update({
  //   "AnnouncementandNewsBannerImage": imgMetadata
  // });
  arr = {
    "type": "thumbnail",
    "fileName": file.name,
    "serverUrl": siteUrl,
    "fieldName": "AnnouncementandNewsBannerImage",
    "serverRelativeUrl": fileUrl
  };
  return arr;
};
export const getUrl = async (sp) => {

  const url = await sp.web.currentUser.getContextInfo();
  console.log(url.WebFullUrl, 'WebFullUrl');
  return url.WebFullUrl
}
const handleFileChange = async (event) => {
  const file = event.target.files[0];
  if (file) {
    try {
      const folderUrl = `/sites//${docLib}`; // Replace with your folder URL
      const fileName = file.name;

      const fileBlob = new Blob([file], { type: file.type });

      await sp.web.getFolderByServerRelativeUrl(folderUrl)
        .files.add(fileName, fileBlob)
        .then((fileAdded) => {
          console.log('File uploaded successfully:', fileAdded);
        });
    } catch (error) {
      console.error('Error uploading file:', error);
    }
  }
};

export const AddAnncouncementanNews = async (_sp) => {
  let arr = []
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.add
    .then((res) => {
      console.log(res);
      arr = res;
    })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}
export const addItem = async (itemData, _sp) => {
  debugger
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGAnnouncementAndNews').items.add(itemData);
    debugger
    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.error('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};
export const updateItem = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGAnnouncementAndNews').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.error('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};
export const DeleteAnnouncementAPI = async (_sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGAnnouncementAndNews').items.getById(id).delete();
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.error('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
}
export const getAnncouncementByID = async (_sp, id) => {
  debugger
  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.getById(id).select("*,Category/ID,Category/Category,Entity/ID,Entity/Entity,AnnouncementandNewsTypeMaster/ID,AnnouncementandNewsTypeMaster/TypeMaster").expand("Category,Entity,AnnouncementandNewsTypeMaster")()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      const bannerimgobject = res.AnnouncementandNewsBannerImage != "{}" && JSON.parse(res.AnnouncementandNewsBannerImage)
      console.log(bannerimgobject[0], 'bannerimgobject');

      bannerimg.push(bannerimgobject);
      const parsedValues = {
        Title: res.Title != undefined ? res.Title : "",
        description: res.Description != undefined ? res.Description : "",
        overview: res.Overview != undefined ? res.Overview : "",
        IsActive: res.IsActive,
        ID: res.ID,
        BannerImage: bannerimg,
        TypeMaster: res?.AnnouncementandNewsTypeMaster?.ID != undefined ? res.AnnouncementandNewsTypeMaster?.ID : "",
        Category: res.Category?.ID != undefined ? res.Category?.ID : "",
        Entity: res.Entity?.ID != undefined ? res.Entity?.ID : "",
        FeaturedAnnouncement: res.FeaturedAnnouncement,
        AnnouncementAndNewsGallaryJSON: res.AnnouncementAndNewsGallaryJSON != null ? JSON.parse(res.AnnouncementAndNewsGallaryJSON) : "",
        AnnouncementAndNewsDocsJSON: res.AnnouncementAndNewsDocsJSON != null ? JSON.parse(res.AnnouncementAndNewsDocsJSON) : "",
        AnnouncementAndNewsGallaryId: res.AnnouncementAndNewsGallaryId,
        AnnouncementsAndNewsDocsId: res.AnnouncementsAndNewsDocsId
        // other fields as needed
      };

      arr.push(parsedValues)
    })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}
export const getAnnouncementDetailsById=async (_sp,idNum)=>
{
  let arr = []
  let arr1 = []

  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.getById(idNum)()
  .then((res) => {
    // arr=res;
    arr1.push(res)
    arr=arr1
  }).catch((error) => {
    console.log("Error fetching data: ", error);
  });
  return arr;
}