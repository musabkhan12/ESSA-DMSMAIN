import Swal from 'sweetalert2';
export const getMedia = async (_sp) => {
  let arr = []
     let str ="Announcements"
  await _sp.web.lists.getByTitle("ARGMediaGallery").items.select("*,EntityMaster/ID,EntityMaster/Entity").expand("EntityMaster").getAll()
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
  let fileSize=0
  try {
    const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file,

      // const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(
      // file.name,
      // file,
      (progress, data) => {
        console.log(progress, data);
        fileSize=progress.fileSize
      },
      true
    );

    const item = await sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("*","ID", "AuthorId", "Modified")
    console.log(item.Id, 'itemitem');
    let arr = {
      ID: item.Id,
      Createdby: item.AuthorId,
      Modified: item.Modified,
      fileUrl: result.data.ServerRelativeUrl,
      fileSize:fileSize,
      fileType:file.type,
      fileName:file.name,
    }
    arrFIleData.push(arr)
    console.log(arrFIleData);

    return arrFIleData;
  } catch (error) {
    console.error("Error uploading file:", error);
    return null; // Or handle error differently
  }
};

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

export const getUrl = async (sp,siteUrl) => {
  let srt = siteUrl;
let dynamicPart = "/sites/";
let UrlArr=[];
try {
    let index = srt.lastIndexOf(dynamicPart); // Find the last occurrence of "/sites/"

    if (index !== -1) {
        let endIndex = srt.indexOf("/", index + dynamicPart.length) !== -1
            ? srt.indexOf("/", index + dynamicPart.length)
            : srt.length;

        let updatedStr = srt.slice(0, index) + srt.slice(endIndex);
        console.log(updatedStr,'updatedStr'); 
        const url = await sp.web.currentUser.getContextInfo();
        console.log(url, 'res');
      
       let UrlArr1=
        {
          DomainUrl:updatedStr,
          WebFullUrl:url.WebFullUrl
      
        }
        UrlArr.push(UrlArr1)
    } else {
        console.log("Pattern not found. No replacement was made.");
    }
} catch (error) {
    console.error("An error occurred:", error.message);
}

  return UrlArr
}
const handleFileChange = async (event) => {
  const file = event.target.files[0];
  if (file) {
    try {
      const folderUrl = `/sites/AlRostmani/${docLib}`; // Replace with your folder URL
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
  await _sp.web.lists.getByTitle("ARGMediaGallery").items.add
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
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGMediaGallery').items.add(itemData);
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
    const newItem = await _sp.web.lists.getByTitle('ARGMediaGallery').items.getById(id).update(itemData);
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
export const DeletemediaAPI = async (_sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGMediaGallery').items.getById(id).delete();
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
export const getMediaByID = async (_sp, id) => {
  debugger
  let arr = []
  let arrs = []
  await _sp.web.lists.getByTitle("ARGMediaGallery").items.getById(id).select("*,EntityMaster/Id,EntityMaster/Entity").expand("EntityMaster")()
      .then((res) => {
          console.log(res, ' let arrs=[]');
               const parsedValues= {
                Title: res.Title,
                ID:res.ID,
                entity: res.EntityMaster?.Id,
                Image: res.Image,
                MediaGalleriesId:res?.MediaGalleriesId,
                MediaGalleryJSON:res?.MediaGalleryJSON
                // other fields as needed
            };

           arr.push(parsedValues)
           arrs = arr
            console.log(arrs, 'arr');
      })
      .catch((error) => {
          console.error("Error fetching data: ", error);
      });
  console.log(arrs, 'arr');
  return arrs;
}