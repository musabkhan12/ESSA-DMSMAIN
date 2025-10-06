
import Swal from 'sweetalert2';
export const getNews = async (_sp) => {
  let arr = []
     let str ="News"
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

export const uploadFileToLibrary = async (file, sp, docLib) => {
  debugger
  let arrFIleData = [];
  try {
    const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file,

      // const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(
      // file.name,
      // file,
      (progress, data) => {
        console.log(progress, data);
      },
      true
    );

    const item = await sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("ID", "AuthorId", "Modified")
    console.log(item.Id, 'itemitem');
    let arr = {
      ID: item.Id,
      Createdby: item.AuthorId,
      Modified: item.Modified,
      fileUrl: result.data.ServerRelativeUrl
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
  var arr ={};
  sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
    console.log(`progress`, data);
  }, true).then(async result => {
    console.log(result, 'result')
  })
  //  const siteUrl=  await getUrl(sp)
  const img = {
    "type": "thumbnail",
    "fileName": file.name,
    "serverUrl": siteUrl,
    "fieldName": "AnnouncementandNewsBannerImage",
    "serverRelativeUrl": '/Shared%20Documents/' + file.name
  };
  arr=img
  return arr;
};

export const getUrl = async (sp) => {
  const url = await sp.web.currentUser.getContextInfo();
  console.log(url.WebFullUrl, 'res');
  return url.WebFullUrl
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
    Swal.fire('Item added successfully', '', 'success');

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
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.getById(id)()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      const parsedValues = {
        Title: res.Title,
        description: res.Description,
        overview: res.Overview,
        IsActive: res.IsActive,
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