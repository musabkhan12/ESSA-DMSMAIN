import Swal from 'sweetalert2';
export const getDynamicBanner = async (_sp) => {
    let arr = []
    await _sp.web.lists.getByTitle("DynamicBanners").items.getAll()
        .then((res) => {
            console.log(res);
            arr = res;
        })
        .catch((error) => {
            console.error("Error fetching data: ", error);
        });
    return arr;
}
export const DeleteBannerAPI = async (_sp, id) => {
    let resultArr = []
    try {
        const newItem = await _sp.web.lists.getByTitle('DynamicBanners').items.getById(id).delete();
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

export const getBannerByID = async (_sp, id) => {
    debugger
    let arr = []
    let arrs = []
    await _sp.web.lists.getByTitle("DynamicBanners").items.getById(id).select("ID","Title","URL","IsImage","BannerImage","Description","BannerImageJSON")()
        .then((res) => {
            console.log(res, ' let arrs=[]');
                 const parsedValues= {
                  title: res.Title,
                  description: res.Description,
                  BannerImage: res.BannerImage,
                  URL: res.URL,
                  IsImage:res.IsImage==true?"on":"off",
                  IsVedio:res.IsImage==true?"on":"off",
                  ID:res.ID
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

export const addItem = async (itemData, _sp) => {
    debugger
    let resultArr = []
    try {
        const newItem = await _sp.web.lists.getByTitle('DynamicBanners').items.add(itemData);
        debugger
        console.log('Item added successfully:', newItem);
        Swal.fire('Item added successfully', '', 'success');

        resultArr = newItem
        // Perform any necessary actions after successful addition
    } catch (error) {
        console.error('Error adding item:', error);
        Swal.fire(' Cancelled', '', 'error')
        // Handle errors appropriately
        resultArr = null
    }
    return resultArr;
};
export const updateItem = async (itemData,_sp,id) => {
    let resultArr=[]
    try {
      const newItem = await _sp.web.lists.getByTitle('DynamicBanners').items.getById(id).update(itemData);
      Swal.fire('Item update successfully', '', 'success');
      resultArr=newItem
      // Perform any necessary actions after successful addition
    } catch (error) {
      console.error('Error adding item:', error);
      Swal.fire(' Cancelled', '', 'error')
      // Handle errors appropriately
      resultArr =null
    }
  return resultArr;
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
      "fieldName": "BannerImage",
      "serverRelativeUrl": fileUrl
    };
    return arr;
  };
  export const getUrl = async (sp) => {
    debugger
    const url = await sp.web.currentUser.getContextInfo();
    console.log(url.WebFullUrl, 'WebFullUrl');
    return url.WebFullUrl
  }