export const fetchMediaGallerycategory = async (_sp) => {
  let arr = []
 
     await _sp.web.lists.getByTitle("ARGMediaGalleryCategory").items.getAll().then((res) => {
      console.log(res);
   
      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}

export const fetchMediaGallerydata = async (_sp) => {
    let arr = []
   
       await _sp.web.lists.getByTitle("ARGMediaGallery").items.select("*,MediaGalleryCategory/Id,MediaGalleryCategory/CategoryName").expand("MediaGalleryCategory").getAll().then((res) => {
        console.log("response-->>>",res);
     
        //res.filter(x=>x.Category?.Category==str)
        arr = res;
      })
      .catch((error) => {
        console.error("Error fetching data: ", error);
      });
    return arr;
  }

  export const fetchMediaGalleryInsideData = async (_sp, Id) => {
    let arr = [];
  
    // Fetch data from the SharePoint list
    await _sp.web.lists
      .getByTitle("ARGMediaGallery")
      .items
      .select("*,MediaGalleryCategory/Id,MediaGalleryCategory/CategoryName")
      .expand("MediaGalleryCategory")
      .filter(`Id eq ${Id}`) // Filter based on the main Id
      .getAll()
      .then((res) => {
        if (res.length > 0) {
          const item = res[0]; // Assuming the result contains one item
          console.log("Main Item:", item);
  
          // Parse the MediaGalleryJSON field
          const mediaGalleryJSON = JSON.parse(item.MediaGalleryJSON);
  
          // Now, filter or process the parsed data (for example, filter by file type or other criteria)
          const filteredData = mediaGalleryJSON.filter(media => media.fileType === "image/png"); // Example filter
  
          console.log("Filtered MediaGalleryJSON Data:", filteredData);
          arr = filteredData; // Store the filtered data
        }
      })
      .catch((error) => {
        console.error("Error fetching data: ", error);
      });
  
    return arr;
  };
  

  
  export const fetchDynamicdata = async (_sp) => {
    let arr = []
   
       await _sp.web.lists.getByTitle("DynamicBanners").items.getAll().then((res) => {
        console.log(res);
     
        //res.filter(x=>x.Category?.Category==str)
        arr = res;
      })
      .catch((error) => {
        console.error("Error fetching data: ", error);
      });
    return arr;
  }