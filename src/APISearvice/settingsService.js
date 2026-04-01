
export const getSettingAPI = async (_sp) => {
    let arr =[]
   await _sp.web.lists.getByTitle("Settings").items.select("Title,ID,ImageIcon,LinkUrl")()
    .then((res) => {
        console.log(res);
        arr= res;
    })
    .catch((error) => {
        console.error("Error fetching data: ", error);
    });
    return arr;
}