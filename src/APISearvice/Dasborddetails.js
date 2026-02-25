export const getAnncouncementone = async (_sp) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.select("*,AnnouncementandNewsTypeMaster/Id,AnnouncementandNewsTypeMaster/TypeMaster,Category/Id,Category/Category").expand("AnnouncementandNewsTypeMaster,Category").filter(`AnnouncementandNewsTypeMaster/TypeMaster eq '${str}'`).top(2)().then((res) => {
    console.log(res);


    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}

export const getNewsone = async (_sp) => {
  let arr = []
  let str = "News"
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.select("*,AnnouncementandNewsTypeMaster/Id,AnnouncementandNewsTypeMaster/TypeMaster,Category/Id,Category/Category").expand("AnnouncementandNewsTypeMaster,Category").filter(`AnnouncementandNewsTypeMaster/TypeMaster eq '${str}'`).top(2)().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}

export const fetchEventdataone = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("ARGEventMaster").items.top(4)().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.error("Error fetching data: ", error);
    });
  return arr;
}
export const fetchUserInformationList = async (sp) => {
  let arr = []
  try {
    const userList = await sp.web.lists.getByTitle("User Information List").items.select("ID", "Title", "EMail", "Department", "JobTitle", "Picture").filter("EMail ne null").top(5)();
    console.log(userList, 'userList');
    arr = userList
    // setUsersArr(userList);
  } catch (error) {
    console.error("Error fetching users:", error);
  }
  return arr;
};
