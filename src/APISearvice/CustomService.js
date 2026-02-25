import Swal from 'sweetalert2';

import { format, subDays } from 'date-fns';

export const getEntity = async (_sp) => {

  let arr = []

  await _sp.web.lists.getByTitle("ARGEntityMaster").items.select("ID,Entity").filter("Active eq 1")()

    .then((res) => {

      console.log(res);

      const newArray = res.map(({ ID, Entity }) => ({ id: ID, name: Entity }));

      console.log(newArray, 'newArray');


      arr = newArray;

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return arr;

}

export const getCategory = async (_sp, id) => {

  let arr = []

  await _sp.web.lists.getByTitle("ARGAnnouncementsandNewsCategory").items.select("ID,Category,AnnouncementandNewsTypeMaster/TypeMaster,AnnouncementandNewsTypeMaster/ID").expand("AnnouncementandNewsTypeMaster").filter(`(Active eq 1) and(AnnouncementandNewsTypeMaster/ID
 eq ${id})`)()

    .then((res) => {

      console.log(res);

      const newArray = res.map(({ ID, Category }) => ({ id: ID, name: Category }));

      console.log(newArray, 'newArray');


      arr = newArray;

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return arr;

}

export const getType = async (_sp) => {

  let arr = []

  await _sp.web.lists.getByTitle("AnnouncementandNewTypeMaster").items.select("ID,TypeMaster")()

    .then((res) => {

      console.log(res);

      const newArray = res.map(({ ID, TypeMaster }) => ({ id: ID, name: TypeMaster }));

      console.log(newArray, 'newArray');


      arr = newArray;

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return arr;

}


export const getChoiceFieldOption = async (_sp, listName, fieldName) => {

  let arr = []

  const field2 = await _sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldName)()

  console.log(field2, 'field2');


  arr = field2["Choices"]


  return arr;

}


export const getCurrentUser = async (_sp, siteUrl) => {

  let arr = []

  await _sp.web.currentUser()

    .then(async (res) => {

      console.log(res);


      arr = res;

      const ProfilePic = `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${res.Email}`

      //await getUserProfilePicture(res.Id,_sp)

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return arr;

}


export const getCurrentUserProfile = async (_sp, siteUrl) => {

  const ProfilePic = []

  await _sp.web.currentUser()

    .then(async (res) => {

      console.log(res);


      //   arr= res;

      ProfilePic = `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${res.Email}`

      //await getUserProfilePicture(res.Id,_sp)

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return ProfilePic;

}


export const getCurrentUserName = async (_sp) => {

  let arr = ""

  await _sp.web.currentUser()

    .then((res) => {

      console.log(res, "response of user email");


      arr = res["Title"];
    
      console.log(arr, "response of user email");
    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return arr;

}

export const getCurrentUserNameId = async (_sp) => {

  let arr = 0;

  await _sp.web.currentUser()

    .then((res) => {

      console.log(res);


      arr = res["Id"];


    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return arr;

}

export async function getUserProfilePicture(authorId, sp) {

  try {

    // Get user information by ID

    const user = await sp.web.getUserById(authorId)();


    if (user) {

      // The 'Picture' field holds the profile picture URL

      const profilePictureUrl = `${user.PictureUrl}`;

      console.log(profilePictureUrl, 'profilePictureUrl');


      return profilePictureUrl;

    }

    return null;

  } catch (error) {

    console.log("Error fetching profile picture:", error);

    return null;

  }

}


export const getCurrentUserProfileEmail = async (sp) => {

  const Email = ""

  await sp.web.currentUser()

    .then(async (res) => {

      console.log(res);


      //   arr= res;

      Email = res.Email;

      //await getUserProfilePicture(res.Id,_sp)

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return Email;

}

export const getCurrentUserProfileEmailForPeople = async (sp) => {

  const UserID = ""

  await sp.web.currentUser()

    .then(async (res) => {

      console.log(res);


      //   arr= res;

      // Email=  res.Email;

      const UserID = await sp.web.ensureUser(res.Email);


      //await getUserProfilePicture(res.Id,_sp)

    })

    .catch((error) => {

      console.log("Error fetching data: ", error);

    });

  return user.data.Id;

}


export const getFollow = async (sp) => {

  let arr = []

  try {

    const currentUser = await sp.web.currentUser();

    const followRecords = await sp.web.lists.getByTitle("ARGFollows").items

      .select("Followed/Title", "Followed/ID", "Followed/EMail", "ID").expand("Followed")

      .filter(`FollowedId eq ${currentUser.Id}`)();

    console.log(followRecords, 'followedRecords');

    arr = followRecords

    // Update the counts based on follow status

  } catch (error) {

    console.error("Error checking follow status:", error);

  }

  return arr;

}

export const getFollowing = async (sp) => {

  let arr = []

  try {

    const currentUser = await sp.web.currentUser();

    const followRecords = await sp.web.lists.getByTitle("ARGFollows").items

      .select("Follower/Title", "Follower/ID", "Follower/EMail", "ID").expand("Follower")

      .filter(`FollowerId eq ${currentUser.Id}`)()


    console.log(followRecords, 'followRecords');


    arr = followRecords

    // Update the counts based on follow status

  } catch (error) {

    console.error("Error checking follow status:", error);

  }

  return arr;

}

export const addActivityLeaderboard = async (sp, Action) => {

  

  const trimmedAction = Action.trim();

 console.log(Action,'Action');

 

  // Retrieve all items, and filter on the client side

  await sp.web.lists

    .getByTitle("ARGLeaderboardConfigure")

    .items.select("Actions", "Points")() // Adjust fields as needed

    .then(async (res) => {

      if (res.length > 0) {

        const matchingItem = res.find(item => item.Actions.trim() === trimmedAction);


        if (matchingItem) {

          await sp.web.lists

            .getByTitle("ARGLeaderboard")

            .items.add({

              Action: trimmedAction,

              Points: matchingItem.Points,

            })

            .then((res) => {

              console.log(res, 'ews');

            });

        }

      }

      // Filter results by trimming Actions field values


    });

};

// export const getLeader = async (sp) => {

//   let arr = [];


//   // Fetching the leaderboard data

//   await sp.web.lists

//     .getByTitle("ARGLeaderboard")

//     .items.select("*,Author/ID,Author/Title,Author/EMail")

//     .expand("Author")

//     .getAll()

//     .then((res) => {

//       console.log(res, "Fetched data");


//       // Storing the fetched data in the array

//       arr = res;

//     });


//   // Grouping and summing points by AuthorId

//   const sumPointsByUser = () => {

//     const userPoints = {};


//     // Loop through each item in the fetched results

//     arr.forEach((item) => {

//       const authorId = item.Author.ID; // Author ID

//       const points = item.Points; // Points for the action


//       // Sum points for each AuthorId

//       if (userPoints[authorId]) {

//         userPoints[authorId] += points;

//       } else {

//         userPoints[authorId] = points;

//       }

//     });


//     // Convert the object to an array of users with total points

//     return Object.keys(userPoints).map((authorId) => ({

//       AuthorId: authorId,

//       TotalPoints: userPoints[authorId],

//     }));

//   };


//   // Return the result

//   return sumPointsByUser();

// };

// export const getLeader = async (sp) => {

//   let arr = [];


 

//   await sp.web.lists

//     .getByTitle("ARGLeaderboard")

//     .items.select("*,Author/ID,Author/Title,Author/EMail,Author/Department")

//     .expand("Author")

//     .getAll()

//     .then((res) => {

//       console.log(res, "Fetched data");


//       arr = res;

//     });


//   const sumPointsByUser = async () => {

//     const userPoints = {};

//     const followRecords = await sp.web.lists.getByTitle("ARGLeaderboardRattingLevel").items

//       .getAll()


//     console.log(followRecords, 'ARGLeaderboardRattingLevel');


   

//     arr.forEach((item) => {

//       const authorId = item.Author.ID; 

//       const points = item.Points; 

//       const authorTitle = item.Author.Title; 

//       const authorEMail = item.Author.EMail; 

//       const authorDepartment = item.Author.Department; 


    

//       if (userPoints[authorId]) {

//         userPoints[authorId].TotalPoints += points;

//       } else {

//         userPoints[authorId] = {

//           TotalPoints: points,

//           AuthorTitle: authorTitle,

//           AuthorEMail: authorEMail,

//           AuthorDepartment: authorDepartment

//         };

//       }

//     });


//     for (let i = 0; i < followRecords.length; i++) {


//       return Object.keys(userPoints).map((authorId) => ({

//         AuthorId: authorId,

//         TotalPoints: userPoints[authorId].TotalPoints,

//         AuthorTitle: userPoints[authorId].AuthorTitle,

//         AuthorEMail: userPoints[authorId].AuthorEMail,

//         AuthorDepartment: userPoints[authorId].AuthorDepartment,

//         AutherRatting: userPoints[authorId].TotalPoints

//       }));

//     }

//   };


//   return sumPointsByUser();

// };



export const GetLeaderboardRattingLevel = async (sp) => {

  let arr = []

  try {

    const currentUser = await sp.web.currentUser();

    const followRecords = await sp.web.lists.getByTitle("ARGLeaderboardRattingLevel").items

      .getAll()


    console.log(ARGLeaderboardRattingLevel, 'ARGLeaderboardRattingLevel');


    arr = followRecords

    // Update the counts based on follow status

  } catch (error) {

    console.error("Error checking follow status:", error);

  }

  return arr;

}



export const getLeaderTop = async (sp) => {

  let arr = [];

  let RATINGSTHRESHOLDS = [];


  // Fetching the leaderboard data from SharePoint

  await sp.web.lists

    .getByTitle("ARGLeaderboard")

    .items.select("*,Author/ID,Author/Title,Author/EMail,Author/Department,Created")

    .expand("Author")

    .getAll()

    .then((res) => {

      console.log(res, "Fetched data");

      // Storing the fetched data in the array

      arr = res;

    });


  const followRecords = await sp.web.lists

    .getByTitle("ARGLeaderboardRattingLevel")

    .items.select("Ratting,Points")

    .getAll();


  console.log(followRecords, "ARGLeaderboardRattingLevel");


  // Dynamic threshold settings for rating

  followRecords.forEach((element) => {

    let arr1 = { minPoints: element.Points, rating: element.Ratting };

    RATINGSTHRESHOLDS.push(arr1);

  });


  // Grouping and summing points by AuthorId

  const sumPointsByUser = () => {

    const userPoints = {};


    // Get current date

    const currentDate = new Date();

    const currentMonth = currentDate.getMonth(); // 0-11 (January is 0)

    const currentYear = currentDate.getFullYear();


    // Filter records for the current month

    const currentMonthRecords = arr.filter((record) => {

      const recordDate = new Date(record.Created); // Assuming `Created` is the date field

      return (

        recordDate.getMonth() === currentMonth -1 &&

        recordDate.getFullYear() === currentYear

      );

    });


    if (currentMonthRecords.length === 0) {

      console.log("No leaderboard is available");

      //alert("No leaderboard is available");

      return []; // No data for the current month

    }


    // Loop through each item in the filtered results

    currentMonthRecords.forEach((item) => {

      const authorId = item.Author.ID; // Author ID

      const points = item.Points; // Points for the action

      const authorTitle = item.Author.Title; // Author Title

      const authorEMail = item.Author.EMail; // Author Email

      const authorDepartment = item.Author.Department; // Author Department

      

      // Sum points for each AuthorId

      if (userPoints[authorId]) {

        userPoints[authorId].TotalPoints += points;

      } else {

        userPoints[authorId] = {

          TotalPoints: points,

          AuthorTitle: authorTitle,

          AuthorEMail: authorEMail,

          AuthorDepartment: authorDepartment,

         

        };

      }

    });


    // Convert the object to an array of users with total points and other details

    const userList = Object.keys(userPoints).map((authorId) => ({

      AuthorId: authorId,

      TotalPoints: userPoints[authorId].TotalPoints,

      AuthorTitle: userPoints[authorId].AuthorTitle,

      AuthorEMail: userPoints[authorId].AuthorEMail,

      AuthorDepartment: userPoints[authorId].AuthorDepartment,

      AutherRatting: userPoints[authorId].TotalPoints,

      Ratting:getRating(RATINGSTHRESHOLDS, userPoints[authorId].TotalPoints)

    }));


    // Sort users by TotalPoints in descending order

    userList.sort((a, b) => b.TotalPoints - a.TotalPoints);


    return userList;

  };


  // Return the result with total points and rating, sorted by points

  return sumPointsByUser();

};

const getRating = (RATINGSTHRESHOLDS, totalPoints) => {

  // Sort thresholds by minPoints in ascending order

  const sortedThresholds = RATINGSTHRESHOLDS.sort((a, b) => a.minPoints - b.minPoints);


  // Find the rating where minPoints <= totalPoints

  let rating = 0; // Default rating if no threshold is met

  for (let i = 0; i < sortedThresholds.length; i++) {

    if (totalPoints >= sortedThresholds[i].minPoints) {

      rating = sortedThresholds[i].rating;

    } else {

      break; // No need to check further

    }

  }


  return rating;

};

export const addNotification = async (itemData, _sp) => {

  

  let resultArr = []

  try {

    const newItem = await _sp.web.lists.getByTitle('ARGNotificationHistory').items.add(itemData);

    

    console.log('Item added successfully:', newItem);

    // Swal.fire('Item added successfully', '', 'success');


    resultArr = newItem

    // Perform any necessary actions after successful addition

  } catch (error) {

    console.log('Error adding item:', error);

    // Handle errors appropriately

    resultArr = null

    Swal.fire(' Cancelled', '', 'error')

  }

  return resultArr;

};

export const UpdateNotification = async (itemId, _sp) => {

  

  let resultArr = []

  try {

    let arr = { ReadStatus: true }

    const newItem = await _sp.web.lists.getByTitle('ARGNotificationHistory').items.getById(itemId).update(arr);

    

    console.log('update added successfully:', newItem);

    // Swal.fire('Item added successfully', '', 'success');


    resultArr = newItem

    // Perform any necessary actions after successful addition

  } catch (error) {

    console.log('Error adding item:', error);

    // Handle errors appropriately

    resultArr = null

    Swal.fire(' Cancelled', '', 'error')

  }

  return resultArr;

};

export const updateNotifications = async (sp) => {

  let resultArr = "waiting";

  try {

    await getARGNotificationHistory(sp).then(async (x) => {

      for (let i = 0; i < x.length; i++) {

        await sp.web.lists

          .getByTitle("ARGNotificationHistory").items.getById(x[i].Id).update({

            ReadStatus: true

          })

      }

      resultArr = "Updated"

    })

  } catch (error) {

    console.log('Error adding item:', error);

    // Handle errors appropriately

    resultArr = null

  }

  return resultArr;

}

export const getARGNotificationHistory = async (sp) => {

  

  let resultArr = []

  const currentUser = await sp.web.currentUser();

  try {

    const newItem = await sp.web.lists.getByTitle('ARGNotificationHistory').items.select("*,NotifiedUser/Id,NotifiedUser/Title,ActionUser/Id,ActionUser/Title").expand("NotifiedUser,ActionUser").filter(`NotifiedUserId eq ${currentUser.Id} and ReadStatus eq false`).getAll();

    

    console.log('Item added successfully:', newItem);

    // Swal.fire('Item added successfully', '', 'success');


    resultArr = newItem

    // Perform any necessary actions after successful addition

  } catch (error) {

    console.log('Error adding item:', error);

    // Handle errors appropriately

    resultArr = null

    Swal.fire(' Cancelled', '', 'error')

  }

  return resultArr;

}

export const getAllARGNotificationHistory = async (sp) => {

  

  let resultArr = []

  const currentUser = await sp.web.currentUser();

  try {

    const newItem = await sp.web.lists.getByTitle('ARGNotificationHistory').items.select("*,NotifiedUser/Id,NotifiedUser/Title,ActionUser/Id,ActionUser/Title").expand("NotifiedUser,ActionUser").filter(`NotifiedUserId eq ${currentUser.Id}`).getAll();

    debugger

    console.log('Item added successfully:', newItem);

    // Swal.fire('Item added successfully', '', 'success');


    resultArr = newItem

    // Perform any necessary actions after successful addition

  } catch (error) {

    console.log('Error adding item:', error);

    // Handle errors appropriately

    resultArr = null

    Swal.fire(' Cancelled', '', 'error')

  }

  return resultArr;

}

export const getTodayARGNotificationHistory = async (sp) => {

  let resultArr = [];

  const currentUser = await sp.web.currentUser();


  // Format today's date

  const today = format(new Date(), 'yyyy-MM-dd');

  const startOfToday = new Date(today); // Start of today

  startOfToday.setHours(0, 0, 0, 0); // Set to 00:00:00

  const endOfToday = new Date(today); // End of today

  endOfToday.setHours(23, 59, 59, 999); // Set to 23:59:59


  const startTimestamp = startOfToday.toISOString();

  const endTimestamp = endOfToday.toISOString();


  try {

      const newItem = await sp.web.lists

          .getByTitle('ARGNotificationHistory')

          .items

          .select("*,NotifiedUser/Id,NotifiedUser/Title,ActionUser/Id,ActionUser/Title,ActionUser/EMail")

          .expand("NotifiedUser,ActionUser")

          .filter(`NotifiedUserId eq ${currentUser.Id} and Created ge '${startTimestamp}' and Created le '${endTimestamp}'`)

          .getAll();


      console.log('Today\'s items retrieved successfully:', newItem);

      resultArr = newItem;

  } catch (error) {

      console.error('Error retrieving today\'s items:', error);

      resultArr = null;

      Swal.fire('Cancelled', '', 'error');

  }


  return resultArr;

};


export const getlastSevenDaysARGNotificationHistory = async (sp) => {

  let resultArr = [];

  const currentUser = await sp.web.currentUser();


  // Calculate the date exactly seven days ago

  const exactSevenDaysAgo = format(subDays(new Date(), 7), 'yyyy-MM-dd');

  const startOfSevenDaysAgo = new Date(exactSevenDaysAgo); // 7 days ago

  startOfSevenDaysAgo.setHours(0, 0, 0, 0); // Set to 00:00:00

  const today = new Date();

  today.setHours(23, 59, 59, 999); // End of today


  const startTimestamp = startOfSevenDaysAgo.toISOString();

  const endTimestamp = today.toISOString();


  try {

      const newItem = await sp.web.lists

          .getByTitle('ARGNotificationHistory')

          .items

          .select("*,NotifiedUser/Id,NotifiedUser/Title,ActionUser/Id,ActionUser/Title,ActionUser/EMail")

          .expand("NotifiedUser,ActionUser")

          .filter(`NotifiedUserId eq ${currentUser.Id} and Created ge '${startTimestamp}' and Created le '${endTimestamp}'`)

          .getAll();


      console.log('Items from last seven days retrieved successfully:', newItem);

      resultArr = newItem;

  } catch (error) {

      console.error('Error retrieving items from last seven days:', error);

      resultArr = null;

      Swal.fire('Cancelled', '', 'error');

  }


  return resultArr;

};


export const getOlderARGNotificationHistory = async (sp) => {

  let resultArr = [];

  const currentUser = await sp.web.currentUser();


  // 8 days ago

  const eightDaysAgo = format(subDays(new Date(), 8), 'yyyy-MM-dd');

  const startOfEightDaysAgo = new Date(eightDaysAgo); // 8 days ago

  startOfEightDaysAgo.setHours(0, 0, 0, 0); // Set to 00:00:00


  const startTimestamp = startOfEightDaysAgo.toISOString();

  

  try {

      const newItem = await sp.web.lists

          .getByTitle('ARGNotificationHistory')

          .items

          .select("*,NotifiedUser/Id,NotifiedUser/Title,ActionUser/Id,ActionUser/Title,ActionUser/EMail")

          .expand("NotifiedUser,ActionUser")

          .filter(`NotifiedUserId eq ${currentUser.Id} and Created lt '${startTimestamp}'`)

          .getAll();


      console.log('8 days ago items retrieved successfully:', newItem);

      resultArr = newItem;

  } catch (error) {

      console.error('Error retrieving older items:', error);

      resultArr = null;

      Swal.fire('Cancelled', '', 'error');

  }


  return resultArr;

};

export const AllowAlphaNumericNotStartWithNumbers = (_name) => {

  var isValidName = true;

  if (_name.match(/^\d/)) //Not Start with numbers.

  {

    isValidName = false;

  }

  if (_name.match(/[!@#$%^&*(),.?":{}|<>]=_-/)) //These special characters are not allowed

  {

    isValidName = false;

  }

  return isValidName;

}

///Alpha numeric only. (,)  are allowed. White space allowed. If required add/remove more filters

export const allowstringandnumber = (_name) => {

  var isValid = true;

  if (_name.match(/^\d/)) //Not start with numbers

  {

    isValid = false;

  }

  if (_name.match(/[!@#$%^&*,.-?":{}|<>=_-]/)) {

    isValid = false;

  }

  return isValid;

}

export const allowstringonly = (_name) => {

  var isValid = true;

  if (_name.match(/^\d+=-_/)) //Not start with numbers

  {

    isValid = false;

  }

  if (_name.match(/[!@#$%^&*(),.?":{}|<>]/)) {

    isValid = false;

  }

  // if (_name.match(/[/""\\~123456789+=-_|-]/)) {

  //   isValid = false;

  //}

  return isValid;

}

//Only Numbers are Allowed.Convert parameter to string before passing it.

export const AllowNumbersOnly = (_val) => {

  var isValid = false;

  if (_val.match(/^[0-9\b]+$/) || _val === "") {

    isValid = true;

  }

  return isValid;

}

//Cannot be numeric

export const validate = (str) => {

  var isValidName = true;

  if (str.match(/^\d/)) //Not Start with numbers.

  {

    isValidName = false;

  }

  if (str.match(/[/""\\~123456789]/)) //These special characters are not allowed

  {

    isValidName = false;

  }

  return isValidName;

}

// import Swal from 'sweetalert2';
// export const getEntity = async (_sp) => {
//     let arr =[]
//    await _sp.web.lists.getByTitle("ARGEntityMaster").items.select("ID,Entity").filter("Active eq 1")()
//     .then((res) => {
//         console.log(res);
//         const newArray = res.map(({ ID, Entity }) => ({ id:ID, name:Entity }));
//         console.log(newArray,'newArray');

//         arr= newArray;
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }
// export const getCategory = async (_sp,id) => {
//     let arr =[]
//    await _sp.web.lists.getByTitle("ARGAnnouncementsandNewsCategory").items.select("ID,Category,AnnouncementandNewsTypeMaster/TypeMaster,AnnouncementandNewsTypeMaster/ID").expand("AnnouncementandNewsTypeMaster").filter(`(Active eq 1) and(AnnouncementandNewsTypeMaster/ID eq ${id})`)()
//     .then((res) => {
//         console.log(res);
//         const newArray = res.map(({ ID, Category }) => ({ id:ID, name:Category }));
//         console.log(newArray,'newArray');

//         arr= newArray;
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }
// export const getType = async (_sp) => {
//     let arr =[]
//    await _sp.web.lists.getByTitle("AnnouncementandNewTypeMaster").items.select("ID,TypeMaster")()
//     .then((res) => {
//         console.log(res);
//         const newArray = res.map(({ ID, TypeMaster }) => ({ id:ID, name:TypeMaster }));
//         console.log(newArray,'newArray');

//         arr= newArray;
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }
// export const getCurrentUser = async (_sp) => {
//     let arr =[]
//    await _sp.web.currentUser()
//     .then(async (res) => {
//         console.log(res);
       
//         arr= res;
//          const  ProfilePic=  `https://officeindia.sharepoint.com/sites/AlRostmani/_layouts/15/userphoto.aspx?size=M&accountname=${res.Email}`
//        //await getUserProfilePicture(res.Id,_sp)
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }
// export const getCurrentUserProfile = async (_sp) => {
//     const ProfilePic =[]
//    await _sp.web.currentUser()
//     .then(async (res) => {
//         console.log(res);
       
//      //   arr= res;
//          ProfilePic=  `https://officeindia.sharepoint.com/sites/AlRostmani/_layouts/15/userphoto.aspx?size=M&accountname=${res.Email}`
//        //await getUserProfilePicture(res.Id,_sp)
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return ProfilePic;
// }

// export const getCurrentUserName = async (_sp) => {
//     let arr =""
//    await _sp.web.currentUser()
//     .then((res) => {
//         console.log(res);
    
//         arr= res["Title"];
        
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }

// export async function getUserProfilePicture(authorId,sp) {
//     try {
//       // Get user information by ID
//       const user = await sp.web.getUserById(authorId)();
  
//       if (user) {
//         // The 'Picture' field holds the profile picture URL
//         const profilePictureUrl = `${user.PictureUrl}`;
//         console.log(profilePictureUrl,'profilePictureUrl');
        
//         return profilePictureUrl;
//       }
//       return null;
//     } catch (error) {
//       console.error("Error fetching profile picture:", error);
//       return null;
//     }
//   }

//   export const getCurrentUserProfileEmail = async (sp) => {
//     const Email =""
//    await sp.web.currentUser()
//     .then(async (res) => {
//         console.log(res);
       
//      //   arr= res;
//          Email=  res.Email;
//        //await getUserProfilePicture(res.Id,_sp)
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return Email;
// }
// export const getCurrentUserProfileEmailForPeople = async (sp) => {
//     const UserID =""
//    await sp.web.currentUser()
//     .then(async (res) => {
//         console.log(res);
       
//      //   arr= res;
//         // Email=  res.Email;
//          const UserID = await sp.web.ensureUser(res.Email);
       
//        //await getUserProfilePicture(res.Id,_sp)
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return user.data.Id; 
// }