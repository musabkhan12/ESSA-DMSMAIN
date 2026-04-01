// HorizontalNavbar.tsx
import * as React from 'react';
// import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
// import { faBars, faExpand, faBell, faSun, faMoon } from '@fortawesome/free-solid-svg-icons';
import "../../horizontalNavBar/components/horizontalNavbar.scss";
import { Bell, ChevronDown, Maximize, Menu, Moon, Search, Settings, User, X } from 'react-feather';
import UserContext from '../../../GlobalContext/context';
// import { SPFI } from '@pnp/sp';
// import { getSP } from '../loc/pnpjsConfig';
import { useRef, useState } from 'react';
import { useMediaQuery } from 'react-responsive';
import { addActivityLeaderboard, 
  getARGNotificationHistory, 
  getCurrentUserName, getCurrentUserProfileEmail,
   UpdateNotification, 
   updateNotifications } from '../../../APISearvice/CustomService';
import "../../../CustomCss/mainCustom.scss"
// import moment from 'moment';
// import NotificationList from '../../../CustomJSComponents/CustomForm/NotificationList';
import NotificationList from './NotificationList';
// import { result } from 'lodash';
import { ListTitleTiSearchCategoryMapping } from './IHorizontalNavBarProps';


interface ListFieldsMapping {
  ARGAnnouncementAndNews: string;
  ARGBlogs: string;
  ARGDiscussionForum: string;
  ARGGroupandTeam: string;
  ARGProject: string;
  ARGSocialFeed: string;
  ARGEventMaster: string;
  ARGMediaGallery: string;
}

type ListTitle = keyof ListFieldsMapping;

interface SearchResult {
  ListTitle: ListTitle;
  [key: string]: any;
}
const HorizontalNavbar = ({ _context, siteUrl }: any) => {
  const listFieldsMapping: { [key: string]: { fields: string[], pageName: string } } = {
    ARGAnnouncementAndNews: { fields: ["Title", "Overview", "Description", "Id", "AnnouncementandNewsTypeMaster/Id", "AnnouncementandNewsTypeMaster/TypeMaster"], pageName: "AnnouncementDetails" },
    ARGBlogs: { fields: ["Title", "Overview", "Description", "Id"], pageName: "BlogDetails" },
    ARGDiscussionForum: { fields: ["Topic", "Overview", "Description", "Id"], pageName: "DiscussionForumDetail" },
    ARGGroupandTeam: { fields: ["GroupName", "Overview", "Id"], pageName: "GroupandTeamDetails" },
    ARGProject: { fields: ["ProjectName", "ProjectOverview", "Id"], pageName: "ProjectDetails" },
    ARGSocialFeed: { fields: ["Contentpost", "Id"], pageName: "SocialFeed" },
    ARGEventMaster: { fields: ["EventName", "Overview", "EventAgenda", "Id"], pageName: "EventDetailsCalendar" },
    ARGMediaGallery: { fields: ["Title", "Id"], pageName: "Mediadetails" }
  };
  console.log(siteUrl, 'siteUrl');

  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);
  const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);
  const { setHide }: any = React.useContext(UserContext);
  const [isDarkMode, setIsDarkMode] = useState(false);
  const [settingArray, setSettingArray] = useState([]);
  const [commentText, setCommentText] = useState<string>('');
  const [comments, setComments] = useState<Comment[]>([]);
  const [replyText, setReplyText] = useState<string>('');
  const [replyingTo, setReplyingTo] = useState<number | null>(null);
  const [isOpen, setIsOpen] = React.useState(false);
  const [isOpenBell, setIsOpenBell] = React.useState(false);
  const [groupedSearchResults, setGroupedSearchResults] = useState<any>({});
  const [currentUser, setCurrentUser] = React.useState("")
  const [currentUserEmail, setCurrentUserEmail] = React.useState("")
  const [searchResults, setSearchResults] = useState<SearchResult[]>([]);

  // Helper function to generate unique IDs
  const generateId = () => Math.floor(Math.random() * 100000);
  const [issearchOpen, setIsSearchOpen] = React.useState(false);
  const { useFullscreen, toggleHide, toggleFullscreen }: any = React.useContext(UserContext);
  const headerRef = useRef(null); // Reference to the header
  const [isSticky, setIsSticky] = useState(false);
  const scrollContainerRef = useRef(null);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState([]);
  const [NotificationArray, setNotificationArray] = useState([]);
  const menuRef = useRef(null);
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  function groupByFn(array: any, keyGetter: any) {
    return array.reduce((result: any, currentItem: any) => {
      const key = keyGetter(currentItem);
      if (!result[key]) {
        result[key] = [];
      }
      result[key].push(currentItem);
      return result;
    }, {});
  }
  const toggleDropdown = () => {
    setIsOpen(!isOpen);
  };
  const toggleDropdownBell = () => {
    setIsOpenBell(!isOpenBell);
    setIsMenuOpen(!isMenuOpen);
  };
  const handleScroll = () => {
    if (headerRef.current) {
      const sticky = headerRef.current.offsetTop;
      if (window.scrollY > sticky) {
        setIsSticky(true);
      } else {
        setIsSticky(false);
      }
    }
  };
  React.useEffect(() => {
    const handleClickOutside = (event: { target: any; }) => {
      if (menuRef.current && !menuRef.current.contains(event.target)) {
        setIsMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  });
  React.useEffect(() => {
    ApiCall();

    const scrollContainer = scrollContainerRef.current;
    if (scrollContainer) {
      scrollContainer.addEventListener("scroll", handleScroll);
    }

    const showNavbar = (
      toggleId: string,
      navId: string,
      bodyId: string,
      headerId: string
    ) => {
      const toggle = document.getElementById(toggleId);
      const nav = document.getElementById(navId);
      const bodypd = document.getElementById(bodyId);
      const headerpd = document.getElementById(headerId);

      if (toggle && nav && bodypd && headerpd) {
        toggle.addEventListener('click', () => {
          nav.classList.toggle('show');
          toggle.classList.toggle('bx-x');
          bodypd.classList.toggle('body-pd');
          headerpd.classList.toggle('body-pd');
        });
      }
    };

    showNavbar('header-toggle', 'nav-bar', 'body-pd', 'header');

    const linkColor = document.querySelectorAll('.nav_link');
    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach(l => l.classList.remove('active'));
        this.classList.add('active');
      }
    }


    linkColor.forEach(l => l.addEventListener('click', colorLink));
    return () => {
      if (scrollContainer) {
        scrollContainer.removeEventListener("scroll", handleScroll);
        window.addEventListener("scroll", handleScroll);
      }
    };

  }, [useHide]);
  const toggleSearchDropdown = () => {
    setIsSearchOpen(!issearchOpen);
  };
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });

  const handleSidebarToggle = (bol: boolean) => {

    setIsSidebarOpen(prevState => !prevState);
    setHide(!bol);
    document.querySelector(".sidebar")?.classList.toggle("close");
  };
  const handleThemeToggle = () => {
    setIsDarkMode((prevState: any) => !prevState);
    document.querySelector("body")?.classList.toggle("dark");
  };
  // const imgLogo = require("../assets/useimg.png");


  const ApiCall = async () => {
    setCurrentUser(await getCurrentUserName(_context))
    setCurrentUserEmail(await getCurrentUserProfileEmail(_context))

    setNotificationArray(await getARGNotificationHistory(_context))
    // console.log(settingsData, 'settingsData');
  };
  console.log(currentUser, siteUrl, 'currentUser');


  const searchAllLists = async (query: string): Promise<any[]> => {
    try {
      const lists = await _context.web.lists();
      let results: any[] = [];

      for (const list of lists) {
        const listTitle = list.Title.trim();
        const listMapping = listFieldsMapping[listTitle];

        if (listMapping) {
          const { fields, pageName } = listMapping;

          // Start building the query
          let queryBuilder = _context.web.lists.getByTitle(listTitle).items.top(100).select(...fields);

          // Conditionally expand for the specific list
          if (listTitle === "ARGAnnouncementAndNews") {
            queryBuilder = queryBuilder.expand("AnnouncementandNewsTypeMaster");
          }

          // Execute the query
          const items = await queryBuilder();

          // Filter items based on the search query
          const filteredItems = items.filter((item: any) =>
            fields.some(field =>
              typeof item[field] === 'string' && item[field].toLowerCase().includes(query.toLowerCase()) ||
              // Check for the expanded lookup field only for the specific list
              (listTitle === "ARGAnnouncementAndNews" &&
                item.AnnouncementandNewsTypeMaster &&
                item.AnnouncementandNewsTypeMaster.TypeMaster &&
                item.AnnouncementandNewsTypeMaster.TypeMaster.toLowerCase().includes(query.toLowerCase()))
            )
          );

          // Add ListTitle and PageName properties to the filtered items
          filteredItems.forEach((item: any) => {
            item.ListTitle = listTitle;
            item.PageName = pageName; // Add the PageName for each item
          });

          // Combine results
          results = [...results, ...filteredItems];
        }
      }



      return results;
    } catch (error) {
      console.error("Error searching lists:", error);
      return [];
    }
  };

  const searchKeyPress = async (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    const queryText = e.target.value;
    setQuery(queryText);

    if (queryText && queryText.length > 2) {
      const searchResults = await searchAllLists(queryText);
      let grped = groupByFn(searchResults, (res: any) => res.ListTitle)
      console.log("grped results", grped);
      setSearchResults(searchResults);
      setGroupedSearchResults(grped);
      console.log("grouped resuls after fncall", groupedSearchResults);
    }
  };
  const handleSearchClick = async (result: any) => {

    await addActivityLeaderboard(_context, "Search Results Click");
    debugger
    setTimeout(() => {
      window.location.href =
        result?.AnnouncementandNewsTypeMaster?.TypeMaster == "News"
          ?
          `${siteUrl}/SitePages/Newsdetails.aspx?${result.Id}`
          : `${siteUrl}/SitePages/${result.pageName}.aspx?${result.Id}`
    }, 2000);

  };

  const OnClearall = async (replyText: any) => {
    console.log("replaytext", replyText);
    let res: any[] = [];
    if (replyText == 'Clear') {
      await updateNotifications(_context).then((x:any) => {
        setNotificationArray(res);
        console.log("isUpdated", res, x);
      })


    }
  }

  const handleNotificationClick = async (result: any) => {

    await UpdateNotification(result.Id, _context);
    debugger
    setTimeout(() => {
      window.location.href =
        `${siteUrl}/SitePages/${result.DeatilPage}.aspx?${result.Id}`
    }, 2000);

  };
  const logout = () => {
    debugger
    localStorage.clear();
    sessionStorage.clear();
    const logoutUrl = `https://login.microsoftonline.com/common/oauth2/logout?post_logout_redirect_uri=${window.location.origin}`;
    window.location.href = logoutUrl;
  };
  return (
    // <nav className="navbar container-fluid" style={{ zIndex: '99' }}>
    //   <div className="logo_item">
    //     <div className="bottom_content">
    //       <div className="bottom expand_sidebar">
    //         <FontAwesomeIcon className={`bx bx`} icon={faBars} size='xs' />
    //       </div>
    //     </div>
    //   </div>

    //   <div className="navbar_content">
    //     <div className="search_bar">
    //       <input type="text" placeholder="Search.." className='searchcss' />
    //     </div>
    //     <FontAwesomeIcon className='bx bx-bell' icon={faExpand} onClick={toggleFullscreen} size='lg' />
    //     <FontAwesomeIcon className='bx bx-bell' icon={faBell} />
    //     <FontAwesomeIcon className={isDarkMode ? 'bx bx-moon' : 'bx bx-sun'} onClick={handleThemeToggle} icon={isDarkMode ? faMoon : faSun} size='lg' />
    //     <div className="dropdown">
    //       <img src={imgLogo} alt="Profile" className="profile dropbtn" onClick={toggleDropdown} />
    //       <div id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
    //         <a href="#home">Home</a>
    //         <a href="#about">About</a>
    //         <a href="#contact">Contact</a>
    //       </div>
    //     </div>
    //     <FontAwesomeIcon className='bx bx-user' icon={faGear} size='lg' />
    //   </div>
    // </nav>
    <div style={{ zIndex: '99' }} ref={headerRef}
      className={isSticky ? "sticky " : "navbar"}
      id="myHeader">
      <div className='navcss' style={{ marginLeft: `${!useHide ? '230px' : '80px'}` }} >
        <div className="" onClick={() => handleSidebarToggle(useHide)}>
          <div className={` ${useHide ? 'sidebar-closedBar' : 'sidebar-openBa'}`} onClick={() => handleSidebarToggle(useHide)}>
            <div className="" onClick={() => handleSidebarToggle(useHide)}>
              <Menu size={22} className='desktoView' />
              <Menu size={80} className='searchcssmobile' />
            </div>
          </div>
        </div>
        <div className={`navbar_content ${useHide ? 'searchcssmobile sidebar-closedBar' : 'searchcssmobile sidebar-openBa'}`} >
          <div className="search_bar">

            <input
              type="text"
              value={query} className='searchcss desktoView'
              onChange={(e) => searchKeyPress(e)}
              onClick={toggleSearchDropdown}
              placeholder="Search..."
            />
          </div>
          <div className="dropdown">
            <Search className='searchcssmobile' size='80' onClick={toggleSearchDropdown} />

            <div id="myDropdown" className={`dropdown-content ${issearchOpen ? 'show' : ''}`}>

              <input
                type="text"
                value={query}
                className='searchcss searchcssmobile'
                onChange={(e) => searchKeyPress(e)}
                placeholder="Search..."
              />
              <div className={searchResults.length > 0 ? 'search-results' : ''}>

                <div className={searchResults.length > 0 ? 'scrollbar' : ''} id={searchResults.length > 0 ? 'style-6' : ''}>
                  {searchResults.length > 0 && <span className='alifnsearch' style={{ padding: '0.85rem' }}>Found {searchResults.length} results</span>}
                  {console.log("grped searchResults dropfown", groupedSearchResults)}
                  {searchResults.length > 0 ? (


                    Object.keys(groupedSearchResults).map((grpreskey: any, grpind: number) => (

                      <div>
                        <div className='alifnsearch1' key={grpind}>{ListTitleTiSearchCategoryMapping[grpreskey]}({groupedSearchResults[grpreskey].length})</div>
                        {
                          groupedSearchResults[grpreskey].map((result: any, index: any) => (
                            <div key={index} className="search-result-item">
                              <a onClick={() => handleSearchClick(result)} style={{ padding: '0.85rem' }}>
                                <h4 className='eclipcsss text-dark' style={{ fontSize: '16px' }}>{result.Title || result.ProjectName || result.EventName || result.Contentpost}</h4>
                                {/* {result.Description && <p dangerouslySetInnerHTML={{ __html: result.Description }}></p>} */}
                                {result.Overview && <p className='eclipcsss text-muted' style={{ fontSize: '14px' }}>{result.Overview}</p>}
                                {result.EventAgenda && <p className='eclipcsss text-muted' style={{ fontSize: '14px' }}>{result.EventAgenda}</p>}
                              </a>
                            </div>

                          ))
                        }
                      </div>


                    ))

                  ) : (
                    null
                  )}
                  <div className="force-overflow"></div>
                </div>
              </div>

            </div>
          </div>
          <Maximize className='bx bx-bell desktoView' size='22' onClick={toggleFullscreen} />
          <div className="dropdown notification-lists" onClick={toggleDropdownBell}>

            <a className="nav-link dropdown-toggle waves-effect waves-light arrow-none" data-bs-toggle="dropdown"
              role="button" aria-haspopup="false" aria-expanded="false">
              <Bell className='bx bx-bell desktoView dropcssBell' size='22' onClick={toggleDropdownBell} style={{ position: 'relative' }} />
              {NotificationArray.length > 0 && <span className="badge bg-danger noti-icon-badge">{NotificationArray.length}</span>}
            </a>

            {isMenuOpen &&
            <div id="myDropdownBell" className={`dropdown-content  ${isOpenBell ? 'show desktoView' : ''}`} style={{ width: '320px' }} ref={menuRef}>


              <NotificationList NotificationArray={NotificationArray} handleNotificationClick={handleNotificationClick} OnClearall={OnClearall} />

            </div>
            }
          </div>
          {/* <Moon size='22' className={isDarkMode ? 'bx bx-moon desktoView' : 'bx bx-sun desktoView'} onClick={handleThemeToggle} /> */}
          <div className="dropdown searchcssmobile ">
            <Bell className='bx bx-bell searchcssmobile dropcssBell' size='80' onClick={toggleDropdownBell} />
            {/* <div id="myDropdownBell" className={`dropdown-content searchcssmobile ${isOpenBell ? 'show' : ''}`}>

              
              <NotificationList NotificationArray={NotificationArray} handleNotificationClick={handleNotificationClick}/>

            </div> */}
          </div>
          <div className="dropdown">
            <div className='d-flex' onClick={toggleDropdown} style={{ gap: '2px', cursor: 'pointer' }}>
              <div >

                <img src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentUserEmail}`}
                  className="rounded-circlecss img-thumbnail desktoView 
                                  avatar-xl"
                  alt="profile-image"
                  style={{ cursor: "pointer" }} />
                <img src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentUser}`}
                  className="rounded-circlecss img-thumbnail searchcssmobile 
                                  avatar-xl"
                  alt="profile-image"
                  style={{ cursor: "pointer" }} />

              </div>
              <div className='dropcssUser desktoView'>
                <div>{currentUser}</div>
                {/* <div><ChevronDown size={12} /></div> */}
              </div>
            </div>
            {/* <div id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
              <a href="#home">Home</a>
              <a href="#about">About</a>
              <a href="#contact">Contact</a>
            </div> */}
          </div>
          <svg style={{ opacity: '0.6', cursor: 'pointer' }} width="33px" height="35px" viewBox="0 -0.5 25 25" fill="none" xmlns="http://www.w3.org/2000/svg" onClick={() => logout()}>
            <path d="M11.75 9.874C11.75 10.2882 12.0858 10.624 12.5 10.624C12.9142 10.624 13.25 10.2882 13.25 9.874H11.75ZM13.25 4C13.25 3.58579 12.9142 3.25 12.5 3.25C12.0858 3.25 11.75 3.58579 11.75 4H13.25ZM9.81082 6.66156C10.1878 6.48991 10.3542 6.04515 10.1826 5.66818C10.0109 5.29121 9.56615 5.12478 9.18918 5.29644L9.81082 6.66156ZM5.5 12.16L4.7499 12.1561L4.75005 12.1687L5.5 12.16ZM12.5 19L12.5086 18.25C12.5029 18.25 12.4971 18.25 12.4914 18.25L12.5 19ZM19.5 12.16L20.2501 12.1687L20.25 12.1561L19.5 12.16ZM15.8108 5.29644C15.4338 5.12478 14.9891 5.29121 14.8174 5.66818C14.6458 6.04515 14.8122 6.48991 15.1892 6.66156L15.8108 5.29644ZM13.25 9.874V4H11.75V9.874H13.25ZM9.18918 5.29644C6.49843 6.52171 4.7655 9.19951 4.75001 12.1561L6.24999 12.1639C6.26242 9.79237 7.65246 7.6444 9.81082 6.66156L9.18918 5.29644ZM4.75005 12.1687C4.79935 16.4046 8.27278 19.7986 12.5086 19.75L12.4914 18.25C9.08384 18.2892 6.28961 15.5588 6.24995 12.1513L4.75005 12.1687ZM12.4914 19.75C16.7272 19.7986 20.2007 16.4046 20.2499 12.1687L18.7501 12.1513C18.7104 15.5588 15.9162 18.2892 12.5086 18.25L12.4914 19.75ZM20.25 12.1561C20.2345 9.19951 18.5016 6.52171 15.8108 5.29644L15.1892 6.66156C17.3475 7.6444 18.7376 9.79237 18.75 12.1639L20.25 12.1561Z" fill="#000000" />
          </svg>
          {/* <Settings className='bx bx-user desktoView' size='22' />
          <Settings className='bx bx-user searchcssmobile' size='80' /> */}
        </div>
      </div>
    </div>
  );
};

export default HorizontalNavbar;
// // HorizontalNavbar.tsx
// import * as React from 'react';
// import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
// import { faBars, faExpand, faBell, faSun, faMoon, faGear } from '@fortawesome/free-solid-svg-icons';
// import "../../horizontalNavBar/components/horizontalNavbar.scss";
// import { Bell, ChevronDown, Maximize, Menu, Moon, Search, Settings, User } from 'react-feather';
// import UserContext from '../../../GlobalContext/context';
// import { SPFI } from '@pnp/sp';
// import { getSP } from '../loc/pnpjsConfig';
// import { useRef, useState } from 'react';
// import { useMediaQuery } from 'react-responsive';
// import { getCurrentUserName, getCurrentUserProfileEmail } from '../../../APISearvice/CustomService';
// import "../../../CustomCss/mainCustom.scss"

// const HorizontalNavbar = ({_context,siteUrl}: any) => {
//   const sp: SPFI = getSP();
//   const { useHide }: any = React.useContext(UserContext);
//   const elementRef = React.useRef<HTMLDivElement>(null);
//   const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);
//   const {  setHide }: any = React.useContext(UserContext);
//   const [isDarkMode, setIsDarkMode] = useState(false);
//   const [settingArray, setSettingArray] = useState([]);
//   const [commentText, setCommentText] = useState<string>('');
//   const [comments, setComments] = useState<Comment[]>([]);
//   const [replyText, setReplyText] = useState<string>('');
//   const [replyingTo, setReplyingTo] = useState<number | null>(null);
//   const [isOpen, setIsOpen] = React.useState(false);
//   const [currentUser, setCurrentUser] = React.useState("")
//   const [currentUserEmail, setCurrentUserEmail] = React.useState("")

//   // Helper function to generate unique IDs
//   const generateId = () => Math.floor(Math.random() * 100000);
//   const [issearchOpen, setIsSearchOpen] = React.useState(false);
//   const { useFullscreen, toggleHide, toggleFullscreen }: any = React.useContext(UserContext);
//   const headerRef = useRef(null); // Reference to the header
//   const [isSticky, setIsSticky] = useState(false);
//   const scrollContainerRef = useRef(null);
//   const toggleDropdown = () => {
//     setIsOpen(!isOpen);
//   };

//   const handleScroll = () => {
//     if (headerRef.current) {
//       const sticky = headerRef.current.offsetTop;
//       if (window.scrollY > sticky) {
//         setIsSticky(true);
//       } else {
//         setIsSticky(false);
//       }
//     }
//   };

//   React.useEffect(() => {
//     ApiCall();
 
//       const scrollContainer = scrollContainerRef.current;
//       if (scrollContainer) {
//         scrollContainer.addEventListener("scroll", handleScroll);
//       }
    
//     const showNavbar = (
//       toggleId: string,
//       navId: string,
//       bodyId: string,
//       headerId: string
//     ) => {
//       const toggle = document.getElementById(toggleId);
//       const nav = document.getElementById(navId);
//       const bodypd = document.getElementById(bodyId);
//       const headerpd = document.getElementById(headerId);

//       if (toggle && nav && bodypd && headerpd) {
//         toggle.addEventListener('click', () => {
//           nav.classList.toggle('show');
//           toggle.classList.toggle('bx-x');
//           bodypd.classList.toggle('body-pd');
//           headerpd.classList.toggle('body-pd');
//         });
//       }
//     };

//     showNavbar('header-toggle', 'nav-bar', 'body-pd', 'header');

//     const linkColor = document.querySelectorAll('.nav_link');
//     function colorLink(this: HTMLElement) {
//       if (linkColor) {
//         linkColor.forEach(l => l.classList.remove('active'));
//         this.classList.add('active');
//       }
//     }
   

//     linkColor.forEach(l => l.addEventListener('click', colorLink));
//     return () => {
//       if (scrollContainer) {
//         scrollContainer.removeEventListener("scroll", handleScroll);
//         window.addEventListener("scroll", handleScroll);
//       }
//     };
//   }, [useHide]);
//   const toggleSearchDropdown = () => {
//     setIsSearchOpen(!issearchOpen);
//   };
//   const isMobile = useMediaQuery({ query: '(max-width: 768px)' });

//   const handleSidebarToggle = (bol: boolean) => {
//     debugger
//     setIsSidebarOpen(prevState => !prevState);
//     setHide(!bol);
//     document.querySelector(".sidebar")?.classList.toggle("close");
//   };
//   const handleThemeToggle = () => {
//     setIsDarkMode((prevState: any) => !prevState);
//     document.querySelector("body")?.classList.toggle("dark");
//   };
//   // const imgLogo = require("../assets/useimg.png");

  
//   const ApiCall = async () => {
//     setCurrentUser(await getCurrentUserName(_context))
//     setCurrentUserEmail(await getCurrentUserProfileEmail(_context))

//     // const settingsData = setSettingArray(await getSettingAPI(sp))
//     // console.log(settingsData, 'settingsData');
//   };
//   console.log(currentUser,siteUrl,'currentUser');
//   return (
//     // <nav className="navbar container-fluid" style={{ zIndex: '99' }}>
//     //   <div className="logo_item">
//     //     <div className="bottom_content">
//     //       <div className="bottom expand_sidebar">
//     //         <FontAwesomeIcon className={`bx bx`} icon={faBars} size='xs' />
//     //       </div>
//     //     </div>
//     //   </div>

//     //   <div className="navbar_content">
//     //     <div className="search_bar">
//     //       <input type="text" placeholder="Search.." className='searchcss' />
//     //     </div>
//     //     <FontAwesomeIcon className='bx bx-bell' icon={faExpand} onClick={toggleFullscreen} size='lg' />
//     //     <FontAwesomeIcon className='bx bx-bell' icon={faBell} />
//     //     <FontAwesomeIcon className={isDarkMode ? 'bx bx-moon' : 'bx bx-sun'} onClick={handleThemeToggle} icon={isDarkMode ? faMoon : faSun} size='lg' />
//     //     <div className="dropdown">
//     //       <img src={imgLogo} alt="Profile" className="profile dropbtn" onClick={toggleDropdown} />
//     //       <div id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
//     //         <a href="#home">Home</a>
//     //         <a href="#about">About</a>
//     //         <a href="#contact">Contact</a>
//     //       </div>
//     //     </div>
//     //     <FontAwesomeIcon className='bx bx-user' icon={faGear} size='lg' />
//     //   </div>
//     // </nav>
//     <div  style={{ zIndex: '99' }} ref={headerRef}
//     className={isSticky ? "sticky " : "navbar"}
//     id="myHeader">
//     <div className='navcss' style={{marginLeft: `${!useHide ? '240px' : '80px'}`}} >
//       <div className="" onClick={() => handleSidebarToggle(useHide)}>
//         <div className={` ${useHide ? 'sidebar-closedBar' : 'sidebar-openBa'}`} onClick={() => handleSidebarToggle(useHide)}>
//           <div className="" onClick={() => handleSidebarToggle(useHide)}>
//             <Menu size={22} className='desktoView' />
//             <Menu size={80} className='searchcssmobile' />
//           </div>
//         </div>
//       </div>
//       <div className={`navbar_content ${useHide ? 'searchcssmobile sidebar-closedBar' : 'searchcssmobile sidebar-openBa'}`} onClick={() => handleSidebarToggle(useHide)}>
//         <div className="search_bar">
//           <input type="text" placeholder="Search.." className='searchcss desktoView' />
//         </div>
//         <div className="dropdown">
//           <Search className='searchcssmobile' size='80' onClick={toggleSearchDropdown} />
//           <div id="myDropdown" className={`dropdown-content ${issearchOpen ? 'show' : ''}`}>
//             <input type="text" placeholder="Search.." className='searchcss searchcssmobile' />
//           </div>
//         </div>
//         <Maximize className='bx bx-bell desktoView' size='22' onClick={toggleFullscreen} />
//         <Bell className='bx bx-bell desktoView' size='22' />
//         <Moon size='22' className={isDarkMode ? 'bx bx-moon desktoView' : 'bx bx-sun desktoView'} onClick={handleThemeToggle} />
//         <Bell className='bx bx-bell searchcssmobile' size='80' />
//         <div className="dropdown">
//           <div className='d-flex' onClick={toggleDropdown} style={{ gap: '2px', cursor: 'pointer' }}>
//             <div >
              
//               <img  src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentUserEmail}`}
//                                   className="rounded-circlecss1 img-thumbnail mt-1 me-2 desktoView 
//                                   avatar-xl"
//                                   alt="profile-image"
//                                   style={{ cursor: "pointer",width:'40px',height:'40px' }}  />
//                                    <img  src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentUserEmail}`}
//                                   className="rounded-circlecss1 img-thumbnail searchcssmobile 
//                                   avatar-xl"
//                                   alt="profile-image"
//                                   style={{ cursor: "pointer",width:'40px',height:'40px' }} />
             
//             </div>
//             <div className='dropcssUser desktoView'>
//               <div>{currentUser}</div>
//               <div><ChevronDown size={12} /></div>
//             </div>
//           </div>
//           <div id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
//             <a href="#home">Home</a>
//             <a href="#about">About</a>
//             <a href="#contact">Contact</a>
//           </div>
//         </div>
//         <Settings className='bx bx-user desktoView' size='22' />
//         <Settings className='bx bx-user searchcssmobile' size='80' />
//       </div>
//     </div>
//   </div>
//   );
// };

// export default HorizontalNavbar;
