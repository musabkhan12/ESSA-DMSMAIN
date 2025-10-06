import * as React from 'react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import "../components/VerticalSidebar.scss";
import "../../horizontalNavBar/components/horizontalNavbar.scss";
import UserContext from '../../../GlobalContext/context';
import { faChevronRight, faChevronUp } from '@fortawesome/free-solid-svg-icons';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import { Airplay, Calendar, File, Image, Clipboard, Bell, Users, Activity, Sun, Moon, Cpu, Rss, Maximize, Settings, Search, ChevronsDown, ChevronDown, Menu, User, Codepen, Command } from 'react-feather';
import classNames from 'classnames';
import { getCurrentUserName } from '../../../APISearvice/CustomService';
import { getSP } from '../loc/pnpjsConfig';

interface NavItem {
  Title: string;
  Url: string;
  Icon: string;
  ParentId?: number;
  ID: number;
}

const VerticalContext = ({ _context }: any) => {
  const sp: SPFI = getSP();

  const imgBigLogo = require("../assets/logodarkBig.png");
  const imgSMLogo = require("../assets/logoImgsm.png");

  const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);
  const [isDarkMode, setIsDarkMode] = React.useState(false);
  const [submenuOpen, setSubmenuOpen] = React.useState<number | null>(null);
  const [navItems, setNavItems] = React.useState<NavItem[]>([]);
  const [isDropdownOpen, setIsDropdownOpen] = React.useState(false);
  const [isMobile, setIsMobile] = React.useState(window.innerWidth < 768);
  const [isOpen, setIsOpen] = React.useState(false);
  const [issearchOpen, setIsSearchOpen] = React.useState(false);

  const [currentUser, setCurrentUser] = React.useState("");
  const [bigLogo, setBigLogo] = React.useState("");
  const [smallLogo, setSmallLogo] = React.useState("");

  // Use local state for active Nav ID
  const [useActive, setUseActive] = React.useState<number>(0);

  const context = React.useContext(UserContext);
  const { setHide, useHide }: any = context; // keep hide functionality from context

  const toggleDropdown = () => setIsOpen(!isOpen);
  const toggleSearchDropdown = () => setIsSearchOpen(!issearchOpen);
  const toggleSidebar = () => setIsSidebarOpen(prev => !prev);

  const handleWindowResize = () => setIsMobile(window.innerWidth < 768);

  const fetchNavItems = async () => {
    // Set useActive from localStorage if present
    const navId = localStorage.getItem("NavId");
    if (navId) setUseActive(Number(navId));

    setCurrentUser(await getCurrentUserName(_context));

    // Fetch Logos
    await _context.web.lists.getByTitle("UtilitySettings").items.getAll().then((res: any) => {
      const ImageUrl = res[0].LogoImage ? JSON.parse(res[0].LogoImage) : { serverUrl: "", serverRelativeUrl: "" };
      const ImagesmUrl = res[0].SmallLogo ? JSON.parse(res[0].SmallLogo) : { serverUrl: "", serverRelativeUrl: "" };
      setBigLogo(ImageUrl.serverUrl + ImageUrl.serverRelativeUrl);
      setSmallLogo(ImagesmUrl.serverUrl + ImagesmUrl.serverRelativeUrl);
    });

    const currentUser = await _context.web.currentUser();
    const userGroups = await _context.web.currentUser.groups();
    const grptitle: string[] = userGroups.map((g: any) => g.Title.toLowerCase());

    const res: any = await _context.web.lists
      .getByTitle("ARGSidebarNavigation")
      .items.select("Title,Url,Icon,ParentId,ID,EnableAudienceTargeting,Audience/Title")
      .expand("Audience")
      .orderBy("Order0", true)
      .getAll();

    const securednavitems = res.filter((nav: any) => {
      return !nav.EnableAudienceTargeting || (nav.EnableAudienceTargeting && nav.Audience && nav.Audience.some((nv1: any) => grptitle.includes(nv1.Title.toLowerCase())));
    });

    setNavItems(securednavitems);
  };

  React.useEffect(() => {
    fetchNavItems();

    const sidebar = document.querySelector(".sidebar");

    const handleMouseEnter = () => {
      if (sidebar?.classList.contains("hoverable")) {
        sidebar.classList.remove("close");
        setIsSidebarOpen(true);
      }
    };
    const handleMouseLeave = () => {
      if (sidebar?.classList.contains("hoverable")) {
        sidebar.classList.add("close");
        setIsSidebarOpen(false);
      }
    };

    sidebar?.addEventListener("mouseenter", handleMouseEnter);
    sidebar?.addEventListener("mouseleave", handleMouseLeave);

    if (window.innerWidth < 768) sidebar?.classList.add("close");

    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === "Escape" && document.fullscreenElement) document.exitFullscreen();
    };

    window.addEventListener("click", closeDropdown);
    window.addEventListener("keydown", handleEscape);
    window.addEventListener("resize", handleWindowResize);

    return () => {
      sidebar?.removeEventListener("mouseenter", handleMouseEnter);
      sidebar?.removeEventListener("mouseleave", handleMouseLeave);
      window.removeEventListener("keydown", handleEscape);
      window.removeEventListener("click", closeDropdown);
      window.removeEventListener("resize", handleWindowResize);
    };
  }, [_context]);

  const closeDropdown = (event: any) => {
    if (!event.target.matches(".dropbtn")) setIsOpen(false);
  };

  const highlightbackground = () => {
    const url = window.location.href;
    const matches = url.match(/\/([^\/]+)\.aspx/);

    if (!matches) return;

    const pageName = matches[1];
    const mapping: { [key: string]: number } = {
      workbench: 1,
      Dashboard: 1,
      MediaGallery: 2,
      Mediadetails: 2,
      EventCalendar: 3,
      EventDetailsCalendar: 3,
      News: 4,
      NewsDetails: 4,
      Announcements: 5,
      AnnouncementDetails: 5,
      CorporateDirectory: 6,
      SocialFeed: 8,
      DiscussionForum: 9,
      DiscussionForumDetail: 9,
      Blogs: 10,
      BlogDetails: 10,
      GroupandTeam: 11,
      GroupandTeamDetails: 11,
      Project: 12,
      ProjectDetails: 12,
      MediaGalleryMaster: 0,
      MediaGalleryForm: 0,
      EventMaster: 0,
      EventMasterForm: 0,
      announcementmaster: 0,
      AddAnnouncement: 0,
      BannerMaster: 0,
      BannerForm: 0,
    };

    if (mapping[pageName] !== undefined) {
      const id = mapping[pageName];
      localStorage.setItem("NavId", String(id));
      setUseActive(id);
    } else {
      localStorage.setItem("NavId", "");
    }
  };

  React.useEffect(() => {
    highlightbackground();
  }, []);

  const gotoPage = (url: string, Id: number) => {
    localStorage.setItem("NavId", String(Id));
    setUseActive(Id);
    if (url) window.location.href = url;
  };

  const renderNavItems = (items: NavItem[], parentId: number | null = null) => {
    return items
      .filter(item => item.ParentId === parentId)
      .map(item => {
        const IconComponent = getIcon(item.Icon);
        return (
          <li key={item.ID} className={classNames("item", { active: submenuOpen === item.ID && !useHide })}>
            <div className={classNames("nav_link submenu_item", { active: item.ID === useActive && !useHide })} onClick={() => setSubmenuOpen(prev => (prev === item.ID ? null : item.ID))}>
              <div style={{ display: "flex", alignItems: "center", lineHeight: "25px" }} onClick={() => gotoPage(item.Url, item.ID)}>
                <span className="navlink_icon">{IconComponent && <IconComponent size={18} />}</span>
                <a className="link_name1" style={{ textDecoration: "unset", paddingLeft: "1rem" }}>
                  <span style={{ display: useHide ? "none" : "block" }}>{item.Title}</span>
                </a>
                {!useHide && items.some(subItem => subItem.ParentId === item.ID) && (
                  <FontAwesomeIcon className="arrow-left" icon={submenuOpen === item.ID ? faChevronUp : faChevronRight} />
                )}
              </div>
              {!useHide && submenuOpen === item.ID && (
                <ul className="menu_items nav_link submenu_item" style={{ background: "#fff", borderRadius: "unset", display: "block", paddingTop: ".25rem" }}>
                  {renderNavItems(items, item.ID)}
                </ul>
              )}
            </div>
          </li>
        );
      });
  };

  const getIcon = (iconName: string) => {
    const iconMap: { [key: string]: any } = {
      home: Airplay,
      calendar: Calendar,
      file: File,
      image: Image,
      clipboard: Clipboard,
      bell: Bell,
      userGroup: Users,
      wifi: Rss,
      waveSquare: Activity,
      sun: Sun,
      moon: Moon,
      approval: Activity,
      gear: Cpu,
      codepen: Codepen,
      command: Command,
    };
    return iconMap[iconName] || null;
  };

  return (
    <div className={classNames("sidebar", { open: !useHide && isMobile, close: useHide })}>
      <div className="menu_content">
        <ul className="menu_items">
          <li className="item mt-1 mb-0 pt-0">
            <div className="logo_item">
              <span>
                <img src={useHide ? smallLogo || imgSMLogo : bigLogo || imgBigLogo} alt="Logo" style={{ objectFit: "cover", width: "100%" }} />
              </span>
            </div>
          </li>
          {renderNavItems(navItems)}
        </ul>
      </div>
    </div>
  );
};

const VerticalSideBar = ({ _context }: any) => {
  return <VerticalContext _context={_context} />;
};

export default VerticalSideBar;
