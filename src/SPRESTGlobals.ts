"use strict";

export
enum UserRequestType {
   TYPE_CURRENT_USER,
   TYPE_USER_ID,
   TYPE_LAST_NAME,
   TYPE_FULL_NAME
};

// use as a bit flag
export
enum PrincipalTypeGroups {
   TYPE_SECURITY_GROUP = 4,
   TYPE_SHARPOINT_GROUP = 8
}

// use as a bit flag
export
enum PrincipalTypeUsers {
   TYPE_PRINCIPAL_ENTITY_USER = 1,
   TYPE_PRINCIPAL_ENTITY_SECURITY_GROUP = 4,
   TYPE_PRINCIPAL_ENTITY_SHAREPIONT_GROUNP = 8
}

export
enum ListTemplateType {
	InvalidType = -1,
	NoListTemplate = 0,
	GenericList = 100,
	DocumentLibrary = 101,
	Survey = 102,
	LinksList = 103,
	AnnouncementsList = 104,
	ContactsList = 105,
	EventsList = 106,
	TasksList = 107,
	DiscussionBoard = 108,
	PictureLibrary = 109,
	DataSources = 110,
	SitePagesLibrary = 111,
	WikiPageLibrary = 112,
	WebPartGallery = 113,
	ListTemplateCatalog=114,
	XMLForm=115,
	MasterPageCatalog=116,
	NoCodeWorkflows=117,
	WorkflowProcess=118,
	WebPageLibrary=119,
	CustomGrid=120,
	SolutionCatalog=121,
	NoCodePublic=122,
	ThemeCatalog=123,
	DesignCatalog=124,
	DataConnectionLibrary=130,
	WorkflowHistory=140,
	GanttTasks=150,
	PromotedLinks=170,
	Tasks=171,
	Meetings=200,
	Agenda=201,
	MeetingUser=202,
	Decision=204,
	MeetingObjective=207,
	TextBox=210,
	ThingsToBring=211,
	HomePageLibrary=212,
	PortalSitesList=300,
	Posts=301,
	Comments=302,
	Categories=303,
	Facility=402,
	Whereabouts=403,
	CallTrack=404,
	Circulation=405,
	Timecard=420,
	Holidays=421,
	ReportLibrary=433,
	IMEDic=499,
	ExternalList=600,
	MySiteDocumentLibrary=700,
	PageLibrary=850,
	AssetLibrary=851,
	IssueTracking=1100,
	AdminTasks=1200,
	HealthRules=1220,
	HealthReports=1221,
	DeveloperSiteDraftApps=1230,
}

export
function getListTemplateTypeLabel(value: number): keyof typeof ListTemplateType | null {
   for (const label in ListTemplateType) {
      if (ListTemplateType[label as keyof typeof ListTemplateType] === value) {
         return label as keyof typeof ListTemplateType;
      }
   }
   return null;
}

export
const ListTemplateTypeText: {[key: number]: string} = {
	[ListTemplateType.InvalidType]: "Invalid Type",
	[ListTemplateType.NoListTemplate]: "No List Template",
	[ListTemplateType.GenericList]: "Generic or Custom List",
	[ListTemplateType.DocumentLibrary]: "Document Library",
	[ListTemplateType.Survey]: "Survey",
	[ListTemplateType.LinksList]: "Links List",
	[ListTemplateType.AnnouncementsList]: "Announcements List",
	[ListTemplateType.ContactsList]: "Contacts List",
	[ListTemplateType.EventsList]: "Events List",
	[ListTemplateType.TasksList]: "Tasks List",
	[ListTemplateType.DiscussionBoard]: "Discussion Board",
	[ListTemplateType.PictureLibrary]: "Picture Library",
	[ListTemplateType.DataSources]: "Data Sources",
	[ListTemplateType.SitePagesLibrary]: "Site Pages Library",
	[ListTemplateType.WikiPageLibrary]: "Wiki Page Library",
	[ListTemplateType.WebPartGallery]: "Web Part Gallery",
	[ListTemplateType.ListTemplateCatalog ]: "List Template Catalog",
	[ListTemplateType.XMLForm ]: "XML Form",
	[ListTemplateType.MasterPageCatalog ]: "Master Page Catalog",
	[ListTemplateType.NoCodeWorkflows ]: "No Code Workflows",
	[ListTemplateType.WorkflowProcess ]: "Workflow Process",
	[ListTemplateType.WebPageLibrary ]: "Web Page Library",
	[ListTemplateType.CustomGrid ]: "Custom Grid",
	[ListTemplateType.SolutionCatalog ]: "Solution Catalog",
	[ListTemplateType.NoCodePublic ]: "No Code Public",
	[ListTemplateType.ThemeCatalog ]: "Theme Catalog",
	[ListTemplateType.DesignCatalog ]: "Design Catalog",
	[ListTemplateType.DataConnectionLibrary ]: "Data Connection Library",
	[ListTemplateType.WorkflowHistory ]: "Workflow History",
	[ListTemplateType.GanttTasks ]: "Gantt Tasks",
	[ListTemplateType.PromotedLinks ]: "Promoted Links",
	[ListTemplateType.Tasks ]: "Tasks",
	[ListTemplateType.Meetings ]: "Meetings",
	[ListTemplateType.Agenda ]: "Agenda",
	[ListTemplateType.MeetingUser ]: "Meeting User",
	[ListTemplateType.Decision ]: "Decision",
	[ListTemplateType.MeetingObjective ]: "Meeting Objective",
	[ListTemplateType.TextBox ]: "Text Box",
	[ListTemplateType.ThingsToBring ]: "Things To Bring",
	[ListTemplateType.HomePageLibrary ]: "Home Page Library",
	[ListTemplateType.PortalSitesList ]: "Portal Sites List",
	[ListTemplateType.Posts ]: "Posts",
	[ListTemplateType.Comments ]: "Comments",
	[ListTemplateType.Categories ]: "Categories",
	[ListTemplateType.Facility ]: "Facility",
	[ListTemplateType.Whereabouts ]: "Whereabouts",
	[ListTemplateType.CallTrack ]: "CallTrack",
	[ListTemplateType.Circulation ]: "Circulation",
	[ListTemplateType.Timecard ]: "Timecard",
	[ListTemplateType.Holidays ]: "Holidays",
	[ListTemplateType.ReportLibrary ]: "Report Library",
	[ListTemplateType.IMEDic ]: "IME Dic",
	[ListTemplateType.ExternalList ]: "External List",
	[ListTemplateType.MySiteDocumentLibrary ]: "My Site Document Library",
	[ListTemplateType.PageLibrary ]: "Page Library",
	[ListTemplateType.AssetLibrary ]: "Asset Library",
	[ListTemplateType.IssueTracking ]: "Issue Tracking",
	[ListTemplateType.AdminTasks ]: "Admin Tasks",
	[ListTemplateType.HealthRules ]: "Health Rules",
	[ListTemplateType.HealthReports ]: "Health Reports",
	[ListTemplateType.DeveloperSiteDraftApps ]: "Developer Site Draft Apps"
}

export
enum RoleTypeKind {
	None = 0 ,
	Guest = 1 ,
	Reader = 2 ,
	Contributor = 3 ,
	WebDesigner = 4 ,
	Administrator = 5 ,
	Editor = 6 ,
	Approver = 7 ,
	Builder = 8 ,
	Manager = 9 ,
	Owner = 10
}

export
enum SpGroupType {
	NONE = 0,
	USER = 1,
	DISTRIBUTION_LIST = 2,
	SECURITY_GROUP = 4,
	SHAREPOINT_GROUP = 8
}

export
enum BasePermissions {

	// Permission to view items in lists or libraries
		ViewListItems = 0x00000001,
	// Permission to add items to lists or libraries
		AddListItems =  0x00000002,
	// Permission to edit items in lists or libraries
		EditListItems =  0x00000004,
	// Permission to delete items from lists or libraries.
	DeleteListItems = 0x00000008,
	// Permission to approve items in lists or libraries.
	ApproveItems = 0x00000010,
	//Permission to view the source of documents with server-side file handlers.
	OpenItems = 0x00000020,
	// Permission to view past versions of a list item or document.
	ViewVersions = 0x00000040,
	// Permission to delete past versions of a list item or document.
	DeleteVersions = 0x00000080,
	// Permission to undo checkouts of list items or documents.
	CancelCheckout = 0x00000100,
	// Permission to create, change, and delete personal views of lists.
	ManagePersonalViews = 0x00000200,
	// Permission to create and delete lists, add or remove columns in a list, and add or remove public views of a list.
	ManageLists = 0x00000400,
	// Permission to view forms, views, and application pages, and to view reports.
	ViewFormPages = 0x00001000,
	// Permission to open a website, list, or folder as a website.
	Open = 0x00010000,
	// Permission to view pages in a website.
	ViewPages = 0x00020000,
	// Permission to add, change, or delete HTML pages or web part pages, and to edit the website using a SharePoint-compatible editor.
	AddAndCustomizePages = 0x00040000,
	// Permission to apply a theme or borders to the website.
	ApplyThemeAndBorder = 0x00080000,
	// Permission to apply a style sheet (.css file) to the website.
	ApplyStyleSheets = 0x00100000,
	// Permission to view reports on website usage.
	ViewUsageData = 0x00200000,
	// Permission to create a website using Self-Service Site Creation (SSC).
	CreateSSCSite = 0x00400000,
	// Permission to create, change, and delete subsites of a website.
	ManageSubwebs = 0x00800000,
	// Permission to create groups.
	CreateGroups = 0x01000000,
	// Permission to create and change permissions levels on the website and assign permissions to users and groups.
	ManagePermissions = 0x10000000,
	// Permission to browse directories and files on a server.
	BrowseDirectories = 0x20000000,
	// Permission to view information about users of the website.
	BrowseUserInfo = 0x40000000,
	// Permission to add or remove personal web parts on a web part page.
	AddDelPrivateWebParts = 0x2000000000,
	// Permission to update personal web parts on a web part page.
	UpdatePersonalWebParts = 0x4000000000,
	// Permission to perform all administration tasks for the website.
	ManageWeb = 0x20000000,
	}


export
const ListItemEntityTypeRE = /EntityType|EntityTypeName|ListItemEntityTypeFullName|metadata/,
		ListFieldEntityTypeRE = /Entity.*Type|metadata/,
		ListEntityTypeRE = /List.*Entity.*Type|Entity.*Type|metadata/,
		ContentTypeEntityTypeRE = /Content.*Entity.*Type|Entity.*Type|metadata/,
		ViewEntityTypeRE = /View.*Entity.*Type|Entity.*Type|metadata/,
      emailAddressRegex =
/(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])/;
