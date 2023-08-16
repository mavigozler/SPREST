

import {
	ListTemplateType,
	RoleTypeKind,
	SpGroupType
} from './SPRESTGlobals';

interface ISPBaseListFieldProperties {
	__metadata: { type: string; }
	Id: number;
	InternalName: string;
	Description: string;
	DefaultValue: string;
	Title: string;
	FromBaseType: boolean;
	CanBeDeleted: boolean;
	Required: boolean;
	FieldTypeKind: number;
	TypeAsString: string;
	Hidden: boolean;
	Choices?: {
		__metadata: {
			type: string;
		}
	}
}

interface ISPBaseListItemProperties {
	__metadata: { type: string; }
	Id: number;
	InternalName: string;
	File: { Name: string; }
	Title: string;
}


// defined by JQuery AJAX
type TJQueryPlainObject = {[key:string]: [] | typeof document | string };
type TXmlHttpRequestData =  [] | string | TJQueryPlainObject | Uint8Array;

export
type TSPResponseDataProperties = {
	Id: string;
	OtherId: string;
	Title: string;
	InternalName?: string;
	ServerRelativeUrl: string;
	WebTemplate?: string;
};

export
type TSPResponseData = {
	d?: {
		results?: TSPResponseDataProperties[],
		GetContextWebInformation?: {
			FormDigestValue: string;
		}
		__next?: string;
		[key:string]: any;
	};
	FormDigestValue?: string;
	__next?: string;
	"odata.nextLink"?: string; // a variant of __next
	value?: object; // where the response is an object of indetermine properties

};

export
type TSuccessCallback = (data: TSPResponseData, status?: string, reqObj?: JQueryXHR) => void;
export
type TErrorCallback = (reqObj: JQueryXHR, status?: string, err?: string) => void;

export
type THttpRequestProtocol = "https" | "http" | "https://" | "http://";

type THttpRequestMethods = "GET" | "POST" | "PUT" | "PATCH" | "HEAD" | "OPTIONS" |
		"DELETE" | "TRACE" | "CONNECT";

type THttpRequestHeaders = {
	"Accept"?: "application/json;odata=verbose" | "application/json;odata=nometadata";
	"Content-Type"?: "application/json;odata=verbose" | "text/plain" | "text/html" | "text/xml";
	"IF-MATCH"?: "*"; // can also use etag
	"X-HTTP-METHOD"?: "MERGE" | "DELETE";
	"X-RequestDigest"?: string;
	[key: string]: any;
};

// this type used to monitor HTTP request calls in an intermediate fashion
type TIntervalControl = {
	currentCount?: number;   // keeps track of current counts of results returned in (AJAX) request
	nextCount?: number;      // keeps track of next count
	interval: number;       // the number of results returned for each interval
	callback: (count: number) => void;  // function to call when a certain count of results are returned in request
} | null;

type THttpRequestParams = {
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	customHeaders?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?:  TXmlHttpRequestData;
	successCallback: TSuccessCallback;
	errorCallback: TErrorCallback;
	ignore?: number[]; // errors to ignore/catch
	progressReport?: TIntervalControl | null
};

type THttpRequestParamsWithPromise = {
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?: TXmlHttpRequestData;
	successCallback?: TSuccessCallback;
	errorCallback?: TErrorCallback;
	ignore?: number[]; // errors to ignore/catch
	progressReport?: TIntervalControl
};


type HttpStatus = 0
  | 100 | 101 | 102
  | 200 | 201 | 202 | 203 | 204 | 205 | 206 | 207 | 208 | 226
  | 300 | 301 | 302 | 303 | 304 | 305 | 306 | 307 | 308
  | 400 | 401 | 402 | 403 | 404 | 405 | 406 | 407 | 408 | 409 | 410 | 411 | 412 | 413 | 414 | 415 | 416 | 417 | 418 | 422 | 423 | 424 | 425 | 426 | 428 | 429 | 431 | 451
  | 500 | 501 | 502 | 503 | 504 | 505 | 506 | 507 | 508 | 510 | 511;

type TFetchInfo = {
	RequestedUrl: string;
	Etag: null | string;
	RequestDigest?: string | null;
	ContentType: string | null;
	HttpStatus: HttpStatus;
	Data: {
		error: {
			message: {
				value: string;
			}
		}
	} | TSPResponseData | undefined;
	ProcessedData: TSPResponseData | undefined;
	ResponseIndex?: number;
};

type THttpRequestBody<T = any> = {
		[key: string]: T;
	}

export
type TArrayToJSON = Record<string, string | boolean | object>

type TQueryProperties = "Expand" | "Filter" | "Select" | "expand" | "filter" | "select";

interface IBatchHTTPRequestParams {
	host: string;
	path: string;
	protocol?: THttpRequestProtocol;
	AllHeaders?: THttpRequestHeaders;
	AllMethods?: THttpRequestMethods;
}

interface IBatchHTTPRequestForm {
	url: string;
	contextinfo: string;
	protocol?: "http" | "https";
	method?: THttpRequestMethods;
	body?: THttpRequestBody;
	headers?: THttpRequestHeaders;
	trackingID?: number;
}

type TBatchResponseRaw = {
	rawData: string;
	ETag: string | null;
	RequestDigest: string | null;
}

type TBatchResponse = {
	success: TFetchInfo[];
	error: TFetchInfo[];
	trackingData?: {
		ETag: string | null;
		RequestDigest: string | null;
	};
};

type TParsedURL = {
	originalUrl: string;
	protocol: THttpRequestProtocol | null;
	server: string;
	hostname: string;
	siteFullPath: string;
	sitePartialPath: string;
	list: string | null;
	listConfirmed: boolean;
	libRelativeUrl: string | null;
	file: string | null;
	query: any[] | null;
};

type TSPDocLibCheckInType =  "minor" | "Minor" | "major" | "Major" | "overwrite" | "Overwrite";

type TUrlDomain = string;
type TUrlPath = string;
type SharePointId = number;
type TUrl2Api = `https://${TUrlDomain}/${TUrlPath}/_api`; // make this a regular expression


// Linearized Lists:  Name and ID only
type TNameIdQuickInfo = {
	Id: string | number;
	Name: string;
};

type TAuthor = {
	Id: number;
	PrincipalType: number;
	Email: string;
	IsSiteAdmin: boolean;
	LoginName: string;
	UserPrincipalName: string;
};

/***********************************************************
 *                    Site-Related
 ***********************************************************/

type TSpSite = {
	Id: string;
	Title: string; // name
	Url: string;
	Description: string;
	LastItemModifiedDate?: string;
	LastItemUserModifiedDate?: string;
	ServerRelativeUrl: string;
	Template?: string;
	WelcomePage?: string;
	RoleAssignments?: TSpRoleAssignments[];
	Created?: Date;
	Author?: TAuthor | null
};

type TEnhancedListQuickInfo = TListQuickInfo & {
	Url: string;
	SiteParentId?: string;
	SiteParentName?: string;
};

type TSpSiteExtra = TSpSite & {
	forApiPrefix: string;  // everything before '/_api'
   Name?: TSpSite["Title"];
	typeInSP?: "ROOT" | "SUBROOT" | "SUBSITE";
	tsType: "TSiteInfo";
	siteParent: TSiteInfo | undefined;
	subsites?: TSiteQuickInfo[] | TSiteInfo[];
	listslibs?: TEnhancedListQuickInfo[] | TListInfo[];
	users?: TUserQuickInfo[];
	groups?: TGroupQuickInfo[];
	parent?: TSiteInfo | null;
	referenceSite?: TSiteInfo;
	permissions?: TSiteListPermission[];
	server?: string; // this will be "https://<domain name>"
};

type TSiteInfo = TSpSiteExtra & TSpSite;

type TSpContentType = {};

type TSpEventReceivers = {};

type TSiteQuickInfo = TNameIdQuickInfo & {
	Url: string;
	SiteParentId?: string;
	SiteParentName?: string;
};



/***********************************************************
 *                    List-Related
 ***********************************************************/
type TSpList = {
	Id: string;
	Title: string;
	Description: string;
	BaseTemplate: number;
	BaseType?: number;
	Url: string;
	ItemCount: number;
	ServerRelativeUrl: string;
	Modified: string;
	Items?: TSpListItem[];
	//Views?: TSpView[];
	RootFolder?: TSpFolder;
	Created: Date;
	Author?: TAuthor | null;
	//EventReceivers?: TSpEventReceiver[];
	//Fields?: TSpField[];
	//Forms?: TSpForm[];
	RoleAssignments?: TSpRoleAssignments[];
	ParentWeb?: TSiteInfo;
	//InformationRightsManagementSettings?: TSpInformationRightsManagementSetting[];
	AllowContentTypes?: boolean;
	ContentTypesEnabled?: boolean;
	EnableVersioning?: boolean;
	ForceCheckout?: boolean;
};

type TSpListExtra = {
   Name?: TSpList["Title"];
   type?: keyof typeof ListTemplateType;
   siteParent: TSiteInfo;
	permissions?: TSiteListPermission[];
};

type TListInfo = TSpList & TSpListExtra;

type TListQuickInfo = TNameIdQuickInfo & {
	Url: string;
	SiteParentId?: string;
	SiteParentName?: string;
};


/*
This type has no properties to expose the files in the collection.
To get files, the `getByIndex` method should be used or an iterator in for loop
*/
type TSpFileCollection = {
	Count: number;  // count of files in collection
	AreItemsAvailable: boolean; // are collection files avalaible?
// 'include' used with load method to specify properties to retrieve
//    default are 'uniqueId,Name,ServerRelativeUrl'
	Include: string;  // specifies which properties of files to retrieve
// SP.CamlQuery object that specifies query to filter files in collection
//	Filter: TSpCamlQuery;
// An SP.ClientObjectCollection object to specify expressions used to retrieve
//  additional property of files in collection
//	RetrievalExpressions: TSpClientObjectCollection;
	SchemaXML: string; // contains XML schema for files in collection
};

type TSpFolderCollection = {
	Count: number;  // count of files in collection
	AreItemsAvailable: boolean; // are collection files avalaible?
// 'include' used with load method to specify properties to retrieve
//    default are 'uniqueId,Name,ServerRelativeUrl'
	Include: string;  // specifies which properties of files to retrieve
// SP.CamlQuery object that specifies query to filter files in collection
//	Filter: TSpCamlQuery;
// An SP.ClientObjectCollection object to specify expressions used to retrieve
//  additional property of files in collection
//	RetrievalExpressions: TSpClientObjectCollection;
	SchemaXML: string; // contains XML schema for files in collection
};

type TSpFolder = {
	Name: string;
	ItemCount: number;
	ServerRelativeUrl: string;
	ParentFolder: TListFolderInfo;
	ParentList: TListInfo;
// The SP.ListItem object that represents list item associated with folder.
// This property is only relevant for folders that have a list item associated with them,
//    such as folders in a document library.
	ListItemAllFields: TListItemInfo;
// A SP.FileCollection object that represents the collection of files in the folder.
	Files: TSpFileCollection;
// A SP.FolderCollection object that represents the collection of subfolders in the folder.
	Folders: TSpFolderCollection;
// A SP.PropertyValues object that represents the collection of custom properties
//   (metadata) associated with the folder.
//	Properties: TPropertyValueInfo;
}

type TListFolder = TSpFolder;

type TListFolderInfo = {

};

type TSpListItem = {
	Id: number;
	Title: string;
//	ContentType: TContentTypeInfo;
//	FileSystemObjectType: File | Folder;
//  An SP.File object that represents file associated with list item.
//   This property is only relevant for document libraries.
//	File: TFileInfo;
// An SP.Folder object that represents the folder associated with list item.
//  This property is only relevant for document libraries.
	Folder: TListFolderInfo;
//  The SP.List object that represents the list that contains list item.
	ParentList: TListInfo;
// A Boolean value that indicates whether the list item has attachments.
	Attachments: boolean;
// The SP.User object that represents the user who created the list item.
	Author: TUserInfo;
// The SP.User object that represents the user who last modified the list item.
	Editor: TUserInfo;
	Created: Date;
	Modified: Date;
// An SP.ListItem object that represents the parent folder of the list item.
// This property is only relevant for items in a document library.
	Item: TListFolderInfo;
};

type TListItemInfo = TSpListItem;

interface IListSetupMinimal {
	server: string;
	site: string;
	include?: string;
	protocol?: string;
}

interface IListSetupWithName extends IListSetupMinimal {
	listName: string;
	listGuid?: never;
}

interface IListSetupWithGuid extends IListSetupMinimal {
	listGuid: string;
	listName?: never;
}

type TListSetup = IListSetupWithName | IListSetupWithGuid;

type TLookupFieldInfo = {
	fieldDisplayName: string,
	fieldInternalName: string;
	fieldLookupFieldName: string;
	choices: {
		id: number;
		value: string;
	}[] | null;
};

type TSpSiteInfo = {
	name:string;
	serverRelativeUrl:string;
	id:string;
	template:string
};







interface ISPBaseListFieldProperties {
	__metadata: { type: string; }
	Id: number;
	InternalName: string;
	Description: string;
	DefaultValue: string;
	Title: string;
	FromBaseType: boolean;
	CanBeDeleted: boolean;
	Required: boolean;
	FieldTypeKind: number;
	TypeAsString: string;
	Hidden: boolean; // if true, it's a system list, if false user-defined list
	Choices?: {
		__metadata: {
			type: string;
		}
	}
}

interface ISPBaseListItemProperties {
	__metadata: { type: string; }
	Id: number;
	InternalName: string;
	File: { Name: string; }
	Title: string;
}




// declare function RequestAgain (next: string, data: any): Promise<any>;

/***********************************************************
 *                    Permissions-Related
 ***********************************************************/

// these are temporary types
type TRawWhatPermission = {
	What: {
		Type: "Site" | "List" | "Folder" | "Item";
		Info: {Name: string; Id: string | number}; // what the thing being access
	};
	How: {
		PermissionValue: TSpBasePermissions;
		PermissionFriendly: string;
	};
	RoleTypeKind: {
		Value: number;
		Text: string;
	};
};

type TRawWhoPermission = {
	Who: {
		Type: "User" | "Group";
		Info: {Name: string; Id: number };
	};
	How: {
		PermissionValue: TSpBasePermissions;
		PermissionFriendly: string;
	};
	RoleTypeKind: {
		Value: number;
		Text: string;
	};
};



type TUserGroupPermission = {
	What: {
		Type: "Site" | "List" | "Folder" | "Item";
		Info: {Name: string; Id: string | number}; // what the thing being access
	};
	How: {
		PermissionValue: TSpBasePermissions;
		PermissionFriendly: string;
		RoleTypeKind: {
			Value: number;
			Text: string;
		};
	}[];
};

type TSiteListPermission = {
	Who: {
		Type: "User" | "Group";
		Info: {Name: string; Id: number };
	};
	How: {
		PermissionValue: TSpBasePermissions;
		PermissionFriendly: string;
		RoleTypeKind: {
			Value: number;
			Text: string;
		};
	}[];
};

type TSpMetadata = {
	id?: string;
	type?: string;
	uri?: string;
}

type TSpRoleAssignments = {
	//__metadata?: {
	//	uri: `${TUrl2Api}/Web/RoleAssignments/GetByPrincipalId(${SharePointId})`,
	//	type: "SP.RoleAssignment"
	//};
	Member?: TSpMember;
	// | {
	//	__metadata?: TSpMetadata;
	//	__deferred?: {
	//		uri: `${TUrl2Api}/Web/RoleAssignments/GetByPrincipalId(${SharePointId})/Member`;
	//	};
	//	Id: number;
	//	PrincipalType: number;
	//	LoginName: string;
	//	Title: string;
	//	Description: string;
	//	AllowMembersEditMembership: boolean;
	//	AllowRequestToJoinLeave: boolean;
	//	AutoAcceptRequestToJoinLeave: boolean;
	//	OnlyAllowMembersViewMembership: boolean;
	//	OwnerTitle: string;
	//	RequestToJoinLeaveEmailSetting: string;
//	};
	RoleDefinitionBindings?: TSpRoleDefinitionBindings[]
	//| {
	//	__deferred: {
	//		uri: `${TUrl2Api}/Web/RoleAssignments/GetByPrincipalId(${SharePointId})/RoleDefinitionBindings`;
	//	};
	//}//;
	PrincipalId: number;
};

type TSpRoleDefinitionBindings = {
	Description: string;
	Id: number;
	Name: string;  // can probably use literals here or refer to a type defined literals
	BasePermissions: TSpBasePermissions;
	RoleTypeKind: number;  // likely an enum here
	// optional
	Hidden?: boolean;
	Order?: number;
	__metadata?: {
		uri: `${TUrl2Api}/Web/RoleDefinitions(${SharePointId})`;
		type: "SP.RoleDefinition";
	};
};



type SpRoleDefinition = {
	BasePermissions: TSpBasePermissions;
	Description: string;
	Hidden: boolean;
	Id: number;
	Name: string;
	Order: number;
	RoleTypeKind: RoleTypeKind;

	// methods
	get_basePermissions: () => TSpBasePermissions;
	set_basePermissions: (bpObj: TSpBasePermissions) => void;
	get_description: () => string;
	set_description: (description: string) => void;
	get_hidden: () => boolean;
	set_hidden: (hidden: boolean) => void;  // can hide the definition
	get_id: () => number;
	get_name: () => string;
	set_name: (name: string) => void;
	get_order: () => number; // 32 bit
	set_order: (order: number) => void;
};

type TSpBasePermissions =  {
	High: string;  // but an interger
	Low: string;
	__metadata?: {
	  type: "SP.BasePermissions";
	},
};

type TSpEffectiveBasePermissions = {
	d: {
		EffectiveBasePermissions: TSpBasePermissions;
	}
};


/***********************************************************
 *                    Group-Related
 ***********************************************************/


type TSpGroup = {
	Id: number; // The ID of the group.
	Title: string; // The title of the group.
	Description: string; // The description of the group.
	OwnerTitle: string; // The title of the user who is the owner of the group.
	OwnerEmail: string; // The email address of the user who is the owner of the group.
	PrincipalType: SpGroupType;
	AllowMembersEditMembership: boolean;
	AllowRequestToJoinLeave: boolean;
	AutoAcceptRequestToJoinLeave: boolean;
	OnlyAllowMembersViewMembership: boolean;
	RequestToJoinLeaveEmailSetting: boolean;
};


type TSpGroupExtra = {
	siteParentId: string;
	Users: TUserQuickInfo[]; /* An object that represents the collection of users that belong
		to the group. This object has a __deferred property that points to the URL
		of the users collection. The string type is in case permission to get users not possible.
		*/
	permissions?: TUserGroupPermission[];
	RoleAssignments: object;
};

export
type TGroupInfo = TSpGroup & TSpGroupExtra;

type TGroupQuickInfo = {
	Name: string;
	Id: number;
	SiteParentId: string;
};

type TSpOwner = {

};


   // this needs detailing
type TSpMember = {
// required
	Id: number;
	OwnerTitle: string;
	Title: string;
	Description: string;
	PrincipalType: number;
// optional
	__metadata?: {
		uri: `${TUrl2Api}/Web/RoleAssignments/GetByPrincipalId(${SharePointId})/Member`;
		type: "SP.Member"
	};
	Owner?: TSpOwner;
	Users?: TSpUser[];
	IsHiddenInUI?: boolean;
	LoginName?: string;
	AllowMembersEditMembership?: boolean;
	AllowRequestToJoinLeave?: boolean;
	AutoAcceptRequestToJoinLeave?: boolean;
	OnlyAllowMembersViewMembership?: boolean;
	RequestToJoinLeaveEmailSetting?: string;
};

/***********************************************************
 *                    User-Related
 ***********************************************************/
type TSpUser = {
	__metadata?: {
		uri: `${TUrl2Api}/Web/GetUserById(${SharePointId})`;
		type: "SP.User";
	};
	Id: number;
	LoginName: string;
	Title: string;
	PrincipalType: number;
	Email: string;
   Expiration?: string;
   IsEmailAuthenticationGuestUser?: boolean;
   IsShareByEmailGuestUser?: boolean;
   IsSiteAdmin?: boolean;
   IsHiddenInUI?: boolean;
   // what is UserId?
   UserId?: {
		__metadata: {
			type: "SP.UserIdInfo";
		},
		NameId: string;
		NameIdIssuer: string;
	},
	UserPrincipalName?: string;
	//Alerts?: TSpAlerts;
	Groups: {
		Id: string;
		Name: string;
	}[];
};

type TSpUserExtras = {
	Name: TSpUser["Title"];
	permissions?: TUserGroupPermission[];
}

type TUserInfo = TSpUser & TSpUserExtras;

type TUserQuickInfo = {
	Id: number;
	Name: string;
};
