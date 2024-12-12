
declare module 'core-js/features/object/from-entries';

import { THttpRequestProtocol } from "./SPHttpReqResp";
import {
	ListTemplateType,
	RoleTypeKind
} from "./SPRESTGlobals";

export {
	TParsedURL,
	TLookupFieldInfo,
	TNameIdQuickInfo,
	TListSetup,
	TSiteInfo,
	TUserInfo,
	TSiteQuickInfo,
	TUserQuickInfo,
	TSpListItem,
	TSPDocLibCheckInType,
	TSpField,
	TSpFieldTypes,
	TSpFieldStandardProperties,
	TSpList,
	TSpUser,
	TSpGroup,
	TListInfo,
	TSpSite,
	SPListRaw,
	SPListItemRaw,
	SPContentTypeRaw,
	SPFileRaw,
	SPFieldRaw,
	SPRootFolderRaw,
	SPWebInformationRaw,
	SPSiteRaw,
	TSpFolder,
	SPWebRaw,
	TUserGroupPermission,
	TAuthor,
	TSiteListPermission,
	TGroupInfo,
	TSpMember,
	TSpRoleAssignments,
	TSpMetadata,
	TEnhancedListQuickInfo,
	TSpRoleDefinitionBindings,
	SPListCustomColumns
};


/***********************************************************
 *                    Site-Related
 ***********************************************************/
type BaseSpDataType = {
	[key: string]: boolean | string | number | object | null | undefined;
};

type TSpSite = BaseSpDataType & {
		__metadata?: {
			type: "SP.Web";
		};
	Title: string; // name
	Url: string;
	ServerRelativeUrl: string;
	RoleAssignments?: TSpRoleAssignments[];
	Id: string;
	Description: string;
	LastItemModifiedDate?: string;
	LastItemUserModifiedDate?: string;
	Template?: string;
	WelcomePage?: string;
	Created?: Date;
	Author?: TAuthor | null
};


type TSiteInfo = TSpSite & {
	forApiPrefix: string;  // everything before '/_api'
   Name?: TSpSite["Title"];
	typeInSP?: "ROOT" | "SUBROOT" | "SUBSITE";
//	tsType: "TSiteInfo";
	siteParent?: TSiteInfo | undefined;
	subsites?: TSiteQuickInfo[] & TSiteInfo[];
	listslibs?: TEnhancedListQuickInfo[] & TListInfo[];
	users?: TUserQuickInfo[];
	groups?: TGroupQuickInfo[];
	parent?: TSiteInfo | null;
	referenceSite?: TSiteInfo;
	permissions?: TSiteListPermission[];
	server?: string; // this will be "https://<domain name>"
};


/*
type TSpSiteInfo = {
	name:string;
	serverRelativeUrl:string;
	id:string;
	template:string
};
*/

type TSpContentType = object;

type TSpEventReceivers = object;

type TSiteQuickInfo = TNameIdQuickInfo & {
	Url: string;
	SiteParentId?: string;
	SiteParentName?: string;
};



/***********************************************************
 *                    List-Related
 ***********************************************************/
type TSpList = BaseSpDataType & {
	__metadata: { type: "SP.List" };
	Id: string;
	Title: string;
	Url: string;
	Description: string;
	ItemCount: number;
	RoleAssignments?: TSpRoleAssignments[];
/*
	BaseTemplate: number;
	BaseType?: number;
	ServerRelativeUrl: string;
	Modified: string;
	Items?: TSpListItem[];
	//Views?: TSpView[];
	RootFolder?: TSpFolder;
	Created: Date;
	Author?: TAuthor | null;
	ListItemEntityTypeFullName: string;
	//EventReceivers?: TSpEventReceiver[];
	//Fields?: TSpField[];
	//Forms?: TSpForm[];

	ParentWeb?: TSiteInfo;
	//InformationRightsManagementSettings?: TSpInformationRightsManagementSetting[];
	AllowContentTypes?: boolean;
	ContentTypesEnabled?: boolean;
	EnableVersioning?: boolean;
	ForceCheckout?: boolean; */
};

type TSpListExtra = TSpList & {
	forApiPrefix: string;
   Name?: TSpList["Title"];
   type?: keyof typeof ListTemplateType;
   siteParent: TSiteInfo;
	permissions?: TSiteListPermission[];
};

type TListInfo = TSpListExtra;

type TEnhancedListQuickInfo = {
	Id: string;
	Name?: string;
	Url: string;
	SiteParentId?: string;
	SiteParentName?: string;
};

// Linearized Lists:  Name and ID only
type TNameIdQuickInfo = {
	Id: string | number;
	Name: string;
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

type TListFolderInfo = object;

type SPListCustomColumns =	{[key: string]: string | boolean | number};

type TSpListItem = {
	Id: number;
	Title: string;
//	ContentType: TContentTypeInfo;
//	FileSystemObjectType: File | Folder;
//  An SP.File object that represents file associated with list item.
//   This property is only relevant for document libraries.
	File: SPFileRaw;
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
} & SPListCustomColumns;


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

type TSpField = {
	__metadata: { type: "SP.Field" }
	Id: number;
	InternalName: string;
	Description: string;
	DefaultValue: string;
	Title: string;
	Group: string; //  "Custom Columns",
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

type TSpFieldStandardProperties = {
	AutoIndexed: boolean;
	CanBeDeleted: boolean;  // true = field can be deleted
				// returns false if FromBaseType == true || Sealed == True
	ClientSideComponentId: string; //
	ClientSideComponentProperties: object;
	ClientValidationFormula: string;
	ClientValidationMessage: string;
	CustomFormatter: string;
	DefaultFormula: string; // string: gets/sets default formula for calculated field
	DefaultValue: string; // string: gets/sets default value for field
	Description: string;
	Direction: string;  // string: "none","LTR","RTL" gets/sets reading order of field
	EnforceUniqueValues: boolean; // boolean: gets/sets to force uniqueness of column values (false=default)
	EntityPropertyName: string; // string: gets name of entity proper for list item entity using field
	Filterable: boolean; // boolean: gets whether field can be filtered
	FromBaseType: boolean; // boolean: gets whetehr field derives from base field type
	Group: string; // string: gets/sets column group to which field belongs
		// Groups: Base, Core Contact and Calendar, Core Document, Core Task and Issue, Custom, Extended, _Hidden, Picture
	Hidden: boolean; // boolean: specifies field display in list
	Id: string; // guid of field
	Indexed: boolean; // boolean: gets/sets if field is indexed
	IndexStatus: string;
	InternalName: string; // string: gets internal name of field
	IsModern: boolean;
	JSLink: string;  // string: gets/sets name of JS file(s) [separated by '|'] for client rendering logic of field
	PinnedToFiltersPane: boolean;
	ReadOnlyField: boolean; // boolean: gets/sets whether field can be modifed
	Required: boolean; // boolean: gets/sets whether user must enter value on New and Edit forms
	SchemaXml: string; // string: gets/sets schema that defines the field
	Scope: string; // string: gets Web site-relative path to list in which field collection is used
	Sealed: boolean; // boolean: gets/sets whether field type can be parent of custom derived field type
	ShowInFiltersPane: boolean;
	Sortable: boolean; // boolean: gets boolean whether field can be used in sort
	StaticName: string; // string: gets/sets static name
	Title: string; // string: gets/sets display name for field
	FieldTypeKind: string; //
	TypeAsString: string; // string: gets the type of field as a string value
	TypeDisplayName: string; // string: gets display name of the field type
	TypeShortDescription: string; // string: gets the description of the field
	ValidationFormula: string; // string: gets/sets formula referenced by field, evaluated when list item added/updated
	ValidationMessage: string; // string: gets/sets message to display if validation fails
};

type TSpFieldTypeDef = {
	name: string;
	metadataType?: string;
	typeId: number;
	extraProperties?: string[];
	//standardProperties: typeof TSpFieldStandarddProperties;
};

type TSpFieldTypes = TSpFieldTypeDef[];


type TLookupFieldInfo = {
	fieldDisplayName: string,
	fieldInternalName: string;
	fieldLookupFieldName: string;
	choices: {
		id: number;
		value: string;
	}[] | null;
	SPproperties: {
		LookupField: string; // gets/set value of internal field name of field used as lookup value
		LookupList: string; // string: gets/sets value of list identifier of list containing field to lookup values
		LookupWebId: string;	// SP.Guid: gets/sets value of GUID identifying site containing list which has field for lookup values
		PrimaryFieldId: string;		// string: gets/sets value specifying primary look field identifier if there is a dependent
					// lookup field; otherwise empty string
	};
};

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
		// TSmallSiteInfo | TSmallListInfo | TSmallListFolderInfo | TSmallListItemInfo;
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
	RoleTypeKind: {
		Value: number;
		Text: string;
	};
	How: {
		PermissionValue: TSpBasePermissions;
		PermissionFriendly: string;
	};
};
type TSiteListPermission = {
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



interface SpRoleDefinition {
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
}

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


type TSpGroup = BaseSpDataType & {
		__metadata?: {
			type: "SP.Group";
		};
	Title: string; // The title of the group.
	Id: number; // The ID of the group.
	Name: string;
/*

	Description: string; // The description of the group.
	OwnerTitle: string; // The title of the user who is the owner of the group.
	OwnerEmail: string; // The email address of the user who is the owner of the group.
	PrincipalType: SpGroupType;
	AllowMembersEditMembership: boolean;
	AllowRequestToJoinLeave: boolean;
	AutoAcceptRequestToJoinLeave: boolean;
	OnlyAllowMembersViewMembership: boolean;
	RequestToJoinLeaveEmailSetting: boolean; */
};


type TGroupInfo = TSpGroup & {
	forApiPrefix: string
	siteParentId: string;
	Users: TUserQuickInfo[];  //An object that represents the collection of users that belong
//		to the group. This object has a __deferred property that points to the URL
//		of the users collection. The string type is in case permission to get users not possible.
	permissions?: TUserGroupPermission[];
//	RoleAssignments: object;
};

type TGroupQuickInfo = {
	Name: string;
	Id: number;
	SiteParentId: string;
};

type TSpOwner = object;


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
type TSpUser = BaseSpDataType & {
	__metadata?: {
		uri: `${TUrl2Api}/Web/GetUserById(${SharePointId})`;
		type: "SP.User";
	};

	Id: number;
	Title: string;
/*	LoginName: string;
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
	}[]; */
};

type TUserInfo = TSpUser & {
	Name: TSpUser["Title"];
	permissions?: TUserGroupPermission[];
};

type TUserQuickInfo = {
	Id: number;
	Name: string;
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
	query: unknown[] | null;
};

type TSPDocLibCheckInType =  "minor" | "Minor" | "major" | "Major" | "overwrite" | "Overwrite";

type TUrlDomain = string;
type TUrlPath = string;
type SharePointId = number;
type TUrl2Api = `https://${TUrlDomain}/${TUrlPath}/_api`; // make this a regular expression



type TAuthor = {
	Id: number;
	PrincipalType: number;
	Email: string;
	IsSiteAdmin: boolean;
	LoginName: string;
	UserPrincipalName: string;
};

type SPSiteRaw = {
	__metadata: {
		id: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site",
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site",
			type: string; // "SP.Site"
	};
	Audit: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/Audit"
		}
	};
	CustomScriptSafeDomains: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/CustomScriptSafeDomains"
		}
	};
	EventReceivers: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/EventReceivers"
		}
	};
	Features: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/Features"
		}
	};
	HubSiteSynchronizableVisitorGroup: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/HubSiteSynchronizableVisitorGroup"
		}
	};
	Owner: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/Owner"
		}
	};
	RecycleBin: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/RecycleBin"
		}
	};
	RootWeb: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/RootWeb"
		}
	};
	SecondaryContact: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/SecondaryContact"
		}
	};
	UserCustomActions: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/UserCustomActions"
		}
	};
	VersionPolicyForNewLibrariesTemplate: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/site/VersionPolicyForNewLibrariesTemplate"
		}
	};
	AllowCreateDeclarativeWorkflow: boolean; // false,
	AllowDesigner: boolean; // true,
	AllowMasterPageEditing: boolean; // false,
	AllowRevertFromTemplate: boolean; // false,
	AllowSaveDeclarativeWorkflowAsTemplate: boolean; // false,
	AllowSavePublishDeclarativeWorkflow: boolean; // false,
	AllowSelfServiceUpgrade: boolean; // true,
	AllowSelfServiceUpgradeEvaluation: boolean; // true,
	AuditLogTrimmingRetention: number; // 90,
	ChannelGroupId: string; // "00000000-0000-0000-0000-000000000000",
	Classification: string; // "",
	CompatibilityLevel: number; // 15,
	CurrentChangeToken: {
		__metadata: {
			type: string; // "SP.ChangeToken"
		};
		StringValue: string; // "1;1;1a546ef3-8c5d-496c-b76c-77e055aeefb0;638498300156900000;341663080"
	};
	DisableAppViews: boolean; // false,
	DisableCompanyWideSharingLinks: boolean; // false,
	DisableFlows: boolean; // false,
	ExternalSharingTipsEnabled: boolean; // false,
	GeoLocation: string; // "NAM",
	GroupId: string; // "a3512a3a-40ac-48eb-8216-0cf8565a52f8",
	HubSiteId: string; // "00000000-0000-0000-0000-000000000000",
	Id: string; // "1a546ef3-8c5d-496c-b76c-77e055aeefb0",
	SensitivityLabelId: null,
	SensitivityLabel: string; // "00000000-0000-0000-0000-000000000000",
	IsHubSite: boolean; // false,
	IsRestrictedAccessControlPolicyEnforcedOnSite: boolean; // false,
	LockIssue: null,
	MaxItemsPerThrottledOperation: number; // 5000,
	MediaTranscriptionDisabled: boolean; // false,
	NeedsB2BUpgrade: boolean; // false,
	ResourcePath: {
		__metadata: {
			type: string; // "SP.ResourcePath"
		};
		DecodedUrl: string; // "https://mhalloran.sharepoint.com/sites/mhteam"
	};
	PrimaryUri: string; // "https://mhalloran.sharepoint.com/sites/mhteam",
	ReadOnly: boolean; // false,
	RequiredDesignerVersion: string; // "15.0.0.0",
	SandboxedCodeActivationCapability: number; // 2,
	ServerRelativeUrl: string; // "/sites/mhteam",
	ShareByEmailEnabled: boolean; // true,
	ShareByLinkEnabled: boolean; // false,
	ShowUrlStructure: boolean; // false,
	TrimAuditLog: boolean; // true,
	UIVersionConfigurationEnabled: boolean; // false,
	UpgradeReminderDate: string; // "1899-12-30T00:00:00",
	UpgradeScheduled: boolean; // false,
	UpgradeScheduledDate: string; // "1753-01-01T00:00:00",
	Upgrading: boolean; // false,
	Url: string; // "https://mhalloran.sharepoint.com/sites/mhteam",
	WriteLocked: boolean; // false`
};

type SPWebRaw = {

	__metadata: {
		id: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web",
		type: string; // "SP.Web"
	},
	FirstUniqueAncestorSecurableObject: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/FirstUniqueAncestorSecurableObject"
		}
	},
	RoleAssignments: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/RoleAssignments"
		}
	},
	AccessRequestsList: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AccessRequestsList"
		}
	},
	Activities: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Activities"
		}
	},
	ActivityLogger: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/ActivityLogger"
		}
	},
	Alerts: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Alerts"
		}
	},
	AllProperties: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AllProperties"
		}
	},
	AppTiles: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AppTiles"
		}
	},
	AssociatedMemberGroup: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AssociatedMemberGroup"
		}
	},
	AssociatedOwnerGroup: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AssociatedOwnerGroup"
		}
	},
	AssociatedVisitorGroup: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AssociatedVisitorGroup"
		}
	},
	Author: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Author"
		}
	},
	AvailableContentTypes: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AvailableContentTypes"
		}
	},
	AvailableFields: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/AvailableFields"
		}
	},
	CanModernizeHomepage: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/CanModernizeHomepage"
		}
	},
	CardDesigns: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/CardDesigns"
		}
	},
	ClientWebParts: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/ClientWebParts"
		}
	},
	ContentTypes: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/ContentTypes"
		}
	},
	CurrentUser: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/CurrentUser"
		}
	},
	DataLeakagePreventionStatusInfo: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/DataLeakagePreventionStatusInfo"
		}
	},
	DescriptionResource: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/DescriptionResource"
		}
	},
	EventReceivers: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/EventReceivers"
		}
	},
	Features: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Features"
		}
	},
	Fields: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Fields"
		}
	},
	Folders: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Folders"
		}
	},
	HostedApps: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/HostedApps"
		}
	},
	Lists: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists"
		}
	},
	ListTemplates: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/ListTemplates"
		}
	},
	MultilingualSettings: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/MultilingualSettings"
		}
	},
	Navigation: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Navigation"
		}
	},
	OneDriveSharedItems: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/OneDriveSharedItems"
		}
	},
	ParentWeb: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/ParentWeb"
		}
	} & SPWebInformationRaw,
	PushNotificationSubscribers: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/PushNotificationSubscribers"
		}
	},
	RecycleBin: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/RecycleBin"
		}
	},
	RegionalSettings: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/RegionalSettings"
		}
	},
	RoleDefinitions: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/RoleDefinitions"
		}
	},
	RootFolder: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/RootFolder"
		}
	},
	SiteCollectionAppCatalog: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/SiteCollectionAppCatalog"
		}
	},
	SiteGroups: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/SiteGroups"
		}
	},
	SiteUserInfoList: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/SiteUserInfoList"
		}
	},
	SiteUsers: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/SiteUsers"
		}
	},
	TenantAppCatalog: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/TenantAppCatalog"
		}
	},
	ThemeInfo: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/ThemeInfo"
		}
	},
	TitleResource: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/TitleResource"
		}
	},
	UserCustomActions: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/UserCustomActions"
		}
	},
	Webs: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Webs"
		}
	},
	WebInfos: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/WebInfos"
		}
	},
	WorkflowAssociations: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/WorkflowAssociations"
		}
	},
	WorkflowTemplates: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/WorkflowTemplates"
		}
	},
	AllowRssFeeds: boolean; // true,
	AlternateCssUrl: string; // "",
	AppInstanceId: string; // "00000000-0000-0000-0000-000000000000",
	ClassicWelcomePage: string | null,
	Configuration: number; // 3,
	Created: string; // "2024-01-23T15:53:53",
	CurrentChangeToken: {
		__metadata: {
		type: string; // "SP.ChangeToken"
		},
		StringValue: string; // "1;2;86b61fa8-cd24-4cd1-8189-ca718dac33b8;638498309655030000;341663327"
	},
	CustomMasterUrl: string; // "/sites/mhteam/Projects/_catalogs/masterpage/seattle.master",
	Description: string; // "",
	DesignPackageId: string; // "00000000-0000-0000-0000-000000000000",
	DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: boolean; // false,
	EnableMinimalDownload: boolean; // false,
	FooterEmphasis: number; // 0,
	FooterEnabled: boolean; // false,
	FooterLayout: number; // 0,
	HeaderEmphasis: number; // 0,
	HeaderLayout: number; // 0,
	HideTitleInHeader: boolean; // false,
	HorizontalQuickLaunch: boolean; // false,
	Id: string; // "86b61fa8-cd24-4cd1-8189-ca718dac33b8",
	IsEduClass: boolean; // false,
	IsEduClassProvisionChecked: boolean; // false,
	IsEduClassProvisionPending: boolean; // false,
	IsHomepageModernized: boolean; // false,
	IsMultilingual: boolean; // true,
	IsRevertHomepageLinkHidden: boolean; // false,
	Language: number; // 1033,
	LastItemModifiedDate: string; // "2024-04-27T13:27:34Z",
	LastItemUserModifiedDate: string; // "2024-04-27T13:27:34Z",
	LogoAlignment: number; // 0,
	MasterUrl: string; // "/sites/mhteam/Projects/_catalogs/masterpage/seattle.master",
	MegaMenuEnabled: boolean; // false,
	NavAudienceTargetingEnabled: boolean; // false,
	NoCrawl: boolean; // false,
	ObjectCacheEnabled: boolean; // false,
	OverwriteTranslationsOnChange: boolean; // false,
	ResourcePath: {
		__metadata: {
		type: string; // "SP.ResourcePath"
		},
		DecodedUrl: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects"
	},
	QuickLaunchEnabled: boolean; // true,
	RecycleBinEnabled: boolean; // true,
	SearchScope: number; // 0,
	ServerRelativeUrl: string; // "/sites/mhteam/Projects",
	SiteLogoUrl: string; // "/sites/mhteam/_api/GroupService/GetGroupImage?id='9bb5afd2-4cfd-4daf-932d-6901e6ee00ea'&hash=637669625585985283",
	SyndicationEnabled: boolean; // true,
	TenantAdminMembersCanShare: number; // 0,
	Title: string; // "Projects",
	TreeViewEnabled: boolean; // false,
	uIVersion: number; // 15,
	uIVersionConfigurationEnabled: boolean; // false,
	url: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects",
	WebTemplate: string; // "STS",
	WelcomePage: string; // "SitePages/Home.aspx"
};

type SPFieldRaw = {  // SP.Field
	__metadata: {
		id: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Fields(guid'03e45e84-1992-4d42-9116-26f756012634')",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Fields(guid'03e45e84-1992-4d42-9116-26f756012634')",
		type: string; // "SP.Field"
	};
	DescriptionResource: {
		__deferred: {
			uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Fields(guid'03e45e84-1992-4d42-9116-26f756012634')/DescriptionResource"
		}
	};
	TitleResource: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Fields(guid'03e45e84-1992-4d42-9116-26f756012634')/TitleResource"
		}
	};
	AutoIndexed: boolean; // false,
	CanBeDeleted: boolean; // false,
	ClientSideComponentId: string; // "00000000-0000-0000-0000-000000000000",
	ClientSideComponentProperties: object; //| null;
	ClientValidationFormula: string | null; // null,
	ClientValidationMessage: string | null; // null
	CustomFormatter: string | null;
	DefaultFormula: string | null;
	DefaultValue: string | null;
	Description: string; // ""
	Direction: string;  // "none",
	EnforceUniqueValues: boolean; // false,
	EntityPropertyName: string; // "ContentTypeId",
	Filterable: boolean; // true,
	FromBaseType: boolean;  // true,
	Group: string; //  "Custom Columns",
	Hidden: boolean; //  true,
	Id: string; // "03e45e84-1992-4d42-9116-26f756012634",
	Indexed: boolean; // false,
	IndexStatus: number; //  0,
	InternalName: string; // "ContentTypeId",
	IsModern: boolean; // false,
	JSLink: string | null;
	PinnedToFiltersPane: boolean; // false,
	ReadOnlyField: boolean; // true,
	Required: boolean; // false,
	SchemaXml: string; //  "<Field ID=\"{03e45e84-1992-4d42-9116-26f756012634}\" RowOrdinal=\"0\" Type=\"ContentTypeId\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" Hidden=\"TRUE\" DisplayName=\"Content Type ID\" Name=\"ContentTypeId\" DisplaceOnUpgrade=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentTypeId\" ColName=\"tp_ContentTypeId\" FromBaseType=\"TRUE\" />",
	Scope: string; // "/sites/mhteam/Projects/Biochemistry Reference Calculator",
	Sealed: boolean; // true,
	ShowInFiltersPane: number; // 0,
	Sortable: boolean; // true,
	StaticName: string; // "ContentTypeId",
	Title: string; // "Content Type ID",
	FieldTypeKind: number; // 25,
	TypeAsString: string; // "ContentTypeId",
	TypeDisplayName: string; // "Content Type Id",
	TypeShortDescription: string; // "Content Type Id",
	ValidationFormula: string | null;
	ValidationMessage: string | null;
}

type SPListRaw = {
	__metadata: {
		id: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')",
		etag: string; // "\"9\"",
		type: string; // "SP.List"
	},
	FirstUniqueAncestorSecurableObject: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/FirstUniqueAncestorSecurableObject"
		}
	},
	RoleAssignments: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RoleAssignments"
		}
	},
	Author: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Author"
		}
	},
	ContentTypes: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes"
		};
		results: SPContentTypeRaw[];
	};
	CreatablesInfo: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/CreatablesInfo"
		}
	},
	DefaultView: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/DefaultView"
		}
	},
	DescriptionResource: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/DescriptionResource"
		}
	},
	EventReceivers: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/EventReceivers"
		}
	},
	Fields: {
		__deferred?: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Fields"
		};
		results: SPFieldRaw[];
	};
	Forms: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Forms"
		}
	},
	InformationRightsManagementSettings: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/InformationRightsManagementSettings"
		}
	},
	Items: {
		__deferred?: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items"
		};
		results?: SPListItemRaw[];
	},
	ParentWeb: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ParentWeb"
		}
	},
	RootFolder: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder"
		};
	} | SPRootFolderRaw,
	Subscriptions: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Subscriptions"
		}
	},
	TitleResource: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/TitleResource"
		}
	},
	UserCustomActions: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/UserCustomActions"
		}
	},
	VersionPolicies: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/VersionPolicies"
		}
	},
	Views: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Views"
		}
	},
	WorkflowAssociations: {
		__deferred: {
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/WorkflowAssociations"
		}
	},
	AllowContentTypes: boolean; // true,
	BaseTemplate: number; // 101,
	BaseType: number; //1,
	ContentTypesEnabled: boolean; // false,
	CrawlNonDefaultViews: boolean; // false,
	Created: string; // "2024-01-23T15:55:35Z",
	CurrentChangeToken: {
		__metadata: {
			type: string; // "SP.ChangeToken"
		},
		StringValue: string; // "1;3;f2735f63-b609-4f6e-b4a6-4b8838dde8cd;638497749907830000;341609631"
	},
	DefaultContentApprovalWorkflowId: string; // "00000000-0000-0000-0000-000000000000",
	DefaultItemOpenUseListSetting: boolean; // false,
	Description: string; // "",
	Direction: string; // "none",
	DisableCommenting: boolean; // false,
	DisableGridEditing: boolean; // false,
	DocumentTemplateUrl: string; // "/sites/mhteam/Projects/Biochemistry Reference Calculator/Forms/template.dotx",
	DraftVersionVisibility: number; //0,
	EnableAttachments: boolean; // false,
	EnableFolderCreation: boolean; // true,
	EnableMinorVersions: boolean; // false,
	EnableModeration: boolean; // false,
	EnableRequestSignOff: boolean; // true,
	EnableVersioning: boolean; // true,
	EntityTypeName: string; // "Biochemistry_x0020_Reference_x0020_Calculator",
	ExemptFromBlockDownloadOfNonViewableFiles: boolean; // false,
	FileSavePostProcessingEnabled: boolean; // false,
	ForceCheckout: boolean; // false,
	HasExternalDataSource: boolean; // false,
	Hidden: boolean; // false,
	Id: string; // "f2735f63-b609-4f6e-b4a6-4b8838dde8cd",
	ImagePath: {
		__metadata: {
			type: string; // "SP.ResourcePath"
		},
		DecodedUrl: string; // "/_layouts/15/images/itdl.png?rev=47"
	},
	ImageUrl: string; // "/_layouts/15/images/itdl.png?rev=47",
	DefaultSensitivityLabelForLibrary: string; // "",
	SensitivityLabelToEncryptOnDownloadForLibrary: null,
	IrmEnabled: boolean; // false,
	IrmExpire: boolean; // false,
	IrmReject: boolean; // false,
	IsApplicationList: boolean; // false,
	IsCatalog: boolean; // false,
	IsPrivate: boolean; // false,
	ItemCount: number; // 32,
	LastItemDeletedDate: string; // "2024-03-06T16:07:40Z",
	LastItemModifiedDate: string; // "2024-04-26T23:03:14Z",
	LastItemUserModifiedDate: string; // "2024-03-06T16:36:38Z",
	ListExperienceOptions: number; //0,
	ListItemEntityTypeFullName: string; // "SP.Data.Biochemistry_x0020_Reference_x0020_CalculatorItem",
	MajorVersionLimit: number; //500,
	MajorWithMinorVersionsLimit: number; //0,
	MultipleDataList: boolean; // false,
	NoCrawl: boolean; // false,
	ParentWebPath: {
		__metadata: {
			type: string; // "SP.ResourcePath"
		},
		DecodedUrl: string; // "/sites/mhteam/Projects"
	},
	ParentWebUrl: string; // "/sites/mhteam/Projects",
	ParserDisabled: boolean; // false,
	ServerTemplateCanCreateFolders: boolean; // true,
	TemplateFeatureId: string; // "00bfea71-e717-4e80-aa17-d0c71b360101",
	Title: string; // "Biochemistry Reference Calculator"
};

type SPListItemRaw = {  // SP.<nameOfList>_Item
	__metadata: {
		id: string; // "091ac82e-129c-4706-9546-9bb827114791",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)",
		etag: string; // "\"2\"",
		type: string; // "SP.Data.Biochemistry_x0020_Reference_x0020_CalculatorItem"
	};
	FirstUniqueAncestorSecurableObject: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/FirstUniqueAncestorSecurableObject"
		}
	},
	RoleAssignments: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/RoleAssignments"
		}
	},
	AttachmentFiles: {
		__deferred: {
			uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/AttachmentFiles"
		}
	};
	ContentType: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/ContentType"
		}
	};
	GetDlpPolicyTip: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/GetDlpPolicyTip"
		}
	},
	FieldValuesAsHtml: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/FieldValuesAsHtml"
		}
	},
	FieldValuesAsText: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/FieldValuesAsText"
		}
	},
	FieldValuesForEdit: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/FieldValuesForEdit"
		}
	};
	File: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/File"
		}
	} & SPFileRaw,
	Folder: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/Folder"
		}
	};
	LikedByInformation: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/LikedByInformation"
		}
	};
	ParentList: {
		__deferred: {
			uri:  string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/ParentList"
		}
	};
	Properties: {
		__deferred: {
			uri: string; //  "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/Properties"
		}
	};
	Versions: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(52)/Versions"
		}
	};
	FileSystemObjectType: number; //  0;
	Id: number; // 52;
	ServerRedirectedEmbedUri: string | null;
	ServerRedirectedEmbedUrl: string; // "";
	ContentTypeId: string; // "0x010100AEBD6EF2AB1818488F56CB69C8BF8C62";
	OData__ColorTag: string | null;
	ComplianceAssetId: string | null;
	Title: string | null,
	ID: number; // 52,
	Created: string; // "2024-03-06T15:05:24Z",
	AuthorId: number; // 11,
	Modified: string; // "2024-03-03T21:50:27Z",
	EditorId: number; // 11,
	OData__CopySource: string | null,
	CheckoutUserId: string | null,
	OData__UIVersionString: string; // "1.0",
	GUID: string; // "d10007bf-c76b-4e43-b922-41dfe8c09b98"
};

type SPFileRaw = { // SP.File
	__metadata: {
		id: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File",
		type: string; // "SP.File"
	};
	Author: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/Author"
		}
	};
	CheckedOutByUser: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/CheckedOutByUser"
		}
	},
	EffectiveInformationRightsManagementSettings: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/EffectiveInformationRightsManagementSettings"
		}
	},
	InformationRightsManagementSettings: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/InformationRightsManagementSettings"
		}
	},
	ListItemAllFields: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/ListItemAllFields"
		}
	} & SPListItemRaw,
	LockedByUser: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/LockedByUser"
		}
	},
	ModifiedBy: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/ModifiedBy"
		}
	},
	Properties: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/Properties"
		}
	},
	VersionEvents: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/VersionEvents"
		}
	},
	VersionExpirationReport: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/VersionExpirationReport"
		}
	},
	Versions: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/Items(50)/File/Versions"
		}
	},
	CheckInComment: string; // "",
	CheckOutType: number;  //2,
	ContentTag: string; // "{45B78441-1D3C-4A4C-8C21-1CBD42547D01},2,2",
	CustomizedPageStatus: number; // 0,
	ETag: string; // "\"{45B78441-1D3C-4A4C-8C21-1CBD42547D01},2\"",
	Exists: true,
	ExistsAllowThrowForPolicyFailures: true,
	ExistsWithException: true,
	IrmEnabled: false,
	Length: string; // "8375",
	Level: number; // 1,
	LinkingUri: null,
	LinkingUrl: string; // "",
	MajorVersion: number; // 1,
	MinorVersion: number; // 0,
	Name: string; // "Dilutions.xhtml",
	ServerRelativeUrl: string; // "/sites/mhteam/Projects/Biochemistry Reference Calculator/html/Dilutions.xhtml",
	TimeCreated: string; // "2024-03-06T15:05:24Z",
	TimeLastModified: string; // "2024-03-04T13:48:57Z",
	Title: null,
	UIVersion: number // 512,
	UIVersionLabel: string; // "1.0",
	UniqueId: string; // "45b78441-1d3c-4a4c-8c21-1cbd42547d01"
};

type SPContentTypeRaw = { // SP.ContentType
	__metadata: {
		id: string // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')",
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')",
		type: string; // SP.ContentType"
	};
	DescriptionResource: {
		__deferred: {
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')/DescriptionResource"
		}
	};
	FieldLinks: {
		__deferred: {
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')/FieldLinks"
		}
	};
	Fields: {
		__deferred: {
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')/Fields"
		}
	};
	NameResource: {
		__deferred: {
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')/NameResource"
		}
	};
	Parent: {
		__deferred: {
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')/Parent"
		}
	};
	WorkflowAssociations: {
		__deferred: {
		uri: string; //"https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/ContentTypes('0x010100AEBD6EF2AB1818488F56CB69C8BF8C62')/WorkflowAssociations"
		}
	};
	ClientFormCustomFormatter: string; // ",
	Description: string; // Create a new document.",
	DisplayFormClientSideComponentId: string; // ",
	DisplayFormClientSideComponentProperties: string; // ",
	DisplayFormTarget: number; // 0,
	DisplayFormTemplateName: string; // DocumentLibraryForm",
	DisplayFormUrl: string; // ",
	DocumentTemplate: string; // /sites/mhteam/Projects/Biochemistry Reference Calculator/Forms/template.dotx",
	DocumentTemplateUrl: string; // /sites/mhteam/Projects/Biochemistry Reference Calculator/Forms/template.dotx",
	EditFormClientSideComponentId: string; // ",
	EditFormClientSideComponentProperties: string; // ",
	EditFormTarget: number; // 0,
	EditFormTemplateName: string; // DocumentLibraryForm",
	EditFormUrl: string; // ",
	Group: string; // Document Content Types",
	Hidden: boolean; // false,
	Id: {
		__metadata: {
			type: string; // SP.ContentTypeId"
		},
		StringValue: string; // 0x010100AEBD6EF2AB1818488F56CB69C8BF8C62"
	};
	JSLink: string; // ",
	MobileDisplayFormUrl: string; // ",
	MobileEditFormUrl: string; // ",
	MobileNewFormUrl: string; // ",
	Name: string; // Document",
	NewFormClientSideComponentId: string | null,
	NewFormClientSideComponentProperties: string; // ",
	NewFormTarget: number; // 0,
	NewFormTemplateName: string; // DocumentLibraryForm",
	NewFormUrl: string; // ",
	ReadOnly: false,
	SchemaXml: string; // <ContentType ID=\"0x010100AEBD6EF2AB1818488F56CB69C8BF8C62\" Name=\"Document\" Group=\"Document Content Types\" Description=\"Create a new document.\" V2ListTemplateName=\"doclib\" Version=\"6\" DelayActivateTemplateBinding=\"GROUP,SPSPERS,SITEPAGEPUBLISHING\" FeatureId=\"{c94c1702-30a7-454c-be15-5a895223428d}\" FeatureIds=\"{c94c1702-30a7-454c-be15-5a895223428d};{695b6570-a48b-4a8e-8ea5-26ea7fc1d162}\"><Fields><Field ID=\"{c042a256-787d-4a6f-8a8a-cf6ab767f12d}\" Type=\"Computed\" DisplayName=\"Content Type\" Name=\"ContentType\" DisplaceOnUpgrade=\"TRUE\" RenderXMLUsingPattern=\"TRUE\" Sortable=\"FALSE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"ContentType\" Group=\"_Hidden\" PITarget=\"MicrosoftWindowsSharePointServices\" PIAttribute=\"ContentTypeID\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ContentTypeId\"/></FieldRefs><DisplayPattern><MapToContentType><Column Name=\"ContentTypeId\"/></MapToContentType></DisplayPattern></Field><Field ID=\"{5f47e085-2150-41dc-b661-442f3027f552}\" ReadOnly=\"TRUE\" Type=\"Computed\" Name=\"SelectFilename\" DisplayName=\"Select\" Hidden=\"TRUE\" CanToggleHidden=\"TRUE\" Sortable=\"FALSE\" Filterable=\"FALSE\" AuthoringInfo=\"(web part connection)\" HeaderImage=\"blank.gif\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"SelectFilename\" FromBaseType=\"TRUE\"><FieldRefs><FieldRef Name=\"ID\"/></FieldRefs><DisplayPattern><IfEqual><Expr1><GetVar Name=\"SelectedID\"/></Expr1><Expr2><Column Name=\"ID\"/></Expr2><Then><HTML><![CDATA[<img align=\"absmiddle\" style=\"cursor: pointer\" src=\"/_layouts/15/images/rbsel.gif?rev=47\" alt=\"]]></HTML><HTML>Selected</HTML><HTML><![CDATA[\"/>]]></HTML></Then><Else><HTML><![CDATA[<a href=\"javascript:SelectField(']]></HTML><GetVar Name=\"View\"/><HTML><![CDATA[',']]></HTML><ScriptQuote NotAddingQuote=\"TRUE\"><Column Name=\"ID\"/></ScriptQuote><HTML><![CDATA[');return false;\" onclick=\"javascript:SelectField(']]></HTML><GetVar Name=\"View\"/><HTML><![CDATA[',']]></HTML><ScriptQuote NotAddingQuote=\"TRUE\"><Column Name=\"ID\"/></ScriptQuote><HTML><![CDATA[');return false;\" target=\"_self\">]]></HTML><HTML><![CDATA[<img border=\"0\" align=\"absmiddle\" style=\"cursor: pointer\" src=\"/_layouts/15/images/rbunsel.gif?rev=47\"  alt=\"]]></HTML><HTML>Normal</HTML><HTML><![CDATA[\"/>]]></HTML><HTML><![CDATA[</a>]]></HTML></Else></IfEqual></DisplayPattern></Field><Field ID=\"{8553196d-ec8d-4564-9861-3dbe931050c8}\" ShowInFileDlg=\"FALSE\" ShowInVersionHistory=\"FALSE\" Type=\"File\" Name=\"FileLeafRef\" DisplayName=\"Name\" AuthoringInfo=\"(for use in forms)\" List=\"Docs\" FieldRef=\"ID\" ShowField=\"LeafName\" JoinColName=\"DoclibRowId\" JoinRowOrdinal=\"0\" JoinType=\"INNER\" Required=\"TRUE\" NoCustomize=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"FileLeafRef\" FromBaseType=\"TRUE\"/><Field ID=\"{8c06beca-0777-48f7-91c7-6da68bc07b69}\" ColName=\"tp_Created\" RowOrdinal=\"0\" ReadOnly=\"TRUE\" Type=\"DateTime\" Name=\"Created\" DisplayName=\"Created\" StorageTZ=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Created\" FromBaseType=\"TRUE\" Hidden=\"TRUE\"/><Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" Type=\"Text\" Name=\"Title\" ShowInNewForm=\"FALSE\" ShowInFileDlg=\"FALSE\" DisplayName=\"Title\" Sealed=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Title\" ColName=\"nvarchar13\" Required=\"FALSE\" ShowInEditForm=\"TRUE\"/><Field ID=\"{28cf69c5-fa48-462a-b5cd-27b6f9d2bd5f}\" ColName=\"tp_Modified\" RowOrdinal=\"0\" ReadOnly=\"TRUE\" Type=\"DateTime\" Name=\"Modified\" DisplayName=\"Modified\" StorageTZ=\"TRUE\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Modified\" FromBaseType=\"TRUE\" Hidden=\"TRUE\"/><Field ID=\"{822c78e3-1ea9-4943-b449-57863ad33ca9}\" ReadOnly=\"TRUE\" Hidden=\"FALSE\" Type=\"Text\" Name=\"Modified_x0020_By\" DisplayName=\"Document Modified By\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Modified_x0020_By\" FromBaseType=\"TRUE\" ColName=\"nvarchar1\"/><Field ID=\"{4dd7e525-8d6b-4cb4-9d3e-44ee25f973eb}\" ReadOnly=\"TRUE\" Hidden=\"FALSE\" Type=\"Text\" Name=\"Created_x0020_By\" DisplayName=\"Document Created By\" SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" StaticName=\"Created_x0020_By\" FromBaseType=\"TRUE\" ColName=\"nvarchar2\"/><Field ID=\"{617f8947-74b2-36bc-9f7e-21ded7029bb5}\" Type=\"Note\" DisplayName=\"MediaServiceMetadata\" Name=\"MediaServiceMetadata\" Group=\"_Hidden\" Hidden=\"TRUE\" Indexed=\"FALSE\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" ShowInNewForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" ShowInEditForm=\"FALSE\" ShowInListSettings=\"FALSE\" Viewable=\"FALSE\" Json=\"TRUE\" SourceID=\"{f2735f63-b609-4f6e-b4a6-4b8838dde8cd}\" StaticName=\"MediaServiceMetadata\" ColName=\"ntext3\" RowOrdinal=\"0\"/><Field ID=\"{b887b6b2-4dcf-34fc-98b1-d5a42c605755}\" Type=\"Note\" DisplayName=\"MediaServiceFastMetadata\" Name=\"MediaServiceFastMetadata\" Group=\"_Hidden\" Hidden=\"TRUE\" Indexed=\"FALSE\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" ShowInNewForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" ShowInEditForm=\"FALSE\" ShowInListSettings=\"FALSE\" Viewable=\"FALSE\" Json=\"TRUE\" SourceID=\"{f2735f63-b609-4f6e-b4a6-4b8838dde8cd}\" StaticName=\"MediaServiceFastMetadata\" ColName=\"ntext4\" RowOrdinal=\"0\"/><Field ID=\"{dc412798-e1ac-4884-afa6-3cb4a10dcdda}\" Type=\"Note\" DisplayName=\"MediaServiceSearchProperties\" Name=\"MediaServiceSearchProperties\" Group=\"_Hidden\" Hidden=\"TRUE\" Indexed=\"FALSE\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" ShowInNewForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" ShowInEditForm=\"FALSE\" ShowInListSettings=\"FALSE\" Viewable=\"FALSE\" Json=\"TRUE\" SourceID=\"{f2735f63-b609-4f6e-b4a6-4b8838dde8cd}\" StaticName=\"MediaServiceSearchProperties\" ColName=\"ntext5\" RowOrdinal=\"0\"/><Field ID=\"{6e9ab27b-184d-411f-aedb-00386b5f89c2}\" Type=\"Text\" DisplayName=\"MediaServiceObjectDetectorVersions\" Name=\"MediaServiceObjectDetectorVersions\" Group=\"_Hidden\" Hidden=\"TRUE\" Indexed=\"TRUE\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" ShowInNewForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" ShowInEditForm=\"FALSE\" ShowInListSettings=\"FALSE\" Viewable=\"FALSE\" Json=\"FALSE\" SourceID=\"{f2735f63-b609-4f6e-b4a6-4b8838dde8cd}\" StaticName=\"MediaServiceObjectDetectorVersions\" ColName=\"nvarchar19\" RowOrdinal=\"0\"/><Field ID=\"{64ae69b5-5b9d-4ba9-b01c-f0ab72af8b7b}\" Type=\"Text\" DisplayName=\"MediaServiceGenerationTime\" Name=\"MediaServiceGenerationTime\" Group=\"_Hidden\" Hidden=\"TRUE\" Indexed=\"FALSE\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" ShowInNewForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" ShowInEditForm=\"FALSE\" ShowInListSettings=\"FALSE\" Viewable=\"FALSE\" Json=\"FALSE\" SourceID=\"{f2735f63-b609-4f6e-b4a6-4b8838dde8cd}\" StaticName=\"MediaServiceGenerationTime\" ColName=\"nvarchar20\" RowOrdinal=\"0\"/><Field ID=\"{ae5f3e55-9e2d-4632-9b90-282575c1c1f3}\" Type=\"Text\" DisplayName=\"MediaServiceEventHashCode\" Name=\"MediaServiceEventHashCode\" Group=\"_Hidden\" Hidden=\"TRUE\" Indexed=\"FALSE\" Sealed=\"TRUE\" ReadOnly=\"TRUE\" ShowInNewForm=\"FALSE\" ShowInDisplayForm=\"FALSE\" ShowInEditForm=\"FALSE\" ShowInListSettings=\"FALSE\" Viewable=\"FALSE\" Json=\"FALSE\" SourceID=\"{f2735f63-b609-4f6e-b4a6-4b8838dde8cd}\" StaticName=\"MediaServiceEventHashCode\" ColName=\"nvarchar21\" RowOrdinal=\"0\"/></Fields><XmlDocuments><XmlDocument NamespaceURI=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>DocumentLibraryForm</Display><Edit>DocumentLibraryForm</Edit><New>DocumentLibraryForm</New></FormTemplates></XmlDocument></XmlDocuments><Folder TargetName=\"Forms/Document\"/></ContentType>",
	Scope: string; // /sites/mhteam/Projects/Biochemistry Reference Calculator",
	Sealed: false,
	StringId: string; // 0x010100AEBD6EF2AB1818488F56CB69C8BF8C62"
};

type SPRootFolderRaw = {
	__metadata: {
		id: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder",
		type: string; // "SP.Folder"
	},
	Files: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder/Files"
		}
	},
	ListItemAllFields: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder/ListItemAllFields"
		}
	},
	ParentFolder: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder/ParentFolder"
		}
	},
	Properties: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder/Properties"
		}
	},
	StorageMetrics: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder/StorageMetrics"
		}
	},
	Folders: {
		__deferred: {
			uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/Web/Lists(guid'f2735f63-b609-4f6e-b4a6-4b8838dde8cd')/RootFolder/Folders"
		}
	},
	Exists: boolean; // true,
	ExistsAllowThrowForPolicyFailures: boolean; // true,
	ExistsWithException: boolean; // true,
	IsWOPIEnabled: boolean; // false,
	ItemCount: number // 3,
	Name: string; // "Biochemistry Reference Calculator",
	ProgID: null,
	ServerRelativeUrl: string; // "/sites/mhteam/Projects/Biochemistry Reference Calculator",
	TimeCreated: string; // "2024-01-23T15:55:35Z",
	TimeLastModified: string; // "2024-03-06T16:07:40Z",
	UniqueId: string; // "1e7158de-b2ba-4bee-a5ed-7a85b70e8cc1",
	WelcomePage: string; // ""
};

type SPWebInformationRaw = { // _api/web/webinfos
	__metadata: {
		id: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/SP.WebInformation5f1dd372-2f37-4263-a8a4-43ea7f3f492b",
		uri: string; // "https://mhalloran.sharepoint.com/sites/mhteam/Projects/_api/SP.WebInformation5f1dd372-2f37-4263-a8a4-43ea7f3f492b",
		type: string; // "SP.WebInformation"
	},
	Configuration: number; // 0,
	Created: string; // "2024-03-16T22:36:06Z",
	Description: string; // "",
	Id: string; // "8ce7e9f8-0674-4d0a-b182-34cefd6f4659",
	Language: number; // 1033,
	LastItemModifiedDate: string; // "2024-03-16T22:36:11Z",
	LastItemUserModifiedDate: string; // "2024-03-16T22:36:11Z",
	ServerRelativeUrl: string; // "/sites/mhteam/Projects/TestProject",
	Title: string; // "Test Project Site",
	WebTemplate: string; // "PROJECTSITE",
	WebTemplateId: number; // 0
};