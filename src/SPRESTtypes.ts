
const
SPListTemplateTypes = {
	enums: [
		{ name: "InvalidType",     typeId: -1 },
		{ name: "NoListTemplate",  typeId: 0 },
		{ name: "GenericList",     typeId: 100 },
		{ name: "DocumentLibrary", typeId: 101 },
		{ name: "Survey",          typeId: 102 },
		{ name: "Links",           typeId: 103 },
		{ name: "Announcements",   typeId: 104 },
		{ name: "Contacts",        typeId: 105 },
		{ name: "Events",          typeId: 106 },
		{ name: "Tasks",           typeId: 107 },
		{ name: "DiscussionBoard", typeId: 108 },
		{ name: "PictureLibrary",  typeId: 109 },
		{ name: "DataSources",     typeId: 110 },
		{ name: "WebTemplateCatalog", typeId: 111 },
		{ name: "UserInformation", typeId: 112 },
		{ name: "WebPartCatalog",  typeId: 113 },
		{ name: "ListTemplateCatalog", typeId: 114 },
		{ name: "XMLForm",         typeId: 115 },
		{ name: "MasterPageCatalog", typeId: 116 },
		{ name: "NoCodeWorkflows", typeId: 117 },
		{ name: "WorkflowProcess", typeId: 118 },
		{ name: "WebPageLibrary",  typeId: 119 },
		{ name: "CustomGrid",      typeId: 120 },
		{ name: "SolutionCatalog", typeId: 121 },
		{ name: "NoCodePublic", typeId: 122 },
		{ name: "ThemeCatalog", typeId: 123 }
	],
	getFieldTypeNameFromTypeId: (typeId: number) => {
		return (SPListTemplateTypes.enums.find(elem => elem.typeId == typeId))!.name;
	},
	getFieldTypeIdFromTypeName: (typeName: string) => {
		return (SPListTemplateTypes.enums.find(elem => elem.name == typeName))!.typeId;
	}
};

type TSPFieldDataTypeName = "Invalid" | "Integer" | "Text" | "Note" | "DateTime" | "Counter" |
	"Choice" | "Lookup" | "Boolean" | "Number" | "Currency" | "URL" | "Computed" | "Threading" |
	"Guid" | "MultiChoice" | "GridChoice" | "Calculated" | "File" | "Attachments" | "User" |
	"Recurrence" | "CrossProjectLink" | "ModStat" | "Error" | "ContentTypeId" | "PageSeparator" |
	"ThreadIndex" | "WorkflowStatus" | "AllDayEvent" | "WorkflowEventType" | "Geolocation" |
	"OutcomeChoice" | "MaxItems";

interface ISPFieldDataTypeInfo {
	name: TSPFieldDataTypeName;
	metadataType: string | null;
	typeId: number;
	extraProperties?: string[];
}

const SPFieldTypes: ISPFieldDataTypeInfo[] = [
		{ name: "Invalid",      metadataType: null,                typeId: 0 },
		{ name: "Integer",      metadataType: "SP.FieldNumber",     typeId: 1,
				extraProperties: [
					"CommaSeparator",   		// boolean
					"CustomUnitName",       // string or null
					"CustomUnitOnRight",    // boolean
					"DisplayFormat",  		// integer # of decimal digits to display
					"MaximumValue",         // double type float: max value of num
					"MinimumValue",			// double type float: min value of num
					"ShowAsPercentage",		// boolean - render value as percentage
					"Unit"						// string
				]
		},
		{ name: "Text",         metadataType: "SP.FieldText",           typeId: 2,
				extraProperties: [
					"MaxLength"   // integer: gets maximum number of characters allowed for field
				]
		},
		{ name: "Note",         metadataType: "SP.FieldMultiLineText",  typeId: 3,
				extraProperties: [
					"AllowHyperlink",   // boolean: get/set whether hyperlink allowed in field value
					"AppendOnly",       // boolean: get/set whether changes to value are displayed in list forms
					"IsLongHyperlink",  // boolean:
					"NumberOfLines",    // integer: get/set # lines to display for field
					"RestrictionMode",  // boolean: get/set whether to support subset of rich formatting
					"RichText",         // boolean: get/set whether to support rich formatting
					"UnlimitedLengthInDocumentLibrary",   // boolean: get/set
					"WikiLinking",      // boolean:  get value whether implementation-specific mechanism for linking wiki pages
				]
		},
		{ name: "DateTime",     metadataType: "SP.FieldDateTime",  typeId: 4,
				extraProperties: [
					"DateTimeCalendarType",   //
					"DateFormat",             //
					"DisplayFormat",          //
					"FriendlyDisplayFormat",  //
					"TimeFormat"
				]
		},
		{ name: "Counter",      metadataType: "SP.Field",        typeId: 5 },
		{ name: "Choice",       metadataType: "SP.FieldChoice",  typeId: 6,
				extraProperties: [
					"FillInChoice",     // boolean: gets/sets whether fill-in value allowed
					"Mappings",  // string:
					"Choices"       //  objects with "results" property that is array of strings with choices
				]
		},
		{ name: "Lookup",       metadataType: "SP.FieldLookup",   typeId: 7,
				extraProperties: [
					"AllowMultipleValues",   		// boolean: whether lookup field allows multiple values
					"DependentLookupInternalNames",  // string or null
					"IsDependentLookup",    // boolean: gets indication if field is 2ndary lookup field that depends on primary field
							// for its relationship with lookup list
					"IsRelationship",  		//boolean: gets/set whether lookup field returned by GetRelatedField from list being looked up
					"LookupField",         // string: gets/set value of internal field name of field used as lookup value
					"LookupList",			// string: gets/sets value of list identifier of list containing field to lookup values
					"LookupWebId",		// SP.Guid: gets/sets value of GUID identifying site containing list which has field for lookup values
					"PrimaryFieldId",		// string: gets/sets value specifying primary look field identifier if there is a dependent
								// lookup field; otherwise empty string
					"RelationshipDeleteBehavior",		// SP.RelationshipDeleteBehaviorType: gets/sets value that specifies delete behavior of lookup field
					"UnlimitedLengthInDocumentLibrary",	 // boolean:
				]
		},
		{ name: "Boolean",      metadataType: "SP.Field",        typeId: 8 },
		{ name: "Number",       metadataType: "SP.FieldNumber",  typeId: 9,
			extraProperties: [
				"CommaSeparator",   		// boolean
				"CustomUnitName",       // string or null
				"CustomUnitOnRight",    // boolean
				"DisplayFormat",  		// integer # of decimal digits to display
				"MaximumValue",         // double type float: max value of num
				"MinimumValue",			// double type float: min value of num
				"ShowAsPercentage",		// boolean - render value as percentage
				"Unit"						// string
			]
		},
		{ name: "Currency",     metadataType: "SP.FieldCurrency",  typeId: 10 },
		{ name: "URL",          metadataType: "SP.FieldUrl",       typeId: 11,
				extraProperties: [
					"DisplayFormat"   // integer: unknown about value sets
				]
		},
		{ name: "Computed",     metadataType: "SP.FieldComputed",    typeId: 12,
				extraProperties: [
					"EnableLookup"  // boolean: gets/sets value specifying whether lookup field can reference the field
				]
		},
		{ name: "Threading",    metadataType: "SP.FieldThreading",   typeId: 13 },
		{ name: "Guid",         metadataType: "SP.FieldGuid",        typeId: 14 }, // so far, just a standard basic field as string
		{ name: "MultiChoice",  metadataType: "SP.FieldMultiChoice", typeId: 15,
				extraProperties: [
					"FillInChoice",     // boolean: gets/sets whether fill-in value allowed
					"Mappings",  // string:
					"Choices"       //  objects with "results" property that is array of strings with choices
				]
		},
		{ name: "GridChoice",   metadataType: "SP.FieldGridChoice",typeId: 16 },
		{ name: "Calculated",   metadataType: "SP.FieldCalculated",typeId: 17 },
		{ name: "File",         metadataType: "SP.FieldFile",      typeId: 18 },
		{ name: "Attachments",  metadataType: "SP.FieldAttachments", typeId: 19 },
		{ name: "User",         metadataType: "SP.FieldUser",       typeId: 20,
				extraProperties: [
					// these properties are from Lookup type 7
					"AllowMultipleValues",   		// boolean: whether lookup field allows multiple values
					"DependentLookupInternalNames",  // string or null
					"IsDependentLookup",    // boolean: gets indication if field is 2ndary lookup field that depends on primary field
							// for its relationship with lookup list
					"IsRelationship",  		//boolean: gets/set whether lookup field returned by GetRelatedField from list being looked up
					"LookupField",         // string: gets/set value of internal field name of field used as lookup value
					"LookupList",			// string: gets/sets value of list identifier of list containing field to lookup values
					"LookupWebId",		// SP.Guid: gets/sets value of GUID identifying site containing list which has field for lookup values
					"PrimaryFieldId",		// string: gets/sets value specifying primary look field identifier if there is a dependent
								// lookup field; otherwise empty string
					"RelationshipDeleteBehavior",		// SP.RelationshipDeleteBehaviorType: gets/sets value that specifies delete behavior of lookup field
					"UnlimitedLengthInDocumentLibrary",	 // boolean:
					// these are special to User
					"AllowDisplay",  		//boolean: gets/set whether to display name of user in survey list
					"Presence",         // boolean: gets/sets whether presence (online status) is enabled on the field
					"SelectionGroup",			// number: gets/sets value of identifiter of SP group whose members can be selected as values of field
					"SelectionMode",		// SP.FieldUserSelectionMode: gets/sets value specifying whether users and groups or only users can be selected
							// SP.FieldUserSelectionMode.peopleAndGroups = 1, SP.FieldUserSelectionMode.peopleOnly = 0
					"UserDisplayOptions"		//
				]
		},
		{ name: "Recurrence",   metadataType: "SP.FieldRecurrence", typeId: 21 },
		{ name: "CrossProjectLink", metadataType: "SP.FieldCrossProjectLink", typeId: 22 },
		{ name: "ModStat",      metadataType: "SP.ModStat", typeId: 23, // Moderation Status is a Choices type
			extraProperties: [
				"FillInChoice",     // boolean: gets/sets whether fill-in value allowed
				"Mappings",  // string:
				"Choices",       //  objects with "results" property that is array of strings with choices
						// usually "results": ["0;#Approved", "1;#Rejected", "2;#Pending", "3;#Draft", "4;#Scheduled"]
				"EditFormat"       // integer value for format type in editing
			]
		},
		{ name: "Error",        metadataType: "SP.FieldError", typeId: 24 },
		{ name: "ContentTypeId", metadataType: "SP.FieldContentTypeId",   typeId: 25 },
		{ name: "PageSeparator", metadataType: "SP.FieldPageSeparator", typeId: 26 },
		{ name: "ThreadIndex",  metadataType: "SP.FieldThreadIndex", typeId: 27 },
		{ name: "WorkflowStatus", metadataType: "SP.FieldWorkflowStatus", typeId: 28 },
		{ name: "AllDayEvent",  metadataType: "SP.FieldAllDayEvent", typeId: 29 },
		{ name: "WorkflowEventType", metadataType: "SP.FieldWorkflowEventType", typeId: 30 },
		{ name: "Geolocation",  metadataType: "SP.FieldGeolocation", typeId: -1 },
		{ name: "OutcomeChoice", metadataType: "SP.FieldOutcomeChoice", typeId: -1},
		{ name: "MaxItems",    metadataType: "SP.FieldMaxItems", typeId: 31 }
	];

const
	standardProperties: string[] = [
		"AutoIndexed",
		"CanBeDeleted",  // true = field can be deleted
					// returns false if FromBaseType == true || Sealed == True
		"ClientSideComponentId", //
		"ClientSideComponentProperties",
		"ClientValidationFormula",
		"ClientValidationMessage",
		"CustomFormatter",
		"DefaultFormula", // string: gets/sets default formula for calculated field
		"DefaultValue", // string: gets/sets default value for field
		"Description",
		"Direction",  // string: "none","LTR","RTL" gets/sets reading order of field
		"EnforceUniqueValues", // boolean: gets/sets to force uniqueness of column values (false=default)
		"EntityPropertyName", // string: gets name of entity proper for list item entity using field
		"Filterable", // boolean: gets whether field can be filtered
		"FromBaseType", // boolean: gets whetehr field derives from base field type
		"Group", // string: gets/sets column group to which field belongs
			// Groups: Base, Core Contact and Calendar, Core Document, Core Task and Issue, Custom, Extended, _Hidden, Picture
		"Hidden", // boolean: specifies field display in list
		"Id", // guid of field
		"Indexed", // boolean: gets/sets if field is indexed
		"IndexStatus",
		"InternalName", // string: gets internal name of field
		"IsModern",
		"JSLink",  // string: gets/sets name of JS file(s) [separated by '|'] for client rendering logic of field
		"PinnedToFiltersPane",
		"ReadOnlyField", // boolean: gets/sets whether field can be modifed
		"Required", // boolean: gets/sets whether user must enter value on New and Edit forms
		"SchemaXml", // string: gets/sets schema that defines the field
		"Scope", // string: gets Web site-relative path to list in which field collection is used
		"Sealed", // boolean: gets/sets whether field type can be parent of custom derived field type
		"ShowInFiltersPane",
		"Sortable", // boolean: gets boolean whether field can be used in sort
		"StaticName", // string: gets/sets static name
		"Title", // string: gets/sets display name for field
		"FieldTypeKind", //
		"TypeAsString", // string: gets the type of field as a string value
		"TypeDisplayName", // string: gets display name of the field type
		"TypeShortDescription", // string: gets the description of the field
		"ValidationFormula", // string: gets/sets formula referenced by field, evaluated when list item added/updated
		"ValidationMessage" // string: gets/sets message to display if validation fails
	];
const getFieldTypeNameFromTypeId = (typeId: number) => {
		return (SPFieldTypes.find(elem => elem.typeId == typeId))!.name;
	},
	getFieldTypeIdFromTypeName = (typeName: string) => {
		return (SPFieldTypes.find(elem => elem.name == typeName))!.typeId;
	}


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

interface ISPListLibFieldInfo {
	propertyName: string;
	fieldDisplayName: string;
	fieldInternalName: string;
	fieldDataType: TSPFieldDataTypeName;
	baseType: boolean;  // deals with group properties in the base types
}

// defined by JQuery AJAX
type TJQueryPlainObject = {[key:string]: [] | typeof document | string };
type TXmlHttpRequestData =  [] | string | TJQueryPlainObject;

type TSPResponseDataProperties = {
	Id: number | string;  // could be a uuid/guid
	Title: string;
	ParentList?: {
		ListItemEntityTypeFullName: string;
		RootFolder: {
			ServerRelativeUrl: string;
		};
	};
// properties below specific for 'items' requests
	ServerRelativeUrl: string;
	WebTemplate?: string;
	FileSystemObjectType?: 0 | 1;
	File?: {
		Name: string;
		CheckOutType: 0 | 1 | 2;
		Length: string; // really a number
		MajorVersion: number;
		MinorVersion: number;
		Title: null | string;  // doesnt seem to be really used
		ETag: string;
		TimeCreated: string; // date TZ format
		TimeLastModified: string; // date TZ format
	};
// properties below would be special to requests for list fields/columns
	InternalName?: string;
	TypeAsString?: string;
	FieldTypeKind?: number;
	Required?: boolean;
	Indexed?: boolean; // has the column been set to be indexed?
	FromBaseType?: boolean;
	ReadOnlyField?: boolean;
	Group?: string; // like "Custom Columns"
	Filterable?: boolean; // determined by system as dynamically filterable
	Sortable?: boolean;
};

type TSPResponseData = {
	d?: {
		results?: TSPResponseDataProperties[],
		GetContextWebInformation?: {
			FormDigestValue: string;
		}
		__next?: string;
		[key:string]: any;
	};
	results?:  TSPResponseDataProperties[];
	FormDigestValue?: string;
	__next?: string;
};

type TSuccessCallback = (data: TSPResponseData, status?: string, reqObj?: JQueryXHR) => void;
type TErrorCallback = (reqObj: JQueryXHR, status?: string, err?: string) => void;

type THttpRequestProtocol = "https" | "http" | "https://" | "http://";

type THttpRequestMethods = "GET" | "POST" | "PUT" | "PATCH" | "HEAD" | "OPTIONS" |
		"DELETE" | "TRACE" | "CONNECT";

type THttpRequestHeaders = {
	"Accept"?: "application/json;odata=verbose" | "application/json;odata=nometadata";
	"Content-Type"?: "application/json;odata=verbose" | "text/plain" | "text/html" | "text/xml";
	"IF-MATCH"?: "*"; // can also use etag
	"X-HTTP-METHOD"?: "MERGE" | "DELETE";
	"X-RequestDigest"?: string;
	[key:string]: any;
};

type THttpRequestParams = ({
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?:  TXmlHttpRequestData;
	successCallback: TSuccessCallback;
	errorCallback: TErrorCallback;
});

type THttpRequestParamsWithPromise = ({
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?: TXmlHttpRequestData;
	successCallback?: TSuccessCallback;
	errorCallback?: TErrorCallback;
});

type THttpRequestBody<T = any> = {
		[key: string]: T;
	}

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
	method?: THttpRequestMethods;
	body?: THttpRequestBody;
	headers?: THttpRequestHeaders;
}

type TParsedURL = {
	originalUrl: string;
	protocol: THttpRequestProtocol;
	server: string;
	hostname: string;
	siteFullPath: string;
	sitePartialPath: string;
	list: string;
	listConfirmed: boolean;
	libRelativeUrl: string;
	file: string;
	query: any[];
};

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
	fieldName: string,
	choices: {
		id: string;
		value: string;
	}[] | null;
};

type TSpSiteInfo = {
	name:string;
	serverRelativeUrl:string;
	id:string;
	template:string
};

type TSiteInfo = {
	name?: string | null;
	server?: string;
	serverRelativeUrl?: string;
	id?: string;
	template?: string;
	children?: TSiteInfo[];
	parent?: TSiteInfo | null;
	referenceSite?: TSiteInfo;
} ;

type TSPDocLibCheckInType =  "minor" | "Minor" | "major" | "Major" | "overwrite" | "Overwrite";

type TUserSearch = {
	userId?: number;
	id?: number;
	firstName?: string;
	lastName?: string;
};

interface IEmailHeaders {
	To: string;
	Subject: string;
	Body: string;
	CC?: string | null;
	BCC?: string | null;
	From?: string | null;
}

// declare function RequestAgain (next: string, data: any): Promise<any>;
