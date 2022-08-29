
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
type TXmlHttpRequestData =  [] | string | TJQueryPlainObject;

type TSPResponseDataProperties = {
	Id: string;
	Title: string;
	InternalName?: string;
	ServerRelativeUrl: string;
	WebTemplate?: string;
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
	[key: string]: any;
};

// this type used to monitor HTTP request calls in an intermediate fashion
type TIntervalControl = {
	currentCount?: number;   // keeps track of current counts of results returned in (AJAX) request
	nextCount?: number;      // keeps track of next count
	interval: number;       // the number of results returned for each interval
	callback: (count: number) => void;  // function to call when a certain count of results are returned in request
} | null;

type THttpRequestParams = ({
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?:  TXmlHttpRequestData;
	successCallback: TSuccessCallback;
	errorCallback: TErrorCallback;
	ignore?: number[]; // errors to ignore/catch
	progressReport?: TIntervalControl
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
	ignore?: number[]; // errors to ignore/catch
	progressReport?: TIntervalControl
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





// declare function RequestAgain (next: string, data: any): Promise<any>;
