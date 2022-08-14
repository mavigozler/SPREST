
export interface ISPBaseListFieldProperties {
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

export interface ISPBaseListItemProperties {
	__metadata: { type: string; }
	Id: number;
	InternalName: string;
	File: { Name: string; }
	Title: string;
}


// defined by JQuery AJAX
export type TJQueryPlainObject = {[key:string]: [] | typeof document | string };
export type TXmlHttpRequestData =  [] | string | TJQueryPlainObject;

export type TSPResponseDataProperties = {
	Id: string;
	Title: string;
	InternalName?: string;
	ServerRelativeUrl: string;
	WebTemplate?: string;
};

export type TSPResponseData = {
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

export type TSuccessCallback = (data: TSPResponseData, status?: string, reqObj?: JQueryXHR) => void;
export type TErrorCallback = (reqObj: JQueryXHR, status?: string, err?: string) => void;

export type THttpRequestProtocol = "https" | "http" | "https://" | "http://";

export type THttpRequestMethods = "GET" | "POST" | "PUT" | "PATCH" | "HEAD" | "OPTIONS" |
		"DELETE" | "TRACE" | "CONNECT";

export type THttpRequestHeaders = {
	"Accept"?: "application/json;odata=verbose" | "application/json;odata=nometadata";
	"Content-Type"?: "application/json;odata=verbose" | "text/plain" | "text/html" | "text/xml";
	"IF-MATCH"?: "*"; // can also use etag
	"X-HTTP-METHOD"?: "MERGE" | "DELETE";
	"X-RequestDigest"?: string;
	[key: string]: any;
};

export type THttpRequestParams = ({
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?:  TXmlHttpRequestData;
	successCallback: TSuccessCallback;
	errorCallback: TErrorCallback;
});

export type THttpRequestParamsWithPromise = ({
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?: TXmlHttpRequestData;
	successCallback?: TSuccessCallback;
	errorCallback?: TErrorCallback;
});

export type THttpRequestBody<T = any> = {
		[key: string]: T;
	}

export type TQueryProperties = "Expand" | "Filter" | "Select" | "expand" | "filter" | "select";

export interface IBatchHTTPRequestParams {
	host: string;
	path: string;
	protocol?: THttpRequestProtocol;
	AllHeaders?: THttpRequestHeaders;
	AllMethods?: THttpRequestMethods;
}

export interface IBatchHTTPRequestForm {
	url: string;
	method?: THttpRequestMethods;
	body?: THttpRequestBody;
	headers?: THttpRequestHeaders;
}

export type TParsedURL = {
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

export interface IListSetupMinimal {
	server: string;
	site: string;
	include?: string;
	protocol?: string;
}

export interface IListSetupWithName extends IListSetupMinimal {
	listName: string;
	listGuid?: never;
}

export interface IListSetupWithGuid extends IListSetupMinimal {
	listGuid: string;
	listName?: never;
}

export type TListSetup = IListSetupWithName | IListSetupWithGuid;

export type TLookupFieldInfo = {
	fieldDisplayName: string,
	fieldInternalName: string;
	choices: {
		id: number;
		value: string;
	}[] | null;
};

export type TSpSiteInfo = {
	name:string;
	serverRelativeUrl:string;
	id:string;
	template:string
};

export type TSiteInfo = {
	name?: string | null;
	server?: string;
	serverRelativeUrl?: string;
	id?: string;
	template?: string;
	children?: TSiteInfo[];
	parent?: TSiteInfo | null;
	referenceSite?: TSiteInfo;
} ;

export type TSPDocLibCheckInType =  "minor" | "Minor" | "major" | "Major" | "overwrite" | "Overwrite";





// declare function RequestAgain (next: string, data: any): Promise<any>;
