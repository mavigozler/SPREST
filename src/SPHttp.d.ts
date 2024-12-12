
export {
	TSPResponseData,
	TSPResponseDataProperties,
	THttpRequestBody,
	THttpRequestHeaders,
	THttpRequestParams,
	THttpResponseHeaders,
	THttpRequestProtocol,
	THttpRequestParamsWithPromise,
	HttpInfo,
	FetchApiResponseInfo,
	TXmlHttpRequestData,
	TIntervalControl,
	TFetchInfo,
	IBatchHTTPRequestForm,
	TBatchResponseRaw,
	TBatchResponse,
	THttpRequestMethods,
	HttpStatusCode,
	TSuccessCallback,
	TErrorCallback,
	TArrayToJSON,
	TOriginType
};

type TJQueryPlainObject = {[key:string]: [] | typeof document | string };
type TXmlHttpRequestData =  ( [] | string | TJQueryPlainObject | Uint8Array ) & BodyInit ;

// export
type TSPResponseData = {
	d?: {
		results?: unknown[],
		GetContextWebInformation?: {
			FormDigestValue: string;
		}
		__next?: string;
		[key:string]: unknown;
	};
	status?: HttpStatusCode;  // HTTP response code if avaiable
	FormDigestValue?: string;
	__next?: string;
	"odata.nextLink"?: string; // a variant of __next
	value?: object; // where the response is an object of indetermine properties
};// & {[key: string]: string | boolean | number};

// export
type TSPResponseDataProperties = {
	Id: string;
	OtherId: string;
	Title: string;
	InternalName?: string;
	ServerRelativeUrl: string;
	WebTemplate?: string;
};

type HttpInfo = THttpRequestParams & FetchApiResponseInfo & THttpResponseHeaders;
type TSuccessCallback = (
	data: TSPResponseData,
	httpInfo: HttpInfo  /* fetch now used */
) => void;
type TErrorCallback = (
	errInfo: unknown,
	httpInfo: HttpInfo
) => void;

type THttpRequestProtocol = "https" | "http" | "https://" | "http://";

type THttpRequestMethods = "GET" | "POST" | "PUT" | "PATCH" | "HEAD" | "OPTIONS" |
		"DELETE" | "TRACE" | "CONNECT";

interface THttpRequestHeaders {
	"Accept"?: "application/json;odata=verbose" | "application/json;odata=nometadata";
	"Content-Type"?: "application/json;odata=verbose" | "text/plain" | "text/html" | "text/xml";
	"IF-MATCH"?: "*"; // can also use etag
	"X-HTTP-METHOD"?: "MERGE" | "DELETE";
	"X-RequestDigest"?: string;
}

interface THttpResponseHeaders {
	"content-type"?: string; // "application/json"
	"date"?: string;  // "Sun, 28 Apr 2024 03:28:47 GMT"
	"etag"?: string;  // "\"{C9EB2CF2-3756-450C-8B6A-0603C80EF420},3\""
	"expires"?: string; // "Sat, 13 Apr 2024 03:28:48 GMT"
	"last-modified"?: string; // "Sat, 27 Apr 2024 22:28:04 GMT"
}

type TOriginType = "basic" | // default not cross-origin, or CORS with correct headers to allow sharing
	"cors" | // response x-origin with correct headers
	"opaque" |  // was cross origin but did not include headers, not accessible to requesting script
	"opaqueredirect";  // initial request redirected to cross-origin resource w/o CORS headers

interface FetchApiResponseInfo {
	bodyUsed: boolean;
	ok: boolean;
	redirected: boolean;
	status: number;
	statusText: string;
	type: TOriginType;
}

// this type used to monitor HTTP request calls in an intermediate fashion
type TIntervalControl = {
	currentCount?: number;   // keeps track of current counts of results returned in (AJAX) request
	nextCount?: number;      // keeps track of next count
	interval: number;       // the number of results returned for each interval
	callback: (count: number) => void;  // function to call when a certain count of results are returned in request
};

type THttpRequestParams = {
	url: string;
	setDigest?: boolean;
	method?: THttpRequestMethods;
	headers?: THttpRequestHeaders;
	customHeaders?: THttpRequestHeaders;
	data?: TXmlHttpRequestData;
	body?:  TXmlHttpRequestData;
	useCORS?: boolean;
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


type HttpStatusCode = 0
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

type THttpRequestBody<T = unknown> = {
	[key: string]: T;
};

// export
type TArrayToJSON = Record<string, string | boolean | object | number>

type TQueryProperties = "Expand" | "Filter" | "Select" | "expand" | "filter" | "select";

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

