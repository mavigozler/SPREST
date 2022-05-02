declare const SPstdHeaders: THttpRequestHeaders;
declare let SelectAllCheckboxes: string, // defined below
UnselectAllCheckboxes: string;
declare class SPServerREST {
    URL: string;
    apiPrefix: string;
    static stdHeaders: THttpRequestHeaders;
    constructor(setup: {
        URL: string;
    });
    static httpRequest(elements: THttpRequestParams): void;
    static httpRequestPromise(parameters: {
        url: string;
        setDigest?: boolean;
        method?: THttpRequestMethods;
        headers?: THttpRequestHeaders;
        data?: TXmlHttpRequestData;
        body?: TXmlHttpRequestData;
    }): Promise<any>;
    getSiteProperties(): Promise<any>;
    getWebProperties(): Promise<any>;
    getEndpoint(endpoint: string): Promise<any>;
    getRootweb(): Promise<any>;
}
declare const MAX_REQUESTS = 500;
/**
 * @function batchRequestingQueue -- when requests are too large in number (> MAX_REQUESTS), the batching
 * 		needs to be broken up in batches
 * @param {IBatchHTTPRequestParams} elements -- same as elements in singleBatchRequest
 *    the BatchHTTPRequestParams object has following properties
 *       host: string -- required name of the server (optional to lead with "https?://")
 *       path: string -- required path to a valid SP site
 *       protocol?: string -- valid use of "http" or "https" with "://" added to it or not
 *       AllHeaders?: THttpRequestHeaders -- object of [key:string]: T; type, headers to apply to all requests in batch
 *       AllMethods?: string -- HTTP method to apply to all requests in batch
 * @param {IBatchHTTPRequestForm} allRequests -- the requests in singleBatchRequest
 *    the BatchHTTPRequestForm object has following properties
 *       url: string -- the valid REST URL to a SP resource
 *       method?: httpRequestMethods -- valid HTTP protocol verb in the request
 */
declare function batchRequestingQueue(elements: IBatchHTTPRequestParams, allRequests: IBatchHTTPRequestForm[]): Promise<any>;
declare function CreateUUID(): string;
/**
 * @function singleBatchRequest -- used to exploit the $batch OData multiple write request
 * @param {object} elements -- properties that must be defined are:
 *        .protocol {string | optional} usually "https://"
 *        .host {string}  usually fully qualified domain name "cawater.sharepoint.com"
 *        .path {string}  follows the host name to the SP site -- must be valid SP site path
 *        .AllHeaders {object | optional}   headers in object format
 *        .AllMethods {string | optional}   if all requests are to be GET, POST, PATCH, etc in the batch
 * @param {arrayOfObject} requests -- each element of the array must be a 'request' object
 *       .protocol -- should be "https" or "http". if not specified, a default "https" will be assumed
 *       .url -- must specify URL for REST request
 *       .method -- the method or if not present, element.AllMethods
 *       .headers -- headers for request. Or if all headers same
 *       .body -- this should be the body for POST request in object form (not JSON or string)
 * @error -- The Promise reject() might return
 *      1. just the AJAX request object
 *      2. the object {reqObj: AJAX request object, addlMessage: <string> message }
 */
declare function singleBatchRequest(elements: IBatchHTTPRequestParams, requests: IBatchHTTPRequestForm[]): Promise<unknown>;
declare function splitRequests(responseSet: string): Promise<any>;
/**
 *
 * @param siteURL -- string representing URL of site
 * @param listNameOrGUID -- name or GUID of list
 * @param sourceColumnIntName -- internal field name to be copied
 * @param destColumnIntName -- internal field name where copies are created
 */
declare function SPListColumnCopy(siteURL: string, listNameOrGUID: string, sourceColumnIntName: string, destColumnIntName: string): Promise<any>;
/**
 * @function formatRESTBody -- creates a valid JSON object as string for specifying SP list item updates/creations
 * @param {object}  JsonBody -- JS object which conforms to JSON rules. Must be "field_properties" : "field_values" format
 *                    If one of the properties is "__SetType__", it will fix the "__metadata" property
 */
declare function formatRESTBody(JsonBody: {
    [key: string]: string | object;
} | string): string;
declare function checkEntityTypeProperty(body: object, typeCheck: string): boolean;
/**
 * @function ParseSPUrl -- analyzes URL specific to SharePoint and determines  its propertiess
 * @param {string} url -- url to be parsed
 * @returns {object} has following propertiies
 * 	originalUrl: url passed as argument
        protocol: returns likely "https://" or "http://"
        server: this will be the protocol + 'www.servername.com' parts
        hostname: the 'www.servername.com' part (between protcol and first slash)
        siteFullPath: starts at first slash and goes to last name with only '/' characters in between
        sitePartialPath: like siteFullPath, but does not have the last name after last slash
        list: if identified clearly, offers a list name
        listConfirmed: boolean to indicate whether list name confirmed
        libRelativeUrl: if identified, offers a lib relative URL
        file: identifiable part of  'rootName.extension'
        query: An array of name=value pairs as objects {key:, value:}
                parsed after the query identifier character '?'
 */
declare function ParseSPUrl(url: string): TParsedURL | null;
/**
 *
 * @param {function} opFunction
 * @param {arrayOfObject} dataset -- this must be an array of objects in JSON format that
 * 				represent the body of a POST request to create SP data: fields, items, etc.
 * @returns -- the return operations are to unwind the recursion in the processing
 *           to get to either resolve or reject operations
 */
declare function serialSPProcessing(opFunction: (arg1: any) => Promise<any>, itemset: any[]): Promise<any>;
declare function constructQueryParameters(parameters: {
    [key: string]: any | any[];
}): string;
/**
 * @function getCheckedInput -- returns the value of a named HTML input object representing radio choices
 * @param {HtmlInputDomObject} inputObj -- the object representing selectabe
 *     input: radio, checkbox
 * @returns {primitive data type | array | null} -- usually numeric or string representing choice from radio input object
 */
declare function getCheckedInput(inputObj: HTMLInputElement | RadioNodeList): null | string | string[];
/**
* @function setCheckedInput -- will set a radio object programmatically
* @param {HtmlRadioInputDomNode} inputObj   the INPUT DOM node that represents the set of radio values
* @param {primitive value} value -- can be numeric, string or null. Using 'null' effectively unsets/clears any
*        radio selection
& @returns boolean  true if value set/utilized, false otherwise
*/
declare function setCheckedInput(inputObj: HTMLInputElement & RadioNodeList, value: string | string[] | null): boolean;
/**
 * @function formatDateToMMDDYYYY -- returns ISO 8061 formatted date string to MM[d]DD[d]YYYY date
 *      string, where [d] is character delimiter
 * @param {string|datetimeObj} dateInput [required] -- ISO 8061-formatted date string
 * @param {string} delimiter [optional] -- character that will delimit the result string; default is '/'
 * @returns {string} -- MM[d]DD[d]YYYY-formatted string.
 */
declare function formatDateToMMDDYYYY(dateInput: string, delimiter?: string): string | null;
declare function fixValueAsDate(date: Date): Date | null;
declare function CSVToArray(strData: string, strDelimiter?: string): string[][];
/**
 * @note -- this function has general use and should be moved to a library for HTML controls
 * @function buildSelectSet -- creates a DIV containing two SELECT and two BUTTON objects. One SELECT is "available" options,
 * 	and other SELECT is "selected" options. Options are set in the 'options' parameter which is an array of object with
 * 	3 properties: 'name' for displayed name, 'value' for set value of option, and 'selected' = false (default) if option
 *    is to be put in 'available' list or = true if option is to be created in 'selected' list
 * @param {DomNodeObject} parentNode the existing DOM node that will contain the control
 * @param {string} nameOrGuid
 * @param {arrayOfObject} options element objects have form
 * 			{text:, value:, selected: T/F, required: T/F} to be used as option Node for select
 * @param {integer} availableSize  size value for Available select object
 * @param {integer} selectedSize  size value for Selected select object
 * @param {function} onChangeHandler   event handler when onchage occurs
 * @returns {string} DOM id attribute of containing DIV so that it can be removed as DOM node by caller
 */
declare function buildSelectSet(parentNode: HTMLElement, // DOM node to append the construct
nameOrGuid: string, // used in tagging
options: {
    selected?: boolean;
    required?: boolean;
    value: any;
    text: string;
}[], // array of option objects {text: text-to-display, value: node value to set}
availableSize: number, // for number of choices displayed for available choices select list
selectedSize: number, // for number of choices displayed for selected choices select list
onChangeHandler: (arg?: any) => {} | void): string | void;
/**
 * @function createSelectUnselectAllCheckboxes --
 * 		this will create two buttons for selecting or unselecting type checkmark
 *        input controls
 * @param {object} parameters -- object represents multiple arguments to function
 *   required properties are:
 *     form {required: object|string} must the 'id' attribute value of a valid form or its DOM node
 * 			that contains the existing input group of checkmarks
 *     formName {required: string} -- this must be the name attribute for all checkboxes that are to be
 * 			controlled by the buttons
 *     clickCallback {required: () => {}} -- continues click event back to calling program
 *   optional properties are:
 *     containerElemType {optional: string} -- usually a "span","p",or"div"
 * 	 includeText {optional: boolean} -- whether to print "select all" and "unselect all" next to buttons
 *     containerNode {object|string} either the DOM node to contain or the 'id' of the container of the buttons
 *     style string or array of string -- valid CSS rule "selector { property:propertyValue;...}
 *          styles can be specified for container of buttons or buttons
 *     images {object} must be JSON form {"select-url": "url to select all image",
 *                  "unselect-url": url to unselect image,
 *                  "select-image": base64-encoded image data, "unselect-image": b64 image data}
 * @returns {HtmlDomNode} -- returns true if containerNode was passed and valid or
 *           returns the DoM Node for the buttons container
 */
declare function createSelectUnselectAllCheckboxes(parameters: {
    form: string | HTMLFormElement;
    formName: string;
    clickCallback: (elem?: string) => void;
    containerElemType?: string;
    containerNode?: string | HTMLElement;
    stylingOptions?: {
        buttonImageWidth?: string;
        includeText?: boolean;
        useText?: string[];
        alignment?: "right" | "center" | "left";
        includeTextFontSize?: string;
        includeTextFontColor?: string;
        float?: "right" | "left";
    };
    style?: string | string[];
    images?: JSON;
    "select-url"?: string;
    "select-image"?: string;
    "unselect-url"?: string;
    "unselect-image"?: string;
}): HTMLElement;
/**
 * @function isIterable -- tests whether variable is iterable
 * @param {object} obj -- basically any variable that may or may not be iterable
 * @returns boolean - true if iterable, false if not
 */
declare function isIterable(obj: any): boolean;
declare function createGuid(): string;
declare function createFileDownload(parameters: {
    href: string;
    downloadFileName: string;
    newTab?: boolean;
}): void;
/**
 * @function openFileUpload -- creates file input by drag and drop using a z=1 div
 * @param dropDivId -- the id attribute value of the DIV that will be the drop zone
 * @param callback -- a function to receive the input file(s)
 * @param dropDivContainerId -- an outer DIV to contain the drag and drop. Useful to
 *     better decorate the div
 */
declare function openFileUpload(callback: (fileList: FileList) => void): void;
declare function disabledClick(evt: Event): void;
/**
* @function RESTrequest -- REST requests interface optimized for SharePoint server
* @param {THttpRequestParams} elements -- all the object properties necessary for the jQuery library AJAX XML-HTTP Request call
*     properties are:
*        url: string;
*        setDigest?: boolean;
*        method?: THttpMethods;
*        headers?: THttpRequestHeaders;
*        data?: string | ArrayBuffer | null;
*        body?:  same as data
*        successCallback: TSuccessCallback;
*        errorCallback: TErrorCallback;
* @returns {object} all the data or error information via callbacks
*/
declare function RESTrequest(elements: THttpRequestParams): void;
declare function RequestAgain(elements: THttpRequestParams, nextUrl: string, aggregateData: TSPResponseDataProperties[]): Promise<any>;
declare const SPListTemplateTypes: {
    enums: {
        name: string;
        typeId: number;
    }[];
    getFieldTypeNameFromTypeId: (typeId: number) => string;
    getFieldTypeIdFromTypeName: (typeName: string) => number;
}, SPFieldTypes: {
    enums: ({
        name: string;
        metadataType: null;
        typeId: number;
        extraProperties?: undefined;
    } | {
        name: string;
        metadataType: string;
        typeId: number;
        extraProperties: string[];
    } | {
        name: string;
        metadataType: string;
        typeId: number;
        extraProperties?: undefined;
    } | {
        name: string;
        typeId: number;
        metadataType?: undefined;
        extraProperties?: undefined;
    } | {
        name: string;
        typeId: number;
        extraProperties: string[];
        metadataType?: undefined;
    } | {
        name: string;
        metadataType?: undefined;
        typeId?: undefined;
        extraProperties?: undefined;
    })[];
    standardProperties: string[];
    getFieldTypeNameFromTypeId: (typeId: number) => string;
    getFieldTypeIdFromTypeName: (typeName: string) => number | undefined;
};
/**
 * @function getTaxonomyValue -- returns values from single-valued Managed Metadata fields
 * @param {} obj
 * @param {*} fieldName
 * @param {*} returnValue --
 * @returns
 */
declare function getTaxonomyValue(obj: {
    [key: string]: any;
}, fieldName: string, returnValue: string): string;
/**
 * @function parseManagedMetadata -- this function returns the text label of the metadata term store for
 *     updating a Managed Metadata field.
 * @param {} results -- should be results of a REST call
 * @param {*} fieldName -- name of the metadata field, e.g. "MetaMultiField"
 * @returns
 */
declare function dedupJSArray(array: string[] | any[]): any[];
declare function parseValue(results: object, fieldName: string): any;
declare function parseManagedMetadata(results: any, fieldName: string): {};
