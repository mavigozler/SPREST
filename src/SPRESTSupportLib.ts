/* eslint-disable @rushstack/security/no-unsafe-regexp */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable eqeqeq */

import {
	TParsedURL,
	SPSiteRaw,
	SPWebRaw,
/*
	THttpRequestParams, THttpRequestMethods,
		TSPResponseData, TSPResponseDataProperties, TBatchResponse,
		IBatchHTTPRequestForm, TFetchInfo, HttpInfo,
		TBatchResponseRaw,TXmlHttpRequestData,
		HttpStatus */
} from "./SPComponentTypes.d";
import { RESTrequest } from "./SPHttpReqResp";
import { TArrayToJSON, THttpRequestProtocol } from "./SPHttp";
import "core-js/features/object/from-entries";
//import { RequestAgain } from "./SPHttpReqResp";

export {
	SPServerREST, // class
	//SPListColumnCopy,
	// formatRESTBody,
	checkEntityTypeProperty,
	ParseSPUrl,
	serialSPProcessing,
	constructQueryParameters,
	//SPListTemplateTypes,

	isIterable,
	createGuid,

	formatDateToMMDDYYYY,
	fixValueAsDate,

	buildSelectSet,
	createSelectUnselectAllCheckboxes,
	createFileDownload,
	openFileUpload,
	//getTaxonomyValue
};

const ListItemEntityTypeRE = /EntityType|EntityTypeName|ListItemEntityTypeFullName|metadata/,
	ListFieldEntityTypeRE = /Entity.*Type|metadata/,
	ListEntityTypeRE = /List.*Entity.*Type|Entity.*Type|metadata/,
	ContentTypeEntityTypeRE = /Content.*Entity.*Type|Entity.*Type|metadata/,
	ViewEntityTypeRE = /View.*Entity.*Type|Entity.*Type|metadata/,

//	emailAddressRegex =
// /(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])/;


// public

export
class SPServerREST {
	URL: string;
	apiPrefix: string;

	constructor (setup: {
		URL: string;
	}) {
		this.URL = setup.URL;
		this.apiPrefix = this.URL + "/_api";
	}

	getSiteProperties(): Promise<SPSiteRaw> {
		return new Promise((resolve, reject) => {
			this.getRequest<SPSiteRaw>(this.apiPrefix + "/site")
				.then((response: SPSiteRaw) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	getWebProperties(): Promise<SPWebRaw> {
		return new Promise((resolve, reject) => {
			this.getRequest<SPWebRaw>(this.apiPrefix + "/web")
				.then((response: SPWebRaw) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	getEndpoint(endpoint: string): Promise<unknown> {
		return new Promise((resolve, reject) => {
			this.getRequest<unknown>(this.apiPrefix + endpoint)
				.then((response: unknown) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	getRootweb(): Promise<SPWebRaw> {
		return new Promise((resolve, reject) => {
			this.getRequest<SPWebRaw>(this.apiPrefix + "/site/rootweb?$select=Id,Title,ServerRelativeUrl")
				.then((response: SPWebRaw) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	getRequest<T>(url: string): Promise<T> {
		return new Promise<T>((resolve, reject) => {
			RESTrequest({
				url: url,
				successCallback: (data, httpInfo) => {
					resolve(data as T);
				},
				errorCallback: (err, info) => {
					reject(err);
				}
			});
		});
	}
}

/**
 *
 * @param siteURL -- string representing URL of site
 * @param listNameOrGUID -- name or GUID of list
 * @param sourceColumnIntName -- internal field name to be copied
 * @param destColumnIntName -- internal field name where copies are created
 */
// public

/*
function SPListColumnCopy(
	siteURL: string,
	listNameOrGUID: string,
	sourceColumnIntName: string,
	destColumnIntName: string // this must already exist on the SP list
): Promise<TBatchResponse | null> {
	return new Promise((resolve, reject) => {
		const guidRE = /[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}/i;

		listNameOrGUID = listNameOrGUID.search(guidRE) >= 0 ? "(guid'" + listNameOrGUID + "')" :
				"/getByTitle('" + listNameOrGUID + "')";

		RESTrequest({
			url: siteURL + "/_api/web/lists" + listNameOrGUID + "/items?$select=Id," + sourceColumnIntName,
			method: "GET",
			headers: {
				"Accept": "application/json;odata=verbose",
				"Content-Type": "application/json;odata=verbose",
			},
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			successCallback: (data: TSPResponseData
				 , text, reqObj ) => {
				const requests: IBatchHTTPRequestForm[] = [];
					//host: string = siteURL.match(/https:\/\/[^\/]+/)![0],
					//responses: string = "";
				let body: {[key: string]: unknown};

				for (const response of data as {[key: string]: unknown}) {
					body = {};
					body.__metadata = {type: response.__metadata.type};
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					body[destColumnIntName] = response[sourceColumnIntName as any];
					requests.push({
						url: siteURL + "/_api/web/lists" + listNameOrGUID + "/items(" + response.ID + ")",
						contextinfo: siteURL,
						body: body,
						method: "PATCH",
						headers: {
						contextinfo: siteURL,
						body: body,
						method: "PATCH",
						headers: {
							"Content-Type": "application/json;odata=verbose",
							"Accept": "application/json;odata=nometadata",
							"IF-MATCH": "*",
					//		"X-HTTP-METHOD": "MERGE"
						}
					});
				}

				batchRequestingQueue(
					});
				}

				batchRequestingQueue(
					requests
				).then((response: TBatchResponse | null) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
				});
			},
			errorCallback: (reqObj: JQueryXHR
				, status, errThrown ) => {
				reject(reqObj);
			}
		});
	});
} */

/*
 * @function format RESTBody -- creates a valid JSON object as string for specifying SP list item updates/creations
 * @param {object}  JsonBody -- JS object which conforms to JSON rules. Must be "field_properties" : "field_values" format
 *                    If one of the properties is "__SetType__", it will fix the "__metadata" property
 */
// public

/*function format  RESTBody(JsonBody: TArrayToJSON ):  {
	//let testString: string,
	//	testBody: object;
/*
   try {
      //testString = JSON.stringify(JsonBody);
      // testBody = JSON.parse(testString);
   } catch (e) {
      throw "The argument is not a JavaScript object with quoted properties and quoted values (JSON format)";
   } */
/*return processObjectLevel(JsonBody) as string;

	function processObjectLevel(
		objPart:  {[key: string]: unknown}
	): unknown {
		const newPart: {[key: string]: unknown} = {};

		if (typeof objPart == "string")
			return objPart;
		for (const property in objPart as {[key: string]: unknown})
			if (property == "__SetType__")
				newPart["__metadata"] = { "type" : objPart[property]  };
			else if (typeof objPart[property] == "object") {
				if (objPart[property] == null)
					newPart[property] = "null";
				else
					newPart[property] = processObjectLevel(objPart[property] as  {[key: string]: unknown});
			} else if (objPart[property] instanceof Date == true)
				newPart[property] = (objPart[property] as Date).toISOString();
			else
				newPart[property] = objPart[property];
		return newPart;
	}

} */

// public

function checkEntityTypeProperty(body: object, typeCheck: string): boolean {
	let checkRE: RegExp;

	if (typeCheck == "item")
		checkRE = //SPRESTGlobals.
			ListItemEntityTypeRE;
		checkRE = //SPRESTGlobals.
			ListItemEntityTypeRE;
	else if (typeCheck == "list")
		checkRE = //SPRESTGlobals.
			ListEntityTypeRE;
		checkRE = //SPRESTGlobals.
			ListEntityTypeRE;
	else if (typeCheck == "field")
		checkRE = //SPRESTGlobals.
			ListFieldEntityTypeRE;
		checkRE = //SPRESTGlobals.
			ListFieldEntityTypeRE;
	else if (typeCheck == "content type")
		checkRE = // SPRESTGlobals.
			ContentTypeEntityTypeRE;
		checkRE = // SPRESTGlobals.
			ContentTypeEntityTypeRE;
	else if (typeCheck == "view")
		checkRE = //SPRESTGlobals.
			ViewEntityTypeRE;
		checkRE = //SPRESTGlobals.
			ViewEntityTypeRE;
	else
		return false;
	for (const property in body)
		if (property.search(checkRE) >= 0)
			return true;
	return false;
}

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

// public

function ParseSPUrl (url: string): TParsedURL | undefined {
	const urlRE = /(https?):\/\/([^/]+)|(\/[^/?]+)/g;
	let index: number,
		urlParts: RegExpMatchArray | null,
		query: URLSearchParams | null = null,
		listConfirmed: boolean = false,
		siteFullPath: string = "",
		sitePartialPath: string | null = "",
		libpath: string | null = null,
		pathDone: boolean = false,
		fName: string | null = null,
		listName: string | null = null;
	const tmpArr: unknown[] = [],
		finalParsedUrl: TParsedURL = {
			originalUrl: "",
			protocol: null,
			server: "",
			hostname: "",
			siteFullPath: "",
			sitePartialPath: "",
			list: null,
			listConfirmed: false,
			libRelativeUrl: null,
			file: null,
			query: null
		};

	if (!url)
		return null;
	if (typeof url != "string")
		throw Error("parameter 'url' must be of type string and be a valid url");
	if ((urlParts = url.match(urlRE)) == null)
		return null;
	finalParsedUrl.originalUrl = location.href;
	finalParsedUrl.protocol = urlParts[0].substring(0, urlParts[0].indexOf("://")) as THttpRequestProtocol;
	finalParsedUrl.hostname = urlParts[0].substring(urlParts[0].indexOf("://") + 3);
	finalParsedUrl.server = finalParsedUrl.hostname;
	if ((index = url.lastIndexOf("?")) >= 0)
		query = new URLSearchParams(url.substring(index));
	for (let i = 1; i < urlParts.length; i++)
		if (urlParts[i] == "/Lists") {
			listName = urlParts[i + 1].substring(1);
			pathDone = true;
			finalParsedUrl.listConfirmed = true;
			finalParsedUrl.listConfirmed = true;
			i++;
		} else if (urlParts[i].search(/^\/SiteAssets/) == 0 ||
					urlParts[i].search(/^\/SitePages|^\/Pages/) == 0) {  // next url part is .aspx
			pathDone = true;
			sitePartialPath = null;
			listName = urlParts[i].substring(1);
			libpath = siteFullPath + "/" + listName;
		} else if (urlParts[i] == "/Forms") {
			pathDone = true;
			sitePartialPath = null;
			siteFullPath = siteFullPath.substring(0, siteFullPath.length - urlParts[i - 1].length);
			listName = urlParts[i - 1].substring(1);
			listConfirmed = true;
			libpath = siteFullPath + "/" + listName;
		} else if (pathDone == false && i < urlParts.length - 1) {
			siteFullPath += urlParts[i];
			sitePartialPath += urlParts[i];
		} else if (i == urlParts.length - 1) {
			if (urlParts[i].search(/\.aspx$/i) > 0)
				fName = urlParts[i].substring(1);
			else
				siteFullPath += urlParts[i];
		}
	if (fName != null)
		finalParsedUrl.file = fName;
	if (libpath != null)
		finalParsedUrl.libRelativeUrl = libpath;
	finalParsedUrl.list = listName;
	finalParsedUrl.listConfirmed = listConfirmed;
	if (siteFullPath == sitePartialPath)
		sitePartialPath = null;
	finalParsedUrl.siteFullPath = siteFullPath;
	if (sitePartialPath != null)
		finalParsedUrl.sitePartialPath = sitePartialPath;
	if (query) {
		for (const pair of query.entries())
			tmpArr.push({key: pair[0], value: pair[1]});
		finalParsedUrl.query = tmpArr;
	}
	return finalParsedUrl;
}

/**
 * @function serialSPProcessing --
		finalParsedUrl.query = tmpArr;
	}
	return finalParsedUrl;
}

/**
 * @function serialSPProcessing --
 * @param {function} opFunction
 * @param {arrayOfObject} dataset -- this must be an array of objects in JSON format that
 * 				represent the body of a POST request to create SP data: fields, items, etc.
 * @returns -- the return operations are to unwind the recursion in the processing
 *           to get to either resolve or reject operations
 */
// public

function serialSPProcessing(
	opFunction: (arg1: unknown) => Promise<unknown>,
	itemset: unknown[]
): Promise<unknown> {
	const responses: unknown[] = [ ];

	return new Promise((resolve) => {
		function iterate(index: number): void {
			let datum: unknown;

			if ((datum = itemset[index]) == null)
				resolve(responses);
			else
				opFunction(datum).then((response: unknown) => {
					responses.push(response);
					iterate(index + 1);
				}).catch((err: unknown) => {
					responses.push(err);
					iterate(index + 1);
				});
		}
		iterate(0);
	});
}

// public
type keyed = {[key:string]: string | string[];};

function constructQueryParameters(
	parameters: string | {query: string} | keyed
): string {
	const odataFunctions = ["filter", "select", "expand", "top", "count", "skip"];
	let query: string = "";

	if (!parameters)
		return "";
	if (typeof parameters == "string")  // not even an object
		return parameters;
	if (parameters.query && typeof parameters.query == "string")
		return parameters.query;
	for (const property in parameters) {
		if (odataFunctions.find(elem => elem == property) == null)
			continue;
		if (query == "")
			query = "?";
		else
			query += "&";
		const keyedparams: keyed = parameters as keyed;
		if (Array.isArray(keyedparams[property]) == true)
			query += (keyedparams[property] as string[]).join(",");
		else if (new RegExp(property).test(keyedparams[property] as string) == true)
			query += keyedparams[property];
		else
			query += "$" + property + "=" + keyedparams[property];
	}
	return query;
}

/**
 * @function formatDateToMMDDYYYY -- returns ISO 8061 formatted date string to MM[d]DD[d]YYYY date
 *      string, where [d] is character delimiter
 * @param {string|datetimeObj} dateInput [required] -- ISO 8061-formatted date string
 * @param {string} delimiter [optional] -- character that will delimit the result string; default is '/'
 * @returns {string} -- MM[d]DD[d]YYYY-formatted string.
 */
// public

function formatDateToMMDDYYYY(
	dateInput: string,
	delimiter: string = "/"
): string | null {
	let matches: RegExpMatchArray | null,
		dateObj: Date;

	if ((matches = dateInput.match(/(^\d{4})-(\d{2})-(\d{2})T/)) != null)
		return matches[2] + delimiter + matches[3] + delimiter + matches[1];
	if ((dateObj = new Date(dateInput)) instanceof Date == true)
		return (dateObj.getMonth() < 9 ? "0" : "") + (dateObj.getMonth() + 1) + delimiter +
			(dateObj.getDate() < 10 ? "0" : "") + dateObj.getDate() + delimiter + dateObj.getFullYear();
	return null;
}


// public

function fixValueAsDate(date: Date): Date | null {
	const datestring: RegExpMatchArray | null = date.toISOString().match(/(\d{4})-(\d{2})-(\d{2})/);

	if (datestring == null)
		return null;
	return new Date(parseInt(datestring[1]), parseInt(datestring[2]) - 1, parseInt(datestring[3]));
}

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
// public

function buildSelectSet(
	parentNode: HTMLElement,   // DOM node to append the construct
	nameOrGuid: string,   // used in tagging
	options: {
		selected?: boolean,
		required?: boolean,
		value: string,
		text: string
	}[],      // array of option objects {text: text-to-display, value: node value to set}
	availableSize: number, // for number of choices displayed for available choices select list
	selectedSize: number, // for number of choices displayed for selected choices select list
	onChangeHandler: (arg?: unknown) => object | void // callback
): string | void {
	const divNode: HTMLDivElement = document.createElement("div");
	let selectElem: HTMLSelectElement,
		optionElem: HTMLOptionElement,
		pNode: HTMLParagraphElement,
		spanElem: HTMLSpanElement,
		bNode: HTMLButtonElement;

	parentNode.appendChild(divNode);
	divNode.style.display = "grid";
	divNode.style.gridTemplateColumns = "auto auto auto";
	divNode.className = "select-set";
	const guid = createGuid();
	divNode.id = guid;

	pNode = document.createElement("p");
	divNode.appendChild(pNode);
	spanElem = document.createElement("span");
	pNode.appendChild(spanElem);
	spanElem.appendChild(document.createTextNode("Available"));
	spanElem.style.font = "italic 9pt Arial,sans-serif";
	spanElem.style.display = "block";
	selectElem = document.createElement("select");
	pNode.appendChild(selectElem);
	selectElem.name = nameOrGuid + "-avail";
	selectElem.multiple = true;
	selectElem.size = availableSize;
	selectElem.style.minWidth = "10em";
	selectElem.addEventListener("change", onChangeHandler);
	for (const option of options)
		if (!option.selected)	{
			optionElem = document.createElement("option");
			selectElem.appendChild(optionElem);
			optionElem.value = option.value;
			if (option.required == true)
				optionElem.className = "required";
			optionElem.appendChild(document.createTextNode(option.text));
		}
	//	selectElem["data-options"] = JSON.stringify(options);

	pNode = document.createElement("p");
	divNode.appendChild(pNode);
	pNode.style.alignSelf = "center";
	pNode.style.padding = "0 0.5em";
	bNode = document.createElement("button");
	pNode.appendChild(bNode);
	bNode.type = "button";
	bNode.style.display = "block";
	bNode.style.margin = "1em auto";
	bNode.appendChild(document.createTextNode("==>>"));
	bNode.addEventListener("click", (event) => {
		// available to chosen
		const form = (event.currentTarget as HTMLInputElement).form,
			availOptions: HTMLOptionsCollection = form![nameOrGuid + "-avail"].selectedOptions,
			chosenSelect: HTMLSelectElement = form![nameOrGuid + "-selected"];
		let option: HTMLOptionElement;

		for (let i: number = 0; i < availOptions.length; i++) {
			option = availOptions[i];
			chosenSelect.appendChild(option.parentNode!.removeChild(option));
		}
		sortSelect(chosenSelect);
	});
	bNode = document.createElement("button");
	pNode.appendChild(bNode);
	bNode.type = "button";
	bNode.style.display = "block";
	bNode.style.margin = "1em auto";
	bNode.appendChild(document.createTextNode("<<=="));
	bNode.addEventListener("click", (event) => {
		// chosen to available
		const form = (event.currentTarget as HTMLFormElement).form,
			availSelect: HTMLSelectElement = form[nameOrGuid + "-avail"],
			chosenOptions: HTMLOptionsCollection = form[nameOrGuid + "-selected"].selectedOptions;
		let option: HTMLElement;

		for (let i = 0; i < chosenOptions.length; i++) {
			option = chosenOptions[i];
			availSelect.appendChild(option.parentNode!.removeChild(option));
		}
		sortSelect(availSelect);
	});

	pNode = document.createElement("p");
	divNode.appendChild(pNode);
	spanElem = document.createElement("span");
	pNode.appendChild(spanElem);
	spanElem.appendChild(document.createTextNode("Selected"));
	spanElem.style.font = "italic 9pt Arial,sans-serif";
	spanElem.style.display = "block";
	selectElem = document.createElement("select");
	pNode.appendChild(selectElem);
	selectElem.name = nameOrGuid + "-selected";
	selectElem.multiple = true;
	selectElem.style.minWidth = "10em";
	selectElem.size = selectedSize;
	selectElem.addEventListener("change", onChangeHandler);
	for (const option of options)
		if (option.selected && option.selected == true)	{
			optionElem = document.createElement("option");
			selectElem.appendChild(optionElem);
			optionElem.value = option.value;
			if (option.required == true)
				optionElem.className = "required";
			optionElem.appendChild(document.createTextNode(option.text));
		}
	sortSelect(selectElem);
	return guid;


	//	selectElem["data-options"] = JSON.stringify([]);
	function sortSelect(selectObj: HTMLSelectElement): void {
		const selectArray: {
			display: string;
			value: string;
			required: boolean;
		}[] = [ ];
		let node: HTMLOptionElement;

		const selectOptions = Array.from(selectObj.options);
		for (const opshun of selectOptions)
			selectArray.push({
				display: (opshun.firstChild as Text).data,
				value: opshun.value,
				required: opshun.className.search(/required/) >= 0 ? true : false
			});
		selectArray.sort((elem1, elem2) => {
			return elem1.display > elem2.display ? 1 :
				elem1.display < elem2.display ? -1 : 0;
		});
		while (selectObj.firstChild)
			selectObj.removeChild(selectObj.firstChild);
		for (const elem of selectArray) {
			node = document.createElement("option");
			selectObj.appendChild(node);
			node.value = elem.value;
			node.appendChild(document.createTextNode(elem.display));
			if (elem.required == true)
				node.className = "required";
		}
	}
}

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
// public

function createSelectUnselectAllCheckboxes(parameters: {
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
}): HTMLElement {
	const containerElemType: string = parameters.containerElemType ?? "p",
		capture: boolean = false;
	let
		includeText: boolean = false,
		useText: string[] = ["select all", "unselect all"],
		buttonImageWidth: string = "15px",
		includeTextFontSize: string = "8pt",
		includeTextFontColor: string = "brown",
		alignment: "right" | "center" | "left" = "left",
		control: RadioNodeList,
		setStyle: string | null = null,
		form: HTMLFormElement,
		parent: HTMLElement | null;

	if (typeof parameters.form == "string")
		form = document.getElementById(parameters.form)! as HTMLFormElement;
	else
		form = parameters.form;
	if (!(form && form.nodeName && form.nodeName.toLowerCase() == "form"))
		throw Error("A form DOM node or its 'id' attribute must be a parameter and contain the checkbox settings");

	// look for a form that contains this
	if (!parameters.formName || typeof parameters.formName != "string" || parameters.formName.length == 0)
		throw Error("The parameter 'formName' is either undefined or not a string of non-zero length");
	if ((control = form[parameters.formName]) == null)
		throw Error("No 'form' with that name attribute is found on the document. It must be present in HTML document.");

	if (typeof parameters.containerNode == "string") { // is an id attribute
		if ((parent = document.getElementById(parameters.containerNode)) == null)
			throw Error("A container node ID but did not correspond to an existing HTML element");
	} else if (typeof parameters.containerNode == "object")
		parent = parameters.containerNode;
	else
		parent = document.createElement(containerElemType);

	if (parameters.stylingOptions) {
		includeText = parameters.stylingOptions.includeText || includeText;
		if (parameters.stylingOptions.useText && parameters.stylingOptions.useText.length > 1)
			useText = parameters.stylingOptions.useText;
		alignment = parameters.stylingOptions.alignment || alignment;
		buttonImageWidth = parameters.stylingOptions.buttonImageWidth || buttonImageWidth;
		includeTextFontSize = parameters.stylingOptions.includeTextFontSize || includeTextFontSize;
		includeTextFontColor = parameters.stylingOptions.includeTextFontColor || includeTextFontColor;
	}
	if (parameters.style && typeof parameters.style == "string" && parameters.style.search(/\{/) < 0)
		parent.setAttribute("style", setStyle = parameters.style);
	if (includeText == true) {
		if (setStyle && setStyle.search(/text-align/) < 0)
			parent.style.textAlign = alignment;
		if (setStyle && setStyle.search(/font-size/) < 0)
			parent.style.fontSize = includeTextFontSize;
		if (setStyle && setStyle.search(/[^-]color/) < 0)
			parent.style.color = includeTextFontColor;
	}
	if (parameters.stylingOptions && parameters.stylingOptions.float)
		parent.style.float = parameters.stylingOptions.float;

	buildButtonWithText(
		control,
		true,
		parent,
		parameters["select-url"] ?? parameters["select-image"] as string,
		"Select All",
		buttonImageWidth,
		useText[0],
		parameters.clickCallback,
		capture,
		"select",
	);
	if (includeText == true)
		parent.appendChild(document.createElement("br"));

	buildButtonWithText(
		control,
		false,
		parent,
		parameters["unselect-url"] ?? parameters["unselect-image"] as string,
		"Unselect All",
		buttonImageWidth,
		useText[1],
		parameters.clickCallback,
		capture,
		"unselect",
	);
	return parent;

	function buildButtonWithText(
		inputNodeList: RadioNodeList,
		checkSet: boolean,
		container: HTMLElement,
		usedImageUrl: string,
		altText: string,
		usedImageWidth: string,
		includedText: string,
		clickCallback: (which: string) => void,
		capture: boolean,
		callbackArg: string
	): void {
		let imgSrc: string;

		const spanElem = document.createElement("span");
		container.appendChild(spanElem);

		const imgElem = document.createElement("img");
		if (typeof includedText == "string" && includedText.length > 0) {
			spanElem.appendChild(document.createTextNode(includedText));
			imgElem.style.marginLeft = "0.3em";
		}
		spanElem.appendChild(imgElem);
		imgElem.addEventListener("click", () => {
			const nodeList = Array.from(inputNodeList);
			for (const checkbox of nodeList)
				(checkbox as HTMLInputElement).checked = checkSet;
			if (clickCallback)
				clickCallback(callbackArg);
		}, capture);
		imgSrc = usedImageUrl;
		if (!imgSrc || imgSrc == null) {
			imgElem.style.width = usedImageWidth;
			imgSrc = UnselectAllCheckboxes;
		}
		imgElem.src = imgSrc;
		imgElem.alt = altText;
	}
}


/**
 * @function isIterable -- tests whether variable is iterable
 * @param {object} obj -- basically any variable that may or may not be iterable
 * @returns boolean - true if iterable, false if not
 */

// public

function isIterable(obj: unknown[]): boolean {
	// checks for null and undefined
	if (obj == null)
		return false;
	return typeof obj[Symbol.iterator] === "function";
}

// off the Internet
// public
function createGuid(): string {
	function S4(): string {
		return (1 + Math.random() * 0x10000 | 0).toString(16).substring(1);
	}
	return (S4() + S4() + "-" + S4() + "-4" + S4().substring(0,3) + "-" + S4() + "-" + S4() + S4() + S4()).toLowerCase();
}


function createFileDownload(parameters: {
	href: string;
	downloadFileName: string;
	newTab?: boolean;
}): void {
	const aNode = document.createElement("a");

	aNode.setAttribute("href", parameters.href);
	aNode.setAttribute("download", parameters.downloadFileName);
	aNode.style.display = "none";
	if (parameters.newTab == true)
		aNode.target = "_blank";
	document.body.appendChild(aNode);
	aNode.click();
	document.body.removeChild(aNode);
}

/* To use openFileUpload, you must use drag and drop
    In the document, create two undisplayed DIV settings
	 1. the outer one is the containing DIV. the inner drop zone DIV can be decorated to
	    stand out
	2. the inner one is the drop zone div. You can also insert text within or outside
	   the div to instruct the user
	Call the function with the id attributes of the appropriate DIVs
*/

/**
 * @function openFileUpload -- creates file input by drag and drop using a z=1 div
 * @param dropDivId -- the id attribute value of the DIV that will be the drop zone
 * @param callback -- a function to receive the input file(s)
 * @param dropDivContainerId -- an outer DIV to contain the drag and drop. Useful to
 *     better decorate the div
 */
// public

function openFileUpload(
	callback: (fileList: FileList) => void
): void {
	const containerDiv: HTMLDivElement = document.createElement("div"),
		ddDiv: HTMLDivElement = document.createElement("div"),
		closeSpan: HTMLSpanElement = document.createElement("span"),
		closeButton: HTMLButtonElement = document.createElement("button"),
		closeImg: HTMLImageElement = document.createElement("img"),
		promptPara: HTMLParagraphElement = document.createElement("p");

	if (document.getElementById("drop-container") == null) { // build only if it exists
		containerDiv.appendChild(closeSpan);
		containerDiv.appendChild(ddDiv);
		closeSpan.appendChild(closeButton);
		closeButton.appendChild(closeImg);
		ddDiv.appendChild(promptPara);
		containerDiv.id = "drop-container";
		ddDiv.id = "drop-zone";
		closeSpan.id = "drag-drop-close-span-container";
		closeButton.type = "button";
		closeButton.className = "wrapping-img";
		closeImg.src = "images/close-icon.png";
		closeImg.alt = "close this!";
		closeImg.id = "cancelingUpload";
		promptPara.appendChild(document.createTextNode(
			"Drag one JSON file representing the agenda templates back up within this " +
			"zone and drop it."
		));
		ddDiv.addEventListener("drop", (evt: Event) => {
			evt.stopPropagation();
			evt.preventDefault();
			containerDiv.style.display = "none";
			document.body.removeEventListener("click", disabledClick, true);
			callback((evt as InputEvent).dataTransfer!.files);
		});
		ddDiv.addEventListener("dragover", (evt: Event) => {
			evt.stopPropagation();
			evt.preventDefault();
		});
		// Styling
		containerDiv.style.display = "none";
		containerDiv.style.zIndex = "2";
		containerDiv.style.backgroundColor = "#f8ffff";
		containerDiv.style.position = "fixed";
		containerDiv.style.top = "30%";
		containerDiv.style.left = "25%";
		containerDiv.style.width = "40%";
		containerDiv.style.height = "40%";
		containerDiv.style.padding = "0.5em 1em 1em 1em";
		containerDiv.style.border = "10px groove darkgreen";
		containerDiv.style.overflow = "hidden";

		ddDiv.style.border = "1px dashed black";
		ddDiv.style.backgroundColor = "#f8fff8";
		ddDiv.style.height = "90%";
		ddDiv.style.width = "100%";
		ddDiv.style.clear = "both";

		promptPara.style.fontSize = "14pt";
		promptPara.style.fontFamily = "'Segoe UI',Tahoma,sans-serif";
		promptPara.style.width = "70%";
		promptPara.style.margin = "10% auto";

		closeButton.style.background = "none";
		closeButton.style.border = "none";
		closeButton.style.margin = "-1em -1em 0 0";

		closeSpan.style.float = "right";

		closeImg.style.width = "20px";


		document.body.appendChild(containerDiv);
	}
	document.body.addEventListener("click", disabledClick, true);
	containerDiv.style.display = "block";
}

function disabledClick(evt: Event): void {
	if ((evt.target as HTMLImageElement).id == "cancelingUpload") {
		document.getElementById("drop-container")!.style.display = "none";
		document.body.removeEventListener("click", disabledClick, true);
	}
	evt.stopImmediatePropagation();
	evt.preventDefault();
}

/**
* @function RESTrequest -- REST requests interface optimized for SharePoint server
* @param {THttpRequestParams} settings -- all the object properties necessary for the jQuery library AJAX XML-HTTP Request call
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

// public

/*
function AjaxRESTrequest(settings: THttpRequestParams): void {
	const editedUrl: string = (settings.url.indexOf("http") == 0 ? settings.url :
			"https://" + settings.url);
	if (settings.setDigest && settings.setDigest == true) {
		const match: RegExpMatchArray = editedUrl.match(/(.*\/_api)/) as RegExpMatchArray;

		$.ajax({  // get digest token
			url: match[1] + "/contextinfo",
			method: "POST",
			headers: {...SPstdHeaders},
			success: function (data) {
				let headers: THttpRequestHeaders = settings.headers as THttpRequestHeaders;

				if (headers) {
					headers["Content-Type"] = "application/json;odata=verbose";
					headers["Accept"] = "application/json;odata=verbose";
				} else
					headers = {...SPstdHeaders};
				headers["X-RequestDigest"] = data.FormDigestValue ? data.FormDigestValue :
					data.d.GetContextWebInformation.FormDigestValue;
				$.ajax({
					url: editedUrl,
					method: settings.method ? settings.method : "GET",
					headers: headers,
					beforeSend: (xhr) => {
						if (settings.customHeaders)
							for (const header in settings.customHeaders)
								xhr.setRequestHeader(header, settings.customHeaders[header]);
					},					data: settings.body || settings.data as string,
					success: function (data: TSPResponseData, status: string, requestObj: JQueryXHR) {
						settings.successCallback!(data, status, requestObj);
					},
					error: function (requestObj: JQueryXHR, status: string, thrownErr: string) {
						settings.errorCallback!(requestObj, status, thrownErr);
					}
				});
			},
			error: function (requestObj, status, thrownErr) {
				settings.errorCallback!(requestObj, status, thrownErr);
			}
		});
	} else {
		if (!settings.headers)
			settings.headers = {...SPstdHeaders};
		else {
			settings.headers["Content-Type"] = "application/json;odata=verbose";
			settings.headers["Accept"] = "application/json;odata=verbose";
		}
		$.ajax({
			url: editedUrl,
			method: settings.method,
			headers: settings.headers,
			beforeSend: (xhr) => {
				if (settings.customHeaders)
					for (const header in settings.customHeaders)
						xhr.setRequestHeader(header, settings.customHeaders[header]);
			},
			data: settings.body || settings.data as string,
			success: function (data: TSPResponseData, status: string, requestObj: JQueryXHR) {
				if (data.d && data.d.__next)
					AjaxRequestAgain(
						settings,
						data.d.__next,
						data.d.results!
					).then((response: unknown) => {
						settings.successCallback!(response as TSPResponseData);
					}).catch((err: JQueryXHR) => {
						settings.errorCallback!(err);
					});
				else
					settings.successCallback!(data, status, requestObj);
			},
			error: function (requestObj, status, thrownErr) {
				settings.errorCallback!(requestObj, status, thrownErr);
			}
		});
	}
}


// public
function AjaxRequestAgain(
	settings: THttpRequestParams,
	nextUrl: string,
	aggregateData: TSPResponseDataProperties[]
): Promise<unknown> {
	return new Promise((resolve, reject) => {
		$.ajax({
			url: nextUrl,
			method: "GET",
			headers: settings.headers,
			success: (data) => {
				if (data.d.__next) {
					aggregateData = aggregateData.concat(data.d.results)
					AjaxRequestAgain(
						settings,
						data.d.__next,
						aggregateData
					).then((response: unknown) => {
						resolve(response);
					}).catch((response: unknown) => {
						reject(response);
					});
				} else
					resolve(aggregateData.concat(data.d.results));
			},
			error: (reqObj) => {
				reject(reqObj);
			}
		});
	});
}
*/

/*
	static httpRequest(settings: THttpRequestParams): void {
		if (settings.setDigest && settings.setDigest == true) {
			const match = settings.url.match(/(.*\/_api)/) as string[];
			fetch(match[1] + "/contextinfo", {  // get digest token
				method: "POST",
				headers: {...SPServerREST.stdHeaders}
			} as RequestInit)
			.then((response: Response) => response.json()
			.then(json => {
				const spResponse: TSPResponseData = json;
				let headers = settings.headers;

				if (typeof headers == "undefined")
					headers = {...SPServerREST.stdHeaders};
				else {
					headers["Content-Type"] = "application/json;odata=verbose";
					headers["Accept"] = "application/json;odata=verbose";
				}
				headers["X-RequestDigest"] = spResponse.FormDigestValue ? spResponse.FormDigestValue :
						spResponse.d?.GetContextWebInformation!.FormDigestValue;
				const fetchSettings: RequestInit = {
					method: !settings.method ? "GET" : settings.method,
					headers: headers as HeadersInit,
					body: (settings.body ?? settings.data) as BodyInit,
				};
				fetch(settings.url, fetchSettings)
				.then((response: Response) => response.json()
				.then(json => {
					const headers = Object.fromEntries(response.headers.entries());
					const httpInfo: HttpInfo = {
						...settings,
						bodyUsed: response.bodyUsed,
						ok: response.ok,
						redirected: response.redirected,
						status: response.status,
						statusText: response.statusText,
						type: response.type,
						url: response.url
					} as HttpInfo;
					settings.successCallback(json, headers, httpInfo);
				}).catch(err => {
					// TODO errinfo here
					settings.errorCallback(err)
				}));
			}).catch((err: unknown) => {
				settings.errorCallback(err);
			}));
		} else {
			if (!settings.headers)
				settings.headers = {...SPServerREST.stdHeaders};
			else {
				settings.headers["Content-Type"] = "application/json;odata=verbose";
				settings.headers["Accept"] = "application/json;odata=verbose";
			}
			const fetchSettings: RequestInit = {
				method: !settings.method ? "GET" : settings.method,
				headers: settings.headers as HeadersInit,
				body: settings.body ?? settings.data as BodyInit,
			};
			fetch(settings.url, fetchSettings)
			.then((response: Response) => response.json()
			.then((json) => {
				const spResponse: TSPResponseData = json;
				const headers = Object.fromEntries(response.headers.entries());
				if (spResponse.d && spResponse.d.__next && spResponse.d.results)
					RequestAgain(
						settings,
						spResponse.d.__next,
						spResponse.d.results as TSPResponseDataProperties[]
					).then((response: TSPResponseDataProperties[]) => {
	/*					const httpInfo: HttpInfo = {
							...settings,
							bodyUsed: response.bodyUsed,
							ok: response.ok,
							redirected: response.redirected,
							status: response.status,
							statusText: response.statusText,
							type: response.type,
							url: response.url
						} as HttpInfo;
						settings.successCallback(response as unknown as TSPResponseData, headers, /*httpInfo);
					}).catch((err: unknown) => {
						// TODO need to convert err to errInfo
						const errInfo = err;
						settings.errorCallback(errInfo);
					});
				else
					settings.successCallback(json, headers, settings);
			}).catch((err: unknown) => {
				settings.errorCallback(err);
			}));
		}
	}

	static httpRequestPromise(parameters: {
		url: string;
		setDigest?: boolean;
		method?: THttpRequestMethods;
		headers?: THttpRequestHeaders;
		data?: TXmlHttpRequestData;
		body?: TXmlHttpRequestData;
	}): Promise<unknown> {
		// eslint-disable-next-line @typescript-eslint/no-unused-vars
		return new Promise((resolve, reject) => {
			SPServerREST.httpRequest({
				setDigest: parameters.setDigest,
				url: parameters.url,
				method: parameters.method,
				headers: parameters.headers,
				data: parameters.data ?? parameters.body as string,
				successCallback: (data: TSPResponseData) => {
					resolve({data: data});
				},
				// eslint-disable-next-line @typescript-eslint/no-unused-vars
				errorCallback: (errInfo: unknown) => {
					//reject(reqObj, errInfo);
				}
			});
		});
	} */

// public
/*
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
   }, */
/*
		standardProperties: [
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
		],
		getFieldTypeNameFromTypeId: (typeId: number) => {
			return (SPFieldTypes.enums.find(elem => elem.typeId == typeId))!.name;
		},
		getFieldTypeIdFromTypeName: (typeName: string) => {
			return (SPFieldTypes.enums.find(elem => elem.name == typeName))!.typeId;
		}
	};*/

/**
 * @function getTaxonomyValue -- returns values from single-valued Managed Metadata fields
 * @param {} obj
 * @param {*} fieldName
 * @param {*} returnValue --
 * @returns
 */
// public

/*
function getTaxonomyValue(
	obj: {[key:string]:unknown},
	fieldName: string,
returnValue: string): string {
	// Below function pulled from here and tweaked to work as I needed it.
	// https://sympmarc.com/2017/06/19/retrieving-multiple-sharepoint-managed-metadata-columns-via-rest/
	//
	// Use this function to pull values from single select managed metadata fields
	//
	// To return only the text value (label) of the managed metadata field, pass in a returnValue of 'labelOnly'; otherwise pass in nothing
	// and a properly formatted string will be returned that can be used to update a managed metadata field.
	//
	let metaString = "";

	// Iterate over the fields in the row of data
	for (const field in obj) {
		// If it's the field we're interested in....
		if (obj.hasOwnProperty(field) && field === fieldName) {
			if (obj[field] !== null) {
				// ... get the WssId from the field ...
				const thisId = obj[field].WssId;
				// ... and loop through the TaxCatchAll data to find the matching Term
				for (let i = 0; i < obj.TaxCatchAll.results.length; i++) {
					if (obj.TaxCatchAll.results[i].ID === thisId) {
					// Only return the text value
						if (returnValue && returnValue === "labelOnly") {
							metaString = obj.TaxCatchAll.results[i].Term;
							break;
						} else {
						// Return a formatted value that can be used to update managed metadata values
							metaString = thisId + ";#" + obj.TaxCatchAll.results[i].Term + "|" + obj[field].TermGuid;
							break;
						}
					}
				}
			}
		}
	}
	return metaString;
}*/

/**
 * @function parseManagedMetadata -- this function returns the text label of the metadata term store for
 *     updating a Managed Metadata field.
 * @param {} results -- should be results of a REST call
 * @param {*} fieldName -- name of the metadata field, e.g. "MetaMultiField"
 * @returns
 */



//declare function parseValue(results: object, fieldName: string): any;
/*
function parseManagedMetadata(results: any, fieldName: string) {
	// Build out the managed metadata string needed for later updating.
	// Some links that helped me figure all this out.
	// http://www.aerieconsulting.com/blog/using-rest-to-update-a-managed-metadata-column-in-sharepoint
	// http://www.aerieconsulting.com/blog/update-using-rest-to-update-a-multi-value-taxonomy-field-in-sharepoint
	// http://stackoverflow.com/questions/17076509/set-taxonomy-field-multiple-values
	// Use this function to pull values from multivalue managed metadata fields. Technically it will return single select labels also, but
	// that will depend on the rest call used.
	// First try to pull out a basic label value. If not found then it's probably a multivalue field.

	let  	metaString = "",
			metaSeparator = "",
		labelValue = parseValue(results, fieldName);

	if (labelValue == "" && results[fieldName].results == undefined)
		return {};
	if (labelValue != "")
		// We have a single value field
		metaString = results[fieldName].WssId + ";#"+ labelValue + "|" + results[fieldName].TermGuid;
	else {
		// We have a multivalue field
		let fieldValues = results[fieldName].results;

		for (let i = 0; i < fieldValues.length; i++) {
			metaString += metaSeparator + fieldValues[i].WssId + ";#" + fieldValues[i].Label + "|" + fieldValues[i].TermGuid;
			metaSeparator = ";#";
		}
	}
	return metaString;
} */


const SelectAllCheckboxes =
    "data:image/png;base64,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";

const UnselectAllCheckboxes =
    "data:image/png;base64,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";


/*************************************************
   Formats for Batch Requests

	These examples write to the server.

POST http://{site_url}/_api/$batch
Content-Type: multipart/mixed; boundary=batch_{batch_GUID}
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding: binary

POST {endpoint_url} HTTP/1.1
Content-Type: application/json;odata=verbose
Content-Length: {length}

{request_body}

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

POST {endpoint_url} HTTP/1.1
Content-Type: application/json;odata=verbose
Content-Length: {length}

{request_body}

--batch_{batch_GUID}--

{site_url}: The URL of the SharePoint site.
{batch_GUID}: A unique identifier for the batch request.
     You can generate this using any method you prefer.
{endpoint_url}: The URL of the REST endpoint for the
   operation you want to perform.
{length}: The length of the request body in bytes.
{request_body}: The body of the request in JSON format.

You can include multiple requests in a single batch request by adding more
Content-Type and Content-Transfer-Encoding headers, followed by the endpoint
URL and request body for each additional request.


//////////////  Format for three GET requests

POST http://{site_url}/_api/$batch
Content-Type: multipart/mixed; boundary=batch_{batch_GUID}
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

GET {endpoint_url_1} HTTP/1.1
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

GET {endpoint_url_2} HTTP/1.1
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

GET {endpoint_url_3} HTTP/1.1
Accept: application/json;odata=verbose

--batch_{batch_GUID}--

In this example, you would replace the following variables:

{site_url}: The URL of the SharePoint site.
{batch_GUID}: A unique identifier for the batch request.
{endpoint_url_1}, {endpoint_url_2}, and {endpoint_url_3}: The URLs
of the REST endpoints for the three GET requests.

Note that each request includes an Accept header that specifies
the desired response format. In this case, application/json;odata=verbose
is specified for all three requests, but you can change this to another
format if you prefer.

Also note that each request is separated by a boundary string
(--batch_{batch_GUID}). This boundary string should be unique
and not appear in any of the request URLs or bodies.

*****************************************************/

/*************************************************
   Formats for Batch Requests

	These examples write to the server.

POST http://{site_url}/_api/$batch
Content-Type: multipart/mixed; boundary=batch_{batch_GUID}
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding: binary

POST {endpoint_url} HTTP/1.1
Content-Type: application/json;odata=verbose
Content-Length: {length}

{request_body}

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

POST {endpoint_url} HTTP/1.1
Content-Type: application/json;odata=verbose
Content-Length: {length}

{request_body}

--batch_{batch_GUID}--

{site_url}: The URL of the SharePoint site.
{batch_GUID}: A unique identifier for the batch request.
     You can generate this using any method you prefer.
{endpoint_url}: The URL of the REST endpoint for the
   operation you want to perform.
{length}: The length of the request body in bytes.
{request_body}: The body of the request in JSON format.

You can include multiple requests in a single batch request by adding more
Content-Type and Content-Transfer-Encoding headers, followed by the endpoint
URL and request body for each additional request.


//////////////  Format for three GET requests

POST http://{site_url}/_api/$batch
Content-Type: multipart/mixed; boundary=batch_{batch_GUID}
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

GET {endpoint_url_1} HTTP/1.1
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

GET {endpoint_url_2} HTTP/1.1
Accept: application/json;odata=verbose

--batch_{batch_GUID}
Content-Type: application/http
Content-Transfer-Encoding:binary

GET {endpoint_url_3} HTTP/1.1
Accept: application/json;odata=verbose

--batch_{batch_GUID}--

In this example, you would replace the following variables:

{site_url}: The URL of the SharePoint site.
{batch_GUID}: A unique identifier for the batch request.
{endpoint_url_1}, {endpoint_url_2}, and {endpoint_url_3}: The URLs
of the REST endpoints for the three GET requests.

Note that each request includes an Accept header that specifies
the desired response format. In this case, application/json;odata=verbose
is specified for all three requests, but you can change this to another
format if you prefer.

Also note that each request is separated by a boundary string
(--batch_{batch_GUID}). This boundary string should be unique
and not appear in any of the request URLs or bodies.

*****************************************************/