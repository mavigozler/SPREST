"use strict";

import * as SPRESTGlobals from './SPRESTGlobals';
import * as SPRESTTypes from './SPRESTtypes';

/* jshint -W069, -W119 */

const SPstdHeaders: SPRESTTypes.THttpRequestHeaders = {
	"Content-Type":"application/json;odata=verbose",
	"Accept":"application/json;odata=verbose"
};

let SelectAllCheckboxes: string, // defined below
	UnselectAllCheckboxes: string;

export class SPServerREST {
	URL: string;
	apiPrefix: string;

	static stdHeaders: SPRESTTypes.THttpRequestHeaders = {
		"Accept": "application/json;odata=verbose",
		"Content-Type": "application/json;odata=verbose"
	};

	constructor (setup: {
		URL: string;
	}) {
		this.URL = setup.URL;
		this.apiPrefix = this.URL + "/_api";
	}

	static httpRequest(elements: SPRESTTypes.THttpRequestParams): void {
		if (elements.setDigest && elements.setDigest == true) {
			let match = elements.url.match(/(.*\/_api)/) as string[];
			$.ajax({  // get digest token
				url: match[1] + "/contextinfo",
				method: "POST",
				headers: {...SPServerREST.stdHeaders},
				success: (data: {
					FormDigestValue?: string;
					d: {GetContextWebInformation: {FormDigestValue: string}}
				}) => {
					let headers = elements.headers;

					if (typeof headers == "undefined")
						headers = {...SPServerREST.stdHeaders};
					else {
						headers["Content-Type"] = "application/json;odata=verbose";
						headers["Accept"] = "application/json;odata=verbose";
					}
					headers["X-RequestDigest"] = data.FormDigestValue ? data.FormDigestValue :
						data.d.GetContextWebInformation.FormDigestValue;
					$.ajax({
						url: elements.url,
						method: !elements.method ? "GET" : elements.method,
						headers: headers,
						data: (elements.body ?? elements.data) as SPRESTTypes.TXmlHttpRequestData,
						success: (data: any, status: string, requestObj: JQueryXHR) => {
							elements.successCallback(data, status, requestObj);
						},
						error: function (requestObj: JQueryXHR, status: string, thrownErr: string) {
							elements.errorCallback(requestObj, status, thrownErr);
						}
					});
				},
				error: function (requestObj: JQueryXHR, status: string, thrownErr: string) {
					elements.errorCallback(requestObj, status, thrownErr);
				}
			});
		} else {
			  if (!elements.headers)
				 elements.headers = {...SPServerREST.stdHeaders};
			 else {
				 elements.headers["Content-Type"] = "application/json;odata=verbose";
				 elements.headers["Accept"] = "application/json;odata=verbose";
			 }
			$.ajax({
				url: elements.url,
				method: !elements.method ? "GET" : elements.method,
				headers: elements.headers,
				data: elements.body ?? elements.data as SPRESTTypes.TXmlHttpRequestData,
				success: (data: SPRESTTypes.TSPResponseData , status: string, requestObj: JQueryXHR) => {
						 if (data.d && data.d.__next && data.d.results)
							  RequestAgain(
								  elements,
									data.d.__next,
									data.d.results
							  ).then((response) => {
									elements.successCallback(response);
							  }).catch((response) => {
									elements.errorCallback(response);
							  });
						else
						 	elements.successCallback(data, status, requestObj);
				},
				error: function (requestObj: JQueryXHR, status: string, thrownErr: string) {
					elements.errorCallback(requestObj, status, thrownErr);
				}
			});
		}
	}

	static httpRequestPromise(parameters: {
		url: string;
		setDigest?: boolean;
		method?: SPRESTTypes.THttpRequestMethods;
		headers?: SPRESTTypes.THttpRequestHeaders;
		data?: SPRESTTypes.TXmlHttpRequestData;
		body?: SPRESTTypes.TXmlHttpRequestData;
	}): Promise<any> {
		return new Promise((resolve, reject) => {
			SPServerREST.httpRequest({
				setDigest: parameters.setDigest,
				url: parameters.url,
				method: parameters.method,
				headers: parameters.headers,
				data: parameters.data ?? parameters.body as string,
				successCallback: (data: SPRESTTypes.TSPResponseData, status?: string, reqObj?: JQueryXHR) => {
					resolve({data: data, message: status, reqObj: reqObj});
				},
				errorCallback: (reqObj: JQueryXHR, text?: string, errThrown?: string) => {
					reject({reqObj: reqObj, text: text, errThrown: errThrown});
				}
			});
		});
	}

	getSiteProperties (): Promise<any> {
		return SPServerREST.httpRequestPromise({
			url: this.apiPrefix + "/site"
		});
	};

	getWebProperties (): Promise<any> {
		return SPServerREST.httpRequestPromise({
			url: this.apiPrefix + "/web"
		});
	};

	getEndpoint (endpoint: string): Promise<any> {
		return SPServerREST.httpRequestPromise({
			url: this.apiPrefix + endpoint
		});
	};

	getRootweb (): Promise<any> {
		return SPServerREST.httpRequestPromise({
			url: this.apiPrefix + "/site/rootweb?$select=Id,Title,ServerRelativeUrl"
		});
	};
}

export const MAX_REQUESTS = 500;

/**
 * @function batchRequestingQueue -- when requests are too large in number (> MAX_REQUESTS), the batching
 * 		needs to be broken up in batches
 * @param {SPRESTTypes.IBatchHTTPRequestParams} elements -- same as elements in singleBatchRequest
 *    the BatchHTTPRequestParams object has following properties
 *       host: string -- required name of the server (optional to lead with "https?://")
 *       path: string -- required path to a valid SP site
 *       protocol?: string -- valid use of "http" or "https" with "://" added to it or not
 *       AllHeaders?: SPRESTTypes.THttpRequestHeaders -- object of [key:string]: T; type, headers to apply to all requests in batch
 *       AllMethods?: string -- HTTP method to apply to all requests in batch
 * @param {SPRESTTypes.IBatchHTTPRequestForm} allRequests -- the requests in singleBatchRequest
 *    the BatchHTTPRequestForm object has following properties
 *       url: string -- the valid REST URL to a SP resource
 *       method?: httpRequestMethods -- valid HTTP protocol verb in the request
 */
 function batchRequestingQueue(
	elements: IBatchHTTPRequestParams,
	allRequests: IBatchHTTPRequestForm[]
): Promise<{success: TFetchInfo[], error: TFetchInfo[]}> {
let allRequestsCopy: IBatchHTTPRequestForm[] = JSON.parse(JSON.stringify(allRequests)),
	successResponses: TFetchInfo[] = [],
	errorResponses: TFetchInfo[] = [],
	subrequests: IBatchHTTPRequestForm[] = [];

	console.log("\n\n=======================" +
	              "\nbatchRequestingQueue()" +
					  "\n=======================" +
					  "\nQueued " + allRequestsCopy.length + " requests");
	return new Promise((resolve, reject) => {
		for (let j = 0, i = 0; j < MAX_REQUESTS && i < allRequestsCopy.length; j++, i++)
			subrequests.push(allRequestsCopy[i]);
		console.log("Batch of " + subrequests.length + " requests proceeding to network");
		allRequestsCopy.splice(0, MAX_REQUESTS);
		singleBatchRequest(elements, subrequests)
			.then((response: {success: TFetchInfo[], error: TFetchInfo[]} | null) => {
			if (response != null) {
				successResponses = successResponses.concat(response.success);
				errorResponses = errorResponses.concat(response.error);
			}
			if (allRequestsCopy.length > 0)
				batchRequestingQueue(elements, allRequestsCopy);
				else
				resolve({
					success: successResponses,
					error: errorResponses
				});
		}).catch((response) => {
			errorResponses = errorResponses.concat(response);
			reject({
				success: null,
				error: errorResponses
			});
		});
	});
}

export function CreateUUID():string {
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
		 let r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);

		 return v.toString(16);
	});
}

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
function singleBatchRequest(
	elements: IBatchHTTPRequestParams,
	requests: IBatchHTTPRequestForm[]
): Promise<{success: TFetchInfo[], error: TFetchInfo[]} | null> {
	return new Promise((resolve, reject) => {
		let multipartBoundary = "batch_" + CreateUUID(),
				changeSetBoundary = "changeset_" + CreateUUID(),
				protocol = elements.protocol ?? "https://",
				requestedUrls: string[] = [],
				allHeaders = "",
				body = "",
				content = "",
				headerContent = "",
				currentMethod: THttpRequestMethods | "" = "",
				previousMethod: THttpRequestMethods | "" = "";

		if (elements.AllHeaders)
			for (let header in elements.AllHeaders)
				allHeaders += "\n" + header + ": " + elements.AllHeaders[header];
// create the body
		if (protocol.search(/\/\/$/) < 0)
			protocol += "//";
		if (elements.host.search(/http/) == 0)
			protocol = "" as SPRESTTypes.THttpRequestProtocol;
		for (let request of requests) {
			currentMethod = request.method ?? elements.AllMethods ?? "GET";
			/ method checking
			if (currentMethod == "POST" && request.url.search(/\/items\(\d{1,}\)/) >= 0) {
				resolve({
					success: [],
					error: [{
						RequestedUrl: request.url,
						Etag: null,
						ContentType: request.headers ? request.headers!["Content-Type"] as string : null,
						HttpStatus: 0,
						Data: "A URL with POST method was found with a GetById() function used. " +
							"POST methods create items, while PATCH methods are used to update items.",
						ProcessedData: []
					}]
				});
				return;
			}if (currentMethod != "GET")
				body += "\n\n--" + changeSetBoundary;
			else if (previousMethod.length > 0 && previousMethod != "GET" && currentMethod == "GET")
				body += "\n\n--" + changeSetBoundary + "--";
			if (currentMethod == "GET")
				body += "\n\n--" + multipartBoundary;

			body += "\nContent-Type: application/http";
			body += "\nContent-Transfer-Encoding: binary";

			body += "\n\n" + currentMethod + " " + request.url + " HTTP/1.1";
			requestedUrls.push(request.url);

 			// header part
			if (request.headers)
				for (let header in request.headers)
					headerContent += "\n" + header + ": " + request.headers[header];
			else
				headerContent = allHeaders;
			if (headerContent.search(/Accept:/i) < 0)
				headerContent += "\nAccept: application/json;odata=nometadata";
			if (currentMethod == "GET")
				headerContent = headerContent.replace(/Content\-Type:[^\r\n]+\r?\n?/, "");
			body += headerContent;
			if (currentMethod != "GET")
				body += "\n\n" + JSON.stringify(request.body);
			previousMethod = currentMethod;
		}
		if (currentMethod != "GET")
			body += "\n\n--" + changeSetBoundary + "--";

		// header
		content += "\n\n--" + multipartBoundary;
		content += "\nContent-Type: multipart/mixed; boundary=" + changeSetBoundary;
		content += "\nContent-Length: " + body.length;
		content += "\nContent-Transfer-Encoding: binary";
		content += "\nAccept: application/json;odata=nometadata";
		content += body;
		// footer
		content += "\n\n--" + multipartBoundary + "--";

		$.ajax({
			url: protocol + elements.host + elements.path + "/_api/contextinfo",
			method: "POST",
			headers: {
				"Content-Type": "application/json;odata=verbose",
				"Accept": "application/json;odata=verbose"
			},
			success: (data: any) => {
				console.log("Top level request header:" +
					"\nPOST " + protocol + elements.host + elements.path + "/_api/$batch");
				console.log(
					"X-RequestDigest: " + (data.FormDigestValue ? data.FormDigestValue :
							data.d.GetContextWebInformation.FormDigestValue));
				console.log(
					"Content-Type: multipart/mixed; boundary=" + multipartBoundary +
					"\nAccept: application/json;odata=verbose");
					console.log("\n\n\n************* START BODY CONTENT ***************** \n" +
					content + "\n************** END BODY CONTENT ******************");
				$.ajax({
					url: protocol + elements.host + elements.path + "/_api/$batch",
					method: "POST",
					headers: {
						// element.multipart boundary usually  "batch_" + guid()
						"X-RequestDigest":  data.FormDigestValue ? data.FormDigestValue :
										data.d.GetContextWebInformation.FormDigestValue,
						"Content-Type": "multipart/mixed; boundary=" + multipartBoundary,
						"Accept": "application/json;odata=verbose"
					},
					data: content,
					success: (data: any) => {
						splitRequests(data, requestedUrls).
							then((response: {success: TFetchInfo[], error: TFetchInfo[]} | null) => {
						}).catch((response: any) => {
							reject(response);
						});
					},
					// error completing the batch request
					error: (reqObj: object) => {
						reject(reqObj);
					}
				});
			},
			// Error getting POST token
			error: (reqObj: object) => {
				reject(reqObj);
			}
		});
	});
}

function splitRequests(
	responseSet: string,
	requestedUrls: string[]
): Promise<{success: TFetchInfo[], error: TFetchInfo[] } | null> {
	return new Promise((resolve, reject) => {
		let idx: number,
			httpCode: number,
			urlIndex: number = 0,
			checkResponse: Promise<any>[] = [],
			match: RegExpMatchArray | null,
			fetchInfo: TFetchInfo[] = [],
			errorResponses: TFetchInfo[] = [],
			allResponses: RegExpMatchArray | null = responseSet.match(/HTTP\/1.1[^\{]+(\{.*\})/g);

		if (allResponses == null)
			return resolve(null);
			for (let response of allResponses!) {
				if ((httpCode = parseInt(response.match(/HTTP\/\d\.\d (\d{3})/)![1])) < 400) {
					idx = fetchInfo.push({
						RequestedUrl: requestedUrls[urlIndex++],
						HttpStatus: httpCode,
						ContentType: (match = response.match(/CONTENT\-TYPE: (.*)\r\n(ETAG|\r\n)/)) != null ? match[1] : null,
						Etag: (match = response.match(/\r\nETAG: ("\d{3}")\r\n/)) != null ? match[1] : null,
						Data: JSON.parse(response.match(/\{.*\}/)![0]),
						ProcessedData: []
					}) - 1;
					checkResponse.push(new Promise((resolve, reject) => {
						collectNext(fetchInfo[idx].Data, []).then((fetched) => {
							resolve(fetched);
						}).catch((response) => {
							reject(response);
						});
					}));
				} else
					errorResponses.push({
						RequestedUrl: requestedUrls[urlIndex++],
						HttpStatus: httpCode,
						ContentType: (match = response.match(/CONTENT\-TYPE: (.*)\r\n(ETAG|\r\n)/)) != null ? match[1] : null,
						Etag: (match = response.match(/\r\nETAG: ("\d{3}")\r\n/)) != null ? match[1] : null,
						Data: JSON.parse(response.match(/\{.*\}/)![0]),
						ProcessedData: []
					});
		Promise.all(checkResponse).then((fetches) => {
			for (idx = 0; idx < fetches.length; idx++)
				fetchInfo[idx].ProcessedData = fetches[idx];
				resolve({
					success: fetchInfo,
					error: errorResponses
			});
		}).catch((response) => {
			reject(response);
		});
	});

	function collectNext(
		responseObj: any,
		carryData: any[]
	): Promise<any> {
		return new Promise((resolve, reject) => {
			let nextLink: string | null;

			if ((responseObj.d || Array.isArray(responseObj) == false) && !responseObj.value) {
				resolve(responseObj);
				return;
			}
			carryData = carryData.concat(responseObj.value);
			nextLink = responseObj["odata.nextLink"] ||
					(responseObj["d"] && responseObj["d"]["__next"] ? responseObj["d"]["__next"] : null) ||
					(responseObj["__next"] ? responseObj["__next"] : null);
			if (nextLink == null) {
				resolve(carryData);
				return;
			}
			$.ajax({
				url: nextLink,
				method: "GET",
				headers: {
					// element.multipart boundary usually  "batch_" + guid()
					"Content-Type": "application/json;odata=verbose",
					"Accept": "application/json;odata=nometadata"
				},
				success: (data) => {
					if (data["odata.nextLink"])
						resolve(collectNext(data, carryData));
					else
						resolve(carryData.concat(data.value));
				},
				error: (reqObj) => {
					reject({
						reqObj: reqObj,
						addlMessage: reqObj.status == 404 ? "A __next link returned 404 error" : ""
					});
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
function SPListColumnCopy(
	siteURL: string,
	listNameOrGUID: string,
	sourceColumnIntName: string,
	destColumnIntName: string // this must already exist on the SP list
): Promise<any> {
	return new Promise((resolve, reject) => {
		let guidRE = /[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}/i;

		listNameOrGUID = listNameOrGUID.search(guidRE) >= 0 ? "(guid'" + listNameOrGUID + "')" :
				"/getByTitle('" + listNameOrGUID + "')";

		RESTrequest({
			url: siteURL + "/_api/web/lists" + listNameOrGUID + "/items?$select=Id," + sourceColumnIntName,
			method: "GET",
			headers: {
				"Accept": "application/json;odata=verbose",
				"Content-Type": "application/json;odata=verbose",
			},
			successCallback: (data: any /*, text, reqObj */) => {
				let requests: SPRESTTypes.IBatchHTTPRequestForm[] = [],
					host: string = siteURL.match(/https:\/\/[^\/]+/)![0],
					responses: string = "",
					body: any;

				for (let response of data) {
					body = {};
					body.__metadata = {type: response.__metadata.type};
					body[destColumnIntName] = response[sourceColumnIntName];
					requests.push({
						url: siteURL + "/_api/web/lists" + listNameOrGUID + "/items(" + response.ID + ")",
						body: body
					});
				}

				batchRequestingQueue({
						host: host,
						path: siteURL.substring(host.length),
						AllMethods: "PATCH",
						AllHeaders: {
							"Content-Type": "application/json;odata=verbose",
							"Accept": "application/json;odata=nometadata",
							"IF-MATCH": "*",
					//		"X-HTTP-METHOD": "MERGE"
						}
					},
					requests
				).then((response) => {
					resolve(response);
				}).catch((response) => {
					reject(response);
				});
			},
			errorCallback: (reqObj: JQueryXHR/*, status, errThrown */) => {
				reject(reqObj);
			}
		});
	});
}

/**
 * @function formatRESTBody -- creates a valid JSON object as string for specifying SP list item updates/creations
 * @param {object}  JsonBody -- JS object which conforms to JSON rules. Must be "field_properties" : "field_values" format
 *                    If one of the properties is "__SetType__", it will fix the "__metadata" property
 */
export function formatRESTBody(JsonBody: { [key: string]: string | object } | string ): string {
   let testString: string, testBody: object,


	temp: any;

   try {
      testString = JSON.stringify(JsonBody);
      testBody = JSON.parse(testString);
   } catch (e) {
      throw "The argument is not a JavaScript object with quoted properties and quoted values (JSON format)";
   }
	function processObjectLevel(
		objPart: { [key: string]: string | object } | string
	): { [key: string]: string | object } | string {
		let newPart: any = {};

		if (typeof objPart == "string")
			return objPart;
      for (let property in objPart)
			if (property == "__SetType__")
				newPart["__metadata"] = { "type" : objPart[property] };
			else if (typeof objPart[property] == "object") {
				if (objPart[property] == null)
					newPart[property] = "null";
				else
					newPart[property] = processObjectLevel(objPart[property] as {[key:string]: string | object});
			} else if (objPart[property] instanceof Date == true)
				newPart[property] = (objPart[property] as Date).toISOString();
         else
				newPart[property] = objPart[property];
		return newPart;
	}
	temp = processObjectLevel(JsonBody);
	return temp;
}

export function checkEntityTypeProperty(body: object, typeCheck: string) {
	let checkRE: RegExp;

	if (typeCheck == "item")
		checkRE = SPRESTGlobals.ListItemEntityTypeRE;
	else if (typeCheck == "list")
		checkRE = SPRESTGlobals.ListEntityTypeRE;
	else if (typeCheck == "field")
		checkRE = SPRESTGlobals.ListFieldEntityTypeRE;
	else if (typeCheck == "content type")
		checkRE = SPRESTGlobals.ContentTypeEntityTypeRE;
	else if (typeCheck == "view")
		checkRE = SPRESTGlobals.ViewEntityTypeRE;
	else
		return false;
	for (let property in body)
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

export function ParseSPUrl (url: string): SPRESTTypes.TParsedURL | null {
	const urlRE = /(https?:\/\/[^\/]+)|(\/[^\/\?]+)/g;
	let index: number,
		urlParts: RegExpMatchArray | null,
		query: URLSearchParams | null = null,
		listConfirmed: boolean = false,
		siteFullPath: string = "",
		sitePartialPath: string | null = "",
		libpath: string | null = null,
		pathDone: boolean = false,
		fName: string | null = null,
		listName: string | null = null,
		tmpArr: any[] = [];

	if (!url)
		url = location.href;
	if (typeof url != "string")
		throw "parameter 'url' must be of type string and be a valid url";
	if ((urlParts = url.match(urlRE)) == null)
		return null;
	if ((index = url.lastIndexOf("?")) >= 0)
		query = new URLSearchParams(url.substring(index));
	for (let i = 1; i < urlParts.length; i++)
		if (urlParts[i] == "/Lists") {
			listName = urlParts[i + 1].substring(1);
			pathDone = true;
			listConfirmed = true;
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
	if (siteFullPath == sitePartialPath)
		sitePartialPath = null;

	if (query)
		for (let pair of query.entries())
			tmpArr.push({key: pair[0], value: pair[1]});

	//if (urlParts[3].charAt(urlParts[3].length - 1) == "/")
	//	urlParts[3] = urlParts[3].substring(0, urlParts[3].length - 1);
	return {
		originalUrl: url,
		protocol: urlParts[0].substring(0, urlParts[0].lastIndexOf("/") + 1) as SPRESTTypes.THttpRequestProtocol,
		server: urlParts[0],
		hostname: urlParts[0].substring(urlParts[0].lastIndexOf("/") + 1),
		siteFullPath: siteFullPath,
		sitePartialPath: sitePartialPath as string,
		list: listName as string,
		listConfirmed: listConfirmed,
		libRelativeUrl: libpath as string,
		file: fName as string,
		query: tmpArr
	};
}

/**
 *
 * @param {function} opFunction
 * @param {arrayOfObject} dataset -- this must be an array of objects in JSON format that
 * 				represent the body of a POST request to create SP data: fields, items, etc.
 * @returns -- the return operations are to unwind the recursion in the processing
 *           to get to either resolve or reject operations
 */
export function serialSPProcessing(opFunction: (arg1: any) => Promise<any>, itemset: any[]): Promise<any> {
	let responses: any[] = [ ];
	return new Promise((resolve) => {
		function iterate(index: number) {
			let datum: any;

			if ((datum = itemset[index]) == null)
				resolve(responses);
			else
				opFunction(datum).then((response) => {
					responses.push(response);
					iterate(index + 1);
				}).catch((response) => {
					responses.push(response);
					iterate(index + 1);
				});
		}
		iterate(0);
	});
}

export function constructQueryParameters(parameters: {[key:string]: any | any[] }): string {
	let query: string = "",
		odataFunctions = ["filter", "select", "expand", "top", "count", "skip"];

	if (!parameters)
		return "";
	if (typeof parameters == "string")  // not even an object
		return parameters;
	if (parameters.query && typeof parameters.query == "string")
		return parameters.query;
	for (let property in parameters) {
		if (odataFunctions.find(elem => elem == property) == null)
			continue;
		if (query == "")
			query = "?";
		else
			query += "&";
		if (Array.isArray(parameters[property]) == true)
			query += parameters[property].join(",");
		else if (new RegExp(property).test(parameters[property] as string) == true)
			query += parameters[property];
		else
			query += "$" + property + "=" + parameters[property];
	}
	return query;
}

/**
 * @function getCheckedInput -- returns the value of a named HTML input object representing radio choices
 * @param {HtmlInputDomObject} inputObj -- the object representing selectabe
 *     input: radio, checkbox
 * @returns {primitive data type | array | null} -- usually numeric or string representing choice from radio input object
 */
 export function getCheckedInput(inputObj: HTMLInputElement | RadioNodeList): null | string | string[] {
	if ((inputObj as RadioNodeList).length) { // multiple checkbox
		let checked: string[] = [];

		for (let cbox of inputObj as RadioNodeList)  // this is a checkbox type
			if ((cbox as HTMLInputElement).checked == true)
				checked.push((cbox as HTMLInputElement).value);
		if (checked.length > 0) {
			if (checked.length == 1)
				return checked[0];
			return checked;
		}
		return null;
	} else if ((inputObj as HTMLInputElement).type == "radio")  // just one value
		return inputObj.value;
	return null;
}

/**
* @function setCheckedInput -- will set a radio object programmatically
* @param {HtmlRadioInputDomNode} inputObj   the INPUT DOM node that represents the set of radio values
* @param {primitive value} value -- can be numeric, string or null. Using 'null' effectively unsets/clears any
*        radio selection
& @returns boolean  true if value set/utilized, false otherwise
*/
export function setCheckedInput(
	inputObj: HTMLInputElement & RadioNodeList,
	value: string | string[] | null
): boolean {
	if (inputObj.length && value != null && Array.isArray(value) == true) {  // a checked list
		if (value.length && value.length > 0) {
			for (let val of value)
				for (let cbox of inputObj)
					if ((cbox as HTMLInputElement).value == val)
						(cbox as HTMLInputElement).checked = true;
		} else
			for (let cbox of inputObj)
			if ((cbox as HTMLInputElement).value == value) {
				(cbox as HTMLInputElement).checked = true;
				return true;
			}
	} else if (value != null) {
		inputObj.value = value as string;
		return true;
	}
	return false;
}

/**
 * @function formatDateToMMDDYYYY -- returns ISO 8061 formatted date string to MM[d]DD[d]YYYY date
 *      string, where [d] is character delimiter
 * @param {string|datetimeObj} dateInput [required] -- ISO 8061-formatted date string
 * @param {string} delimiter [optional] -- character that will delimit the result string; default is '/'
 * @returns {string} -- MM[d]DD[d]YYYY-formatted string.
 */
export function formatDateToMMDDYYYY(
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


function fixValueAsDate(date: Date): Date | null {
	let datestring: RegExpMatchArray | null = date.toISOString().match(/(\d{4})\-(\d{2})\-(\d{2})/);

	if (datestring == null)
		return null;
	return new Date(parseInt(datestring[1]), parseInt(datestring[2]) - 1, parseInt(datestring[3]));
}

// ref: http://stackoverflow.com/a/1293163/2343
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.
export function CSVToArray(
		strData: string,
		strDelimiter: string = "," // default delimiter is ','
): string[][] {
    // Create a regular expression to parse the CSV values.
    let objPattern = new RegExp(
        (
            // Delimiters.
            "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

            // Quoted fields.
            "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

            // Standard fields.
            "([^\"\\" + strDelimiter + "\\r\\n]*))"
        ),
        "gi"
        );

    // Create an array to hold our data. Give the array
    // a default empty first row.
    let arrData: Array<Array<string>> = [[]];

    // Create an array to hold our individual pattern
    // matching groups.
    let arrMatches: RegExpExecArray | null = null;

    // Keep looping over the regular expression matches
    // until we can no longer find a match.
    while (arrMatches = objPattern.exec( strData )){

        // Get the delimiter that was found.
        let strMatchedDelimiter = arrMatches[ 1 ];

        // Check to see if the given delimiter has a length
        // (is not the start of string) and if it matches
        // field delimiter. If id does not, then we know
        // that this delimiter is a row delimiter.
        if (
            strMatchedDelimiter.length &&
            strMatchedDelimiter !== strDelimiter
            ){

            // Since we have reached a new row of data,
            // add an empty row to our data array.
            arrData.push( [] );
        }

        let strMatchedValue;
        // Now that we have our delimiter out of the way,
        // let's check to see which kind of value we
        // captured (quoted or unquoted).
        if (arrMatches[ 2 ]){

            // We found a quoted value. When we capture
            // this value, unescape any double quotes.
            strMatchedValue = arrMatches[ 2 ].replace(
                new RegExp( "\"\"", "g" ),
                "\""
                );
        } else {

            // We found a non-quoted value.
            strMatchedValue = arrMatches[ 3 ];

        }
        // Now that we have our value string, let's add
        // it to the data array.
        arrData[ arrData.length - 1 ].push( strMatchedValue );
    }

    // Return the parsed data.
    return arrData;
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
 export function buildSelectSet(
	parentNode: HTMLElement,   // DOM node to append the construct
	nameOrGuid: string,   // used in tagging
	options: {
		selected?:boolean,
		required?:boolean,
		value:any,
		text:string
	}[],      // array of option objects {text: text-to-display, value: node value to set}
	availableSize: number, // for number of choices displayed for available choices select list
	selectedSize: number, // for number of choices displayed for selected choices select list
	onChangeHandler: (arg?: any) => {} | void // callback
): string | void {
	let selectNode: HTMLSelectElement,
		oNode: HTMLOptionElement,
		pNode: HTMLParagraphElement,
		sNode: HTMLSpanElement,
		bNode: HTMLButtonElement,
		guid: string,
		divNode: HTMLDivElement = document.createElement("div");

	parentNode.appendChild(divNode);
	divNode.style.display = "grid";
	divNode.style.gridTemplateColumns = "auto auto auto";
	divNode.className = "select-set";
	guid = createGuid();
	divNode.id = guid;

	pNode = document.createElement("p");
	divNode.appendChild(pNode);
	sNode = document.createElement("span");
	pNode.appendChild(sNode);
	sNode.appendChild(document.createTextNode("Available"));
	sNode.style.font = "italic 9pt Arial,sans-serif";
	sNode.style.display = "block";
	selectNode = document.createElement("select");
	pNode.appendChild(selectNode);
	selectNode.name = nameOrGuid + "-avail";
	selectNode.multiple = true;
	selectNode.size = availableSize;
	selectNode.style.minWidth = "10em";
	selectNode.addEventListener("change", onChangeHandler);
	for (let option of options)
		if (!option.selected)	{
			oNode = document.createElement("option");
			selectNode.appendChild(oNode);
			oNode.value = option.value;
			if (option.required == true)
				oNode.className = "required";
			oNode.appendChild(document.createTextNode(option.text));
		}
	//	selectNode["data-options"] = JSON.stringify(options);

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
		let option: HTMLOptionElement,
			form = (event.currentTarget as HTMLInputElement).form,
			availOptions: HTMLOptionsCollection = form![nameOrGuid + "-avail"].selectedOptions,
			chosenSelect: HTMLSelectElement = form![nameOrGuid + "-selected"];

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
		let option: HTMLElement,
			form = (event.currentTarget as HTMLFormElement).form,
			availSelect: HTMLSelectElement = form[nameOrGuid + "-avail"],
			chosenOptions: HTMLOptionsCollection = form[nameOrGuid + "-selected"].selectedOptions;

		for (let i = 0; i < chosenOptions.length; i++) {
			option = chosenOptions[i];
			availSelect.appendChild(option.parentNode!.removeChild(option));
		}
		sortSelect(availSelect);
	});

	pNode = document.createElement("p");
	divNode.appendChild(pNode);
	sNode = document.createElement("span");
	pNode.appendChild(sNode);
	sNode.appendChild(document.createTextNode("Selected"));
	sNode.style.font = "italic 9pt Arial,sans-serif";
	sNode.style.display = "block";
	selectNode = document.createElement("select");
	pNode.appendChild(selectNode);
	selectNode.name = nameOrGuid + "-selected";
	selectNode.multiple = true;
	selectNode.style.minWidth = "10em";
	selectNode.size = selectedSize;
	selectNode.addEventListener("change", onChangeHandler);
	for (let option of options)
		if (option.selected && option.selected == true)	{
			oNode = document.createElement("option");
			selectNode.appendChild(oNode);
			oNode.value = option.value;
			if (option.required == true)
				oNode.className = "required";
			oNode.appendChild(document.createTextNode(option.text));
		}
	sortSelect(selectNode);
	return guid;


	//	selectNode["data-options"] = JSON.stringify([]);
	function sortSelect(selectObj: HTMLSelectElement) {
		let node: HTMLOptionElement,
			selectArray: {
				display: string;
				value: string;
				required: boolean;
			}[] = [ ],
			opshun: HTMLOptionElement;

		for (opshun of selectObj.options)
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
		for (let elem of selectArray) {
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
	let capture: boolean = false,
		includeText: boolean = false,
		useText: string[] = ["select all", "unselect all"],
		buttonImageWidth: string = "15px",
		includeTextFontSize: string = "8pt",
		includeTextFontColor: string = "brown",
		alignment: "right" | "center" | "left" = "left",
		control: RadioNodeList,
		setStyle: string | null = null,
		containerElemType: string = parameters.containerElemType ?? "p",
		form: HTMLFormElement,
		parent: HTMLElement | null;

	if (typeof parameters.form == "string")
		 form = document.getElementById(parameters.form) as HTMLFormElement;
	else
		 form = parameters.form;
	if (!(form && form.nodeName && form.nodeName.toLowerCase() == "form"))
		 throw "A form DOM node or its 'id' attribute must be a parameter and contain the checkbox elements";

		 // look for a form that contains this
	if (!parameters.formName || typeof parameters.formName != "string" || parameters.formName.length == 0)
		 throw "The parameter 'formName' is either undefined or not a string of non-zero length";
	if ((control = form[parameters.formName]) == null)
		 throw "No 'form' with that name attribute is found on the document. It must be present in HTML document.";

	if (typeof parameters.containerNode == "string") { // is an id attribute
		if ((parent = document.getElementById(parameters.containerNode)) == null)
			throw "A container node ID but did not correspond to an existing HTML element";
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
		if (setStyle && setStyle.search(/text\-align/) < 0)
			parent.style.textAlign = alignment;
		if (setStyle && setStyle.search(/font\-size/) < 0)
			parent.style.fontSize = includeTextFontSize;
		if (setStyle && setStyle.search(/[^\-]color/) < 0)
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
		let sNode: HTMLSpanElement,
			iNode: HTMLImageElement,
			imgSrc: string;

		sNode = document.createElement("span");
		container.appendChild(sNode);

		iNode = document.createElement("img");
		if (typeof includedText == "string" && includedText.length > 0) {
			sNode.appendChild(document.createTextNode(includedText));
			iNode.style.marginLeft = "0.3em";
		}
		sNode.appendChild(iNode);
		iNode.addEventListener("click", () => {
			for (let checkbox of inputNodeList)
				(checkbox as HTMLInputElement).checked = checkSet;
			if (clickCallback)
				clickCallback(callbackArg);
		}, capture);
		imgSrc = usedImageUrl;
		if (!imgSrc || imgSrc == null) {
			 iNode.style.width = usedImageWidth
			 imgSrc = UnselectAllCheckboxes;
		}
		iNode.src = imgSrc;
		iNode.alt = altText;
	}
}


/**
 * @function isIterable -- tests whether variable is iterable
 * @param {object} obj -- basically any variable that may or may not be iterable
 * @returns boolean - true if iterable, false if not
 */
 export function isIterable(obj: any) {
	// checks for null and undefined
	if (obj == null)
		 return false;
	return typeof obj[Symbol.iterator] === 'function';
 }

// off the Internet
export function createGuid(){
	function S4() {
		return (1 + Math.random() * 0x10000 | 0).toString(16).substring(1);
	}
	return (S4() + S4() + "-" + S4() + "-4" + S4().substring(0,3) + "-" + S4() + "-" + S4() + S4() + S4()).toLowerCase();
}

export function createFileDownload(parameters: {
	href: string;
	downloadFileName: string;
	newTab?: boolean;
}): void {
	let aNode = document.createElement("a");

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
    In the document, create two undisplayed DIV elements
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
export function openFileUpload(
	callback: (fileList: FileList) => void
): void {
	let containerDiv: HTMLDivElement = document.createElement("div"),
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

function disabledClick(evt: Event) {
	let img: HTMLImageElement;

	if ((img = evt.target as HTMLImageElement).id == "cancelingUpload") {
		document.getElementById("drop-container")!.style.display = "none";
		document.body.removeEventListener("click", disabledClick, true);
	}
	evt.stopImmediatePropagation();
	evt.preventDefault();
}

/**
* @function RESTrequest -- REST requests interface optimized for SharePoint server
* @param {SPRESTTypes.THttpRequestParams} elements -- all the object properties necessary for the jQuery library AJAX XML-HTTP Request call
*     properties are:
*        url: string;
*        setDigest?: boolean;
*        method?: THttpMethods;
*        headers?: SPRESTTypes.THttpRequestHeaders;
*        data?: string | ArrayBuffer | null;
*        body?:  same as data
*        successCallback: TSuccessCallback;
*        errorCallback: TErrorCallback;
* @returns {object} all the data or error information via callbacks
*/

export function RESTrequest(elements: SPRESTTypes.THttpRequestParams): void {
	if (elements.setDigest && elements.setDigest == true) {
 		let match: RegExpMatchArray = elements.url.match(/(.*\/_api)/) as RegExpMatchArray;

		 $.ajax({  // get digest token
	 		url: match[1] + "/contextinfo",
	 		method: "POST",
	 		headers: {...SPstdHeaders},
	 		success: function (data) {
			  let headers: SPRESTTypes.THttpRequestHeaders = elements.headers as SPRESTTypes.THttpRequestHeaders;

				if (headers) {
					headers["Content-Type"] = "application/json;odata=verbose";
					headers["Accept"] = "application/json;odata=verbose";
				} else
					headers = {...SPstdHeaders};
		 		headers["X-RequestDigest"] = data.FormDigestValue ? data.FormDigestValue :
			 			data.d.GetContextWebInformation.FormDigestValue;
				$.ajax({
					url: elements.url,
					method: elements.method ? elements.method : "GET",
					headers: headers,
					data: elements.body || elements.data as string,
					success: function (data: SPRESTTypes.TSPResponseData, status: string, requestObj: JQueryXHR) {
						elements.successCallback!(data, status, requestObj);
					},
					error: function (requestObj: JQueryXHR, status: string, thrownErr: string) {
						elements.errorCallback!(requestObj, status, thrownErr);
					}
				});
	 		},
	 		error: function (requestObj, status, thrownErr) {
				elements.errorCallback!(requestObj, status, thrownErr);
	 		}
		});
	} else {
		if (!elements.headers)
	  		elements.headers = {...SPstdHeaders};
		else {
			elements.headers["Content-Type"] = "application/json;odata=verbose";
			elements.headers["Accept"] = "application/json;odata=verbose";
		}
		$.ajax({
			url: elements.url,
			method: elements.method,
			headers: elements.headers,
			data: elements.body || elements.data as string,
			success: function (data: SPRESTTypes.TSPResponseData, status: string, requestObj: JQueryXHR) {
				if (data.d && data.d.__next)
					RequestAgain(
						elements,
						data.d.__next,
						data.d.results!
					).then((response) => {
						elements.successCallback!(response);
					}).catch((response) => {
						elements.errorCallback!(response);
					});
				else
					elements.successCallback!(data, status, requestObj);
			},
			error: function (requestObj, status, thrownErr) {
				elements.errorCallback!(requestObj, status, thrownErr);
			}
		});
	}
}

export function RequestAgain(
		elements: SPRESTTypes.THttpRequestParams,
		nextUrl: string,
		aggregateData: SPRESTTypes.TSPResponseDataProperties[]
): Promise<any> {
	return new Promise((resolve, reject) => {
		$.ajax({
			url: nextUrl,
			method: "GET",
			headers: elements.headers,
			success: (data) => {
				if (data.d.__next) {
					aggregateData = aggregateData.concat(data.d.results)
					RequestAgain(
						elements,
						data.d.__next,
						aggregateData
					).then((response) => {
						resolve(response);
					}).catch((response) => {
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
   },
   SPFieldTypes = {
		enums: [
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
			{ name: "Currency",     typeId: 10 },
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
			{ name: "Threading",    typeId: 13 },
			{ name: "Guid",         metadataType: "SP.FieldGuid",        typeId: 14 }, // so far, just a standard basic field as string
			{ name: "MultiChoice",  metadataType: "SP.FieldMultiChoice", typeId: 15,
					extraProperties: [
						"FillInChoice",     // boolean: gets/sets whether fill-in value allowed
						"Mappings",  // string:
						"Choices"       //  objects with "results" property that is array of strings with choices
					]
			},
			{ name: "GridChoice",   typeId: 16 },
			{ name: "Calculated",   typeId: 17 },
			{ name: "File",         metadataType: "SP.Field",          typeId: 18 },
			{ name: "Attachments",  typeId: 19 },
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
			{ name: "Recurrence",   typeId: 21 },
			{ name: "CrossProjectLink", typeId: 22 },
			{ name: "ModStat",      typeId: 23, // Moderation Status is a Choices type
				extraProperties: [
					"FillInChoice",     // boolean: gets/sets whether fill-in value allowed
					"Mappings",  // string:
					"Choices",       //  objects with "results" property that is array of strings with choices
							// usually "results": ["0;#Approved", "1;#Rejected", "2;#Pending", "3;#Draft", "4;#Scheduled"]
					"EditFormat"       // integer value for format type in editing
				]
			},
			{ name: "Error",        typeId: 24 },
			{ name: "ContentTypeId", metadataType: "SP.Field",   typeId: 25 },
			{ name: "PageSeparator", typeId: 26 },
			{ name: "ThreadIndex",  typeId: 27 },
			{ name: "WorkflowStatus", typeId: 28 },
			{ name: "AllDayEvent",  typeId: 29 },
			{ name: "WorkflowEventType", typeId: 30 },
			{ name: "Geolocation"   },
			{ name: "OutcomeChoice" },
			{ name: "MaxItems",     typeId: 31 }
		],
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
	};

/**
 * @function getTaxonomyValue -- returns values from single-valued Managed Metadata fields
 * @param {} obj
 * @param {*} fieldName
 * @param {*} returnValue --
 * @returns
 */
function getTaxonomyValue(obj: {[key:string]:any}, fieldName: string, returnValue: string): string {
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
	for (let field in obj) {
		// If it's the field we're interested in....
		if (obj.hasOwnProperty(field) && field === fieldName) {
			if (obj[field] !== null) {
				// ... get the WssId from the field ...
				let thisId = obj[field].WssId;
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
}

/**
 * @function parseManagedMetadata -- this function returns the text label of the metadata term store for
 *     updating a Managed Metadata field.
 * @param {} results -- should be results of a REST call
 * @param {*} fieldName -- name of the metadata field, e.g. "MetaMultiField"
 * @returns
 */

export function dedupJSArray(array: string[] | any[]): any[] {
	let newArray: any[] | string[] = [];

	for (let elem of array)
		newArray.push(JSON.stringify(elem));
	newArray = [...new Set(newArray)];
	for (let i = 0; i < newArray.length; i++)
		newArray[i] = JSON.parse(newArray[i]);
	return newArray;
}


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

SelectAllCheckboxes =
    "data:image/png;base64,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";

UnselectAllCheckboxes =
    "data:image/png;base64,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";
