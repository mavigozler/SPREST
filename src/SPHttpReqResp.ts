/* eslint-disable guard-for-in */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable eqeqeq */
import type { THttpRequestHeaders, TSPResponseData, THttpRequestParams, TSPResponseDataProperties,
	TBatchResponse, TBatchResponseRaw, IBatchHTTPRequestForm, TXmlHttpRequestData,
	THttpRequestMethods, HttpStatusCode, TFetchInfo, TSuccessCallback, TErrorCallback,
	FetchApiResponseInfo, HttpInfo, THttpResponseHeaders, TOriginType
} from "./SPHttp";
import { CreateUUID } from "@GenLib/etc";
import "core-js/features/object/from-entries";
// import { CreateUUID } from "../GenLib/etc.js";

export { SPstdHeaders, RESTrequest, RequestAgain, batchRequestingQueue };

const SPstdHeaders: THttpRequestHeaders = {
	"Content-Type":"application/json;odata=verbose",
	"Accept":"application/json;odata=verbose"
};

function RESTrequest(param: {
	url: string,
	successCallback: TSuccessCallback,
	setDigest?: boolean,
	method?: THttpRequestMethods,
	headers?: THttpRequestHeaders,
	body?: TXmlHttpRequestData,
	errorCallback?: TErrorCallback,
	useCORS?: boolean
}): void {
	const editedUrl: string = (param.url.indexOf("http") == 0 ? param.url : "https://" + param.url);

	/*
	let winFetch: typeof fetch,
		iframeWin: HTMLIFrameElement;

	if (settings.useCORS == true) { // set up temporary iframe, use then delete
		iframeWin = document.createElement("iframe");
		document.body.insertBefore(iframeWin, document.body.lastChild);
		winFetch = iframeWin.contentWindow!.fetch!;
	} else
		winFetch = fetch;
*/

	param.headers = typeof param.headers !== "undefined" ? param.headers : SPstdHeaders;
	const basicRequestSettings: THttpRequestParams = {
		url: param.url,
		setDigest: param.setDigest,
		method: param.method,
		headers: param.headers,
		useCORS: param.useCORS,
		body: param.body
	} as THttpRequestParams;
	if (param.setDigest == true) {
		const match: RegExpMatchArray = editedUrl.match(/(.*\/_api)/) as RegExpMatchArray;
		const contextinfoRequestSettings: RequestInit = {
			method: "POST",
			headers: param.headers as HeadersInit,
			// credentials: "include", "same-origin", "omit"
			// "cache": // 6 options here
			// "redirect": // 3 options
			// referrerPolicy: 8 options
			// body:  usually JSON.stringify(data)
			mode: param.useCORS == true ? "no-cors" : "same-origin"
		};
		fetch(match[1] + "/contextinfo", contextinfoRequestSettings)
			.then((data: Response) => data.json()
				.then(json => {
					const spResponse: TSPResponseData = json;
				param.headers!["X-RequestDigest"] = spResponse.FormDigestValue ? spResponse.FormDigestValue :
					spResponse.d?.GetContextWebInformation?.FormDigestValue;
				const requestSettings: RequestInit = {
					method: param.method ? param.method : "GET",
					headers: param.headers as HeadersInit,
					body: (param.body || data) as string,
					mode: param.useCORS == true ? "no-cors" : "same-origin"
				};
				fetch(editedUrl, requestSettings)
					.then((data: Response) => data.json()
						.then(json => {
							// determine if there is a "__next" property and continue Request Again
							const responseHeaders: THttpResponseHeaders = Object.fromEntries(data.headers.entries());
							const responseParams: FetchApiResponseInfo = {
								bodyUsed: data.bodyUsed,
								ok: data.ok,
								redirected: data.redirected,
								status: data.status,
								statusText: data.statusText,
								type: data.type as TOriginType
							};
							const httpInfo: HttpInfo = {
								...responseParams,
								...basicRequestSettings,
								...responseHeaders,
							} as HttpInfo;
							if (spResponse.d?.__next) // Request again handles until done
								RequestAgain(
									basicRequestSettings,
									spResponse.d.__next,
							spResponse.d.results! as TSPResponseDataProperties[]
								).then((response: unknown[]) => {
									return param.successCallback(
								response as unknown as TSPResponseData,
								httpInfo
									);
								}).catch((errInfo: unknown) => {
							param.errorCallback!(errInfo, httpInfo);
								});
					param.successCallback!(json, httpInfo);
						}).catch((err: unknown) => {
							if (param.errorCallback)
								param.errorCallback(err, requestSettings as unknown as HttpInfo);
						})
					).catch((err: unknown) => {
						if (param.errorCallback)
							param.errorCallback(err, requestSettings as unknown as HttpInfo);
					});
				//	iframeWin ? iframeWin.parentNode?.removeChild(iframeWin) : null;
				})).catch((err: unknown) => {
				if (param.errorCallback)
					param.errorCallback!(err, contextinfoRequestSettings as unknown as HttpInfo);
			});
	} else {
		const headers = param.headers || {...SPstdHeaders};
		const body = (param.method == "GET" || param.method == "HEAD") ?
				null : param.body as BodyInit,
			fetchSettings: RequestInit = {
				method: param.method,
				headers: headers as HeadersInit,
				body: body,
				mode: basicRequestSettings.useCORS == true ? "no-cors" : "same-origin"
			};
		fetch(editedUrl, fetchSettings)
			.then((data: Response) => {
				data.text()
					.then((content) => {
						const responseHeaders: FetchApiResponseInfo =
						Object.fromEntries(data.headers.entries()) as unknown as FetchApiResponseInfo;
						const responseParams: FetchApiResponseInfo = {
							bodyUsed: data.bodyUsed,
							ok: data.ok,
							redirected: data.redirected,
							status: data.status,
							statusText: data.statusText,
							type: data.type as TOriginType
						};
						const httpInfo: HttpInfo = {
							...responseParams,
							...basicRequestSettings,
							...responseHeaders,
						} as HttpInfo;

						try {
							const json = JSON.parse(content);
							if (json.d?.__next)
								RequestAgain(
									basicRequestSettings,
									json.d.__next,
							json.d.results! as TSPResponseDataProperties[]
								).then((response: unknown[]) => {
									return param.successCallback(
								response as unknown as TSPResponseData,
								httpInfo
									);
								}).catch((errInfo: unknown) => {
							param.errorCallback!(errInfo, httpInfo);
								});
							else
						param.successCallback!(json, httpInfo);
						} catch {
					param.successCallback!(content as TSPResponseData, httpInfo);
						}
					}).catch((errInfo) => {
				param.errorCallback!(errInfo, basicRequestSettings as HttpInfo);
					});
			}).catch((errInfo) => {
			// the fetch request had an internal error
			param.errorCallback!(errInfo, basicRequestSettings as HttpInfo);
			});
	}
}

// public
function RequestAgain(
	settings: THttpRequestParams,
	nextUrl: string,
	aggregateData: TSPResponseDataProperties[]
): Promise<TSPResponseDataProperties[]> {
	return new Promise((resolve, reject) => {
		fetch(nextUrl, {
			method: "GET",
			headers: settings.headers,
		} as RequestInit) .then((data: Response) => {
			data.json().then((json) => {
				const spResponse: TSPResponseData = json;
				if (spResponse.d?.__next) {
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					aggregateData = aggregateData.concat(spResponse.d.results as any[])
					RequestAgain(
						settings,
						spResponse.d.__next,
						aggregateData
					).then((response: TSPResponseDataProperties[]) => {
						resolve(response);
					}).catch((err: unknown) => {
						reject(err);
					});
				} else
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					resolve(aggregateData.concat(spResponse.d?.results as any[]));
			}).catch((err: unknown) => {
				reject(err);
			});
		}).catch((err: unknown) => {
			reject(err);
		});
	});
}


const MAX_REQUESTS = 75;

/**
 * @function batchRequestingQueue -- when requests are too large in number (> MAX_REQUESTS), the batching
 * 		needs to be broken up in batches
 * @param {IBatchHTTPRequestParams} settings -- same as settings in singleBatchRequest
 *    the BatchHTTPRequestParams object has following properties
 *       host: string -- required name of the server (optional to lead with "https?://")
 *       path: string -- required path to a valid SP site
 *       protocol?: string -- valid use of "http" or "https" with "://" added to it or not
 *       AllHeaders?: THttpRequestHeaders -- object of [key:string]: T; type, headers to apply to all requests in batch
 *       AllMethods?: string -- HTTP method to apply to all requests in batch
 * @param {IBatchHTTPRequestForm[]} allRequests -- the requests in singleBatchRequest
 *    the BatchHTTPRequestForm object has following properties
 *       url: string -- the valid REST URL to a SP resource
 *       method?: httpRequestMethods -- valid HTTP protocol verb in the request
 */
// public
function batchRequestingQueue(
	allRequests: IBatchHTTPRequestForm[]
): Promise<TBatchResponse | undefined> {
	const allRequestsCopy: IBatchHTTPRequestForm[] = JSON.parse(JSON.stringify(allRequests)),
		batches: IBatchHTTPRequestForm[][] = [],
		batchRequests: Promise<TBatchResponse | undefined>[] = [];
	let subbatches: IBatchHTTPRequestForm[][];

	//	console.log("\n\n=======================" +
	//	              "\nbatchRequestingQueue()" +
	//					  "\n=======================" +
	//					  "\nQueued " + allRequestsCopy.length + " requests");
	return new Promise<TBatchResponse | undefined>((resolve, reject) => {

		/* CONDITIONS:  Batch must have one protocol/site URL. Maximum of 100 requests in a batch */

		//  sort requests by contextinfo
		allRequestsCopy.sort((req1, req2) => {
			return req1.contextinfo > req2.contextinfo ? 1 : req1.contextinfo < req2.contextinfo ? -1 : 0;
		});
		batches[0] = [];
		for (let j = 0, i = 0; i < allRequestsCopy.length; i++) {
			if (i > 0 && (allRequestsCopy[i - 1].contextinfo !== allRequestsCopy[i].contextinfo)) {
				batches[j++] = [];
			}
			batches[j].push(allRequestsCopy[i]);
		}
		for (const batch of batches) {
			subbatches = [];
			while (batch.length > 0)
				subbatches.push(batch.splice(0, MAX_REQUESTS));
			//	console.log("Batch of " + batches.length + " requests proceeding to network");
			//	allRequestsCopy.splice(0, MAX_REQUESTS);
			for (let i = 0; i < subbatches.length; i++) {
				batchRequests.push(new Promise<TBatchResponse | undefined>((resolve) => {
					singleBatchRequest(subbatches[i])
						.then((response: TBatchResponse | undefined) => {
							//	console.log("Storing success to Promise.all() in batchRequestingQueue():" +
							//	`\n  Successes: ${response ? response.success.length : 'null'}, Errors: ${response ? response.error.length: 'null'}`);
							resolve(response);
						}).catch((response: HttpInfo) => {
							//	console.log("Storing error to Promise.all() in batchRequestingQueue():" +
							//	`\n  caught error but resolving: ${JSON.stringify(response, null, '  ')}`);
							reject(response);
						});
				}));
			}
			Promise.all(batchRequests).then((responses: (TBatchResponse | undefined)[]) => {
				let successes: TFetchInfo[] = [],
					errors: TFetchInfo[] = [];
				if (responses) {
					for (const response of responses)
						if (response) {
							successes = successes.concat(response.success);
							errors = errors.concat(response.error);
						}
					//	console.log("Promise.all() in batchRequestingQueue():" +
					//			`\n  Successes: ${successes.length}, Errors: ${errors.length}`);
					resolve({success: successes, error: errors});
				} else {
					//	console.log("Promise.all() in batchRequestingQueue():" +
					//					`\n   responses were null ')}`);
					resolve(undefined);
				}
			}).catch((err: unknown) => {
				reject(err);
			});
		}
	});
}


/*
	Format for showing batch requesting shown in comments at bottom of this
	source file!
*/

/**
 * @function singleBatchRequest -- used to exploit the $batch OData multiple write request
 * @param {object} settings -- properties that must be defined are:
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
	requests: IBatchHTTPRequestForm[]
): Promise<TBatchResponse | undefined> {
	const requestedUrls: string[] = [],
		multipartBoundary = "batch_" + CreateUUID(),
		changeSetBoundary = "changeset_" + CreateUUID(),
		contentPackage: {
			content: string;
			RequestDigest: string;
		} = {
			content: "",
			RequestDigest: ""
		};

	/****************************************************
 * A single batch must have ONE PROTOCOL and ONE SITE URL or it does not get past contextinfo
 */
	return new Promise<TBatchResponse | undefined>((resolve, reject) => {
		const headerAccept: string = "application/json;odata=nometadata",
			contextinfo = requests[0].contextinfo;
		let retval: {action: "resolve", value: object} | null;

		// check contextinfo

		if (requests.every((request) => {
			return request.contextinfo == contextinfo;
		}) == false)
			throw Error("singleBatchRequest():  all contextinfo values in a single batch must be identical");

		if (requests.every((request) => {
			return typeof request.method === "undefined" || request.method == "GET";
		}) == true)
			onlyGETRequests();
		else if ((retval = changeSetRequests()) != null)
			if (retval.action == "resolve")
				resolve(retval.value as {success: TFetchInfo[], error: TFetchInfo[]} | undefined);

		fetch(contextinfo + "/_api/contextinfo", {  //  gets the digest token
			method: "POST",
			headers: {
				"Content-Type": "application/json;odata=verbose",
				"Accept": headerAccept!
			}}).then((response: Response) => response.json()
			.then((json) => {
				const spReponse: TSPResponseData = json;
				contentPackage.RequestDigest = spReponse.FormDigestValue ? spReponse.FormDigestValue :
				spReponse.d!.GetContextWebInformation!.FormDigestValue;
				//				console.log("\n\n\n************* START BODY CONTENT *****************\n" +
				//					contentPackage.content + "\n\n\n************** END BODY CONTENT ******************");
				fetch(contextinfo + "/_api/$batch", {
					method: "POST",
					headers: {
						"X-RequestDigest":  contentPackage.RequestDigest,
						"Content-Type": "multipart/mixed; boundary=" + multipartBoundary,
						"Accept": "application/json;odata=verbose"
					},
					body: contentPackage.content,
				}).then((response: Response) => response.json()
					.then((json) => {
						const batchData: TBatchResponseRaw = {
							rawData: json,
							ETag: response.headers.get("ETag"),
							RequestDigest: response.headers.get("X-RequestDigest")
						};
						processRawResponses(batchData, requestedUrls)
							.then((response: TBatchResponse | undefined) => {
								resolve(response);
							}).catch((err: unknown) => {
								reject(err);
							});
					}).catch((err: unknown) => {
						reject(err);
				})).catch((err: unknown) => {reject(err)});
			})).catch((err: unknown) => {reject(err)});
	});

	function onlyGETRequests(): void {
		let content: string = "";

		for (const request of requests) {
			content += "\n\n--" + multipartBoundary;
			content += "\nContent-Type: application/http";
			content += "\nContent-Transfer-Encoding: binary";
			content += `\n\nGET ${request.url} HTTP/1.1`;
			content += "\nAccept: application/json;odata=verbose";
			requestedUrls.push(request.url);
		}
		content += `\n\n--${multipartBoundary}--`;
		contentPackage.content = content;
	}

	function changeSetRequests(): null | {action: "resolve", value: object} {
		let headerContent = "",
			headerAccept: string,
			content = "",
			body = "",
			currentMethod: THttpRequestMethods | "" = "",
			previousMethod: THttpRequestMethods | "" = "",
			usedChangeSetBoundary: boolean = false;

		for (const request of requests) {
			currentMethod = request.method ?? "GET";
			// method checking
			if (currentMethod == "POST" && request.url.search(/\/items\(\d{1,}\)/) >= 0)
				return {
					action: "resolve",
					value: ({
						success: [],
						error: [{
							RequestedUrl: request.url,
							Etag: null,
							ContentType: request.headers ? request.headers!["Content-Type"] as string : null,
							HttpStatus: 0,
							Data: {
								error:{
									message: {
										value: "A URL with POST method was found with a GetById() function used. " +
										"POST methods create items, while PATCH methods are used to update items."
									}
								}
							},
							ProcessedData: undefined
						}]
					})
				};
			// change of method ==> changeset
			if (currentMethod != "GET") {
				body += "\n\n--" + changeSetBoundary;
				usedChangeSetBoundary = true;
			} else if (previousMethod.length > 0 && previousMethod != "GET" && currentMethod == "GET") {
				body += "\n\n--" + changeSetBoundary + "--";
				usedChangeSetBoundary = true;
			}
			if (currentMethod == "GET")
				body += "\n\n--" + multipartBoundary;
			body += "\nContent-Type: application/http";
			body += "\nContent-Transfer-Encoding: binary";

			body += "\n\n" + currentMethod + " " + request.url + " HTTP/1.1";
			requestedUrls.push(request.url);

			// header part
			if (request.headers)
				for (const header in request.headers)
					// eslint-disable-next-line @typescript-eslint/no-explicit-any
					headerContent += "\n" + header + ": " + (request.headers as any)[header];
			//		else
			//			headerContent = allHeaders;
			if (headerContent.search(/Accept:/i) < 0) {
				headerAccept = "application/json;odata=nometadata";
				headerContent = "\nAccept: " + headerAccept;
			}
			if (currentMethod == "GET")
				headerContent = headerContent.replace(/Content-Type:[^\r\n]+\r?\n?/, "");
			body += headerContent;
			// added to body, the header content for request

			// this adds the actual body part for a non-GET request
			if (currentMethod != "GET")
				body += "\n\n" + JSON.stringify(request.body);
			previousMethod = currentMethod;
		}
		if (currentMethod != "GET") {
			body += "\n\n--" + changeSetBoundary + "--";
			usedChangeSetBoundary = true;
		}
		// header + body
		content += "\n\n--" + multipartBoundary;
		if (usedChangeSetBoundary == true)
			content += "\nContent-Type: multipart/mixed; boundary=" + changeSetBoundary;
		else
			//content += "\nContent-Type: multipart/mixed; boundary=" + multipartBoundary;
			content += "\nContent-Type: application/http"

		content += "\nContent-Transfer-Encoding: binary";
		content += "\nAccept: " + headerAccept!;
		content += body;
		// footer
		content += "\n--" + multipartBoundary + "--";
		contentPackage.content = content;
		return null;
	}
}

function processRawResponses(
	responseSet: TBatchResponseRaw,
	requestedUrls: string[]
): Promise<TBatchResponse | undefined> {
	return new Promise((resolve, reject) => {
		const successResponses: TFetchInfo[] = [],
			errorResponses: TFetchInfo[] = [],
			allResponses: RegExpMatchArray | null = responseSet.rawData.match(/HTTP\/1.1[^{]+(\{.*\})/g),
			RequestDigest: string | null = (responseSet && responseSet.RequestDigest) ?
				responseSet.RequestDigest : null;
		let idx: number,
			responseIndex: number = 0,
			httpCode: HttpStatusCode,
			urlIndex: number = 0,
			match: RegExpMatchArray | null;

		//		parsedData: TSPResponseData;


		//console.log("ENTRY_PAGE REST API Batching: splitRequests()");
		if (allResponses == null) {
			//console.log("  splitRequests() allResponses == null, returning" +
			//			   `\n     tracking: RequestDigest=${responseSet.RequestDigest}, ETag=${responseSet.ETag}`);
			return resolve(undefined);
		}
		//console.log(`  splitRequests() allResponses: ${allResponses.length} total, against ${requestedUrls.length} requests` +
		//"\n\nordered responses:");

		//for (let response of allResponses)
		//console.log(`  response: ${response}`);
		for (const response of allResponses!)
			if ((httpCode = <HttpStatusCode>parseInt(response.match(/HTTP\/\d\.\d (\d{3})/)![1])) < 400)
				idx = successResponses.push({
					ResponseIndex: responseIndex++,
					RequestedUrl: requestedUrls[urlIndex++],
					HttpStatus: httpCode,
					ContentType: (match = response.match(/CONTENT-TYPE: (.*)\r\n(ETAG|\r\n)/)) != null ? match[1] : null,
					Etag: (match = response.match(/\r\nETAG: ("\d{3}")\r\n/)) != null ? match[1] : null,
					RequestDigest: RequestDigest,
					Data: JSON.parse(response.match(/\{.*\}/)![0]),
					ProcessedData: [] as unknown as TSPResponseData
				}) - 1;
			else
				errorResponses.push({
					ResponseIndex: responseIndex++,
					RequestedUrl: requestedUrls[urlIndex++],
					HttpStatus: httpCode,
					ContentType: (match = response.match(/CONTENT-TYPE: (.*)\r\n(ETAG|\r\n)/)) != null ? match[1] : null,
					Etag: (match = response.match(/\r\nETAG: ("\d{3}")\r\n/)) != null ? match[1] : null,
					RequestDigest: RequestDigest,
					Data: JSON.parse(response.match(/\{.*\}/)![0]),
					ProcessedData: undefined
				});
		const followupNextRequests: Promise<TSPResponseData>[] = [];
		let successResponse: TSPResponseData;

		for (const response of successResponses) {
			successResponse = response.Data as TSPResponseData;
			if (successResponse.d && successResponse.d.__next)
				followupNextRequests.push(new Promise((resolve, reject) => {
					collectNext(successResponse).then((response: TSPResponseData) => {
						resolve(response);
					}).catch((response) => {
						reject(response);
					});
				}));
		}
		// check for an collect
		Promise.all(followupNextRequests).then((fetches: TSPResponseData[]) => {
			for (idx = 0; idx < fetches.length; idx++)
				successResponses[idx].ProcessedData = fetches[idx];
			resolve({
				success: successResponses,
				error: errorResponses
			});
		}).catch((err: unknown) => {
			reject(err);
		});
	});

	function collectNext(
		responseObj: TSPResponseData,
		carryData?: unknown[],
	): Promise<TSPResponseData> {
		return new Promise((resolve, reject) => {
			// conditions where __next is property is not found
			if (!((responseObj.d && responseObj.d.__next) || responseObj.value || responseObj.__next)) {
				resolve(responseObj);
				return;
			}
			if (!carryData)
				carryData = [];
			else
				carryData = carryData.concat(responseObj.value);
			const nextLink = responseObj["odata.nextLink"] ||
					(responseObj.d && responseObj.d["__next"] ? responseObj.d["__next"] : null) ||
					(responseObj["__next"] ? responseObj["__next"] : null);
			if (nextLink == null) {
				// TODO  fix this
				resolve(carryData as unknown as TSPResponseData);
				return;
			}
			fetch(nextLink, {
				method: "GET",
				headers: {
					// element.multipart boundary usually  "batch_" + guid()
					"Content-Type": "application/json;odata=verbose",
					"Accept": "application/json;odata=nometadata"
				}
			}).then((response: Response) => response.json()
				.then(json => {
					if (json["odata.nextLink"])
						resolve(collectNext(json, carryData));
					//				else
					//					resolve(carryData!.concat(json));
				}).catch(err => {
				// TODO  fetch error fixup
					reject({
						reqObj: err,
						addlMessage: err.status == 404 ? "A __next link returned 404 error" : ""
					});
				})).catch(err => {
					// TODO  fetch error fixup
						reject({
							reqObj: err,
							addlMessage: "fetch formatting to JSON problem"
						});
					});
		});
	}
}
