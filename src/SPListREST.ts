"use strict";

/* jshint -W119, -W069, -W083 */

import * as SPRESTTypes from './SPRESTtypes';
import { SPSiteREST } from './SPSiteREST';
import * as SPRESTSupportLib from './SPRESTSupportLib';
import * as SPRESTGlobals from './SPRESTGlobals';

/***************************************************************
 *  Basic interface
 * 		spList = new SPListREST(setup: TListSetup)
 *
 *
 *
 ***************************************************************/

/**
 * @class SPListREST  -- constructor for interface to making REST request for SP Lists
 */
export class SPListREST {
	protocol: string;
	server: string;
	site: string;
	apiPrefix: string;
	listName: string = "";
	listGuid: string = "";
	baseUrl: string = "";
	serverRelativeUrl: string = "";
	creationDate: Date = new Date(0);
	baseTemplate: string = "";
	allowContentTypes: boolean = Boolean(null);
	itemCount: number = -1;
	listItemEntityTypeFullName: string = "";
	rootFolder: string = "";
	contentTypes: object = {};
	listIds: number[] = [];
	fields: any[] = [];
	fieldInfo: any[] = [];
	lookupFieldInfo: SPRESTTypes.TLookupFieldInfo[] = [];
	lookupInfoCached: boolean = Boolean(null);
	linkToDocumentContentTypeId: string = "";
	currentListIdIndex: number = -1;
	setup: SPRESTTypes.TListSetup;
	sitePedigree: SPRESTTypes.TSiteInfo = {};
	arrayPedigree: SPRESTTypes.TSiteInfo[] = [];

	static stdHeaders: SPRESTTypes.THttpRequestHeaders = {
		"Accept": "application/json;odata=verbose",
		"Content-Type": "application/json;odata=verbose"
	};

	/**
	 * @constructor
	 * @param {object:
	 *	server: {string} server domain, equivalent of location.hostname return value
	* 	site: {string}  URL string part to the site
	*		list: {string}  must be a valid list name. To create a list, use SPSiteREST() object
	*    include: {string}  comma-separated properties to be part of OData $expand for more properties
	*	} setup -- object to initialize the
	*/
	constructor(setup: SPRESTTypes.TListSetup) {
		let matches: string[] | null;

		if (this instanceof SPListREST == false)
			throw "Was 'new' operator used? Use of this function object should be with constructor";
		if (setup.protocol)
			this.protocol = setup.protocol;
		else
			this.protocol = "https";
		if (this.protocol.search(/https?$/) == 0)
			this.protocol += "://";
		if (!setup.server)
			throw "List REST constructor requires defining server as arg { server: <server-name> } e.g. \"https://tenant.sharepoint.com\"";
		this.server = setup.server;
		if (this.server.search(/https?/) == 0 && (matches = this.server.match(/https?:\/\/(.*)/)) != null)
			this.server = matches[1];
		if (!setup.site)
			throw "List REST constructor requires defining site as arg { site: <site-name> } e.g. \"//teams/path/to/my/site\"";
		this.site = setup.site;
		if (setup.site.charAt(0) != "/")
			this.site = "/" + this.site;
		if (setup.listName)
			this.listName = setup.listName;
		else if (setup.listGuid)
			this.listGuid = setup.listGuid;
		if (this.listGuid && (matches = this.listGuid.match(/\{?([\d\w\-]+)\}?/)) != null)
			this.listGuid = matches[1];
		if (!(this.listName || this.listGuid))
			throw "Neither a list name or list GUID were provided to identify a list";
		this.setup = setup;
		this.apiPrefix = this.protocol + this.server + this.site + "/_api";
/*
		if (setup.linkToDocumentContentTypeId)
			this.linkToDocumentContentTypeId = setup.linkToDocumentContentTypeId; */
	}
	static escapeApostrophe (string: string): string {
		return encodeURIComponent(string).replace(/'/g, '\'\'');
	};

	/**
	 * @method _init -- sets many list characteristics whose data is required
	 * 	by asynchronous request
	 * @returns Promise
	 */
	init() {
		return new Promise((resolve, reject) => {
			$.ajax({
				url:  this.apiPrefix + "/web/lists" +
					(this.listName ? "/getByTitle('" + this.listName + "')" :
					"(guid'" + this.listGuid + "')") +
					"?$expand=RootFolder,ContentTypes,Fields" +
					(this.setup && this.setup.include ? "," + this.setup.include : ""),
				method: "GET",
				headers: SPListREST.stdHeaders,
				success: (data) => {
					let loopData: any[],
						lookupFieldTypeNum = SPRESTSupportLib.SPFieldTypes.getFieldTypeIdFromTypeName("Lookup");

					data = data.d;
					this.baseUrl = this.server + this.site;
					this.serverRelativeUrl = data.RootFolder.ServerRelativeUrl;
					this.creationDate = data.Created;
					this.baseTemplate = data.BaseTemplate;
					this.allowContentTypes = data.AllowContentTypes;
					this.listGuid = data.Id;
					this.listName = data.Title;
					this.itemCount = data.ItemCount;
					this.listItemEntityTypeFullName = data.ListItemEntityTypeFullName;
					this.rootFolder = data.RootFolder;
					this.contentTypes = data.ContentTypes.results;
					this.fields = data.Fields.results;
					if (this.setup.include) {
						let components: string[] = this.setup.include.split(",");

						for (let component of components)
							//this[component] = data[component];
							Object.defineProperty(this, data[component], component);
					}
							// for doc libs, look for "Link To Document" content type and store its content type ID
					// one more trip to the network to get any lookup field information
					this.getLookupFieldsInfo().then((response) => {
						if (SPRESTSupportLib.isIterable(loopData = data.ContentTypes.results) == true) {
							let i;

							for (i = 0; i < loopData.length; i++)
								if (loopData[i].Name == "Link to a Document")
									break;
							if (i < loopData.length)
								this.linkToDocumentContentTypeId = loopData[i].Id.StringValue;
						}
						resolve(true);
					}).catch((response) => {
						reject("List info error\n" + JSON.stringify(response, null, "  "));
					});
				},
				error: (reqObj) => {
					reject("List info error\n" + JSON.stringify(reqObj, null, "  "));
				}
			});
		});
	} // _init()

	/**
	 * @method getLookupFieldsInfo -- builds this.lookupFieldsInfo when a lib/list has lookup fields
	 * 		THIS FUNCTION MAY NEED DEBUGGING
	 * @returns
	 */
	getLookupFieldsInfo() {
		return new Promise((resolve, reject) => {
	// first get the site pedigree of current site
			let siteREST = new SPSiteREST({
				server: this.server,
				site: this.site
			});
			siteREST.init().then(() => {
				siteREST.getSitePedigree(null).then(({pedigree, arrayPedigree}) => {
	// after site pedigree, collect the lookup fields and lists as
					let lookupFields: any[] = [],
						requests: Promise<any>[] = [];

					this.sitePedigree = pedigree;
					this.arrayPedigree = arrayPedigree;
					for (let field of this.fields)
						if (field.FieldTypeKind == 7 && field.FromBaseType == false) // lookup type, so get the data
							lookupFields.push(field);

					for (let field of lookupFields)
						requests.push(this.retrieveLookupFieldInfo({
							useFunction: "siteREST",
							LookupWebId: field.LookupWebId,
							LookupList: field.LookupList,
							LookupField: field.LookupField
						}));

					Promise.all(requests).then((responses) => {
						console.log("responses count = " + responses.length + "\n" +
							JSON.stringify(responses, null, "  "));
						for (let i of responses) {
							let data: any[] = i.data,
									fld: any,
									choices: {id: number; value: string }[] = [],
									idx: number;

							found1:
							for (fld of lookupFields)
								for (let datum of data)
									if (datum[fld.Title] != null)
										break found1;
							idx = this.lookupFieldInfo.push({
								fieldDisplayName: fld.Title,
								fieldInternalName: fld.InternalName,
								choices: null
							}) - 1;
							for (let fldVal of data)
								choices.push({
									id: fldVal.Id,
									value: fldVal[fld.InternalName]
								});
							this.lookupFieldInfo[idx].choices = choices;
						}
						resolve(true);
					}).catch((response) => {
						reject(response);
					});

				}).catch((response) => {
					reject(response);
				});
			}).catch((response) => {
				reject(response);
			});
		});
	}

/**
 * @method retrieveLookupFieldInfo -- Promise to return choices for a lookup value
 * @param {object
 * 		.useFunction: {string}  "search" (default)|"enumWebs"
 *      .LookupWebId:  must be GUID form of the site containing the list that is lookup
 *      .LookupList:   must be GUID form the list containing column lookup
 *      .LookupField:  must be the Internal Name form of the field in the list looked up
 *    } parameters -- object properties specified
 * @returns {Promise} the array of values that form the choice list of the lookup
 */
	retrieveLookupFieldInfo(parameters: {
		useFunction?: "enumWebs" | "siteREST";
		LookupWebId: string;
		LookupList: string;
		LookupField: string;
	}): Promise<any> {
		return new Promise((resolve, reject) => {
			if (parameters.useFunction == "siteREST") {
				let site: SPRESTTypes.TSiteInfo = {},
					lookupListId = parameters.LookupList.replace(/[{}]/g, "");

				for (site of this.arrayPedigree)
					if (site.id == parameters.LookupWebId)
						break;
				SPListREST.httpRequestPromise({
					url: this.protocol + this.server + site.serverRelativeUrl + "/_api/lists(guid'" +
							lookupListId + "')/items?$select=Id," + parameters.LookupField
				}).then((response) => {
					resolve(response);
				}).catch((response) => {
					reject(response);
				});

			} else // no useFunction specified
				$.ajax({  // use SharePoint's search query tool
					// search for the web id
					url: this.apiPrefix + "/search/query?querytext=guid'" + parameters.LookupWebId + "'",
					method: "GET",
					headers: SPListREST.stdHeaders,
					success: (data) => {
						// the data returned are the lookup choice values in the target
						let siteName,
							lookupList,
							siteResult = null,
							lookupChoices: {
								id: string;  // item.Id
								choiceValue: string  // item[parameters.LookupField]
							}[] = [],
								searchRows = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results,
								// this variabled function will get the choices from the lookup field
								collectItemData = (url: string) => {
									// this named function will be used
									$.ajax({
										url: url,
										method: "GET",
										headers: SPListREST.stdHeaders,
										success: (data) => {
											for (let item of data.d.results)
												lookupChoices.push({
													id: item.Id,
													choiceValue: item[parameters.LookupField]
												});
											if (data.d.__next)
												collectItemData(data.d.__next);
											else
												resolve(lookupChoices);
										},
										error: (reqObj) => {
											reject(reqObj);
										}
									});
								};
						// This complex multiple loop is trying to find the results in the complex response
						//   of the SharePoint search query
						for (let searchRow of searchRows) {
							for (let searchCell of searchRow.Cells.results) {
								if (searchCell.Key == "WebId" && searchCell.Value == parameters.LookupWebId) {
									for (let searchCell2 of searchRow.Cells.results)
										if (searchCell2.Key == "SPWebUrl") {
											siteResult = searchCell2.Value;
											break;
										}
								}
								if (siteResult != null)
									break;
							}
							if (siteResult != null)
								break;
						}
						// if the loop found the result and set it, this condition is ready
						if (siteResult != null) {
							siteName = SPRESTSupportLib.ParseSPUrl(siteResult)!.siteFullPath;
							lookupList = parameters.LookupList.match(/^\{?([^\}]+)\}?$/)![1];
							collectItemData(this.protocol + this.server + siteName + "/_api/web/lists(guid'" + lookupList +
										"')/items?$select=Id," + parameters.LookupField);
						}
						resolve(null); // could not find a result in search
					},
					error: () => {
						reject("REST Search request failed");
					}
				});
		});
	}

	/**
	 * @method getLookupFieldValue -- retrieves as a Promise the value to set up for a lookup field
	 * @param {string} fieldName --
	 * @returns the value to set for the lookup field, null if not available
	 */
	getLookupFieldValue(fieldName: string, fieldValue: string): number | null {
		for (let info of this.lookupFieldInfo)
			if (info.fieldDisplayName == fieldName)
				for (let i: number = 0; i < info.choices!.length; i++)
					if (fieldValue == info.choices![i].value)
						return info.choices![i].id;
		return null;
	}

	/**
	 * @method fixBodyForSpecialFields -- support function to update data for a REST request body
	 *    where the column data is special, such as multichoice or lookup or managed metadata
	 * @param {object} body -- usually the XMLHTTP request body passed to a POST-type request
	 * @returns -- a Promise (this is an async call)
	 */
	fixBodyForSpecialFields(body: SPRESTTypes.THttpRequestBody) {
		let newBody: {[key:string]: any},
				value: number | null;

		return new Promise((resolve, reject) => {
			if (this.lookupInfoCached == true) {
// this is a quick way to get lookup value without going back to the server
				newBody = { };
				for (let key in body)
					if ((value = this.getLookupFieldValue(key, body[key])) == body[key])
						newBody[key] = value;
					else
						newBody[key + "Id"] = value;
				resolve(newBody);
			} else
// with this block, we go to the server to get lookup values
				this.getFields({
					filter: "FromBaseType eq false"
				}).then((response: any) => {
					let requests = [ ];

					// get the lists fields and store the info
					if (this.fields)
						this.fields.concat(response.data.d.results);
					else
						this.fields = response.data.d.results;
					this.fieldInfo = [ ];
					// this will be multiple requests, so find Promise.all()
					for (let fld of response.data.d.results)
						requests.push(new Promise((resolve, reject) => {
							// lookup type field
							if (fld.FieldTypeKind == 7 && typeof fld.LookupWebId == "string" &&
										typeof fld.LookupList == "string" && typeof fld.LookupField == "string" &&
										fld.LookupList.length > 0 && fld.LookupField.length > 0) { // lookup
								this.retrieveLookupFieldInfo({
									LookupWebId: fld.LookupWebId,
									LookupList: fld.LookupList,
									LookupField: fld.LookupField
								}).then((response: any) => {
									this.lookupFieldInfo.push({
										fieldDisplayName: fld.Title,
										fieldInternalName: fld.InternalName,
										choices: response
									});
								}).catch((response) => {
									reject(response);
								});
							// multi-choice field (not lookup)
							} else if (fld.FieldTypeKind == 15) { // multi-choice field
								resolve(this.fieldInfo.push({
									intName: fld.InternalName,
									multiple: 15
								}));
							} else if (fld.FieldTypeKind == 8) { // boolean
								resolve(this.fieldInfo.push({
									intName: fld.InternalName,
									boolean: true
								}));
							} else
								resolve(null); // not a lookup field, so don't store it
						}));
					Promise.all(requests).then(() => {
						// build the body from stored lookup fields, if they exist
						newBody = { };

						for (let key in body)
							if (Array.isArray(body[key])) {
								for (let fld of this.fields)
									if (fld.InternalName == key) {
										newBody[key] = {
											__metadata: {
												type: fld.Choices.__metadata.type
											},
											results: body[key]
										};
										break;
									}
							} else if (typeof body[key] == "boolean" ||
										(typeof body[key] == "string" && body[key].search(/Y|Yes|No|N|true|false/i) == 0)) {
								for (let fld of this.fieldInfo)
									if (fld.intName == key)
										if (fld.boolean)
											if (body[key] == true || body[key].search(/Y|Yes|true/i) == 0)
												newBody[key] = true;
											else
												newBody[key] = false;
										else
											newBody[key] = body[key];
							} else if ((value = this.getLookupFieldValue(key, body[key])) == body[key])
								newBody[key] = value; // this is the case
							else
								newBody[key + "Id"] = value;
						resolve(newBody);
					}).catch((response) => {
						reject(response);
					});
				}).catch(() => {
					reject(".getFields() called failed");
				});
		});
	}

	// if parameters.success not found the request will be SYNCHRONOUS and not ASYNC
	static httpRequest(elements: SPRESTTypes.THttpRequestParams) {
		if (!elements.successCallback)
			throw "HTTP Request for SPListREST requires defining 'successCallback' in parameter 'elements'";
		if (!elements.errorCallback)
			throw "HTTP Request for SPListREST requires defining 'errorCallback' in parameter 'elements'";
		if (typeof elements.headers == "undefined")
			elements.headers = SPListREST.stdHeaders;
		else {
			if (!elements.headers["Accept"])
				elements.headers["Accept"] = "application/json;odata=verbose";
			if (!elements.headers["Content-Type"])
				elements.headers["Content-Type"] = "application/json;odata=verbose";
		}
		if (elements.setDigest && elements.setDigest == true) {
			let match = elements.url.match(/(.*\/_api)/);

			if (match == null)
				throw "Problem parsing '/_api/' of URL to get request digest value";
			$.ajax({  // get digest token
				url: match[1] + "/contextinfo",
				method: "POST",
				headers: SPListREST.stdHeaders,
				success: (data: SPRESTTypes.TSPResponseData) => {
					elements.headers!["X-RequestDigest"] = data.FormDigestValue ? data.FormDigestValue :
						data.d!.GetContextWebInformation!.FormDigestValue;
					$.ajax({
						url: elements.url,
						method: elements.method,
						headers: elements.headers,
						data: elements.body as any ?? elements.data as any,
						success: (data: SPRESTTypes.TSPResponseData, status: string, requestObj: JQueryXHR) => {
							elements.successCallback!(data, status, requestObj);
						},
						error: (requestObj: JQueryXHR, status: string, thrownErr: string) => {
							elements.errorCallback!(requestObj, status, thrownErr);
						}
					});
				},
				error: (requestObj: JQueryXHR, status: string, thrownErr: string) => {
					elements.errorCallback!(requestObj, status, thrownErr);
				}
			});
		} else {
			if (!elements.method)
			 	elements.method = "GET";
			$.ajax({
				url: elements.url,
				method: elements.method,
				headers: elements.headers,
				data: elements.body as any ?? elements.data as any,
				success: (data: SPRESTTypes.TSPResponseData, status: string, requestObj: JQueryXHR) => {
						 if (data.d && data.d.__next)
							  this.RequestAgain(
									data.d.__next,
									data.d.results as any[]
							  ).then((response: any) => {
									elements.successCallback!(response);
							  }).catch((response) => {
									elements.errorCallback!(response);
							  });
						 else
						 	elements.successCallback!(data, status, requestObj);
				},
				error: (requestObj: JQueryXHR, status: string, thrownErr: string) => {
					elements.errorCallback!(requestObj, status, thrownErr);
				}
			});
		}
	}

	static RequestAgain(nextUrl: string, aggregateData: any[]) {
		return new Promise((resolve, reject) => {
			$.ajax({
				url: nextUrl,
				method: "GET",
				headers: SPListREST.stdHeaders,
				success: (data) => {
					if (data.d.__next) {
						this.RequestAgain(
							data.d.__next,
							data.d.results
						).then((response) => {
							resolve(aggregateData.concat(response));
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

	static httpRequestPromise(parameters: SPRESTTypes.THttpRequestParamsWithPromise) {
		return new Promise((resolve, reject) => {
			SPListREST.httpRequest({
				setDigest: parameters.setDigest,
				url: parameters.url,
				method: parameters.method,
				headers: parameters.headers,
				data: parameters.data ?? parameters.body,
				successCallback: (data: SPRESTTypes.TSPResponseData, status: string | undefined,
							reqObj: JQueryXHR | undefined) => {
					resolve({data: data, message: status, reqObj: reqObj});
				},
				errorCallback: (reqObj: JQueryXHR, text: string | undefined, errThrown: string | undefined) => {
					reject({reqObj: reqObj, text: text, errThrown: errThrown});
				}
			});
		});
	}

	// query: [optional]
	getProperties (parameters: any): Promise<any> {
		return new Promise((resolve, reject) => {
			let query: string = SPRESTSupportLib.constructQueryParameters(parameters);

			SPListREST.httpRequestPromise({
				url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')" + query,
				method: "GET",
			}).then((response) => {
				resolve(response);
			}).catch((response) => {
				reject(response);
			});
		});
	}

	getListProperties(parameters?: any): Promise<any> {
		return this.getProperties(parameters);
	}

	getListInfo(parameters?: any): Promise<any> {
		return this.getProperties(parameters);
	}

	/**
	 * @method getListItemData
	 * @param {object --
	 * 		itemId: specific ID of item data to be retrieved
	 * 		lowId: starts a range with the low ID
	 * 		highId: starts a range with the high ID
	 * 	} parameters
	 * @returns {Promise} HTTP Request
	 */
	getListItemData(parameters: {
		itemId: number;
		lowId?: number;
		highId?: number;
		select?: string | null,
		expand?: string | null,
		filter?: string | null,
		selectDisplay?: string[],
		[key: string]: any;
	}): Promise<any> {
		return new Promise((resolve, reject) => {
			let query: string = "",
					filter: string = "";

			if (parameters.selectDisplay && parameters.selectDisplay.length > 0)
			if (parameters && parameters.lowId)
				filter += "Id ge " + parameters.lowId as string;
			if (parameters && parameters.lowId && parameters.highId)
				if (parameters.lowId >= parameters.highId)
					filter += " or ";
				else
					filter += " and ";
			if (parameters && parameters.highId && parameters.highId > 0)
				filter += "Id le " + parameters.highId as string;
			if (parameters.filter)
				filter = "(" + parameters.filter + ") and (" + filter + ")";
			if (parameters.select && parameters.select.length > 0)
				query = "?$select=" + parameters.select;
			if (parameters.expand && parameters.expand.length > 0)
				if (query.length > 0)
					query += "&$expand=" + parameters.expand;
				else
					query = "?$expand=" + parameters.expand;
			if (filter.length > 0)
				if (query.length > 0)
					query += "&$filter=" + filter;
				else
					query = "?filter=" + filter;
			SPListREST.httpRequestPromise({
				url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/items" +
						(parameters && parameters.itemId > 0 ? "(" + parameters.itemId + ")" : "") + query
			}).then((response: any) => {
				if (response.data.__next)
					this.getListItemData(response.data.__next);
				else
					resolve(response);
			}).catch((response) => {
				reject(response);
			});
		});
	}

	getListItem(parameters: any): Promise<any> {
		return this.getListItemData(parameters);
	}

	getItemData(parameters: any): Promise<any> {
		return this.getListItemData(parameters);
	}

	getAllListItems(): Promise<any> {
		return this.getListItemData({itemId: -1});
	}

	getListItemsWithQuery(parameters: {
		select: string | null,
		expand: string | null,
		filter: string | null,
		selectDisplay?: string[] // this must be used to get columns/fields data with that display name
	}): Promise<any> {
		let newParameters = {
			select: parameters.select,
			expand: parameters.expand,
			filter: parameters.filter,
			selectDisplay: parameters.selectDisplay,
			itemId: -1
		};
		return this.getListItemData(newParameters);
	}

	getAllListItemsOptionalQuery(parameters: any): Promise<any> {
		return this.getListItems(parameters);
	}

	getListItems(parameters?: any): Promise<any> {
		return this.getListItemData(parameters);
	}

	// @param {object} parameters - should have at least {body:,success:}
	// body format: {string} " 'fieldInternalName': 'value', ...}
	createListItem(item: {body: SPRESTTypes.THttpRequestBody}) {
		let body: SPRESTTypes.THttpRequestBody | string = item.body;

		if (!body || body == null)
			throw "The object argument to createListItem() must have a 'body' property";
//		if (this.checkEntityTypeProperty(body, "item") == false)
//			body["__SetType__"] = this.listItemEntityTypeFullName;
		body = SPRESTSupportLib.formatRESTBody(body);
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid +
					"')/items",
			method: "POST",
			body: body
		});
	}

	/**
	 *
	 * @param {arrayOfObject} items -- objects should be objects whose properties are fields
	 * 		within the lists and using Internal name property of field
	 * @returns a promise
	 */
	createListItems(items: any[], useBatch: boolean) {
		if (useBatch == true) {
			let requests = [];
			// need to re-configure
			for (let item of items)
				requests.push({
					url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/items",
					body: item.body
				});
			return SPRESTSupportLib.batchRequestingQueue({
				host: this.server,
				path: this.site,
				AllHeaders: {
					"Content-Type": "application/json;odata=verbose"
				},
				AllMethods: "POST"
			}, requests);
		}
	// the alternative to batching
		return new Promise((resolve, reject) => {
			SPRESTSupportLib.serialSPProcessing(this.createListItem, items).then((response: any) => {
				resolve(response);
			}).catch((response: any) => {
				reject(response);
			});
		});
	}

	/**
	 * @method updateListItem -- modifies existing list item
	 * @param {object} parameters -- requires 2 parameters {itemId:, body: [JSON] }
	 * @returns
	 */
	updateListItem(parameters: {
		itemId: number;
		body: SPRESTTypes.THttpRequestBody;
	}) {
		let body: SPRESTTypes.THttpRequestBody | string = parameters.body;

		return new Promise((resolve, reject) => {
			this.fixBodyForSpecialFields(body as SPRESTTypes.THttpRequestBody).then((body: any) => {
				if (SPRESTSupportLib.checkEntityTypeProperty(body, "item") == false)
					body["__SetType__"] = this.listItemEntityTypeFullName;
				body = SPRESTSupportLib.formatRESTBody(body);
				SPListREST.httpRequestPromise({
					setDigest: true,
					url: this.apiPrefix + "/web/lists(guid'" + this.listGuid +
								"')/items(" + parameters.itemId + ")",
					method: "POST",
					headers: {
						"X-HTTP-METHOD": "MERGE",
						"IF-MATCH": "*" // can also use etag
					},
					body: body,
				}).then((response) => {
					resolve(response);
				}).catch((response) => {
					reject(response);
				});
			}).catch((response) => {
				reject(response);
			});
		});
	}

	updateListItems(itemsArray: any[], useBatch: boolean) {
		if (useBatch == true) {
			let requests: {url: string; body: object}[] = [];
			// need to re-configure
			for (let item of itemsArray)
				requests.push({
					url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/items(" +
							item.id + ")",
					body: item.body
				});
			return SPRESTSupportLib.batchRequestingQueue({
				host: this.server,
				path: this.site,
				AllHeaders: {
					"Content-Type": "application/json;odata=verbose",
					"IF-MATCH": "*"
				},
				AllMethods: "PATCH"
			}, requests);
		}
	// the alternative to batching
		return new Promise((resolve, reject) => {
			SPRESTSupportLib.serialSPProcessing(this.updateListItem, itemsArray).then((response: any) => {
				resolve(response);
			}).catch((response: any) => {
				reject(response);
			});
		});
	}

	/**
	applies to all lists
	parameters properties should be
	.itemId [required]
	.recycle [optional, default=true]. Use false to get hard delete
	*/
	deleteListItem(parameters: {itemId: number; recycle: boolean;}): Promise<any> {
		let recycle = parameters.recycle;

		if (!recycle)
			throw "Parameter 'recycle' is required. If 'true', item goes to recycle bin. " +
					"If 'false', item is permanently deleted.";
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists" +
					(this.listName ?
						"/GetByTitle('" + this.listName + "')" :
						"(guid'" + this.listGuid + "')") +
					"/items(" + parameters.itemId + ")" +
					(recycle == true ? "/recycle()" : ""),
			method: recycle == true ? "POST" : "DELETE",
			headers: recycle == true ? { "IF-MATCH": "*" } : { "X-HTTP-METHOD": "DELETE"}
		});
	}

	getListItemCount (doFresh: boolean = false): Promise<any> {
		return new Promise((resolve, reject) => {
			if (doFresh == false)
				resolve(this.itemCount);
			else
				SPListREST.httpRequest({
					url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')",
					successCallback: (data: SPRESTTypes.TSPResponseData) => {
						resolve(data.d!.ItemCount);
					},
					errorCallback: (reqObj: JQueryXHR) => {
						reject(reqObj);
					}
				});
		});
	}

	loadListItemIds() {
		return new Promise((resolve, reject) => {
			if (this.listIds && this.listIds.length > 0)
				resolve(this.listIds);
			else
				this.getListItemCount().then(() => {
					this.getAllListItemsOptionalQuery({
						query: "?$select=ID"
					}).then((response) => {
						let i, results, IDs = [];

						if (response.data.d)
							results = response.data.d.results;
						else if (response.data) // this occurs with response.data.__next
							results = response.data;
						for (i = 0; i < results.length; i++)
							IDs.push(results[i].Id);
						this.listIds = IDs;
						resolve(this.listIds);
					}).catch((response) => {
						reject(response);
					});
				}).catch((response) => {
					reject(response);
				});
		});
	}

	isValidID(itemId: number): Promise<any> {
		return new Promise((resolve, reject) => {
			this.getListItemIds().then((response: number[]) => {
				resolve(response.includes(itemId));
			}).catch((response) => {
				reject(response);
			});
		});
	}

	getListItemIds(): Promise<any> {
		return new Promise((resolve, reject) => {
			this.loadListItemIds().then( () =>  {
				resolve(this.listIds);
			}).catch((response) => {
				reject(response);
			});
		});
	}

	getNextListId(): Promise<any> {
		return new Promise((resolve, reject) => {
			this.loadListItemIds().then( () =>  {
				if (this.currentListIdIndex == this.listIds.length)
					resolve(this.listIds[this.currentListIdIndex]);
				else
					resolve(this.listIds[++this.currentListIdIndex]);
			}).catch((response) => {
				reject (response);
			});
		});
	}

	getPreviousListId() {
		let thisInstance = this;
		return new Promise(function (resolve, reject) {
			thisInstance.loadListItemIds().then( () =>  {
				if (thisInstance.currentListIdIndex == 0)
					resolve(thisInstance.listIds[thisInstance.currentListIdIndex]);
				else
					resolve(thisInstance.listIds[--thisInstance.currentListIdIndex]);
			}).catch((response: any) => {
				reject (response);
			});
		});
	}

	getFirstListId(): Promise<any> {
		let thisInstance = this;
		return new Promise(function (resolve, reject) {
			thisInstance.loadListItemIds().then( () =>  {
				resolve(thisInstance.listIds[thisInstance.currentListIdIndex = 0]);
			}).catch((response) => {
				reject (response);
			});
		});
	}

	setCurrentListId(itemId: number | string): boolean {
		if (!this.listIds)
			throw "Cannot call IListRESTRequest.setCurrentListId() unless List IDs are loaded";
		if ((this.currentListIdIndex = this.listIds.indexOf(parseInt(itemId as string))) < 0)
			return false;
		return true;
	}

	// ==============================================================================================
	// ======================================  DOCLIB related REST requests =========================
	// ==============================================================================================
	getAllDocLibFiles(): Promise<any> {
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/Files",
			method: "GET",
		});
	}

	getFolderFilesOptionalQuery(parameters: {
		folderPath: string;
	}): Promise<any> {
		let query: string = SPRESTSupportLib.constructQueryParameters(parameters);
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix +	"/web/getFolderByServerRelativeUrl('" + this.baseUrl +
						parameters.folderPath + "')/Files" + query,
			method: "GET"
		});
	}

	/** @method getDocLibItemByFileName
	retrieves item by file name and also returns item metadata
	* @param {Object} parameters
	* @param {string} parameters.fileName - name of file
	*/
	getDocLibItemByFileName(parameters: {fileName?:string,itemName?:string}): Promise<any> {
		if (!(parameters.fileName || parameters.itemName))
			throw "Method requires 'parameters.fileName' or 'parameters.itemName' " +
				"to be defined";
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/getFolderByServerRelativeUrl('" +
						this.baseUrl + "')/Files('" + SPListREST.escapeApostrophe(parameters.fileName!) +
				"')?$expand=ListItemAllFields",
			method: "GET"
		});
	}

	/** @method getDocLibItemMetadata
	 * @param {Object} parameters
	 * @param {number|string} parameters.itemId - ID of item whose data is wanted
	 * @returns  {Object} returns only the metadata about the file item
	 */
	getDocLibItemMetadata(parameters: {itemId:number}): Promise<any> {
		if (!parameters.itemId)
			throw "Method requires 'parameters.itemId' to be defined";
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid +
						"')/items(" + parameters.itemId + ")",
		});
	}
	/** @method getDocLibItemFileAndMetaData
	 * @param {Object} parameters
	 * @param {number|string} parameters.itemId - ID of item whose data is wanted
	 * @returns {Object} data about the file and metadata of the library item
	 */

	getDocLibItemFileAndMetaData(parameters: {itemId: number}): Promise<any> {
		if (!parameters.itemId)
			throw "Method requires 'parameters.itemId' to be defined";
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid +
					"')/items(" + parameters.itemId + ")/File?$expand=ListItemAllFields",
		});
	}

	/**
	arguments as parameters properties:
	.fileName or .itemName -- required which will be name applied to file data
	.folderPath -- optional, if omitted, uploaded to root folder
	.body  required file data (not metadata)
	.willOverwrite [optional, default = false]
	*/
	uploadItemToDocLib(parameters: {
		fileName: string;
		body: SPRESTTypes.TXmlHttpRequestData;
		folderPath: string;
		willOverwrite: boolean;
	}): Promise<any> {
		let path: string = this.serverRelativeUrl;
		if (parameters.folderPath.charAt(0) == "/")
			path += parameters.folderPath;
		else
			path += "/" + parameters.folderPath;
		if (typeof parameters.willOverwrite == "undefined")
			parameters.willOverwrite = false;
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/getFolderByServerRelativeUrl('" + path + "')" +
					"/Files/add(url='" + SPListREST.escapeApostrophe(parameters.fileName) +
					"',overwrite=" + (parameters.willOverwrite == true ? "true" : "false") +
					")?$expand=ListItemAllFields",
			method: "POST",
			body: parameters.body
		});
	}

	createFolder(parameters: {folderName: string}): Promise<any> {
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/folders",
			method: "POST",
			body: JSON.stringify({
				"__metadata": {
					"type": "SP.Folder"
				},
				"ServerRelativeUrl": this.serverRelativeUrl + "/" + parameters.folderName
			})
		});
	}

/*/
	/**
	required arguments as parameters properties:
	.fileName or .itemName
	.url   required string which is URL representing the link
	.fileType   required string
	.willOverwrite [optional, default = false]

	createLinkToDocItemInDocLib(parameters) {
		return new Promise(function (resolve, reject) {
			if (!this.linkToDocumentContentTypeId) {
				this.loadListBasics().then(() => {
					resolve(this.continueCreateLinkToDocItemInDocLib(parameters));
				}).catch(() => {
					reject("createLinkToDocItemInDocLib() requires IListRESTRequest " +
						"property 'linkToDocumentContentTypeId' to be defined\n\n" +
						"Use setLinkToDocumentContentTypeId() method or define object with " +
						"initializing parameter");
					alert("There was a system error. Contact the administrator.");
				});
			}
			else
				resolve(this.continueCreateLinkToDocItemInDocLib(parameters));
		});
	}
	/**
	 *  helper function for one above

	continueCreateLinkToDocItemInDocLib(parameters) {
		if (!(parameters.fileName || parameters.itemName))
			throw "Method requires 'parameters.fileName' or 'parameters.itemName' " +
				"to be defined";
		if (parameters.itemName)
			parameters.fileName = parameters.itemName;
		if (typeof parameters.willOverwrite == "undefined")
			parameters.willOverwrite = false;
		else
			parameters.willOverwrite = (parameters.willOverwrite == true ? "true" : "false");
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/getFolderByServerRelativeUrl('" +
						this.baseUrl + "')/Files/add(url='" + SPListREST.escapeApostrophe(
						parameters.fileName) + "',overwrite=" + parameters.willOverwrite +
						")?$expand=ListItemAllFields",
			method: "POST",
			body: this.LinkToDocumentASPXTemplate.replace(/\$\$LinkToDocUrl\$\$/g, parameters.url).
						replace(/\$\$LinkToDocumentContentTypeId\$\$/,
					this.linkToDocumentContentTypeId).replace(/\$\$filetype\$\$/, parameters.fileType)
		});
	}
/*/

	/**
	 * @method checkOutDocLibItem  checks out a document library item where checkout is required
	 * @param {string} .fileName or .itemName -- required which will be name applied to file data
	 * @param {string} .folderPath -- optional, if omitted, uploaded to root folder
	 * @param {byte} exceptions - flag on how to handle exceptions
	 *     only acceptable options are IGNORE_THIS_USER_CHECKOUT, IGNORE_NO_EXIST
	 * @param {object} exceptionsData. For IGNORE_THIS_USER_CHECKOUT, a string that
	 *     is part of the user name should be passed as {uname: <username string>}
	 */
	checkOutDocLibItem(parameters: {
		fileName?: string, itemName?: string,
		folderPath?: string
	}): Promise<any> {
		// parameters.tryCheckOutResolve == true will
		//   perform a user request of resolving without checking
		//   if user is identical with checked out user
		let path: string;
		if (!(parameters.fileName || parameters.itemName))
			throw "Method requires 'parameters.fileName' or 'parameters.itemName' " +
				"to be defined";
		if (parameters.itemName)
			parameters.fileName = parameters.itemName;
		path = this.serverRelativeUrl;
		if (parameters.folderPath)
			if (parameters.folderPath.charAt(0) == "/")
				path += parameters.folderPath;
			else
				path += "/" + parameters.folderPath;
		path += "/" + parameters.fileName;
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/getFileByServerRelativeUrl('" + path + "')/CheckOut()",
			method: "POST"
		});
	}


	/**
	 * @method checkInDocLibItem -- check in one or more files
	 * @param {object:
	 * 		{string[]?} itemNames -- must be array representing folder paths + file names
	 *       {string} checkInType -- same check in type for one or multiple files
	 * 				must be anyon of 6 strings: "[mM]inor" (0), "[mM]ajor" (1), "[oO]verwrite" (2)
	 *       {string} .checkInComment -- applied to all files if multipe
	 			* } parameters -- the following properties are recognized
	 * @returns Promise
	 */
	checkInDocLibItem(parameters: {
		itemNames: string[],
		checkInType?: SPRESTTypes.TSPDocLibCheckInType,
		checkInComment?: string
	}): Promise<any> {
		// checkintype = 0: minor version, = 1: major version, = 2: overwrite
		let requests: {url:string;body?:{}}[] = [],
			checkinType: number;
		if (!parameters.itemNames || parameters.itemNames.length == 0)
			throw "Method requires 'parameters.itemNames' as non-zero length arrays of strings";
		switch (parameters.checkInType) {
			case null:
			case undefined:
			case "minor":
			case "Minor":
				checkinType = 0;
				break;
			case "major":
			case "Major":
				checkinType = 1;
				break;
			case "overwrite":
			case "Overwrite":
				checkinType = 2;
				break;
		} // build the batch request
		// for IDs, must recover the name
		if (!parameters.checkInComment)
			parameters.checkInComment = "";
		for (let itemName of parameters.itemNames)
			requests.push({
				url: this.apiPrefix + "/web/GetFileByServerRelativeUrl('" + itemName +
						"')/CheckIn(comment='" + parameters.checkInComment + "',checkintype=" + checkinType + ")"
			});
		return SPRESTSupportLib.batchRequestingQueue({
			AllMethods: "POST"
		} as SPRESTTypes.IBatchHTTPRequestParams, requests);
	}

	/**
	 * @method discardCheckOutDocLibItem
	 * @param {Object} parameters
	 * @param {string} parameters.(fileName|itemName) - name of file checked out
	 * @param {string} parameters.folderPath - path to file name in doc lib
	 */
	discardCheckOutDocLibItem(parameters: {
		file: string;
		folderPath: string;
	}): Promise<any> {
		let path: string;

		path = this.baseUrl;
		if (parameters.folderPath.charAt(0) == "/")
			path += parameters.folderPath;
		else
			path += "/" + parameters.folderPath;
		path += "/" + parameters.file;
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/GetFileByServerRelativeUrl('" + path + "')" +
					"/UndoCheckout()",
			method: "POST"
		});
	}


	renameItem(parameters: any) {
		return this.renameFile(parameters);
	}

	renameFile(parameters: {
		itemId: number;
		itemName: string;
		newName: string;
	}): Promise<any> {
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid +
					"')/items(" + parameters.itemId + ")",
			method: "POST",
			headers: {
				"X-HTTP-METHOD": "MERGE",
				"IF-MATCH": "*" // can also use etag
			},
			body: SPRESTSupportLib.formatRESTBody({
				FileLeafRef: parameters.itemName ? parameters.itemName : parameters.newName
			}),
		});
	}


	// 3 arguments are necessary:
	// .itemId: ID of file item
	// .newFileName or .newName: name to be used in rename
	renameItemWithCheckout(parameters: {
		itemId: number;
		newName: string;
	}): Promise<any> {
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')" +
					"/items(" + parameters.itemId + ")",
			method: "POST",
			headers: {
				"X-HTTP-METHOD": "MERGE",
				"IF-MATCH": "*" // can also use etag
			},
			body: {
				"FileLeadRef": parameters.newName
			} as any
		});
	}

	/** @method updateLibItemWithCheckout
	 * @param {Object} parameters
	 * @param {string|number} parameters.(fileName|itemId) - id of lib item or name of file checked out
	 */
	updateLibItemWithCheckout(parameters: {
		itemId: number;
		fileName: string;
		body: string;
		checkInType?: SPRESTTypes.TSPDocLibCheckInType;
	}) {
		let thisInstance = this;
		return new Promise(function (resolve, reject) {
			if ((!parameters.itemId || isNaN(parameters.itemId) == true) &&
				!parameters.fileName)
				throw "argument must contain {itemId:<value>} or {fileName:<name>}";
			if (!parameters.body)
				throw "parameters.body not found/defined: must be formatted JSON string";
			if (!parameters.fileName)
				thisInstance.getListItemData({
					itemId: parameters.itemId,
					expand: "File"
				}).then((response: any) => {
					parameters.fileName = response.responseJSON.d.File.Name;
					thisInstance.continueupdateLibItemWithCheckout(parameters).then((response: any) => {
						resolve(response);
					}).catch((response: any) => {
						reject(response);
					});
				}).catch((response: any) => {
					reject(response);
				});
			else
				thisInstance.continueupdateLibItemWithCheckout(parameters).then((response: any) => {
					resolve(response);
				}).catch((response: any) => {
					reject(response);
				});
		});
	}

	/** @method continueupdateLibItemWithCheckout
	 * @param {Object} parameters
	 * @param {string|number} parameters.(fileName|itemId) - id of lib item or name of file checked out
	 */
	continueupdateLibItemWithCheckout(parameters: {
		itemId: number;
		fileName: string;
		checkInType?: SPRESTTypes.TSPDocLibCheckInType;
		body: string;
	}) {
		let thisInstance = this;
		return new Promise(function (resolve, reject) {
			thisInstance.checkOutDocLibItem({
				itemName: parameters.fileName,
			}).then(() => {
				SPListREST.httpRequestPromise({
					setDigest: true,
					url: thisInstance.apiPrefix + "/web/lists(guid'" + thisInstance.listGuid +
							"')/items(" + parameters.itemId + ")",
					method: "POST",
					body: "{ '__metadata': { 'type': '" +
						thisInstance.listItemEntityTypeFullName + "' }, " +
						parameters.body + " }",
					headers: {
						"X-HTTP-METHOD": "MERGE",
						"IF-MATCH": "*" // can also use etag
					},
				}).then(() => {
					thisInstance.checkInDocLibItem({
						itemNames: [ parameters.fileName ],
						checkInType: parameters.checkInType
					}).then((response: any) => {
						resolve(response);
					}).catch((response: any) => {
						reject(response);
					});
				}).catch((response: any) => {
					reject(response);
				});
			}).catch((response: any) => {
				reject(response);
			});
		});
	}

	/** @method discardCheckOutDocLibItem
	 * @param {Object} parameters
	 * @param {string} parameters.(fileName|itemName) - name of file checked out
	 * @param {string} parameters.folderPath - path to file name in doc lib
	 */
	updateDocLibItemMetadata(parameters: {
		itemId: number;
		body: SPRESTTypes.THttpRequestBody
	}) {
		let thisInstance = this;
		return new Promise(function (resolve, reject) {
			let itemName: string,
					majorResponse: any;
			thisInstance.getDocLibItemFileAndMetaData({
				itemId: parameters.itemId
			}).then((response: any) => {
				thisInstance.checkOutDocLibItem({
					itemName: itemName = response.responseJSON.d.Name
				}).then(() => {
					thisInstance.updateListItem({
						itemId: parameters.itemId,
						body: parameters.body
					}).then((response: any) => {
						majorResponse = response;
						thisInstance.checkInDocLibItem({
							itemNames: [ itemName ]
						}).then(() => {
							resolve(majorResponse);
						}).catch((response: any) => { // check in failure
							reject(response);
						});
					}).catch((response: any) => { // update failure
						reject(response);
					});
				}).catch((response: any) => { // check out failure
					reject(response);
				});
			});
		});
	}

	/**
	required arguments as parameters properties:
	.sourceFileName [required]  the path to the source name
	.destinationFileName [required]  the path to the location to be copied,
	can have file name different than source name
	.willOverwrite [optional, default = false]
	*/
	copyDocLibItem(parameters: {
		sourceFileName: string;
		destinationFileName: string;
		willOverwrite?: string;
	}): Promise<any> {
		let uniqueId;
		if (!parameters.sourceFileName)
			throw "parameters.sourceFileName not found/defined";
		if (!parameters.destinationFileName)
			throw "parameters.destinationFileName not found/defined";
		if (typeof parameters.willOverwrite == "undefined")
			parameters.willOverwrite = "false";
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/GetFileByServerRelativeUrl('" +
					parameters.sourceFileName + "')?$expand=ListItemAllFields",
			method: "GET"
		}).then((response: any) => {
			uniqueId = response.responseJSON.d.UniqueId;
			SPListREST.httpRequestPromise({
				url: this.apiPrefix + "/web/GetFileById(guid'" + uniqueId +
						"')/CopyTo(strNewUrl='" + parameters.destinationFileName +
						"',bOverWrite=" + parameters.willOverwrite + ")",
				method: "POST"
			});
		});
	}

	// .path
	// .folderPath
	// .fileName
	// .includeBaseUrl [optional, default=true]. Set to false if passing
	//    the item's .ServerRelativeUrl value
	//   set parameters.recycle = false to create a HARD delete (no recycle)
	// @return:
	//   for soft delete, responseJSON.d will be .Recycle property with GUID value
	deleteDocLibItem(parameters: {
		folderPath?: string;
		fileName?: string;
		path?: string;
		includeBaseUrl?: boolean;
		recycle?: boolean;
	}): Promise<any> {
		if (parameters.folderPath && parameters.fileName)
			parameters.path = parameters.folderPath + "/" + parameters.fileName;
		if (typeof parameters.includeBaseUrl == "undefined")
			parameters.includeBaseUrl = true;
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/GetFileByServerRelativeUrl('" +
				(parameters.includeBaseUrl == false ? "" : this.baseUrl) + parameters.path + "')" +
				((parameters.recycle && parameters.recycle == false as boolean) ? "" : "/recycle()"),
			method: "POST"
		});
	}


/*******************************************************************************
 *     LIST PROPERTY / STRUCTURE REQUESTS: fields, content types, views
 *******************************************************************************/

/****************  Fields  ****************************/
	/**
	 * @method getField -- returns one or more field objects, with properties all or selected, and possible
	 *     filtering of results
	 * @param {object} parameters -- object takes two properties or can be undefined
	 *			if parameters is null or undefined, it will return all fields and their properties for the list
	 * 		 fields: an array of objects or null or empty array, where object {name: <field-name>} or {guid: <field-guid>}
	 * 	    query: an object with the following properties: "select", "filter", "expand" which operate as O-Data
	 *               query is optional
	 * @returns -- will return to then() an array of results representing fields and their properties
	 */
	getField(parameters?: {
		fields?: {name?: string; guid?: string;}[];
		filter?: string;
	} | undefined) {
		let query = "", filter;

		if (parameters && parameters.fields != null && parameters.fields.length > 0) {
			if (parameters.filter)
				filter = parameters.filter;
			else
				filter = "";
			for (let fld of parameters.fields) {
				if (filter.length > 0)
					filter += " or ";
				if (fld.name)
					filter += "Title eq '" + fld.name + "'";
				else
					filter = "Id eq '" + fld.guid + "'";
			}
			parameters.filter = filter;
		}
	//	query = constructQueryParameters(parameters);
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/fields" + query
		});
	}

	/**
	 * @method getFields -- will return information on all list fields
	 * @returns Promise with data in then() response
	 */
	getFields(parameters?: {filter: string;}): Promise<any> {
		return this.getField(parameters);
	}

	/**
	 *
	 * @param {JSONofParms} fieldProperties 'Title': 'field title', 'FieldTypeKind': FieldType value,
	 *    Required': 'true/false', 'EnforceUniqueValues': 'true/false','StaticName': 'field name'}
	 * @returns REST API response
	 */
	createField(fieldProperties: {
		"Title": string;
		"FieldTypeKind": number;
		"Required": boolean;
		"EnforceUniqueValues": boolean;
		"StaticName": string;
	}): Promise<any> {
		let body: SPRESTTypes.THttpRequestBody | string = fieldProperties;

		if (SPRESTSupportLib.checkEntityTypeProperty(body, "field") == false)
			body.__SetType__ = "SP.Field";
		body = SPRESTSupportLib.formatRESTBody(body);
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/fields",
			method: "POST",
			body: body
		});
	}

	/**
	 *
	 * @param {arrayOfFieldObject} fields -- this should be an array of the
	 * 			       objects used to define fields for SPListREST.createField()
	 * @returns
	 */
	createFields(fields: any[]) {
		// field creation can not occur co-synchronously. One must complete before the
		// next field creation can begin.
		return new Promise((resolve, reject) => {
			let body: SPRESTTypes.THttpRequestBody,
				requests: SPRESTTypes.IBatchHTTPRequestForm[] = [],
				responses: string = "";

			for (let fld of fields) {
				body = fld;
				if (SPRESTSupportLib.checkEntityTypeProperty(body, "field") == false)
					body.__SetType__ = "SP.Field";
				requests.push({
					url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/fields",
					method: "POST",
					body: JSON.parse(SPRESTSupportLib.formatRESTBody(body))
				});
			}
			SPRESTSupportLib.batchRequestingQueue({
				host: this.server,
				path: this.site,
				AllHeaders: SPListREST.stdHeaders,
				AllMethods: "POST"
			}, requests).then((response: any) => {
				resolve(response);
			}).catch((response: any) => {
				reject(response);
			});
/*
			serialSPProcessing(this.createField, fields).then((response) => {
				resolve(response);
			}).catch((response) => {
				reject(response);
			});  */
		});
	}

	renameField(parameters: {
		oldName?: string;
		currentName?: string;
		newName: string | undefined;
	}) {
		let body: SPRESTTypes.THttpRequestBody;

		if (!parameters.oldName && !parameters.currentName)
			throw "A parameter for 'oldName' or 'currentName' was not found";
		if (parameters.oldName)
			parameters.currentName = parameters.oldName;
		if (!parameters.newName)
			throw "A parameters for 'newName' was not found";
		body = {
			"Title": parameters.newName
		};
		if (SPRESTSupportLib.checkEntityTypeProperty(body, "field") == false)
			body["__SetType__"] = "SP.Field";
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid +
					"')/fields/GetByTitle('" + parameters.currentName + "')",
			method: "POST",
			headers: {
				"X-HTTP-METHOD": "MERGE",
				"IF-MATCH": "*"
			},
			body: SPRESTSupportLib.formatRESTBody(body)
		});
	}

	/** @method .updateChoiceTypeField
	 *       @param {Object} parameters
	 *       @param {string} parameters.id - the field guid/id to be updated
	 *       @param {(array|arrayAsString)} parameters.choices - the elements that will
	 *     form the choices for the field, as an array or array written as string
	 */
	updateChoiceTypeField (parameters: {id: string; choices?: string[]; choice?: string}) {
		let choices: string;
		if (Array.isArray(parameters.choices) == true) {
			choices = JSON.stringify(parameters.choices);
			choices = choices.replace(/"/g, "'");
		}
		else if (typeof parameters.choice == "string")
			choices = parameters.choice.replace(/"/g, "'");
		else
			throw "parameters must include '.choices' property that is " +
				"either array or string";
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" +
						this.listGuid + "')/fields(guid'" + parameters.id + "')",
			method: "POST",
			headers: {
				"IF-MATCH": "*",
				"X-HTTP-METHOD": "MERGE"
			},
			body: "{ '__metadata': { 'type': 'SP.FieldChoice' }, " +
				"'Choices': { '__metadata': { 'type': 'Collection(Edm.String)' }, " +
					"'results': " + choices + "} }"
		});
	}

	/****************  Content Types  ****************************/
	/**
	 * @method getContentType -- returns one or more contentType objects, with properties all or selected, and possible
	 *     filtering of results
	 * @param {object} parameters -- object takes two properties or can be undefined
	 *			if parameters is null or undefined, it will return all fields and their properties for the list
	 * 		 contentTypes: an array of objects or null or empty array, where object {name: <content type-name>} or {guid: <field-guid>}
	 * 	    query: an object with the following properties: "select", "filter", "expand" which operate as O-Data
	 *               query is optional
	 * @returns -- will return to then() an array of results representing content types and their properties
	 */
	 getContentType(parameters: {
		 contentTypes?: {name?: string; guid?: string}[];
		 query?: {filter: string;}
	 } | undefined): Promise<any> {
		let query = "", filter;

		if (parameters && parameters.contentTypes != null && parameters.contentTypes.length > 0) {
			if (parameters.query && parameters.query.filter)
				filter = parameters.query.filter;
			else
				filter = "";
			for (let ct of parameters.contentTypes) {
				if (filter.length > 0)
					filter += " or ";
				if (ct.name)
					filter += "Title eq '" + ct.name + "'";
				else
					filter = "Id eq '" + ct.guid + "'";
			}
		}
		if (parameters && parameters.query)
			query = SPRESTSupportLib.constructQueryParameters(parameters.query);
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists(guid'" +
					this.listGuid + "')/contentTypes" + query,
			method: "GET"
		});
	}

	/**
	 *
	 * @param {arrayOfcontentTypeObject} parameters -- this should be an array of the
	 * 			       object parameter passed to .createcontentType
	 * @returns
	 */
/*/
	createContentTypes(contentTypes) {
		// contentType creation can not occur co-synchronously. One must complete before the
		// next contentType creation can begin.

		return new Promise((resolve, reject) => {
			serialSPProcessing(this.createContentType, contentTypes).then((response) => {
				resolve(response);
			}).catch((response) => {
				reject(response);
			});
		});
	}
/**/
	/****************  Views  ****************************/
	/**
	 * @method getView -- returns one or more contentType objects, with properties all or selected, and possible
	 *     filtering of results
	 * @param {object} parameters -- object takes two properties or can be undefined
	 *			if parameters is null or undefined, it will return all fields and their properties for the list
	 * 		 contentTypes: an array of objects or null or empty array, where object {name: <content type-name>} or {guid: <field-guid>}
	 * 	    query: an object with the following properties: "select", "filter", "expand" which operate as O-Data
	 *               query is optional
	 * @returns -- will return to then() an array of results representing content types and their properties
	 */
	 getView(parameters?: {
		 	views?: {
				name?: string;
				guid?:string;
			}[],
			query?: {
				filter: string;
			}}): Promise<any> {
		let query:string = "",
			filter: string;

		if (parameters && parameters.views != null && parameters.views.length > 0) {
			if (parameters.query!.filter)
				filter = parameters.query!.filter;
			else
				filter = "";
			for (let view of parameters.views) {
				if (filter.length > 0)
					filter += " or ";
				if (view.name)
					filter += "Title eq '" + view.name + "'";
				else
					filter = "Id eq '" + view.guid + "'";
			}
		}
		if (parameters && parameters.query)
			query = SPRESTSupportLib.constructQueryParameters(parameters.query);
		return SPListREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists(guid'" +
					this.listGuid + "')/views" + query,
			method: "GET"
		});
	}

	/**
	 *
	 * @param {JSONofParms} viewProperties }
	 * @returns REST API response
	 */
	createView(viewsProperties: SPRESTTypes.THttpRequestBody): Promise<any> {
		let body: SPRESTTypes.THttpRequestBody | string = viewsProperties;

		if (SPRESTSupportLib.checkEntityTypeProperty(body, "view") == false)
			body.__SetType__ = "SP.View";
		body = SPRESTSupportLib.formatRESTBody(body);
		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists(guid'" + this.listGuid + "')/views",
			method: "POST",
			body: body
		});
	}

	setView(viewName: string, fieldNames: string[]) {
		let viewFields = "";

		for (let field of fieldNames)
			viewFields += "<FieldRef Name=\"" + field + "\"/>";

		return SPListREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/GetList(@a1)/Views(@a2)/SetViewXml()?@a1='" +
					this.serverRelativeUrl + "'&@a2='" + this.listGuid + "'",
			method: "POST",
			body: {"ViewXml": "<View Name=\"{" + SPRESTSupportLib.createGuid() + "}\" DefaultView=\"FALSE\" MobileView=\"TRUE\" " +
					"MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" " +
					"Url=\"" + this.serverRelativeUrl + "/" + viewName + ".aspx\" Level=\"1\" " +
					"BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=47\">" +
					"<Query><OrderBy><FieldRef Name=\"ID\" Ascending=\"FALSE\"/></OrderBy></Query>" +
					"<ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"Created\"/><FieldRef Name=\"Modified\"/>" +
					viewFields + "</ViewFields>" +
					"<RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink>" +
					"<Toolbar Type=\"Standard\"/><Aggregations Value=\"Off\"/></View>"} as any
		});
	}
	/**
	 *
	 * @param {arrayOfFieldObject} views -- this should be an array of the
	 * 			       objects used to define views for SPListREST.createViews()
	 * @returns
	 */
	createViews (views: any[]) {
		// field creation can not occur co-synchronously. One must complete before the
		// next field creation can begin.
		return new Promise((resolve, reject) => {
			SPRESTSupportLib.serialSPProcessing(this.createView, views).then((response: any) => {
				resolve(response);
			}).catch((response: any) => {
				reject(response);
			});
		});

/*
		return new Promise((resolve, reject) => {
			function recurse(index) {
				let viewProperties;

				if ((viewProperties = views[index]) == null) {
					if (index > 0)
						return true;
					else
						resolve(true);
				}
				this.createView(viewProperties).then(() => {
					if (recurse(index + 1) == true) {
						if (index > 1)
							return true;
						else
							resolve(true);
					}
				}).catch((response) => {
					reject(response);
				});
			}
			recurse(0);
		});	*/
	}
}

