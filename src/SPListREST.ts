/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unused-vars */
"use strict";

/* jshint -W119, -W069, -W083 */

import {
	TLookupFieldInfo,
	TListSetup,
	TSpListItem,
	TSiteInfo,
	TSpField,
	//	THttpRequestParams,
	//	THttpRequestParamsWithPromise,
	TSPDocLibCheckInType,
	//		IBatchHTTPRequestParams,
	SPListRaw,
	SPListItemRaw,
	SPContentTypeRaw,
	SPFileRaw,
	SPFieldRaw,
	SPRootFolderRaw
} from "./SPComponentTypes.d";
import { RESTrequest, batchRequestingQueue, SPstdHeaders } from "./SPHttpReqResp";
import { THttpRequestBody, TSPResponseData, TIntervalControl, HttpInfo,
	TXmlHttpRequestData, TFetchInfo, IBatchHTTPRequestForm, TArrayToJSON, HttpStatusCode
} from "./SPHttp";
import { SPSiteREST } from "./SPSiteREST";
import {
	isIterable,
	ParseSPUrl,
	constructQueryParameters,
//	formatRESTBody,
	serialSPProcessing,
	checkEntityTypeProperty,
	createGuid
} from "./SPRESTSupportLib";

export { SPListREST };

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

class SPListREST {
	protocol: string;
	server: string;
	site: string;
	forApiPrefix: string;
	forApiPrefix: string;
	listName: string = "";
	listGuid: string = "";
	baseUrl: string = "";
	serverRelativeUrl: string = "";
	creationDate: Date = new Date(0);
	baseTemplate: number = -1;
	allowContentTypes: boolean = Boolean(null);
	itemCount: number = -1;
	listItemEntityTypeFullName: string = "";
	rootFolder: string = "";
	contentTypes: object = {};
	listIds: number[] = [];
	fields: TSpField[] = [];
	fieldInfo: any[] = [];
	lookupFieldInfo: TLookupFieldInfo[] = [];
	lookupInfoCached: boolean = Boolean(null);
	linkToDocumentContentTypeId: string = "";
	currentListIdIndex: number = -1;
	setup: TListSetup;
	sitePedigree: TSiteInfo = {} as TSiteInfo;
	sitePedigree: TSiteInfo = {} as TSiteInfo;
	arrayPedigree: TSiteInfo[] = [];

	/**
	 * @constructor
	 * @param {object:
	 *	server: {string} server domain, equivalent of location.hostname return value
	* 	site: {string}  URL string part to the site
	*		list: {string}  must be a valid list name. To create a list, use SPSiteREST() object
	*    include: {string}  comma-separated properties to be part of OData $expand for more properties
	*	} setup -- object to initialize the
	*/
	constructor(setup: TListSetup) {
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
		if (this.listGuid && (matches = this.listGuid.match(/\{?([\d\w-]+)\}?/)) != null)
			this.listGuid = matches[1];
		if (!(this.listName || this.listGuid))
			throw "Neither a list name or list GUID were provided to identify a list";
		this.setup = setup;
		this.forApiPrefix = this.protocol + this.server + this.site + "/_api";
		/*
		if (setup.linkToDocumentContentTypeId)
			this.linkToDocumentContentTypeId = setup.linkToDocumentContentTypeId; */
	}
	static escapeApostrophe (string: string): string {
		return encodeURIComponent(string).replace(/'/g, "''");
	}

	/**
	 * @method _init -- sets many list characteristics whose data is required
	 * 	by asynchronous request
	 * @returns Promise
	 */
	init(): Promise<boolean> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/lists" +
					(this.listName ? "/getByTitle('" + this.listName + "')" :
						"(guid'" + this.listGuid + "')") +
					"?$expand=RootFolder,ContentTypes,Fields" +
					(this.setup && this.setup.include ? "," + this.setup.include : ""),
				method: "GET",
				successCallback: (data/*, text, reqObj */) => {
					let loopData: unknown[];
					//	lookupFieldTypeNum = SPFieldTypes.getFieldTypeIdFromTypeName("Lookup");


					const d: SPListRaw = data.d as SPListRaw;
					this.baseUrl = this.server + this.site;
					this.serverRelativeUrl = (d.RootFolder as SPRootFolderRaw).ServerRelativeUrl;
					this.creationDate = new Date(d.Created);
					this.baseTemplate = d.BaseTemplate;
					this.allowContentTypes = d.AllowContentTypes;
					this.listGuid = d.Id;
					this.listName = d.Title;
					this.itemCount = d.ItemCount;
					this.listItemEntityTypeFullName = d.ListItemEntityTypeFullName;
					this.rootFolder = (d.RootFolder as SPRootFolderRaw).Name;
					this.contentTypes = d.ContentTypes.results;
					this.fields = d.Fields.results.map((elem: any) => {
						const field: TSpField = {} as TSpField;
						for (const fieldProp in field)
							if (fieldProp in elem)
								(field as any)[fieldProp] = elem[fieldProp];
						return field;
					});
					if (this.setup.include) {
						const components: string[] = this.setup.include.split(",");

						for (const component of components)
							//this[component] = data[component];
							Object.defineProperty(this, (d as any)[component], component);
					}
					// for doc libs, look for "Link To Document" content type and store its content type ID
					// one more trip to the network to get any lookup field information
					this.getLookupFieldsInfo().then(() => {
						if (//

							isIterable(loopData = d.ContentTypes.results) == true) {
							let i: number,

								loopData: any;

							for (i = 0; i < loopData.length; i++)
								if (loopData[i].Name == "Link to a Document")
									break;
							if (i < loopData.length)
								this.linkToDocumentContentTypeId = loopData[i].Id.StringValue;
						}
						resolve(true);
					}).catch((err: unknown) => {
						reject("List info error\n" + JSON.stringify(err, null, "  "));
					});

				},
				errorCallback: (reqObj/*, status, errThrown */) => {
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
	getLookupFieldsInfo(): Promise<boolean> {
		return new Promise<boolean>((resolve, reject) => {
			// first get the site pedigree of current site
			const siteREST = new SPSiteREST({
				server: this.server,
				site: this.site
			});
			siteREST.init().then(() => {
				siteREST.getSitePedigree(null).then(
					({pedigree: pedigree, arrayPedigree: arrayPedigree}) => {
						// after site pedigree, collect the lookup fields and lists as
						const lookupFields: TLookupFieldInfo[] = [],
							requests: Promise<unknown>[] = [];

						this.sitePedigree = pedigree;
						this.arrayPedigree = arrayPedigree;
						for (const field of this.fields)
							if (field.FieldTypeKind == 7 && field.FromBaseType == false) // lookup type, so get the data
								lookupFields.push(field as unknown as TLookupFieldInfo);

						for (const field of lookupFields)
							requests.push(this.retrieveLookupFieldInfo({
								useFunction: "siteREST",
								LookupWebId: field.SPproperties.LookupWebId,
								LookupList: field.SPproperties.LookupList,
								LookupField: field.SPproperties.LookupField
							}));


						Promise.all(requests).then((responses: any) => {
						//console.log("responses count = " + responses.length + "\n" +
						//	JSON.stringify(responses, null, "  "));
							for (const i of responses) {

								const data: any[] = i.data,
									choices: {id: number; value: string }[] = [];

								let fld: any;

								if (Array.isArray(data) == false)
									continue;
								found1:
								for (let i = 0; i < lookupFields.length; i++)
									for (const datum of data)
										if (datum[lookupFields[i].SPproperties.LookupField] != null) {
											fld = lookupFields.splice(i, 1)[0];
											break found1;
										}
								const idx = this.lookupFieldInfo.push({
									fieldDisplayName: fld.Title,
									fieldInternalName: fld.InternalName,
									fieldLookupFieldName: fld.LookupField,
									choices: null
								} as TLookupFieldInfo) - 1;
								for (const fldVal of data) {
									if (fldVal[fld.LookupField] == null)
										break;
									choices.push({
										id: fldVal.Id,
										value: fldVal[fld.LookupField]
									});
								}
								this.lookupFieldInfo[idx].choices = choices;
							}
							resolve(true);
						}).catch((err: unknown) => {
							reject(err);
						});

					}).catch((err: unknown) => {
					reject(err);
				});
			}).catch((err: unknown) => {
				reject(err);
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
		useFunction?: "SPSearch" | "siteREST";
		LookupWebId: string;
		LookupList: string;
		LookupField: string;
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			if (parameters.useFunction == "siteREST") {
				let site: TSiteInfo = {} as TSiteInfo;
				const lookupListId = parameters.LookupList.replace(/[{}]/g, "");

				for (site of this.arrayPedigree)
					if (site.Id == parameters.LookupWebId)
					if (site.Id == parameters.LookupWebId)
						break;
				RESTrequest({
					url: this.protocol + this.server + site.ServerRelativeUrl + "/_api/lists(guid'" +
							lookupListId + "')/items?$select=Id," + parameters.LookupField,
					successCallback: (data, httpInfo) => {
						resolve(data);
					},
					errorCallback: (err, httpInfo) => {
						reject(err);
					}
				});
			} else if (parameters.useFunction == "SPSearch") // no useFunction specified
				RESTrequest({
					url: this.forApiPrefix + "/_api/search/query?querytext=guid'" + parameters.LookupWebId + "'",
					successCallback: (data, httpInfo) => {
					// the data returned are the lookup choice values in the target
						const d = data.d;
						let siteName,
							lookupList,
							siteResult = null;
						const lookupChoices: {
							id: string;  // item.Id
							choiceValue: string  // item[parameters.LookupField]
						}[] = [],

							searchRows = (d!.query as any)!.PrimaryQueryResult.RelevantResults.Table.Rows.results,
							// this variabled function will get the choices from the lookup field
							collectItemData = (url: string) => {
							// this named function will be used
								RESTrequest({
									url: url,
									successCallback: (data, httpInfo) => {
										const d = data.d;

										for (const item of d!.results!)
											lookupChoices.push({

												id: (item as any).Id,

												choiceValue: (item as any)[parameters.LookupField]
											});
										if (d!.__next)
											collectItemData(d!.__next);
										else
											resolve(lookupChoices);

									},
									errorCallback: (err, info) => {
										reject(err);

									}
								});
							};
						// This complex multiple loop is trying to find the results in the complex response
						//   of the SharePoint search query
						for (const searchRow of searchRows) {
							for (const searchCell of searchRow.Cells.results) {
								if (searchCell.Key == "WebId" && searchCell.Value == parameters.LookupWebId) {
									for (const searchCell2 of searchRow.Cells.results)
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
							siteName = //
						ParseSPUrl(siteResult)!.siteFullPath;
							lookupList = parameters.LookupList.match(/^\{?([^}]+)\}?$/)![1];
							collectItemData(this.protocol + this.server + siteName + "/_api/web/lists(guid'" + lookupList +
								"')/items?$select=Id," + parameters.LookupField);
						}
						resolve(null); // could not find a result in search
					},
					errorCallback: (err, info) => {
						reject("REST Search request failed");
					}
				});
		});
	}

	pullLookupFieldInfo(parameters: {
		displayName?: string;
		internalName?: string;
		lookupFieldName?: string;
	}): TLookupFieldInfo | null {
		for (const fldInfo of this.lookupFieldInfo)
			if (parameters.displayName && parameters.displayName == fldInfo.fieldDisplayName)
				return fldInfo;
			else if (parameters.internalName && parameters.internalName == fldInfo.fieldInternalName)
				return fldInfo;
			else if (parameters.lookupFieldName && parameters.lookupFieldName == fldInfo.fieldLookupFieldName)
				return fldInfo;
		return null;
	}

	/**
	 * @method getLookupFieldValue -- retrieves as a Promise the value to set up for a lookup field
	 * @param {string} fieldName --
	 * @returns the value to set for the lookup field, null if not available
	 */
	getLookupFieldValue(
		fieldName: string,
		fieldValue: string
	): number | null {
		for (const info of this.lookupFieldInfo)
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
	fixBodyForSpecialFields(body: THttpRequestBody): Promise<unknown> {
		let newBody: {[key:string]: unknown},
			value: number | null;

		return new Promise((resolve, reject) => {
			if (this.lookupInfoCached == true) {
				// this is a quick way to get lookup value without going back to the server
				newBody = { };
				for (const key in body)
					if ((value = this.getLookupFieldValue(key, body[key] as string)) == body[key])
						newBody[key] = value;
					else
						newBody[key + "Id"] = value;
				resolve(newBody);
			} else
			// with this block, we go to the server to get lookup values
				this.getFields({
					filter: "FromBaseType eq false"

				}).then((response: any) => {

					const requests: any = [ ];

					// get the lists fields and store the info
					if (this.fields)
						this.fields.concat(response.data.d.results);
					else
						this.fields = response.data.d.results;
					//this.fieldInfo = [ ];
					// this will be multiple requests, so find Promise.all()
					for (const fld of response.data.d.results)
						requests.push(new Promise((resolve, reject) => {
							// lookup type field
							if (fld.FieldTypeKind == 7 && typeof fld.LookupWebId == "string" &&
										typeof fld.LookupList == "string" && typeof fld.LookupField == "string" &&
										fld.LookupList.length > 0 && fld.LookupField.length > 0) { // lookup
								this.retrieveLookupFieldInfo({
									LookupWebId: fld.LookupWebId,
									LookupList: fld.LookupList,
									LookupField: fld.LookupField

								}).then(response => {
									this.lookupFieldInfo.push({
										fieldDisplayName: fld.Title,
										fieldInternalName: fld.InternalName,
										fieldLookupFieldName: fld.LookupField,
										choices: response
									} as TLookupFieldInfo);
								}).catch((err: unknown) => {
									reject(err);
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

						for (const key in body)
							if (Array.isArray(body[key])) {
								for (const fld of this.fields)
									if (fld.InternalName == key) {
										newBody[key] = {
											__metadata: {
												type: fld.Choices!.__metadata.type
											},
											results: body[key]
										};
										break;
									}
							} else if (typeof body[key] == "boolean" ||
										(typeof body[key] == "string" && (body[key] as string).search(/Y|Yes|No|N|true|false/i) == 0)) {
								for (const fld of this.fieldInfo)
									if (fld.intName == key)
										if (fld.boolean)
											if (body[key] == true || (body[key] as string).search(/Y|Yes|true/i) == 0)
												newBody[key] = true;
											else
												newBody[key] = false;
										else
											newBody[key] = body[key];
							} else if ((value = this.getLookupFieldValue(key, body[key] as string)) == body[key])
								newBody[key] = value; // this is the case
							else
								newBody[key + "Id"] = value;
						resolve(newBody);
					}).catch((err: unknown) => {
						reject(err);
					});
				}).catch(() => {
					reject(".getFields() called failed");
				});
		});
	}

	// query: [optional]
	getProperties (parameters: unknown): Promise<unknown> {
		return new Promise((resolve, reject) => {
			const query: string = //
					constructQueryParameters(parameters);

			RESTrequest({
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')" + query,
				method: "GET",
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	getListProperties(parameters: unknown): Promise<unknown> {
		return this.getProperties(parameters);
	}

	getListInfo(parameters: unknown): Promise<unknown> {
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
		progressReport?: TIntervalControl,
		[key: string]: unknown;
	}): Promise<TSpListItem[]> {
		return new Promise((resolve, reject) => {
			let query: string = "",
				filter: string = "",
				select: string = "",
				expand: string = "";

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
			if (parameters.selectDisplay && parameters.selectDisplay.length > 0) {
				for (const displayName of parameters.selectDisplay) {
					let found = false;
					for (const lookupField of this.lookupFieldInfo) {
						if (lookupField.fieldDisplayName == displayName) {
							expand += (expand.length > 0 ? "," : "") + lookupField.fieldInternalName;
							select += (select.length > 0 ? "," : "") + lookupField.fieldInternalName + "/" + lookupField.fieldLookupFieldName;
							found = true;
							break;
						}
					}
					if (found == false)
						select += (select.length > 0 ? "," : "") + displayName;
				}
				select += (parameters.select && parameters.select.length > 0) ? "," : "";
			}
			if (parameters.select && parameters.select.length > 0)
				query = "?$select=" + select + parameters.select;
			if (parameters.expand && parameters.expand.length > 0) {
				if (expand.length > 0)
					expand += "," + parameters.expand;
				else
					expand = parameters.expand;
				if (query.length > 0)
					query += "&$expand=" + expand;
				else
					query = "?$expand=" + expand;
			}
			if (filter.length > 0)
				if (query.length > 0)
					query += "&$filter=" + filter;
				else
					query = "?filter=" + filter;
			RESTrequest({
				// SPListItemRaw type returned
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/items" +
						(parameters && parameters.itemId > 0 ? "(" + parameters.itemId + ")" : "") + query,
				//			progressReport: parameters.progressReport,
				successCallback: (response, httpInfo) => {
					if (response.d?.results)
						resolve(response.d.results as TSpListItem[]);
					else if (response.d)
						resolve([response.d as TSpListItem]);
					else
						resolve([]);
				},
				errorCallback: (errInfo, info) => {
					reject(errInfo);
				}
			});
		});
	}

	getListItem(parameters: any): Promise<TSpListItem | null> {
		return new Promise<TSpListItem | null>((resolve, reject) => {
			this.getListItemData(parameters).then((response: TSpListItem[] | TSpListItem | null) => {
				if (Array.isArray(response) == true)
					response = null;
				resolve(response as TSpListItem | null);
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}

	getItemData(parameters: any): Promise<TSpListItem[] | TSpListItem | null> {
		return this.getListItemData(parameters);
	}

	getAllListItems(): Promise<TSpListItem[]> {
		return new Promise((resolve, reject) => {
			this.getListItemData({itemId: -1}).then((response) => {
				if (response != null && Array.isArray(response) == false)
					response = [];
				resolve(response);
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}

	getListItemsWithQuery(parameters: {
		select: string | null,
		expand: string | null,
		filter: string | null,
		selectDisplay?: string[] // this must be used to get columns/fields data with that display name
		progressReport?: TIntervalControl
	}): Promise<TSpListItem[] | TSpListItem | null> {
		const newParameters = {
			select: parameters.select,
			expand: parameters.expand,
			filter: parameters.filter,
			selectDisplay: parameters.selectDisplay,
			progressReport: parameters.progressReport,
			itemId: -1
		};
		return this.getListItemData(newParameters);
	}

	getAllListItemsOptionalQuery(parameters: any): Promise<TSpListItem[]> {
		return this.getListItems(parameters);
	}

	getListItems(parameters: any): Promise<TSpListItem[]> {
		return new Promise<TSpListItem[]>((resolve, reject) => {
			this.getListItemData(parameters).then((response) => {
				resolve(response);
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}

	// @param {object} parameters - should have at least {body:,success:}
	// body format: {string} " 'fieldInternalName': 'value', ...}
	createListItem(item: {body: THttpRequestBody}): Promise<HttpStatusCode> {
		return new Promise<HttpStatusCode>((resolve, reject) => {
			const body: THttpRequestBody | string = item.body;

			if (!body || body == null)
				throw "The object argument to createListItem() must have a 'body' property";
			//		if (this.checkEntityTypeProperty(body, "item") == false)
			//			body["__SetType__"] = this.listItemEntityTypeFullName;
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/items",
				method: "POST",
				body: body as TXmlHttpRequestData,
				successCallback: (data) => {
					resolve(data.status as HttpStatusCode);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	/**
	 *
	 * @param {arrayOfObject} items -- objects should be objects whose properties are fields
	 * 		within the lists and using Internal name property of field
	 * @returns a promise
	 */
	createListItems(items: {[key:string]: string}[], useBatch: boolean):
				Promise<{success: TFetchInfo[]; error: TFetchInfo[]}> | Promise<unknown> {
		if (useBatch == true) {
			const requests: IBatchHTTPRequestForm[] = [];
			// need to re-configure
			for (const item of items)
				requests.push({
					// type SPListItemRaw[] // need specified entity name
					url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/items",
					contextinfo: this.forApiPrefix,
					method: "POST",
					headers: {
						"Content-Type": "application/json;odata=verbose"
					},
					body: item
				});
			return batchRequestingQueue(requests);
		}
		// the alternative to batching
		return new Promise((resolve, reject) => {
			serialSPProcessing(
				this.createListItem as ((arg1: unknown) => Promise<unknown>),
				items
			).then((response: unknown) => {
				resolve(response);
			}).catch((err: unknown) => {
				reject(err);
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
		body: TSpListItem;
	}): Promise<number> {
		const body: THttpRequestBody | string = parameters.body;

		return new Promise<HttpStatusCode>((resolve, reject) => {
			this.fixBodyForSpecialFields(body as THttpRequestBody)
				.then((body: any) => {
					if (checkEntityTypeProperty(body, "item") == false)
						body["__SetType__"] = this.listItemEntityTypeFullName;
					RESTrequest({
					// SPListItem
						setDigest: true,
						url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid +
								"')/items(" + parameters.itemId + ")",
						method: "POST",
						headers: {
							...SPstdHeaders,
							"X-HTTP-METHOD": "MERGE",
							"IF-MATCH": "*" // can also use etag
						},
						body: body,
						successCallback: (data) => {
							resolve(data.status as HttpStatusCode);
						},
						errorCallback: (err, headers) => {
							reject(err);
						}
					});
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	updateListItems(
		itemsArray: {[key:string]:string;}[],
		useBatch: boolean
	): Promise<{success: TFetchInfo[]; error: TFetchInfo[]} | null> {
		if (useBatch == true) {
			const requests: IBatchHTTPRequestForm[] = [];
			// need to re-configure
			for (const item of itemsArray)
				requests.push({
					url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/items(" + item.id + ")",
					contextinfo: this.forApiPrefix,
					headers: {
						"Content-Type": "application/json;odata=verbose",
						"IF-MATCH": "*"
					},
					method: "PATCH",
					body: item
				});
			return batchRequestingQueue(requests);
		}
		// the alternative to batching
		return new Promise((resolve, reject) => {
			serialSPProcessing(
				this.updateListItems as ((arg1: unknown) => Promise<unknown>),
				itemsArray
			)
				.then((response: any) => {
					resolve(response as {success: TFetchInfo[]; error: TFetchInfo[]} | null);
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	/**
	applies to all lists
	parameters properties should be
	.itemId [required]
	.recycle [optional, default=true]. Use false to get hard delete
	*/
	deleteListItem(parameters: {
		itemId: number;
		recycle: boolean;
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			const recycle = parameters.recycle;

			if (!recycle)
				throw "Parameter 'recycle' is required. If 'true', item goes to recycle bin. " +
						"If 'false', item is permanently deleted.";
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/lists" +
						(this.listName ?
							"/GetByTitle('" + this.listName + "')" : "(guid'" + this.listGuid + "')") +
						"/items(" + parameters.itemId + ")" + (recycle == true ? "/recycle()" : ""),
				method: recycle == true ? "POST" : "DELETE",
				headers: recycle == true ? { "IF-MATCH": "*" } : { "X-HTTP-METHOD": "DELETE"},
				successCallback: (data) => {

				},
				errorCallback: (err) => {}
			});
		});
	}

	getListItemCount (doFresh: boolean = false): Promise<number> {
		return new Promise<number>((resolve, reject) => {
			if (doFresh == false)
				resolve(this.itemCount);
			else
				RESTrequest({
					//  SPListRaw
					url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')",
					successCallback: (data: TSPResponseData) => {
						resolve((data.d as SPListRaw).ItemCount);
					},
					errorCallback: (errInfo) => {
						reject(errInfo);
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
					}).then((response: TSpListItem[] | null) => {
						if (response)
							for (const result of response)
								this.listIds.push(result.Id);
						resolve(this.listIds);
					}).catch((err: unknown) => {
						reject(err);
					});
				}).catch((err: unknown) => {
					reject(err);
				});
		});
	}

	isValidID(itemId: number): Promise<unknown> {
		return new Promise((resolve, reject) => {
			this.getListItemIds().then((response: number[]) => {
				resolve(response.includes(itemId));
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}

	getListItemIds(): Promise<number[]> {
		return new Promise((resolve, reject) => {
			this.loadListItemIds().then(() =>  {
				resolve(this.listIds);
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}

	getNextListId(): Promise<unknown> {
		return new Promise((resolve, reject) => {
			this.loadListItemIds().then( () =>  {
				if (this.currentListIdIndex == this.listIds.length)
					resolve(this.listIds[this.currentListIdIndex]);
				else
					resolve(this.listIds[++this.currentListIdIndex]);
			}).catch((err: unknown) => {
				reject (err);
			});
		});
	}

	getPreviousListId() {
		return new Promise((resolve, reject) => {
			this.loadListItemIds()
				.then(() => {
					if (this.currentListIdIndex == 0)
						resolve(this.listIds[this.currentListIdIndex]);
					else
						resolve(this.listIds[--this.currentListIdIndex]);
				}).catch((err: unknown) => {
					reject (err);
				});
		});
	}

	getFirstListId(): Promise<unknown> {
		return new Promise((resolve, reject) => {
			this.loadListItemIds().then(() => {
				resolve(this.listIds[this.currentListIdIndex = 0]);
			}).catch((response: unknown) => {
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
	getAllDocLibFiles(): Promise<SPFileRaw[]> {
		return new Promise<SPFileRaw[]>((resolve, reject) => {
			RESTrequest({
				// SPFile[]
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/Files",
				successCallback: (data) => {
					resolve(data as SPFileRaw[]);
				},
				errorCallback: (errInfo) => {
					reject(errInfo);
				}
			});
		});
	}

	getFolderFilesOptionalQuery(parameters: {
		folderPath: string;
	}): Promise<unknown> {
		const query: string = constructQueryParameters(parameters);
		return new Promise((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/getFolderByServerRelativeUrl('" + this.baseUrl +
							parameters.folderPath + "')/Files" + query,
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	/** @method getDocLibItemByFileName
	retrieves item by file name and also returns item metadata
	* @param {Object} parameters
	* @param {string} parameters.fileName - name of file
	*/
	getDocLibItemByFileName(fileName:string): Promise<TSpListItem> {
		return new Promise<TSpListItem>((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/getFolderByServerRelativeUrl('" +
							this.baseUrl + "')/Files('" + SPListREST.escapeApostrophe(fileName) +
					"')?$expand=ListItemAllFields",
				successCallback: (data) => {
					resolve(data as TSpListItem);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	/** @method getDocLibItemMetadata
	 * @param {Object} parameters
	 * @param {number|string} parameters.itemId - ID of item whose data is wanted
	 * @returns  {Object} returns only the metadata about the file item
	 */
	getDocLibItemMetadata(parameters: {itemId:number}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			if (!parameters.itemId)
				throw "Method requires 'parameters.itemId' to be defined";
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid +
							"')/items(" + parameters.itemId + ")",
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err, info) => {
					reject(err);
				}
			});
		});
	}
	/** @method getDocLibItemFileAndMetaData
	 * @param {Object} parameters
	 * @param {number|string} parameters.itemId - ID of item whose data is wanted
	 * @returns {Object} data about the file and metadata of the library item
	 */

	getDocLibItemFileAndMetaData(parameters: {itemId: number}): Promise<SPFileRaw> {
		return new Promise<SPFileRaw>((resolve, reject) => {
			if (!parameters.itemId)
				throw "Method requires 'parameters.itemId' to be defined";
			RESTrequest({
				// SPFileRaw with ListItemAllFields Expanded (SP.Data.listnameItem)
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid +
						"')/items(" + parameters.itemId + ")/File?$expand=ListItemAllFields",
				successCallback: (data) => {
					resolve(data as SPFileRaw);
				},
				errorCallback: (err, info) => {
					reject(err);
				}
			});
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
		body: TXmlHttpRequestData;
		folderPath: string;
		willOverwrite: boolean;
	}): Promise<HttpStatusCode> {
		return new Promise<HttpStatusCode>((resolve, reject) => {
			let path: string = this.serverRelativeUrl;
			if (parameters.folderPath.charAt(0) == "/")
				path += parameters.folderPath;
			else
				path += "/" + parameters.folderPath;
			if (typeof parameters.willOverwrite == "undefined")
				parameters.willOverwrite = false;
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/getFolderByServerRelativeUrl('" + path + "')" +
						"/Files/add(url='" + SPListREST.escapeApostrophe(parameters.fileName) +
						"',overwrite=" + (parameters.willOverwrite == true ? "true" : "false") +
						")?$expand=ListItemAllFields",
				method: "POST",
				body: parameters.body,
				successCallback: (data: TSPResponseData) => {
					resolve(data.status as HttpStatusCode);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	createFolder(parameters: {folderName: string}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/folders",
				method: "POST",
				body: JSON.stringify({
					"__metadata": {
						"type": "SP.Folder"
					},
					"ServerRelativeUrl": this.serverRelativeUrl + "/" + parameters.folderName
				}),
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
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
		RESTrequest({
			setDigest: true,
			url: this.forApiPrefix + "/_api/web/getFolderByServerRelativeUrl('" +
			url: this.forApiPrefix + "/_api/web/getFolderByServerRelativeUrl('" +
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
	}): Promise<number> {
		// parameters.tryCheckOutResolve == true will
		//   perform a user request of resolving without checking
		//   if user is identical with checked out user
		return new Promise<number>((resolve, reject) => {
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
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/getFileByServerRelativeUrl('" + path + "')/CheckOut()",
				method: "POST",
				body: "",
				successCallback: (data, httpInfo) => {
					resolve(httpInfo.status);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
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
		checkInType?: TSPDocLibCheckInType,
		checkInComment?: string
	}): Promise<number> {
		// checkintype = 0: minor version, = 1: major version, = 2: overwrite
		const requests: IBatchHTTPRequestForm[] = [];
		let checkinType: number;
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
		for (const itemName of parameters.itemNames)
			requests.push({
				url: this.forApiPrefix + "/_api/web/GetFileByServerRelativeUrl('" + itemName +
						"')/CheckIn(comment='" + parameters.checkInComment + "',checkintype=" + checkinType + ")",
				contextinfo: this.forApiPrefix,
				method: "POST"
			});
		return new Promise<number>((resolve, reject) => {
			batchRequestingQueue(requests);
			resolve(1);
		});
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
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			let path: string;

			path = this.baseUrl;
			if (parameters.folderPath.charAt(0) == "/")
				path += parameters.folderPath;
			else
				path += "/" + parameters.folderPath;
			path += "/" + parameters.file;
			RESTrequest({
				// library function
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/GetFileByServerRelativeUrl('" + path + "')" +
						"/UndoCheckout()",
				method: "POST",
				body: "", // TODO a post method with no body?
				successCallback: (data, httpInfo) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	renameItem(parameters: any) {
		return this.renameFile(parameters);
	}

	renameFile(parameters: {
		itemId: number;
		itemName: string;
		newName: string;
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				// SPListItemRaw
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid +
						"')/items(" + parameters.itemId + ")",
				method: "POST",
				headers: {
					"X-HTTP-METHOD": "MERGE",
					"IF-MATCH": "*" // can also use etag
				},
				body: JSON.stringify({
					FileLeafRef: parameters.itemName ? parameters.itemName : parameters.newName
				}) as TXmlHttpRequestData,
				successCallback: (data, httpInfo) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}


	// 3 arguments are necessary:
	// .itemId: ID of file item
	// .newFileName or .newName: name to be used in rename
	renameItemWithCheckout(parameters: {
		itemId: number;
		newName: string;
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')" +
						"/items(" + parameters.itemId + ")",
				method: "POST",
				headers: {
					"X-HTTP-METHOD": "MERGE",
					"IF-MATCH": "*" // can also use etag
				},
				body: JSON.stringify({
					"FileLeadRef": parameters.newName
				}),
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
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
		checkInType?: TSPDocLibCheckInType;
	}) {
		return new Promise((resolve, reject) => {
			if ((!parameters.itemId || isNaN(parameters.itemId) == true) &&
				!parameters.fileName)
				throw "argument must contain {itemId:<value>} or {fileName:<name>}";
			if (!parameters.body)
				throw "parameters.body not found/defined: must be formatted JSON string";
			if (!parameters.fileName)
				this.getListItemData({
					itemId: parameters.itemId,
					expand: "File"
				}).then((response: TSpListItem[] | TSpListItem | null) => {
					if (response && Array.isArray(response) == false)
						parameters.fileName = (response as TSpListItem).File.Name;
					this.continueupdateLibItemWithCheckout(parameters).then((response: unknown) => {
						resolve(response);
					}).catch((err: unknown) => {
						reject(err);
					});
				}).catch((err: unknown) => {
					reject(err);
				});
			else
				this.continueupdateLibItemWithCheckout(parameters).then((response: unknown) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
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
		checkInType?: TSPDocLibCheckInType;
		body: string;
	}) {
		return new Promise((resolve, reject) => {
			this.checkOutDocLibItem({
				itemName: parameters.fileName,
			}).then(() => {
				RESTrequest({
					setDigest: true,
					url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid +
							"')/items(" + parameters.itemId + ")",
					method: "POST",
					body: "{ '__metadata': { 'type': '" +
						this.listItemEntityTypeFullName + "' }, " +
						parameters.body + " }",
					headers: {
						"X-HTTP-METHOD": "MERGE",
						"IF-MATCH": "*" // can also use etag
					},
					successCallback: (data) => {
						this.checkInDocLibItem({
							itemNames: [ parameters.fileName ],
							checkInType: parameters.checkInType
						}).then((response: unknown) => {
							resolve(response);
						}).catch((err: unknown) => {
							reject(err);
						});
					},
					errorCallback: (err) => {}
				});
			}).catch((err: unknown) => {
				reject(err);
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
		body: THttpRequestBody
	}): Promise<boolean> {
	//	const thisInstance = this;
		return new Promise<boolean>((resolve, reject) => {
			let itemName: string,
				majorResponse: number;
			this.getDocLibItemFileAndMetaData({
				itemId: parameters.itemId
			}).then((response: SPFileRaw) => {
				this.checkOutDocLibItem({
					itemName: itemName = response.Name
				}).then((response: number) => {
					this.updateListItem({
						itemId: parameters.itemId,
						body: parameters.body as TSpListItem
					}).then((response: number) => {
						majorResponse = response;
						this.checkInDocLibItem({
							itemNames: [ itemName ]
						}).then(() => {
							if (majorResponse < 400)
								resolve(true);
						}).catch((err: unknown) => { // check in failure
							reject(err);
						});
					}).catch((err: unknown) => { // update failure
						reject(err);
					});
				}).catch((err: unknown) => { // check out failure
					reject(err);
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
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			let uniqueId;
			if (!parameters.sourceFileName)
				throw "parameters.sourceFileName not found/defined";
			if (!parameters.destinationFileName)
				throw "parameters.destinationFileName not found/defined";
			if (typeof parameters.willOverwrite == "undefined")
				parameters.willOverwrite = "false";
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/GetFileByServerRelativeUrl('" +
						parameters.sourceFileName + "')?$expand=ListItemAllFields",
				successCallback: (data, httpInfo) => {
					uniqueId = data.d?.UniqueId;
					RESTrequest({
						url: this.forApiPrefix + "/_api/web/GetFileById(guid'" + uniqueId +
							"')/CopyTo(strNewUrl='" + parameters.destinationFileName +
							"',bOverWrite=" + parameters.willOverwrite + ")",
						method: "POST",
						successCallback: (data, httpInfo) => {
							resolve(data);
						},
						errorCallback: (err) => {
							reject(err);
						}
					});
				}
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
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			if (parameters.folderPath && parameters.fileName)
				parameters.path = parameters.folderPath + "/" + parameters.fileName;
			if (typeof parameters.includeBaseUrl == "undefined")
				parameters.includeBaseUrl = true;
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/GetFileByServerRelativeUrl('" +
					(parameters.includeBaseUrl == false ? "" : this.baseUrl) + parameters.path + "')" +
					((parameters.recycle && parameters.recycle == false as boolean) ? "" : "/recycle()"),
				method: "POST",
				body: "", // TODO  this needs a body?
				successCallback: (data, httpInfo) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
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
	} | undefined): Promise<SPFieldRaw[]> {
		return new Promise<SPFieldRaw[]>((resolve, reject) => {
			const query = "";
			let filter: string;

			if (parameters && parameters.fields != null && parameters.fields.length > 0) {
				if (parameters.filter)
					filter = parameters.filter;
				else
					filter = "";
				for (const fld of parameters.fields) {
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
			RESTrequest({
				// SPField[]
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/fields" + query,
				successCallback: (data, httpInfo) => {
					resolve(data as SPFieldRaw[]);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	/**
	 * @method getFields -- will return information on all list fields
	 * @returns Promise with data in then() response
	 */
	getFields(parameters?: {filter: string;}): Promise<unknown> {
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
	}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			const body: THttpRequestBody | string = fieldProperties;

			if (checkEntityTypeProperty(body, "field") == false)
				body.__SetType__ = "SP.Field";
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/fields",
				method: "POST",
				body: body as TXmlHttpRequestData,
				successCallback: (data, httpInfo) => {

				},
				errorCallback: (errInfo: unknown, fetchInfo) => {

				}
			});
		});
	}

	/**
	 *
	 * @param {arrayOfFieldObject} fields -- this should be an array of the
	 * 			       objects used to define fields for SPListREST.createField()
	 * @returns
	 */
	createFields(fields: TSpField[]): Promise<unknown> {
		// field creation can not occur co-synchronously. One must complete before the
		// next field creation can begin.
		return new Promise((resolve, reject) => {
			const requests: IBatchHTTPRequestForm[] = [];
			let body: THttpRequestBody;

			for (const fld of fields) {
				body = fld;
				if (checkEntityTypeProperty(body, "field") == false)
					body.__SetType__ = "SP.Field";
				requests.push({
					url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/fields",
					contextinfo: this.forApiPrefix,
					url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/fields",
					contextinfo: this.forApiPrefix,
					method: "POST",
					body: body
				});
			}
			batchRequestingQueue(requests).then((response: unknown) => {
				resolve(response);
			}).catch((response: unknown) => {
				reject(response);
			});
			/*
			serialSPProcessing(this.createField, fields).then((response: unknown) => {
				resolve(response);
			}).catch((response: unknown) => {
				reject(response);
			});  */
		});
	}

	renameField(parameters: {
		oldName?: string;
		currentName?: string;
		newName: string | undefined;
	}) {
		if (!parameters.oldName && !parameters.currentName)
			throw "A parameter for 'oldName' or 'currentName' was not found";
		if (parameters.oldName)
			parameters.currentName = parameters.oldName;
		if (!parameters.newName)
			throw "A parameters for 'newName' was not found";
		const body = {
			"Title": parameters.newName
		};
		if (checkEntityTypeProperty(body, "field") == false)
			(body as any)["__SetType__"] = "SP.Field";
		return new Promise((resolve, reject) => {
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid +
						"')/fields/GetByTitle('" + parameters.currentName + "')",
				method: "POST",
				headers: {
					"X-HTTP-METHOD": "MERGE",
					"IF-MATCH": "*"
				},
				body: JSON.stringify(body),
				successCallback: (data, httpInfo) => {
					resolve(data);
				},
				errorCallback: (errInfo: unknown, httpInfo: HttpInfo) => {
					reject(errInfo);
				}
			});
		});
	}

	/** @method .updateChoiceTypeField
	 *       @param {Object} parameters
	 *       @param {string} parameters.id - the field guid/id to be updated
	 *       @param {(array|arrayAsString)} parameters.choices - the elements that will
	 *     form the choices for the field, as an array or array written as string
	 */
	updateChoiceTypeField (
		parameters: {
			id: string;
			choices?: string[];
			choice?: string
		}
	) {
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
		RESTrequest({
			setDigest: true,
			url: this.forApiPrefix + "/_api/web/lists(guid'" +
			url: this.forApiPrefix + "/_api/web/lists(guid'" +
						this.listGuid + "')/fields(guid'" + parameters.id + "')",
			method: "POST",
			headers: {
				"IF-MATCH": "*",
				"X-HTTP-METHOD": "MERGE"
			},
			body: "{ '__metadata': { 'type': 'SP.FieldChoice' }, " +
				"'Choices': { '__metadata': { 'type': 'Collection(Edm.String)' }, " +
					"'results': " + choices + "} }",
			successCallback: (data) => {},
			errorCallback: (err) => {}
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
	} | undefined): Promise<SPContentTypeRaw> {
		return new Promise<SPContentTypeRaw>((resolve, reject) => {
			let query = "",
				filter: string;

			if (parameters && parameters.contentTypes != null && parameters.contentTypes.length > 0) {
				if (parameters.query && parameters.query.filter)
					filter = parameters.query.filter;
				else
					filter = "";
				for (const ct of parameters.contentTypes) {
					if (filter.length > 0)
						filter += " or ";
					if (ct.name)
						filter += "Title eq '" + ct.name + "'";
					else
						filter = "Id eq '" + ct.guid + "'";
				}
			}
			if (parameters && parameters.query)
				query = constructQueryParameters(parameters.query);
			RESTrequest({
				// SPContentType[]
				url: this.forApiPrefix + "/_api/web/lists(guid'" +
						this.listGuid + "')/contentTypes" + query,
				successCallback: (data) => {
					resolve(data as SPContentTypeRaw);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
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
			serialSPProcessing(this.createContentType, contentTypes).then((response: unknown) => {
				resolve(response);
			}).catch((response: unknown) => {
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
	}}): Promise<unknown> {
		return new Promise((resolve, reject) => {
			let query:string = "",
				filter: string;

			if (parameters && parameters.views != null && parameters.views.length > 0) {
				if (parameters.query!.filter)
					filter = parameters.query!.filter;
				else
					filter = "";
				for (const view of parameters.views) {
					if (filter.length > 0)
						filter += " or ";
					if (view.name)
						filter += "Title eq '" + view.name + "'";
					else
						filter = "Id eq '" + view.guid + "'";
				}
			}
			if (parameters && parameters.query)
				query = constructQueryParameters(parameters.query);
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/lists(guid'" +
						this.listGuid + "')/views" + query,
				successCallback: (data, httpInfo) => {
					resolve(data);
				},
				errorCallback: (err, fetchInfo) => {
					reject(err);
				}
			});
		});
	}

	/**
	 *
	 * @param {JSONofParms} viewProperties }
	 * @returns REST API response
	 */
	createView(viewsProperties: THttpRequestBody): Promise<unknown> {
		return new Promise((resolve, reject) => {
			let body: THttpRequestBody | string = viewsProperties;

			if (checkEntityTypeProperty(body, "view") == false)
				body.__SetType__ = "SP.View";
			body = JSON.stringify(body);
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + this.listGuid + "')/views",
				method: "POST",
				body: body,
				successCallback: (response, httpInfo) => {
					resolve(response);
				},
				errorCallback: (errInfo) => {
					reject(errInfo);
				}
			});
		});
	}

	setView(viewName: string, fieldNames: string[]) {

		let viewFields = "";

		for (const field of fieldNames)
			viewFields += "<FieldRef Name=\"" + field + "\"/>";

		RESTrequest({
			setDigest: true,
			url: this.forApiPrefix + "/_api/web/GetList(@a1)/Views(@a2)/SetViewXml()?@a1='" +
			url: this.forApiPrefix + "/_api/web/GetList(@a1)/Views(@a2)/SetViewXml()?@a1='" +
					this.serverRelativeUrl + "'&@a2='" + this.listGuid + "'",
			method: "POST",
			body: JSON.stringify({"ViewXml": "<View Name=\"{" + createGuid() + "}\" DefaultView=\"FALSE\" MobileView=\"TRUE\" " +
					"MobileDefaultView=\"TRUE\" Type=\"HTML\" DisplayName=\"All Items\" " +
					"Url=\"" + this.serverRelativeUrl + "/" + viewName + ".aspx\" Level=\"1\" " +
					"BaseViewID=\"1\" ContentTypeID=\"0x\" ImageUrl=\"/_layouts/15/images/generic.png?rev=47\">" +
					"<Query><OrderBy><FieldRef Name=\"ID\" Ascending=\"FALSE\"/></OrderBy></Query>" +
					"<ViewFields><FieldRef Name=\"Title\"/><FieldRef Name=\"Created\"/><FieldRef Name=\"Modified\"/>" +
					viewFields + "</ViewFields>" +
					"<RowLimit Paged=\"TRUE\">30</RowLimit><JSLink>clienttemplates.js</JSLink><XslLink Default=\"TRUE\">main.xsl</XslLink>" +
					"<Toolbar Type=\"Standard\"/><Aggregations Value=\"Off\"/></View>"}),
			successCallback: (data, httpInfo) => {

			}
		});
	}
	/**
	 *
	 * @param {arrayOfFieldObject} views -- this should be an array of the
	 * 			       objects used to define views for SPListREST.createViews()
	 * @returns
	 */
	createViews (views: unknown[]): Promise<unknown> {
		// field creation can not occur co-synchronously. One must complete before the
		// next field creation can begin.
		return new Promise((resolve, reject) => {
			serialSPProcessing(
				this.createView as (arg1: unknown) => Promise<unknown>,
				views
			)
				.then((response: unknown) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
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
				}).catch((response: unknown) => {
					reject(response);
				});
			}
			recurse(0);
		});	*/
	}
}


// if parameters.success not found the request will be SYNCHRONOUS and not ASYNC
/**
	 * @method httpRequest -- this is a static method
	 * @param {ThttpRequestParams} elements
	 * 	this object contains several properties crucial to setting up a JQuery ajax() call:
	 * 		url, method, success, error, headers, data
	 *    it also contains a property 'progressReport' which is object with properties 'interval' (type number)
	 * 		and 'callback' (type (number) => void). This is intended for network request that might take
	 * 		a long time with several results cycles. interval is the count of results collected, and
	 * 		the callback is to a function so that an update in the callback can be done.
	 * @returns {JQueryXHR}  standard JQuery ajax() return type
	 */
/*
	static httpRequest(elements: THttpRequestParams): JQueryXHR {
		let intervalControl: TIntervalControl,
			aggregateResults: unknown[] = [];

		if (!elements.successCallback)
			throw "HTTP Request for SPListREST requires defining 'successCallback' in parameter 'elements'";
		if (!elements.errorCallback)
			throw "HTTP Request for SPListREST requires defining 'errorCallback' in parameter 'elements'";
		if (typeof elements.headers == "undefined")
			elements.headers = SPListREST.stdHeaders;
		else { // default headers
			if (!elements.headers["Accept"])
				elements.headers["Accept"] = "application/json;odata=verbose";
			if (!elements.headers["Content-Type"])
				elements.headers["Content-Type"] = "application/json;odata=verbose";
		}
		if (elements.progressReport) // count of results when multiple network calls needed
			intervalControl = {
				currentCount: 0,
				nextCount: elements.progressReport.interval,
				interval: elements.progressReport.interval,
				callback: elements.progressReport.callback
			};
		else
			elements.progressReport = null;

		// for any method other thatn "GET"
		if (elements.setDigest && elements.setDigest == true) {
			const match = elements.url.match(/(.*\/_api)/);

			if (match == null)
				throw "Problem parsing '/_api/' of URL to get request digest value";

			RESTrequest({
				url: match[1] + "/contextinfo",
				method: "POST",
				headers: SPListREST.stdHeaders,
				successCallback: (data/*, text, reqObj ) => {
					elements.headers!["X-RequestDigest"] = data.FormDigestValue ? data.FormDigestValue :
					data.d!.GetContextWebInformation!.FormDigestValue;
					RESTrequest({
						url: elements.url,
						method: elements.method,
						headers: elements.headers,
						data: elements.body as unknown ?? elements.data as unknown,
						successCallback: (data/*, text, reqObj ) => {

						},
						errorCallback: (reqObj/*, status, errThrown ) => {
							if (elements.ignore)
								for (let i = 0; i < elements.ignore.length; i++)
									if (elements.ignore[i] == requestObj.status)
										elements.successCallback(data, status, requestObj);
							elements.errorCallback!(requestObj, status, thrownErr);
						}
					});
				},
				errorCallback: (reqObj/*, status, errThrown ) => {

				}
			});

		} else {  // GET methods
			if (!elements.method)
			 	elements.method = "GET";
			return $.ajax({
				url: elements.url,
				method: elements.method,
				headers: elements.headers,
//				data: elements.body as unknown ?? elements.data as unknown,
				success: (data: TSPResponseData, status: string, requestObj: JQueryXHR) => {
					aggregateResults = aggregateResults.concat(data.d?.results);
					if (data.d && data.d.__next) {  // result data too numerous
						this.RequestAgain(
							data.d.__next,
							data.d.results as unknown[], // aggregate data
							intervalControl
						).then((response: unknown) => {
							elements.successCallback!(response);
						}).catch((response: unknown) => {
							elements.errorCallback!(response);
						});
					} else
						elements.successCallback!(data, status, requestObj);
				},
				error: (requestObj: JQueryXHR, status: string, thrownErr: string) => {
					let ignored = false;

					// ignore is an array of HTTP Response codes to ignore
					// if the returned
					if (elements.ignore)
						for (let i = 0; i < elements.ignore.length; i++)
							if (elements.ignore[i] == requestObj.status) {
								ignored = true;
								elements.successCallback(aggregateResults as unknown as TSPResponseData, status, requestObj);
							}
					if (ignored == false)
						elements.errorCallback!(requestObj, status, thrownErr);
				}
			});
		}
	}

	static RequestAgain(
		nextUrl: string,
		aggregateData: unknown[],
		intervalControl: TIntervalControl
	): Promise<unknown> {
		return new Promise((resolve, reject) => {
			$.ajax({
				url: nextUrl,
				method: "GET",
				headers: SPListREST.stdHeaders,
				success: (data) => {
					if (data.d.__next) {
						if (intervalControl && intervalControl.interval > 0) {
							intervalControl.currentCount += data.d.results ? data.d.results!.length : 0;
							if (intervalControl.currentCount! >= intervalControl.nextCount!) {
								intervalControl.nextCount! += intervalControl.interval;
								intervalControl.callback(intervalControl.currentCount!);
							}
						}
						this.RequestAgain(
							data.d.__next,
							data.d.results,
							intervalControl
						).then((response: unknown) => {
							resolve(aggregateData.concat(response));
						}).catch((err: unknown) => {
							reject(err);
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

	static httpRequestPromise(parameters: THttpRequestParamsWithPromise): Promise<unknown> {
		return new Promise((resolve, reject) => {
			SPListREST.httpRequest({
				setDigest: parameters.setDigest,
				url: parameters.url,
				method: parameters.method,
				headers: parameters.headers,
				data: parameters.data ?? parameters.body,
				ignore: parameters.ignore,
				progressReport: parameters.progressReport,
				successCallback: (data: TSPResponseData, status: string | undefined,
							reqObj: JQueryXHR | undefined) => {
					resolve({data: data, message: status, reqObj: reqObj});
				},
				errorCallback: (reqObj: JQueryXHR, text: string | undefined, errThrown: string | undefined) => {
					reject({reqObj: reqObj, text: text, errThrown: errThrown});
				}
			});
		});
	}
*/
