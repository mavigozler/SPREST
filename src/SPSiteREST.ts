"use strict";

/* jshint -W119, -W069 */

/**
 * @class SPSiteREST
 */
class SPSiteREST {
	server: string;
	sitePath: string;
	apiPrefix: string;
	id: string = "";
	serverRelativeUrl: string = "";
	isHubSite: boolean = Boolean();
	webId: string = "";
	webGuid: string = "";
	creationDate: Date = new Date(0);
	siteName: string = "";
	homePage: string = "";
	siteLogoUrl: string = "";
	template: string = "";
	static stdHeaders: THttpRequestHeaders = {
		"Accept": "application/json;odata=verbose",
		"Content-Type": "application/json;odata=verbose"
	};
	arrayPedigree: TSiteInfo[] = [];

	/**
	 * @constructor SPSiteREST -- sets up properties and methods instance to describe a SharePoint site
	 * @param {object} setup -- representing parameters to establish interface with a site
	 *             {server: name of location.host, site: path to site}
	 * @private
	 */
	constructor (setup: {
		server: string;
		site: string;
	}) {
		if (this instanceof SPSiteREST == false)
			throw "Was 'new' operator used? Use of this function object should be with constructor";

		/*
		this.server;
		this.site;
	*/
		if (!setup || typeof setup != "object" || !(setup.server || setup.site))
			throw "Use of SPSiteREST() constructor must include server and site";
		if (!setup.server)
			throw "Constructor requires defining server as arg { server:<server> }";
		this.server = setup.server;
		if (setup.server.search(/^http/) < 0)
			this.server = "https://" + setup.server;
		else
			this.server = setup.server;
		this.sitePath = setup.site;
		if (setup.site.charAt(0) != "/")
			this.sitePath = "/" + this.sitePath;
		this.apiPrefix = this.server + this.sitePath + "/_api";
	}

	/**
	 * @method create -- sets many list characteristics whose data is required
	 * 	by asynchronous request
	 * @returns Promise
	 */
	init() {
		return new Promise((resolve, reject) => {
			$.ajax({
				url:  this.apiPrefix + "/site",
				method: "GET",
				headers: SPSiteREST.stdHeaders,
				success: (data) => {
					data = data.d;
					this.id = data.Id;
					this.serverRelativeUrl = data.ServerRelativeUrl;
					this.isHubSite = data.IsHubSite;
					this.webId = this.webGuid = data.Id;
					$.ajax({
						url: this.apiPrefix + "/web",
						method: "GET",
						headers: SPSiteREST.stdHeaders,
						success: (data) => {
							data = data.d;
							this.creationDate = data.Created;
							this.siteName = data.Title;
							this.homePage = data.WelcomePage;
							this.siteLogoUrl = data.SiteLogoUrl;
							this.template = data.WebTemplate;
							resolve(true);
						},
						error: (reqObj) => {
							reject(reqObj);
						}
					});
				},
				error: (reqObj) => {
					reject(reqObj);
				}
			});
		});
	} // _init()

	// if parameters.success not found the request will be SYNCHRONOUS and not ASYNC
	static httpRequest(elements: THttpRequestParams) {
		if (elements.setDigest && elements.setDigest == true) {
			let match = elements.url.match(/(.*\/_api)/);

			if (match == null)
				throw "Could not extra '/_api/' part of URL to get request digest";
			$.ajax({  // get digest token
				url: match[1] + "/contextinfo",
				method: "POST",
				headers: {...SPSiteREST.stdHeaders},
				success: function (data) {
						 let headers: THttpRequestHeaders | undefined = elements.headers;

						 if (headers) {
							  headers["Content-Type"] = "application/json;odata=verbose";
							  headers["Accept"] = "application/json;odata=verbose";
						 } else
							  headers = {...SPSiteREST.stdHeaders};
					headers["X-RequestDigest"] = data.d.FormDigestValue ? data.d.FormDigestValue :
						data.d.GetContextWebInformation.FormDigestValue;
					$.ajax({
						url: elements.url,
						method: elements.method,
						headers: headers,
						data: elements.body ?? elements.data as string,
						success: function (data: TSPResponseData, status: string, requestObj: JQueryXHR) {
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
				 elements.headers = {...SPSiteREST.stdHeaders};
			 else {
				 elements.headers["Content-Type"] = "application/json;odata=verbose";
				 elements.headers["Accept"] = "application/json;odata=verbose";
			 }
			$.ajax({
				url: elements.url,
				method: elements.method ? elements.method : "GET",
				headers: elements.headers,
				data: elements.body ?? elements.data as string,
				success: function (data: TSPResponseData, status: string, requestObj: JQueryXHR) {
					if (data.d && data.__next)
						RequestAgain(
							elements,
							data.__next,
							data.d.results!
						).then((response: any) => {
							elements.successCallback!(response);
						}).catch((response: any) => {
							elements.errorCallback!(response);
						});
					else
						elements.successCallback!(data, status, requestObj);
				},
				error: function (requestObj: JQueryXHR, status: string, thrownErr: string) {
					elements.errorCallback!(requestObj, status, thrownErr);
				}
			});
		}
	}

	static  httpRequestPromise(parameters: THttpRequestParamsWithPromise) {
		return new Promise((resolve, reject) => {
			SPSiteREST.httpRequest({
				setDigest: parameters.setDigest,
				url: parameters.url,
				method: parameters.method ? parameters.method : "GET",
				headers: parameters.headers,
				data: parameters.data ?? parameters.body as TXmlHttpRequestData,
				successCallback: (data: TSPResponseData, message: string | undefined, reqObj: JQueryXHR | undefined) => {
					resolve({data: data, message: message, reqObj: reqObj});
				},
				errorCallback: (reqObj: JQueryXHR, status: string | undefined, err: string | undefined) => {
					reject({reqObj: reqObj, text: status, errThrown: err});
				}
			});
		});
	}

	/**
	 *
	 * @param {*} parameters -- object which may contain select, filter, expand properties
	 */
	getProperties(parameters?: any) {
		return SPSiteREST.httpRequestPromise({
			url: this.apiPrefix + "/site" + constructQueryParameters(parameters)
		});
	}

	getSiteProperties(parameters?: any) {
		return this.getProperties(parameters);
	}

	getWebProperties(parameters?: any) {
		return SPSiteREST.httpRequestPromise({
			url: this.apiPrefix + "/web" + constructQueryParameters(parameters)
		});
	}

	getEndpoint(endpoint: string) {
		return SPSiteREST.httpRequestPromise({
			url: this.apiPrefix + endpoint
		});
	}

	getSubsites() {
		return new Promise((resolve, reject) => {
			SPSiteREST.httpRequest({
				url: this.apiPrefix + "/web/webinfos",
				method: "GET",
				successCallback: (data: TSPResponseData) => {
					let sites: TSiteInfo[] = [ ],
						results = data.d!.results;

					for (let result of results!)
						sites.push({
							Name: result.Title,
							ServerRelativeUrl: result.ServerRelativeUrl,
							Id: result.Id as string,
							Template: result.WebTemplate,
							Title: result.Title,
							Url: "",
							Description: "",
							forApiPrefix: "",
							tsType: "TSiteInfo",
							siteParent: undefined
						});
					resolve(sites);
				},
				errorCallback: (reqObj: JQueryXHR, status: string | undefined, text: string | undefined) => {
					reject({reqObj: reqObj, status: status, text: text});
				}
			});
		});
	}

	getParentWeb(siteUrl: string): Promise<any> {
		return new Promise((resolve, reject) => {
			SPSiteREST.httpRequestPromise({
				url: siteUrl + "/_api/web?$expand=ParentWeb"
			}).then((response: any) => {
				let data = response.data.d;
				if (data.ParentWeb.__metadata)
					resolve(data.ParentWeb);
				else
					resolve(null);
			}).catch((response) => {
				reject(response);
			});
		});
	}

	getSitePedigree(siteUrl: string | null): Promise<{pedigree: TSiteInfo; arrayPedigree: TSiteInfo[]}> {
		return new Promise((resolve, reject) => {
			let pedigree: TSiteInfo = {} as TSiteInfo,

					siteInfo: TSiteInfo = {} as TSiteInfo,
					repeatGetParent = (siteInfo: TSiteInfo) => {
						this.getParentWeb(siteInfo.server! + siteInfo.ServerRelativeUrl).then((ParentWeb) => {
							if (ParentWeb) {  // != null
								let siteParent: TSiteInfo = {} as TSiteInfo;
								ParentWeb = ParentWeb.data.d;
								siteParent.Name = ParentWeb.Title;
								siteParent.server = siteInfo.server;
								siteParent.ServerRelativeUrl = ParentWeb.ServerRelativeUrl;
								siteParent.Id = ParentWeb.Id;
								siteParent.Template = ParentWeb.WebTemplate;
								siteInfo.parent = siteParent;
								siteParent.subsites = [];
								siteParent.subsites.push(siteInfo as any);
								repeatGetParent(siteParent);
							} else {
								this.fillOutPedigree(siteInfo).then((response) => {
									resolve({pedigree: pedigree, arrayPedigree: this.arrayPedigree});
								}).catch((response) => {
									reject(response);
								});;  // use the last site
							}
						}).catch((response) => {
							reject(response);
						});
					};

			if (!siteUrl) {
				siteInfo.server = this.server;
				siteInfo.ServerRelativeUrl = this.serverRelativeUrl;
			} else {
				let parsedUrl: TParsedURL | null = ParseSPUrl(siteUrl);

				if (parsedUrl == null)
					throw "Parameter 'siteUrl' was not a parseable SharePoint URL";
				siteInfo.server = parsedUrl.server;
				siteInfo.ServerRelativeUrl = parsedUrl.siteFullPath;
			}
			pedigree.referenceSite = {
			 ///pedigree.referenceSite = {
				Name: undefined,
				server: siteInfo.server,
				ServerRelativeUrl: this.serverRelativeUrl,
				Id: this.id,
				Template: this.template,
				parent: null,
				subsites: [],
				Title: siteInfo.Title,
				Description: siteInfo.Description,
				Url: siteInfo.Url,
				forApiPrefix: "",
				tsType: "TSiteInfo",
				siteParent: undefined
			};
			this.arrayPedigree = [];
			this.arrayPedigree.push(pedigree.referenceSite as TSiteInfo);
			//repeatGetParent(pedigree.referenceSite!);
			repeatGetParent(pedigree.referenceSite!);
		});
	}

	fillOutPedigree (parentWeb: TSiteInfo): Promise<any> {
		return new Promise((resolve, reject) => {
			let subSite: SPSiteREST;

			if (!parentWeb)
				throw "fillOutPedigree():  parameter 'parentWeb' is not defined";
			subSite = new SPSiteREST({
					server: this.server as string,
					site: parentWeb.ServerRelativeUrl as string
				});
			subSite.init().then(() => {
				subSite.getSubsites().then((webInfos: any) => {
					if (typeof webInfos == "undefined")
						reject("Promise.success() returned nothing");
					else {
						if (typeof parentWeb.subsites == "undefined")
							parentWeb.subsites = [];
						if (webInfos.length > 0) {
							let idx: number,
									webInfoRequests: Promise<any>[] = [];

							for (let webinfo of webInfos) {
								idx = parentWeb.subsites.push({
									Name: webinfo.name,
									server: parentWeb.server,
									ServerRelativeUrl: webinfo.serverRelativeUrl,
									Id: webinfo.id,
									Url: webinfo.url,
									Template: webinfo.template,
									parent: parentWeb,
									subsites: [],
									tsType: "TSiteInfo",
									Title: "",
									Description: "",
									forApiPrefix: "",
									siteParent: undefined
								}) - 1;
								this.arrayPedigree.push((parentWeb.subsites as TSiteInfo[])[idx]);
								webInfoRequests.push(this.fillOutPedigree((parentWeb.subsites as TSiteInfo[])[idx]));
							}

							Promise.all(webInfoRequests).then((response) => {
								resolve(true);
							}).catch((response) => {
								reject(response);
							});
						} else
							resolve(true);
					}
				}).catch((response) => {
					reject(response);
				});
			});
		});
	}

	getSiteColumns(parameters: any): Promise<any> {
		return SPSiteREST.httpRequestPromise({
			url: this.apiPrefix + "/web/fields" + constructQueryParameters(parameters)
		});
	}

	getSiteContentTypes(parameters: any): Promise<any> {
		return SPSiteREST.httpRequestPromise({
			url: this.apiPrefix + "/web/ContentTypes" + constructQueryParameters(parameters)
		});
	}

	getLists(parameters?: any):Promise<any> {
		return SPSiteREST.httpRequestPromise({
			url: this.apiPrefix + "/web/lists" + constructQueryParameters(parameters)
		});
	}

	/**
	 *
	 * @param {*} parameters -- need to control these!
	 * @returns
	 */
	createList(parameters: {body: THttpRequestBody}) {
		let body: THttpRequestBody | string = parameters.body;

		if (checkEntityTypeProperty(body, "item") == false)
			body["__SetType__"] = "SP.List";
		body = formatRESTBody(body);
		return SPSiteREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists",
			method: "POST",
			body: body
		});
	}

	/**
	 * @method updateListByMerge
	 * @param {object} parameters -- need a 'Title' parameter
	 * @returns
	 */
	updateListByMerge(parameters: {body: THttpRequestBody, listGuid: string}) {
		let body: THttpRequestBody | string = parameters.body;

		if (checkEntityTypeProperty(body, "item") == false)
			body["__SetType__"] = "SP.List";
		body = formatRESTBody(body);
		return SPSiteREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists" +
					"(guid'" + parameters.listGuid + "')", /// this cannot be correct
			method: "POST",
			headers: {
				"IF-MATCH": "*",
				"X-HTTP-METHOD": "MERGE"
			}
		});
	}

	deleteList(parameters: {id?: string, guid?: string}) {
		let id = parameters.id ?? parameters.guid;

		if (!id)
			throw "List deletion requires an 'id' or 'guid' parameter for the list to be deleted.";
		return SPSiteREST.httpRequestPromise({
			setDigest: true,
			url: this.apiPrefix + "/web/lists" +
					"(guid'" + id + "')",
			method: "POST",
			headers: {
				"X-HTTP-METHOD": "DELETE",
				"IF-MATCH": "*" // can also use etag
			}
		});
	}

	makeLibCopyWithItems(
		sourceLibName: string,
		destLibSitePath: string,
		destLibName: string,
		itemCopyCount?: number
	): Promise<any> {
		// steps: (1) get source lib info on fields,
		//  (2) create dest list/lib, then its fields, then views, try content types
		//  (3) copy the items
		return new Promise((resolve, reject) => {
			let sourceListREST = new SPListREST({
				server: this.server,
				site: this.sitePath,
				listName: sourceLibName
			});
			sourceListREST.init().then(() => {
				this.createList({
					body: {
						"AllowContentTypes": sourceListREST.allowContentTypes,
						"BaseTemplate": sourceListREST.baseTemplate,
						"Title": destLibName
					}
				}).then((response: any) => {
					let iDestListREST = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: response.data.d.Title
					});
					iDestListREST.init().then((response: any) => {
						this.copyFields(sourceListREST, iDestListREST).then(() => {
							this.copyFolders(sourceListREST, iDestListREST).then(() => {
								this.copyFiles(sourceListREST, iDestListREST).then((response) => {
									resolve(response);
								}).catch((response) => {
									reject(response);
								});
							}).catch((response) => {
								reject(response);
							});
						}).catch((response) => {
							reject(response);
						});
					}).catch((response: any) => {
						reject(response);
					});
				}).catch((response) => {
					if (response.reqObj.status == 500 &&
							response.reqObj.responseJSON.error.message.value.search(
								/list.*already exists.*Please choose another title\./
								) >= 0)
						sourceListREST.getListItems({
								top: itemCopyCount
							}).then((response: any) => {
							/*
								iDestListREST.createListItems(response.data.d.ressults).then((response) => {
									resolve(response);
								}).catch((response) => {
									reject(response);
								}); */
							}).catch((response: any) => {
								reject(response);
							});
						else {
							reject(response);
						}
				});
			}).catch((response: any) => {
				reject (response);
			});
		});
	}

	copyFields(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST
	): Promise<any> {
		return new Promise((resolve, reject) => {
			let fieldDefs: {[key:string]: any}[] = [ ],
					fieldExclusions: {[key:string]: boolean | string[] | number[] } = {
						CanBeDeleted: false,
						ReadOnlyField: true,
						FieldTypeKind: [0, 7, 11, 12],
						Title: ["^Shortcut", "^Icon", "^URL"]
					};

			if (typeof sourceLib == "string")
				sourceLib = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: sourceLib
					});
			if (sourceLib == null)
				reject("Source lib not found in site");
			sourceLib.getFields().then((response: any) => {
				let sourceFields = response.data.d.results;

				for (let field of sourceFields) {
					for (let property in fieldExclusions) {
						if (Array.isArray(fieldExclusions[property]) == true)
							for (let elem of fieldExclusions[property] as (string[] | number[])) {
								if (field[property] == elem || typeof elem == "string" &&
											elem.charAt(0) == "^" &&
											new RegExp(elem.substring(1)).test(field[property]) == true) {
									field = null;
									break;
								}
							}
						else if (field[property] == fieldExclusions[property])
							field = null;
						if (field == null)
							break;
					}
					if (field == null)
						continue;
					fieldDefs.push({
						"Title": field.Title,
						"InternalName": field.InternalName,
						"FieldTypeKind": field.FieldTypeKind,
						"Required": field.Required,
						"Group": field.Group
					});
				}
				// field information assembled, now time to create
				if (typeof destLib == "string")
					destLib = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: destLib
					});
				if (destLib == null)
					reject("Destination library not found in site: must create library before copying fields to it.");
				destLib.createFields(fieldDefs).then((response: any) => {
					resolve(response);
				}).catch((response: any) => {
					reject(response);
				});
			}).catch((response: any) => {
				reject(response);
			});
		});
	}

	copyFolders(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST
	): Promise<any> {
		return new Promise((resolve, reject) => {
			if (typeof sourceLib == "string")
				sourceLib = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: sourceLib
					});
			if (sourceLib == null)
				reject("Source lib not found in site");
			sourceLib.getListItems({
				expand: "Folder",
				select: "Folder/Name,Folder/ServerRelativeUrl"
			}).then((response: any) => {
				let folders: any[] = [],
						results: any[] = response.data ? response.data : response,
						requests: IBatchHTTPRequestForm[] = [],
						returnvalue: string = "";

				for (let item of results)
					if (item.FileSystemObjectType == 1) { // folder
						item.folderLevel = (item.ServerRelativeUrl.match(/\//g) || []).length;
						folders.push(item);
					}
				// sort out the folders
				folders.sort((el1, el2) => {
					return el1.folderLevel > el2.folderLevel ? 1 : el1.folderLevel < el2.folderLevel ? -1 : 0;
				});
				if (typeof destLib == "string")
					destLib = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: destLib
					});
				if (destLib == null)
					reject("Destination library not found in site: must create library before copying fields to it.");

				for (let item of folders)
					requests.push({
						url: destLib.forApiPrefix + "/web/folders",
						method: "POST",
						body: {
							"__metadata": {
							  "type": "SP.Folder"
							},
							"ServerRelativeUrl": item.serverRelativeUrl
						},
						contextinfo: ""
					});
					batchRequestingQueue(
					{host: destLib.server, path: destLib.site,
						AllHeaders: SPSiteREST.stdHeaders, AllMethods: "POST"},
					requests
				).then((response: any) => {
					resolve(response);
				}).catch((response: any) => {
					reject(response);
				});
			}).catch((response: any) => {
				reject(response);
			})
		});
	}

	copyFiles(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST,
		maxitems?: number
	): Promise<any> {
		return new Promise((resolve, reject) => {
			if (typeof sourceLib == "string")
				sourceLib = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: sourceLib
					});
			if (sourceLib == null)
				reject("Source lib not found in site");
			sourceLib.init().then((response: any) => {
			if (typeof destLib == "string")
				destLib = new SPListREST({
					server: this.server,
					site: this.sitePath,
					listName: destLib
				});
			if (destLib == null)
				reject("Destination library not found in site: must create library before copying fields to it.");
			destLib.init().then((response: any) => {
			(sourceLib as SPListREST).getAllListItems().then((response: any) => {
				let parts: RegExpMatchArray,
					requests: {sourceUrl: string; destUrl: string; fileName: string}[] = [];

				for (let item of response.data)
					if (item.File.ServerRelativeUrl) {
						parts = item.File.ServerRelativeUrl.match(/^(.*\/)([^\/]+)$/);
						requests.push({
							sourceUrl: item.File.ServerRelativeUrl,
							destUrl: (destLib as SPListREST).serverRelativeUrl,
							fileName: parts[2]
						});
					}
				this.workRequests(requests, 0).then((response) => {
					resolve(true);
				}).catch((response) => {
					reject(response);
				});
			/*
					body = { };
					for (let field of fieldDefs)
						body[field.InternalName] = item[field.InternalName];
					itemMetadata.push({body:body});
				iDestListREST.createListItems(itemMetadata, true).then((response) => {
					resolve(response);
				}).catch((response) => {
					reject(response);
				}); */
			});
			});
			});
		});
	}

	workRequests(requests: {
		sourceUrl: string;
		destUrl: string;
		fileName: string
	}[], index: number): Promise<any> {
		return new Promise((resolve, reject) => {
			$.ajax({
				url: this.apiPrefix + "/web/getFileByServerRelativeUrl('" +
					requests[index].sourceUrl + "')/copyto(strnewurl='" +
					requests[index].destUrl + "/" +
					requests[index].fileName + "',boverwrite=false)",
				method: "POST",
				success: (data) => {
					if (index + 1 > requests.length)
						resolve("All successful: Completed " + (index + 1) + " requests");
					else
						this.workRequests(requests, index + 1);
				},
				error: (reqObj) => {
					reject("Completed up to " + (index + 1) + " requests.\n\n" + reqObj);
				}
			});
		});
	}

	copyMetadata(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST
	): Promise<any> {
		return new Promise((resolve, reject) => {
			resolve(true);
			reject();
		});
	}

	copyViews(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST
	): Promise<any> {
		return new Promise((resolve, reject) => {
			if (typeof sourceLib == "string")
			sourceLib = new SPListREST({
					server: this.server,
					site: this.sitePath,
					listName: sourceLib
				});
			if (sourceLib == null)
				reject("Source lib not found in site");
			if (typeof destLib == "string")
				destLib = new SPListREST({
					server: this.server,
					site: this.sitePath,
					listName: destLib
				});
			if (destLib == null)
				reject("Destination library not found in site: must create library before copying fields to it.");
			(sourceLib as SPListREST).getView().then((response: any) => {
				let viewDefs = [ ];

				for (let sourceView of response.data.d.results)
					viewDefs.push({
						"Title": sourceView.Title,
						"ViewType": sourceView.ViewType,
						"PersonalView": sourceView.PersonalView,
						"ViewQuery": sourceView.ViewQuery,
						"RowLimit": sourceView.RowLimit,
						"ViewFields": sourceView.ViewFields,
						"DefaultView": sourceView.DefaultView
					});
				(destLib as SPListREST).createViews(viewDefs).then(() => {
					resolve(true);
				}).catch((response: any) => {
					reject(response);
				});
			}).catch((response: any) => {
				reject(response);
			});
		});
	}
}

SPSiteREST.stdHeaders = {
	"Accept": "application/json;odata=verbose",
	"Content-Type": "application/json;odata=verbose"
};