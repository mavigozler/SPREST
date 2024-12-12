/* eslint-disable @typescript-eslint/no-explicit-any */
"use strict";

import {
	constructQueryParameters,
//	formatRESTBody,
	checkEntityTypeProperty,
	ParseSPUrl,
} from "./SPRESTSupportLib";
import {
	RESTrequest, batchRequestingQueue
} from "./SPHttpReqResp";
import {
	// THttpRequestHeaders, THttpRequestParams, 	 THttpRequestParamsWithPromise, THttpRequestBody
	TSPResponseData, TXmlHttpRequestData,	IBatchHTTPRequestForm,  TArrayToJSON, HttpInfo
} from "./SPHttp";
import {
	SPListREST
} from "./SPListREST";
import {
	TSiteInfo, TParsedURL, SPWebInformationRaw, SPSiteRaw, SPWebRaw,
	SPFieldRaw, SPContentTypeRaw, SPListRaw, TSpField, TSiteQuickInfo
	// TSiteQuickInfo,
} from "./SPComponentTypes.d";


export {
	SPSiteREST
};

/**
 * @class SPSiteREST
 */
// export
class SPSiteREST {
	server: string;
	sitePath: string;
	forApiPrefix: string;
	forApiPrefix: string;
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
	Title: string = "";
	Description: string = "";
	Url: string = "";
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
		this.forApiPrefix = this.server + this.sitePath;
		this.forApiPrefix = this.server + this.sitePath;
	}

	/**
	 * @method create -- sets many list characteristics whose data is required
	 * 	by asynchronous request
	 * @returns Promise
	 */
	init(): Promise<SPSiteREST> {
	init(): Promise<SPSiteREST> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				url:  this.forApiPrefix + "/_api/site",
				successCallback: (data) => {
					const d: SPSiteRaw = data.d as SPSiteRaw;
					this.id = d.Id;
					this.serverRelativeUrl = d.ServerRelativeUrl;
					this.isHubSite = d.IsHubSite;
					this.webId = this.webGuid = d.Id;
					RESTrequest({
						url: this.forApiPrefix + "/_api/web",
						successCallback: (data) => {
							const webD: SPWebRaw = data.d as SPWebRaw;
							this.creationDate = new Date(webD.Created);
							this.siteName = webD.Title;
							this.homePage = webD.WelcomePage;
							this.siteLogoUrl = webD.SiteLogoUrl;
							this.template = webD.WebTemplate;
							resolve(this);
						},
						errorCallback: (err) => {
							reject(err);
						}
					});
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	} // _init()

	/**
	 *
	 * @param {*} parameters -- object which may contain select, filter, expand properties
	 */
	getProperties(parameters?: any): Promise<unknown> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/site" + constructQueryParameters(parameters),
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	getSiteProperties(parameters?: any) {
		return this.getProperties(parameters);
	}

	getWebProperties(parameters?: any): Promise<SPWebRaw> {
		return new Promise<SPWebRaw>((resolve, reject) => {
			return RESTrequest({
				url: this.forApiPrefix + "/_api/web" + constructQueryParameters(parameters),
				successCallback: (data) => {
					resolve(data as SPWebRaw);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	getEndpoint(endpoint: string): Promise<unknown> {
		return new Promise<unknown>((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/" + endpoint,
				successCallback: (data) => {
					resolve(data);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	WebTemplateIDs = [
		{ ID: "STS#3",  name: "Team Site (no M635 Group)", group: "Collaboration" },
		{ ID: "STS#0",  name: "Team Site (classic experience)", group: "Collaboration" },
		{ ID: "PROJECTSITE#0",  name: "Project Site", group: "Collaboration" },
		{ ID: "BDR#0",  name: "Document Center", group: "Enterprise" },
		{ ID: "OFFILE#1",  name: "Records Center", group: "Enterprise"  },
		{ ID: "SRCHCENTERLITE#0",  name: "Basic Search Center", group: "Enterprise"  },
		{ ID: "visprus#0",  name: "Visio Process Repository", group: "Enterprise"  },
		{ ID: "SAPWorkflowSite#0",  name: "SAP Workflow Site", group: "Duet Enterprise"  }
	];

	createSubsite(
		title: string,
		description: string | undefined,
		url: string | undefined,
		webTemplateID: string,
		useUniquePermissions: boolean
	): Promise<TSiteInfo> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				url: this.server + this.sitePath + url,
				method: "POST",
				body: {  // this is actually TSpList type
					__metadata: { type: "SP.Site" },
					Language: 1033,
					BaseTemplate: webTemplateID,
					UseUniquePermissions: useUniquePermissions,
					Description: description,
					Title: title
				} as unknown as TXmlHttpRequestData,
				successCallback: (response/* , httpInfo */) => {
					resolve(response as unknown as TSiteInfo);
				},
				errorCallback: (errInfo/* , httpInfo */) => {
					reject(errInfo);
				}
			});
		});
	}

	getSubsites(): Promise<TSiteInfo[]> {
		return new Promise<TSiteInfo[]>((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/webinfos",
				successCallback: (data: TSPResponseData) => {
					const sites: TSiteInfo[] = [ ],
						results: SPWebInformationRaw[] = data.d!.results as SPWebInformationRaw[];

					for (const result of results!)
						sites.push({
							Name: result.Title,
							ServerRelativeUrl: result.ServerRelativeUrl,
							Id: "", // TODO  fix this result.OtherId,
							Template: result.WebTemplate,
							Title: result.Title,
							Url: "", //  TODO fix this result.Url,
							Description: result.Description,
							Created: new Date(result.Created),
							tsType: "TSiteInfo",
							siteParent: undefined,
							forApiPrefix: this.server + result.ServerRelativeUrl
						}) as unknown as TSiteInfo;
					resolve(sites);
				},
				errorCallback: (errInfo) => {
					reject(errInfo);
				}
			});
		});
	}

	getParentWeb(siteUrl: string): Promise<SPWebInformationRaw | null> {
		return new Promise<SPWebInformationRaw | null>((resolve, reject) => {
			RESTrequest({
				url: siteUrl + "/_api/web?$expand=ParentWeb",
				successCallback: (data) => {
					const spData: SPWebRaw = data as SPWebRaw;
					if (spData.ParentWeb.__metadata)
						resolve(spData.ParentWeb);
					else
						resolve(null);
				},
				errorCallback: (errInfo) => {
					reject(errInfo);
				}
			});
		});
	}

	getSitePedigree(siteUrl: string | null):
				Promise<{pedigree: TSiteInfo; arrayPedigree: TSiteInfo[]}> {
	getSitePedigree(siteUrl: string | null):
				Promise<{pedigree: TSiteInfo; arrayPedigree: TSiteInfo[]}> {
		return new Promise((resolve, reject) => {
			const pedigree: TSiteInfo = {} as TSiteInfo,

				siteInfo: TSiteInfo = {} as TSiteInfo,
				repeatGetParent = (siteInfo: TSiteInfo) => {
					this.getParentWeb(siteInfo.server! + siteInfo.ServerRelativeUrl).then((ParentWeb) => {
						if (ParentWeb) {  // != null
							const siteParent: TSiteInfo = {} as TSiteInfo;
							siteParent.Title = ParentWeb.Title;
							siteParent.server = siteInfo.server;
							siteParent.ServerRelativeUrl = ParentWeb.ServerRelativeUrl;
							siteParent.Id = ParentWeb.Id;
							siteParent.Template = ParentWeb.WebTemplate;
							siteInfo.parent = siteParent;
							siteParent.subsites = [];
							siteParent.subsites.push(siteInfo as TSiteQuickInfo);
							repeatGetParent(siteParent);
						} else {
							this.fillOutPedigree(siteInfo).then(() => {
								resolve({pedigree: pedigree, arrayPedigree: this.arrayPedigree});
							}).catch((response: any) => {
								reject(response);
							});  // use the last site
						}
					}).catch((response: any) => {
						reject(response);
					});
				};

			if (!siteUrl) {
				siteInfo.server = this.server;
				siteInfo.ServerRelativeUrl = this.serverRelativeUrl;
				siteInfo.ServerRelativeUrl = this.serverRelativeUrl;
			} else {
				const parsedUrl: TParsedURL | null = //SPRESTSupportLib.
					ParseSPUrl(siteUrl);

				if (parsedUrl == null)
					throw "Parameter 'siteUrl' was not a parseable SharePoint URL";
				siteInfo.server = parsedUrl.server;
				siteInfo.ServerRelativeUrl = parsedUrl.siteFullPath;
				siteInfo.ServerRelativeUrl = parsedUrl.siteFullPath;
			}
			pedigree.referenceSite = {
				///pedigree.referenceSite = {
				Name: this.Title,
				server: siteInfo.server,
				ServerRelativeUrl: this.serverRelativeUrl,
				Id: this.id,
				Template: this.template,
				ServerRelativeUrl: this.serverRelativeUrl,
				Id: this.id,
				Template: this.template,
				parent: null,
				subsites: [],
				Title: this.Title,
				typeInSP: "SUBSITE",
				tsType: "TSiteInfo",
				Url: this.Url,
				Description: this.Description,
				siteParent: undefined,
				forApiPrefix: siteInfo.server + this.serverRelativeUrl
				subsites: [],
				Title: this.Title,
				typeInSP: "SUBSITE",
				tsType: "TSiteInfo",
				Url: this.Url,
				Description: this.Description,
				siteParent: undefined,
				forApiPrefix: siteInfo.server + this.serverRelativeUrl
			};
			this.arrayPedigree = [];
			this.arrayPedigree.push(pedigree.referenceSite);
			//repeatGetParent(pedigree.referenceSite!);
			repeatGetParent(pedigree.referenceSite!);
		});
	}

	fillOutPedigree (parentWeb: TSiteInfo): Promise<any> {
		return new Promise((resolve, reject) => {
			if (!parentWeb)
				throw "fillOutPedigree():  parameter 'parentWeb' is not defined";
			const subSite = new SPSiteREST({
				server: this.server as string,
				site: parentWeb.ServerRelativeUrl as string
			});
			subSite.init().then(() => {
				subSite.getSubsites().then((webInfos: TSiteInfo[]) => {
				subSite.getSubsites().then((webInfos: TSiteInfo[]) => {
					if (typeof webInfos == "undefined")
						reject("Promise.success() returned nothing");
					else {
						if (typeof parentWeb.subsites == "undefined")
							parentWeb.subsites = [];
						if (webInfos.length > 0) {
							let idx: number;
							const webInfoRequests: Promise<any>[] = [];

							for (const webinfo of webInfos) {
								idx = parentWeb.subsites.push({
									Name: webinfo.Title,
									Title: webinfo.Title,
									server: parentWeb.server,
									ServerRelativeUrl: webinfo.ServerRelativeUrl,
									Id: webinfo.Id,
									Template: webinfo.Template,
									subsites: [],
									typeInSP: "SUBSITE",
									tsType: "TSiteInfo",
									Url: this.Url,
									Description: this.Description,
									siteParent: parentWeb,
									forApiPrefix: parentWeb.server! + webinfo.ServerRelativeUrl
									ServerRelativeUrl: webinfo.ServerRelativeUrl,
									Id: webinfo.Id,
									Template: webinfo.Template,
									subsites: [],
									typeInSP: "SUBSITE",
									tsType: "TSiteInfo",
									Url: this.Url,
									Description: this.Description,
									siteParent: parentWeb,
									forApiPrefix: parentWeb.server! + webinfo.ServerRelativeUrl
								}) - 1;
								this.arrayPedigree.push((parentWeb.subsites as TSiteInfo[])[idx]);
								webInfoRequests.push(this.fillOutPedigree((parentWeb.subsites as TSiteInfo[])[idx]));
								this.arrayPedigree.push((parentWeb.subsites as TSiteInfo[])[idx]);
								webInfoRequests.push(this.fillOutPedigree((parentWeb.subsites as TSiteInfo[])[idx]));
							}

							Promise.all(webInfoRequests).then((response: TSiteInfo[]) => {
								resolve(response);
							}).catch((response: HttpInfo) => {
								reject(response);
							});
						} else
							resolve(true);
					}
				}).catch((response: any) => {
				}).catch((response: any) => {
					reject(response);
				});
			});
		});
	}

	getSiteColumns(parameters: any): Promise<SPFieldRaw[]> {
		return new Promise<SPFieldRaw[]>((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/fields" + constructQueryParameters(parameters),
				successCallback: (data) => {
					const results: SPFieldRaw[] = data.d?.results as SPFieldRaw[];
					resolve(results);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	getSiteContentTypes(parameters: any): Promise<SPContentTypeRaw[]> {
		return new Promise<SPContentTypeRaw[]>((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/ContentTypes" + constructQueryParameters(parameters),
				successCallback: (data) => {
					const results: SPContentTypeRaw[] = data.d?.results as SPContentTypeRaw[];
					resolve(results);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	getLists(parameters?: any):Promise<SPListRaw[]> {
		return new Promise<SPListRaw[]>((resolve, reject) => {
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/lists" + constructQueryParameters(parameters),
				successCallback: (data) => {
					const results: SPListRaw[] = data.d?.results as SPListRaw[];
					resolve(results);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	/**
	 *
	 * @param {*} parameters -- need to control these!
	 * @returns
	 */
	createList(
		title: string,
		listTemplate: number,
		description: string,
		allowContentTypes: boolean,
		enableContentTypes: boolean
	): Promise<boolean> {
		return new Promise<boolean>((resolve, reject) => {
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists",
				method: "POST",
				body: {  // this is actually TSpList type
					__metadata: { type: "SP.List" },
					AllowContentTypes: allowContentTypes,
					BaseTemplate: listTemplate,
					ContentTypesEnabled: enableContentTypes,
					Description: description,
					Title: title
				} as unknown as TXmlHttpRequestData,
				successCallback: (data) => {
					resolve(data as unknown as boolean);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	/**
	 * @method updateListByMerge
	 * @param {object} parameters -- need a 'Title' parameter
	 * @returns
	 */
	updateListByMerge(parameters: {body: TArrayToJSON, listGuid: string}): Promise<boolean> {
		const body: TArrayToJSON | string = parameters.body;

		if (checkEntityTypeProperty(body, "item") == false)
			body["__SetType__"] = "SP.List";
	//	body = formatRESTBody(body);
		return new Promise<boolean>((resolve, reject) => {
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists" +
					"(guid'" + parameters.listGuid + "')", /// this cannot be correct
				method: "POST",
				headers: {
					"IF-MATCH": "*",
					"X-HTTP-METHOD": "MERGE"
				},
				body: body as TXmlHttpRequestData,
				successCallback: (data) => {
					resolve(data as unknown as boolean);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
		});
	}

	deleteList(parameters: {id?: string, guid?: string}) {
		const id = parameters.id ?? parameters.guid;

		if (!id)
			throw "List deletion requires an 'id' or 'guid' parameter for the list to be deleted.";
		return new Promise<boolean>((resolve, reject) => {
			RESTrequest({
				setDigest: true,
				url: this.forApiPrefix + "/_api/web/lists(guid'" + id + "')",
				method: "POST",
				headers: {
					"X-HTTP-METHOD": "DELETE",
					"IF-MATCH": "*" // can also use etag
				},
				successCallback: (data) => {
					resolve(data as unknown as boolean);
				},
				errorCallback: (err) => {
					reject(err);
				}
			});
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
			const sourceListREST = new SPListREST({
				server: this.server,
				site: this.sitePath,
				listName: sourceLibName
			});
			sourceListREST.init().then(() => {
				this.createList(
					destLibName,
					sourceListREST.baseTemplate,
					"Copied list",
					sourceListREST.allowContentTypes,
					false
				).then((response: any) => {
					const iDestListREST = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: response.data.d.Title
					});
					iDestListREST.init().then(() => {
						this.copyFields(sourceListREST, iDestListREST)
							.then(() => {
								this.copyFolders(sourceListREST, iDestListREST)
									.then(() => {
										this.copyFiles(sourceListREST, iDestListREST)
											.then((response: any) => {
												resolve(response);
											}).catch((err: unknown) => {
												// copyFiles
												reject(err);
											});
									}).catch((err: unknown) => {
										// copyFolders
										reject(err);
									});
							}).catch((err: unknown) => {
							// copyFields
								reject(err);
							});
					}).catch((err: any) => {
						reject(err);
					});
				}).catch((err: any) => {
					//  createList
					if (err.status == 500 &&
							err.reqObj.responseJSON.error.message.value.search(
								/list.*already exists.*Please choose another title\./
							) >= 0)
						sourceListREST
							.getListItems({
								top: itemCopyCount
							}).then(() => {
							/*
								iDestListREST.createListItems(response.data.d.ressults).then((response: any) => {
								iDestListREST.createListItems(response.data.d.ressults).then((response: any) => {
									resolve(response);
								}).catch((response: any) => {
									reject(response);
								}); */
							}).catch((err: unknown) => {
								// getListItems
								reject(err);
							});
					else {
						reject(err);
					}
				});
			}).catch((err: unknown) => {
				reject (err);
			});
		});
	}

	copyFields(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST
	): Promise<any> {
		return new Promise((resolve, reject) => {
			const fieldDefs: TSpField[] = [ ],
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
			sourceLib.getFields()
				.then((response: any) => {
					const sourceFields = response.data.d.results;

					for (let field of sourceFields) {
						for (const property in fieldExclusions) {
							if (Array.isArray(fieldExclusions[property]) == true)
								for (const elem of fieldExclusions[property] as (string[] | number[])) {
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
							Title: field.Title,
							InternalName: field.InternalName,
							FieldTypeKind: field.FieldTypeKind,
							Required: field.Required,
							Group: field.Group
						} as TSpField);
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
					destLib.createFields(fieldDefs)
						.then((response: any) => {
							resolve(response);
						}).catch((err: unknown) => {
							// problem with createFields()
							reject(err);
						});
				}).catch((err: unknown) => {
				// problem with getFields()
					reject(err);
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
				const folders: any[] = [],
					results: any[] = response.data ? response.data : response,
					requests: IBatchHTTPRequestForm[] = [];
				//,returnvalue: string = "";

				for (const item of results)
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

				for (const item of folders)
					requests.push({
						url: destLib.forApiPrefix + "/_api/web/folders",
						contextinfo: destLib.server + destLib.serverRelativeUrl,
						url: destLib.forApiPrefix + "/_api/web/folders",
						contextinfo: destLib.server + destLib.serverRelativeUrl,
						method: "POST",
						headers: SPSiteREST.stdHeaders,
						body: {
							"__metadata": {
								"type": "SP.Folder"
							},
							"ServerRelativeUrl": item.serverRelativeUrl
						}
					});
				batchRequestingQueue(requests).then((response: any) => {
					resolve(response);
				}).catch((err: unknown) => {
					reject(err);
				});
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}

	copyFiles(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST,
	//	maxitems?: number
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
			sourceLib.init().then(() => {
				if (typeof destLib == "string")
					destLib = new SPListREST({
						server: this.server,
						site: this.sitePath,
						listName: destLib
					});
				if (destLib == null)
					reject("Destination library not found in site: must create library before copying fields to it.");
				destLib.init().then(() => {
					(sourceLib as SPListREST).getAllListItems().then((response: any) => {
						let parts: RegExpMatchArray;
						const requests: {sourceUrl: string; destUrl: string; fileName: string}[] = [];

						for (const item of response.data)
							if (item.File.ServerRelativeUrl) {
								parts = item.File.ServerRelativeUrl.match(/^(.*\/)([^/]+)$/);
								requests.push({
									sourceUrl: item.File.ServerRelativeUrl,
									destUrl: (destLib as SPListREST).serverRelativeUrl,
									fileName: parts[2]
								});
							}
						this.workRequests(requests, 0).then(() => {
							resolve(true);
						}).catch((response: any) => {
							reject(response);
						});
						/*
					body = { };
					for (let field of fieldDefs)
						body[field.InternalName] = item[field.InternalName];
					itemMetadata.push({body:body});
				iDestListREST.createListItems(itemMetadata, true).then((response: any) => {
				iDestListREST.createListItems(itemMetadata, true).then((response: any) => {
					resolve(response);
				}).catch((response: any) => {
				}).catch((response: any) => {
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
			RESTrequest({
				url: this.forApiPrefix + "/_api/web/getFileByServerRelativeUrl('" +
					requests[index].sourceUrl + "')/copyto(strnewurl='" +
					requests[index].destUrl + "/" +
					requests[index].fileName + "',boverwrite=false)",
				method: "POST",
				successCallback: () => {
					if (index + 1 > requests.length)
						resolve("All successful: Completed " + (index + 1) + " requests");
					else
						this.workRequests(requests, index + 1);
				},
				errorCallback: (reqObj) => {
					reject("Completed up to " + (index + 1) + " requests.\n\n" + reqObj);
				}
			});
		});
	}

	// TODO   what was this supposed to be?
	/*
	copyMetadata(
		sourceLib: string | SPListREST,
		destLib: string | SPListREST
	): Promise<any> {
		return new Promise((resolve, reject) => {
			resolve(true);
			reject();
		});
	}*/

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
				const viewDefs: {
					"Title": string;
					"ViewType": string;
					"PersonalView": string;
					"ViewQuery": string;
					"RowLimit": string;
					"ViewFields": string;
					"DefaultView": string;
				}[] = [ ];

				for (const sourceView of response.data.d.results)
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
				}).catch((err: unknown) => {
					reject(err);
				});
			}).catch((err: unknown) => {
				reject(err);
			});
		});
	}
}

/*


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
						//SPRESTSupportLib.
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
			RESTrequest({
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

SPSiteREST.stdHeaders = {
	"Accept": "application/json;odata=verbose",
	"Content-Type": "application/json;odata=verbose"
}; */