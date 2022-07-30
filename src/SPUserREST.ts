"use strict";

/*  SP User via REST

How to get the current user:
	getSharePointCurrentUserInfo().then(function (response) {
		response is UserInfo object!
	});
How to get info on another user by ID:
	getSharePointUserInfo({
		userId: <id-value>
	}).then(function (response) {
		response is UserInfo object!
	});

How to get info on another user by name:
	getSharePointUserInfo({
		firstName: "firstName",
		lastName: "lastName"
	}).then(function (response) {
		response is UserInfo object!
	});
How to get info on all users through siteuserinfolist:
	getAllSharePointUsersInfo({
		server: [required]
		query: [ optional ]
	}).then(function (response) {
		response is results array of all users
	});

*/


/** @function UserInfo -- class to store user information from the SharePoint system
 *  @param  {object} search
 *       {object} [required] any of the following property combinations will be evaluated
 *          {} - empty object: current user info will be returned
 *                                      {.userId: <valid-numeric-ID-for-user> }  user whose ID is used will have info returned
 *                                      {.lastName: "<last-name>" }   first user with that last name will be returned
 *                                      {.lastName: "<last-name>", .firstName: "<first-name>" }
 *          {.debugging: true...sets for debugger}
 */
 class UserInfo {
	search: TUserSearch;
	userId: number = -1; // numeric SP ID
	loginName: string = ""; //i:0#.w|domain\\user name
	title: string = ""; // "<last name>, <first name> <org>
	emailAddress: string = "";
	userName: string = "";
	firstName: string = "";
	lastName: string = "";
	workPhone: string = "";
	created: Date = new Date(0);
	modified: Date = new Date(0);
	jobTitle: string = "";
	dataComplete: boolean = false;
	//        storeData = storeDataEx.bind(null, this);
	constructor(search: TUserSearch) {
		if (!search || typeof search != "object")
			throw "Input to the constructor must include a search property with defined object: { .search: {...} }";
		this.search = search;
		if (search.id)
			this.search.userId = search.id;
	}

	populateUserData (userRestReqObj: SPUserREST): Promise<any> {
		return new Promise((resolve, reject) => {
			let type: number,
				args: {
					userId?: number;
					lastName?: string;
					firstName?: string;
				} = {};

			if (this.search.userId) {
				type = requestType.TYPE_ID;
				args = {
					userId: this.search.userId
				};
			} else if (this.search.lastName)
			   if (this.search.firstName) {
					type = requestType.TYPE_FULL_NAME;
					args = {
						lastName: this.search.lastName,
						firstName: this.search.firstName
					};
				} else {
					type = requestType.TYPE_LAST_NAME;
					args = {
						lastName: this.search.lastName
					};
				}
			else  {
				type = requestType.TYPE_CURRENT_USER;
			}
			userRestReqObj.requestUserInfo({
					uiObject: this,
					type: type,
					args: args
			}).then((response) => {
				resolve(UserInfo.storeData(response.uiObject, response.result));
			}).catch((response) => {
				reject(response);
			});
		});
	};

	static storeData(userInfoObj: any, userData: any) {
		userInfoObj.email = userInfoObj.emailAddress = userData.EMail;
		userInfoObj.userId = userData.Id;
		userInfoObj.loginName = userData.Name;
		userInfoObj.title = userData.Title;
		userInfoObj.jobTitle = userData.JobTitle;
		userInfoObj.lastName = userData.LastName;
		userInfoObj.firstName = userData.FirstName;
		userInfoObj.workPhone = userData.WorkPhone;
		userInfoObj.userName = userData.UserName;
		userInfoObj.created = userData.Created;
		userInfoObj.modified = userData.Modified;
		userInfoObj.dataComplete = true;
		return userInfoObj;
	}
	getFullName (): string {
		return this.firstName + " " + this.lastName;
	};
	getLastName (): string {
		return this.lastName;
	};
	getFirstName (): string {
		return this.firstName;
	};
	getUserId (): number {
		return this.userId;
	};
	getUserLoginName (): string {
		return this.loginName;
	};
	getUserEmail (): string {
		return this.emailAddress;
	};
}

/** @function SPUserREST -- class to set up REST interface to get user info on SharePoint server
 *  @param {object} the following properties are defined
 *        .protocol {string} [optional]...either "http" (default used) or "https"
 *        .server  {string} [required]   domain name of server
 *        .site {string} [optional]  site within server site collection, empty string is default
 *        .debugging {boolean}  if true, set to debugging
 */
 class SPUserREST {
	server: string;
	site: string;
	constructor(parameters: {
		server: string;
		site: string;
	}) {
		if (!parameters.server)
			throw "Input to the constructor must include a server property: { .server: <string> }";
		if (parameters.server.search(/^https?\:\/\/[^\/]+\/?/) != 0)
			throw "'parameters.server' does not appear to follow the pattern 'http[s]://host.name.com/. It must include protocol & host name";
		this.server = parameters.server;
		if (!(this.site = parameters.site))
			throw "A site path is required in 'parameters.site'";
		this.site = parameters.site;
	}
	/** @method  .requestUserInfo -- class to set up REST interface to get user info on SharePoint server
	 *  @param {object} has following properties
	 *                              .uiObject: {UserInfo object} [required] need to associate
	 *           .type: {numeric} required to be TYPE_ID, TYPE_LAST_NAME, TYPE_FULL_NAME
	 *           .args: {object} optional. depends on .type setting, properties should be
	 *                 .id: {numeric} ID of user on SP system
	 *                 .lastName: {string}  present if TYPE_LAST_NAME or TYPE_FULL_NAME
	 *                 .firstName: {string}  must be present if TYPE_FULL_NAME
	 */
	requestUserInfo (parameters: {
		uiObject: UserInfo;
		type: number;
		args?: {
			userId?: number;
			lastName?: string;
			firstName?: string;
		}
	}): Promise<any> {
		if (parameters.uiObject instanceof UserInfo == false)
			throw "SPUserREST.requestUserInfo(): missing 'uiObject' parameter or parameter not UserInfo class";
		return new Promise((resolve, reject) => {
			if (parameters.type == requestType.TYPE_CURRENT_USER)
				this.processRequest({
					url: this.server + this.site + "/_api/web/currentuser"
				}).then((response) => {
					this.processRequest({
						url: this.server + this.site + "/_api/web/siteuserinfolist/items(" + response.d.Id + ")"
					}).then((response) => {
						resolve({
							uiObject: parameters.uiObject,
							result: response.d
						});
					}).catch((response) => {
						reject(response);
					});
				}).catch((response) => {
					reject(response);
				});
			else if (parameters.type == requestType.TYPE_ID)
				this.processRequest({
					url: this.server + this.site + "/_api/web/siteuserinfolist/items(" + parameters.args!.userId + ")"
				}).then((response) => {
					resolve({
						uiObject: parameters.uiObject,
						result: response.d
					});
				}).catch((response) => {
					reject(response);
				});
			else { // type == TYPE_FULL_NAME or TYPE_LAST_NAME
				let filter: string = "$filter=lastName eq '" + parameters.args!.lastName + "'";
				if (parameters.type == requestType.TYPE_FULL_NAME)
					filter += " and firstName eq '" + parameters.args!.firstName + "'";
				this.processRequest({
					url: this.server + this.site + "/_api/web/siteuserinfolist/items?" + filter
				}).then((response) => {
					resolve({
						uiObject: parameters.uiObject,
						result: response.d
					});
				}).catch((response) => {
					reject(response);
				});
			}
		});
	};

	processRequest (parameters: {
		url: string;
		method?: string;
	}): Promise<any> {
		return new Promise((resolve, reject) => {
			$.ajax({
				method: parameters.method ? parameters.method : "GET",
				url: parameters.url,
				headers: {
					"Content-Type": "application/json;odata=verbose",
					"Accept": "application/json;odata=verbose"
				},
				success: (data) => {
					resolve(data);
				},
				error: (reqObj) => {
					reject(reqObj);
				}
			});
		});
	};
}

function getSharePointCurrentUserInfo(parameters: {
	server: string;
	site: string;
}): Promise<any> {
	return new Promise((resolve, reject) => {
		let iUserRequest = new SPUserREST({
				server: parameters.server,
				site: parameters.site
			}),
			uInfo = new UserInfo({});
		uInfo.populateUserData(iUserRequest).then((response) => {
			resolve(response);
		}).catch((response) => {
			reject(response);
		});
	});
}

function getSharePointUserInfo(parameters: {
	server: string;
	site: string;
	userId: number;
	firstName: string;
	lastName: string;
}): Promise<any> {
	return new Promise((resolve, reject) => {
		let uInfo, iUserRequest = new SPUserREST({
				server: parameters.server,
				site: parameters.site
			});
		if (parameters.userId)
			uInfo = new UserInfo({
				userId: parameters.userId
			});
		else
			uInfo = new UserInfo({
				firstName: parameters.firstName,
				lastName: parameters.lastName
			});
		uInfo.populateUserData(iUserRequest).then((response) => {
			resolve(response);
		}).catch((response) => {
			reject(response);
		});
	});
}

let CollectedResults: any[];
function getAllSharePointUsersInfo(parameters: {
	url?: string;
	server?: string;
	site?: string;
	query?: string;
}) {
	return new Promise((resolve, reject) => {
		$.ajax({
			method: "GET",
			url: parameters.url ? parameters.url : "https://" + parameters.server + parameters.site +
					"/_api/web/siteuserinfolist/items" +
					(parameters.query ? "?" + parameters.query : ""),
			headers: {
				"Content-Type": "application/json;odata=verbose",
				"Accept": "application/json;odata=verbose"
			},
			success: (data: TSPResponseData) => {
				if (!CollectedResults)
					CollectedResults = data.d!.results as any[];
				else
					CollectedResults = CollectedResults.concat(data.d!.results);
				if (data.d!.__next)
					resolve(getAllSharePointUsersInfo({
						url: data.d!.__next
					}));
				else
					resolve(data.d!.results);
			},
			error: (reqObj, responseStatus, responseMessage) => {
				reject(reqObj);
			}
		});
	});
}
