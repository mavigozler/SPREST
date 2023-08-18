"use strict";

/*  SP User via REST

1. get instance of SPUserREST()
	supply the server name, and then the site name.  This will retrieve the users of the site
2. Use the class methods
	a) getSharePointCurrentUser()
	b) getSharePointSiteUser(Id)
	c) getSharePointSiteUser(LastName)
	d) getSharePointSiteUser(FirstName, LastName)
	e) getSharePointSiteUsers(query?: string;)
*/

type TUserSearch = {
	userId?: number;
	id?: number;
	firstName?: string;
	lastName?: string;
};

type UserData = {
	Email: string;
	Id: string | number;
	LoginName?: string;
	Title?: string;
	JobTitle?: string;
	LastName: string;
	FirstName: string;
	WorkPhone: string;
	UserName: string;
	Created?: Date;
};


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
	CollectedResults: any[] = [];

	constructor(
		server: string,
		site: string
	) {
		if (!server)
			throw "Input to the constructor must include a server property: { .server: <string> }";
		if (server.search(/^https?\:\/\/[^\/]+\/?/) != 0)
			throw "'parameters.server' does not appear to follow the pattern 'http[s]://host.name.com/. It must include protocol & host name";
		this.server = server;
		if (!(this.site = site))
			throw "A site path is required in 'parameters.site'";
		this.site = site;
	}

	requestUserInfo(args: {
		userId?: number | null,
		lastName?: string | null,
		firstName?: string | null
	}): Promise<{UserInfo: UserData; RawData: TSPResponseData}> {
		return new Promise((resolve, reject) => {
			let user: User = {} as User;

			if (!args.userId || !args.lastName)
				user = new User(
					this.server,
					this.site,
					{}
				);
			else if (args.userId)
				user = new User(
					this.server,
					this.site,
					{userId: args.userId}
				);
			else if (args.lastName)
				user = new User(
					this.server,
					this.site,
					{lastName: args.lastName, firstName: args.firstName ? args.firstName : undefined}
				);
			user.getUserData().then((response: {userInfo: UserData, rawData: TSPResponseData}) => {
				resolve({UserInfo: response.userInfo, RawData: response.rawData});
			}).catch((response: JQueryXHR) => {
				reject(response);
			});
		});
	}

	getSharePointCurrentUserInfo(): Promise<{UserInfo: UserData; RawData: TSPResponseData}> {
		return this.requestUserInfo({});
	}

	getSharePointUserInfoByLastName(lastName: string): Promise<{UserInfo: UserData; RawData: TSPResponseData}> {
		return this.requestUserInfo({lastName: lastName});
	}
	getSharePointUserInfoByUserId(userId: number): Promise<{UserInfo: UserData; RawData: TSPResponseData}> {
		return this.requestUserInfo({userId: userId});
	}
	getSharePointUserInfoByFullName(firstName: string, lastName: string): Promise<{UserInfo: UserData; RawData: TSPResponseData}> {
		return this.requestUserInfo({lastName: lastName, firstName: firstName});
	}

	getAllSharePointUsersInfo(query?: string): Promise<TSPResponseDataProperties[]> {
		return new Promise((resolve, reject) => {
			RESTrequest({
				url: `https://${this.server}${this.site}/_api/web/siteuserinfolist/items` + query ? "?" + query : "",
				method: "GET",
				successCallback: (data: TSPResponseData/*, text, reqObj */) => {
					if (!this.CollectedResults)
						this.CollectedResults = data.d!.results as any[];
					else
						this.CollectedResults = this.CollectedResults.concat(data.d!.results);
					if (data.d!.__next)
						resolve(this.getAllSharePointUsersInfo(data.d!.__next));
					else
						resolve(data.d!.results!);
				},
				errorCallback: (reqObj/*, status, errThrown */) => {
					reject(reqObj);
				}
			});
		});
	}
}

class User {
	server: string;
	site: string;
	search: TUserSearch;

	Email: string = "";
	Id: number = -1;
	LoginName: string = "";
	Title: string = "";
	JobTitle: string = "";
	LastName: string = "";
	FirstName: string = "";
	WorkPhone: string = "";
	UserName: string = "";
	Created: Date = new Date(0);

	RawUserData: TSPResponseData | null = null;
	//        storeData = storeDataEx.bind(null, this);
	constructor(
		server: string,
		site: string,
		search: TUserSearch
	) {
		this.server = server;
		this.site = site;
		this.search = search;
		if (search.id)
			this.search.userId = search.id;
	}

	getUserData(): Promise<{userInfo: UserData, rawData: TSPResponseData}> {
		return new Promise((resolve, reject) => {
			let url: string;

			if (this.search.userId)
				url = `${this.server}${this.site}/_api/web/siteuserinfolist/items(${this.search.userId})`;
			else if (this.search.lastName) {
				url = `${this.server}${this.site}/_api/web/siteuserinfolist/items?$filter=lastNName eq ${this.search.lastName}`;
				if (this.search.firstName)
					url += ` and firstName eq ${this.search.firstName}`;
			} else
				url = `${this.server}${this.site}/_api/web/currentuser`;
			RESTrequest({
				url: url,
				method: "GET",
				successCallback: (data: TSPResponseData/*, text, reqObj */) => {
					let userData: UserData = data.d as UserData;

					this.storeData(userData);
					resolve({userInfo: this, rawData: data.d as TSPResponseData})
				},
				errorCallback: (reqObj/*, status, errThrown */) => {
					reject(reqObj);
				}
			});
		});
	};

	storeData(userData: UserData) {
		this.Email = userData.Email;
		this.Id = userData.Id as number;
		this.LoginName = userData.LoginName!;
		this.Title = userData.Title as string;
		this.JobTitle = userData.JobTitle as string;
		this.LastName = userData.LastName;
		this.FirstName = userData.FirstName;
		this.WorkPhone = userData.WorkPhone;
		this.UserName = userData.UserName;
		this.Created = userData.Created as Date;
	}

	getFullName (): string {
		return this.FirstName + " " + this.LastName;
	};
	getLastName (): string {
		return this.LastName;
	};
	getFirstName (): string {
		return this.FirstName;
	};
	getUserId (): number {
		return this.Id;
	};
	getUserLoginName (): string {
		return this.LoginName;
	};
	getUserEmail (): string {
		return this.Email;
	};
}