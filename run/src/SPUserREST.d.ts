/** @function UserInfo -- class to store user information from the SharePoint system
 *  @param  {object} search
 *       {object} [required] any of the following property combinations will be evaluated
 *          {} - empty object: current user info will be returned
 *                                      {.userId: <valid-numeric-ID-for-user> }  user whose ID is used will have info returned
 *                                      {.lastName: "<last-name>" }   first user with that last name will be returned
 *                                      {.lastName: "<last-name>", .firstName: "<first-name>" }
 *          {.debugging: true...sets for debugger}
 */
declare class UserInfo {
    search: TUserSearch;
    userId: number;
    loginName: string;
    title: string;
    emailAddress: string;
    userName: string;
    firstName: string;
    lastName: string;
    workPhone: string;
    created: Date;
    modified: Date;
    jobTitle: string;
    dataComplete: boolean;
    constructor(search: TUserSearch);
    populateUserData(userRestReqObj: SPUserREST): Promise<any>;
    static storeData(userInfoObj: any, userData: any): any;
    getFullName(): string;
    getLastName(): string;
    getFirstName(): string;
    getUserId(): number;
    getUserLoginName(): string;
    getUserEmail(): string;
}
/** @function SPUserREST -- class to set up REST interface to get user info on SharePoint server
 *  @param {object} the following properties are defined
 *        .protocol {string} [optional]...either "http" (default used) or "https"
 *        .server  {string} [required]   domain name of server
 *        .site {string} [optional]  site within server site collection, empty string is default
 *        .debugging {boolean}  if true, set to debugging
 */
declare class SPUserREST {
    server: string;
    site: string;
    constructor(parameters: {
        server: string;
        site: string;
    });
    /** @method  .requestUserInfo -- class to set up REST interface to get user info on SharePoint server
     *  @param {object} has following properties
     *                              .uiObject: {UserInfo object} [required] need to associate
     *           .type: {numeric} required to be TYPE_ID, TYPE_LAST_NAME, TYPE_FULL_NAME
     *           .args: {object} optional. depends on .type setting, properties should be
     *                 .id: {numeric} ID of user on SP system
     *                 .lastName: {string}  present if TYPE_LAST_NAME or TYPE_FULL_NAME
     *                 .firstName: {string}  must be present if TYPE_FULL_NAME
     */
    requestUserInfo(parameters: {
        uiObject: UserInfo;
        type: number;
        args?: {
            userId?: number;
            lastName?: string;
            firstName?: string;
        };
    }): Promise<any>;
    processRequest(parameters: {
        url: string;
        method?: string;
    }): Promise<any>;
}
/**
 * @function getSharePointCurrentUserInfo -- return infor from SP server about current user
 * @param parameters the server hostname and the site
 * @returns Promise with response being several parameters
 *    email
 * 	userId
 * 	loginName
 * 	title
 * 	jobTitle
 * 	lastName
 * 	firstName
 * 	workPhone
 * 	userName
 * 	created
 * 	modified
 */
declare function getSharePointCurrentUserInfo(parameters: {
    server: string;
    site: string;
}): Promise<any>;
declare function getSharePointUserInfo(parameters: {
    server: string;
    site: string;
    userId: number;
    firstName: string;
    lastName: string;
}): Promise<any>;
declare let CollectedResults: any[];
declare function getAllSharePointUsersInfo(parameters: {
    url?: string;
    server?: string;
    site?: string;
    query?: string;
}): Promise<unknown>;
