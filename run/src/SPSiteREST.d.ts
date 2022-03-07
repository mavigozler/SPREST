/**
 * @class SPSiteREST
 */
declare class SPSiteREST {
    server: string;
    sitePath: string;
    apiPrefix: string;
    id: string;
    serverRelativeUrl: string;
    isHubSite: boolean;
    webId: string;
    webGuid: string;
    creationDate: Date;
    siteName: string;
    homePage: string;
    siteLogoUrl: string;
    template: string;
    static stdHeaders: THttpRequestHeaders;
    /**
     * @constructor SPSiteREST -- sets up properties and methods instance to describe a SharePoint site
     * @param {object} setup -- representing parameters to establish interface with a site
     *             {server: name of location.host, site: path to site}
     * @private
     */
    constructor(setup: {
        server: string;
        site: string;
    });
    /**
     * @method create -- sets many list characteristics whose data is required
     * 	by asynchronous request
     * @returns Promise
     */
    init(): Promise<unknown>;
    static httpRequest(elements: THttpRequestParams): void;
    static httpRequestPromise(parameters: THttpRequestParamsWithPromise): Promise<unknown>;
    /**
     *
     * @param {*} parameters -- object which may contain select, filter, expand properties
     */
    getProperties(parameters?: any): Promise<unknown>;
    getSiteProperties(parameters?: any): Promise<unknown>;
    getWebProperties(parameters?: any): Promise<unknown>;
    getEndpoint(endpoint: string): Promise<unknown>;
    getSubsites(): Promise<unknown>;
    getParentWeb(siteUrl: string): Promise<any>;
    getSitePedigree(siteUrl: string | null): Promise<any>;
    getSiteColumns(parameters: any): Promise<any>;
    getSiteContentTypes(parameters: any): Promise<any>;
    getLists(parameters?: any): Promise<any>;
    /**
     *
     * @param {*} parameters -- need to control these!
     * @returns
     */
    createList(parameters: {
        body: THttpRequestBody;
    }): Promise<unknown>;
    /**
     * @method updateListByMerge
     * @param {object} parameters -- need a 'Title' parameter
     * @returns
     */
    updateListByMerge(parameters: {
        body: THttpRequestBody;
        listGuid: string;
    }): Promise<unknown>;
    deleteList(parameters: {
        id?: string;
        guid?: string;
    }): Promise<unknown>;
    makeLibCopyWithItems(sourceLibName: string, destLibSitePath: string, destLibName: string, itemCopyCount?: number): Promise<any>;
    copyFields(sourceLib: string | SPListREST, destLib: string | SPListREST): Promise<any>;
    copyFolders(sourceLib: string | SPListREST, destLib: string | SPListREST): Promise<any>;
    copyFiles(sourceLib: string | SPListREST, destLib: string | SPListREST, maxitems?: number): Promise<any>;
    workRequests(requests: {
        sourceUrl: string;
        destUrl: string;
        fileName: string;
    }[], index: number): Promise<any>;
    copyMetadata(sourceLib: string | SPListREST, destLib: string | SPListREST): Promise<any>;
    copyViews(sourceLib: string | SPListREST, destLib: string | SPListREST): Promise<any>;
}
