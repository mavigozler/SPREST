/**
 *
 * @class  representing parameters to establish interface with a site
 *             {server: name of location.host, site: path to site}
 */
declare class SPSearchREST {
    server: string;
    apiPrefix: string;
    static stdHeaders: THttpRequestHeaders;
    constructor(server: string);
    static httpRequest(elements: THttpRequestParams): void;
    static httpRequestPromise(parameters: THttpRequestParamsWithPromise): Promise<unknown>;
    queryText(parameters: {
        query?: string;
        querytext?: string;
    }): Promise<unknown>;
}
