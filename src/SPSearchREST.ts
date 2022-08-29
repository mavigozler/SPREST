"use strict";

/**
 *
 * @class  representing parameters to establish interface with a site
 *             {server: name of location.host, site: path to site}
 */
class SPSearchREST {
	server: string;
	apiPrefix: string;
	static stdHeaders: THttpRequestHeaders = {
		"Accept": "application/json;odata=verbose",
		"Content-Type": "application/json;odata=verbose"
	};

	constructor(server: string) {
	   if (!server || typeof server != "string")
			throw "Use of SPSearchREST() constructor must include string parameter specifying URL to server";
		this.server = server;
   	this.apiPrefix = this.server + "/_api";
	}
	// if parameters.success not found the request will be SYNCHRONOUS and not ASYNC
	static httpRequest(elements:THttpRequestParams) {
		if (elements.setDigest && elements.setDigest == true) {
			let match = elements.url.match(/(.*\/_api)/);
			$.ajax({  // get digest token
				url: match![1] + "/contextinfo",
				method: "POST",
				headers: {...SPSearchREST.stdHeaders},
				success: function (data: any) {
						 let headers: THttpRequestHeaders = elements.headers as THttpRequestHeaders;

						 if (elements.headers) {
							  headers["Content-Type"] = "application/json;odata=verbose";
							  headers["Accept"] = "application/json;odata=verbose";
						 } else
							  headers = {...SPSearchREST.stdHeaders};
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
				error: (requestObj: any, status: string, thrownErr: string) => {
					elements.errorCallback!(requestObj, status, thrownErr);
				}
			});
		} else {
			  if (!elements.headers)
				 elements.headers = {...SPSearchREST.stdHeaders};
			 else {
				 elements.headers["Content-Type"] = "application/json;odata=verbose";
				 elements.headers["Accept"] = "application/json;odata=verbose";
			 }
			$.ajax({
				url: elements.url,
				method: elements.method,
				headers: elements.headers,
				data: elements.body ?? elements.data as string,
				success: (data: TSPResponseData, status: string, requestObj: JQueryXHR) => {
					if (data.d && data.__next)
					RequestAgain(
							elements,
							data.__next,
							data.results!
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

	static httpRequestPromise(parameters: THttpRequestParamsWithPromise) {
		return new Promise((resolve, reject) => {
			SPSearchREST.httpRequest({
				setDigest: parameters.setDigest,
				url: parameters.url,
				method: parameters.method ? parameters.method : "GET",
				headers: parameters.headers,
				data: parameters.data ?? parameters.body as TXmlHttpRequestData,
				successCallback: (data, message, reqObj) => {
					resolve({data: data, message: message, reqObj: reqObj});
				},
				errorCallback: (reqObj, text, errThrown) => {
					reject({reqObj: reqObj, text: text, errThrown: errThrown});
				}
			});
		});
	}

   queryText(parameters: {
		query?: string;
		querytext?: string;
	}) {
      let querytext = "";

      if (parameters.query)
         querytext = "?querytext='" + parameters.query + "'";
      else if (parameters.querytext)
         querytext = "?querytext='" + parameters.querytext + "'";
      return SPSearchREST.httpRequestPromise({
         url: this.apiPrefix + "/search/query" + querytext
      });
   }
}