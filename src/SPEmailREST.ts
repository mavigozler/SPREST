"use strict";

import type { TSPResponseData } from "./SPHttp.d.ts";
import { RESTrequest, SPstdHeaders } from "./SPHttpReqResp.js";

export {
	SPUtilityEmailService,
	emailDeveloper
};

/*
 *  requires use of jQuery AJAX object
 *  include the jquery.js script
 *
 *   USAGE
 *   1. instantiate the object SPUtilityEmailService with its
 *      two parameters 'server' (e.g. 'mdsd') and site (for a subsite in collection)
 *             var emailService = new SPUtilityEmailService({.server: , .site: });
 *   2. Set callbacks:
 *         .setSuccess(callbackFunc) for handling the success (HTTP 200 response)
 *         .setError(callbackFunc)  for handling the error
 *      Be sure to define the callbacks in the calling code file
 *      Both callbacks will get 3 parameters in this order
 *        i) responseData -- an object
 *       ii) returnStatusString -- a string that may say "success" vs "error"
 *      iii) reqObj -- usually the Xml HTTP Request object
 * *   3. call the .sendEmail(headers) method to send email
 *        headers is an object with several properties that are standard email headers
 *        for email address parameters, use a valid email address
 *        when there are multiple email addresses in a header, all addresses are
 *           to be separated by a SEMI-COLON!
 */

// import { TSPResponseData } from './SPComponentTypes';

interface IEmailHeaders {
	To: string;
	Subject: string;
	Body: string;
	CC?: string | null;
	BCC?: string | null;
	From?: string | null;
}

class SPUtilityEmailService {
	server: string;
	site: string;

	constructor (parameters: {
      server: string;
      site: string;
   }) {
		this.server = parameters.server;
		if (this.server.search(/^https?:\/\//) < 0)
			this.server += "https://" + this.server;
		this.site = parameters.site;
	}

	structureMultipleAddresses (addressList: string): string {
		let i: number;

		if (!addressList || addressList == null) return "";
		if (addressList.search(/,/) >= 0) throw "A multiple email address list appears to be comma-separated in a string\n" + "Use a semicolon-separated set of addresses";
		if (addressList.search(/;/) < 0) return addressList;
		const temp = addressList.split(";");
		for (i = 0; i < temp.length; i++)
			if (i == 0) temp[i] = temp[i] + "'";
			else if (i == temp.length - 1) temp[i] = "'" + temp[i];
			else temp[i] = "'" + temp[i] + "'";
		return temp.join(",");
	}

	/** @method .sendEmail
* @param {object} properties should include {From:, To:, CC:, BCC:, Subject:, Body:}
*/
	sendEmail (headers: IEmailHeaders) : Promise<TSPResponseData> {
		return new Promise((resolve, reject) => {
			let To: string,
				CC: string | null,
				BCC: string | null,
				From: string,
				Subject: string,
				Body:string,
				uri: string = this.server;

			if (!headers.To)
				reject({status: 0, error: { message: { value:
               "Object arg to SPUtilityEmailService.sendEmail() method must include .To: property" }}});
			if (!headers.Subject)
				reject({status: 0, error: { message: { value:
               "Object arg to SPUtilityEmailService.sendEmail() method must include defined .Subject: property" }}});
			if (!headers.Body)
				reject({status: 0, error: { message: { value:
            "Object arg to SPUtilityEmailService.sendEmail() method must include defined .Body: property" }}});
			window.onerror = null;
			To = this.structureMultipleAddresses(headers.To);
			CC = headers.CC ? this.structureMultipleAddresses(headers.CC) : "";
			BCC = headers.BCC ? this.structureMultipleAddresses(headers.BCC) : "";
			From = headers.From ? headers.From.replace(/"/g, "\\\"") : "";
			From = From.replace(/'/g, "\\'");
			To = To.replace(/"/g, "\\\"");
			To = To.replace(/'/g, "\\'");
			CC = CC.replace(/"/g, "\\\"");
			CC = CC.replace(/'/g, "\\'");
			BCC = BCC.replace(/"/g, "\\\"");
			BCC = BCC.replace(/'/g, "\\'");
			Subject = headers.Subject.replace(/"/g, "\\\"");
			Subject = Subject.replace(/'/g, "\\'");
			Body = headers.Body.replace(/\\/g, "\\\\");
			Body = Body.replace(/"/g, "\\\"");
			Body = Body.replace(/'/g, "\\'");
			if (this.site && this.site.length > 0)
				uri = uri + (this.site.charAt(0) == "/" ? "" : "/") + this.site;
			RESTrequest({
				url: uri + "/_api/contextinfo",
				method: "POST",
				successCallback: (responseJSON: TSPResponseData) => {
					const digValue = responseJSON!.d!.GetContextWebInformation!.FormDigestValue;
					// the actual transmission
					RESTrequest({
						url: uri + "/_api/SP.Utilities.Utility.SendEmail",
						method: "POST",
						headers: {
							...SPstdHeaders,
							"X-RequestDigest": digValue
						},
						body: JSON.stringify({
							properties : {
								__metadata: { type: "SP.Utilities.EmailProperties" },
								To: { results : [  To  ] },
								CC: { results : [  CC  ] },
								BCC: { results : [ BCC ] },
								From: From,
								Subject: Subject,
								Body: Body
							}
						}),
						successCallback: (data: TSPResponseData) => {
							resolve(data);
						},
						errorCallback: (errInfo, /* httpInfo */) => {
							reject(errInfo);
						}
					});
				},
				errorCallback: (errInfo, /* httpInfo */) => {
					reject(errInfo);
				}
			});
		});
	}
}

// parameters should be object with following acceptable properties
//   .server : this is required to establish the service unless InitSettings.SERVER_NAME defined
//   .site : optional site part of URL unless SITE_PATH is defined
//   .errorObj : should be the response object from a jQuery AJAX call
//   .subject : string for the subject header of email message
//   .body : string allowed to contain HTML markup for email message body
//   .CC or .Cc : string formatted for CC header and containing email addresses
//   .To : To addressee; optional if the global constant
//           DEVELOPER_MAIL_ADDRESS is defined with proper email address


function emailDeveloper(parameters: {
    subject: string;
    body: string;
    server?: string;
    site?: string;
    To?: string;
    CC?: string | null;
    Cc?: string | null;
    head?: string;
    errorObj?: object;
}): Promise<TSPResponseData> {
	return new Promise((resolve, reject) => {
		let bodyStart: string;

		if (!parameters.server) throw "Cannot emailDeveloper() with having a defined server name";
		const emailService = new SPUtilityEmailService({
			server: parameters.server,
			site: parameters.site as string
		});
		bodyStart = "<html><head>";
		if (parameters.head) bodyStart += parameters.head;
		bodyStart += "</head><body>";
		const bodyEnd = "</body></html>";
		emailService.sendEmail({
			From: "",
			To: parameters.To as string,
			Subject: parameters.subject,
			Body: bodyStart + parameters.body + bodyEnd,
			CC: parameters.CC ? parameters.CC : (parameters.Cc ? parameters.Cc : null)
		}).then((response: TSPResponseData) => {
			resolve(response);
		}).catch((error) => {
			reject(error);
		});
	});
}

/*/
function emailUser(parameters: {
    userEmail: string;
    body: string;
    subject: string;
    CC?: string;
}): void {
    var emailService = new SPUtilityEmailService({
            server: InitSettings.SERVER_NAME,
            site: SITE_PATH
        }),
        bodyStart = "<div style=\"font:normal 11pt Verdana,sans-serif;\">",
        bodyEnd = "</div>";
    emailService.sendEmail({
        From: "",
        To: parameters.userEmail,
        Subject: "CAT ERROR:  " + parameters.subject,
        Body: bodyStart + parameters.body + bodyEnd,
        CC: parameters.CC ? parameters.CC : null
    });
}
/**/
/*
function emailHttpErrors(restRequest, config, itemId): void {
    var emailService, body, prop,
                headers = restRequest.getRequestHeaders();
    if (config == "list") config = "listItemCommon()";
    else if (config == "library") config = "libraryItemCommon()";
    body = "Request Type <span class=\"red\"><b>" + restRequest.getRequestTypeReadable() + "</b></span> called in " + config;
    if (itemId && itemId != null) body += " on Item " + itemId + ".";
    else body += " (no item ID).";
    body += "<ul><li>Method: " + restRequest.getRequestMethod() + "</li>" + "<li>URL: " + restRequest.getRequestUrl() + "</li>" + "<li>Headers: <ul>";
    for (prop in headers) body += "<li>'" + prop + "': '" + headers[prop] + "'</li>";
    body += "</ul></li>" + "<li>Body:<br />" + restRequest.getRequestBody() + "</li>" + "<li>Status: " + restRequest.getResponseStatus() + "</li>" + "<li>Response Text: <br />" + restRequest.getResponseText() + "</li></ul>";
    emailService = emailService = new SPUtilityEmailService({
        server: CatCONST.InitSettings.SERVER_NAME,
        site: CatCONST.SITE_PATH
    });
    emailService.sendEmail({
        To: CatCONST.DEVELOPER_EMAIL_ADDRESS,
        Subject: "Correspondence Control System: HTTP Error",
        Body: body
    });
}
*/