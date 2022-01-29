"use strict";

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

class SPUtilityEmailService {
    server: string;
    site: string;

    constructor (parameters: {
        server: string;
        site: string;
    }) {
        this.server = parameters.server;
        this.site = parameters.site;
    }

    structureMultipleAddresses (addressList: string): string {
        let i: number, temp: any;
        if (!addressList || addressList == null) return "";
        if (addressList.search(/,/) >= 0) throw "A multiple email address list appears to be comma-separated in a string\n" + "Use a semicolon-separated set of addresses";
        if (addressList.search(/;/) < 0) return addressList;
        temp = addressList.split(";");
        for (i = 0; i < temp.length; i++)
            if (i == 0) temp[i] = temp[i] + "'";
            else if (i == temp.length - 1) temp[i] = "'" + temp[i];
        else temp[i] = "'" + temp[i] + "'";
        return temp.join(",");
    };

    /** @method .sendEmail
     * @param {object} properties should include {From:, To:, CC:, BCC:, Subject:, Body:}
     */
    sendEmail (headers: IEmailHeaders) : Promise<any> {
        return new Promise((resolve, reject) => {
            let To: string, 
            CC: string | null, 
            BCC: string | null, 
            From: string, 
            Subject: string, 
            Body:string, 
            uri: string = this.server;

            if (!headers.To) 
                throw "Object arg to SPUtilityEmailService.sendEmail() method must include .To: property";
            if (!headers.Subject) 
                throw "Object arg to SPUtilityEmailService.sendEmail() method must include defined .Subject: property";
            if (!headers.Body) 
                throw "Object arg to SPUtilityEmailService.sendEmail() method must include defined .Body: property";
            window.onerror = null;
            To = this.structureMultipleAddresses(headers.To);
            CC = headers.CC ? this.structureMultipleAddresses(headers.CC) : "";
            BCC = headers.BCC ? this.structureMultipleAddresses(headers.BCC) : "";
            From = headers.From ? headers.From.replace(/"/g, '\\"') : "";
            From = From.replace(/'/g, "\\'");
            To = To.replace(/"/g, '\\"');
            To = To.replace(/'/g, "\\'");
            CC = CC.replace(/"/g, '\\"');
            CC = CC.replace(/'/g, "\\'");
            BCC = BCC.replace(/"/g, '\\"');
            BCC = BCC.replace(/'/g, "\\'");
            Subject = headers.Subject.replace(/"/g, '\\"');
            Subject = Subject.replace(/'/g, "\\'");
            Body = headers.Body.replace(/\\/g, "\\\\");
            Body = Body.replace(/"/g, '\\"');
            Body = Body.replace(/'/g, "\\'");
            if (this.site && this.site.length > 0)
                uri += "/" + this.site;
            $.ajax({
                url: "https://" + uri + "/_api/contextinfo",
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose"
                },
                contentType: "application/json;odata=verbose",
                success: (responseJSON: any) => {
                    var digValue,
                        // the actual transmission
                        digValue = responseJSON.d.GetContextWebInformation.FormDigestValue;
                    $.ajax({
                        url: "https://" + uri + "/_api/SP.Utilities.Utility.SendEmail",
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "X-RequestDigest": digValue
                        },
                        data: JSON.stringify({ properties : {
                                __metadata: { type: "SP.Utilities.EmailProperties" },
                                To: { results : [  To  ] }, 
                                CC: { results : [  CC  ] },
                                BCC: { results : [ BCC ] },
                                From: From,
                                Subject: Subject,
                                Body: Body
                            } 
                        }),
                        contentType: "application/json;odata=verbose",
                        success: (data: any, message: string, reqObj: JQueryXHR) => {
                            resolve(reqObj.status);
                            console.log("Email success");
                        },
                        error: (reqObj: any) => {
                            reject(reqObj);
                            console.log("Email failure: " + "\n  HTTP Status: " + reqObj.status + "\n  Message: " + reqObj.responseJSON.error.message.value);
                        }
                    });
                },
                error: function(responseObj: any) {
                    // this = request object
                    alert("ERROR!\n\n" + "Contact the developer with this error message below\n\n" + "URL: " + this.url + "\n" + "Message: " + 
                            responseObj.name + ": " + responseObj.message);
                }
            });
        });
    }
}

// parameters should be object with following acceptable properties
//   .server : this is required to establish the service unless SERVER_NAME defined
//   .site : optional site part of URL unless SITE_NAME is defined
//   .errorObj : should be the response object from a jQuery AJAX call
//   .subject : string for the subject header of email message
//   .body : string allowed to contain HTML markup for email message body
//   .CC or .Cc : string formatted for CC header and containing email addresses
//   .To : To addressee; optional if the global constant
//           DEVELOPER_MAIL_ADDRESS is defined with proper email address

function emailDeveloper(parameters : {
    server: string;
    site: string;
    errorObj?: object;
    To: string;
    subject: string;
    body: string;
    CC?: string | null;
    Cc?: string | null;
    head?: string;
}): Promise<any> {
    return new Promise((resolve, reject) => {
        let emailService: SPUtilityEmailService, 
        bodyStart: string, 
        bodyEnd: string;

        if (!parameters.server) throw "Cannot emailDeveloper() with having a defined server name";
        emailService = new SPUtilityEmailService({
            server: parameters.server,
            site: parameters.site as string
        });
        bodyStart = "<html><head>";
        if (parameters.head) bodyStart += parameters.head;
        bodyStart += "</head><body>";
        bodyEnd = "</body></html>";
        emailService.sendEmail({
            From: "",
            To: parameters.To as string,
            Subject: parameters.subject,
            Body: bodyStart + parameters.body + bodyEnd,
            CC: parameters.CC ? parameters.CC : (parameters.Cc ? parameters.Cc : null)
        }).then((response) => {
            resolve(response);
        }).catch((response) => {
            reject(response);
        });;
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
            server: SERVER_NAME,
            site: SITE_NAME
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
        server: CatCONST.SERVER_NAME,
        site: CatCONST.SITE_NAME
    });
    emailService.sendEmail({
        To: CatCONST.DEVELOPER_EMAIL_ADDRESS,
        Subject: "Correspondence Control System: HTTP Error",
        Body: body
    });
}
*/