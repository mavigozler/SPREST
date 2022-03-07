declare class SPUtilityEmailService {
    server: string;
    site: string;
    constructor(parameters: {
        server: string;
        site: string;
    });
    structureMultipleAddresses(addressList: string): string;
    /** @method .sendEmail
     * @param {object} properties should include {From:, To:, CC:, BCC:, Subject:, Body:}
     */
    sendEmail(headers: IEmailHeaders): Promise<any>;
}
declare function emailDeveloper(parameters: {
    server: string;
    site: string;
    errorObj?: object;
    To: string;
    subject: string;
    body: string;
    CC?: string | null;
    Cc?: string | null;
    head?: string;
}): Promise<any>;
