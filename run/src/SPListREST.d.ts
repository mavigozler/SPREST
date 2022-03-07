/**
 * @class SPListREST  -- constructor for interface to making REST request for SP Lists
 */
declare class SPListREST {
    protocol: string;
    server: string;
    site: string;
    apiPrefix: string;
    listName: string;
    listGuid: string;
    baseUrl: string;
    serverRelativeUrl: string;
    creationDate: Date;
    baseTemplate: string;
    allowContentTypes: boolean;
    itemCount: number;
    listItemEntityTypeFullName: string;
    rootFolder: string;
    contentTypes: object;
    listIds: number[];
    fields: any[];
    fieldInfo: any[];
    lookupFieldInfo: TLookupFieldInfo[];
    lookupInfoCached: boolean;
    linkToDocumentContentTypeId: string;
    currentListIdIndex: number;
    setup: TListSetup;
    static stdHeaders: THttpRequestHeaders;
    /**
     * @constructor
     * @param {object:
     *	server: {string} server domain, equivalent of location.hostname return value
    * 	site: {string}  URL string part to the site
    *		list: {string}  must be a valid list name. To create a list, use SPSiteREST() object
    *    include: {string}  comma-separated properties to be part of OData $expand for more properties
    *	} setup -- object to initialize the
    */
    constructor(setup: TListSetup);
    static escapeApostrophe(string: string): string;
    /**
     * @method _init -- sets many list characteristics whose data is required
     * 	by asynchronous request
     * @returns Promise
     */
    init(): Promise<unknown>;
    getLookupFieldsInfo(): Promise<unknown>;
    /**
     * @method retrieveLookupFieldInfo -- Promise to return choices for a lookup value
     * @param {object
     * 		.useFunction: {string}  "search" (default)|"enumWebs"
     *      .LookupWebId:  must be GUID form of the site containing the list that is lookup
     *      .LookupList:   must be GUID form the list containing column lookup
     *      .LookupField:  must be the Internal Name form of the field in the list looked up
     *    } parameters -- object properties specified
     * @returns {Promise} the array of values that form the choice list of the lookup
     */
    retrieveLookupFieldInfo(parameters: {
        useFunction?: string;
        LookupWebId: string;
        LookupList: string;
        LookupField: string;
    }): Promise<any>;
    /**
     * @method getLookupFieldValue -- retrieves as a Promise the value to set up for a lookup field
     * @param {string} fieldName --
     * @returns the value to set for the lookup field, null if not available
     */
    getLookupFieldValue(fieldName: string, fieldValue: string): string | null;
    /**
     * @method fixBodyForSpecialFields -- support function to update data for a REST request body
     *    where the column data is special, such as multichoice or lookup or managed metadata
     * @param {object} body -- usually the XMLHTTP request body passed to a POST-type request
     * @returns -- a Promise (this is an async call)
     */
    fixBodyForSpecialFields(body: THttpRequestBody): Promise<unknown>;
    static httpRequest(elements: THttpRequestParams): void;
    static RequestAgain(nextUrl: string, aggregateData: any[]): Promise<unknown>;
    static httpRequestPromise(parameters: THttpRequestParamsWithPromise): Promise<unknown>;
    getProperties(parameters: any): Promise<any>;
    getListProperties(parameters?: any): Promise<any>;
    getListInfo(parameters?: any): Promise<any>;
    /**
     * @method getListItemData
     * @param {object --
     * 		itemId: specific ID of item data to be retrieved
     * 		lowId: starts a range with the low ID
     * 		highId: starts a range with the high ID
     * 	} parameters
     * @returns {Promise} HTTP Request
     */
    getListItemData(parameters: {
        itemId: number;
        lowId?: number;
        highId?: number;
        [key: string]: any;
    }): Promise<any>;
    getListItem(parameters: any): Promise<any>;
    getItemData(parameters: any): Promise<any>;
    getAllListItems(): Promise<any>;
    getAllListItemsOptionalQuery(parameters: any): Promise<any>;
    getListItems(parameters?: any): Promise<any>;
    createListItem(item: {
        body: THttpRequestBody;
    }): Promise<unknown>;
    /**
     *
     * @param {arrayOfObject} items -- objects should be objects whose properties are fields
     * 		within the lists and using Internal name property of field
     * @returns a promise
     */
    createListItems(items: any[], useBatch: boolean): Promise<any>;
    /**
     * @method updateListItem -- modifies existing list item
     * @param {object} parameters -- requires 2 parameters {itemId:, body: [JSON] }
     * @returns
     */
    updateListItem(parameters: {
        itemId: number;
        body: THttpRequestBody;
    }): Promise<unknown>;
    updateListItems(itemsArray: any[], useBatch: boolean): Promise<any>;
    /**
    applies to all lists
    parameters properties should be
    .itemId [required]
    .recycle [optional, default=true]. Use false to get hard delete
    */
    deleteListItem(parameters: {
        itemId: number;
        recycle: boolean;
    }): Promise<any>;
    getListItemCount(doFresh?: boolean): Promise<any>;
    loadListItemIds(): Promise<unknown>;
    isValidID(itemId: number): Promise<any>;
    getListItemIds(): Promise<any>;
    getNextListId(): Promise<any>;
    getPreviousListId(): Promise<unknown>;
    getFirstListId(): Promise<any>;
    setCurrentListId(itemId: number | string): boolean;
    getAllDocLibFiles(): Promise<any>;
    getFolderFilesOptionalQuery(parameters: {
        folderPath: string;
    }): Promise<any>;
    /** @method getDocLibItemByFileName
    retrieves item by file name and also returns item metadata
    * @param {Object} parameters
    * @param {string} parameters.fileName - name of file
    */
    getDocLibItemByFileName(parameters: {
        fileName?: string;
        itemName?: string;
    }): Promise<any>;
    /** @method getDocLibItemMetadata
     * @param {Object} parameters
     * @param {number|string} parameters.itemId - ID of item whose data is wanted
     * @returns  {Object} returns only the metadata about the file item
     */
    getDocLibItemMetadata(parameters: {
        itemId: number;
    }): Promise<any>;
    /** @method getDocLibItemFileAndMetaData
     * @param {Object} parameters
     * @param {number|string} parameters.itemId - ID of item whose data is wanted
     * @returns {Object} data about the file and metadata of the library item
     */
    getDocLibItemFileAndMetaData(parameters: {
        itemId: number;
    }): Promise<any>;
    /**
    arguments as parameters properties:
    .fileName or .itemName -- required which will be name applied to file data
    .folderPath -- optional, if omitted, uploaded to root folder
    .body  required file data (not metadata)
    .willOverwrite [optional, default = false]
    */
    uploadItemToDocLib(parameters: {
        fileName: string;
        body: TXmlHttpRequestData;
        folderPath: string;
        willOverwrite: boolean;
    }): Promise<any>;
    createFolder(parameters: {
        folderName: string;
    }): Promise<any>;
    /**
     * @method checkOutDocLibItem  checks out a document library item where checkout is required
     * @param {string} .fileName or .itemName -- required which will be name applied to file data
     * @param {string} .folderPath -- optional, if omitted, uploaded to root folder
     * @param {byte} exceptions - flag on how to handle exceptions
     *     only acceptable options are IGNORE_THIS_USER_CHECKOUT, IGNORE_NO_EXIST
     * @param {object} exceptionsData. For IGNORE_THIS_USER_CHECKOUT, a string that
     *     is part of the user name should be passed as {uname: <username string>}
     */
    checkOutDocLibItem(parameters: {
        fileName?: string;
        itemName?: string;
        folderPath?: string;
    }): Promise<any>;
    /**
     * @method checkInDocLibItem -- check in one or more files
     * @param {object:
     * 		{string[]?} itemNames -- must be array representing folder paths + file names
     *       {string} checkInType -- same check in type for one or multiple files
     * 				must be anyon of 6 strings: "[mM]inor" (0), "[mM]ajor" (1), "[oO]verwrite" (2)
     *       {string} .checkInComment -- applied to all files if multipe
                * } parameters -- the following properties are recognized
     * @returns Promise
     */
    checkInDocLibItem(parameters: {
        itemNames: string[];
        checkInType?: TSPDocLibCheckInType;
        checkInComment?: string;
    }): Promise<any>;
    /**
     * @method discardCheckOutDocLibItem
     * @param {Object} parameters
     * @param {string} parameters.(fileName|itemName) - name of file checked out
     * @param {string} parameters.folderPath - path to file name in doc lib
     */
    discardCheckOutDocLibItem(parameters: {
        file: string;
        folderPath: string;
    }): Promise<any>;
    renameItem(parameters: any): Promise<any>;
    renameFile(parameters: {
        itemId: number;
        itemName: string;
        newName: string;
    }): Promise<any>;
    renameItemWithCheckout(parameters: {
        itemId: number;
        newName: string;
    }): Promise<any>;
    /** @method updateLibItemWithCheckout
     * @param {Object} parameters
     * @param {string|number} parameters.(fileName|itemId) - id of lib item or name of file checked out
     */
    updateLibItemWithCheckout(parameters: {
        itemId: number;
        fileName: string;
        body: string;
        checkInType?: TSPDocLibCheckInType;
    }): Promise<unknown>;
    /** @method continueupdateLibItemWithCheckout
     * @param {Object} parameters
     * @param {string|number} parameters.(fileName|itemId) - id of lib item or name of file checked out
     */
    continueupdateLibItemWithCheckout(parameters: {
        itemId: number;
        fileName: string;
        checkInType?: TSPDocLibCheckInType;
        body: string;
    }): Promise<unknown>;
    /** @method discardCheckOutDocLibItem
     * @param {Object} parameters
     * @param {string} parameters.(fileName|itemName) - name of file checked out
     * @param {string} parameters.folderPath - path to file name in doc lib
     */
    updateDocLibItemMetadata(parameters: {
        itemId: number;
        body: THttpRequestBody;
    }): Promise<unknown>;
    /**
    required arguments as parameters properties:
    .sourceFileName [required]  the path to the source name
    .destinationFileName [required]  the path to the location to be copied,
    can have file name different than source name
    .willOverwrite [optional, default = false]
    */
    copyDocLibItem(parameters: {
        sourceFileName: string;
        destinationFileName: string;
        willOverwrite?: string;
    }): Promise<any>;
    deleteDocLibItem(parameters: {
        folderPath?: string;
        fileName?: string;
        path?: string;
        includeBaseUrl?: boolean;
        recycle?: boolean;
    }): Promise<any>;
    /*******************************************************************************
     *     LIST PROPERTY / STRUCTURE REQUESTS: fields, content types, views
     *******************************************************************************/
    /****************  Fields  ****************************/
    /**
     * @method getField -- returns one or more field objects, with properties all or selected, and possible
     *     filtering of results
     * @param {object} parameters -- object takes two properties or can be undefined
     *			if parameters is null or undefined, it will return all fields and their properties for the list
     * 		 fields: an array of objects or null or empty array, where object {name: <field-name>} or {guid: <field-guid>}
     * 	    query: an object with the following properties: "select", "filter", "expand" which operate as O-Data
     *               query is optional
     * @returns -- will return to then() an array of results representing fields and their properties
     */
    getField(parameters: {
        fields: {
            name?: string;
            guid?: string;
        }[];
        filter?: string;
    } | undefined): Promise<unknown>;
    getFields(parameters?: any): Promise<any>;
    /**
     *
     * @param {JSONofParms} fieldProperties 'Title': 'field title', 'FieldTypeKind': FieldType value,
     *    Required': 'true/false', 'EnforceUniqueValues': 'true/false','StaticName': 'field name'}
     * @returns REST API response
     */
    createField(fieldProperties: {
        "Title": string;
        "FieldTypeKind": number;
        "Required": boolean;
        "EnforceUniqueValues": boolean;
        "StaticName": string;
    }): Promise<any>;
    /**
     *
     * @param {arrayOfFieldObject} fields -- this should be an array of the
     * 			       objects used to define fields for SPListREST.createField()
     * @returns
     */
    createFields(fields: any[]): Promise<unknown>;
    renameField(parameters: {
        oldName?: string;
        currentName?: string;
        newName: string | undefined;
    }): Promise<unknown>;
    /** @method .updateChoiceTypeField
     *       @param {Object} parameters
     *       @param {string} parameters.id - the field guid/id to be updated
     *       @param {(array|arrayAsString)} parameters.choices - the elements that will
     *     form the choices for the field, as an array or array written as string
     */
    updateChoiceTypeField(parameters: {
        id: string;
        choices?: string[];
        choice?: string;
    }): Promise<unknown>;
    /****************  Content Types  ****************************/
    /**
     * @method getContentType -- returns one or more contentType objects, with properties all or selected, and possible
     *     filtering of results
     * @param {object} parameters -- object takes two properties or can be undefined
     *			if parameters is null or undefined, it will return all fields and their properties for the list
     * 		 contentTypes: an array of objects or null or empty array, where object {name: <content type-name>} or {guid: <field-guid>}
     * 	    query: an object with the following properties: "select", "filter", "expand" which operate as O-Data
     *               query is optional
     * @returns -- will return to then() an array of results representing content types and their properties
     */
    getContentType(parameters: {
        contentTypes?: {
            name?: string;
            guid?: string;
        }[];
        query?: {
            filter: string;
        };
    } | undefined): Promise<any>;
    /**
     *
     * @param {arrayOfcontentTypeObject} parameters -- this should be an array of the
     * 			       object parameter passed to .createcontentType
     * @returns
     */
    /****************  Views  ****************************/
    /**
     * @method getView -- returns one or more contentType objects, with properties all or selected, and possible
     *     filtering of results
     * @param {object} parameters -- object takes two properties or can be undefined
     *			if parameters is null or undefined, it will return all fields and their properties for the list
     * 		 contentTypes: an array of objects or null or empty array, where object {name: <content type-name>} or {guid: <field-guid>}
     * 	    query: an object with the following properties: "select", "filter", "expand" which operate as O-Data
     *               query is optional
     * @returns -- will return to then() an array of results representing content types and their properties
     */
    getView(parameters?: {
        views?: {
            name?: string;
            guid?: string;
        }[];
        query?: {
            filter: string;
        };
    }): Promise<any>;
    /**
     *
     * @param {JSONofParms} viewProperties }
     * @returns REST API response
     */
    createView(viewsProperties: THttpRequestBody): Promise<any>;
    setView(viewName: string, fieldNames: string[]): Promise<unknown>;
    /**
     *
     * @param {arrayOfFieldObject} views -- this should be an array of the
     * 			       objects used to define views for SPListREST.createViews()
     * @returns
     */
    createViews(views: any[]): Promise<unknown>;
}
