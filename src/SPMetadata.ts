"use strict";

/* Purpose of code is to get current metadata structures for SharePoint
 * data and object types */

import path from "path";
import type { TSpUser, TSpGroup, TSpSite, TSpList } from "./SPComponentTypes.d.ts";
import type { TSPResponseData } from "./SPHttp.d.ts";
import { RESTrequest } from "./SPHttpReqResp.js";
export { buildType, buildTypes, TypePropertyMetadata,
	//SPUser, SPGroup, SPSite, SPList,
	SPTypesDefs, SPTypeInfo, TItemDetailPropertiesDetails, SP_METADATA_CONFIG_KEY
};


type SPTypeInfo = {
	definedType: TSpSite | TSpList | TSpUser | TSpGroup,
	metadata:  { [key: string]: TypePropertyMetadata }
};

let SPTypesDefs: SPTypeInfo[]
	/*,
	SPUser: TUserInfo,
	SPGroup: TGroupInfo,
	SPSite: TSiteInfo,
	SPList: TListInfo */
;

const SP_METADATA_CONFIG_KEY = "SPMetadataConfigKey",
	entryScript = document.currentScript ? (document.currentScript as HTMLScriptElement).src : null,
	DefaultMetadataJson = "./metadata.json",
	_apiEndpoints: {
   endpoint: string;
   spType: string;
}[] = [
   { endpoint: "user", spType: "SP.User" },
   { endpoint: "group", spType: "SP.Group" },
   { endpoint: "site", spType: "SP.Site" },
   { endpoint: "web", spType: "SP.Web" },
   { endpoint: "list", spType: "SP.list" },
];

interface TypePropertyMetadata {
   type: string;
   display: boolean;
   description?: string;
   label?: string; // label used in a web app
   format?: { // f
      searchRE: string | undefined;  // s
      characterFlag: {  // c
         flagValue: string; // v // what value flag represents
         key: { // k
            charMeaningFlag: string; // m
            charMeaning: string; // r
         }[];
         renderingFormat?: string; // w
      }[];
   };
   sptype?: string;  // SP's own type name
}


type SPtype = {
	_t:string;
};

type MetadataJson = {
	revisionDate: Date;
	user: TItemDetailPropertiesDetails[];
	group: TItemDetailPropertiesDetails[];
	site: TItemDetailPropertiesDetails[];
	list: TItemDetailPropertiesDetails[];
	ref: {[key: string]: TItemDetailPropertiesDetails[] & SPtype};
		// references
};

type TItemDetailPropertiesDetails = {
	p:string; // property name on SP
	t:string; // property type
	_t?:string;
	d?:boolean;  // whether to display in SPAmin tool
	n?:string; // description of the property
	x?:string|number|boolean; // an example value of the property
	l?:string;  // designates label wording for property in UI, or use "p" wording
	f?: { // "flagged"
		s:string; // search RE for flags string
		c: { // character meanings
			v:string;// value meaning for the charater
			k?: {  // character + meaning value set
				m:string; // character to match
				r:string; // meaning/description of character position
			}[];
			w?:string; // specifies a specific writing format
		}[];
	} //"f"ormat for property: "s"earchable RE,
	o?:{ // optioned property
		v:number; // integer value indicating option
		m:string; // meaning of option
	}[];
		//e"x"pression for response to found "s" values
	e?:boolean; // "e"xpanded property, not part of SP if true
};


const buildType = (
   baseType: TSpSite | TSpList | TSpUser | TSpGroup,
   config: TItemDetailPropertiesDetails[]
): SPTypeInfo => {

   const metadata: { [key: string]: TypePropertyMetadata } = {};
   config.forEach(item => {
		baseType[item.p] = item.t;
		metadata[item.p] = {
			type: item.t,
			display: item.d ? item.d : false,
			description: item.n ? item.n : undefined,
			label: item.l ? item.l : undefined,
			format: {
				searchRE: item.f?.s ? item.f.s : undefined,
				characterFlag: item.f?.c ? item.f?.c.map(subitem => {
					return {
						flagValue: subitem.v,
						key: subitem.k ? subitem.k.map(keyitem => {
							return {
								charMeaningFlag: keyitem.m,
								charMeaning: keyitem.r
							};
						}) : [],
						renderingFormat: subitem.w
					};
				}) : []
			}
		}
	});
   return {definedType: baseType, metadata: metadata};
};

function buildTypes(metadataFilePath?: string): Promise<SPTypeInfo[]> {
	if (!metadataFilePath)
		metadataFilePath = DefaultMetadataJson;
	return new Promise<SPTypeInfo[]>((resolve, reject) => {
		readMetadataJson(metadataFilePath!).then((config: MetadataJson) => {
			const builtTypes: SPTypeInfo[] = [];
			for (const typeConfig of ["user", "group", "site", "list"])
				switch (typeConfig) {
				case "user":
					builtTypes.push(buildType({} as TSpUser, config.user));
					break;
				case "group":
					builtTypes.push(buildType({} as TSpGroup, config.group));
					break;
				case "site":
					builtTypes.push(buildType({} as TSpSite, config.site));
					break;
				case "list":
					builtTypes.push(buildType({} as TSpList, config.list));
					break;
				}
			resolve(builtTypes);
		}).catch((err: unknown) => {
			reject(err);
		});
	});
}

function readMetadataJson(metadataJsonFilePath: string): Promise<MetadataJson> {
	return new Promise<MetadataJson>((resolve, reject) => {
		/// try to find the config name in current directory or try to get awareness up one level
		// look for local storage
		const storedConfig: string | null = localStorage.getItem(SP_METADATA_CONFIG_KEY);
		if (checkMetadataIntegrity(storedConfig) == false) {
			console.log(`appConfig.ts::readConfigJson(): reading '${metadataJsonFilePath}'`);
			requestMetadataFile(metadataJsonFilePath).then((response: MetadataJson) => {
				resolve(response);
			}).catch((errInfo) => {
				if ((errInfo as Error).message.search(/Failed to fetch/) >= 0 && entryScript) {
					const fileRetry = path.join(path.dirname(entryScript), path.basename(DefaultMetadataJson));
					requestMetadataFile(fileRetry).then((response: MetadataJson) => {
						resolve(response);
					}).catch((errInfo: Error) => {
						reject(errInfo);
					});
				} else
					reject(errInfo);
			});
		} else
			resolve(JSON.parse(storedConfig as string));
	});
}

function requestMetadataFile(metadataJsonFilePath: string): Promise<MetadataJson> {
	return new Promise<MetadataJson>((resolve, reject) => {
		RESTrequest({
			url: metadataJsonFilePath,
			method: "GET",
			successCallback: (data: TSPResponseData) => {
				const modifiedConfig: MetadataJson = data as MetadataJson;
				modifiedConfig.revisionDate = new Date();
				localStorage.setItem(SP_METADATA_CONFIG_KEY, JSON.stringify(data as string));
				resolve(modifiedConfig);
			},
			errorCallback: (errInfo) => {
				reject(errInfo);
			}
		});
	});
}

function checkMetadataIntegrity(config: string | null): boolean {
	if (config == null) {
		console.log("SPMetadata.ts::SPMetadata(): param 'config' is null");
		return false;
	}
	let metadataObj: MetadataJson;
	try {
		metadataObj = JSON.parse(config);
	} catch {
		return false;
	}
	if (!metadataObj || !metadataObj.user || !metadataObj.site ||
			!metadataObj.list || !metadataObj.group || !metadataObj.ref)
		return false;
	console.log("SPMetadata.ts::SPMetadata(): integrity confirms true");
	return true;
}


// eslint-disable-next-line @typescript-eslint/no-unused-vars
function printSPMetadata(
   server:  string,
   format: "SP" | "configForm",
   expandOnly?: boolean
) {
   for (const path of _apiEndpoints)
      RESTrequest({
         url: `${server}/${path.endpoint}`,
         successCallback: (data) => {
            if (expandOnly && expandOnly == true) {
               let build = "";
               build = objPart(data, build);
               printSPMetadata(
                  `${server}/${path.endpoint}?$expand=${build}`,
                  format
               );
            } else {
               let result: unknown | null = null, results: unknown[] | null = null;

               if (data.d && data.d.results)
                  results = data.d.results;
               else
                  result = data.d;
               if (format == "SP")
                  JSON.stringify(result || results);
               else {
                  let build = "{\n";
                  if (results)
                     for (const item of results)
                        build = objPart(item, build);
                  else
                     build = objPart(result, build);
                  build += "\n}"
               }
            }
         },
         errorCallback: (errInfo) => {
            console.log(errInfo)
         }
      })

   function objPart(partIn: unknown, build: string, expandOnly?: boolean): string {
      if (expandOnly && expandOnly == true) {
         // eslint-disable-next-line @typescript-eslint/no-explicit-any
         for (const prop in partIn as any[])
            // eslint-disable-next-line no-prototype-builtins, @typescript-eslint/no-explicit-any
            if ((partIn as any).hasOwnProperty(prop) == true)
               build += "," + prop;
         build = build.substring(1);
      } else {
         const datetimeRE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/;
         for (const prop in partIn as unknown[]) {
            if (prop == "__metadata")
               // eslint-disable-next-line @typescript-eslint/no-explicit-any
               build += `\n\t{"_t":"${(partIn as any)[prop].type}"}`
            else {
               build += `\n\t{"p":"${prop}","t":`;
               // eslint-disable-next-line @typescript-eslint/no-explicit-any
               const val = (partIn as any)[prop];
               if (typeof val == "string")
                  if (val.search(datetimeRE) == 0)
                     build += '"datetime"';
                  else
                     build += '"string"';
               else if (Array.isArray(val) == true) {
                  if (val.length > 0)
                     build += `"${typeof val[0]}[]"`;
                  else
                     build += '"unknown[]"';
               } else
                  build += `"${typeof val}"`;
               build += "},";
            }
         }
      }
      return build;
   }
}


/*
printSPMetadata(
   "https://mhalloran.sharepoint.com",
   "configForm"
);
*/