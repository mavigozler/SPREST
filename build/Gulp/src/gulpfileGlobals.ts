"use strict";

/* eslint-disable @typescript-eslint/no-unused-vars */

export { glog, tasksCalled, winstlog,
   SelectedGulpfileConfig, GulpDict, GulpfileConfigs, gulpSrcAllowEmpty,
	fwdSlash
};

import { SelectedConfig, GulpfileConfigsInfo } from "./gulpfile.d";

import winston from "winston";
import "winston-daily-rotate-file";
import "winston-color";

const // JsMapDeleteDefault = true,
	doubleStarSlash = "**/",
	tasksCalled: string[] = [],
	gulpSrcAllowEmpty = {allowEmpty: true};

const GulpfileConfigs: GulpfileConfigsInfo = {} as GulpfileConfigsInfo;

const SelectedGulpfileConfig: SelectedConfig = {
	task: undefined,
	stage: undefined,
	config: undefined
};

enum LogLevel {
	ERROR,
	WARN,
	INFO,
	HTTP,
	VERBOSE,
	DEBUG,
	SILLY
}

const TSFiles: {
		includedFileSet: string[];
		excludedFileSet: string[];
		finalFileSet: string[];
		whatsLeft: string[];
	} = {
		includedFileSet: [],
		excludedFileSet: [],
		finalFileSet: [],
		whatsLeft: []
	};

const
	winstlog: {
		logger: winston.Logger | null
	} = {
		logger: null
	};

function glog(
	message: string,
	level?: LogLevel
): void {
//	console.log(message);
	if (message.search(/ERROR/) >= 0)
		level = LogLevel.ERROR;
	else if (message.search(/WARN/) >= 0)
		level = LogLevel.WARN;
	switch (level) {
	case LogLevel.ERROR:
		winstlog.logger!.error(message);
		break;
	case LogLevel.WARN:
		winstlog.logger!.warn(message);
		break;
	case LogLevel.INFO:
	default:
		winstlog.logger!.info(message);
		break;
	case LogLevel.HTTP:
		winstlog.logger!.http(message);
		break;
	case LogLevel.VERBOSE:
		winstlog.logger!.verbose(message);
		break;
	case LogLevel.DEBUG:
		winstlog.logger!.debug(message);
		break;
	case LogLevel.SILLY:
		winstlog.logger!.silly(message);
		break;
	}
}

function fwdSlash(theString: string): string {
	return theString.replace(/\\/g, "/");
}
class GulpDictionarySystem {
	SubstitutionsDictionary: {[key: string]: string} = {};

	get(key: string): string {
		return this.SubstitutionsDictionary[key];
	}

	set(key: string, def: string) {
		this.SubstitutionsDictionary[key] = def;
	}

	resolveString(text: string): string {
		let matches: RegExpMatchArray | null;
		if (typeof text == "string")
			while ((matches = text.match(/\$\{([^}]+)\}/)) != null)
				for (let i = 1; i < matches.length; i++)
					text = text.replace(new RegExp("\\$\\{" + matches[i] + "\\}"), this.SubstitutionsDictionary[matches[i]]);
		return text;
	}

	resolveArrayOfStrings(texts: string[]): string[] {
		const newTexts: string[] = [];
		for (const item of texts)
			newTexts.push(this.resolveString(item));
		return newTexts;
	}
}

const GulpDict = new GulpDictionarySystem();

