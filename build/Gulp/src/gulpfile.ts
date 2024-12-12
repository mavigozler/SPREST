/* eslint-disable @typescript-eslint/no-unused-vars */
"use strict";

/**********
 *  1. ALWAYS VERIFY GULP VERSION 4 IS INSTALLED!
 *  2. If using gulp-typescript, 'npm install gulp-typescript@latest --save-dev
 * *************/

// 1. purpose of this gulpfile and 2. required set up at bottom of file
//   3. consider watch() function

/****** Exports
 * default:  use args to 'gulp' to shape task use
 ******/
/*
type dirEntryInfo = {
	entryPath: string;
	entryType: "FOLDER" | "FILE" | "UNKNOWN";
	isDeleted: boolean;
};
*/


import { ConfigJson, GulpfileConfigStage,
	GulpCallback, HTMLdef, tsconfigPathMaps, EditInfo,
	SelectedConfig, deleteAsyncOptions, ConfigStage } from "./gulpfile.d";

import {
	tasksCalled, glog, winstlog,
	SelectedGulpfileConfig, GulpDict, GulpfileConfigs, gulpSrcAllowEmpty,
	fwdSlash
} from "./gulpfileGlobals";

import gulpfilelist from "./gulp-filelistexmem";
//import gulpgetfnames from "./gulp-getfnames";

import * as fs from "fs";
import parseArgs from "minimist";
import * as pathMod from "path";
import { deleteAsync } from "del";
import winston from "winston";
import "winston-daily-rotate-file";
import "winston-color";

// gulp and gulp plugins
import gulpCJS from "gulp";
const {
	task,
	series,
	parallel,
	src,
	dest,
	emit,
	watch
} = gulpCJS;
import gulpts from "gulp-typescript";
import gulpsrcmaps from "gulp-sourcemaps";
// import gulpFlatten from "gulp-flatten";
// import rename from "gulp-rename";
// import htmlReplace from "gulp-html-replace";

// import filelog from 'gulp-filelog';
// import debug from 'gulp-debug';
// import filter from 'gulp-filter';
// const filter = require('gulp-filter');

import JSON5 from "json5";
// import webpack from "webpack";
//import webpackStream from "webpack-stream";
//import { webpackConfig } from "./webpack.config";
//import browserSync from "browser-sync";
import { exec, spawn } from "child_process";
import * as globpkg from "glob";
const { glob, globSync } = globpkg;
// import ts from "typescript";
import inquirer from "inquirer";

import { createRequire } from "module";
const require = createRequire(import.meta.url);
const importExportRE =
/^\s*import\s+{[^}]+}\s+from\s+"[^"]+";?|^\s*import\s+.*\s+from\s+"[^"]+";?$|^\s*import\s+"[^"]+";?$|^\s*export\s+{[^};]+};?$/mg;
const importExportReplace = "";//"\n\n";

// import { promisify } from 'util';
/*
const readFileAsync = promisify(fs.readFile); */

// import { safe as jsonc } from 'jsonc';
function loggerSetup() {
	// log rotation
	const { format, transports, createLogger } = winston;
	const { combine, timestamp, label, printf, colorize } = format;
	const loggingFormat = winston.format.cli();
	/* combine(
		timestamp({format: "dd-MMM-YYYY HH:mm:ss"}),
		label({ label: 'gulpfile-logger'}),
		printf(({level, message, timestamp}) => {
			return `${colorize({
				all: true,
				colors: { info: 'blue', warn: 'yellow', error: 'red' }
			})} [${timestamp}]\n${message}`;
		})
	) */
	const fileRotateTransport = new transports.DailyRotateFile({
		filename: 'combined-%DATE%.log',
		// dirname: default="."  directory to save to
		// stream: write to custom stream, bypass rotation
		datePattern: 'YYYY-MM-DD', // controls frequency of rotation 'moment.js' date format
		zippedArchive: true, // gzipping
		// frequency: timed rotations, #m '5m' or #h '3h' if null->datePattern
		maxSize: '20m',
		maxFiles: '14d'  // when to delete aged log files
		// options: see NodeJS/API/fs.html#fs_fs_createwritestream_path_options
		// auditFile: string for audit file name
		// utc: use the UTC time for date in filename
		// extension to use for filename (default="")
	});
	/*   can set event handlers related with log rotation
	fired when a log file is created
		fileRotateTransport.on('new', (filename) => {});
	fired when a log file is rotated
		fileRotateTransport.on('rotate', (oldFilename, newFilename) => {});
	fired when a log file is archived
		fileRotateTransport.on('archive', (zipFilename) => {});
	fired when a log file is deleted
		fileRotateTransport.on('logRemoved', (removedFilename) => {});
	*/
	winstlog.logger = createLogger({
		// or try winston.config.syslog.levels:
		//   https://datatracker.ietf.org/doc/html/rfc5424#page-10
		levels: winston.config.npm.levels,
		// winston reference: https://betterstack.com/community/guides/logging/how-to-install-setup-and-use-winston-and-morgan-to-log-node-js-applications/#configuring-transports-in-winston
		level: process.env.LOG_LEVEL || "info",
		format: loggingFormat,   //https://www.npmjs.com/package//winston#formats
		// defaultMeta: { service: 'user-service' },
		transports: [
			//
			// - Write all logs with importance level of `error` or less to `error.log`
			// - Write all logs with importance level of `info` or less to `combined.log`
			//
			fileRotateTransport,
			new transports.Console(),
			new transports.File({ filename: "error.log", level: "error" }),
			new transports.File({ filename: "combined.log" }),
		],
	});
}
/*
function help(cb: GulpCallback) {
`   flags/options:
       --configFilePath=</path/to/tsconfig.json>
           specifies relative (from gulpfile.js) or absolute path to a
          \`tsconfig.json\` file to be run by 'tsc' task function if used
       -c
           cleanup any built files before new files are built as specified
           in a \`tsconfig.json\` file used. Only works if 'tsc' task function used
       --preclean=<JSON.parse()able string of array of path patterns for
           cleaning (deleting) files before any build process occurs.
           This flag is ignored if '-c' flag is specified and there a \`tsconfig.json\`
           file is found and utilized
`
}
*/

function readGulpfileConfigJson(argList: parseArgs.ParsedArgs) {
	/*if (!argList.s) {
		console.error("Gulpfile execution requires specification of two items in\n" +
				"a special argument format: 'gulp -t <task-name> -s <stage-name>'\n" +
				" (1) a task name like 'default' or 'buildSpecificProject'\n" +
				"   if a task name not specified, 'default' is assumed\n" +
				" (2) an development stage indicator: 'test' or 'release' are allowed\n" +
				"   a stage is required to be specified\n\n" +
				"Both task and stage names can be specified as one argument with colon-separated 'task:stage' or\n" +
				" two arguments which is a space-separated 'task stage' format.\n\n" +
				"Exiting with code '1'");
		process.exit(1);
	} else { */

	//}

	/* argv[0]  "C:\\Program Files\\nodejs\\node.exe"
			[1]  "D:\\dev\\GenLib\\node_modules\\gulp\\bin\\gulp.js"
			[2]  <task-name> e.g. 'default'
			[3]  config: 'buildJSONex'  <--- 'config=buildJSONex'
	// Load plugins*/
	// ...

	// the -c argument for running this gulpfile points to the path of a configuration JSON file for specifying
	//   useful parameters to direct the processing by Gulp. If missing, the path will be assumed to be
	//   the file 'gulpfile.config.json' in the directory where gulpfile.js is located
	let fileJSON: ConfigJson;
	const GULPFILE_CONFIG_JSON_PATH = argList.c ?
		fwdSlash(argList.c) :
		"../config/gulpfile.config.json";
	try {
		// note this is commented JSON, so use a json-c parser
		fileJSON = JSON5.parse(
			fs.readFileSync(GULPFILE_CONFIG_JSON_PATH).toString("utf-8")
		);
	} catch (err: unknown) {
		glog("ERROR reading 'gulpfile.config.json'\n" +
				(err as Error).message);
		process.exit(1);
	}
	/* get the Gulp Config JSON content and put it in memory */
	GulpfileConfigs.configNames =  fileJSON.configNames;
	GulpfileConfigs.outputReportPath = fileJSON.outputReportPath ? fileJSON.outputReportPath : undefined;
	GulpfileConfigs.fileJson = fileJSON;
	GulpfileConfigs.defaultConfig = "default";
	GulpfileConfigs.gulpfileJSdirectory = pathMod.dirname(process.argv[1]);
	GulpfileConfigs.configs = [];

	let storedDetail: Record<string, GulpfileConfigStage>,
		gulpConfigIdx: number;

	// check for project root and abort if not present
	if (!fileJSON.base || !fileJSON.base.projectRoot || fileJSON.base.projectRoot.length == 0) {
		glog("ERROR could not find {base.projectRoot} in the stored JSON. This must be specified");
		process.exit(1);
	}
	for (const configName of fileJSON.configNames) {
		storedDetail = fileJSON[configName];
		const configDetail = {
			configName: configName,
			stages: []
		};
		gulpConfigIdx = GulpfileConfigs.configs.push(configDetail) - 1;
		for (const prop in storedDetail)
			GulpfileConfigs.configs[gulpConfigIdx].stages.push({
					stageName: prop,
					info: storedDetail[prop]
				});
	}
	// look for a 'default' config definition. If it is does not exist, set FIRST config as default
	if (GulpfileConfigs.configNames.findIndex(elem => elem == "default") < 0)
		GulpfileConfigs.defaultConfig = GulpfileConfigs.configNames[0];
}

function systemSetup(done: GulpCallback) {
	tasksCalled.push("systemSetup");
	loggerSetup();
	glog("\ntask: ----   system setup");

	const argList: parseArgs.ParsedArgs = parseArgs(process.argv.slice(2));
	readGulpfileConfigJson(argList);
	for (const key in GulpfileConfigs.fileJson.base)
		GulpDict.set(key, (GulpfileConfigs.fileJson.base)[key]);

	if (!argList.t)
		argList.t = GulpfileConfigs.defaultConfig;

	// get user input about configuration here
	let version: string;
	exec("npm list gulp --json", { cwd: GulpDict.get("projectRoot") }, (err, stdout, stderr) => {
		if (err)
			glog("ERROR: systemSetup()::exec(<for gulp version>)" +
				`\n  errmsg: ${err.message}`);
		else {
			version = JSON.parse(stdout).dependencies.gulp.version;

			let setTask: string | undefined, setStage: string | undefined;
			const defaultAnswers = {
					task: GulpfileConfigs.defaultConfig,  // Replace with the default task name
					stage: "test"   // Replace with the default stage name
				},
				systemDetail = (selectedConfig: SelectedConfig) => {
					return "\n  ====== System Detail ======" +
					`\n  Current working directory: ${GulpfileConfigs.gulpfileJSdirectory}` +
					`\n  Gulp version: ${version}` +
					`\n  Gulp args:  '${process.argv.join("', '")}'` +
					`\n  Selected Gulp config:\n     task = '${selectedConfig.task}'\n    stage = '${selectedConfig.stage}'\n` +
					"\n  ===========================";
				};
			if (process.env) {
				if (process.env.NODE_ENV)
					glog(`\`process.env.NODE_ENV\` is defined: ${process.env.NODE_ENV}`);
				else
					glog("`process.env.NODE_ENV` is not defined");
			} else
				glog("Object `process.env` is not defined");

			if (argList.t && argList.s) {
				setTask = argList.t;
				setStage = argList.s;
			} else if (process.env?.NODE_ENV === "development") {
				setTask = defaultAnswers.task;
				setStage = defaultAnswers.stage;
			}
			if (setTask && setStage) {
				SelectedGulpfileConfig.task = setTask;
				SelectedGulpfileConfig.stage = setStage;
				SelectedGulpfileConfig.config = {} as GulpfileConfigStage;
				SelectedGulpfileConfig.config = GulpfileConfigs.fileJson[SelectedGulpfileConfig.task!][SelectedGulpfileConfig.stage!];

				/*
				GulpfileConfig.base = {
					projectRoot: storedJson.base.projectRoot,
					browserLoadPage: storedJson.base.browserLoadPage
				};
				GulpfileConfig = resolveBaseSubstitutions(GulpfileConfig, GulpfileConfig.base) as GulpfileConfigStage; */
				glog(systemDetail(SelectedGulpfileConfig));
				done();
			} else {
				let selectedTask: { configName: string; stages: ConfigStage[] } | undefined;
				inquirer.prompt([
					{
						type: "list",
						name: "task",
						message: "Select the task:",
						choices: GulpfileConfigs.configNames,
						default: defaultAnswers.task,
					},
					{
						type: "list",
						name: "stage",
						message: "Select the stage of task ",
						choices: (answers) => {
							selectedTask = GulpfileConfigs.configs.find(
									elem => elem.configName == answers.task
							);
							return selectedTask!.stages.map(elem => elem.stageName);
						},
						default: defaultAnswers.stage,
					}
				]).then((answers) => {
					SelectedGulpfileConfig.task = answers.task;
					SelectedGulpfileConfig.stage = answers.stage;
					SelectedGulpfileConfig.config = selectedTask!.stages.find(elem => elem.stageName)!.info;
				/*
					GulpfileConfig.base = {
						projectRoot: storedJson.base.projectRoot,
						browserLoadPage: storedJson.base.browserLoadPage
					};
					resolveBaseSubstitutions(GulpfileConfig, GulpfileConfig.base); */
					glog(systemDetail(SelectedGulpfileConfig));
					done();
				}).catch((err) => {
					glog("ERROR gulpfile.js::inquirer.prompt() did not execute properly--");
					glog(`....${err.message}`);
					glog(systemDetail(SelectedGulpfileConfig));
					done();
				});
			}
		}
	});
}


/*
function gleanForList(gulpedFiles: string[]): Promise<boolean> {
	return new Promise((resolve, reject) => {
		src(gulpedFiles)
			.pipe(gulpFilelist(GulpfileConfig.preclean));
		resolve(true);
//		.pipe(dest("./"));
	});
}
*/
/**
 * @function preclean -- private task better used to delete one or more files
 *     set up as an array of string patterns to paths in same way as glob
 * @param done {pre-clean callback}
 */
function preclean(): Promise<void> {
	const gulpfileConfig = SelectedGulpfileConfig.config!;
	if (!gulpfileConfig.preclean)
		return Promise.resolve(glog("No preclean specified in config JSON"));
	else if (gulpfileConfig.preclean == null)
		return Promise.resolve(glog("preclean set to 'null': no preclean action"));
	else if (!gulpfileConfig.preclean.patterns)
		return Promise.reject(glog("ERROR: 'GulpfileConfig.preclean must specify a 'patterns' property if used"));
	else if (Array.isArray(gulpfileConfig.preclean.patterns) == false)
		return Promise.reject(glog("ERROR: 'GulpfileConfig.preclean.patterns' not specified as array in 'gulpfile.config.json'"));
	else
		return new Promise<void>((resolve, reject) => {
			tasksCalled.push("preClean");
			const options = gulpfileConfig.preclean.options ?? {} as deleteAsyncOptions;
			if (options?.preview)
				options.dryRun = options.preview;
			if (!options.dryRun) {
				glog(
"\n   In using Node 'deleteAsync' module, the default is that 'dryRun' is set 'true' for 'gulpfile.js' use." +
"\n   Actual deletions require that the 'preclean' property of 'gulpfile.config.json'" +
"\n   have `\"options\":{\"dryRun\":false}` or `\"options\":{\"preview\":false}` be set.");
				options.dryRun = true;
			}
			if (options.dryRun == true)
				glog("preclean has been set in 'dryRun'/'preview' mode");
			deleteAsync(
				GulpDict.resolveArrayOfStrings(gulpfileConfig.preclean.patterns),
				options
			).then((whatDone: string[]) => {
				if (options.dryRun == false)
					glog( "preclean: all items actually irretrievably deleted" );
				resolve(glog(JSON.stringify(whatDone, null, "  ")));
			}).catch((whatNotDone: unknown) => {
				reject(glog(JSON.stringify(whatNotDone, null, "  ")));
			});
		});
}

function tsCompile() {
	//	TSProject(gulpts.reporter.fullReporter(false));
	const gulpfiletsconfig = SelectedGulpfileConfig.config!.tsconfig;
	for (const prop in gulpfiletsconfig)
		gulpfiletsconfig[prop] = GulpDict.resolveString(gulpfiletsconfig[prop]);
	let tsConfigPath = GulpDict.resolveString(gulpfiletsconfig.path || "${projectRoot}/tsconfig.json");
	if (fs.existsSync(tsConfigPath) == false) {
		glog(`WARN: tsCompile():: tsConfigPath not found!\n  tsconfigPath = '${tsConfigPath}'` +
			`\ncurrent working directory: '${process.cwd()}'`);
		const newPath = pathMod.join(GulpDict.resolveString("${projectRoot}"),
					pathMod.basename(tsConfigPath));
		if (newPath !== tsConfigPath && fs.existsSync(newPath) == true) {
			tsConfigPath = newPath;
			glog(`tsCompile():: tsConfigPath found that works: '${tsConfigPath}'`);
		}
	}
	const TSProject = gulpts.createProject(tsConfigPath);
	gulpfiletsconfig.json = {
		files: TSProject.config.files,
		include: TSProject.config.include,
		exclude: TSProject.config.exclude,
		compilerOptions: TSProject.config.compilerOptions
	};
	gulpfiletsconfig.projectDirectory = TSProject.projectDirectory;
	{
		let collectedFiles: string[] = [],
			includeFiles: string[],
			excludeFiles: string[];
		if (gulpfiletsconfig.json.include) {
			includeFiles = globSync(gulpfiletsconfig.json.include);
			if (gulpfiletsconfig.json.exclude) {
				excludeFiles = globSync(gulpfiletsconfig.json.exclude);
				excludeFiles.forEach(element => {
					includeFiles = includeFiles.filter(item => item !== element);
				});
			}
			collectedFiles = includeFiles;
		}
		if (gulpfiletsconfig.json.files)
			collectedFiles =  collectedFiles.concat(gulpfiletsconfig.json.files);
		const resolvedFiles: string[] = [];
		for (const item of collectedFiles)
			if (pathMod.isAbsolute(item) == true)
				resolvedFiles.push(item);
			else
				resolvedFiles.push(fwdSlash(pathMod.join(gulpfiletsconfig.projectDirectory, item)));
		SelectedGulpfileConfig.config!.tsconfig.allInputFiles = resolvedFiles;
	}
	const compilerOptions = gulpfiletsconfig.compilerOptions;
	// post-modification of overriding options
	// https://www.npmjs.com/package/gulp-typescript
	const tscJsonOutDir = TSProject.config.compilerOptions.outDir;
	let gulpfileJsonOutDir = SelectedGulpfileConfig.config?.tsconfig.compilerOptions.outDir;
	const thisUndefinedOutDirOverrides = SelectedGulpfileConfig.config?.tsconfig.compilerOptions.thisUndefinedOutDirOverrides;

	gulpfileJsonOutDir = typeof gulpfileJsonOutDir == "string" && gulpfileJsonOutDir == "undefined" ?
			undefined : gulpfileJsonOutDir;

	let outDir = tscJsonOutDir && pathMod.isAbsolute(tscJsonOutDir) == true ?
			pathMod.join(TSProject.projectDirectory, tscJsonOutDir) : tscJsonOutDir;
	if (!outDir) {
		if (gulpfileJsonOutDir) {
			if (pathMod.isAbsolute(gulpfileJsonOutDir) == true)
				outDir = gulpfileJsonOutDir;
			else
				outDir = pathMod.join(TSProject.projectDirectory, gulpfileJsonOutDir);
		} else if (thisUndefinedOutDirOverrides == true)
			outDir = gulpfileJsonOutDir;
	} else if (thisUndefinedOutDirOverrides == true)
		outDir = gulpfileJsonOutDir;
	const sourceMap = TSProject.config.compilerOptions.sourceMap || gulpfiletsconfig.compilerOptions.sourceMap;
	glog("running TScompile()\n  Overriding options--" +
		`\n  gulpconfig-outdir undefined overrides: '${thisUndefinedOutDirOverrides}'` +
		`\n  tsconfig-outDir   = '${tscJsonOutDir}'` +
		`\n  gulpconfig-outDir = '${gulpfileJsonOutDir}'` +
		`\n  final-outDir = '${outDir}'` +
		`\n  tsconfig-sourceMap   = '${TSProject.config.compilerOptions.sourceMap}'` +
		`\n  gulpconfig-soureMap = '${gulpfiletsconfig.compilerOptions.sourceMap}'` +
		`\n  final-sourceMap = '${sourceMap}'`
	);
	{
		const filesList: {
			list: string[];
			path: string;
			xpiledPath: string;
			pathTo: string;
			flatten: boolean;
			filesCopyTo: string;
		} = {
			list: [],
			path: "filelist.json",
			xpiledPath: "",
			pathTo: "",
			flatten: false,
			filesCopyTo: ""
		};
		const TSProjectListing = gulpts.createProject(tsConfigPath);
		TSProjectListing.src()
		.pipe(gulpsrcmaps.init())
		.pipe(TSProjectListing(gulpts.reporter.fullReporter(true))).js
		.pipe(gulpsrcmaps.write(".", {includeContent: false, sourceRoot: "./"}))
		.pipe(gulpfilelist(filesList, { absolute: true }))
		.pipe(dest(outDir));
		SelectedGulpfileConfig.config!.tsconfig.filesList = filesList;
	}
	if (!outDir) {
		return new Promise<void>((resolve, reject) => {

			const cmdline = `tsc -p ${GulpDict.resolveString('${projectRoot}/tsconfig.nooutdir.json')} --listFiles` +
					((sourceMap == true) ? " --sourceMap" : "");
			exec(cmdline, (err, stdout, stderr) => {
				if (err)
					reject(glog(`ERROR: tsCompile()::exec('${cmdline}')\nmessage: ${err.message}`));
				else {
					SelectedGulpfileConfig.config!.tsconfig.filesList.list =
							stdout.split(/\r?\n/).filter(line => !line.includes('node_modules'));
					resolve(glog(`COMPLETED using tsCompile()::exec('${cmdline}')`));
				}
			});
		});
	} else {
		if (outDir.search(/\/build\/|\/Gulp\//) >= 0)
			outDir = fwdSlash(GulpDict.resolveString(pathMod.join("${projectRoot}", pathMod.basename(outDir))));
		if (pathMod.isAbsolute(outDir) == false)
			outDir = fwdSlash(GulpDict.resolveString(pathMod.join("${projectRoot}", outDir)));
		if (sourceMap == true)
			return TSProject.src()
				//.pipe(())
				.pipe(gulpsrcmaps.init())
				.pipe(TSProject(gulpts.reporter.fullReporter(true))).js
				.pipe(gulpsrcmaps.write(".", {includeContent: false, sourceRoot: "./"}))
				.pipe(dest(outDir));
		else
			return TSProject.src()
				.pipe(TSProject(gulpts.reporter.fullReporter(true))).js
				.pipe(dest(outDir));
	}
}

function processHtml(done: GulpCallback) {
	const gulpfileConfig = SelectedGulpfileConfig.config!;
	gulpfileConfig.tsconfig.filesList.list = gulpfileConfig.tsconfig.filesList.list.map(elem => elem.replace(/"/g, ""));
	const filesList: string[] = gulpfileConfig.tsconfig.filesList.list.map(elem => {
		return elem.search(/\.js$/) != -1 ? elem : undefined
	}).filter(Boolean) as string[];
	const subfolderPaths =  gulpfileConfig.html?.subfolderPaths ??  null,
		htmlConfig: HTMLdef | null = gulpfileConfig.html;
	if (subfolderPaths)
		for (const path in subfolderPaths)
			GulpDict.set(path, subfolderPaths[path]);
	if (!htmlConfig || !htmlConfig.templateTransform || htmlConfig.templateTransform.length == 0 || htmlConfig == null)
		done();
	else {
		const htmlPath = GulpDict.resolveString(htmlConfig.templateTransform[0].from);
//		getJsFilesCopyList().then((jsFilesCopyList: {from:string[];to:string}[]) => {
			fs.readFile(GulpDict.resolveString(htmlPath), (err, contentBuffer) => {
				if (err) {
					glog(`ERROR reading the HTML file: ${htmlPath}\nmessage: ${err.message}`);
					done();
				} else {
					let content: string = contentBuffer.toString("utf-8"),
						gulpfileItems: ({name:string;value:string}[] | string)[] = [];

					// replace links and scripts sections
					const sectiondefs: {
						secName: string;
						marker: string;
						element: {
							name: string;
							etago: boolean;
							attrib: string;
							other?: {name:string;value:string};
						};
					}[] = [ {secName: "links", marker: "<!-- %%%LINK ELEMENTS%%% -->",
							element:  {name: "link", etago: false, attrib: "href", other:{name:"rel",value:"stylesheet"}} },
						/* 'links': (string | { rel: "stylesheet" | string; href: string; secName: string' })[];
								  "href": "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css", */
						{secName: "scripts", marker: "<!-- %%%SCRIPT ELEMENTS%%% -->",
							element: {name: "script", etago: true, attrib: "src"}}
					];
					let subsection;
					if (gulpfileConfig.html)
						for (const sectiondef of sectiondefs) { // as defined above
							for (const gulpconfigSection in gulpfileConfig?.html)
								if (gulpconfigSection == sectiondef.secName) {
									if (sectiondef.secName == "links")
										// eslint-disable-next-line @typescript-eslint/no-explicit-any
										gulpfileItems = gulpfileItems.concat(setUpLinkScriptInput((gulpfileConfig.html as any)[gulpconfigSection]));
									else if (sectiondef.secName == "scripts") {
										for (const subsectionProp in gulpfileConfig?.html.scripts) {
											// eslint-disable-next-line @typescript-eslint/no-explicit-any
											subsection = (gulpfileConfig.html.scripts as any)[subsectionProp];
											if ((subsectionProp == "special") != null && subsection.length > 0)
												gulpfileItems = gulpfileItems.concat(setUpLinkScriptInput(subsection));
											else if ((subsectionProp == "module") != null) {
												if (subsection.tsconfigPathMaps)
													for (const pathMap of subsection.tsconfigPathMaps as tsconfigPathMaps[])
														if (pathMap.from) {
															const fromPattern = pathMap.from.match(/%%([^%]+)%%/);
															SelectedGulpfileConfig.config!.tsconfig.filesList.flatten =
																	pathMap.flatten ? pathMap.flatten : false;
															SelectedGulpfileConfig.config!.tsconfig.filesList.pathTo =
																		GulpDict.resolveString(pathMap.to);
															SelectedGulpfileConfig.config!.tsconfig.filesList.filesCopyTo =
																		GulpDict.resolveString(pathMap.filesCopyTo)
															if (fromPattern) {
																let toTarget: string;
																const toTargets: string[] = [],
																	basePath = `${GulpDict.get("projectRoot")}/${fromPattern[1]}`;
																SelectedGulpfileConfig.config!.tsconfig.filesList.xpiledPath = basePath;
																// const globs = globSync(`${basePath}/** /*.js`);
																const excepts = pathMap.except;
																for (const xpiledFile of filesList) {
																	if (pathMap.flatten == true)
																		toTarget = fwdSlash(GulpDict.resolveString(
																				`${pathMap.to}/${pathMod.basename(xpiledFile)}`));
																	else // pathMap.flatten == false
																		toTarget = fwdSlash(GulpDict.resolveString(
																				`${pathMap.to}/${xpiledFile.substring(basePath.length + 1)}`));
																	toTargets.push(toTarget);
																	if (excepts &&
																			excepts.findIndex(elem => xpiledFile.search(elem) >= 0) >= 0)
																		continue;
																	if (pathMap.typemodule  === true)
																		gulpfileItems.push([{name:"src",value:toTarget},{name:"type",value:"module"}]);
																	else
																		gulpfileItems.push(toTarget);
																}
												//				SelectedGulpfileConfig.config!.tsconfig.filesList.htmlsrctexts =
												//						toTargets.map(elem => fwdSlash(pathMod.join(basePath, pathMod.basename(elem))));
															}
														}
											}
										}
									}
									if (gulpfileItems.length > 0) {
										// eslint-disable-next-line no-useless-escape
										content = content.replace(new RegExp(`\\r?\\n\\s*${sectiondef.marker}`),	buildSection(sectiondef, gulpfileItems));
										gulpfileItems = [];
									}
								}
							}
					// perform replace, add, deletion
					const ops = [ "replace", "add", "del" ];
					let items: /* del */ string[] & /* replace */ {fromPattern: string; toText: string;}[] &
							/* add */ {insertionPoint: string; insertionText: string}[];
					let sectionContent = content;
					for (const op of ops) {
						// eslint-disable-next-line @typescript-eslint/no-explicit-any
						const htmlConfigOp = (htmlConfig as any)[op];
						if ((items = htmlConfigOp).length > 0)
							for (const item of items)
								switch (op) {
								case "replace":
									sectionContent = sectionContent.replace(item.fromPattern, item.toText);
									break;
								case "add":
									sectionContent = sectionContent.replace(item.insertionPoint, item.insertionPoint + item.toText);
									break;
								case "del":
									sectionContent = sectionContent.replace(item, "");
									break;
								}
					}
					content = sectionContent;
					const htmlWritePath = GulpDict.resolveString(htmlConfig.templateTransform[0].to);
					fs.writeFile(htmlWritePath, content, (err) => {
						if (err)
							glog(`ERROR writing the HTML file: '${htmlWritePath}'\nmessage: ${err?.message}`);
						else {
							glog(`processHtml()::fs.writeFile('${htmlWritePath}')`);
							done();
						}
					});
				}
			});
//		}).catch((err) => {
//			glog(`ERROR: processHtml()::getJsFilesCopyList()\n  errmsg: ${err.message}`);
//		});
	}

	/**
	 * @function setUpLinkScriptInput -- required to prepare attributes of the HTML
	 * 	written string element form
	 * @param info
	 * @returns either the string which is just the 'src' of 'script' elements
	 * 	or 'href' of 'link' elements, or of name=value format attributes
	 */
	function setUpLinkScriptInput(
		info: (string | Record<string, string>)[]
	): (string | {name:string;value:string}[])[] {
		const theSetup: (string | {name:string;value:string}[])[] = [];
		for (const item of info)
			if (typeof item === "string")
				theSetup.push(GulpDict.resolveString(item));
			else {
				const objItem: {name:string;value:string;}[] = [];
				for (const prop in item)
					if (prop.length > 0)
						objItem.push({
							name: prop,
							value: GulpDict.resolveString(item[prop])
						});
				theSetup.push(objItem);
			}
		glog("processHtml()::setUpLinkScriptInput()")
		for (const item of theSetup)
			glog(`  item: ${JSON.stringify(item)}`);
		return theSetup;
	}  // end of function

	/**
	 * @function buildTag -- creates written form of element (tag)
	 * @param elemInfo
	 * @returns the constructed HTML tag string
	 */
	function buildTag(
		elemInfo: {
			elemName: string;
			attributes: {name:string;value:string}[];
			etago: boolean;
		}
	): string {
		const tagLineLength = 70, halfWindow = 5;
		let tag: string = `\n    <${elemInfo.elemName} `;
		for (const attribute of elemInfo.attributes)
			if (attribute.name.length > 0)
				tag += ` ${attribute.name}="${attribute.value}"`;
		tag += ">";
		if (elemInfo.etago == true)
			tag += `</${elemInfo.elemName}>`;
		// breaking up line length:
		//  1) start with a position of the current position, add tagLineLength
		//  2) if position > tag.length, go to 7
		//  3) subtract to the halfWindow looking for SPACE char, go to #4
		//  4) if halfWindow reached, re-position to #1, start increasing position to find space, go to #5
		//  5) if #2 condition met, follow it; otherwise insert '\n' character at space
		//  6) loop back to #1
		//  7) end loop
		const tagChars = tag.split("");
		for (let segStart = 0, segEnd = tagLineLength, segPosn = segEnd - halfWindow;
				segEnd < tag.length;
				segPosn++ ) {
					if (tagChars[segStart + segPosn] == " ") {
						tagChars[segStart + segPosn] = "\n";
						segStart += segPosn;
						segEnd = segStart + tagLineLength;
						segPosn = segEnd - halfWindow;
					} else if (segStart + segPosn >= tag.length)
						break;
			}
		return tagChars.join("");
	} // end of function

	/**
	 * @function buildSection -- takes the strings built by support functions
	 * 	'buildTag' and 'setupLinkScriptInput' and constructs sections to
	 * 	replace in a template with HTML-commented markers
	 * @param sectionInfo -- data required to build the section skeleton
	 * @param items -- constructed tags
	 * @returns -- a string representing the entire replaceable HTML page part
	 */
	function buildSection(
		sectionInfo: {
			secName: string;
			marker: string;
			element: {
				name: string;
				etago: boolean;
				attrib: string;
				other?: {name:string;value:string};
			};
		},
		items: (string | {name:string;value:string}[])[]
	): string {
		// works on links and scripts.special and scripts.module.tsconfigPathMaps
		let sectionContent = "",
			attributes: {name: string; value: string}[];

		for (const item of items) { // should be iterable, item will be string or object as an element
			// if only a string, then attributes must be built
			if (typeof item === "string")
				attributes = [
					{
						name:sectionInfo.element.attrib,
						value: item
					},
					sectionInfo.element.other ? sectionInfo.element.other :
					{name:"", value:""}
				];
			else // if an object, the attributes are already built
				attributes = item ;

			sectionContent += buildTag({
				elemName: sectionInfo.element.name,
				attributes: attributes,
				etago: sectionInfo.element.etago
			});
		}
		glog(`--- HTML Section build: ${sectionInfo.secName}\n${JSON.stringify(sectionContent)}`);
		return sectionContent;
	} // end of function
}

function getSpecialFiles(
	specialTag: "$$TSCONFIGFILES" | "$$TSCONFIGINCLUDE"
): Promise<[string, string][]> {
	const list: [string, string][] = [],
		gulpfileConfig = SelectedGulpfileConfig.config!;

	return new Promise<[string, string][]>((resolve, reject) => {
		if (specialTag == "$$TSCONFIGFILES") {
			if (gulpfileConfig.tsconfig.json.files) {
				for (const file of gulpfileConfig.tsconfig.json.files)
					list.push([fwdSlash(gulpfileConfig.tsconfig.projectDirectory),	file]);
			}
			resolve(list);
		} else if (specialTag == "$$TSCONFIGINCLUDE") {
			if (gulpfileConfig.tsconfig.json.include) {
				const globRequests: Promise<string[]>[] = [];
				for (const pattern of gulpfileConfig.tsconfig.json.include)
					globRequests.push(new Promise<string[]>((resolve, reject) => {
						glob(pattern).then((files: string[]) => {
							resolve(files);
						}).catch((err: Error) => {
							glog(`ERROR: getSpecialFiles()::glob('${pattern}')` +
								`\nerrmsg: ${err.message}`);
							reject(err);
						});
					}));
				Promise.all(globRequests).then((files: string[][]) => {
					let joinedPatterns: string[] = [];
					for (const filesSet of files)
						joinedPatterns = joinedPatterns.concat(filesSet);
					for (const item of joinedPatterns)
						list.push([fwdSlash(gulpfileConfig.tsconfig.projectDirectory), item]);
					resolve(list);
				}).catch((err) => {
					reject(glog(`ERROR: getSpecialFiles()::glob().Promise.all(requests)` +
						`\nerrmsg: ${err.message}`));
				});
			}
		}
	});
}

function jsMapEdits(
	jsMapFilePaths: string[]
): Promise<void> {
	return new Promise<void>((resolve, reject) => {
		// find the common path
		let ithCol: number;
		const splitFilePaths = jsMapFilePaths.map(elem => elem.split(/[/\\]/));
		const output = splitFilePaths[0].map((_, colIndex) => splitFilePaths.map(row => row[colIndex]));
		for (ithCol = 0; ithCol < output.length; ithCol++)
			if (Array.from(new Set(output[ithCol])).length > 1)
				break;
		const newMapPaths = splitFilePaths.map((elem) => {
			return `./${elem.slice(ithCol).join("/")}`;
		});

		const jsMapFilesEdits: Promise<void>[] = [];
		for (const file of jsMapFilePaths)
			jsMapFilesEdits.push(new Promise<void>((resolve, reject) => {
				fs.readFile(file, (err, buffer) => {
					if (err)
						reject(glog(`ERROR: jsMapEdits('${file}')\n  errmsg:${err.message}`));
					else {
						const newMapPath = newMapPaths.find(
							elem => elem.search(pathMod.basename(file).replace(/\.js.map/, "")) > 0
						);
						const content = buffer.toString("utf8").replace(
							/"sources":\["[^"]+"\]/, `"sources":["${newMapPath?.replace(/\.js.map/, ".ts")}"]`
						);
						fs.writeFile(file, content, (err) => {
							if (err)
								reject(glog(`ERROR: jsMapEdits('${file}')\n  errmsg:${err.message}`));
							else
								resolve(glog(`edit completed: jsMapEdits('${file}')`));
						});
					}
				});
			}));
		Promise.all(jsMapFilesEdits).then((edit) => {
			resolve(glog("jsMapEdit()::Promise.all() succeeded"));
		}).catch((err) => {
			reject(glog(`ERROR: jsMapEdit()::Promise.all()\n  errmsg:${err.message}`));
		});
	});
}

/**
 * @function copyOperations -- copy source files to destinations task
 *   useful gulp plugins: gulp-rename, gulp-if, guil
 * @returns
 *   JSON guidelines "copy" property has 3 properties "js", "css", "img"
 *      each of these will have
 * 		"fromN": string[] (N = 1, 2,..) -- globbed use gulp.src()
 * 		"toN": string  (N = 1, 2, ...)   use gulp.dest()
 */

function copyOperations(): Promise<void> {
	tasksCalled.push("copyOperations");
	return new Promise<void>((resolve, reject) => {
		const gulpfileConfig = SelectedGulpfileConfig.config!;
		if (!gulpfileConfig.copy ||
					gulpfileConfig.copy.find(elem => elem.bypass)?.bypass == true)
			return resolve(glog(
				"copyOperations():: bypass was specified"
			));
		// set tsconfig copying done
		if (gulpfileConfig.tsconfig.filesList.list.length > 0) {
			const copyFromList = gulpfileConfig.tsconfig.filesList.list,
				newList: (string | {src:string;dest:string})[] = [],
				xpiledPathLength = gulpfileConfig.tsconfig.filesList.xpiledPath.length,
				srcPathLength = `${gulpfileConfig.tsconfig.projectDirectory}/src`.length,
				flatten = gulpfileConfig.tsconfig.filesList.flatten,
				pathTo = gulpfileConfig.tsconfig.filesList.filesCopyTo,
				srcMaps = gulpfileConfig.tsconfig.compilerOptions.sourceMap,
				srcFiles = gulpfileConfig.tsconfig.allInputFiles,
				jsMapPaths: string[] = [];
			let neededPaths: string[] = [];
			for (const item of copyFromList) {
				let tsFile: string | undefined = undefined;
				if (srcMaps == true && item.search(/\.js\.map/) > 0) {
					const basenameRoot = pathMod.basename(item).replace(/\.js\.map/, "");
					tsFile = srcFiles.find(elem => elem.search(basenameRoot) != -1);
				}
				if (flatten == true) {
					newList.push(item);
					if (tsFile) {
						newList.push(tsFile);
						jsMapPaths.push(`${pathTo}/${pathMod.basename(tsFile!).replace(/\.ts/, ".js.map")}`);
					}
					neededPaths.push(pathMod.dirname(item));
				} else {
					newList.push({
						src: item,
						dest: `${pathTo}/${item.substring(xpiledPathLength + 1)}`
					});
					if (tsFile) {
						newList.push({
							src: tsFile,
							dest: `${pathTo}/${tsFile.substring(srcPathLength + 1)}`
						});
						jsMapPaths.push(`${pathTo}/${tsFile.substring(srcPathLength + 1).replace(/\.ts/, ".js.map")}`);
					}
					neededPaths.push(pathMod.dirname(`${pathTo}/${item.substring(xpiledPathLength + 1)}`));
				}
			}
			neededPaths = Array.from(new Set(neededPaths));
			const pathCheck: Promise<void>[] = [];
			for (const neededPath of neededPaths)
				pathCheck.push(new Promise<void>((resolve, reject) => {
					verifyPath(neededPath).then(() => {
						resolve(glog(`copyOperations()::verifying needed path '${neededPath}'`));
					}).catch((err) => {
						reject(glog(`ERROR copyOperations()::verifying needed path '${neededPath}\n  ${err.message}'`))
					});
				}));
			Promise.all(pathCheck).then(() => {
				glog(`copyOperations()::all needed paths verified`);
				performCopyAction(
					newList,
					flatten == true ? gulpfileConfig.tsconfig.filesList.pathTo : undefined
				).then((response) => {
					glog(response.toString());
					// JS MAP file edits here
					jsMapEdits(jsMapPaths) .then(() => {
						gulpCopyOperation().then(() => {
							resolve(glog("copyOperations() all copy operations done"));
						}).catch((err) => {
							reject(glog(`ERROR: copyOperations():: all copy operations failure\n  errmsg:${err.message}`));
						});
					}).catch((err) => {
						reject(glog(`ERROR: copyOperations():: jsMapEdits()  errmsg:${err.message}`));
					});
				}).catch((err) => {
					reject(glog(`ERROR: copyOperations()::performCopyAction('Script JS files')\n  errmsg:${err.message}`));
				});
			}).catch((err) => {
				reject(glog(`copyOperations() paths verification check failed\n ${err.message}`));
			});
		}
		// now the 'copy' section
		function gulpCopyOperation() {
			return new Promise((resolve, reject) => {
				if (!gulpfileConfig.copy)
					resolve(glog(`copyOperations(): `));
				else if (gulpfileConfig.copy.length == 0)
					resolve(glog(`copyOperations(): 'Copy' property 'gulpfile.config.json' but has no items`));
				else {
					const eachCopyAction: Promise<void>[] = [];
					let copyActionCount: number = 0;
					glog(`copyOperation::gulpCopyOperation ${gulpfileConfig.copy.length} copy items`);
					for (const copyAction of gulpfileConfig.copy) {
						if (("bypass" in copyAction) == true)
							continue;
						eachCopyAction.push(new Promise<void>((resolve, reject) => {
							copyActionCount++;
							const eachGlob: Promise<boolean>[] = [];
							const assertCopyAction: {from: string[]; to: string;flatten?:boolean} = copyAction;
							glog(`${assertCopyAction.from.length} assertCopyAction.from elements`);
							let globsCount: number = 0;
							for (const fromItem of assertCopyAction.from) {
								eachGlob.push(new Promise<boolean>((resolve, reject) => {
									globsCount++;
									glog(`#${globsCount} glob item`);
									const resolvedPattern = GulpDict.resolveString(fromItem);
									if (resolvedPattern.search(/TSCONFIGFILES|TSCONFIGINCLUDE/) > 0)
										getSpecialFiles(resolvedPattern as "$$TSCONFIGFILES" | "$$TSCONFIGINCLUDE")
										.then((filesSet: [string, string][]) => {
											// an array of a 3-string array which are all the TS files
											// 1. the effective project root path, 2. relative file path
											let fromPaths: string[] = [];
											for (const fileSet of filesSet)
												if (fileSet[1].search(/\.d\.ts$/) < 0)
													fromPaths.push(fwdSlash(pathMod.join(gulpfileConfig.tsconfig.projectDirectory, fileSet[1])));
											fromPaths = fromPaths.concat(
													gulpfileConfig.tsconfig.filesList.list.map(elem => elem.replace(/"/g, ""))
											);
											// only need first element with target path
											assertCopyAction.to = GulpDict.resolveString(assertCopyAction.to);
											verifyPath(assertCopyAction.to)
											.then((message: string) => {
												glog(`copyOperations::gulpCopyOperation::verifyPath success: ${message}`);
												performCopyAction(fromPaths, assertCopyAction.to).then(() => {
													resolve(true);
												}).catch((err) => {
													reject(glog(`ERROR copyOperations(): performCopyAction()\n` +
														`  errmsg: '${err.message}'`));
												});
											}).catch((message) => {
												glog(`ERROR: copyOperations::gulpCopyOperation::verifyPath failure: ${message}`);
											});
										}).catch((err) => {
											reject(glog(`ERROR copyOperations(): getSpecialFiles()\n` +
											`  errmsg: '${err.message}'`));
										});
									else
									// all other files not TS or JS
										glob(resolvedPattern).then((matches) => {
											glog(`copyOperations() -->\n  glob('${resolvedPattern}')`);
											if (matches.length == 0) {
												glog("copyOperations()::glob(): no matches to specified pattern" +
													"\n   check 'gulpfile.config.json' 'copy' specification to see if should be present.");
												resolve(true);
											} else {
												const toPath = GulpDict.resolveString(assertCopyAction.to);
												verifyPath(toPath).then((info) => {
													glog(`copyOperations() -->\n  verifyPath('${toPath}'): '${info}'` +
														"\n  performCopyAction(" +
														`\n  matches: ${JSON.stringify(matches)}\n)`);
													performCopyAction(matches, toPath).then(() => {
														resolve(true);
													}).catch((err) => {
														reject(glog(`ERROR copyOperations(): performCopyAction(${toPath.length} files)` +
														`\n  errmsg: '${err.message}'`));
													});
												}).catch((err) => {
													glog(`ERROR copyOperations()::gulpCopyOperations()::glob()::verifyPath(${toPath.length} files)` +
															`\n  errmsg: '${err}'`);
													reject(err);
												});
											}
										}).catch((err) => {
											reject(glog(`ERROR copyOperations()::gulpCopyOperations()::glob('${GulpDict.resolveString(fromItem)}')` +
											`\n  errmsg: '${err.message}'`));
										});
								}));
								Promise.all(eachGlob).then((responses) => {
									resolve(glog(`copyOperations(): Promise.all(GLOB) responses = ${responses.length}) globs count: ${globsCount}\n`));
								}).catch((err) => {
									reject(glog(`ERROR copyOperations(): copyAction::Promise.all(GLOB): called counts = ${globsCount}` +
										`\n  errmsg: '${JSON.stringify(err)}'`));
								});
							}
						}));
					}
					Promise.all(eachCopyAction).then((responses) => {
						resolve(glog(`copyOperations()::Promise.all(COPY ACTION): returned counts = ${responses.length}`));
					}).catch((err) => {
						reject(glog(`ERROR copyOperations()::Promise.all(COPY ACTION): called counts = ${copyActionCount}` +
						`\n  errmsg: '${err.message}'`));
					});
				}
			});
		}
	});

	/* NodeJS fs.copyFile() requires both src and dest have file name
		This function does multiple file copying to ONE destination folder
		The path basename on the destination will be appended to the destination folder!!
	*/
	function performCopyAction(
		copyItemSet: (string | { src: string; dest: string })[],
		dest?: string    // this value should be path to directory
	): Promise<boolean> {
		return new Promise((resolve, reject) => {
			let msg: string = "";
			if (typeof copyItemSet[0] == "string" &&
					typeof (copyItemSet[0] as unknown as {src:string;dest:string}).src == "undefined") {
				msg += "\tArg 'copyItemSet' is `string[]` type";
				if (typeof dest == "undefined")
					msg += "\tArg 'dest?' is undefined:" +
						"\nERROR arg 'dest' cannot be undefined if 'copyItemSet' is `string[]` type ";
				else
					msg += `\tArg 'dest?' is defined: '${dest}'`;
			} else if (typeof (copyItemSet[0] as unknown as {src:string;dest:string}).src == "undefined") {
				msg += "\tArg 'copyItemSet' is not of type `string[]` so should be type `{src:string;dest:string}[]";
			}
			glog(`copyOperations()::performCopyAction() variable analysis\n${msg}`);
			const eachCopyAction: Promise<boolean>[] = [];
			let copyActionCount: number = 1;
			glog(`${copyItemSet.length} copy actions to do\n${msg}`);
			for (const item of copyItemSet)
				eachCopyAction.push(new Promise((resolve, reject) => {
					let from: string, to: string;
					if (dest) {
						from = item as string;
						to = `${dest}/${pathMod.basename(item as string)}`;
					} else {
						const itemObj = item as {src:string;dest:string};
						from = itemObj.src
						to = itemObj.dest;
					}
					glog(`${copyActionCount++} copy:\n     '${from}'\n      -> ${to}`);
					fs.copyFile(from, to, (err) => {
						if (err)
							reject(glog(`ERROR performCopyAction(): request fs.copyFile(` +
								`\n   '${from}' ->` +
								`\n            '${to}')` +
								`\n  errmsg: '${err.message}'`));
						else {
							glog(`copied: '${from}' -->\n           '${to}'`);
							resolve(true);
						}
					});
				}));
			Promise.all(eachCopyAction).then((responses) => {
				glog(`copyOperations(): performCopyAction::Promise.all() responses = ${responses.length})` +
					`\n          copyActionCount = ${copyActionCount}`);
				resolve(true);
			}).catch((err) => {
				reject(glog(`ERROR copyOperations(): performCopyAction():: Promise.all(), copyActionCount = ${copyActionCount}` +
				`\n  errmsg: '${err.message}'`));
			});
		});
	}
}

function verifyPath(path: string): Promise<string> {
	return new Promise<string>((resolve, reject) => {
		if (fs.existsSync(path) == false)
			fs.mkdir(path, (err) => {
				if (err)
					reject(`ERROR: verifyPath('${path}')\n errmsg:${err.message}`);
				else
					resolve(`path ${path} created`)
			});
		else
			resolve(`path ${path} already exists`);
	});
}

function moveRename(): Promise<void> {
	return new Promise((resolve, reject) => {
		const gulpfileConfig = SelectedGulpfileConfig.config!;
		if (gulpfileConfig.moveRename.find(elem => elem.bypass)?.bypass == true)
			return resolve(glog(
				"moveRename()::  bypass specifed"
			));
		const moves: [string, string][] = gulpfileConfig.moveRename;
// first process globs
		const getFilesRequests: Promise<[string, string][]>[] = [];
		for (const move of moves)
			getFilesRequests.push(new Promise<[string, string][]>((resolve, reject) => {
				if (move[0] == "$$TSCONFIGFILES" || move[0] == "$$TSCONFIGINCLUDE")
					getSpecialFiles(move[0]).then((gottenfiles: [string, string][]) => {
					// an array of a 3-string array is returned:
					// 1. the effective project root path, 2. relative file path, 3. destination path
						resolve(gottenfiles);
					}).catch((err) => {
						reject(glog(`ERROR: moveRename()::getSpecialFiles('${move[0]}', '${move[1]}')\nerrmsg: ${err.message}`));
					});
				else
					resolve([[move[0], move[1]]]);
			}));
		Promise.all(getFilesRequests).then((collection: [string, string][][]) => {
			let reorderedCollection: [string, string][] = [],
				newItem: [string, string],
				newSet: [string, string][];
			for (const set of collection) {
				newSet = [];
				for (const item of set) {
					newItem = [GulpDict.resolveString(item[0]), GulpDict.resolveString(item[1])];
					newSet.push(newItem);
				}
				reorderedCollection = reorderedCollection.concat(newSet);
			}
			const globbingRequests: Promise<[string, string][]>[] = [];
			for (const item of reorderedCollection)
				globbingRequests.push(new Promise<[string, string][]>((resolve, reject) => {
					glob(item[0]).then((files) => {
						const tupled: [string, string][] = [];
						const addBasename: boolean = files.length > 1 ? true : false;
						let item1: string;
						for (const file of files) {
							if (addBasename == true)
								item1 = pathMod.join(item[1], pathMod.basename(file));
							else
								item1 = item[1];
							tupled.push([file, item1]);
						}
						resolve(tupled);
					}).catch((err) => {
						reject(err);
					});
				}));
			Promise.all(globbingRequests).then((results: [string, string][][]) => {
				let fullTuple: [string, string][] = [];
				for (const result of results)
					fullTuple  = fullTuple.concat(result)
				const moveOps: Promise<void>[] = [];
				for (const unitMove of fullTuple)
					moveOps.push(new Promise<void>((resolve) => {
						fs.rename(
							GulpDict.resolveString(unitMove[0]),
							GulpDict.resolveString(unitMove[1]),
							(err) => {
							if (err) {
								glog(`ERROR: moveRename()--fs.rename()\n${err.message}`);
								resolve();
							} else {
								glog(`moved/rename: ${unitMove[1]}`);
								resolve();
							}
						});
					}));
				Promise.all(moveOps).then(() => {
					resolve(glog(`COMPLETE moveRename()::moveOps[total = ${moveOps.length}]`));
				}).catch((err) => {
					reject(glog(`ERROR: moveRename()::globbingRequests[]\nerrmsg: ${err.message}`));
				});
			}).catch((err) => {
				reject(glog(`ERROR: moveRename()::globbingRequests[]\nerrmsg: ${err.message}`));
			});
		}).catch((err) => {
			reject(glog(`ERROR: moveRename()::getFilequests[]\nerrmsg: ${err.message}`));
		});
	});
}

function editOperations(): Promise<void> {
	return new Promise((resolve, reject) => {
		const gulpfileConfig = SelectedGulpfileConfig.config!;
// edits to all *.js files to remove 'import/export' for browser runtimes
		if (SelectedGulpfileConfig.task == "browser") {
			const flist = gulpfileConfig.tsconfig.filesList;
			for (const file of flist.list)
				if (file.search(/\.js$/) > 0) {
					const fpath = `${flist.filesCopyTo}` +
							(flist.flatten == true ? pathMod.basename(file) :
								file.substring(flist.xpiledPath.length));
					fs.readFile(fpath, (err, buffer) => {
						if (err)
							glog(`editOperations()::import/export edit-readFile('${pathMod.basename(fpath)}') error` +
								`  msg:'${err.message}'`);
						else {
							const content = buffer.toString("utf8")
								.replace(importExportRE, importExportReplace);
							fs.writeFile(fpath, content, (err) => {
								if (err)
									glog(`editOperations()::import/export edit-writeFile('${pathMod.basename(fpath)}') error` +
										`  msg:'${err.message}'`);
								else
									glog(`editOperations()::import/export edit-writeFile('${pathMod.basename(fpath)}') SUCCESS`);
							});
						}
					});
				}
		}

// now check for config.json edit specifications
		if (!gulpfileConfig.edits)
			resolve(glog("editOperations()::  gulpfile.config.json specifies no 'edits' property"));
		else if (gulpfileConfig.edits.find(elem => elem.bypass)?.bypass == true)
			resolve(glog("editOperations()::  bypass specified"));
		else {
			const edits: EditInfo[] = gulpfileConfig.edits;
			const editActionRequests: Promise<void>[] = [];
			for (const editInfo of edits) {
				const fPath = GulpDict.resolveString(editInfo.filepath);
				glog(`editOperations()::glob('${fPath}')`);
				editActionRequests.push(new Promise((resolve, reject) => {
					glob(fPath).then((files) => {
						if (files.length == 0)
							resolve(glog(`No files found for path '${fPath}'`));
						else {
							let fileNum = 0;
							glog(`editOps()::glob('${fPath}') returns ${files.length} file(s)`);
							for (const file of files)
								fs.readFile(file, (err, buffer) => {
									fileNum++;
									if (err)
										glog(`ERROR: editOperations()::fs.readFile('${file}')\n  ${err.message}`);
									else {
										let content = buffer.toString("utf-8");
										if (Array.isArray(editInfo.fixes) == false)
											glog("ERROR: editOperations()::the Gulp config json" +
												" for 'edits.fixes must be an array of objects. Fix the JSON");
										else {
											for (const fix of editInfo.fixes) {
												const re = new RegExp(GulpDict.resolveString(fix.target), fix.flags ? fix.flags : "");
												if (Array.isArray(fix.replace) == true)
													fix.replace = (fix.replace as string[]).join("");
												glog(`  Target pattern '${fix.target}' has ${content.match(re)?.length}` +
													" matches to content");
												content = content.replace(re, fix.replace as string);
											}
											fs.writeFile(file, content, (err) => {
												if (err) {
													if (fileNum == files.length)
														reject(glog(`ERROR: editOperations()--writeFile()\n${err.message}`));
													else
														reject(glog(`ERROR: editOperations()--writeFile()\n${err.message}`));
												} else if (fileNum == files.length)
													resolve(glog(`editOperations('${file}') completed with no error`));
												else
													resolve(glog(`editOperations('${file}') completed with no error`));
											});
										}
									}
								});
						}
					}).catch((err) => {
						reject(glog(`ERROR: editOperations()::glob(${fPath})\n  ${err.message}`));
					});
				}));
			}
			Promise.all(editActionRequests).then((response) => {
				resolve(glog(`All ${response.length} edit operations completed `));
			}).catch((err) => {
				reject();
			});
		}
	});
}

function packaging(): Promise<boolean> {
	if (SelectedGulpfileConfig.stage !== "release") {
		glog("Task 'packaging' can only be utilized for 'release' stage");
		return Promise.reject();
	}
	const releaseDir = "release";
	return new Promise((resolve, reject) => {
		const copyOps: Promise<boolean>[] = [];
		const copies: string[] = [];
		const packageRoot = GulpDict.get("projectRoot");

		const packageFiles: string[] = SelectedGulpfileConfig.config!.package;

		copyOps.push(new Promise<boolean>((resolve, reject) => {
			for (const file of packageFiles) {
				const componentPath = `${packageRoot}/${releaseDir}/${file}`;
				if (fs.existsSync(componentPath) == false) {
					glog(`ABORT packaging(): all specified files in 'package' Config JSON property ` +
						`must exist for this operation: file '${componentPath}' not found`);
					reject();
				}
				copies.push(`${packageRoot}/${file}`);
				fs.copyFile(componentPath, packageRoot, (err) => {
					if (err) {
						glog(`ERROR packaging() - copyFile()\n${err.message}`);
						reject(false);
					} else {
						glog("file copy succeeded");
						resolve(true);
					}
				});
			}
		}));
		Promise.all(copyOps).then((response) => {
			// now do the npm pack
			spawn("npm", ["pack"])
			.stdout.on("data", (data) => {
				console.log(data.toString());
			})
			.on("close", (code: number) => {
				if (code !== 0)
					glog(`ERROR: packaging() -- npm pack\ncode = ${code}`);
				else {
					const deleteOps: Promise<unknown>[] = [];
					for (const copy of copies)
						deleteOps.push(new Promise<unknown>((resolve, reject) => {
							deleteAsync(copy).then((response) => {
								resolve(response);
							}).catch((err) => {
								reject(false);
							});
						}));
					Promise.all(deleteOps).then((response) => {
						resolve(true);
					}).catch((err) => {
						reject(false);
					});
				}
			});
		}).catch((err) => {
			glog("ERROR: packaging()")
		});
	});
}


/* cleanup contains array of configuration objects to specify
	one or more actions to cleanup processeing. Each configuation will use the following parameters (only 'id' and 'path' required)
	id: string or number to identify a config (user can designate)
	path: sets the path where operations will be done during deletion.
		the current directory will be set to path to prevent possible wanted unanticipated deletions
	filter: this if a file mask pattern. If not set, default pattern = "*" to delete ALL files
	recursive: boolean that will include deletion of the masked pattern files in subdirectories
	deleteTree: boolean that will delete all subdirectories of the tree
	preview: instead of doing the deletion, this will create a report of all files/folders to be deleted

	del/delAsync options (Node module)
		Files and Directories Patterns: glob or regular expressions for matching file/dirs for deleteion
				deleteAsync(['dist/** /*.js']); // delete all *.js files under 'dist' tree
		Force: override any permissions
				deleteAsync(['temp'], { force: true });  // forced delete of all
		Dry Run: testing what would be deleted
				deleteAsync(['logs/** /*.log'], { dryRun: true }) // will return list of files to be deleted
		Concurrency: max number of deletion operations at one time.
				del(['cache'], { concurrency: 5 });
		Ignore: Specify exclusions from inclusion pattern
				del(['dist/**', '!dist/main.js']);  // dont delete main.js from this
		Dotfiles: exclude/include files starting with '.'
				del(['.*'], { dot: true })
		Match Base: match either basename or full path
				del(['** /test*.js'], { matchBase: true });
		Follow Symlinks: You can specify whether to follow symbolic links and delete the target files or directories.
				del(['dist/** /*'], { followSymlinks: true }); */
function cleanUp(done: GulpCallback) {
	tasksCalled.push("cleanUp");
	new Promise<string>((resolve, reject) => {
		const cleanupConfigs = SelectedGulpfileConfig.config!.cleanup;
		// set the path specifier as cwd; if no path, already done
		if (!cleanupConfigs)
			reject("No cleanup config specifed or no 'path' specified for cleanup");
		else
			for (const cleanupConfig of cleanupConfigs) {
				const id = cleanupConfig.id;
				if (!id)
					reject("Cleanup config requires an 'id' specifier (can be number or string)");
				else if (typeof id !== "number" && typeof id !== "string")
					reject("Cleanup config requires an 'id' of either type \"number\" or \"string\"");
				if (!cleanupConfig.path)
					reject("Cleanup config requires an 'path' specified");
				const path = GulpDict.resolveString(cleanupConfig.path);
				if (fs.existsSync(path) == false)
					reject(`The specified cleanup config 'path': '${path}' does not exist`);
				else { // all clear
					const options: deleteAsyncOptions = {
						dryRun: false,
						followSymlinks: true,
//						force: false,
						dot: true
					};
					let pattern = `${path}/` +
						(cleanupConfig.recursive == true ? "**" : "") +
						(cleanupConfig.filter ? cleanupConfig.filter : "");
					if (cleanupConfig.preview == true)
						options.dryRun = true;
					if (cleanupConfig.deleteTree)
						pattern = `${path}/*`;
					deleteAsync(GulpDict.resolveString(pattern), options).then((result) => {
						glog(`deleteAsync() report:\n${result}`);
						resolve(`Cleanup configuration ${cleanupConfig.id} successful`);
					}).catch((err) => {
						reject(`ERROR: cleanUp()::deleteAsync('${pattern}', ${JSON.stringify(options)})`);
					});
				}
			}
	}).then((result) => {
		glog(result);
		done();
	}).catch((err) => {
		glog(`ERROR cleanUp() task--see log\n${err}`);
		done();
	});
}

/*
function webpackBundle() {
	tasksCalled.push("webpackBundle");
	if (!GulpfileConfig.webpack)
		return Promise.resolve();
	return src(GulpDict.resolveString(GulpfileConfig.webpack?.entryJS)) // Replace 'src/entry.js' with the path to your entry file
		.pipe(webpackStream(webpackConfig))
		.pipe(dest(GulpDict.resolveString(GulpfileConfig.webpack?.dest))); // Replace 'dist/' with the desired output directory
}
*/
/*
const BrowserSetup = () => {
	browserSync.init({
		server: "./", // Serve files from the current directory
	});

	// Watch HTML, CSS, JS files and reload on changes
	watch(["*.html", "css/*.css", "js/*.js"]).on("change", browserSync.reload);
};
*/
/**********************************
 * EXPORTED FUNCTIONS, VARIABLES
 *********************************/
export default
series(
	systemSetup,
	preclean,
	tsCompile, // output to lib
	processHtml,
	copyOperations,
	moveRename,
	editOperations,
//	packaging,
//	webpackBundle,
	cleanUp
);

/*
function folderEntriesInfo(folderPath: string): dirEntryInfo[] {
	const entries: string[] = fs.readdirSync(folderPath);
	const entryDispo: dirEntryInfo[] = [];

	for (const entry of entries) {
		const entryPath = `${folderPath}/${entry}`,
			entryInfo = fs.statSync(entryPath);
		entryDispo.push({
			entryPath: entryPath,
			entryType: (entryInfo.isFile() == true) ? "FILE" : (entryInfo.isDirectory() == true) ? "FOLDER" : "UNKNOWN",
			isDeleted: false
		});
	}
	return entryDispo;
}
*/
/***************************************************
 * This gulpfile
 * gulp 'buildJSONex'
 * 1. Compile JSONtool.ts with tsconfig.JSONtool.json
 * 2. Copy these to './test': JS, html (for Node), any CSS
 * gulp 'browserTestJSONex'
 * 1. Check for files in test subfolder as above
 * 2. Open start a browser and load for debugging
 ***************************************************/


/****************************************************'
 * Required setup for Gulp
 *   1. npm install --global gulp-cli (Global install of gulp-cli)
 *   2. package.json devDependences of gulp
 *      npm install --save-dev gulp in project folder
 *   3. Create & test simple gulpfile
 *            function defaultTask(cb) {
 *               place code for your default task here
 *               cb();  // placed at end of task code to return
 *            }
 *            exports.default = defaultTask
 *   4. 'gulp' or 'gulp default' on command line
 *  https://gulpjs.com/docs/en/getting-started/quick-start
 ****************************************************/