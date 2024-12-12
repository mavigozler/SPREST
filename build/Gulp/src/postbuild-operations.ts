"use strict";

// postbuild-operations.ts
type FileChangeOp = {
	opTitle: string;
	opId: number;
	dependsOn?: number[]; // this should be a contentChange ID
	contentChange?: {
		filePattern: string;
		edits: {from: RegExp | string; to: string}[];
	};
	copyFile?: {from: string; to: string;};
	moveRenameFile?: {from: string; to: string; };
	callbacks?: ((...args: FileChangeOp[]) => Promise<string[]>)[];
	done: boolean;
};

import fs from "fs";
import * as globPkg from "glob";
const { glob } = globPkg;
import process from "process";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url); // get the resolved path to the file
const __dirname = path.dirname(__filename); // get the name of the directory

const startingExecutionDirectory = process.cwd();
console.log(`=========== running ${process.argv[1]}` +
   `\ncwd: ${startingExecutionDirectory}` +
	`\nsetting working directory to: ${__dirname}`
);
process.chdir(__dirname);

const projectRoot = "D:/dev/_Projects/SPREST";
const excludePattern = `${__dirname}/postbuild-operations.js`
console.log(`Project root: ${projectRoot}`)

// files to modify
const JSchanges: {from: RegExp | string; to: string}[] = [
   { from: /\/\* insert eslint DISABLING for compiled JS here \*\//, to: "/* eslint-disable */" },
   { from:  /^import (.*) from "((\/?\.\.?\/)+)([^\n]+)(\.js)?";?/mg,
      to: "import $1 from \"$2$4.js\"" },
	{ from: /\.json\.js/, to: ".json" },
	{ from: /\.js\.js/, to: ".js" }
];

const fileChangeRequestList: FileChangeOp[] = [
	{
		opTitle: "Content change:  make changes to JS files in './build/Gulp/run/' folder",
		opId: 1,
		contentChange: { filePattern: `${projectRoot}/build/Gulp/run/**/*.js`, edits: JSchanges },
		done: false,
	},
	/*
	{
		opTitle: "Rename:  'gulpfile.js' to 'gulpfile.mjs'",
		opId: 2,
		dependsOn: [1],
		moveRenameFile: { from: `${projectRoot}/lib/gulpfile.js`, to: `${projectRoot}/lib/gulpfile.mjs` },
		done: false,
	}, */
];

fileChangeTasks(fileChangeRequestList);
process.chdir(startingExecutionDirectory);
//copyJSON();

function fileChangeTasks(
	fileChangesSpec: FileChangeOp[],
): void {
   for (const item of fileChangesSpec) {
		console.log(`\nFile change task: ${item.opTitle}`)
		if (item.dependsOn)
			// the dependency gets an entry in a callback list with opId
			for (const dependOnTask of item.dependsOn) {
				const request = fileChangesSpec.find(spec => spec.opId == dependOnTask);
				if (request) {
					if (!request.callbacks)
						request.callbacks = [];
					request.callbacks.push(taskProcess.bind(null, item));
				}
			}
		else
			taskProcess(item).then((responses: string[]) => {
				// deal with the report
				for (const response of responses)
					console.log(response);
				if (item.callbacks) // contains opId of req wanting callback
					for (const cb of item.callbacks)
						cb();
			}).catch((err) => {
				if (Array.isArray(err) == true)
					for (const item of err)
						console.log(item);
				else
					console.log(err);
			});
	}
}

function taskProcess(taskSpec: FileChangeOp): Promise<string[]> {
	return new Promise<string[]>((resolve) => {
		const action: Promise<string>[] = [];
		if (taskSpec.contentChange)  // contentChange
			action.push(new Promise<string>((resolve) => {
				contentChange(taskSpec.contentChange!).then((report: string) => {
					resolve(report + "\nEnd Content Change operation\n===========\n");
				}).catch((err) => {
					resolve(err);
				});
			}));
		else if (taskSpec.copyFile)
			action.push(new Promise<string>((resolve) => {
				copyFile(taskSpec.copyFile!).then((report: string) => {
					resolve(report + "\nEnd Copy File operation\n===========\n");
				}).catch((err) => {
					resolve(err);
				});
			}));
		else if (taskSpec.moveRenameFile)
			action.push(new Promise<string>((resolve) => {
				moveRenameFile(taskSpec.moveRenameFile!).then((report: string) => {
					resolve(report + "\nEnd Move/Rename File operation\n===========\n");
				}).catch((err) => {
					resolve(err);
				});
			}));
		Promise.all(action).then((response) => {
			taskSpec.done = true;
			resolve(response.concat(["\n===== Task operation complete\n===========\n"]));
		}).catch((err) => {
			resolve(err);
		});
	});
}

function contentChange(changeInfo: {
	filePattern: string;
	edits: {from: RegExp | string; to: string}[];
}): Promise<string> {
	return new Promise((resolve) => {
		let report =
`\n==========\nContent Change: pattern='${changeInfo.filePattern}', ${changeInfo.edits.length} edits specified`;
		glob(changeInfo.filePattern, {ignore: excludePattern})
		.then((files) => {
			report += `\n\nglob("${changeInfo.filePattern}")`;
			report += `\n  glob() => count: ${files.length}`;
			for (let i = 0; i < files.length; i++)
				report += `\n  [${(i + 1)}] ${files[i]}`;
			const contentChangeRead: Promise<string>[] = [];
			for (let i = 0; i < files.length; i++) {
				contentChangeRead.push(new Promise<string>((resolve) => {
					fs.readFile(files[i], "utf8", (err, content) => {
						report + `\n======\nFile #${i + 1}:  fs.readFile('${files[i]}')`;
						if (err) {
							report += `\nERROR: ${err.message}`;
							resolve(report);
						} else {
							let changeCount = 0;
							report += `\n\n file '${files[i]}' opened for reading`;
							for (const edit of changeInfo.edits) {
								const matches: RegExpMatchArray | null = content.match(edit.from);
								if (matches == null) {
									report += `\n\nNot found in content:\n  change '${edit.from}' ->\n  '${edit.to}'`;
									continue;
								}
								report += `\n\nChange pattern found:\n  ${edit.from} ->\n      \`${edit.to}\`` +
									`\n           match count = ${matches.length}`;
								for (let i = 0; i < matches.length; i++)
									report += `  [${(i + 1)}] ${matches[i]}`;
								content = content.replace(edit.from, edit.to);
								changeCount++;
							}
					//		report (header);
							if (changeCount > 0) {
								fs.writeFile(files[i], content, "utf8", (err) => {
									report += `\nFile #${i + 1}: writeFile(${files[i]}`;
									if (err)
										report += `\n  ERROR: writing file '${files[i]}'\n    ${err.message}\n`;
									else
										report += `\n  file '${files[i]}' successfully modified`;
									resolve(report);
								});
							} else
								resolve(report += `\n\nFile #${i + 1}: '${files[i]}' no need to modify`);
						}
					});
				}));
			}
			Promise.all(contentChangeRead).then((responses) => {
				for (const response of responses)
					report += response;
				resolve(report);
			}).catch((err) => {
				resolve(err);
			});
		}).catch((err) => {
			report += `\nERROR: ${err.message}`;
			resolve(report);
		});
	});
}

function copyFile(copyInfo: {from: string; to: string;}, isRename?: boolean): Promise<string> {
	return new Promise((resolve) => {
		let report = (isRename == false) ? "\n\n\n==========\nCopy File" : "\n==========\nMove/Rename File",
			operation:  (src: fs.PathLike, dest: fs.PathLike, callback: fs.NoParamCallback) => void;

		report += `\n  '${copyInfo.from}' -> '${copyInfo.to}'`;
		if (isRename == false)
			operation = fs.copyFile;
		else
			operation = fs.rename;

		operation(copyInfo.from, copyInfo.to, (err) => {
			if (err)
				report +=
						`\nERROR: ${err.message}` +
						`\n   resolved copy source path:  ${path.resolve(copyInfo.from)}` +
						`\n   resolved copy target path:  ${path.resolve(copyInfo.to)}\n`;
			else
				report += "\n    File copy success";
			resolve(report);
		});
	});
}

function moveRenameFile(renameInfo: {from: string; to: string;}): Promise<string> {
	return copyFile(renameInfo, true);
}

/*
const targetFolder = `${projectRoot}/html`;
if (fs.existsSync(targetFolder) == false) {
    console.log(`\nERROR:  post-transpile.js::fs.existsSync(): target folder '${targetFolder}' does not exist!`)
}

function copyJSON() {
	// copy JSON files to output
	const filePattern = `${projectRoot}/src/*.json`;
	glob(filePattern, (err, files) => {
		console.log("\nglob() -- File filtering operation");
	if (err)
			console.log(`ERROR 'postbuild-operations.js'::glob()\n'${err.message}'`);
		else
			for (const file of files) {
				if (fs.existsSync(file) == false) {
					console.log(`ERROR:  post-transpile.js::fs.existsSync(): file '${file}' does not exist!`);
				} else {
					const targetFile = `${targetFolder}/${path.basename(file)}`;
					fs.copyFile(file, targetFile, (err) => {
						console.log("\nfs.copyFile() -- File copy operation")
						if (err)
							copyError(err, file, targetFolder);
						else
							console.log(`copy:  '${file}'\n  -> ${targetFile}`);
					});
				}
			}
	});
}
*/

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function copyError(err: NodeJS.ErrnoException, file: string, folder: string) {
   console.log(`ERROR 'postbuild-operations.js'::fs.copyFile()\n'${err.message}'`);
   if (err.code == "EPERM") {
      console.log("EPERM issue");
      fs.stat(file, (err, stats) => {
         if (err)
            console.log(`ERROR 'postbuild-operations.js'::fs.stat() Problem getting info\n'${err.message}'`);
         else
            console.log(`File '${file}':    ${modeReport(stats.mode)}`)
      });
      fs.stat(folder, (err, stats) => {
         if (err)
            console.log(`ERROR 'postbuild-operations.js'::fs.stat() Problem getting info\n'${err.message}'`);
         else {
            console.log(`Folder '${folder}':   ${modeReport(stats.mode)}`)
         }
      });
//      isFileLocked(file);
   }
}

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function isFileLocked(file: string) {
   fs.open(file, 'r+', (err, fd) => {
      if (err) {
         if (err.code === 'EACCES') {
            console.error('File is locked');
         } else {
            console.error(`Error opening file: ${err}`);
         }
         return;
      }
      console.log('File is not locked');
      fs.close(fd, (err) => {
         if (err) {
            console.error(`Error closing file: ${err}`);
         }
      });
   });
}

function modeReport(mode: number) {
   return `\n   File mode: ${(mode & 0o7777).toString()}` +
      `\n   Is directory? ${(mode & 0o170000) === 0o40000}` +
      `\n   Read permission: ${(mode & 0o400) !== 0}` +
      `\n   Write permission: ${(mode & 0o200) !== 0}\n`;
}