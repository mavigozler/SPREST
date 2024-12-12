'use strict';

type GulpFilelistOptions = {
	absolute?: boolean;
	flatten?: boolean;
	relative?: boolean;
	removeExtensions?: boolean;
	destRowTemplate?: string | ((path: string) => string);
};

/* insert eslint DISABLING for compiled JS here */
// eslint-disable-next-line @typescript-eslint/no-var-requires
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const pkg = require("./package.json");
import PluginError from "plugin-error";
import through from "through2";
import pathMod from "path";
import VinylMod from "vinyl";

function gulpFileListMem(
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	out: {
		list: string[]; // to be used to store the file list and return to caller
		path: string; // necessary to continue any stream
	},  // output to variable as well as file (named by string)
	options?: GulpFilelistOptions
) {
	const fileList: string[] = [];
	let filePath: string;

	const setOptions = options || {} as GulpFilelistOptions;

	const stream = through.obj((
			file: VinylMod,
			enc: string,
			cb: (
				error?: Error | null,
				filePart?: unknown
			) => void
		) => {
			if (file.isNull())
				return cb(null, file);
			if (file.isStream())
				return cb(new PluginError(pkg.name, 'Streams not supported'));
			if (setOptions.absolute)
				filePath = pathMod.normalize(file.path);
			else if (setOptions.flatten)
				filePath = pathMod.basename(file.path);
			else if (setOptions.relative)
				filePath = file.relative;
			else
				filePath = pathMod.relative(file.cwd, file.path);

			if (setOptions.removeExtensions) {
				const extension = pathMod.extname(filePath);

				if (extension.length)
					filePath = filePath.slice(0, -extension.length);
			}
			filePath = filePath.replace(/\\/g, "/");
			if (setOptions.destRowTemplate && typeof setOptions.destRowTemplate === 'string')
				fileList.push(setOptions.destRowTemplate.replace(/@filePath@/g, filePath));
			else if (setOptions.destRowTemplate && typeof setOptions.destRowTemplate === 'function')
				fileList.push(setOptions.destRowTemplate(filePath));
			else
				fileList.push(filePath);
			cb();
		},
		function (cb: () => void) {
			const buffer = (setOptions.destRowTemplate) ? Buffer.from(fileList.join('')) :
					Buffer.from(JSON.stringify(fileList, null, '  '));
			if (out.list && Array.isArray(out.list) === true)
				out.list = buffer.toString().match(/("[^"]+")/g) as string[];
			const fileListFile = new VinylMod({
				path: out.path,
				contents: buffer
			});
			// eslint-disable-next-line @typescript-eslint/no-explicit-any
			(this as unknown as any[]).push(fileListFile);
			cb();
		}
	);
	return stream;
}

export default gulpFileListMem;