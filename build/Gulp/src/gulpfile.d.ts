
export { ConfigJson, GulpCallback,
	GulpfileConfigStage, GulpfileConfigsInfo, HTMLdef, ConfigStage,
	tsconfigPathMaps, EditInfo, SelectedConfig, deleteAsyncOptions
};

type GulpCallback = () => void;

type EditInfo = {
	filepath: string;
	fixes: { target: string; replace: string | string[]; flags?: string}[];
};

type tsconfigPathMaps = {default: string;} & {
	from:string;
	to:string;
	filesCopyTo: string;
	typemodule?:boolean;
	except?:string[];
	flatten?:boolean;
};

type deleteAsyncOptions = {
	force?: boolean;
	dryRun?: boolean;
	preview?: boolean;
	concurrency?: number;
	dot?: boolean;
	matchBase?: boolean;
	followSymlinks?: boolean;
}

type HTMLdef = {
	templateTransform: {from: string; to: string}[];
	subfolderPaths?: { [key: string]: string };
	links?: (Record<string, string> | string)[];
	scripts?: {
		special: (Record<string, string> | string)[];
		module: {
			tsconfigPathMaps: tsconfigPathMaps[];
			// (Record<string, string> | string)[];
		}
	} ;
	replace?: { fromPattern: string; toText: string }[];
	add?: { insertionPoint: string; insertionText: string }[];
	del?: string[];
};

type TSConfigJSON = {
	compilerOptions?: {[key: string]: string};
	files?: string[];
	include?: string[];
	exclude?: string[];
};

type SelectedConfig = {
	task: string | undefined;
	stage: string | undefined;
	config: GulpfileConfigStage | undefined;
};

interface GulpfileConfigStage {
	base: {
		projectRoot: string;
		browserLoadPage: string;
	};
	preclean: {
		patterns: string[];
		options?: deleteAsyncOptions;
	};
	tsconfig: {
		path: string;
		compilerOptions: {
			// gulp-typescript specified
			sourceMap: boolean;
			outDir: string;
			thisUndefinedOutDirOverrides: boolean;
			rootDir: string; // use with outDir to control output dir structure
		}
		filesList: {
			list: string[];  // the file list retrieved by module
			path: string; // required for use of gulp-filelist modification
			xpiledPath: string;  // path to there transpiled files are delivered
			filesCopyTo: string; // from tsconfigPathMaps 'filesCopyTo'
			pathTo: string;  // from tsconfigPathMaps 'to'
			flatten: boolean; // from tsconfigPathMaps 'flatten'
		};
		json: TSConfigJSON;
		allInputFiles: string[];
		projectDirectory: string;
	};
	copy: ({ from: string[]; to: string;} & {bypass: boolean})[];
	moveRename: ([ string, string ] & {bypass: boolean})[];
	edits: ({   // use regular expressions
		filepath: string;
		fixes: {target: string; flags: string; replace: string;}[];
	} & {bypass: boolean})[];
	package: string[];
	html: HTMLdef | null;
	webpack?: {
		entryJS: string;
		dest: string;
	}
	cleanup: {
		id: number | string;
		path: string;
		filter?: string;
		recursive?: boolean;
		deleteTree?: boolean;
		preview?: boolean;

	}[];
}

type ConfigStage = {
	stageName: string;
	info: GulpfileConfigStage;
};

// This is the format of the actual file
type ConfigJson = {
	configNames: string[];
	base: {
		projectRoot: string;
		browserLoadPage: string;
	};
	outputReportPath?: string | undefined;
} & {
	[key: string]: Record<string, GulpfileConfigStage>;
};

// This is the memory structure organization of Gulpfile.config.json in code
interface GulpfileConfigsInfo {
	gulpfileJSdirectory: string;
	configNames: string[];
	defaultConfig: string;
	outputReportPath: string | undefined;
	fileJson: ConfigJson;
	configs: {
		configName: string;
		stages: ConfigStage[];
	}[];
}

