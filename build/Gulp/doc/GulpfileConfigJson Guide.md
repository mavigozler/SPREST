# Gulpfile.Config.JSON Guide

This should be a guide to the gulpfile.config.json object and properties. Gulp starts with a task argument in the command line: 'gulp task-name'. If there is no task-name, the task default is 'default'.

This configuration file is not a feature of the Gulp system as composed by its creators. A gulpfile is supposed to contain all the hard coding to configure the file and not rely on external configuration data as a supplement in non-coding file form.

It is more organized for the gulpfile to have a JSON config itself to define its parameters. This config must not define any parameters that would be defined for TypeScript compilation that would go into a `tsconfig.json` file. This exists only for hard-coded stuff that would go in a gulpfile and pulls it out.

## Project directory structure/organization

```DOS
app-project/
├─ .vscode/
│  ├─ <project-name>.code-workspace [JSON]
│  ├─ launch.json
│  ├─ tasks.json
├─ build/
│  ├─ Gulp/
|  │  ├─ gulpfile.config.json
|  │  ├─ gulpfile.d.ts
|  │  ├─ gulpfile.ts
|  │  ├─ postbuild-operations.ts
|  │  ├─ tsconfig-copy.gulpfile.json
|  │  └─ webpack.config.ts
|  └─ run/
|     └─ build JS/JS.MAP files
├─ css/
|  └─ project CSS files
├─ doc/
|  └─ project documentation: txt, md files
├─ html/
│  ├─ Templates/
|  │  └─ <HTMLname>Template.html: HTML, ASPX
|  └─ project finished HTML or ASPX files
├─ js/
|  └─ project ready-for-test JS files
├─ lib/
|  └─ trans/compiled JS storage
├─ node_modules/
├─ release/ (symbolic link to final loc)
│  ├─ js/
│  ├─ css/
│  ├─ img/
|  └─ entryPoint.hmtl/.aspx, other markup files (html/xhtml)
├─ src/
|  └─ all TS files, subfolder grouped
├─ package.json/package-lock.json
└─ tsconfig.<scoped>.json
```

## Development/Test stages

1. source & config building: ts, css, html, json
2. compilation to '.lib' folder: ts -> js
3. move ("rename") to test:
   file moves to ./test the following subdirectories
   root: html/aspx, ./js: js + js.map, ./css: css only, ./img: all images
  html/aspx file edits
a) &lt;script&gt; to ./lib/js file checks, type=module? defer/async?
b) link to css

## Release

1. Create './release' subfolder under project folder
2. Create the subfolder structure as above to place resources for browser

```JSON
{
  "configNames": [ "default" ],
  "base": {
    "projectRoot": "D:/dev/_Projects/RESTRequesting",
    "browserLoadPage": ""
  },
  "default": {
    "test": {
      "preclean": null,
      "tsconfig": {
        "path": null  // will go to tsconfig.json
      },
      "copy": [
        { "from": [
          // for workingPathIndex, use -1 for current dir
          "${projectRoot}/lib/**/{*.js,*.js.map}"
          // "D:/dev/GenLib/src/html-docx.js"
           ],
           "to": "${projectRoot}/js"
        }
        /* {
          "from": [ "${projectRoot}/css/*.css" ],
          "to": "${projectRoot}/test/css"
        },
        {
           "from": [
           //"${projectRoot}/img/add-icon.jpg",
           ],
     "to": "${projectRoot}/test/img"
    } */
   ],
   "moveRename": [
    //["${projectRoot}/lib/index.js.map", "${projectRoot}/test/index.js.map"]
   ],
   "edits": [ // if "^" insert at start, if "$" append replace
   /* {
     "filepath": "${projectRoot}/lib/index.js",
     "fixes": [
      { "target": "import plugin from \"./bin/gulp-filelistex\";",
      "replace":  "import plugin from \"./bin/gulp-filelistex.js\";"}
     ]
    } */
   ],
   "html": {
    "templateTransform": [
     { "from": "${HTML_TEMPLATES}/RESTRequestingTemplate.html",
       "to": "${HTML}/RESTRequesting.html" }
    ],
    "subfolderPaths": {
     "JS": "../js",
     "CSS": "../css",
     "HTML": "${projectRoot}/html",
     "HTML_TEMPLATES": "${projectRoot}/html/Templates",
     "TEST": "../test"
    },
    "links": [
     "${CSS}/RestRequesting.css"
     //{
     // "rel": "stylesheet",
     // "href": "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css",
     // "integrity": "sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh",
     // "crossorigin": "anonymous"
     //}
    ],
    "scripts": {
     "special": [
      {
       "src": "https://code.jquery.com/jquery-3.6.0.min.js",
       "integrity": "sha256-/xUj+3OJU5yExlq6GSYGSHk7tPXikynS7ogEvDej/m4=",
       "crossorigin": "anonymous"
      }
     // { "src": "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js",
     //  "integrity": "sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6",
     //  "crossorigin": "anonymous" },
     //
     ],
     "module": {
      // format: "%%<string>%%" match a substring in from than map
      "tsconfigPathMaps": [
      // { "default": "./${JS}" },
       { "from": "%%lib%%", "to":"${JS}", "typemodule": false }
      ]
     }
    },
    "replace": [],
    "add": [],
    "del": []
   },
   "webpack": {
    "entryJS": "${projectRoot}/${JS}/RestRequestingjs",
    "dest": "${projectRoot}/${JS}"
   },
 /*   "cleanup": [{
    "path": "./js",
    "filter": "^(?:(?!\\.js$|\\.js\\.map$|\\.tsbuildinfo$).)*$",
    "recursive": true,
    "preview": true
   }] */
   "flags": [ "--verbose", "--force" ]
  },
/***** RELEASE STAGE ********/
  "release": {
   "preclean": null,
   "tsconfig": {
    "path": null  // will go to tsconfig.json
   },
   "copy": [
    { "from": [
     // for workingPathIndex, use -1 for current dir
      "${projectRoot}/lib/**/{*.js,*.js.map}"
     // "D:/dev/GenLib/src/html-docx.js"
     ],
     "to": "${projectRoot}/js"
    }
    /* {
     "from": [ "${projectRoot}/css/*.css" ],
     "to": "${projectRoot}/test/css"
    },
    {
     "from": [
      //"${projectRoot}/img/add-icon.jpg",
     ],
     "to": "${projectRoot}/test/img"
    } */
   ],
   "moveRename": [
    //["${projectRoot}/lib/index.js.map", "${projectRoot}/test/index.js.map"]
   ],
   "edits": [ // if "^" insert at start, if "$" append replace
   /* {
     "filepath": "${projectRoot}/lib/index.js",
     "fixes": [
      { "target": "import plugin from \"./bin/gulp-filelistex\";",
      "replace":  "import plugin from \"./bin/gulp-filelistex.js\";"}
     ]
    } */
   ],
   "html": {
    "templateTransform": [
     { "from": "${HTML_TEMPLATES}/RESTRequestingTemplate.html",
       "to": "${HTML}/RESTRequesting.html" }
    ],
    "subfolderPaths": {
     "JS": "../js",
     "CSS": "../css",
     "HTML": "${projectRoot}/html",
     "HTML_TEMPLATES": "${projectRoot}/html/Templates",
     "TEST": "../test"
    },
    "links": [
     "${CSS}/RestRequesting.css"
     //{
     // "rel": "stylesheet",
     // "href": "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css",
     // "integrity": "sha384-Vkoo8x4CGsO3+Hhxv8T/Q5PaXtkKtu6ug5TOeNV6gBiFeWPGFN9MuhOf23Q9Ifjh",
     // "crossorigin": "anonymous"
     //}
    ],
    "scripts": {
     "special": [
      {
       "src": "https://code.jquery.com/jquery-3.6.0.min.js",
       "integrity": "sha256-/xUj+3OJU5yExlq6GSYGSHk7tPXikynS7ogEvDej/m4=",
       "crossorigin": "anonymous"
      }
     // { "src": "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/js/bootstrap.min.js",
     //  "integrity": "sha384-wfSDF2E50Y2D1uUdj0O3uMBJnjuUD4Ih7YwaYd1iqfktj0Uod8GCExl3Og8ifwB6",
     //  "crossorigin": "anonymous" },
     //
     ],
     "module": {
      // format: "%%<string>%%" match a substring in from than map
      "tsconfigPathMaps": [
      // { "default": "./${JS}" },
       { "from": "%%lib%%", "to":"${JS}", "typemodule": false }
      ]
     }
    },
    "replace": [],
    "add": [],
    "del": []
   },
   "webpack": {
    "entryJS": "${projectRoot}/${JS}/RestRequestingjs",
    "dest": "${projectRoot}/${JS}"
   },
 /*   "cleanup": [{
    "path": "./js",
    "filter": "^(?:(?!\\.js$|\\.js\\.map$|\\.tsbuildinfo$).)*$",
    "recursive": true,
    "preview": true
   }] */
   "flags": [ "--verbose", "--force" ]
  }
 }
}
```
