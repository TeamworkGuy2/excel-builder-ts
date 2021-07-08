# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.7.1](N/A) - 2021-07-08
#### Changed
* Extract several parameter and field types into public interfaces


--------
### [0.7.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/9a1e834a13c844fc4cdc944e8b8409c3695207ef) - 2021-06-12
#### Changed
* Update to TypeScript 4.3


--------
### [0.6.2](https://github.com/TeamworkGuy2/excel-builder-ts/commit/c0ed9a93212b19b16118f59f55cbb8b4bf9b74f4) - 2020-11-25
#### Changed
* Fix issue with `StyleSheet` `BorderProperty.color` and `FontStyle.color` fields not supporting type `string`
* Split `StyleSheet` `Fill` interface into `PatternFill` and `GradientFill` and make `Fill` a union type

#### Fixed
* `Fill` (now changed to `PatternFill`) `bgColor` and `fgColor` fields having an incorrect custom color type with a misspelled `rbg` field, now those fields use the correct `Color` interface


--------
### [0.6.1](https://github.com/TeamworkGuy2/excel-builder-ts/commit/3f06fa54d88ffb1c22773da579c89fd53c56b4e9) - 2020-11-24
#### Changed
* Improve types in `StyleSheet`
* Reorganize `Util`
* Add first unit tests for `Util`


--------
### [0.6.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/c82cfe597bf45e83ac0be84a992272d9cbf07f23) - 2020-09-04
#### Changed
* Update to TypeScript 4.0


--------
### [0.5.2](https://github.com/TeamworkGuy2/excel-builder-ts/commit/df4ce3b7a57d68c64cc1a285b5cd80fca8deedca) - 2019-11-08
#### Changed
* Update to TypeScript 3.7


--------
### [0.5.1](https://github.com/TeamworkGuy2/excel-builder-ts/commit/b3a14d9f8e559e3071297c85aaa6bb6537137082) - 2019-07-16
#### Changed
* Update to TypeScript 3.5
* Cache attributes array lookup in `Utils.defaults`
* A few documentation tweaks


--------
### [0.5.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/66179ecb3f34789f9406eb97fe27d11027737841) - 2019-01-12
#### Added
* `Worksheet.dataValidations` with `toXML()` output
* `Worksheet` added `setAutoFilter()`, `setDataValidations()`, and `setPageSetupAndMargins()` (which got split out from `addPagePrintSetup()`)

#### Removed
* `Worksheet.addPagePrintSetup()`


--------
### [0.4.3](https://github.com/TeamworkGuy2/excel-builder-ts/commit/a9edc537d38c6e705e0aad42ec3edbbd18ddd36a) - 2018-12-29
#### Changed
* Update to TypeScript 3.2
* Update dev and @types/ dependencies


--------
### [0.4.2](https://github.com/TeamworkGuy2/excel-builder-ts/commit/698bc9cdfdeb9fdda47e6569a7714918b95e9ca3) - 2018-10-17
#### Changed
* Update to TypeScript 3.1
* Update dev dependencies and @types
* Enable `tsconfig.json` `strict` and fix compile errors
* Removed compiled bin tarball in favor of git tags


--------
### [0.4.1](https://github.com/TeamworkGuy2/excel-builder-ts/commit/41e334ce51bf376b53f4a7eee86efcffe6835bc7) - 2018-06-22
#### Added
* Support for worksheet `<autoFilter>` and workbook `<definedName name="_FilterDatabase">` which add filter/search dropdowns on column header cells


--------
### [0.4.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/42ade4c544f9c9d5810a25b17262467a1a83ccd6) - 2018-04-14
#### Changed
* Update to TypeScript 2.8
* Update tsconfig.json with `strictNullChecks: true`, `noImplicitReturns: true` and `forceConsistentCasingInFileNames: true`
* Added release tarball and npm script `build-package` to package.json referencing external process to generate tarball
* Cleanup JSZip dependency:
  * `ExcelBuilder.createFile()` now requires a JSZip instance or object with `file()` method instead of creating a JSZip instance from a constructor function and generating the zip file. i.e. use `ExcelBuilder.createFile(new JSZip(), ...).generateAsync(...)` instead of previously using `ExcelBuilder.createFile(JSZip, ...)`
  * `JSZip` is only required if using `ZipWorker`


--------
### [0.3.2](https://github.com/TeamworkGuy2/excel-builder-ts/commit/ce1061e1897a1b369c97a2d70d3da05510926b20) - 2018-02-28
#### Changed
* Update to TypeScript 2.7
* Update dependencies: mocha, @types/chai, @types/mocha
* Enable tsconfig.json `noImplicitAny` and add/refine missing types


--------
### [0.3.1](https://github.com/TeamworkGuy2/excel-builder-ts/commit/aae03a4578b41d6eef08af6d3908c875e6c5e4fc) - 2017-10-26
#### Fixed
* Fix `importScripts()` definition in ZipWorker to match the definition in `WorksheetExportWorker`


--------
### [0.3.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/eb4f2bb801e30a9549c34b1873d8dc545fcccb97) - 2017-10-26
#### Changed
* Re-organized the project into sub-folders: `drawings/`, `export/`, `util/`, `workbook/`, `worksheet/`, and `xml/`
* Tweaked some variable names and comments/documentation
* Improved some TypeScript types
* Updated README with note about eventual desire to merge/deprecate project in favor of xlsx-spec-models and xlsx-spec-utils libraries
* Upgraded to TypeScript 2.4


--------
### [0.2.1](https://github.com/TeamworkGuy2/excel-builder-ts/commit/da98cefcb04335ecd7387510aceae8b397bb9082) - 2017-06-29
#### Changed
* Fixed some StyleSheet type issues


--------
### [0.2.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/f8e4a5b0a06ca8c26154441b6b81ed7e0746b903) - 2017-06-29
#### Changed
* Added missing types and improved existing types (mostly in StyleSheet)
* Removed `new ActiveXObject("Microsoft.XMLDOM")` fallback from Util.createXmlDoc() since `document.implementation.createDocument()` is supported by all major browsers
* Throw new Error(string) instances instead of strings


--------
### [0.1.5](https://github.com/TeamworkGuy2/excel-builder-ts/commit/108cb24fdee9553c379e67c70abfc2bf92a74687) - 2017-05-09
#### Changed
* Simplified some documentation for Visual Studio
* Added some missing types
* Upgraded to TypeScript 2.3, added tsconfig.json, use npm @types/ definitions


--------
### [0.1.4](https://github.com/TeamworkGuy2/excel-builder-ts/commit/2aa41518ff614d1fa9e7e5e71326aace53cbe367) - 2016-12-31
#### Changed
* TypeScript 2.0 compatibility tweaks
* Merged up to latest excel-builder.js commit from 2016-10-30 (https://github.com/stephenliberty/excel-builder.js/commit/162004ecb6b745f33924fff3f22692638a296306)


--------
### [0.1.3](https://github.com/TeamworkGuy2/excel-builder-ts/commit/393f0edbe9189e49a2df9cd842a504af5401f569) - 2016-06-21
#### Changed
* Merged up to latest excel-builder.js commit from 2016-06-02 (https://github.com/stephenliberty/excel-builder.js/commit/86129145d6242d973a6ade31c1a9a2c80420b2f0)


--------
### [0.1.2](https://github.com/TeamworkGuy2/excel-builder-ts/commit/bf2bbb96b52e8a5dc6fb8156533abd74e9b05e59) - 2016-06-01
#### Changed
* Corrected/added some type definitions


--------
### [0.1.1](https://github.com/TeamworkGuy2/excel-builder-ts/commit/c0d76ebed850b73aeb7eed4c52830f27e1df7bae) - 2016-05-31
#### Changed
* Switched all applicable strings from single to double quotes
* Added some missing and best guess data types

#### Fixed
* Fixed some typeof comparison bugs
* ExcelBuilder createFile() and createFileAsync() were mistakenly instance functions, now static
* Picture prototype and constructor type are now correct
* Workbook.Drawing interface had incorrect type
* Workbook.addMedia() 'contentType' is now correctly optional


--------
### [0.1.0](https://github.com/TeamworkGuy2/excel-builder-ts/commit/67ec7eedbcb88d43ac4ad1c02130183c8b8126ef) - 2016-05-30
#### Added
Initial commit of TypeScript port of the [excel-builder.js](https://github.com/stephenliberty/excel-builder.js) library.

#### Changed
JSZip dependency in favor of requiring the caller to pass an instance of JSZip to this library

#### Removed
Removed underscore and require.js dependencies in favor of native javascript and CommonJS style imports/exports.