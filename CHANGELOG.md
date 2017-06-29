# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.2.0](N/A) - 2017-06-29
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