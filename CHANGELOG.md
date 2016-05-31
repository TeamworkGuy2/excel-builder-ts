# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.1.1](N/A) - 2016-05-31
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