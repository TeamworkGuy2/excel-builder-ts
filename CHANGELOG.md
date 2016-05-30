# Change Log
All notable changes to this project will be documented in this file.
This project does its best to adhere to [Semantic Versioning](http://semver.org/).


--------
### [0.1.0](N/A) - 2016-05-30
#### Added
Initial commit of TypeScript port of the [excel-builder.js](https://github.com/stephenliberty/excel-builder.js) library.

#### Changed
JSZip dependency in favor of requiring the caller to pass an instance of JSZip to this library

#### Removed
Removed underscore and require.js dependencies in favor of native javascript and CommonJS style imports/exports.