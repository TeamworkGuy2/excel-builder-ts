/// <reference path="./excel-builder.d.ts" />
/// <reference path="../definitions/jszip/jszip.d.ts" />
"use strict";
var Workbook = require("./Workbook");
/**
 * @name Excel
 * @public
 * @author Stephen Liberty
 * @requires Excel/Workbook
 * @requires JSZIP
 * @exports excel-builder
 */
var ExcelBuilder = (function () {
    function ExcelBuilder() {
    }
    /** Creates a new workbook.
     */
    ExcelBuilder.createWorkbook = function () {
        return new Workbook();
    };
    /** Turns a workbook into a downloadable file.
     * @param {Excel/Workbook} workbook The workbook that is being converted
     * @param {Object} options
     * @param {boolean} options.base64 Whether to 'return' the generated file as a base64 string
     * @param {Function} options.success The callback function to run after workbook creation is successful.
     * @param {Function} options.error The callback function to run if there is an error creating the workbook.
     * @param {string} options.requireJsPath (Optional) The path to requirejs. Will use the id 'requirejs' to look up the script if not specified.
     */
    ExcelBuilder.createFileAsync = function (workbook, options, jszipPath, zipWorkerPath, worksheetExportWorkerPath) {
        workbook.generateFilesAsync({
            requireJsPath: options.requireJsPath,
            success: function (files) {
                var worker = new Worker(zipWorkerPath); //require.toUrl('./Excel/ZipWorker.js')
                worker.addEventListener("message", function (event) {
                    if (event.data.status == "done") {
                        options.success(event.data.data);
                    }
                });
                worker.postMessage({
                    files: files,
                    ziplib: jszipPath,
                    base64: (!options || options.base64 !== false)
                });
            },
            error: function () {
                options.error();
            }
        }, worksheetExportWorkerPath);
    };
    /** Turns a workbook into a downloadable file.
     * @param {JSZip} jszip A JSZip equivalent library to use to generate/zip the excel file
     * @param {Excel/Workbook} workbook The workbook that is being converted
     * @param {Object} options - options to modify how the excel doc is created. Only accepts a base64 boolean at the moment.
     */
    ExcelBuilder.createFile = function (jszip, workbook, options) {
        var zip = new jszip();
        var files = workbook.generateFiles();
        Object.keys(files).forEach(function (path) {
            var content = files[path];
            path = path.substr(1);
            if (path.indexOf(".xml") !== -1 || path.indexOf(".rel") !== -1) {
                zip.file(path, content, { base64: false });
            }
            else {
                zip.file(path, content, { base64: true, binary: true });
            }
        });
        return zip.generate({
            base64: (!options || options.base64 !== false)
        });
    };
    return ExcelBuilder;
}());
module.exports = ExcelBuilder;
