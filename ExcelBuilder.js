"use strict";
var Workbook = require("./workbook/Workbook");
/**
 * @name Excel
 * @public
 * @author Stephen Liberty
 * @requires Excel/Workbook
 * @requires JSZIP
 * @exports excel-builder
 */
var ExcelBuilder = /** @class */ (function () {
    function ExcelBuilder() {
    }
    /** Creates a new workbook.
     */
    ExcelBuilder.createWorkbook = function () {
        return new Workbook();
    };
    /** Turns a workbook into a downloadable file.
     * @param workbook The workbook that is being converted
     * @param options
     * options.base64 Whether to 'return' the generated file as a base64 string
     * options.success The callback function to run after workbook creation is successful.
     * options.error The callback function to run if there is an error creating the workbook.
     * options.requireJsPath The path to requirejs.
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
    /** Generates the xml/binary data files for a workbook and loads them into the provided object.
     * @param zip A JSZip library equivalent object with a file() method to add all the xlsx files to
     * @param workbook The workbook that is being converted
     * @param options options to modify how the excel doc is created. Only accepts a base64 boolean at the moment.
     * @returns the JSZip style object
     */
    ExcelBuilder.createFile = function (zip, workbook) {
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
        return zip;
    };
    return ExcelBuilder;
}());
module.exports = ExcelBuilder;
