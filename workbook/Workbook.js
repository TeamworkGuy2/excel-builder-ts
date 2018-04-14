"use strict";
var Util = require("../util/Util");
var Paths = require("../worksheet/Paths");
var RelationshipManager = require("../worksheet/RelationshipManager");
var SharedStrings = require("../worksheet/SharedStrings");
var StyleSheet = require("../worksheet/StyleSheet");
var Worksheet = require("../worksheet/Worksheet");
/** Return base64 encoded data for the printer seeings binary file for a default portrait,
 * 0.25 margin, Excel .xlsx spreadsheet
 * @returns a base64 encoded string with no initial 'data:...,' marker, just a base64 string of binary data
 */
function getXlsxPrinterSettings1binBase64() {
    return "UABEAEYAQwByAGUAYQB0AG8AcgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAFwDU++A" +
        "AQEAAQDqCm8IQgABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFBSSVbi" +
        "MAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAYAAAAAAAQJxAnECcAABAnAAAAAAAAAACIAFwDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAwAAAAAAAAAAABAAXEsD" +
        "AGhDBAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAA57FLTAMAAAAFAAoA/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAACIAAAAU01USgAAAAAQAHgAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
}
/**
 * @module Excel/Workbook
 */
var Workbook = /** @class */ (function () {
    function Workbook(config) {
        this.worksheets = [];
        this.tables = [];
        this.drawings = [];
        this.media = {};
        this.initialize(config);
    }
    Workbook.prototype.initialize = function (config) {
        this.id = Util._uniqueId("Workbook");
        this.styleSheet = new StyleSheet();
        this.sharedStrings = new SharedStrings();
        this.relations = new RelationshipManager();
        this.relations.addRelation(this.styleSheet, "stylesheet");
        this.relations.addRelation(this.sharedStrings, "sharedStrings");
    };
    Workbook.prototype.createWorksheet = function (config) {
        var cfg = (config != null ? config : {});
        if (cfg.name == null) {
            cfg.name = "Sheet ".concat(this.worksheets.length + 1);
        }
        return new Worksheet(cfg);
    };
    Workbook.prototype.getStyleSheet = function () {
        return this.styleSheet;
    };
    Workbook.prototype.addTable = function (table) {
        this.tables.push(table);
    };
    Workbook.prototype.addDrawings = function (drawings) {
        this.drawings.push(drawings);
    };
    /** Set number of rows to repeat for this sheet.
     * @param inSheet sheet name
     * @param inRowCount number of rows to repeat from the top
     */
    Workbook.prototype.setPrintTitleTop = function (inSheet, inRowCount) {
        if (this.printTitles == null) {
            this.printTitles = {};
        }
        if (this.printTitles[inSheet] == null) {
            this.printTitles[inSheet] = {};
        }
        this.printTitles[inSheet].top = inRowCount;
    };
    /** Set number of rows to repeat for this sheet.
     * @param inSheet sheet name
     * @param inColumn number of columns to repeat from the left
     */
    Workbook.prototype.setPrintTitleLeft = function (inSheet, inColumn) {
        if (this.printTitles == null) {
            this.printTitles = {};
        }
        if (this.printTitles[inSheet] == null) {
            this.printTitles[inSheet] = {};
        }
        //WARN: this does not handle AA, AB, etc.
        this.printTitles[inSheet].left = String.fromCharCode(64 + inColumn);
    };
    Workbook.prototype.addMedia = function (type, fileName, fileData, contentType) {
        var fileNamePieces = fileName.split('.');
        var extension = fileNamePieces[fileNamePieces.length - 1];
        if (!contentType) {
            switch (extension.toLowerCase()) {
                case "jpeg":
                case "jpg":
                    contentType = "image/jpeg";
                    break;
                case "png":
                    contentType = "image/png";
                    break;
                case "gif":
                    contentType = "image/gif";
                    break;
                default:
                    contentType = null;
                    break;
            }
        }
        if (!this.media[fileName]) {
            this.media[fileName] = {
                id: fileName,
                data: fileData,
                fileName: fileName,
                contentType: contentType,
                extension: extension
            };
        }
        return this.media[fileName];
    };
    Workbook.prototype.addWorksheet = function (worksheet) {
        this.relations.addRelation(worksheet, "worksheet");
        worksheet.setSharedStringCollection(this.sharedStrings);
        this.worksheets.push(worksheet);
    };
    Workbook.prototype.createContentTypes = function () {
        var doc = Util.createXmlDoc(Util.schemas.contentTypes, "Types");
        var types = doc.documentElement;
        types.appendChild(Util.createElement(doc, "Default", [
            ["Extension", "rels"],
            ["ContentType", "application/vnd.openxmlformats-package.relationships+xml"]
        ]));
        types.appendChild(Util.createElement(doc, "Default", [
            ["Extension", "xml"],
            ["ContentType", "application/xml"]
        ]));
        var extensions = {};
        for (var filename in this.media) {
            extensions[this.media[filename].extension] = this.media[filename].contentType;
        }
        for (var extension in extensions) {
            types.appendChild(Util.createElement(doc, "Default", [
                ["Extension", extension],
                ["ContentType", extensions[extension]]
            ]));
        }
        types.appendChild(Util.createElement(doc, "Override", [
            ["PartName", "/xl/workbook.xml"],
            ["ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"]
        ]));
        types.appendChild(Util.createElement(doc, "Override", [
            ["PartName", "/xl/sharedStrings.xml"],
            ["ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"]
        ]));
        types.appendChild(Util.createElement(doc, "Override", [
            ["PartName", "/xl/styles.xml"],
            ["ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"]
        ]));
        for (var i = 0, l = this.worksheets.length; i < l; i++) {
            types.appendChild(Util.createElement(doc, "Override", [
                ["PartName", "/xl/worksheets/sheet" + (i + 1) + ".xml"],
                ["ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"]
            ]));
        }
        for (var i = 0, l = this.tables.length; i < l; i++) {
            types.appendChild(Util.createElement(doc, "Override", [
                ["PartName", "/xl/tables/table" + (i + 1) + ".xml"],
                ["ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"]
            ]));
        }
        for (var i = 0, l = this.drawings.length; i < l; i++) {
            types.appendChild(Util.createElement(doc, "Override", [
                ["PartName", "/xl/drawings/drawing" + (i + 1) + ".xml"],
                ["ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml"]
            ]));
        }
        return doc;
    };
    Workbook.prototype.toXML = function () {
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, "workbook");
        var wb = doc.documentElement;
        wb.setAttribute("xmlns:r", Util.schemas.relationships);
        var sheets = Util.createElement(doc, "sheets");
        for (var i = 0, l = this.worksheets.length; i < l; i++) {
            var sheet = doc.createElement("sheet");
            sheet.setAttribute("name", this.worksheets[i].name);
            sheet.setAttribute("sheetId", i + 1);
            sheet.setAttribute("r:id", this.relations.getRelationshipId(this.worksheets[i]));
            sheets.appendChild(sheet);
        }
        wb.appendChild(sheets);
        // now to add repeating rows
        var definedNames = Util.createElement(doc, "definedNames");
        var ctr = 0;
        for (var name in this.printTitles) {
            if (!this.printTitles.hasOwnProperty(name)) {
                continue;
            }
            var entry = this.printTitles[name];
            var definedName = doc.createElement("definedName");
            definedName.setAttribute("name", "_xlnm.Print_Titles");
            definedName.setAttribute("localSheetId", ctr++);
            var value = "";
            if (entry.top) {
                value += name + "!$1:$" + entry.top;
                if (entry.left) {
                    value += ",";
                }
            }
            if (entry.left) {
                value += name + "!$A:$" + entry.left;
            }
            definedName.appendChild(doc.createTextNode(value));
            definedNames.appendChild(definedName);
        }
        wb.appendChild(definedNames);
        return doc;
    };
    Workbook.prototype.createWorkbookRelationship = function () {
        var doc = Util.createXmlDoc(Util.schemas.relationshipPackage, "Relationships");
        var relationships = doc.documentElement;
        relationships.appendChild(Util.createElement(doc, "Relationship", [
            ["Id", "rId1"],
            ["Type", Util.schemas.officeDocument],
            ["Target", "xl/workbook.xml"]
        ]));
        return doc;
    };
    Workbook.prototype._generateCorePaths = function (files) {
        Paths[this.styleSheet.id] = "styles.xml";
        Paths[this.sharedStrings.id] = "sharedStrings.xml";
        Paths[this.id] = "/xl/workbook.xml";
        for (var i = 0, l = this.tables.length; i < l; i++) {
            files["/xl/tables/table" + (i + 1) + ".xml"] = this.tables[i].toXML();
            Paths[this.tables[i].id] = "/xl/tables/table" + (i + 1) + ".xml";
        }
        for (var fileName in this.media) {
            var media = this.media[fileName];
            files["/xl/media/" + fileName] = media.data;
            Paths[fileName] = "/xl/media/" + fileName;
        }
        for (var i = 0, l = this.drawings.length; i < l; i++) {
            files["/xl/drawings/drawing" + (i + 1) + ".xml"] = this.drawings[i].toXML();
            Paths[this.drawings[i].id] = "/xl/drawings/drawing" + (i + 1) + ".xml";
            files["/xl/drawings/_rels/drawing" + (i + 1) + ".xml.rels"] = this.drawings[i].relations.toXML();
        }
    };
    Workbook.prototype._prepareFilesForPackaging = function (files) {
        var contentTypes = this.createContentTypes();
        // adds reference for xl/printerSettings/printerSettings1.bin
        if (files["/xl/printerSettings/printerSettings1.bin"]) {
            contentTypes.documentElement.appendChild(Util.createElement(contentTypes, "Default", [
                ["Extension", "bin"],
                ["ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"]
            ]));
        }
        Object.assign(files, {
            "/[Content_Types].xml": contentTypes,
            "/_rels/.rels": this.createWorkbookRelationship(),
            "/xl/styles.xml": this.styleSheet.toXML(),
            "/xl/workbook.xml": this.toXML(),
            "/xl/sharedStrings.xml": this.sharedStrings.toXML(),
            "/xl/_rels/workbook.xml.rels": this.relations.toXML()
        });
        Object.keys(files).forEach(function (key) {
            var value = files[key];
            if (key.indexOf(".xml") != -1 || key.indexOf(".rels") != -1) {
                var resStr = files[key] = value.xml || new XMLSerializer().serializeToString(value);
                var content = resStr.replace(/xmlns=""/g, '');
                content = content.replace(/NS[\d]+:/g, '');
                content = content.replace(/xmlns:NS[\d]+=""/g, '');
                files[key] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n" + content;
            }
        });
    };
    Workbook.prototype.generateFilesAsync = function (options, worksheetExportWorkerPath) {
        var requireJsPath = options.requireJsPath;
        var self = this;
        if (!options.requireJsPath) {
            requireJsPath = document.getElementById("requirejs") ? document.getElementById("requirejs")["src"] : '';
        }
        if (!requireJsPath) {
            throw new Error("Please add 'requirejs' to the script that includes requirejs, or specify the path as an argument");
        }
        var files = {};
        var doneCount = this.worksheets.length;
        var stringsCollectedCount = this.worksheets.length;
        var workers = [];
        var result = {
            status: "Not Started",
            terminate: function () {
                for (var i = 0; i < workers.length; i++) {
                    workers[i].terminate();
                }
            }
        };
        this._generateCorePaths(files);
        function done() {
            if (--doneCount === 0) {
                self._prepareFilesForPackaging(files);
                for (var i = 0; i < workers.length; i++) {
                    workers[i].terminate();
                }
                options.success(files);
            }
        }
        function stringsCollected() {
            if (--stringsCollectedCount === 0) {
                for (var i = 0; i < workers.length; i++) {
                    workers[i].postMessage({
                        instruction: "export",
                        sharedStrings: self.sharedStrings.exportData()
                    });
                }
            }
        }
        for (var i = 0, l = this.worksheets.length; i < l; i++) {
            workers.push(this._createWorker(requireJsPath, i, function (worksheetIndex) {
                return {
                    error: function () {
                        for (var i = 0; i < workers.length; i++) {
                            workers[i].terminate();
                        }
                        //message, filename, lineno
                        options.error.apply(this, arguments);
                    },
                    stringsCollected: function () {
                        stringsCollected();
                    },
                    finished: function (data) {
                        files["/xl/worksheets/sheet" + (worksheetIndex + 1) + ".xml"] = { xml: data };
                        Paths[self.worksheets[worksheetIndex].id] = "worksheets/sheet" + (worksheetIndex + 1) + ".xml";
                        files["/xl/worksheets/_rels/sheet" + (worksheetIndex + 1) + ".xml.rels"] = self.worksheets[worksheetIndex].relations.toXML();
                        done();
                    }
                };
            }(i), worksheetExportWorkerPath));
        }
        return result;
    };
    Workbook.prototype._createWorker = function (requireJsPath, worksheetIndex, callbacks, worksheetExportWorkerPath) {
        var worker = new Worker(worksheetExportWorkerPath); //require.toUrl('./WorksheetExportWorker.js')
        var self = this;
        worker.addEventListener("error", callbacks.error);
        worker.addEventListener("message", function (event, data) {
            //console.log("Called back by the worker!\n", event.data);
            switch (event.data.status) {
                case "ready":
                    worker.postMessage({
                        instruction: "start",
                        data: self.worksheets[worksheetIndex].exportData()
                    });
                    break;
                case "sharedStrings":
                    for (var i = 0; i < event.data.data.length; i++) {
                        self.sharedStrings.addString(event.data.data[i]);
                    }
                    callbacks.stringsCollected();
                    break;
                case "finished":
                    callbacks.finished(event.data.data);
                    break;
            }
        }, false);
        worker.postMessage({
            instruction: "setup",
            requireJsPath: requireJsPath,
        });
        return worker;
    };
    Workbook.prototype.generateFiles = function () {
        var files = {};
        this._generateCorePaths(files);
        // TODO work-in-progress
        var anyPrintOptions = false;
        for (var i = 0, l = this.worksheets.length; i < l; i++) {
            files["/xl/worksheets/sheet" + (i + 1) + ".xml"] = this.worksheets[i].toXML();
            Paths[this.worksheets[i].id] = "worksheets/sheet" + (i + 1) + ".xml";
            if (this.worksheets[i]._printerSettings) {
                //anyPrintOptions = true;
                //Paths[this.worksheets[i]._printerSettings.id] = '../printerSettings/printerSettings1.bin';
            }
            files["/xl/worksheets/_rels/sheet" + (i + 1) + ".xml.rels"] = this.worksheets[i].relations.toXML();
        }
        if (anyPrintOptions) {
            files["/xl/printerSettings/printerSettings1.bin"] = getXlsxPrinterSettings1binBase64();
        }
        this._prepareFilesForPackaging(files);
        return files;
    };
    return Workbook;
}());
module.exports = Workbook;
