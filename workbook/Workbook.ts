import Util = require("../util/Util");
import Drawings = require("../drawings/Drawings");
import Paths = require("../worksheet/Paths");
import RelationshipManager = require("../worksheet/RelationshipManager");
import SharedStrings = require("../worksheet/SharedStrings");
import StyleSheet = require("../worksheet/StyleSheet");
import Worksheet = require("../worksheet/Worksheet");
import XmlDom = require("../xml/XmlDom");

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
class Workbook {
    worksheets: Worksheet[];
    tables: Workbook.Table[];
    drawings: Workbook.Drawing[];
    media: { [id: string]: Workbook.Media };
    printTitles: {
        [sheetName: string]: {
            top?: number;
            left?: string;
        }
    };
    filterDatabases: {
        [sheetName: string]: {
            left: string; // Letter portion of cell name that defines left to right where filter database starts at
            right: string; // Letter portion of cell name that defines left to right where filter database ends at
            top: number; // Numeric portion of cell name that defines top to bottom where filter database starts at
            bottom: number; // Numeric portion of cell name that defines top to bottom where filter database ends at
        }
    };
    id: string;
    styleSheet: StyleSheet;
    sharedStrings: SharedStrings;
    relations: RelationshipManager;


    constructor(config?: any) {
        this.worksheets = [];
        this.tables = [];
        this.drawings = [];
        this.media = {};
        this.printTitles = {};
        this.filterDatabases = {};
        // copy from initialize() to appease TypeScript
        this.id = Util._uniqueId("Workbook");
        this.styleSheet = new StyleSheet();
        this.sharedStrings = new SharedStrings();
        this.relations = new RelationshipManager();
        this.relations.addRelation(this.styleSheet, "stylesheet");
        this.relations.addRelation(this.sharedStrings, "sharedStrings");
    }


    public initialize(config?: any) {
        this.id = Util._uniqueId("Workbook");
        this.styleSheet = new StyleSheet();
        this.sharedStrings = new SharedStrings();
        this.relations = new RelationshipManager();
        this.relations.addRelation(this.styleSheet, "stylesheet");
        this.relations.addRelation(this.sharedStrings, "sharedStrings");
    }


    public createWorksheet(config?: { name?: string; columns: Worksheet.Column[]; }) {
        var cfg = (config != null ? config : <any>{});
        if (cfg.name == null) {
            cfg.name = "Sheet ".concat(<any>this.worksheets.length + 1);
        }
        return new Worksheet(<typeof cfg & { name: string }>cfg);
    }


    public getStyleSheet() {
        return this.styleSheet;
    }


    public addTable(table: Workbook.Table) {
        this.tables.push(table);
    }


    public addDrawings(drawings: Workbook.Drawing) {
        this.drawings.push(drawings);
    }


    /** Set number of rows to repeat for this sheet.
     * @param inSheet sheet name
     * @param inRowCount number of rows to repeat from the top
     */
    public setPrintTitleTop(inSheet: string, inRowCount: number) {
        if (this.printTitles == null) {
            this.printTitles = {};
        }
        if (this.printTitles[inSheet] == null) {
            this.printTitles[inSheet] = {};
        }
        this.printTitles[inSheet].top = inRowCount;
    }

    
    /** Set number of rows to repeat for this sheet.
     * @param inSheet sheet name
     * @param inColumn number of columns to repeat from the left
     */
    public setPrintTitleLeft(inSheet: string, inColumn: number) {
        if (this.printTitles == null) {
            this.printTitles = {};
        }
        if (this.printTitles[inSheet] == null) {
            this.printTitles[inSheet] = {};
        }
        //WARN: this does not handle AA, AB, etc.
        this.printTitles[inSheet].left = String.fromCharCode(64 + inColumn);
    }


    public addMedia(type: any, fileName: string, fileData: any, contentType?: string) {
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
                    contentType = <any>null;
                    break;
            }
        }
        if (!this.media[fileName]) {
            this.media[fileName] = {
                id: fileName,
                data: fileData,
                fileName: fileName,
                contentType: <string>contentType,
                extension: extension
            };
        }
        return this.media[fileName];
    }


    public addWorksheet(worksheet: Worksheet) {
        this.relations.addRelation(worksheet, "worksheet");
        worksheet.setSharedStringCollection(this.sharedStrings);
        this.worksheets.push(worksheet);
    }


    public createContentTypes() {
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

        var extensions: { [id: string]: string } = {};
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
    }


    public toXML() {
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, "workbook");
        var wb = doc.documentElement;
        wb.setAttribute("xmlns:r", Util.schemas.relationships);

        var sheets = Util.createElement(doc, "sheets");
        for (var i = 0, l = this.worksheets.length; i < l; i++) {
            var sheet = doc.createElement("sheet");
            sheet.setAttribute("name", this.worksheets[i].name);
            sheet.setAttribute("sheetId", i + 1);
            sheet.setAttribute("r:id", this.relations.getRelationshipId(this.worksheets[i]))
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
                    value += ","
                }
            }
            if (entry.left) {
                value += name + "!$A:$" + entry.left;
            }

            definedName.appendChild(doc.createTextNode(value));
            definedNames.appendChild(definedName);
        }

        ctr = 0;
        for (var sheetName in this.filterDatabases) {
            if (!this.filterDatabases.hasOwnProperty(sheetName)) {
                continue;
            }
            var filterDatabase = this.filterDatabases[sheetName];
            var definedName = doc.createElement("definedName");
            definedName.setAttribute("name", "_xlnm._FilterDatabase");
            definedName.setAttribute("hidden", "1");
            definedName.setAttribute("localSheetId", ctr++);

            // Excel needs this format for a _FilterDatabase: "'Worksheet Name'!$A$11:$K$18"
            definedName.appendChild(doc.createTextNode("'" + sheetName + "'!$" + filterDatabase.left + "$" + filterDatabase.top + ":$" + filterDatabase.right + "$" + filterDatabase.bottom));
            definedNames.appendChild(definedName);
        }

        wb.appendChild(definedNames);

        return doc;
    }


    public createWorkbookRelationship() {
        var doc = Util.createXmlDoc(Util.schemas.relationshipPackage, "Relationships");
        var relationships = doc.documentElement;
        relationships.appendChild(Util.createElement(doc, "Relationship", [
            ["Id", "rId1"],
            ["Type", Util.schemas.officeDocument],
            ["Target", "xl/workbook.xml"]
        ]));
        return doc;
    }


    public _generateCorePaths(files: { [path: string]: any }) {
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
    }


    public _prepareFilesForPackaging(files: { [id: string]: XmlDom | string | { xml: string } }) {
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

        Object.keys(files).forEach((key) => {
            var value = files[key];
            if (key.indexOf(".xml") != -1 || key.indexOf(".rels") != -1) {
                var resStr = files[key] = (<{ xml: string }>value).xml || new XMLSerializer().serializeToString(<any><XmlDom>value);
                var content = resStr.replace(/xmlns=""/g, '');
                content = content.replace(/NS[\d]+:/g, '');
                content = content.replace(/xmlns:NS[\d]+=""/g, '');
                files[key] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n" + content;
            }
        });
    }


    public generateFilesAsync(options: { requireJsPath: string; success(files: { [id: string]: any }): void; error(...args: any[]): void; }, worksheetExportWorkerPath: string) {
        var requireJsPath = options.requireJsPath;
        var self = this;
        if (!options.requireJsPath) {
            requireJsPath = document.getElementById("requirejs") ? (<any>document.getElementById("requirejs"))["src"] : '';
        }
        if (!requireJsPath) {
            throw new Error("Please add 'requirejs' to the script that includes requirejs, or specify the path as an argument");
        }

        var files: { [id: string]: XmlDom | { xml: string } } = {};
        var doneCount = this.worksheets.length;
        var stringsCollectedCount = this.worksheets.length;
        var workers: Worker[] = [];

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
            workers.push(this._createWorker(requireJsPath, i, function (worksheetIndex: number) {
                    return {
                        error: function () {
                            for (var i = 0; i < workers.length; i++) {
                                workers[i].terminate();
                            }
                            //message, filename, lineno
                            options.error.apply(this, <any[]><any>arguments);
                        },
                        stringsCollected: function () {
                            stringsCollected();
                        },
                        finished: function (data: any) {
                            files["/xl/worksheets/sheet" + (worksheetIndex + 1) + ".xml"] = { xml: data };
                            Paths[self.worksheets[worksheetIndex].id] = "worksheets/sheet" + (worksheetIndex + 1) + ".xml";
                            files["/xl/worksheets/_rels/sheet" + (worksheetIndex + 1) + ".xml.rels"] = self.worksheets[worksheetIndex].relations.toXML();
                            done();
                        }
                    };
                } (i), worksheetExportWorkerPath)
            );
        }

        return result;
    }


    public _createWorker(requireJsPath: string, worksheetIndex: number, callbacks: { error(err: ErrorEvent): any; stringsCollected(): void; finished(data: any): void; }, worksheetExportWorkerPath: string) {
        var worker = new Worker(worksheetExportWorkerPath); //require.toUrl('./WorksheetExportWorker.js')
        var self = this;
        worker.addEventListener("error", callbacks.error);
        worker.addEventListener("message", <any>function (event: MessageEvent, data: any) {
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
            requireJsPath,
        });
        return worker;
    }


    public generateFiles(): { [key: string]: XmlDom | string } {
        var files: { [id: string]: XmlDom | string } = {};
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
    }

}

module Workbook {

    export interface Drawing {
        id: string;
        relations: { toXML(): XmlDom | string; };
        toXML(): XmlDom | string;
    }


    export interface Table {
        id: string;
        toXML(): string;
    }


    export interface Media {
        id: string;
        data: any;
        fileName: string;
        contentType: string;
        extension: string;
    }

}

export = Workbook;