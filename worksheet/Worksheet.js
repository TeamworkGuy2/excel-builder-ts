"use strict";
var Util = require("../util/Util");
var RelationshipManager = require("./RelationshipManager");
/**
 * This module represents an excel worksheet in its basic form - no tables, charts, etc. Its purpose is
 * to hold data, the data's link to how it should be styled, and any links to other outside resources.
 */
var Worksheet = /** @class */ (function () {
    /**
     * @constructor
     */
    function Worksheet(config) {
        this.relations = null;
        this.columns = [];
        this.columnFormats = [];
        this.data = [];
        this.mergedCells = [];
        this._headers = [];
        this._footers = [];
        this._tables = [];
        this._drawings = [];
        this._orientation = null;
        this._margin = null;
        this.sharedStrings = { strings: {} };
        this.customCellAttributes = [];
        this.customRowAttributes = [];
        this._printerSettings;
        this.pageMargins = null;
        this.autoFilter = null;
        this.pageSetup = null;
        this.dataValidations = null;
        var cfg = (config != null ? config : {});
        this.name = cfg.name;
        this.id = Util._uniqueId("Worksheet");
        this._timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
        if (cfg.columns) {
            this.setColumns(cfg.columns);
        }
        this.relations = new RelationshipManager();
    }
    /** Returns an object that can be consumed by a WorksheetExportWorker
     * @returns export ready data object containing this worksheet's data
     */
    Worksheet.prototype.exportData = function () {
        return {
            relations: this.relations.exportData(),
            columnFormats: this.columnFormats,
            data: this.data,
            columns: this.columns,
            mergedCells: this.mergedCells,
            _headers: this._headers,
            _footers: this._footers,
            _tables: this._tables,
            name: this.name,
            id: this.id
        };
    };
    /** Imports data - to be used while inside of a WorksheetExportWorker.
     * @param data
     */
    Worksheet.prototype.importData = function (data) {
        this.relations.importData(data.relations);
        delete data.relations;
        Object.assign(this, data);
    };
    Worksheet.prototype.setSharedStringCollection = function (stringCollection) {
        this.sharedStrings = stringCollection;
    };
    Worksheet.prototype.addTable = function (table) {
        this._tables.push(table);
        this.relations.addRelation(table, "table");
    };
    Worksheet.prototype.addDrawings = function (table) {
        this._drawings.push(table);
        this.relations.addRelation(table, "drawingRelationship");
    };
    //public addPagePrintSetup() {
    //    this._printerSettings = { id: Util.uniqueId('PrinterSettings') };
    //    this.relations.addRelation(this._printerSettings, 'printerSettings');
    //}
    Worksheet.prototype.setAutoFilter = function (autoFilter) {
        this.autoFilter = autoFilter;
    };
    Worksheet.prototype.setDataValidations = function (dataValidations) {
        this.dataValidations = dataValidations;
    };
    Worksheet.prototype.setPageSetupAndMargins = function (pageSetup, pageMargins) {
        this.pageSetup = pageSetup;
        this.pageMargins = pageMargins;
    };
    /** Expects an array length of three.
     * @see compilePageDetailPiece()
     * @see <a href='/cookbook/addingHeadersAndFooters.html'>Adding headers and footers to a worksheet</a>
     * @param headers [left, center, right]
     */
    Worksheet.prototype.setHeader = function (headers) {
        if (!Array.isArray(headers)) {
            throw new Error("Invalid argument type - setHeader expects an array of three instructions");
        }
        this._headers = headers;
    };
    /** Expects an array length of three.
     * @see compilePageDetailPiece()
     * @see <a href='/cookbook/addingHeadersAndFooters.html'>Adding headers and footers to a worksheet</a>
     * @param footers [left, center, right]
     */
    Worksheet.prototype.setFooter = function (footers) {
        if (!Array.isArray(footers)) {
            throw new Error("Invalid argument type - setFooter expects an array of three instructions");
        }
        this._footers = footers;
    };
    /** Turns page header/footer details into the proper format for Excel.
     * @param pieces
     * @returns a page details string header/footer string
     */
    Worksheet.prototype.compilePageDetailPackage = function (pieces) {
        pieces = pieces || "";
        return [
            "&L", this.compilePageDetailPiece(pieces[0] || ""),
            "&C", this.compilePageDetailPiece(pieces[1] || ""),
            "&R", this.compilePageDetailPiece(pieces[2] || "")
        ].join('');
    };
    /** Turns instructions on page header/footer details into something usable by Excel.
     * @param piece
     * @returns string | reduce
     */
    Worksheet.prototype.compilePageDetailPiece = function (piece) {
        if (typeof piece === "string") {
            return '&"-,Regular"'.concat(piece);
        }
        else if (Array.isArray(piece)) {
            var self = this;
            return piece.reduce(function (m, v) {
                return m.concat(self.compilePageDetailPiece(v));
            }, "");
        }
        else if (typeof piece === "object") {
            var str = "";
            if (piece.font || piece.bold) {
                var weighting = piece.bold ? "Bold" : "Regular";
                str += '&"' + (piece.font || '-');
                str += ',' + weighting + '"';
            }
            else {
                str += '&"-,Regular"';
            }
            if (piece.underline) {
                str += "&U";
            }
            if (piece.fontSize) {
                str += "&" + piece.fontSize;
            }
            str += piece.text;
            return str;
        }
        return undefined;
    };
    /** Creates the header node.
     * @todo implement the ability to do even/odd headers
     * @param doc XmlDom
     * @returns a new XmlDom.XMLNode instance
     */
    Worksheet.prototype.exportHeader = function (doc) {
        var oddHeader = doc.createElement("oddHeader");
        oddHeader.appendChild(doc.createTextNode(this.compilePageDetailPackage(this._headers)));
        return oddHeader;
    };
    /** Creates the footer node.
     * @todo implement the ability to do even/odd footers
     * @param doc XmlDom
     * @returns a new XmlDom.XMLNode instance
     */
    Worksheet.prototype.exportFooter = function (doc) {
        var oddFooter = doc.createElement("oddFooter");
        oddFooter.appendChild(doc.createTextNode(this.compilePageDetailPackage(this._footers)));
        return oddFooter;
    };
    /** This creates some nodes ahead of time, which cuts down on generation time due to
     * most cell definitions being essentially the same, but having multiple nodes that need
     * to be created. Cloning takes less time than creation.
     * @param doc XmlDom
     * @returns
     */
    Worksheet.prototype._buildCache = function (doc) {
        var numberNode = doc.createElement("c");
        var value = doc.createElement("v");
        value.appendChild(doc.createTextNode("--temp--"));
        numberNode.appendChild(value);
        var formulaNode = doc.createElement("c");
        var formulaValue = doc.createElement("f");
        formulaValue.appendChild(doc.createTextNode("--temp--"));
        formulaNode.appendChild(formulaValue);
        var stringNode = doc.createElement("c");
        stringNode.setAttribute("t", "s");
        var stringValue = doc.createElement("v");
        stringValue.appendChild(doc.createTextNode("--temp--"));
        stringNode.appendChild(stringValue);
        return {
            number: numberNode,
            date: numberNode,
            string: stringNode,
            formula: formulaNode
        };
    };
    /** Runs through the XML document and grabs all of the strings that will
     * be sent to the 'shared strings' document.
     * @returns list of cell 'value' strings with 'metadata.type' equal to 'text'
     */
    Worksheet.prototype.collectSharedStrings = function () {
        var data = this.data;
        var maxX = 0;
        var strings = {};
        for (var r = 0, dLen = data.length; r < dLen; r++) {
            var row = data[r];
            var rLen = row.length;
            maxX = rLen > maxX ? rLen : maxX;
            for (var c = 0; c < rLen; c++) {
                var cellValue = row[c];
                if (typeof row[c] === "object") {
                    cellValue = row[c].value;
                }
                var metadata = row[c].metadata || {};
                if (!metadata.type) {
                    if (typeof cellValue === "number") {
                        metadata.type = "number";
                    }
                }
                if (metadata.type == "text" || !metadata.type) {
                    if (typeof strings[cellValue] === "undefined") {
                        strings[cellValue] = true;
                    }
                }
            }
        }
        return Object.keys(strings);
    };
    Worksheet.prototype.toXML = function () {
        var data = this.data;
        var columns = this.columns || [];
        var customCellAttributes = this.customCellAttributes || [];
        var customRowAttributes = this.customRowAttributes || [];
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, "worksheet");
        var worksheet = doc.documentElement;
        worksheet.setAttribute("xmlns:r", Util.schemas.relationships);
        worksheet.setAttribute("xmlns:mc", Util.schemas.markupCompat);
        worksheet.setAttribute("mc:Ignorable", "x14ac");
        worksheet.setAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
        var maxX = 0;
        var sheetData = Util.createElement(doc, "sheetData");
        var cellCache = this._buildCache(doc);
        var sharedStrs = this.sharedStrings;
        for (var r = 0, dLen = data.length; r < dLen; r++) {
            var row = data[r];
            var rLen = row.length;
            maxX = rLen > maxX ? rLen : maxX;
            var rowNode = doc.createElement("row");
            for (var c = 0; c < rLen; c++) {
                columns[c] = columns[c] || {};
                var cellValue = row[c];
                if (cellValue != null && typeof cellValue === "object") {
                    cellValue = cellValue.value;
                }
                //fix undefined or null value
                var metadata = row[c] ? (row[c].metadata || {}) : {};
                if (!metadata.type) {
                    if (typeof cellValue === "number") {
                        metadata.type = "number";
                    }
                    // Allows for empty cells in switch statement below
                    if (cellValue == null) {
                        metadata.type = "empty";
                    }
                }
                switch (metadata.type) {
                    case "number":
                        var cell = cellCache.number.cloneNode(true);
                        cell.firstChild.firstChild.nodeValue = cellValue;
                        break;
                    case "date":
                        var cell = cellCache.date.cloneNode(true);
                        cell.firstChild.firstChild.nodeValue = 25569.0 + ((cellValue - this._timezoneOffset) / (60 * 60 * 24 * 1000));
                        break;
                    case "formula":
                        var cell = cellCache.formula.cloneNode(true);
                        cell.firstChild.firstChild.nodeValue = cellValue;
                        break;
                    // empty cell that contains no value, valid in Excel
                    case "empty":
                        var cell = doc.createElement("c");
                        break;
                    case "text":
                    default:
                        var id = sharedStrs.strings[cellValue] || (sharedStrs.addString && sharedStrs.addString(cellValue));
                        var cell = cellCache.string.cloneNode(true);
                        cell.firstChild.firstChild.nodeValue = id;
                        break;
                }
                ;
                if (metadata.style) {
                    cell.setAttribute("s", metadata.style);
                }
                cell.setAttribute("r", Util.positionToLetterRef(c + 1, r + 1));
                // add any additional custom attributes to this cell's XML element
                if (r < customCellAttributes.length && customCellAttributes[r] != null && c < customCellAttributes[r].length) {
                    var attribs = customCellAttributes[r][c];
                    for (var attrib in attribs) {
                        cell.setAttribute(attrib, attribs[attrib]);
                    }
                }
                rowNode.appendChild(cell);
            }
            rowNode.setAttribute("r", r + 1);
            // add any additional custom attributes to this row's XML element
            if (r < customRowAttributes.length) {
                var rowAttribs = customRowAttributes[r];
                for (var attrib in rowAttribs) {
                    rowNode.setAttribute(attrib, rowAttribs[attrib]);
                }
            }
            sheetData.appendChild(rowNode);
        }
        var cellOrRangeRef = (maxX !== 0
            ? Util.positionToLetterRef(1, 1) + ':' + Util.positionToLetterRef(maxX, data.length)
            : Util.positionToLetterRef(1, 1));
        worksheet.appendChild(Util.createElement(doc, "dimension", [
            ["ref", cellOrRangeRef]
        ]));
        if (this.columns.length) {
            worksheet.appendChild(this.exportColumns(doc));
        }
        worksheet.appendChild(sheetData);
        this.exportPageSettings(doc, worksheet);
        if (this._headers.length > 0 || this._footers.length > 0) {
            var headerFooter = doc.createElement("headerFooter");
            if (this._headers.length > 0) {
                headerFooter.appendChild(this.exportHeader(doc));
            }
            if (this._footers.length > 0) {
                headerFooter.appendChild(this.exportFooter(doc));
            }
            worksheet.appendChild(headerFooter);
        }
        if (this._tables.length > 0) {
            var tables = doc.createElement("tableParts");
            tables.setAttribute("count", this._tables.length);
            for (var i = 0, tLen = this._tables.length; i < tLen; i++) {
                var table = doc.createElement("tablePart");
                table.setAttribute("r:id", this.relations.getRelationshipId(this._tables[i]));
                tables.appendChild(table);
            }
            worksheet.appendChild(tables);
        }
        if (this.mergedCells.length > 0) {
            var mergeCells = doc.createElement("mergeCells");
            for (var i = 0, mcLen = this.mergedCells.length; i < mcLen; i++) {
                var mergeCell = doc.createElement("mergeCell");
                mergeCell.setAttribute("ref", this.mergedCells[i][0] + ':' + this.mergedCells[i][1]);
                mergeCells.appendChild(mergeCell);
            }
            mergeCells.setAttribute("count", this.mergedCells.length);
            worksheet.appendChild(mergeCells);
        }
        // Add autoFilter element if there are custom autofilter attributes
        if (this.autoFilter) {
            var autoFilterEl = doc.createElement("autoFilter");
            autoFilterEl.setAttribute("ref", (this.autoFilter.left + this.autoFilter.rowNum + ":" + this.autoFilter.right + this.autoFilter.rowNum));
            worksheet.appendChild(autoFilterEl);
        }
        // Add pageMargins element if there are custom page margin attributes
        if (this.pageMargins) {
            var pageMarginsEl = doc.createElement("pageMargins");
            for (var attr in this.pageMargins) {
                pageMarginsEl.setAttribute(attr, this.pageMargins[attr]);
            }
            worksheet.appendChild(pageMarginsEl);
        }
        // Add pageSetup element if there are custom page setup/printing attributes
        if (this.pageSetup) {
            var pageSetupEl = doc.createElement("pageSetup");
            //pageSetupEl.setAttribute("r:id", this.relations.getRelationshipId(this._printerSettings));
            for (var attr in this.pageSetup) {
                pageSetupEl.setAttribute(attr, this.pageSetup[attr]);
            }
            worksheet.appendChild(pageSetupEl);
        }
        for (var i = 0, dLen = this._drawings.length; i < dLen; i++) {
            var drawing = doc.createElement("drawing");
            drawing.setAttribute("r:id", this.relations.getRelationshipId(this._drawings[i]));
            worksheet.appendChild(drawing);
        }
        // Add extLst element if there are custom dataValidations attributes
        if (this.dataValidations) {
            var extLstEl = doc.createElement("extLst");
            var extChildEl = doc.createElement("ext");
            extChildEl.setAttribute("xmlns:x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            extChildEl.setAttribute("uri", "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}");
            var dataValidationsEl = doc.createElement("x14:dataValidations");
            dataValidationsEl.setAttribute("xmlns:xm", "http://schemas.microsoft.com/office/excel/2006/main");
            dataValidationsEl.setAttribute("count", this.dataValidations.length.toString());
            for (var i = 0; i < this.dataValidations.length; i++) {
                var dataValadation = this.dataValidations[i];
                var dataValidationEl = doc.createElement("x14:dataValidation");
                dataValidationEl.setAttribute("showErrorMessage", "1");
                dataValidationEl.setAttribute("showInputMessage", "1");
                dataValidationEl.setAttribute("allowBlank", "1");
                dataValidationEl.setAttribute("type", "list");
                var formulaEl = doc.createElement("x14:formula1");
                var xmfEl = doc.createElement("xm:f");
                // Build formula to specify range where data set lies. Ex: SheetName!$A$5:$A$10
                xmfEl.appendChild(doc.createTextNode(dataValadation.dataSheetName + "!$" + dataValadation.column + "$" + dataValadation.topRow + ":$" + dataValadation.column + "$" + dataValadation.bottomRow));
                var sqrefEl = doc.createElement("xm:sqref");
                sqrefEl.appendChild(doc.createTextNode(dataValadation.sqref));
                formulaEl.appendChild(xmfEl);
                dataValidationEl.appendChild(formulaEl);
                dataValidationEl.appendChild(sqrefEl);
                dataValidationsEl.appendChild(dataValidationEl);
            }
            extChildEl.appendChild(dataValidationsEl);
            extLstEl.appendChild(extChildEl);
            worksheet.appendChild(extLstEl);
        }
        return doc;
    };
    /**
     * @param doc XmlDom
     * @returns a new XmlDom.XMLNode instance
     */
    Worksheet.prototype.exportColumns = function (doc) {
        var cols = Util.createElement(doc, "cols");
        for (var i = 0, l = this.columns.length; i < l; i++) {
            var cd = this.columns[i];
            var col = Util.createElement(doc, "col", [
                ["min", cd.min || i + 1],
                ["max", cd.max || i + 1]
            ]);
            if (cd.hidden) {
                col.setAttribute("hidden", 1);
            }
            if (cd.bestFit) {
                col.setAttribute("bestFit", 1);
            }
            if (cd.customWidth || cd.width) {
                col.setAttribute("customWidth", 1);
            }
            if (cd.width) {
                col.setAttribute("width", cd.width);
            }
            else {
                col.setAttribute("width", 9.140625);
            }
            cols.appendChild(col);
        }
        ;
        return cols;
    };
    /** Sets the page settings on a worksheet node.
     * @param doc XmlDom
     * @param worksheet XmlDom.XMLNode
     */
    Worksheet.prototype.exportPageSettings = function (doc, worksheet) {
        if (this._margin) {
            var defaultVal = 0.7;
            var left = this._margin.left ? this._margin.left : defaultVal;
            var right = this._margin.right ? this._margin.right : defaultVal;
            var top = this._margin.top ? this._margin.top : defaultVal;
            var bottom = this._margin.bottom ? this._margin.bottom : defaultVal;
            defaultVal = 0.3;
            var header = this._margin.header ? this._margin.header : defaultVal;
            var footer = this._margin.footer ? this._margin.footer : defaultVal;
            worksheet.appendChild(Util.createElement(doc, "pageMargins", [
                ["top", top],
                ["bottom", bottom],
                ["left", left],
                ["right", right],
                ["header", header],
                ["footer", footer],
            ]));
        }
        if (this._orientation) {
            worksheet.appendChild(Util.createElement(doc, "pageSetup", [
                ["orientation", this._orientation]
            ]));
        }
    };
    /** Set page details in inches.
     * use this structure:
     *  top: 0.7,
     *  bottom: 0.7,
     *  left: 0.7,
     *  right: 0.7,
     *  header: 0.3,
     *  footer: 0.3,
     */
    Worksheet.prototype.setPageMargin = function (input) {
        this._margin = input;
    };
    /** Can be one of 'portrait' or 'landscape'.
     * http://www.schemacentral.com/sc/ooxml/t-ssml_ST_Orientation.html
     * @param orientation
     */
    Worksheet.prototype.setPageOrientation = function (orientation) {
        this._orientation = orientation;
    };
    /** Expects an array of column definitions. Each column definition needs to have a width assigned to it.
     * @param columns
     */
    Worksheet.prototype.setColumns = function (columns) {
        this.columns = columns;
    };
    /** Expects an array of data to be translated into cells.
     * @param data Two dimensional array - [ [A1, A2], [B1, B2] ]
     * @see <a href='/cookbook/addingDataToAWorksheet.html'>Adding data to a worksheet</a>
     */
    Worksheet.prototype.setData = function (data) {
        this.data = data;
    };
    /** Merge cells in given range
     * @param cell1 - A1, A2...
     * @param cell2 - A2, A3...
     */
    Worksheet.prototype.mergeCells = function (cell1, cell2) {
        this.mergedCells.push([cell1, cell2]);
    };
    /** Expects an array containing an object full of column format definitions.
     * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
     * bestFit
     * collapsed
     * customWidth
     * hidden
     * max
     * min
     * outlineLevel
     * phonetic
     * style
     * width
     */
    Worksheet.prototype.setColumnFormats = function (columnFormats) {
        this.columnFormats = columnFormats;
    };
    return Worksheet;
}());
module.exports = Worksheet;
