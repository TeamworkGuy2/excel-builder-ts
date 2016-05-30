import Util = require("./Util");
import RelationshipManager = require("./RelationshipManager");
import XmlDom = require("./XmlDom");


/**
 * This module represents an excel worksheet in its basic form - no tables, charts, etc. Its purpose is 
 * to hold data, the data's link to how it should be styled, and any links to other outside resources.
 * 
 * @module Excel/Worksheet
 */
class Worksheet {
    relations: RelationshipManager;
    columnFormats: Worksheet.ColumnFormat[];
    data: any[];
    mergedCells: any[];
    columns: Worksheet.Column[];
    _headers: [string, string, string]; // [left, center, right]
    _footers: [string, string, string]; // [left, center, right]
    _tables: { id: string; schema?: string; }[];
    _drawings: any[];

    id: string;
    name: string;
    sharedStrings: { strings: { [key: string]: number }; addString?(str: string): number; };
    _timezoneOffset: number;
    // The page orientation
    _orientation: string;

    // custom-code 2014-6-27
    // A two dimensional array of objects with custom XML attributes to add this worksheet's cells
    // for example an object { style: 12b } at index [1][2] would add a {@code style="12b"} attribute to cell 'C2'
    customCellAttributes: any[];
    // A one dimensional array with the same purpose as the custom cell attributes array except the custom
    // attributes are applied to the worksheet's rows
    customRowAttributes: any[];
    // The ID and settings for pageMargins and pageSetup
    _printerSettings: any;
    // An array of attributes to apply to the 'pageMargins' element of the spreadsheet
    pageMargins: any;
    // An array of attributes to apply to the 'pageSetup' element of the spreadsheet
    pageSetup: any;
    // custom-code-end


    /**
     * @constructor
     */
    constructor(config?: { name: string; columns: Worksheet.Column[]; }) {
        this.relations = null;
        this.columnFormats = [];
        this.data = [];
        this.mergedCells = [];
        this.columns = [];
        this._headers = <any>[];
        this._footers = <any>[];
        this._tables = [];
        this._drawings = [];

        // custom-code 2014-6-27
        // A two dimensional array of objects with custom XML attributes to add this worksheet's cells
        // for example an object { style: 12b } at index [1][2] would add a {@code style="12b"} attribute to cell 'C2'
        this.customCellAttributes = [];
        // A one dimensional array with the same purpose as the custom cell attributes array except the custom
        // attributes are applied to the worksheet's rows
        this.customRowAttributes = [];
        // The ID and settings for pageMargins and pageSetup
        this._printerSettings;
        // An array of attributes to apply to the 'pageMargins' element of the spreadsheet
        this.pageMargins;
        // An array of attributes to apply to the 'pageSetup' element of the spreadsheet
        this.pageSetup;
        // custom-code-end

        this.initialize(config);
    }


    private initialize(config?: { name: string; columns: Worksheet.Column[]; }) {
        config = config || <any>{};
        this.name = config.name;
        this.id = Util._uniqueId('Worksheet');
        this._timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
        if (config.columns) {
            this.setColumns(config.columns);
        }

        this.relations = new RelationshipManager();
    }


    /** Returns an object that can be consumed by a WorksheetExportWorker
     * @returns {Object}
     */
    public exportData() {
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
    }


    /** Imports data - to be used while inside of a WorksheetExportWorker.
     * @param {Object} data
     */
    public importData(data: { relations: any }) {
        this.relations.importData(data.relations);
        delete data.relations;
        Object.assign(this, data);
    }


    public setSharedStringCollection(stringCollection: { strings: { [key: string]: number }; addString?(str: string): number; }) {
        this.sharedStrings = stringCollection;
    }


    public addTable(table: { id: string }) {
        this._tables.push(table);
        this.relations.addRelation(table, 'table');
    }


    public addDrawings(table: { id: string }) {
        this._drawings.push(table);
        this.relations.addRelation(table, 'drawingRelationship');
    }


    // not-original 2015-5-1
    public addPagePrintSetup(pageSetup: any, pageMargins: any) {
        this.pageSetup = pageSetup;
        this.pageMargins = pageMargins;
        //this._printerSettings = { id: _.uniqueId('PrinterSettings') };
        //this.relations.addRelation(this._printerSettings, 'printerSettings');
    }
    // not-original-end


    /** Expects an array length of three.
     * 
     * @see Excel/Worksheet compilePageDetailPiece 
     * @see <a href='/cookbook/addingHeadersAndFooters.html'>Adding headers and footers to a worksheet</a>
     * 
     * @param {Array} headers [left, center, right]
     */
    public setHeader(headers: [string, string, string]) {
        if (!Array.isArray(headers)) {
            throw "Invalid argument type - setHeader expects an array of three instructions";
        }
        this._headers = headers;
    }


    /**
     * Expects an array length of three.
     * 
     * @see Excel/Worksheet compilePageDetailPiece 
     * @see <a href='/cookbook/addingHeadersAndFooters.html'>Adding headers and footers to a worksheet</a>
     * 
     * @param {Array} footers [left, center, right]
     */
    public setFooter(footers: [string, string, string]) {
        if (!Array.isArray(footers)) {
            throw "Invalid argument type - setFooter expects an array of three instructions";
        }
        this._footers = footers;
    }


    /** Turns page header/footer details into the proper format for Excel.
     * @param {type} data
     * @returns {string}
     */
    public compilePageDetailPackage(data: [string, string, string]) {
        data = data || <any>"";
        return [
            "&L", this.compilePageDetailPiece(data[0] || ""),
            "&C", this.compilePageDetailPiece(data[1] || ""),
            "&R", this.compilePageDetailPiece(data[2] || "")
        ].join('');
    }


    /** Turns instructions on page header/footer details into something
     * usable by Excel.
     * 
     * @param {type} data
     * @returns {string|@exp;_@call;reduce}
     */
    public compilePageDetailPiece(data: string | { font?; bold?; underline?; fontSize?; text?; [id: string]: any } | any[]) {
        if (typeof data === "string") {
            return '&"-,Regular"'.concat(data);
        }
        if (typeof data === "[object Object]" && !Array.isArray(data)) {
            var string = "";
            if (data.font || data.bold) {
                var weighting = data.bold ? "Bold" : "Regular";
                string += '&"' + (data.font || '-');
                string += ',' + weighting + '"';
            } else {
                string += '&"-,Regular"';
            }
            if (data.underline) {
                string += "&U";
            }
            if (data.fontSize) {
                string += "&" + data.fontSize;
            }
            string += data.text;

            return string;
        }

        if (Array.isArray(data)) {
            var self = this;
            return data.reduce(function (m, v) {
                return m.concat(self.compilePageDetailPiece(v));
            }, "");
        }
    }


    /** Creates the header node. 
     * 
     * @todo implement the ability to do even/odd headers
     * @param {XML Doc} doc
     * @returns {XML Node}
     */
    public exportHeader(doc: XmlDom) {
        var oddHeader = doc.createElement('oddHeader');
        oddHeader.appendChild(doc.createTextNode(this.compilePageDetailPackage(this._headers)));
        return oddHeader;
    }


    /** Creates the footer node.
     * 
     * @todo implement the ability to do even/odd footers
     * @param {XML Doc} doc
     * @returns {XML Node}
     */
    public exportFooter(doc: XmlDom) {
        var oddFooter = doc.createElement('oddFooter');
        oddFooter.appendChild(doc.createTextNode(this.compilePageDetailPackage(this._footers)));
        return oddFooter;
    }


    /** This creates some nodes ahead of time, which cuts down on generation time due to 
     * most cell definitions being essentially the same, but having multiple nodes that need
     * to be created. Cloning takes less time than creation.
     * 
     * @private
     * @param {XML Doc} doc
     * @returns
     */
    private _buildCache(doc: XmlDom) {
        var numberNode = doc.createElement('c');
        var value = doc.createElement('v');
        value.appendChild(doc.createTextNode("--temp--"));
        numberNode.appendChild(value);

        var formulaNode = doc.createElement('c');
        var formulaValue = doc.createElement('f');
        formulaValue.appendChild(doc.createTextNode("--temp--"));
        formulaNode.appendChild(formulaValue);

        var stringNode = doc.createElement('c');
        stringNode.setAttribute('t', 's');
        var stringValue = doc.createElement('v');
        stringValue.appendChild(doc.createTextNode("--temp--"));
        stringNode.appendChild(stringValue);


        return {
            number: numberNode,
            date: numberNode,
            string: stringNode,
            formula: formulaNode
        }
    }


    /** Runs through the XML document and grabs all of the strings that will
     * be sent to the 'shared strings' document. 
     * 
     * @returns {Array}
     */
    public collectSharedStrings() {
        var data = this.data;
        var maxX = 0;
        var strings: { [id: string]: boolean; } = {};
        for (var row = 0, l = data.length; row < l; row++) {
            var dataRow = data[row];
            var cellCount = dataRow.length;
            maxX = cellCount > maxX ? cellCount : maxX;
            for (var c = 0; c < cellCount; c++) {
                var cellValue = dataRow[c];
                if (typeof dataRow[c] == 'object') {
                    cellValue = dataRow[c].value;
                }
                var metadata = dataRow[c].metadata || {};

                if (!metadata.type) {
                    if (typeof cellValue == 'number') {
                        metadata.type = 'number';
                    }
                }
                if (metadata.type == "text" || !metadata.type) {
                    if (typeof strings[cellValue] == 'undefined') {
                        strings[cellValue] = true;
                    }
                }
            }
        }
        return Object.keys(strings);
    }


    public toXML() {
        var data = this.data;
        var columns = this.columns || [];

        // custom-code 2014-6-27
        var customCellAttributes = this.customCellAttributes || [];
        var customRowAttributes = this.customRowAttributes || [];
        // custom-code-end

        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, 'worksheet');
        var worksheet = doc.documentElement;
        worksheet.setAttribute('xmlns:r', Util.schemas.relationships);
        worksheet.setAttribute('xmlns:mc', Util.schemas.markupCompat);

        var maxX = 0;
        var sheetData = Util.createElement(doc, 'sheetData');

        var cellCache = this._buildCache(doc);

        for (var row = 0, l = data.length; row < l; row++) {
            var dataRow = data[row];
            var cellCount = dataRow.length;
            maxX = cellCount > maxX ? cellCount : maxX;
            var rowNode = doc.createElement('row');

            for (var c = 0; c < cellCount; c++) {
                columns[c] = columns[c] || {};
                var cellValue = dataRow[c];
                if (dataRow[c] != null && typeof dataRow[c] == 'object') {
                    cellValue = dataRow[c].value;
                }
                //fix undefined or null value
                var cell: XmlDom.XMLNode,
                    metadata = dataRow[c] ? (dataRow[c].metadata || {}) : {};


                if (!metadata.type) {
                    if (typeof cellValue == 'number') {
                        metadata.type = 'number';
                    }
                    // custom-code 2014-6-30
                    // Allows for empty cells in switch statement below
                    if (cellValue === null || typeof cellValue === undefined) {
                        metadata.type = "empty";
                    }
                    // custom-code-end
                }

                switch (metadata.type) {
                    case "number":
                        cell = cellCache.number.cloneNode(true);
                        (<XmlDom.XMLNode>cell.firstChild).firstChild.nodeValue = cellValue;
                        break;
                    case "date":
                        cell = cellCache.date.cloneNode(true);
                        (<XmlDom.XMLNode>cell.firstChild).firstChild.nodeValue = <any>25569.0 + ((cellValue - this._timezoneOffset) / (60 * 60 * 24 * 1000));
                        break;
                    case "formula":
                        cell = cellCache.formula.cloneNode(true);
                        (<XmlDom.XMLNode>cell.firstChild).firstChild.nodeValue = cellValue;
                        break;
                    // custom-code 2014-6-30
                    // empty cell that contains no value, valid in Excel
                    case "empty":
                        cell = doc.createElement("c");
                        break;
                    // custom-code-end
                    case "text":
                    default:
                        var id: number;
                        if (typeof this.sharedStrings.strings[cellValue] != 'undefined') {
                            id = this.sharedStrings.strings[cellValue];
                        } else {
                            id = this.sharedStrings.addString(cellValue);
                        }
                        cell = cellCache.string.cloneNode(true);
                        (<XmlDom.XMLNode>cell.firstChild).firstChild.nodeValue = <any>id;
                        break;
                };
                if (metadata.style) {
                    cell.setAttribute('s', metadata.style);
                }
                cell.setAttribute('r', Util.positionToLetterRef(c + 1, row + 1));

                // custom-code 2014-6-27
                // add any additional custom attributes to this cell's XML element
                if (row < customCellAttributes.length && customCellAttributes[row] !== null && c < customCellAttributes[row].length) {
                    var attribs = customCellAttributes[row][c];
                    for (var attrib in attribs) {
                        cell.setAttribute(attrib, attribs[attrib]);
                    }
                }
                // custom-code-end

                rowNode.appendChild(cell);
            }
            rowNode.setAttribute('r', row + 1);

            // custom-code 2014-6-27
            // add any additional custom attributes to this row's XML element
            if (row < customRowAttributes.length) {
                var rowAttribs = customRowAttributes[row];
                for (var attrib in rowAttribs) {
                    rowNode.setAttribute(attrib, rowAttribs[attrib]);
                }
            }
            // custom-code-end

            sheetData.appendChild(rowNode);
        }

        if (maxX !== 0) {
            worksheet.appendChild(Util.createElement(doc, 'dimension', [
                ['ref', Util.positionToLetterRef(1, 1) + ':' + Util.positionToLetterRef(maxX, data.length)]
            ]));
        } else {
            worksheet.appendChild(Util.createElement(doc, 'dimension', [
                ['ref', Util.positionToLetterRef(1, 1)]
            ]));
        }

        if (this.columns.length) {
            worksheet.appendChild(this.exportColumns(doc));
        }
        worksheet.appendChild(sheetData);

        this.exportPageSettings(doc, worksheet);

        if (this._headers.length > 0 || this._footers.length > 0) {
            var headerFooter = doc.createElement('headerFooter');
            if (this._headers.length > 0) {
                headerFooter.appendChild(this.exportHeader(doc));
            }
            if (this._footers.length > 0) {
                headerFooter.appendChild(this.exportFooter(doc));
            }
            worksheet.appendChild(headerFooter);
        }

        if (this._tables.length > 0) {
            var tables = doc.createElement('tableParts');
            tables.setAttribute('count', this._tables.length);
            for (var i = 0, l = this._tables.length; i < l; i++) {
                var table = doc.createElement('tablePart');
                table.setAttribute('r:id', this.relations.getRelationshipId(this._tables[i]));
                tables.appendChild(table);
            }
            worksheet.appendChild(tables);
        }

        if (this.mergedCells.length > 0) {
            var mergeCells = doc.createElement('mergeCells');
            for (var i = 0, l = this.mergedCells.length; i < l; i++) {
                var mergeCell = doc.createElement('mergeCell');
                mergeCell.setAttribute('ref', this.mergedCells[i][0] + ':' + this.mergedCells[i][1]);
                mergeCells.appendChild(mergeCell);
            }
            // custom-code 2014-7-2
            mergeCells.setAttribute("count", this.mergedCells.length);
            // custom-code-end
            worksheet.appendChild(mergeCells);
        }

        // custom-code 2014-6-30
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
        // custom-code-end

        for (var i = 0, l = this._drawings.length; i < l; i++) {
            var drawing = doc.createElement('drawing');
            drawing.setAttribute('r:id', this.relations.getRelationshipId(this._drawings[i]));
            worksheet.appendChild(drawing);
        }

        return doc;
    }


    /**
     * @param {XML Doc} doc
     * @returns {XML Node}
     */
    public exportColumns(doc: XmlDom) {
        var cols = Util.createElement(doc, 'cols');
        for (var i = 0, l = this.columns.length; i < l; i++) {
            var cd = this.columns[i];
            var col = Util.createElement(doc, 'col', [
                ['min', cd.min || i + 1],
                ['max', cd.max || i + 1]
            ]);
            if (cd.hidden) {
                col.setAttribute('hidden', <any>1);
            }
            if (cd.bestFit) {
                col.setAttribute('bestFit', <any>1);
            }
            if (cd.customWidth || cd.width) {
                col.setAttribute('customWidth', <any>1);
            }
            if (cd.width) {
                col.setAttribute('width', cd.width);
            } else {
                col.setAttribute('width', <any>9.140625);
            }

            cols.appendChild(col)
        };
        return cols;
    }


    /** Sets the page settings on a worksheet node.
     * 
     * @param {XML Doc} doc
     * @param {XML Node} worksheet
     * @returns {undefined}
     */
    public exportPageSettings(doc: XmlDom, worksheet: XmlDom.XMLNode) {
        if (this._orientation) {
            worksheet.appendChild(Util.createElement(doc, "pageSetup", [
                ['orientation', this._orientation]
            ]));
        }
    }


    /** Can be one of 'portrait' or 'landscape'.
     * http://www.schemacentral.com/sc/ooxml/t-ssml_ST_Orientation.html
     * 
     * @param {String} orientation
     * @returns {undefined}
     */
    public setPageOrientation(orientation: string) {
        this._orientation = orientation;
    }


    /** Expects an array of column definitions. Each column definition needs to have a width assigned to it. 
     * 
     * @param {Array} Columns
     */
    public setColumns(columns: Worksheet.Column[]) {
        this.columns = columns;
    }


    /** Expects an array of data to be translated into cells. 
     * 
     * @param {Array} data Two dimensional array - [ [A1, A2], [B1, B2] ]
     * @see <a href='/cookbook/addingDataToAWorksheet.html'>Adding data to a worksheet</a>
     */
    public setData(data: any[]) {
        this.data = data;
    }


    /** Merge cells in given range
     *
     * @param cell1 - A1, A2...
     * @param cell2 - A2, A3...
     */
    public mergeCells(cell1, cell2) {
        this.mergedCells.push([cell1, cell2]);
    }


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
    public setColumnFormats(columnFormats: Worksheet.ColumnFormat[]) {
        this.columnFormats = columnFormats;
    }

}

module Worksheet {

    export interface Column {
        min?: number;
        max?: number;
        hidden?: boolean;
        bestFit?: boolean;
        customWidth?: boolean;
        width?: any;
    }


    export interface ColumnFormat {
        bestFit: boolean;
        collapsed: boolean;
        customWidth: boolean;
        hidden: boolean;
        max: number;
        min: number;
        outlineLevel: number;
        phonetic: boolean;
        style: number;
        width: number;
    }

}

export = Worksheet;