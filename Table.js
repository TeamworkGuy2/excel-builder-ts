"use strict";
var Util = require("./Util");
/**
 * @module Excel/Table
 */
var Table = (function () {
    function Table(config) {
        var _this = this;
        var defaults = {
            autoFilter: null,
            dataCellStyle: null,
            dataDfxId: null,
            displayName: "",
            headerRowBorderDxfId: null,
            headerRowCellStyle: null,
            headerRowCount: 1,
            headerRowDxfId: null,
            insertRow: false,
            insertRowShift: false,
            name: "",
            ref: null,
            sortState: null,
            styleInfo: {},
            tableBorderDxfId: null,
            totalsRowBorderDxfId: null,
            totalsRowCellStyle: null,
            totalsRowCount: 0,
            totalsRowDxfId: null,
            tableColumns: [],
        };
        Object.keys(defaults).forEach(function (key) {
            if (_this[key] == null) {
                _this[key] = defaults[key];
            }
        });
        this.initialize(config);
    }
    Table.prototype.initialize = function (config) {
        this.displayName = Util._uniqueId("Table");
        this.name = this.displayName;
        this.id = this.name;
        this.tableId = this.id.replace("Table", '');
        if (config != null) {
            Object.assign(this, config);
        }
    };
    Table.prototype.setReferenceRange = function (start, end) {
        this.ref = [start, end];
    };
    Table.prototype.setTableColumns = function (columns) {
        var _this = this;
        columns.forEach(function (column) {
            _this.addTableColumn(column);
        });
    };
    /** Expects an object with the following optional properties:
     * name (required)
     * dataCellStyle
     * dataDxfId
     * headerRowCellStyle
     * headerRowDxfId
     * totalsRowCellStyle
     * totalsRowDxfId
     * totalsRowFunction
     * totalsRowLabel
     * columnFormula
     * columnFormulaIsArrayType (boolean)
     * totalFormula
     * totalFormulaIsArrayType (boolean)
     */
    Table.prototype.addTableColumn = function (column) {
        var col = column;
        if (typeof column === "string") {
            col = {
                name: column
            };
        }
        if (!col.name) {
            throw new Error("Invalid argument for addTableColumn - minimum requirement is a name property");
        }
        this.tableColumns.push(col);
    };
    /** Expects an object with the following properties:
     * caseSensitive (boolean)
     * dataRange
     * columnSort (assumes true)
     * sortDirection
     * sortRange (defaults to dataRange)
     */
    Table.prototype.setSortState = function (state) {
        this.sortState = state;
    };
    Table.prototype.toXML = function () {
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, "table");
        var table = doc.documentElement;
        table.setAttribute("id", this.tableId);
        table.setAttribute("name", this.name);
        table.setAttribute("displayName", this.displayName);
        var s = this.ref[0];
        var e = this.ref[1];
        table.setAttribute("ref", Util.positionToLetterRef(s[0], s[1]) + ":" + Util.positionToLetterRef(e[0], e[1]));
        /** TOTALS **/
        table.setAttribute("totalsRowCount", this.totalsRowCount);
        /** HEADER **/
        table.setAttribute("headerRowCount", this.headerRowCount);
        if (this.headerRowDxfId) {
            table.setAttribute("headerRowDxfId", this.headerRowDxfId);
        }
        if (this.headerRowBorderDxfId) {
            table.setAttribute("headerRowBorderDxfId", this.headerRowBorderDxfId);
        }
        if (!this.ref) {
            throw new Error("Needs at least a reference range");
        }
        if (!this.autoFilter) {
            this.addAutoFilter(this.ref[0], this.ref[1]);
        }
        table.appendChild(this.exportAutoFilter(doc));
        table.appendChild(this.exportTableColumns(doc));
        table.appendChild(this.exportTableStyleInfo(doc));
        return table;
    };
    Table.prototype.exportTableColumns = function (doc) {
        var tableColumns = doc.createElement("tableColumns");
        tableColumns.setAttribute("count", this.tableColumns.length);
        var tcs = this.tableColumns;
        for (var i = 0, l = tcs.length; i < l; i++) {
            var col = tcs[i];
            var tableColumn = doc.createElement("tableColumn");
            tableColumn.setAttribute("id", i + 1);
            tableColumn.setAttribute("name", col.name);
            if (col.totalsRowFunction) {
                tableColumn.setAttribute("totalsRowFunction", col.totalsRowFunction);
            }
            if (col.totalsRowLabel) {
                tableColumn.setAttribute("totalsRowLabel", col.totalsRowLabel);
            }
            tableColumns.appendChild(tableColumn);
        }
        return tableColumns;
    };
    Table.prototype.exportAutoFilter = function (doc) {
        var autoFilter = doc.createElement("autoFilter");
        var s = this.autoFilter[0];
        var e = this.autoFilter[1];
        autoFilter.setAttribute("ref", Util.positionToLetterRef(s[0], s[1]) + ":" + Util.positionToLetterRef(e[0], e[1] - this.totalsRowCount));
        return autoFilter;
    };
    Table.prototype.exportTableStyleInfo = function (doc) {
        var ts = this.styleInfo;
        var tableStyle = doc.createElement("tableStyleInfo");
        tableStyle.setAttribute("name", ts.themeStyle);
        tableStyle.setAttribute("showFirstColumn", ts.showFirstColumn ? "1" : "0");
        tableStyle.setAttribute("showLastColumn", ts.showLastColumn ? "1" : "0");
        tableStyle.setAttribute("showColumnStripes", ts.showColumnStripes ? "1" : "0");
        tableStyle.setAttribute("showRowStripes", ts.showRowStripes ? "1" : "0");
        return tableStyle;
    };
    Table.prototype.addAutoFilter = function (startRef, endRef) {
        this.autoFilter = [startRef, endRef];
    };
    return Table;
}());
module.exports = Table;
