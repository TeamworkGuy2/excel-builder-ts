import Util = require("../util/Util");
import XmlDom = require("../xml/XmlDom");


interface SortState {
    caseSensitive: boolean;
    dataRange: any;
    columnSort: boolean; //(assumes true);
    sortDirection: any;
    sortRange: any; //(defaults to dataRange)
}


/**
 * @module Excel/Table
 */
class Table {
    autoFilter: [[number, number], [number, number]];
    displayName: string;
    headerRowBorderDxfId: string;
    headerRowCount: number;
    headerRowDxfId: string | number;
    id: string;
    name: string;
    ref: [[number, number], [number, number]];
    sortState: SortState;
    styleInfo: any;
    tableId: string;
    totalsRowCount: number;
    tableColumns: {
        name: string;
        [prop: string]: any;
    }[];


    constructor(config?: any) {
        var defaults = {
            autoFilter: <[[number, number], [number, number]]>null,
            dataCellStyle: <any>null,
            dataDfxId: <any>null,
            displayName: "",
            headerRowBorderDxfId: <string>null,
            headerRowCellStyle: <any>null,
            headerRowCount: 1,
            headerRowDxfId: <string | number>null,
            insertRow: false,
            insertRowShift: false,
            name: "",
            ref: <[[number, number], [number, number]]>null,
            sortState: <SortState>null,
            styleInfo: {},
            tableBorderDxfId: <any>null,
            totalsRowBorderDxfId: <any>null,
            totalsRowCellStyle: <any>null,
            totalsRowCount: 0,
            totalsRowDxfId: <any>null,
            tableColumns: <{ name: string; [prop: string]: any; }[]>[],
        };
        (<(keyof typeof defaults)[]>Object.keys(defaults)).forEach((key) => {
            if ((<any>this)[key] == null) {
                (<any>this)[key] = defaults[key];
            }
        });
        this.initialize(config);
    }


    public initialize(config?: any) {
        this.displayName = Util._uniqueId("Table");
        this.name = this.displayName;
        this.id = this.name;
        this.tableId = this.id.replace("Table", '');
        if (config != null) {
            Object.assign(this, config);
        }
    }


    public setReferenceRange(start: [number, number], end: [number, number]) {
        this.ref = [start, end];
    }


    public setTableColumns(columns: any[]) {
        columns.forEach((column) => {
            this.addTableColumn(column);
        });
    }


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
    public addTableColumn(column: string | { name: string; }) {
        var col: { name: string; } = <any>column;
        if (typeof column === "string") {
            col = {
                name: column
            };
        }
        if (!col.name) {
            throw new Error("Invalid argument for addTableColumn - minimum requirement is a name property");
        }
        this.tableColumns.push(col);
    }


    /** Expects an object with the following properties:
     * caseSensitive (boolean)
     * dataRange
     * columnSort (assumes true)
     * sortDirection
     * sortRange (defaults to dataRange)
     */
    public setSortState(state: SortState) {
        this.sortState = state;
    }


    public toXML() {
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
    }


    public exportTableColumns(doc: XmlDom) {
        var tableCols = doc.createElement("tableColumns");
        tableCols.setAttribute("count", this.tableColumns.length);
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
            tableCols.appendChild(tableColumn);
        }
        return tableCols;
    }


    public exportAutoFilter(doc: XmlDom) {
        var autoFilter = doc.createElement("autoFilter");
        var s = this.autoFilter[0];
        var e = this.autoFilter[1]
        autoFilter.setAttribute("ref", Util.positionToLetterRef(s[0], s[1]) + ":" + Util.positionToLetterRef(e[0], e[1] - this.totalsRowCount));
        return autoFilter;
    }


    public exportTableStyleInfo(doc: XmlDom) {
        var ts = this.styleInfo;
        var tableStyle = doc.createElement("tableStyleInfo");
        tableStyle.setAttribute("name", ts.themeStyle);
        tableStyle.setAttribute("showFirstColumn", ts.showFirstColumn ? "1" : "0");
        tableStyle.setAttribute("showLastColumn", ts.showLastColumn ? "1" : "0");
        tableStyle.setAttribute("showColumnStripes", ts.showColumnStripes ? "1" : "0");
        tableStyle.setAttribute("showRowStripes", ts.showRowStripes ? "1" : "0");
        return tableStyle;
    }


    public addAutoFilter(startRef: [number, number], endRef: [number, number]) {
        this.autoFilter = [startRef, endRef];
    }

}

export = Table;