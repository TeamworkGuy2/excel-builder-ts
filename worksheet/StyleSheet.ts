import Util = require("../util/Util");
import XmlDom = require("../xml/XmlDom");

/**
 * @module Excel/StyleSheet
 */
class StyleSheet {
    borders: StyleSheet.Border[];
    cellStyles: { name: string; xfId: string | number; builtinId: string | number; [key: string]: any; }[];
    defaultTableStyle: boolean;
    differentialStyles: StyleSheet.DifferentialStyle[];
    id: string;
    masterCellFormats: StyleSheet.CellFormat[];
    masterCellStyles: StyleSheet.CellStyle[];
    fills: StyleSheet.Fill[];
    fonts: StyleSheet.FontStyle[];
    numberFormatters: StyleSheet.NumberingFormat[];
    tableStyles: StyleSheet.TableStyle[];


    constructor(config?: any) {
        this.id = Util._uniqueId("StyleSheet");
        this.cellStyles = [{
            name: "Normal",
            xfId: "0",
            builtinId: "0"
        }];
        this.defaultTableStyle = false;
        this.differentialStyles = [<any>{}];
        this.masterCellFormats = [{
            numFmtId: 0,
            fontId: 0,
            fillId: 0,
            borderId: 0,
            xfid: 0
        }];
        this.masterCellStyles = [{
            numFmtId: 0,
            fontId: 0,
            fillId: 0,
            borderId: 0
        }];
        this.fonts = [<any>{}];
        this.numberFormatters = [];
        this.fills = [<any>{}, {
            type: "pattern",
            patternType: "gray125",
            fgColor: "FF333333",
            bgColor: "FF333333"
        }];
        this.borders = [{
            top: <any>{},
            left: <any>{},
            right: <any>{},
            bottom: <any>{},
            diagonal: <any>{}
        }];
        this.tableStyles = [];
    }


    public createSimpleFormatter(type: string): { id: number; numFmtId?: number; } {
        var sid = this.masterCellFormats.length;
        var style = {
            id: sid,
            numFmtId: <number><any>undefined,
        };
        switch (type) {
            case "date":
                style.numFmtId = 14;
                break;
        }
        this.masterCellFormats.push(style);
        return style;
    }


    public createFill(fillInstructions: StyleSheet.Fill): StyleSheet.Fill & { id: number } {
        var id = this.fills.length;
        var fill = fillInstructions;
        fill.id = id;
        this.fills.push(<StyleSheet.Fill>fill);
        return <any>fill;
    }


    public createNumberFormatter(formatInstructions: string) {
        var id = this.numberFormatters.length + 100;
        var format: StyleSheet.NumberingFormat = {
            id: id,
            formatCode: formatInstructions
        }
        this.numberFormatters.push(format);
        return format;
    }


    /** alignment:
     *  horizontal: http://www.schemacentral.com/sc/ooxml/t-ssml_ST_HorizontalAlignment.html
     *  vertical: http://www.schemacentral.com/sc/ooxml/t-ssml_ST_VerticalAlignment.html
     */
    public createFormat(styleInstructions: StyleSheet.CellFormatOpts) {
        var sid = this.masterCellFormats.length;
        var style: StyleSheet.CellFormat & { id: number } = {
            id: sid,
            fontId: undefined,
            numFmtId: undefined,
            borderId: undefined,
            fillId: undefined,
            alignment: undefined,
        };
        if (isObj(styleInstructions.font)) {
            style.fontId = this.createFontStyle(styleInstructions.font).id;
        }
        else if (styleInstructions.font) {
            if (isNaN(parseInt(<any>styleInstructions.font, 10))) {
                throw new Error("Passing a non-numeric font id is not supported");
            }
            style.fontId = styleInstructions.font;
        }

        if (isStr(styleInstructions.format)) {
            style.numFmtId = this.createNumberFormatter(styleInstructions.format).id;
        }
        else if (styleInstructions.format) {
            if (isNaN(parseInt(<any>styleInstructions.format))) {
                throw new Error("Invalid number formatter id");
            }
            style.numFmtId = styleInstructions.format;
        }

        if (isObj(styleInstructions.border)) {
            style.borderId = this.createBorderFormatter(styleInstructions.border).id;
        }
        else if (styleInstructions.border) {
            if (isNaN(parseInt(<any>styleInstructions.border))) {
                throw new Error("Passing a non-numeric border id is not supported");
            }
            style.borderId = styleInstructions.border;
        }

        if (isObj(styleInstructions.fill)) {
            style.fillId = this.createFill(styleInstructions.fill).id;
        }
        else if (styleInstructions.fill) {
            if (isNaN(parseInt(<any>styleInstructions.fill))) {
                throw new Error("Passing a non-numeric fill id is not supported");
            }
            style.fillId = styleInstructions.fill;
        }

        if (isObj(styleInstructions.alignment)) {
            style.alignment = Util.pick<Partial<StyleSheet.Alignment>, keyof Partial<StyleSheet.Alignment>>(styleInstructions.alignment, [
                "horizontal",
                "justifyLastLine",
                "readingOrder",
                "relativeIndent",
                "shrinkToFit",
                "textRotation",
                "vertical",
                "wrapText"
            ]);
        }

        this.masterCellFormats.push(style);
        return style;
    }


    public createDifferentialStyle(styleInstructions: { font?: StyleSheet.FontStyle; border?: StyleSheet.Border; fill?: StyleSheet.Fill; alignment?: object; format?: string; }) {
        var id = this.differentialStyles.length;
        var style: StyleSheet.DifferentialStyle & { id: number } = {
            id: id,
            alignment: undefined,
            border: undefined,
            fill: undefined,
            font: undefined,
            numFmt: undefined,
        }
        if (isObj(styleInstructions.font)) {
            style.font = styleInstructions.font;
        }
        if (isObj(styleInstructions.border)) {
            style.border = Util.defaults(styleInstructions.border, {
                top: <any>{},
                left: <any>{},
                right: <any>{},
                bottom: <any>{},
                diagonal: <any>{}
            });
        }
        if (isObj(styleInstructions.fill)) {
            style.fill = styleInstructions.fill;
        }
        if (isObj(styleInstructions.alignment)) {
            style.alignment = styleInstructions.alignment;
        }
        if (isStr(styleInstructions.format)) {
            style.numFmt = this.createNumberFormatter(styleInstructions.format);
        }
        this.differentialStyles[id] = style;
        return style;
    }


    /** Should be an object containing keys that match with one of the keys from this list:
     * http://www.schemacentral.com/sc/ooxml/t-ssml_ST_TableStyleType.html
     * The value should be a reference to a differential format (dxf)
     */
    public createTableStyle(instructions: StyleSheet.TableStyle) {
        this.tableStyles.push(instructions);
    }


    /** All params optional
     * Expects: {
     * top: {},
     * left: {},
     * right: {},
     * bottom: {},
     * diagonal: {},
     * outline: boolean,
     * diagonalUp: boolean,
     * diagonalDown: boolean
     * }
     * Each border should follow:
     * {
     * style: styleString, http://www.schemacentral.com/sc/ooxml/t-ssml_ST_BorderStyle.html
     * color: ARBG color (requires the A, so for example FF006666)
     * }
     */
    public createBorderFormatter(border?: StyleSheet.BorderOpts) {
        var res = <StyleSheet.Border & { id: number }>Util.defaults(<Exclude<typeof border, undefined>>border, {
            top: {},
            left: {},
            right: {},
            bottom: {},
            diagonal: {},
            id: this.borders.length
        });
        this.borders.push(res);
        return res;
    }


    /** Supported font styles:
     * bold
     * italic
     * underline (single, double, singleAccounting, doubleAccounting)
     * size
     * color
     * fontName
     * strike (strikethrough)
     * outline (does this actually do anything?)
     * shadow (does this actually do anything?)
     * superscript
     * subscript
     *
     * Color is a future goal - at the moment it's looking a bit complicated
     */
    public createFontStyle(instructions: StyleSheet.FontStyleOpts) {
        var fontId = this.fonts.length;
        var fontStyle: StyleSheet.FontStyle & { id: number } = {
            id: fontId,
            bold: undefined,
            color: undefined,
            fontName: undefined,
            italic: undefined,
            outline: undefined,
            shadow: undefined,
            size: undefined,
            strike: undefined,
            vertAlign: undefined,
            underline: undefined,
        };
        if (instructions.bold) {
            fontStyle.bold = true;
        }
        if (instructions.italic) {
            fontStyle.italic = true;
        }
        if (instructions.superscript) {
            fontStyle.vertAlign = "superscript";
        }
        if (instructions.subscript) {
            fontStyle.vertAlign = "subscript";
        }
        if (instructions.underline) {
            if (["double", "singleAccounting", "doubleAccounting"].indexOf(<string>instructions.underline) != -1) {
                fontStyle.underline = <any>instructions.underline;
            }
            else {
                fontStyle.underline = true;
            }
        }
        if (instructions.strike) {
            fontStyle.strike = true;
        }
        if (instructions.outline) {
            fontStyle.outline = true;
        }
        if (instructions.shadow) {
            fontStyle.shadow = true;
        }
        if (instructions.size) {
            fontStyle.size = instructions.size;
        }
        if (instructions.color) {
            fontStyle.color = instructions.color;
        }
        if (instructions.fontName) {
            fontStyle.fontName = instructions.fontName;
        }
        this.fonts.push(fontStyle);
        return fontStyle;
    }


    public exportBorders(doc: XmlDom) {
        var borders = doc.createElement("borders");
        borders.setAttribute("count", this.borders.length);

        for (var i = 0, l = this.borders.length; i < l; i++) {
            borders.appendChild(this.exportBorder(doc, this.borders[i]));
        }
        return borders;
    }


    public exportBorder(doc: XmlDom, data: StyleSheet.Border) {
        var border = doc.createElement("border");
        var self = this;
        function borderGenerator(name: keyof StyleSheet.Border) {
            var b = doc.createElement(name);
            border.appendChild(b);
            var propVal = <StyleSheet.BorderProperty>data[name];
            if (propVal.style) {
                b.setAttribute("style", propVal.style);
            }
            if (propVal.color) {
                b.appendChild(self.exportColor(doc, propVal.color));
            }
            return b;
        };
        border.appendChild(borderGenerator("left"));
        border.appendChild(borderGenerator("right"));
        border.appendChild(borderGenerator("top"));
        border.appendChild(borderGenerator("bottom"));
        border.appendChild(borderGenerator("diagonal"));
        return border;
    }


    public exportColor(doc: XmlDom, color: string | StyleSheet.Color) {
        var colorEl = doc.createElement("color");
        if (isStr(color)) {
            colorEl.setAttribute("rgb", color);
        }
        else {
            if (color.tint != null) {
                colorEl.setAttribute("tint", color.tint);
            }
            if (color.auto != null) {
                colorEl.setAttribute("auto", !!color.auto);
            }
            if (color.theme != null) {
                colorEl.setAttribute("theme", color.theme);
            }
        }
        return colorEl;
    }


    public exportMasterCellFormats(doc: XmlDom) {
        var cellFormats = Util.createElement(doc, "cellXfs", [
            ["count", this.masterCellFormats.length]
        ]);
        for (var i = 0, l = this.masterCellFormats.length; i < l; i++) {
            var mformat = this.masterCellFormats[i];
            cellFormats.appendChild(this.exportCellFormatElement(doc, mformat));
        }
        return cellFormats;
    }


    public exportMasterCellStyles(doc: XmlDom) {
        var records = Util.createElement(doc, "cellStyleXfs", [
            ["count", this.masterCellStyles.length]
        ]);
        for (var i = 0, l = this.masterCellStyles.length; i < l; i++) {
            var mstyle = this.masterCellStyles[i];
            records.appendChild(this.exportCellFormatElement(doc, mstyle));
        }
        return records;
    }


    public exportCellFormatElement(doc: XmlDom, styleInstructions: StyleSheet.CellFormat) {
        var xf = doc.createElement("xf"), i = 0;
        var allowed = ["applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat",
            "applyProtection", "borderId", "fillId", "fontId", "numFmtId", "pivotButton", "quotePrefix", "xfId"]
        var attributes = <(keyof StyleSheet.CellFormat)[]>Object.keys(styleInstructions).filter((key) => allowed.indexOf(key) != -1);

        if (styleInstructions.alignment) {
            var alignmentData = styleInstructions.alignment;
            xf.appendChild(this.exportAlignment(doc, alignmentData));
        }
        var a = attributes.length;
        while (a--) {
            xf.setAttribute(attributes[a], styleInstructions[attributes[a]]);
        }
        if (styleInstructions.fillId) {
            xf.setAttribute("applyFill", '1');
        }
        return xf;
    }


    public exportAlignment(doc: XmlDom, alignmentData: { [key: string]: any }) {
        var alignment = doc.createElement("alignment");
        var keys = Object.keys(alignmentData);
        for (var i = 0, len = keys.length; i < len; i++) {
            alignment.setAttribute(keys[i], alignmentData[keys[i]]);
        }
        return alignment;
    }


    public exportFonts(doc: XmlDom) {
        var fonts = doc.createElement("fonts");
        fonts.setAttribute("count", this.fonts.length);
        for (var i = 0, l = this.fonts.length; i < l; i++) {
            var fd = this.fonts[i];
            fonts.appendChild(this.exportFont(doc, fd));
        }
        return fonts;
    }


    public exportFont(doc: XmlDom, fd: StyleSheet.FontStyle) {
        var font = doc.createElement("font");
        if (fd.size) {
            var size = doc.createElement("sz");
            size.setAttribute("val", fd.size);
            font.appendChild(size);
        }

        if (fd.fontName) {
            var fontName = doc.createElement("name");
            fontName.setAttribute("val", fd.fontName);
            font.appendChild(fontName);
        }

        if (fd.bold) {
            font.appendChild(doc.createElement("b"));
        }
        if (fd.italic) {
            font.appendChild(doc.createElement("i"));
        }
        if (fd.vertAlign) {
            var vertAlign = doc.createElement("vertAlign");
            vertAlign.setAttribute("val", fd.vertAlign);
            font.appendChild(vertAlign);
        }
        if (fd.underline) {
            var u = doc.createElement("u");
            if (fd.underline !== true) {
                u.setAttribute("val", fd.underline);
            }
            font.appendChild(u);
        }
        if (fd.strike) {
            font.appendChild(doc.createElement("strike"));
        }
        if (fd.shadow) {
            font.appendChild(doc.createElement("shadow"));
        }
        if (fd.outline) {
            font.appendChild(doc.createElement("outline"));
        }
        if (fd.color) {
            font.appendChild(this.exportColor(doc, fd.color));
        }
        return font;
    }


    public exportFills(doc: XmlDom) {
        var fills = doc.createElement("fills");
        fills.setAttribute("count", this.fills.length);
        for (var i = 0, l = this.fills.length; i < l; i++) {
            var fd = this.fills[i];
            fills.appendChild(this.exportFill(doc, fd));
        }
        return fills;
    }


    public exportFill(doc: XmlDom, fd: StyleSheet.Fill) {
        var fill = doc.createElement("fill");
        if (fd.type == "pattern") {
            var fillDef = this.exportPatternFill(doc, fd);
            fill.appendChild(fillDef);
        }
        else if (fd.type == "gradient") {
            var fillDef = this.exportGradientFill(doc, fd);
            fill.appendChild(fillDef);
        }
        return fill;
    }


    public exportGradientFill(doc: XmlDom, data: StyleSheet.Fill) {
        var fillDef = doc.createElement("gradientFill");
        if (data.degree) {
            fillDef.setAttribute("degree", data.degree);
        }
        else if (data.left) {
            fillDef.setAttribute("left", data.left);
            fillDef.setAttribute("right", data.right);
            fillDef.setAttribute("top", data.top);
            fillDef.setAttribute("bottom", data.bottom);
        }

        var start = doc.createElement("stop");
        var dataStart = <NonNullable<typeof data.start>><any>data.start;
        start.setAttribute("position", (<any>dataStart).pureAt || 0);
        var startColor = doc.createElement("color");
        if (isStr(dataStart) || dataStart.color) {
            startColor.setAttribute("rgb", (<any>dataStart).color || dataStart);
        }
        else if (typeof dataStart.theme) {
            startColor.setAttribute("theme", dataStart.theme);
        }

        var end = doc.createElement("stop");
        var endColor = doc.createElement("color");
        var dataEnd = <NonNullable<typeof data.end>><any>data.end;
        end.setAttribute("position", (<any>dataEnd).pureAt || 1);
        if (isStr(dataEnd) || dataEnd.color) {
            endColor.setAttribute("rgb", (<any>dataEnd).color || dataEnd);
        }
        else if (typeof dataEnd.theme) {
            endColor.setAttribute("theme", dataEnd.theme);
        }

        start.appendChild(startColor);
        end.appendChild(endColor);
        fillDef.appendChild(start);
        fillDef.appendChild(end);
        return fillDef;
    }


    /**
     * Pattern types: http://www.schemacentral.com/sc/ooxml/t-ssml_ST_PatternType.html
     */
    public exportPatternFill(doc: XmlDom, data: StyleSheet.Fill) {
        var fillDef = Util.createElement(doc, "patternFill", [
            ["patternType", data.patternType]
        ]);
        var bgColor = (!data.bgColor ? data.bgColor = "FFFFFFFF" : data.bgColor);
        var fgColor = (!data.fgColor ? data.fgColor = "FFFFFFFF" : data.fgColor);

        var bgColorElem = doc.createElement("bgColor");
        if (isStr(bgColor)) {
            bgColorElem.setAttribute("rgb", bgColor)
        }
        else {
            if (bgColor.theme) {
                bgColorElem.setAttribute("theme", bgColor.theme);
            }
            else {
                bgColorElem.setAttribute("rgb", bgColor.rbg);
            }
        }

        var fgColorElem = doc.createElement("fgColor");
        if (isStr(fgColor)) {
            fgColorElem.setAttribute("rgb", fgColor)
        }
        else {
            if (fgColor.theme) {
                fgColorElem.setAttribute("theme", fgColor.theme);
            }
            else {
                fgColorElem.setAttribute("rgb", fgColor.rbg);
            }
        }

        fillDef.appendChild(fgColorElem);
        fillDef.appendChild(bgColorElem);
        return fillDef;
    }


    public exportNumberFormatters(doc: XmlDom) {
        var formatters = doc.createElement("numFmts");
        formatters.setAttribute("count", this.numberFormatters.length);
        for (var i = 0, l = this.numberFormatters.length; i < l; i++) {
            var fd = this.numberFormatters[i];
            formatters.appendChild(this.exportNumberFormatter(doc, fd));
        }
        return formatters;
    }


    public exportNumberFormatter(doc: XmlDom, fd: StyleSheet.NumberingFormat) {
        var numFmt = doc.createElement("numFmt");
        numFmt.setAttribute("numFmtId", fd.id);
        numFmt.setAttribute("formatCode", fd.formatCode);
        return numFmt;
    }


    public exportCellStyles(doc: XmlDom) {
        var cellStylesElem = doc.createElement("cellStyles");
        cellStylesElem.setAttribute("count", this.cellStyles.length);

        for (var i = 0, l = this.cellStyles.length; i < l; i++) {
            var style = this.cellStyles[i];
            delete style.id; //Remove internal id
            var record = Util.createElement(doc, "cellStyle");
            cellStylesElem.appendChild(record);
            var attributes = Object.keys(style);
            var a = attributes.length;
            while (a--) {
                record.setAttribute(attributes[a], style[attributes[a]]);
            }
        }

        return cellStylesElem;
    }


    public exportDifferentialStyles(doc: XmlDom) {
        var dxfs = doc.createElement("dxfs");
        dxfs.setAttribute("count", this.differentialStyles.length);

        for (var i = 0, l = this.differentialStyles.length; i < l; i++) {
            var style = this.differentialStyles[i];
            dxfs.appendChild(this.exportDFX(doc, style));
        }

        return dxfs;
    }


    public exportDFX(doc: XmlDom, style: StyleSheet.DifferentialStyle) {
        var dxf = doc.createElement("dxf");
        if (style.font) {
            dxf.appendChild(this.exportFont(doc, style.font));
        }
        if (style.fill) {
            dxf.appendChild(this.exportFill(doc, style.fill));
        }
        if (style.border) {
            dxf.appendChild(this.exportBorder(doc, style.border));
        }
        if (style.numFmt) {
            dxf.appendChild(this.exportNumberFormatter(doc, style.numFmt));
        }
        if (style.alignment) {
            dxf.appendChild(this.exportAlignment(doc, style.alignment));
        }
        return dxf;
    }


    public exportTableStyles(doc: XmlDom) {
        var tableStyles = doc.createElement("tableStyles");
        tableStyles.setAttribute("count", this.tableStyles.length);
        if (this.defaultTableStyle) {
            tableStyles.setAttribute("defaultTableStyle", this.defaultTableStyle);
        }
        for (var i = 0, l = this.tableStyles.length; i < l; i++) {
            tableStyles.appendChild(this.exportTableStyle(doc, this.tableStyles[i]));
        }
        return tableStyles;
    }


    public exportTableStyle(doc: XmlDom, style: StyleSheet.TableStyle) {
        var tableStyle = doc.createElement("tableStyle");
        tableStyle.setAttribute("name", style.name);
        tableStyle.setAttribute("pivot", 0);
        var i = 0;

        Object.keys(style).forEach(function (key) {
            var value = style[key];
            if (key == "name") { return; }
            i++;
            var styleEl = doc.createElement("tableStyleElement");
            styleEl.setAttribute("type", key);
            styleEl.setAttribute("dxfId", value);
            tableStyle.appendChild(styleEl);
        });
        tableStyle.setAttribute("count", i);
        return tableStyle;
    }


    public toXML() {
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, "styleSheet");
        var styleSheet = doc.documentElement;
        styleSheet.appendChild(this.exportNumberFormatters(doc));
        styleSheet.appendChild(this.exportFonts(doc));
        styleSheet.appendChild(this.exportFills(doc));
        styleSheet.appendChild(this.exportBorders(doc));
        styleSheet.appendChild(this.exportMasterCellStyles(doc));
        styleSheet.appendChild(this.exportMasterCellFormats(doc));
        styleSheet.appendChild(this.exportCellStyles(doc));
        styleSheet.appendChild(this.exportDifferentialStyles(doc));
        if (this.tableStyles.length) {
            styleSheet.appendChild(this.exportTableStyles(doc));
        }
        return doc;
    }

}

module StyleSheet {
    // All ST_* types have exact names taken from W3C XML §A.2

    type ST_BorderStyle = ("none" | "thin" | "medium" | "dashed" | "dotted" | "thick" | "double" | "hair" | "mediumDashed" | "dashDot" | "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot");


    export interface Alignment {
        horizontal?: string;
        //indent?: number; // TODO
        justifyLastLine?: 0 | 1;
        readingOrder?: number;
        relativeIndent?: number;
        shrinkToFit?: 0 | 1;
        textRotation?: number;
        vertical?: string;
        wrapText?: 0 | 1;
    }


    export interface Border {
        id?: number;
        start?: BorderProperty;
        end?: BorderProperty;
        top?: BorderProperty;
        bottom?: BorderProperty;
        diagonal?: BorderProperty;
        vertical?: BorderProperty;
        horizontal?: BorderProperty;
        diagonalDown?: 0 | 1 | boolean;
        diagonalUp?: 0 | 1 | boolean;
        outline?: 0 | 1 | boolean;

        // not part of the spec but required by MS Office
        left?: BorderProperty;
        right?: BorderProperty;
    }


    export interface BorderProperty {
        color?: Color;
        style?: ST_BorderStyle;
    }

    export interface BorderOpts {
        start?: BorderProperty;
        end?: BorderProperty;
        top?: BorderProperty;
        bottom?: BorderProperty;
        diagonal?: BorderProperty;
        vertical?: BorderProperty;
        horizontal?: BorderProperty;
        diagonalDown?: 0 | 1 | boolean;
        diagonalUp?: 0 | 1 | boolean;
        outline?: 0 | 1 | boolean;
        // not part of the spec but required by MS Office
        left?: BorderProperty;
        right?: BorderProperty;
        [id: string]: any;
    }


    export interface CellFormat {
        id?: number;
        numFmtId?: number;
        fontId?: number;
        fillId?: number;
        borderId?: number;
        xfid?: number;
        alignment?: Partial<Alignment>;
    }


    export interface CellFormatOpts {
        font?: FontStyleOpts | number;
        format?: string | number;
        border?: BorderOpts | number;
        fill?: Fill | number;
        alignment?: Partial<Alignment>;
    }


    export interface CellStyle {
        borderId?: number;
        fillId?: number;
        fontId?: number;
        numFmtId?: number;
    }


    export interface Color {
        auto?: boolean;
        indexed?: number;
        rgb?: string | number;
        theme?: number;
        tint?: number;
    }


    export interface DifferentialStyle {
        id?: number;
        alignment?: { [key: string]: any };
        border?: Border;
        fill?: Fill;
        font?: FontStyle;
        numFmt?: NumberingFormat;
    }


    export interface Fill {
        id?: number;
        type: string; // 'pattern'
        patternType: string;
        // Pattern fill
        bgColor?: string | { theme?: string; rbg?: string; }; // ARGB
        fgColor?: string | { theme?: string; rbg?: string; }; // ARGB
        // Gradient fill
        degree?: any;
        left?: any;
        right?: any;
        top?: any;
        bottom?: any;
        start?: string | { pureAt?: number; color?: string; theme?: string; };
        end?: string | { pureAt?: number; color?: string; theme?: string; };
    }


    export interface FontStyle {
        id?: number;
        bold?: boolean;
        color?: string;
        fontName?: string;
        italic?: boolean;
        outline?: boolean;
        shadow?: boolean;
        strike?: boolean;
        size?: number;
        vertAlign?: string;
        underline?: boolean | ("single" | "double" | "singleAccounting" | "doubleAccounting" | "none");
    }


    export interface FontStyleOpts {
        bold?: boolean;
        color?: string;
        fontName?: string;
        italic?: boolean;
        outline?: boolean;
        shadow?: boolean;
        strike?: boolean;
        size?: number;
        superscript?: boolean;
        subscript?: boolean;
        underline?: FontStyle["underline"];
    }


    export interface NumberingFormat {
        id: number;
        formatCode: string;
    }


    export interface TableStyle {
        name: string;
        [key: string]: any;
    }

}


var toStrFunc = Object.prototype.toString;

function isObj(obj: any): obj is object {
    return obj && toStrFunc.call(obj) === "[object Object]";
}

function isStr(str: any): str is string {
    return str && toStrFunc.call(str) === "[object String]";
}


export = StyleSheet;