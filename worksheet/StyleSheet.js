"use strict";
var Util = require("../util/Util");
/**
 * @module Excel/StyleSheet
 */
var StyleSheet = /** @class */ (function () {
    function StyleSheet(config) {
        this.id = Util._uniqueId("StyleSheet");
        this.cellStyles = [{
                name: "Normal",
                xfId: "0",
                builtinId: "0"
            }];
        this.defaultTableStyle = false;
        this.differentialStyles = [{}];
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
        this.fonts = [{}];
        this.numberFormatters = [];
        this.fills = [{}, {
                type: "pattern",
                patternType: "gray125",
                fgColor: "FF333333",
                bgColor: "FF333333"
            }];
        this.borders = [{
                top: {},
                left: {},
                right: {},
                bottom: {},
                diagonal: {}
            }];
        this.tableStyles = [];
    }
    StyleSheet.prototype.createSimpleFormatter = function (type) {
        var sid = this.masterCellFormats.length;
        var style = {
            id: sid,
            numFmtId: undefined,
        };
        switch (type) {
            case "date":
                style.numFmtId = 14;
                break;
        }
        this.masterCellFormats.push(style);
        return style;
    };
    StyleSheet.prototype.createFill = function (fillInstructions) {
        var id = this.fills.length;
        var fill = fillInstructions;
        fill.id = id;
        this.fills.push(fill);
        return fill;
    };
    StyleSheet.prototype.createNumberFormatter = function (formatInstructions) {
        var id = this.numberFormatters.length + 100;
        var format = {
            id: id,
            formatCode: formatInstructions
        };
        this.numberFormatters.push(format);
        return format;
    };
    /** alignment:
     *  horizontal: http://www.schemacentral.com/sc/ooxml/t-ssml_ST_HorizontalAlignment.html
     *  vertical: http://www.schemacentral.com/sc/ooxml/t-ssml_ST_VerticalAlignment.html
     */
    StyleSheet.prototype.createFormat = function (styleInstructions) {
        var sid = this.masterCellFormats.length;
        var style = {
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
            if (isNaN(parseInt(styleInstructions.font, 10))) {
                throw new Error("Passing a non-numeric font id is not supported");
            }
            style.fontId = styleInstructions.font;
        }
        if (isStr(styleInstructions.format)) {
            style.numFmtId = this.createNumberFormatter(styleInstructions.format).id;
        }
        else if (styleInstructions.format) {
            if (isNaN(parseInt(styleInstructions.format))) {
                throw new Error("Invalid number formatter id");
            }
            style.numFmtId = styleInstructions.format;
        }
        if (isObj(styleInstructions.border)) {
            style.borderId = this.createBorderFormatter(styleInstructions.border).id;
        }
        else if (styleInstructions.border) {
            if (isNaN(parseInt(styleInstructions.border))) {
                throw new Error("Passing a non-numeric border id is not supported");
            }
            style.borderId = styleInstructions.border;
        }
        if (isObj(styleInstructions.fill)) {
            style.fillId = this.createFill(styleInstructions.fill).id;
        }
        else if (styleInstructions.fill) {
            if (isNaN(parseInt(styleInstructions.fill))) {
                throw new Error("Passing a non-numeric fill id is not supported");
            }
            style.fillId = styleInstructions.fill;
        }
        if (isObj(styleInstructions.alignment)) {
            style.alignment = Util.pick(styleInstructions.alignment, [
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
    };
    StyleSheet.prototype.createDifferentialStyle = function (styleInstructions) {
        var id = this.differentialStyles.length;
        var style = {
            id: id,
            alignment: undefined,
            border: undefined,
            fill: undefined,
            font: undefined,
            numFmt: undefined,
        };
        if (isObj(styleInstructions.font)) {
            style.font = styleInstructions.font;
        }
        if (isObj(styleInstructions.border)) {
            style.border = Util.defaults(styleInstructions.border, {
                top: {},
                left: {},
                right: {},
                bottom: {},
                diagonal: {}
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
    };
    /** Should be an object containing keys that match with one of the keys from this list:
     * http://www.schemacentral.com/sc/ooxml/t-ssml_ST_TableStyleType.html
     * The value should be a reference to a differential format (dxf)
     */
    StyleSheet.prototype.createTableStyle = function (instructions) {
        this.tableStyles.push(instructions);
    };
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
    StyleSheet.prototype.createBorderFormatter = function (border) {
        var res = Util.defaults(border, {
            top: {},
            left: {},
            right: {},
            bottom: {},
            diagonal: {},
            id: this.borders.length
        });
        this.borders.push(res);
        return res;
    };
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
    StyleSheet.prototype.createFontStyle = function (instructions) {
        var fontId = this.fonts.length;
        var fontStyle = {
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
            if (["double", "singleAccounting", "doubleAccounting"].indexOf(instructions.underline) != -1) {
                fontStyle.underline = instructions.underline;
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
    };
    StyleSheet.prototype.exportBorders = function (doc) {
        var borders = doc.createElement("borders");
        borders.setAttribute("count", this.borders.length);
        for (var i = 0, l = this.borders.length; i < l; i++) {
            borders.appendChild(this.exportBorder(doc, this.borders[i]));
        }
        return borders;
    };
    StyleSheet.prototype.exportBorder = function (doc, data) {
        var border = doc.createElement("border");
        var self = this;
        function borderGenerator(name) {
            var b = doc.createElement(name);
            border.appendChild(b);
            var propVal = data[name];
            if (propVal.style) {
                b.setAttribute("style", propVal.style);
            }
            if (propVal.color) {
                b.appendChild(self.exportColor(doc, propVal.color));
            }
            return b;
        }
        ;
        border.appendChild(borderGenerator("left"));
        border.appendChild(borderGenerator("right"));
        border.appendChild(borderGenerator("top"));
        border.appendChild(borderGenerator("bottom"));
        border.appendChild(borderGenerator("diagonal"));
        return border;
    };
    StyleSheet.prototype.exportColor = function (doc, color) {
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
    };
    StyleSheet.prototype.exportMasterCellFormats = function (doc) {
        var cellFormats = Util.createElement(doc, "cellXfs", [
            ["count", this.masterCellFormats.length]
        ]);
        for (var i = 0, l = this.masterCellFormats.length; i < l; i++) {
            var mformat = this.masterCellFormats[i];
            cellFormats.appendChild(this.exportCellFormatElement(doc, mformat));
        }
        return cellFormats;
    };
    StyleSheet.prototype.exportMasterCellStyles = function (doc) {
        var records = Util.createElement(doc, "cellStyleXfs", [
            ["count", this.masterCellStyles.length]
        ]);
        for (var i = 0, l = this.masterCellStyles.length; i < l; i++) {
            var mstyle = this.masterCellStyles[i];
            records.appendChild(this.exportCellFormatElement(doc, mstyle));
        }
        return records;
    };
    StyleSheet.prototype.exportCellFormatElement = function (doc, styleInstructions) {
        var xf = doc.createElement("xf"), i = 0;
        var allowed = ["applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat",
            "applyProtection", "borderId", "fillId", "fontId", "numFmtId", "pivotButton", "quotePrefix", "xfId"];
        var attributes = Object.keys(styleInstructions).filter(function (key) { return allowed.indexOf(key) != -1; });
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
    };
    StyleSheet.prototype.exportAlignment = function (doc, alignmentData) {
        var alignment = doc.createElement("alignment");
        var keys = Object.keys(alignmentData);
        for (var i = 0, len = keys.length; i < len; i++) {
            alignment.setAttribute(keys[i], alignmentData[keys[i]]);
        }
        return alignment;
    };
    StyleSheet.prototype.exportFonts = function (doc) {
        var fonts = doc.createElement("fonts");
        fonts.setAttribute("count", this.fonts.length);
        for (var i = 0, l = this.fonts.length; i < l; i++) {
            var fd = this.fonts[i];
            fonts.appendChild(this.exportFont(doc, fd));
        }
        return fonts;
    };
    StyleSheet.prototype.exportFont = function (doc, fd) {
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
    };
    StyleSheet.prototype.exportFills = function (doc) {
        var fills = doc.createElement("fills");
        fills.setAttribute("count", this.fills.length);
        for (var i = 0, l = this.fills.length; i < l; i++) {
            var fd = this.fills[i];
            fills.appendChild(this.exportFill(doc, fd));
        }
        return fills;
    };
    StyleSheet.prototype.exportFill = function (doc, fd) {
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
    };
    StyleSheet.prototype.exportGradientFill = function (doc, data) {
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
        var dataStart = data.start;
        start.setAttribute("position", dataStart.pureAt || 0);
        var startColor = doc.createElement("color");
        if (isStr(dataStart) || dataStart.color) {
            startColor.setAttribute("rgb", dataStart.color || dataStart);
        }
        else if (typeof dataStart.theme) {
            startColor.setAttribute("theme", dataStart.theme);
        }
        var end = doc.createElement("stop");
        var endColor = doc.createElement("color");
        var dataEnd = data.end;
        end.setAttribute("position", dataEnd.pureAt || 1);
        if (isStr(dataEnd) || dataEnd.color) {
            endColor.setAttribute("rgb", dataEnd.color || dataEnd);
        }
        else if (typeof dataEnd.theme) {
            endColor.setAttribute("theme", dataEnd.theme);
        }
        start.appendChild(startColor);
        end.appendChild(endColor);
        fillDef.appendChild(start);
        fillDef.appendChild(end);
        return fillDef;
    };
    /**
     * Pattern types: http://www.schemacentral.com/sc/ooxml/t-ssml_ST_PatternType.html
     */
    StyleSheet.prototype.exportPatternFill = function (doc, data) {
        var fillDef = Util.createElement(doc, "patternFill", [
            ["patternType", data.patternType]
        ]);
        var bgColor = (!data.bgColor ? data.bgColor = "FFFFFFFF" : data.bgColor);
        var fgColor = (!data.fgColor ? data.fgColor = "FFFFFFFF" : data.fgColor);
        var bgColorElem = doc.createElement("bgColor");
        if (isStr(bgColor)) {
            bgColorElem.setAttribute("rgb", bgColor);
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
            fgColorElem.setAttribute("rgb", fgColor);
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
    };
    StyleSheet.prototype.exportNumberFormatters = function (doc) {
        var formatters = doc.createElement("numFmts");
        formatters.setAttribute("count", this.numberFormatters.length);
        for (var i = 0, l = this.numberFormatters.length; i < l; i++) {
            var fd = this.numberFormatters[i];
            formatters.appendChild(this.exportNumberFormatter(doc, fd));
        }
        return formatters;
    };
    StyleSheet.prototype.exportNumberFormatter = function (doc, fd) {
        var numFmt = doc.createElement("numFmt");
        numFmt.setAttribute("numFmtId", fd.id);
        numFmt.setAttribute("formatCode", fd.formatCode);
        return numFmt;
    };
    StyleSheet.prototype.exportCellStyles = function (doc) {
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
    };
    StyleSheet.prototype.exportDifferentialStyles = function (doc) {
        var dxfs = doc.createElement("dxfs");
        dxfs.setAttribute("count", this.differentialStyles.length);
        for (var i = 0, l = this.differentialStyles.length; i < l; i++) {
            var style = this.differentialStyles[i];
            dxfs.appendChild(this.exportDFX(doc, style));
        }
        return dxfs;
    };
    StyleSheet.prototype.exportDFX = function (doc, style) {
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
    };
    StyleSheet.prototype.exportTableStyles = function (doc) {
        var tableStyles = doc.createElement("tableStyles");
        tableStyles.setAttribute("count", this.tableStyles.length);
        if (this.defaultTableStyle) {
            tableStyles.setAttribute("defaultTableStyle", this.defaultTableStyle);
        }
        for (var i = 0, l = this.tableStyles.length; i < l; i++) {
            tableStyles.appendChild(this.exportTableStyle(doc, this.tableStyles[i]));
        }
        return tableStyles;
    };
    StyleSheet.prototype.exportTableStyle = function (doc, style) {
        var tableStyle = doc.createElement("tableStyle");
        tableStyle.setAttribute("name", style.name);
        tableStyle.setAttribute("pivot", 0);
        var i = 0;
        Object.keys(style).forEach(function (key) {
            var value = style[key];
            if (key == "name") {
                return;
            }
            i++;
            var styleEl = doc.createElement("tableStyleElement");
            styleEl.setAttribute("type", key);
            styleEl.setAttribute("dxfId", value);
            tableStyle.appendChild(styleEl);
        });
        tableStyle.setAttribute("count", i);
        return tableStyle;
    };
    StyleSheet.prototype.toXML = function () {
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
    };
    return StyleSheet;
}());
var toStrFunc = Object.prototype.toString;
function isObj(obj) {
    return obj && toStrFunc.call(obj) === "[object Object]";
}
function isStr(str) {
    return str && toStrFunc.call(str) === "[object String]";
}
module.exports = StyleSheet;
