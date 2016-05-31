"use strict";
var XmlDom = require("./XmlDom");
/**
 * @module Excel/util
 */
var Util = (function () {
    function Util() {
    }
    Util._uniqueId = function (space) {
        var id = ++Util._id;
        return space + id;
    };
    /**
     * Returns a number based on a namespace. So, running with 'Picture' will return 1. Run again, you will get 2. Run with 'Foo', you'll get 1.
     * @param {string} space
     * @returns {number}
     */
    Util.uniqueId = function (space) {
        if (!this._idSpaces[space]) {
            this._idSpaces[space] = 1;
        }
        return this._idSpaces[space]++;
    };
    Util.pick = function (obj, props) {
        var res = {};
        for (var i = 0, size = props.length; i < size; i++) {
            var key = props[i];
            if (key in obj) {
                res[key] = obj[key];
            }
        }
        return res;
    };
    Util.defaults = function (obj, overrides) {
        for (var key in overrides) {
            if (overrides.hasOwnProperty(key) && (obj[key] === undefined)) {
                obj[key] = overrides[key];
            }
        }
        return obj;
    };
    /**
     * Attempts to create an XML document. Due to limitations in web workers,
     * it may return a 'fake' xml document created from the XMLDOM.js file.
     *
     * Takes a namespace to start the xml file in, as well as the root element
     * of the xml file.
     *
     * @param {type} ns
     * @param {type} base
     * @returns {ActiveXObject|@exp;document@pro;implementation@call;createDocument|@new;XMLDOM}
     */
    Util.createXmlDoc = function (ns, base) {
        if (typeof document === "undefined") {
            return new XmlDom(ns || null, base, null);
        }
        if (document.implementation && document.implementation.createDocument) {
            return document.implementation.createDocument(ns || null, base, null);
        }
        else if (window["ActiveXObject"]) {
            var doc = new ActiveXObject("Microsoft.XMLDOM");
            var rootNode = doc.createElement(base);
            rootNode.setAttribute("xmlns", ns);
            doc.appendChild(rootNode);
            return doc;
        }
        throw new Error("No xml document generator");
    };
    /**
     * Creates an xml node (element). Used to simplify some calls, as IE is
     * very particular about namespaces and such.
     *
     * @param {XMLDOM} doc An xml document (actual DOM or fake DOM, not a string)
     * @param {type} name The name of the element
     * @param {type} attributes
     * @returns {XML Node}
     */
    Util.createElement = function (doc, name, attributes) {
        var el = doc.createElement(name);
        var ie = !el.setAttributeNS;
        attributes = attributes || [];
        var i = attributes.length;
        while (i--) {
            if (!ie && attributes[i][0].indexOf("xmlns") != -1) {
                el.setAttributeNS("http://www.w3.org/2000/xmlns/", attributes[i][0], attributes[i][1]);
            }
            else {
                el.setAttribute(attributes[i][0], attributes[i][1]);
            }
        }
        return el;
    };
    Util.positionToLetterRef = function (x, y) {
        var digit = 1;
        var num = x;
        var string = "";
        var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        if (Util.LETTER_REFS[x]) {
            return Util.LETTER_REFS[x].concat(y);
        }
        var index;
        while (num > 0) {
            num -= Math.pow(26, digit - 1);
            index = num % Math.pow(26, digit);
            num -= index;
            index = index / Math.pow(26, digit - 1);
            string = alphabet.charAt(index) + string;
            digit += 1;
        }
        Util.LETTER_REFS[x] = string;
        return string.concat(y);
    };
    Util._idSpaces = {};
    Util._id = 0;
    Util.LETTER_REFS = {};
    Util.schemas = {
        "worksheet": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "sharedStrings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        "stylesheet": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "relationships": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "relationshipPackage": "http://schemas.openxmlformats.org/package/2006/relationships",
        "contentTypes": "http://schemas.openxmlformats.org/package/2006/content-types",
        "spreadsheetml": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "markupCompat": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
        "officeDocument": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "package": "http://schemas.openxmlformats.org/package/2006/relationships",
        "table": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        "spreadsheetDrawing": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "drawing": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "drawingRelationship": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
        "image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "chart": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    };
    return Util;
}());
module.exports = Util;
