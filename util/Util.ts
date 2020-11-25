import XmlDom = require("../xml/XmlDom");

/**
 * @module Excel util
 */
module Util {

    export type SchemaName = keyof typeof Util.schemas;

    export interface Pos {
        x: number;
        y: number;
        width: number;
        height: number;
    }

    export interface OffsetConfig {
        x?: number;
        y?: number;
        xOff: number;
        yOff: number;
    }

    export interface ElementLike {
        setAttributeNS(ns: string, name: string, value: string): void;
        setAttribute(name: string, value: string): void;
    }


    export var schemas = {
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
        "chart": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    };

    var _idSpaces: { [id: string]: number } = { };
    var _id = 0;


    export function _uniqueId(space: string) {
        var id = ++_id;
        return space + id;
    }


    /** Returns a number based on a namespace. So, running with 'Picture' will return 1. Run again, you will get 2. Run with 'Foo', you'll get 1.
     * @param space
     * @returns a unique ID identifying the string
     */
    export function uniqueId(space: string): number {
        if (!_idSpaces[space]) {
            _idSpaces[space] = 1;
        }
        return _idSpaces[space]++;
    }


    export function pick<T extends object, K extends keyof T>(obj: T, props: K[]): { [P in K]: T[P] } {
        var res = <{ [P in K]: T[P] }><any>{};
        for (var i = 0, size = props.length; i < size; i++) {
            var key = props[i];
            if (key in obj) {
                res[key] = obj[key];
            }
        }
        return res;
    }


    export function defaults<T1 extends object, T2 extends object>(obj: T1, overrides: T2): T1 & T2 {
        for (var key in overrides) {
            if (overrides.hasOwnProperty(key) && (obj[<keyof T1><string>key] === undefined)) {
                obj[<keyof T1><string>key] = <any>overrides[key];
            }
        }
        return <any>obj;
    }


    /** Attempts to create an XML document. Due to limitations in web workers, 
     * it may return a 'fake' xml document created from the XmlDom.js file.
     * 
     * Takes a namespace to start the xml file in, as well as the root element
     * of the xml file.
     * 
     * @param ns a namespace string
     * @param base node name
     * @returns document.implementation.createDocument() | new XmlDom()
     */
    export function createXmlDoc(ns: string, base: string): XmlDom {
        if (typeof document === "undefined") {
            return new XmlDom(ns || null, base, null);
        }
        else if (document.implementation && document.implementation.createDocument) {
            return <any>document.implementation.createDocument(ns || null, base, null);
        }
        throw new Error("No XML document generator");
    }


    /** Creates an xml node (element). Used to simplify some calls, as IE is
     * very particular about namespaces and such. 
     * 
     * @param doc An xml document (actual DOM or fake DOM, not a string)
     * @param name The name of the element
     * @param attributes
     * @returns ElementLike implementation
     */
    export function createElement<E extends Util.ElementLike>(doc: { createElement(tagName: string): E; }, name: string, attributes?: [string, string | number][]): E {
        var el = doc.createElement(name);
        var ie = !el.setAttributeNS;
        attributes = attributes || [];
        var i = attributes.length;
        while (i--) {
            var attr = attributes[i];
            if (!ie && attr[0].indexOf("xmlns") != -1) {
                el.setAttributeNS("http://www.w3.org/2000/xmlns/", attr[0], <any>attr[1]);
            }
            else {
                el.setAttribute(attr[0], <any>attr[1]);
            }
        }
        return el;
    }
        

    var LETTER_REFS: { [id: string]: string } = {};


    /** Convert two numbers representing a column and a row into an 'A1' style spreadsheet location reference.
     * Example: positionToLetterRef(34, 5) returns AH5
     * @param x the value to convert to letter(s)
     * @param y the row number to append to the letters
     * @returns a string starting with 'x' converted to letters and the string value of 'y' appended
     */
    export function positionToLetterRef(x: number, y: number): string {
        var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        // check for precalculated value in cache
        if (LETTER_REFS[x]) {
            return LETTER_REFS[x] + y;
        }

        var num = x;
        var str = "";
        var digitPos = 1;
        var prevPow = 1; // Math.pow(26, digitPos - 1) = 1 for the 1st digit's scale

        // tricky but cleaver algorithm, credit to stephenliberty's excel-builder.js original: https://github.com/stephenliberty/excel-builder.js/commit/b23cf0384a18bcfbfb25a0a9556636e2d5fb4b04
        while (num > 0) {
            num -= prevPow; // this works for the first digit because the 'alphabet' of symbols is base-0 so subtracting 1 on the first loop makes the 'charAt()' lookup work, subsequent loops work the same way
            var digitPow = Math.pow(26, digitPos); // the current digit's scale, in this case alphabetic digits: 26, 676, 17576 (e.g. 10 for the 1st digit base ten)
            var remainder = num % digitPow; // remainder of dividing the number by the current digit's scale (e.g. '6' in 136 for the 1st digit base ten)
            num -= remainder; // subtract the remainder so the number is evenly divisible by the current digit's scale
            var digit = remainder / prevPow; // divide the remainder by the previous digit's scale (e.g. '6' when dividing 6 by 1 for the 1st digit base ten)
            prevPow = digitPow; // for the next loop to reduce Math.pow() calls
            str = alphabet.charAt(digit) + str; // convert the digit to it's symbol and prepend since we're dividing and moving up from the least significant digit each loop
            digitPos++; // next digit
        }
        LETTER_REFS[x] = str; // save calculated results in a cache for reuse, could be an issue if the user generates 10's of thousands of unique letter positions

        return str + y;
    }

}

export = Util;