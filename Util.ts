import XmlDom = require("./XmlDom");

/**
 * @module Excel util
 */
class Util {
    static _idSpaces: { [id: string]: number } = {};
    static _id = 0;


    static _uniqueId(space: string) {
        var id = ++Util._id;
        return space + id;
    }


    /**
     * Returns a number based on a namespace. So, running with 'Picture' will return 1. Run again, you will get 2. Run with 'Foo', you'll get 1.
     * @param space
     * @returns a unique ID identifying the string
     */
    static uniqueId(space: string): number {
        if (!this._idSpaces[space]) {
            this._idSpaces[space] = 1;
        }
        return this._idSpaces[space]++;
    }


    static pick<T extends object, K extends keyof T>(obj: T, props: K[]): { [P in K]: T[P] } {
        var res = {};
        for (var i = 0, size = props.length; i < size; i++) {
            var key = props[i];
            if (key in obj) {
                (<any>res)[key] = obj[key];
            }
        }
        return <any>res;
    }


    static defaults<T1, T2>(obj: T1, overrides: T2): T1 & T2 {
        for (var key in overrides) {
            if (overrides.hasOwnProperty(key) && (obj[<string>key] === undefined)) {
                obj[<string>key] = overrides[key];
            }
        }
        return <any>obj;
    }


    /**
     * Attempts to create an XML document. Due to limitations in web workers, 
     * it may return a 'fake' xml document created from the XMLDOM.js file. 
     * 
     * Takes a namespace to start the xml file in, as well as the root element
     * of the xml file. 
     * 
     * @param ns a namespace string
     * @param base node name
     * @returns document.implementation.createDocument() | new XmlDom()
     */
    static createXmlDoc(ns: string, base: string): XmlDom {
        if (typeof document === "undefined") {
            return new XmlDom(ns || null, base, null);
        }
        else if (document.implementation && document.implementation.createDocument) {
            return <any>document.implementation.createDocument(ns || null, base, null);
        }
        throw new Error("No xml document generator");
    }


    /**
     * Creates an xml node (element). Used to simplify some calls, as IE is
     * very particular about namespaces and such. 
     * 
     * @param doc An xml document (actual DOM or fake DOM, not a string)
     * @param name The name of the element
     * @param attributes
     * @returns ElementLike implementation
     */
    static createElement<E extends Util.ElementLike>(doc: { createElement(tagName: string): E; }, name: string, attributes?: [string, string | number][]): E {
        var el = doc.createElement(name);
        var ie = !el.setAttributeNS;
        attributes = attributes || [];
        var i = attributes.length;
        while (i--) {
            if (!ie && attributes[i][0].indexOf("xmlns") != -1) {
                el.setAttributeNS("http://www.w3.org/2000/xmlns/", attributes[i][0], <any>attributes[i][1]);
            } else {
                el.setAttribute(attributes[i][0], <any>attributes[i][1]);
            }
        }
        return el;
    }
        

    static LETTER_REFS: { [id: string]: string } = {};


    static positionToLetterRef(x: number, y: number) {
        var digit = 1;
        var num = x;
        var str = "";
        var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        if (Util.LETTER_REFS[x]) {
            return Util.LETTER_REFS[x].concat(<any>y);
        }

        var idx: number;
        while (num > 0) {
            num -= Math.pow(26, digit - 1);
            idx = num % Math.pow(26, digit);
            num -= idx;
            idx = idx / Math.pow(26, digit - 1);
            str = alphabet.charAt(idx) + str;
            digit += 1;
        }
        Util.LETTER_REFS[x] = str;
        return str.concat(<any>y);
    }

		
    static schemas = {
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

}

module Util {

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

}

export = Util;