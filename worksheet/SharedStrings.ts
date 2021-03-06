import Util = require("../util/Util");
import XmlDom = require("../xml/XmlDom");

/**
 * @module Excel/SharedStrings
 */
class SharedStrings {
    id: string;
    strings: { [key: string]: number };
    stringArray: string[];


    constructor() {
        this.strings = {};
        this.stringArray = [];
        this.id = Util._uniqueId("SharedStrings");
    }


    /** Adds a string to the shared string file, and returns the ID of the
     * string which can be used to reference it in worksheets.
     * @param str the string to add
     * @returns int the string index
     */
    public addString(str: string): number {
        this.strings[str] = this.stringArray.length;
        this.stringArray[this.stringArray.length] = str;
        return this.strings[str];
    }


    public exportData() {
        return this.strings;
    }


    public toXML() {
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetml, "sst");
        var sharedStringTable = doc.documentElement;
        this.stringArray.reverse();
        var l = this.stringArray.length;
        sharedStringTable.setAttribute("count", l);
        sharedStringTable.setAttribute("uniqueCount", l);

        var template = doc.createElement("si");
        var templateValue = doc.createElement("t");
        templateValue.appendChild(doc.createTextNode("--placeholder--"));
        template.appendChild(templateValue);
        var strings = this.stringArray;

        while (l--) {
            var clone = template.cloneNode(true);
            (<XmlDom.XMLNode>(<XmlDom.XMLNode>clone.firstChild).firstChild).nodeValue = strings[l];
            sharedStringTable.appendChild(clone);
        }

        return doc;
    }

}

export = SharedStrings;