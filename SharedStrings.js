"use strict";
var Util = require("./Util");
/**
 * @module Excel/SharedStrings
 */
var SharedStrings = (function () {
    function SharedStrings() {
        this.strings = {};
        this.stringArray = [];
        this.id = Util._uniqueId("SharedStrings");
    }
    /** Adds a string to the shared string file, and returns the ID of the
     * string which can be used to reference it in worksheets.
     *
     * @param string {String}
     * @return int
     */
    SharedStrings.prototype.addString = function (str) {
        this.strings[str] = this.stringArray.length;
        this.stringArray[this.stringArray.length] = str;
        return this.strings[str];
    };
    SharedStrings.prototype.exportData = function () {
        return this.strings;
    };
    SharedStrings.prototype.toXML = function () {
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
            clone.firstChild.firstChild.nodeValue = strings[l];
            sharedStringTable.appendChild(clone);
        }
        return doc;
    };
    return SharedStrings;
}());
module.exports = SharedStrings;
