"use strict";
var Util = require("../Util");
var OneCellAnchor = (function () {
    /**
     * @param {Object} config
     * @param {number} config.x The cell column number that the top left of the picture will start in
     * @param {number} config.y The cell row number that the top left of the picture will start in
     * @param {number} config.width Width in EMU's
     * @param {number} config.height Height in EMU's
     * @constructor
     */
    function OneCellAnchor(config) {
        this.x = null;
        this.y = null;
        this.xOff = null;
        this.yOff = null;
        this.width = null;
        this.height = null;
        if (config) {
            this.setPos(config.x, config.y, config.xOff, config.yOff);
            this.setDimensions(config.width, config.height);
        }
    }
    OneCellAnchor.prototype.setPos = function (x, y, xOff, yOff) {
        this.x = x;
        this.y = y;
        if (xOff !== undefined) {
            this.xOff = xOff;
        }
        if (yOff !== undefined) {
            this.yOff = yOff;
        }
    };
    OneCellAnchor.prototype.setDimensions = function (width, height) {
        this.width = width;
        this.height = height;
    };
    OneCellAnchor.prototype.toXML = function (xmlDoc, content) {
        var root = Util.createElement(xmlDoc, "xdr:oneCellAnchor");
        var from = Util.createElement(xmlDoc, "xdr:from");
        var fromCol = Util.createElement(xmlDoc, "xdr:col");
        fromCol.appendChild(xmlDoc.createTextNode(this.x));
        var fromColOff = Util.createElement(xmlDoc, "xdr:colOff");
        fromColOff.appendChild(xmlDoc.createTextNode(this.xOff || 0));
        var fromRow = Util.createElement(xmlDoc, "xdr:row");
        fromRow.appendChild(xmlDoc.createTextNode(this.y));
        var fromRowOff = Util.createElement(xmlDoc, "xdr:rowOff");
        fromRowOff.appendChild(xmlDoc.createTextNode(this.yOff || 0));
        from.appendChild(fromCol);
        from.appendChild(fromColOff);
        from.appendChild(fromRow);
        from.appendChild(fromRowOff);
        root.appendChild(from);
        var dimensions = Util.createElement(xmlDoc, "xdr:ext");
        dimensions.setAttribute("cx", this.width);
        dimensions.setAttribute("cy", this.height);
        root.appendChild(dimensions);
        root.appendChild(content);
        root.appendChild(Util.createElement(xmlDoc, "xdr:clientData"));
        return root;
    };
    return OneCellAnchor;
}());
module.exports = OneCellAnchor;
