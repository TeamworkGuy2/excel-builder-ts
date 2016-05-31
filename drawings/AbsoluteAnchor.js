"use strict";
var Util = require("../Util");
var AbsoluteAnchor = (function () {
    /**
     * @param {Object} config
     * @param {number} config.x X offset in EMU's
     * @param {number} config.y Y offset in EMU's
     * @param {number} config.width Width in EMU's
     * @param {number} config.height Height in EMU's
     * @constructor
     */
    function AbsoluteAnchor(config) {
        this.x = null;
        this.y = null;
        this.width = null;
        this.height = null;
        if (config) {
            this.setPos(config.x, config.y);
            this.setDimensions(config.width, config.height);
        }
    }
    /** Sets the X and Y offsets.
     *
     * @param {number} x
     * @param {number} y
     */
    AbsoluteAnchor.prototype.setPos = function (x, y) {
        this.x = x;
        this.y = y;
    };
    /** Sets the width and height of the image.
     *
     * @param {number} width
     * @param {number} height
     */
    AbsoluteAnchor.prototype.setDimensions = function (width, height) {
        this.width = width;
        this.height = height;
    };
    AbsoluteAnchor.prototype.toXML = function (xmlDoc, content) {
        var root = Util.createElement(xmlDoc, "xdr:absoluteAnchor");
        var pos = Util.createElement(xmlDoc, "xdr:pos");
        pos.setAttribute("x", this.x);
        pos.setAttribute("y", this.y);
        root.appendChild(pos);
        var dimensions = Util.createElement(xmlDoc, "xdr:ext");
        dimensions.setAttribute("cx", this.width);
        dimensions.setAttribute("cy", this.height);
        root.appendChild(dimensions);
        root.appendChild(content);
        root.appendChild(Util.createElement(xmlDoc, "xdr:clientData"));
        return root;
    };
    return AbsoluteAnchor;
}());
module.exports = AbsoluteAnchor;
