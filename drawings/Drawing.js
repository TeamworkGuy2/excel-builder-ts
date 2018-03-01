"use strict";
var Util = require("../util/Util");
var AbsoluteAnchor = require("./AbsoluteAnchor");
var OneCellAnchor = require("./OneCellAnchor");
var TwoCellAnchor = require("./TwoCellAnchor");
/** This is mostly a global spot where all of the relationship managers can get and set
 * path information from/to.
 * @module Excel/Drawing
 */
var Drawing = /** @class */ (function () {
    /**
     * @constructor
     */
    function Drawing() {
        this.id = Util._uniqueId("Drawing");
    }
    /**
     * @param type can be "absoluteAnchor", "oneCellAnchor", or "twoCellAnchor".
     * @param config Shorthand - pass the created anchor coords that can normally be used to construct it.
     * @returns a cell anchor object
     */
    Drawing.prototype.createAnchor = function (type, config) {
        config = config || {};
        config.drawing = this;
        switch (type) {
            case "absoluteAnchor":
                this.anchor = new AbsoluteAnchor(config);
                break;
            case "oneCellAnchor":
                this.anchor = new OneCellAnchor(config);
                break;
            case "twoCellAnchor":
                this.anchor = new TwoCellAnchor(config);
                break;
        }
        return this.anchor;
    };
    return Drawing;
}());
module.exports = Drawing;
