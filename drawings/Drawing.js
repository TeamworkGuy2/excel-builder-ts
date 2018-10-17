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
        this.anchor = null;
        this.id = Util._uniqueId("Drawing");
    }
    /**
     * @param type can be "absoluteAnchor", "oneCellAnchor", or "twoCellAnchor".
     * @param config Shorthand - pass the created anchor coords that can normally be used to construct it.
     * @returns a cell anchor object
     */
    Drawing.prototype.createAnchor = function (type, config) {
        var cfg = (config != null ? config : {});
        cfg.drawing = this;
        switch (type) {
            case "absoluteAnchor":
                return this.anchor = new AbsoluteAnchor(cfg);
            case "oneCellAnchor":
                return this.anchor = new OneCellAnchor(cfg);
            case "twoCellAnchor":
                return this.anchor = new TwoCellAnchor(cfg);
        }
    };
    return Drawing;
}());
module.exports = Drawing;
