"use strict";
var Util = require("./Util");
var RelationshipManager = require("./RelationshipManager");
/**
 * @module Excel/Drawings
 */
var Drawings = (function () {
    function Drawings() {
        this.drawings = [];
        this.relations = new RelationshipManager();
        this.id = Util._uniqueId("Drawings");
    }
    /**
     * Adds a drawing (more likely a subclass of a Drawing) to the 'Drawings' for a particular worksheet.
     *
     * @param {Drawing} drawing
     */
    Drawings.prototype.addDrawing = function (drawing) {
        this.drawings.push(drawing);
    };
    Drawings.prototype.getCount = function () {
        return this.drawings.length;
    };
    Drawings.prototype.toXML = function () {
        var doc = Util.createXmlDoc(Util.schemas.spreadsheetDrawing, "xdr:wsDr");
        var drawingsElem = doc.documentElement;
        //drawings.setAttribute('xmlns:xdr', util.schemas.spreadsheetDrawing);
        drawingsElem.setAttribute("xmlns:a", Util.schemas.drawing);
        var existingRelationships = {};
        for (var i = 0, l = this.drawings.length; i < l; i++) {
            var rId = this.relations.getRelationshipId(this.drawings[i].getMediaData());
            if (!rId) {
                rId = this.relations.addRelation(this.drawings[i].getMediaData(), this.drawings[i].getMediaType()); //chart
            }
            this.drawings[i].setRelationshipId(rId);
            drawingsElem.appendChild(this.drawings[i].toXML(doc));
        }
        return doc;
    };
    return Drawings;
}());
module.exports = Drawings;
