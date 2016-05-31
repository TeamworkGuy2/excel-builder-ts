import Util = require("./Util");
import RelationshipManager = require("./RelationshipManager");
import XmlDom = require("./XmlDom");

/**
 * @module Excel/Drawings
 */
class Drawings {
    drawings: Drawings.Drawing[];
    relations: RelationshipManager;
    id: any;


    constructor() {
        this.drawings = [];
        this.relations = new RelationshipManager();
        this.id = Util._uniqueId("Drawings");
    }


    /**
     * Adds a drawing (more likely a subclass of a Drawing) to the 'Drawings' for a particular worksheet.
     * 
     * @param {Drawing} drawing
     */
    public addDrawing(drawing: Drawings.Drawing) {
        this.drawings.push(drawing);
    }


    public getCount() {
        return this.drawings.length;
    }


    public toXML() {
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
    }

}

module Drawings {

    export interface Drawing {
        getMediaData();
        getMediaType();
        setRelationshipId(rId: string): void;
        toXML(doc: XmlDom);
    }

}

export = Drawings;