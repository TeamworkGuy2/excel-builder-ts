import Util = require("../util/Util");
import RelationshipManager = require("../worksheet/RelationshipManager");
import XmlDom = require("../xml/XmlDom");

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


    /** Adds a drawing (more likely a subclass of a Drawing) to the 'Drawings' for a particular worksheet.
     * 
     * @param drawing
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

        for (var i = 0, l = this.drawings.length; i < l; i++) {
            var drwI = this.drawings[i];
            var rId = this.relations.getRelationshipId(drwI.getMediaData());
            if (!rId) {
                rId = this.relations.addRelation(drwI.getMediaData(), drwI.getMediaType()); //chart
            }
            drwI.setRelationshipId(rId);
            drawingsElem.appendChild(drwI.toXML(doc));
        }
        return doc;
    }

}

module Drawings {

    export interface Drawing {
        getMediaData(): { id: string; schema?: Util.SchemaName; };
        getMediaType(): Util.SchemaName;
        setRelationshipId(rId: string): void;
        toXML(doc: XmlDom): XmlDom.NodeBase;
    }

}

export = Drawings;