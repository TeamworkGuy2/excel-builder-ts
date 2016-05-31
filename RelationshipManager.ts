import Util = require("./Util");
import Paths = require("./Paths");

/**
 * @module Excel/RelationshipManager
 */
class RelationshipManager {

    private Cctor = (function () {
        Util._uniqueId("rId"); //priming
    }());


    relations: { [id: string]: { id: string; schema: string; } };
    lastId: number;


    constructor() {
        this.relations = <any>{};
        this.lastId = 1;
    }


    public importData(data: { relations: any; lastId: number; }) {
        this.relations = data.relations;
        this.lastId = data.lastId;
    }


    public exportData(): RelationshipManager.ExportData {
        return {
            relations: this.relations,
            lastId: this.lastId
        };
    }


    public addRelation(object: { id: string; schema?: string; }, type: string) {
        var newRelation = this.relations[object.id] = {
            id: Util._uniqueId("rId"),
            schema: Util.schemas[type]
        };
        return newRelation.id;
    }


    public getRelationshipId(object: { id: string; schema?: string; }) {
        return this.relations[object.id] ? this.relations[object.id].id : null;
    }


    public toXML() {
        var doc = Util.createXmlDoc(Util.schemas.relationshipPackage, "Relationships");
        var relationships = doc.documentElement;

        var rels = this.relations;
        Object.keys(rels).forEach((id) => {
            var data = rels[id];
            var relationship = Util.createElement(doc, "Relationship", [
                ["Id", data.id],
                ["Type", data.schema],
                ["Target", Paths[id]]
            ]);
            relationships.appendChild(relationship);
        });
        return doc;
    }

}

module RelationshipManager {

    export interface ExportData {
        relations: { [id: string]: { id: string; schema: string; } };
        lastId: number;
    }

}

export = RelationshipManager;