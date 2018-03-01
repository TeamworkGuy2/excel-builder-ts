import Util = require("../util/Util");
import Paths = require("./Paths");

/**
 * @module Excel/RelationshipManager
 */
class RelationshipManager {
    private Cctor = (function () {
        Util._uniqueId("rId"); //priming
    }());

    relations: { [id: string]: RelationshipManager.Relation };
    lastId: number;


    constructor() {
        this.relations = <any>{};
        this.lastId = 1;
    }


    public importData(data: RelationshipManager.ExportData) {
        this.relations = data.relations;
        this.lastId = data.lastId;
    }


    public exportData(): RelationshipManager.ExportData {
        return {
            relations: this.relations,
            lastId: this.lastId
        };
    }


    public addRelation(obj: { id: string; schema?: Util.SchemaName; }, type: Util.SchemaName) {
        var newRelation = this.relations[obj.id] = {
            id: Util._uniqueId("rId"),
            schema: Util.schemas[type]
        };
        return newRelation.id;
    }


    public getRelationshipId(obj: { id: string; schema?: string; }) {
        return this.relations[obj.id] ? this.relations[obj.id].id : null;
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

    export interface Relation {
        id: string;
        schema: string;
    }


    export interface ExportData {
        relations: { [id: string]: Relation };
        lastId: number;
    }

}

export = RelationshipManager;