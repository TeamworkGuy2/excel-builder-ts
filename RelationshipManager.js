"use strict";
var Util = require("./Util");
var Paths = require("./Paths");
/**
 * @module Excel/RelationshipManager
 */
var RelationshipManager = (function () {
    function RelationshipManager() {
        this.Cctor = (function () {
            Util._uniqueId('rId'); //priming
        }());
        this.relations = {};
        this.lastId = 1;
    }
    RelationshipManager.prototype.importData = function (data) {
        this.relations = data.relations;
        this.lastId = data.lastId;
    };
    RelationshipManager.prototype.exportData = function () {
        return {
            relations: this.relations,
            lastId: this.lastId
        };
    };
    RelationshipManager.prototype.addRelation = function (object, type) {
        var newRelation = this.relations[object.id] = {
            id: Util._uniqueId('rId'),
            schema: Util.schemas[type]
        };
        return newRelation.id;
    };
    RelationshipManager.prototype.getRelationshipId = function (object) {
        return this.relations[object.id] ? this.relations[object.id].id : null;
    };
    RelationshipManager.prototype.toXML = function () {
        var doc = Util.createXmlDoc(Util.schemas.relationshipPackage, 'Relationships');
        var relationships = doc.documentElement;
        var rels = this.relations;
        Object.keys(rels).forEach(function (id) {
            var data = rels[id];
            var relationship = Util.createElement(doc, 'Relationship', [
                ['Id', data.id],
                ['Type', data.schema],
                ['Target', Paths[id]]
            ]);
            relationships.appendChild(relationship);
        });
        return doc;
    };
    return RelationshipManager;
}());
module.exports = RelationshipManager;
