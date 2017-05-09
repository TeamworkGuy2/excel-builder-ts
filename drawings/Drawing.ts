import Util = require("../Util");
import XmlDom = require("../XmlDom");
import Drawings = require("../Drawings");
import AbsoluteAnchor = require("./AbsoluteAnchor");
import OneCellAnchor = require("./OneCellAnchor");
import TwoCellAnchor = require("./TwoCellAnchor");

/** This is mostly a global spot where all of the relationship managers can get and set
 * path information from/to. 
 * @module Excel/Drawing
 */
abstract class Drawing implements Drawings.Drawing {
    anchor: Drawing.AnchorLike;
    id: string;


    /**
     * @constructor
     */
    constructor() {
        this.id = Util._uniqueId("Drawing");
    }


    /**
     * @param type can be "absoluteAnchor", "oneCellAnchor", or "twoCellAnchor". 
     * @param config Shorthand - pass the created anchor coords that can normally be used to construct it.
     * @returns a cell anchor object
     */
    public createAnchor(type: "absoluteAnchor" | "oneCellAnchor" | "twoCellAnchor", config?: { drawing?; from?: Util.OffsetConfig; to?: Util.OffsetConfig; } & Util.Pos): Drawing.AnchorLike {
        config = config || <any>{};
        config.drawing = this;
        switch (type) {
            case "absoluteAnchor":
                this.anchor = new AbsoluteAnchor(config);
                break;
            case "oneCellAnchor":
                this.anchor = new OneCellAnchor(config);
                break;
            case "twoCellAnchor":
                this.anchor = new TwoCellAnchor(<{ from: Util.OffsetConfig; to: Util.OffsetConfig; }><any>config);
                break;
        }
        return this.anchor;
    }


    public abstract setRelationshipId(rId: string): void;

    public abstract toXML(xmlDoc: XmlDom): XmlDom.NodeBase;

    public abstract getMediaData(): { id: string; schema?: string; };

    public abstract getMediaType(): string;

}

module Drawing {

    export type AnchorLike = AbsoluteAnchor | OneCellAnchor | TwoCellAnchor;

}

export = Drawing;