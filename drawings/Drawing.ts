import Util = require("../Util");
import AbsoluteAnchor = require("./AbsoluteAnchor");
import OneCellAnchor = require("./OneCellAnchor");
import TwoCellAnchor = require("./TwoCellAnchor");


type AnchorLike = AbsoluteAnchor | OneCellAnchor | TwoCellAnchor;


/** This is mostly a global spot where all of the relationship managers can get and set
 * path information from/to. 
 * @module Excel/Drawing
 */
class Drawing {
    anchor: AnchorLike;
    id: string;


    /**
     * @constructor
     */
    constructor() {
        this.id = Util._uniqueId('Drawing');
    }


    /**
     * @param {String} type Can be 'absoluteAnchor', 'oneCellAnchor', or 'twoCellAnchor'. 
     * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
     * @returns {Anchor}
     */
    public createAnchor(type: 'absoluteAnchor' | 'oneCellAnchor' | 'twoCellAnchor', config?: { drawing?; from?: Util.OffsetConfig; to?: Util.OffsetConfig; } & Util.Pos) {
        config = config || <any>{};
        config.drawing = this;
        switch (type) {
            case 'absoluteAnchor':
                this.anchor = new AbsoluteAnchor(config);
                break;
            case 'oneCellAnchor':
                this.anchor = new OneCellAnchor(config);
                break;
            case 'twoCellAnchor':
                this.anchor = new TwoCellAnchor(<{ from: Util.OffsetConfig; to: Util.OffsetConfig; }><any>config);
                break;
        }
        return this.anchor;
    }

}

export = Drawing;