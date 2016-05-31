import Util = require("../Util");

class AbsoluteAnchor {
    x: number;
    y: number;
    width: number;
    height: number;


    /**
     * @param {Object} config
     * @param {number} config.x X offset in EMU's
     * @param {number} config.y Y offset in EMU's
     * @param {number} config.width Width in EMU's
     * @param {number} config.height Height in EMU's
     * @constructor
     */
    constructor(config?: Util.Pos) {
        this.x = null;
        this.y = null;
        this.width = null;
        this.height = null;
        if (config) {
            this.setPos(config.x, config.y);
            this.setDimensions(config.width, config.height);
        }
    }


    /** Sets the X and Y offsets.
     * 
     * @param {number} x
     * @param {number} y
     */
    public setPos(x: number, y: number) {
        this.x = x;
        this.y = y;
    }


    /** Sets the width and height of the image.
     * 
     * @param {number} width
     * @param {number} height
     */
    public setDimensions(width: number, height: number) {
        this.width = width;
        this.height = height;
    }


    public toXML(xmlDoc: XMLDocument, content: Node) {
        var root = Util.createElement(xmlDoc, "xdr:absoluteAnchor");
        var pos = Util.createElement(xmlDoc, "xdr:pos");
        pos.setAttribute("x", <any>this.x);
        pos.setAttribute("y", <any>this.y);
        root.appendChild(pos);

        var dimensions = Util.createElement(xmlDoc, "xdr:ext");
        dimensions.setAttribute("cx", <any>this.width);
        dimensions.setAttribute("cy", <any>this.height);
        root.appendChild(dimensions);

        root.appendChild(content);

        root.appendChild(Util.createElement(xmlDoc, "xdr:clientData"));
        return root;
    }

}

export = AbsoluteAnchor;