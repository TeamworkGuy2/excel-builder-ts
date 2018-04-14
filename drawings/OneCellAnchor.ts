import Util = require("../util/Util");
import XmlDom = require("../xml/XmlDom");

class OneCellAnchor {
    x: number | null;
    y: number | null;
    xOff: number | null;
    yOff: number | null;
    width: number | null;
    height: number | null;


    /**
     * @param config
     * config.x The cell column number that the top left of the picture will start in
     * config.y The cell row number that the top left of the picture will start in
     * config.width Width in EMU's
     * config.height Height in EMU's
     * @constructor
     */
    constructor(config?: Util.Pos & { xOff?: number; yOff?: number; }) {
        this.x = null;
        this.y = null;
        this.xOff = null;
        this.yOff = null;
        this.width = null;
        this.height = null;
        if (config) {
            this.setPos(config.x, config.y, config.xOff, config.yOff);
            this.setDimensions(config.width, config.height);
        }
    }


    public setPos(x: number, y: number, xOff: number | undefined, yOff: number | undefined) {
        this.x = x;
        this.y = y;
        if (xOff !== undefined) {
            this.xOff = xOff;
        }
        if (yOff !== undefined) {
            this.yOff = yOff;
        }
    }


    public setDimensions(width: number, height: number) {
        this.width = width;
        this.height = height;
    }


    public toXML(xmlDoc: XmlDom, content: XmlDom.NodeBase) {
        var root = Util.createElement(xmlDoc, "xdr:oneCellAnchor");
        var from = Util.createElement(xmlDoc, "xdr:from");
        var fromCol = Util.createElement(xmlDoc, "xdr:col");
        fromCol.appendChild(xmlDoc.createTextNode(<any>this.x));
        var fromColOff = Util.createElement(xmlDoc, "xdr:colOff");
        fromColOff.appendChild(xmlDoc.createTextNode(<any>this.xOff || 0));
        var fromRow = Util.createElement(xmlDoc, "xdr:row");
        fromRow.appendChild(xmlDoc.createTextNode(<any>this.y));
        var fromRowOff = Util.createElement(xmlDoc, "xdr:rowOff");
        fromRowOff.appendChild(xmlDoc.createTextNode(<any>this.yOff || 0));
        from.appendChild(fromCol);
        from.appendChild(fromColOff);
        from.appendChild(fromRow);
        from.appendChild(fromRowOff);

        root.appendChild(from);

        var dimensions = Util.createElement(xmlDoc, "xdr:ext");
        dimensions.setAttribute("cx", <any>this.width);
        dimensions.setAttribute("cy", <any>this.height);
        root.appendChild(dimensions);

        root.appendChild(content);

        root.appendChild(Util.createElement(xmlDoc, "xdr:clientData"));
        return root;
    }

}

export = OneCellAnchor;