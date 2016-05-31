import Util = require("../Util");

class TwoCellAnchor {
    from: Util.OffsetConfig;
    to: Util.OffsetConfig;


    constructor(config?: { from: Util.OffsetConfig; to: Util.OffsetConfig; }) {
        this.from = { xOff: 0, yOff: 0 };
        this.to = { xOff: 0, yOff: 0 };
        if (config) {
            this.setFrom(config.from.x, config.from.y, config.to.xOff, config.to.yOff);
            this.setTo(config.to.x, config.to.y, config.to.xOff, config.to.yOff);
        }
    }


    public setFrom(x: number, y: number, xOff: number, yOff: number) {
        this.from.x = x;
        this.from.y = y;
        if (xOff !== undefined) {
            this.from.xOff = xOff;
        }
        if (yOff !== undefined) {
            this.from.yOff = xOff;
        }
    }


    public setTo(x: number, y: number, xOff: number, yOff: number) {
        this.to.x = x;
        this.to.y = y;
        if (xOff !== undefined) {
            this.to.xOff = xOff;
        }
        if (yOff !== undefined) {
            this.to.yOff = xOff;
        }
    }


    public toXML(xmlDoc: XMLDocument, content: Node) {
        var root = Util.createElement(xmlDoc, "xdr:twoCellAnchor");

        var from = Util.createElement(xmlDoc, "xdr:from");
        var fromCol = Util.createElement(xmlDoc, "xdr:col");
        fromCol.appendChild(xmlDoc.createTextNode(<any>this.from.x));
        var fromColOff = Util.createElement(xmlDoc, "xdr:colOff");
        fromColOff.appendChild(xmlDoc.createTextNode(<any>this.from.xOff));
        var fromRow = Util.createElement(xmlDoc, "xdr:row");
        fromRow.appendChild(xmlDoc.createTextNode(<any>this.from.y));
        var fromRowOff = Util.createElement(xmlDoc, "xdr:rowOff");
        fromRowOff.appendChild(xmlDoc.createTextNode(<any>this.from.yOff));

        from.appendChild(fromCol);
        from.appendChild(fromColOff);
        from.appendChild(fromRow);
        from.appendChild(fromRowOff);

        var to = Util.createElement(xmlDoc, "xdr:to");
        var toCol = Util.createElement(xmlDoc, "xdr:col");
        toCol.appendChild(xmlDoc.createTextNode(<any>this.to.x));
        var toColOff = Util.createElement(xmlDoc, "xdr:colOff");
        toColOff.appendChild(xmlDoc.createTextNode(<any>this.from.xOff));
        var toRow = Util.createElement(xmlDoc, "xdr:row");
        toRow.appendChild(xmlDoc.createTextNode(<any>this.to.y));
        var toRowOff = Util.createElement(xmlDoc, "xdr:rowOff");
        toRowOff.appendChild(xmlDoc.createTextNode(<any>this.from.yOff));

        to.appendChild(toCol);
        to.appendChild(toColOff);
        to.appendChild(toRow);
        to.appendChild(toRowOff);


        root.appendChild(from);
        root.appendChild(to);

        root.appendChild(content);

        root.appendChild(Util.createElement(xmlDoc, "xdr:clientData"));
        return root;
    }

}

export = TwoCellAnchor;