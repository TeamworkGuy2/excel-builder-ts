import Util = require("../Util");
import XmlDom = require("../XmlDom");
import Drawing = require("./Drawing");
import Drawings = require("../Drawings");

class Picture {
    static Cctor = (function () {
        var thisProto = Picture.prototype;
        Picture.prototype = new (<any>Drawing)();
        Object.assign(Picture.prototype, thisProto);
    } ());

    anchor: Drawing.AnchorLike;
    description: string;
    fill: any;
    id: string;
    media: any;
    mediaData: { rId?: string; id: string; fileName: string; };
    pictureId: number;


    constructor() {
        this.media = null;
        this.id = Util._uniqueId("Picture");
        this.pictureId = Util.uniqueId("Picture");
        this.fill = {};
        this.mediaData = null;
    }


    public setMedia(mediaRef: { rId?: string; id: string; fileName: string; [id: string]: any }) {
        this.mediaData = mediaRef;
    }


    public setDescription(description: string) {
        this.description = description;
    }


    public setFillType(type: any) {
        this.fill.type = type;
    }


    public setFillConfig(config: any) {
        Util.defaults(this.fill, config);
    }


    public getMediaType() {
        return "image";
    }


    public getMediaData(): { id: string; schema?: string; } {
        return this.mediaData;
    }


    public setRelationshipId(rId: string) {
        this.mediaData.rId = rId;
    }


    public toXML(xmlDoc: XmlDom): XmlDom.NodeBase {
        var pictureNode = Util.createElement(xmlDoc, "xdr:pic");

        var nonVisibleProperties = Util.createElement(xmlDoc, "xdr:nvPicPr");

        var nameProperties = Util.createElement(xmlDoc, "xdr:cNvPr", [
            ["id", this.pictureId],
            ["name", this.mediaData.fileName],
            ["descr", this.description || ""]
        ]);
        nonVisibleProperties.appendChild(nameProperties);
        var nvPicProperties = Util.createElement(xmlDoc, "xdr:cNvPicPr");
        nvPicProperties.appendChild(Util.createElement(xmlDoc, "a:picLocks", [
            ["noChangeAspect", '1'],
            ["noChangeArrowheads", '1']
        ]));
        nonVisibleProperties.appendChild(nvPicProperties);
        pictureNode.appendChild(nonVisibleProperties);
        var pictureFill = Util.createElement(xmlDoc, "xdr:blipFill");
        pictureFill.appendChild(Util.createElement(xmlDoc, "a:blip", [
            ["xmlns:r", Util.schemas.relationships],
            ["r:embed", this.mediaData.rId]
        ]));
        pictureFill.appendChild(Util.createElement(xmlDoc, "a:srcRect"));
        var stretch = Util.createElement(xmlDoc, "a:stretch");
        stretch.appendChild(Util.createElement(xmlDoc, "a:fillRect"));
        pictureFill.appendChild(stretch);
        pictureNode.appendChild(pictureFill);

        var shapeProperties = Util.createElement(xmlDoc, "xdr:spPr", [
            ["bwMode", "auto"]
        ]);

        var transform2d = Util.createElement(xmlDoc, "a:xfrm");
        shapeProperties.appendChild(transform2d);

        var presetGeometry = Util.createElement(xmlDoc, "a:prstGeom", [
            ["prst", "rect"]
        ]);
        shapeProperties.appendChild(presetGeometry);

        pictureNode.appendChild(shapeProperties);
        //     <xdr:spPr bwMode="auto">
        //         <a:xfrm>
        //             <a:off x="1" y="1"/>
        //             <a:ext cx="1640253" cy="1885949"/>
        //         </a:xfrm>
        //         <a:prstGeom prst="rect">
        //             <a:avLst/>
        //         </a:prstGeom>
        //         <a:noFill/>
        //         <a:extLst>
        //             <a:ext uri="{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}">
        //                 <a14:hiddenFill xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
        //                     <a:solidFill>
        //                         <a:srgbClr val="FFFFFF"/>
        //                     </a:solidFill>
        //                 </a14:hiddenFill>
        //             </a:ext>
        //         </a:extLst>
        //     </xdr:spPr>

        return this.anchor.toXML(xmlDoc, pictureNode);
    }

}

interface PictureDrawing extends Picture, Drawing, Drawings.Drawing { }
export = <{ new (): PictureDrawing }><any>Picture;