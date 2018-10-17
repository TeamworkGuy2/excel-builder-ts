import Util = require("../util/Util");
import XmlDom = require("../xml/XmlDom");
import Drawing = require("./Drawing");
import Drawings = require("./Drawings");

class Picture implements Drawings.Drawing {
    static Cctor = (function () {
        var thisProto = Picture.prototype;
        Picture.prototype = new (<any>Drawing)();
        Object.assign(Picture.prototype, thisProto);
    } ());

    anchor: Drawing.AnchorLike | null = null;
    description: string | null;
    fill: any;
    id: string;
    media: any;
    mediaData: { rId?: string; id: string; fileName: string; };
    pictureId: number;


    constructor() {
        this.description = null;
        this.media = null;
        this.id = Util._uniqueId("Picture");
        this.pictureId = Util.uniqueId("Picture");
        this.fill = {};
        this.mediaData = <any>null;
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


    public getMediaType(): Util.SchemaName {
        return "image";
    }


    public getMediaData(): { id: string; schema?: Util.SchemaName; } {
        return this.mediaData;
    }


    public setRelationshipId(rId: string) {
        this.mediaData.rId = rId;
    }


    public toXML(xmlDoc: XmlDom): XmlDom.NodeBase {
        var pictureNode = Util.createElement(xmlDoc, "xdr:pic");

        var nonVisibleProps = Util.createElement(xmlDoc, "xdr:nvPicPr");

        var nameProps = Util.createElement(xmlDoc, "xdr:cNvPr", [
            ["id", this.pictureId],
            ["name", this.mediaData.fileName],
            ["descr", this.description || ""]
        ]);
        nonVisibleProps.appendChild(nameProps);
        var nvPicProps = Util.createElement(xmlDoc, "xdr:cNvPicPr");
        nvPicProps.appendChild(Util.createElement(xmlDoc, "a:picLocks", [
            ["noChangeAspect", '1'],
            ["noChangeArrowheads", '1']
        ]));
        nonVisibleProps.appendChild(nvPicProps);
        pictureNode.appendChild(nonVisibleProps);

        var pictureFill = Util.createElement(xmlDoc, "xdr:blipFill");
        pictureFill.appendChild(Util.createElement(xmlDoc, "a:blip", [
            ["xmlns:r", Util.schemas.relationships],
            ["r:embed", <any>this.mediaData.rId]
        ]));
        pictureFill.appendChild(Util.createElement(xmlDoc, "a:srcRect"));
        var stretch = Util.createElement(xmlDoc, "a:stretch");
        stretch.appendChild(Util.createElement(xmlDoc, "a:fillRect"));
        pictureFill.appendChild(stretch);
        pictureNode.appendChild(pictureFill);

        var shapeProps = Util.createElement(xmlDoc, "xdr:spPr", [
            ["bwMode", "auto"]
        ]);

        var transform2d = Util.createElement(xmlDoc, "a:xfrm");
        shapeProps.appendChild(transform2d);

        var presetGeometry = Util.createElement(xmlDoc, "a:prstGeom", [
            ["prst", "rect"]
        ]);
        shapeProps.appendChild(presetGeometry);

        pictureNode.appendChild(shapeProps);
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

        var ach = this.anchor;
        if (ach == null) {
            throw new Error("picture " + this.id + " anchor null, cannot conver to XML");
        }
        else {
            return ach.toXML(xmlDoc, pictureNode);
        }
    }

}

interface PictureDrawing extends Picture, Drawing, Drawings.Drawing { }
export = <{ new (): PictureDrawing }><any>Picture;