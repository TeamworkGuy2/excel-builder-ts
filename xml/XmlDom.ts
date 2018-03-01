
class XmlDom {
    documentElement: XmlDom.XMLNode;


    constructor(ns: string, rootNodeName: string, documentType: any) {
        this.documentElement = this.createElement(rootNodeName);
        this.documentElement.setAttribute("xmlns", ns);
    }


    public createElement(name: string) {
        return new XmlDom.XMLNode({ nodeName: name });
    }


    public createTextNode(text: string) {
        return new XmlDom.TextNode(text);
    }


    public toString() {
        return this.documentElement.toString();
    }


    static createNode(config: XmlDom.NodeInfo): XmlDom.NodeBase {
        switch (config.type) {
            case "XML":
                return new XmlDom.XMLNode(config);
            case "TEXT":
                return new XmlDom.TextNode(config.nodeValue);
            default:
                return null;
        }
    }

}


module XmlDom {

    export interface StringMap<T> {
        [key: string]: T;
    }


    export interface NodeLike {
        nodeValue: string;
    }


    export interface NodeInfo {
        nodeValue?: string;
        type: "XML" | "TEXT";
    }


    export interface NodeBase {
        nodeValue?: string;
        toJSON?(): NodeInfo;
    }




    export class TextNode implements NodeBase {
        nodeValue: string;


        constructor(text: string) {
            this.nodeValue = text;
        }


        public toJSON(): NodeInfo {
            return {
                nodeValue: this.nodeValue,
                type: "TEXT"
            };
        }


        public toString() {
            return this.nodeValue;
        }

    }




    export class XMLNode implements NodeBase, StringMap<any> {
        nodeName: string;
        nodeValue: string;
        children: NodeBase[];
        firstChild: NodeBase;
        attributes: { [key: string]: any };


        constructor(config: { nodeName?: string; nodeValue?: string; children?: NodeInfo[]; attributes?: { [key: string]: any }; }) {
            this.nodeName = config.nodeName;
            this.children = [];
            this.nodeValue = config.nodeValue || "";
            this.attributes = {};

            if (config.children) {
                for (var i = 0; i < config.children.length; i++) {
                    this.appendChild(XmlDom.createNode(config.children[i]));
                }
            }

            if (config.attributes) {
                for (var attr in config.attributes) {
                    this.setAttribute(attr, config.attributes[attr]);
                }
            }
        }


        public toString() {
            var str = "<" + this.nodeName + " ";
            var attrs: string[] = [];
            for (var attr in this.attributes) {
                attrs.push(attr + "=\"" + this.attributes[attr] + "\"");
            }
            str += attrs.join(" ") + ">";

            for (var i = 0, l = this.children.length; i < l; i++) {
                str += this.children[i].toString();
            }

            str += "</" + this.nodeName + ">";
            return str;
        }


        public toJSON(): (NodeInfo & StringMap<any>) {
            var children: NodeInfo[] = [];
            for (var i = 0, l = this.children.length; i < l; i++) {
                children.push(this.children[i].toJSON());
            }
            return {
                nodeName: this.nodeName,
                children: children,
                nodeValue: this.nodeValue,
                attributes: this.attributes,
                type: "XML"
            };
        }


        public setAttribute(name: string, val: any) {
            if (val === null) {
                delete this.attributes[name];
                delete (<any>this)[name];
                return;
            }
            this.attributes[name] = val;
            (<any>this)[name] = val;
        }


        public setAttributeNS(ns: string, name: string, val: any) {
            this.setAttribute(name, val);
        }


        public appendChild(child: NodeBase) {
            this.children.push(child);
            this.firstChild = this.children[0];
        }


        public cloneNode(deep?: boolean) {
            return new XmlDom.XMLNode(this.toJSON());
        }

    }

}

export = XmlDom;