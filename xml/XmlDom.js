"use strict";
var XmlDom = /** @class */ (function () {
    function XmlDom(ns, rootNodeName, documentType) {
        this.documentElement = this.createElement(rootNodeName);
        this.documentElement.setAttribute("xmlns", ns);
    }
    XmlDom.prototype.createElement = function (name) {
        return new XmlDom.XMLNode({ nodeName: name });
    };
    XmlDom.prototype.createTextNode = function (text) {
        return new XmlDom.TextNode(text);
    };
    XmlDom.prototype.toString = function () {
        return this.documentElement.toString();
    };
    XmlDom.createNode = function (config) {
        switch (config.type) {
            case "XML":
                return new XmlDom.XMLNode(config);
            case "TEXT":
                return new XmlDom.TextNode(config.nodeValue);
            default:
                return null;
        }
    };
    return XmlDom;
}());
(function (XmlDom) {
    var TextNode = /** @class */ (function () {
        function TextNode(text) {
            this.nodeValue = text;
        }
        TextNode.prototype.toJSON = function () {
            return {
                nodeValue: this.nodeValue,
                type: "TEXT"
            };
        };
        TextNode.prototype.toString = function () {
            return this.nodeValue;
        };
        return TextNode;
    }());
    XmlDom.TextNode = TextNode;
    var XMLNode = /** @class */ (function () {
        function XMLNode(config) {
            this.nodeName = config.nodeName;
            this.children = [];
            this.firstChild = null;
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
        XMLNode.prototype.toString = function () {
            var str = "<" + this.nodeName + " ";
            var attrs = [];
            for (var attr in this.attributes) {
                attrs.push(attr + "=\"" + this.attributes[attr] + "\"");
            }
            str += attrs.join(" ") + ">";
            for (var i = 0, l = this.children.length; i < l; i++) {
                str += this.children[i].toString();
            }
            str += "</" + this.nodeName + ">";
            return str;
        };
        XMLNode.prototype.toJSON = function () {
            var children = [];
            for (var i = 0, l = this.children.length; i < l; i++) {
                children.push(this.children[i].toJSON());
            }
            return {
                nodeName: this.nodeName,
                nodeValue: this.nodeValue,
                children: children,
                attributes: this.attributes,
                type: "XML"
            };
        };
        XMLNode.prototype.setAttribute = function (name, val) {
            if (val === null) {
                delete this.attributes[name];
                delete this[name];
                return;
            }
            this.attributes[name] = val;
            this[name] = val;
        };
        XMLNode.prototype.setAttributeNS = function (ns, name, val) {
            this.setAttribute(name, val);
        };
        XMLNode.prototype.appendChild = function (child) {
            this.children.push(child);
            this.firstChild = this.children[0];
        };
        XMLNode.prototype.cloneNode = function (deep) {
            return new XmlDom.XMLNode(this.toJSON());
        };
        return XMLNode;
    }());
    XmlDom.XMLNode = XMLNode;
})(XmlDom || (XmlDom = {}));
module.exports = XmlDom;
