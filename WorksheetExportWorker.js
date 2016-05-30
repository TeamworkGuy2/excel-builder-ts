"use strict";
var Worksheet = require("./Worksheet");
var worksheet;
var console = {
    log: postMessage
};
var start = function (data) {
    worksheet = new Worksheet();
    worksheet.importData(data);
    postMessage({ status: 'sharedStrings', data: worksheet.collectSharedStrings() }, undefined);
};
onmessage = function (event) {
    var data = event.data;
    if (typeof data == "object") {
        switch (data.instruction) {
            case "setup":
                importScripts(data.requireJsPath);
                postMessage({ status: "ready" }, undefined);
                break;
            case "start":
                start(data.data);
                break;
            case "export":
                worksheet.setSharedStringCollection({
                    strings: data.sharedStrings
                });
                postMessage({ status: "finished", data: worksheet.toXML().toString() }, undefined);
                break;
        }
    }
};
