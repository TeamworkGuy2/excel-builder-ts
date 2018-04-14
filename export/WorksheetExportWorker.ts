import Worksheet = require("../worksheet/Worksheet");

declare function importScripts(...urls: string[]): void;

interface WorksheetExportWorkerData {
    data: any;
    instruction: "setup" | "start" | "export";
    requireJsPath: string;
    sharedStrings: { [key: string]: number };
}


var worksheet: Worksheet;

var start = function (data: any) {
    worksheet = new Worksheet();
    worksheet.importData(data);
    postMessage({ status: "sharedStrings", data: worksheet.collectSharedStrings() }, <any>undefined);
};

onmessage = function (event: { data: WorksheetExportWorkerData }) {
    var data = event.data;
    if (typeof data === "object") {
        switch (data.instruction) {
            case "setup":
                importScripts(data.requireJsPath);
                postMessage({ status: "ready" }, <any>undefined);
                break;
            case "start":
                start(data.data);
                break;
            case "export":
                worksheet.setSharedStringCollection({
                    strings: data.sharedStrings
                });
                postMessage({ status: "finished", data: worksheet.toXML().toString() }, <any>undefined);
                break;
        }
    }
};
