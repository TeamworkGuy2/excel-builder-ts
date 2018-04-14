
declare function importScripts(...urls: string[]): void;
declare var JSZip: any;

onmessage = function (event: { data: any; }) {
    importScripts(event.data.ziplib);

    var zip = new JSZip();
    var files = event.data.files;
    for(var path in files) {
        var content = files[path];
        path = path.substr(1);
        zip.file(path, content, {base64: false});
    }

    postMessage({
        base64: !!event.data.base64
    }, <any>undefined);

    zip.generateAsync({
        base64: !!event.data.base64
    }).then(function (data: any) {
        postMessage({
            status: 'done',
            data: data
        }, <any>undefined);
    });
};
