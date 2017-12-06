'use strict';

Object.defineProperty(exports, "__esModule", {
    value: true
});
exports.save = save;
exports.insertDocxSync = insertDocxSync;
exports.insertDocx = insertDocx;
var fs = require('fs');

function save(docx, filepath, err) {
    var template = fs.readFileSync(__dirname + '/../template.docx', 'binary');
    var buf = docx.render(template, 'nodebuffer');
    fs.writeFile(filepath, buf, err);
}

function insertDocxSync(docx, path) {
    var data = fs.readFileSync(path, 'binary');
    docx.insertDocxSync(data);
}

function insertDocx(docx, path, callback) {
    fs.readFile(path, 'binary', function (e, data) {
        if (e) {
            return callback(e);
        }

        docx.insertDocxSync(data);
        callback(null);
    });
}