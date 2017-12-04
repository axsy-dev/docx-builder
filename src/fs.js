const fs = require('fs');

export function save(docx, filepath, err) {
    const template = fs.readFileSync(__dirname + '/../template.docx', 'binary');
    const buf = docx.render(template, 'nodebuffer');
    fs.writeFile(filepath, buf, err);
}

export function insertDocxSync(docx, path) {
    const data = fs.readFileSync(path, 'binary');
    docx.insertDocxSync(data);
}

export function insertDocx(docx, path, callback) {
    fs.readFile(path, 'binary', (e, data) => {
        if (e) {
            return callback(e);
        }

        docx.insertDocxSync(data);
        callback(null);
    });
}
