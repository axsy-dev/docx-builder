const builder = require('docx-builder');
const docx = new builder.Document();
const docxFs = builder.fs;

docxFs.insertDocxSync(docx, __dirname + '/Doc1.docx');

docx.insertSection({ orientation: 'landscape' });
docx.insertSection({ type: 'oddPage' });
docx.insertSection({ type: 'continuous' });
docx.insertText('Text');

docxFs.insertDocxSync(docx, __dirname + '/Doc2.docx');

docxFs.save(docx, __dirname + '/output-merging.docx', err => {
    if (err) console.log(err);
});
