const builder = require('docx-builder');
const docx = new builder.Document();
const docxFs = builder.fs;

docx.insertSection({ orientation: 'landscape', type: 'continuous' });
docx.insertSection({ type: 'oddPage' });
docx.insertText('Text');

docxFs.save(docx, __dirname + '/output-sections.docx', err => {
    if (err) console.log(err);
});
