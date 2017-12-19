const JSZip = require('jszip');
const Docxtemplater = require('docxtemplater');
const xmldom = require('xmldom');

module.exports = Document;

const systemXmlRelIds = {
    'styles.xml': 'rId1',
    'settings.xml': 'rId2',
    'webSettings.xml': 'rId3',
    'footnotes.xml': 'rId4',
    'endnotes.xml': 'rId5',
    'header1.xml': 'rId6',
    'footer1.xml': 'rId7',
    'fontTable.xml': 'rId8',
    'theme1.xml': 'rId9',
    'footer2.xml': 'rId10',
    'numbering.xml': 'rId11'
};

function Document() {
    let body = [];
    let header = [];
    let footer = [];
    let builder = body;
    let bold = false;
    let italic = false;
    let underline = false;
    let font = null;
    let size = null;
    let alignment = null;
    let rels = [];

    function replaceRIds(xml, replacements) {
        var xmlBuilder = [];
        var startingIndex = 0;
        for (var i = 0; i < xml.length; i++) {
            if (
                xml[i] == '"' &&
                xml[i + 1] == 'r' &&
                xml[i + 2] == 'I' &&
                xml[i + 3] == 'd'
            ) {
                var oldRId = ['rId'];
                i = i + 4;
                while (xml[i] != '"') {
                    oldRId.push(xml[i]);
                    i++;
                }

                oldRId = oldRId.join('');
                var newRId = replacements[oldRId] || oldRId;

                xmlBuilder.push('"');
                xmlBuilder.push(newRId);
                xmlBuilder.push('"');
            } else xmlBuilder.push(xml[i]);
        }

        return xmlBuilder.join('');
    }

    function utf8ArrayToString(array) {
        var out, i, len, c;
        var char2, char3;

        out = '';
        len = array.length;
        i = 0;
        while (i < len) {
            c = array[i++];
            switch (c >> 4) {
                case 0:
                case 1:
                case 2:
                case 3:
                case 4:
                case 5:
                case 6:
                case 7:
                    // 0xxxxxxx
                    out += String.fromCharCode(c);
                    break;
                case 12:
                case 13:
                    // 110x xxxx   10xx xxxx
                    char2 = array[i++];
                    out += String.fromCharCode(
                        ((c & 0x1f) << 6) | (char2 & 0x3f)
                    );
                    break;
                case 14:
                    // 1110 xxxx  10xx xxxx  10xx xxxx
                    char2 = array[i++];
                    char3 = array[i++];
                    out += String.fromCharCode(
                        ((c & 0x0f) << 12) |
                            ((char2 & 0x3f) << 6) |
                            ((char3 & 0x3f) << 0)
                    );
                    break;
            }
        }

        return out;
    }

    function normalizeBody(body) {
        const doc = new xmldom.DOMParser().parseFromString(body);
        const sectPrs = doc.getElementsByTagName('w:sectPr');

        if (sectPrs.length > 1) {
            for (let i = 0; i < sectPrs.length; i++) {
                const node = sectPrs[0];
                if (
                    node.firstChild.tagName !== 'w:headerReference' ||
                    node.firstChild.tagName !== 'w:footerReference'
                ) {
                    node.parentNode.removeChild(node);
                }
            }

            const serializer = new xmldom.XMLSerializer();
            body = serializer.serializeToString(doc);

            // A namespace is added to each node.  This removes it.
            body = body.replace(/xmlns\:\w+\=\"\w*\"/g, '');
        }

        return body;
    }

    this.beginHeader = function() {
        builder = header;
    };

    this.endHeader = function() {
        builder = body;
    };

    this.beginFooter = function() {
        builder = footer;
    };

    this.endFooter = function() {
        builder = body;
    };

    this.setBold = function() {
        bold = true;
    };

    this.unsetBold = function() {
        bold = false;
    };

    this.setItalic = function() {
        italic = true;
    };

    this.unsetItalic = function() {
        italic = false;
    };

    this.setUnderline = function() {
        underline = true;
    };

    this.unsetUnderline = function() {
        underline = false;
    };

    this.setFont = function(font) {
        font = font;
    };

    this.unsetFont = function() {
        font = null;
    };

    this.setSize = function(size) {
        size = size;
    };

    this.unsetSize = function() {
        size = null;
    };

    this.rightAlign = function() {
        alignment = 'right';
    };

    this.centerAlign = function() {
        alignment = 'center';
    };

    this.leftAlign = function() {
        alignment = null;
    };

    this.insertPageBreak = function() {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        builder.push(pb);
    };

    this.beginTable = function(options) {
        if (!options) {
            builder.push('<w:tbl>');
        } else {
            options = options || { borderSize: 4, borderColor: 'auto' };
            builder.push(
                '<w:tbl><w:tblPr><w:tblBorders> \
				<w:top w:val="single" w:space="0" w:color="' +
                    options.borderColor +
                    '" w:sz="' +
                    options.borderSize +
                    '"/> \
				<w:left w:val="single" w:space="0" w:color="' +
                    options.borderColor +
                    '" w:sz="' +
                    options.borderSize +
                    '"/> \
				<w:bottom w:val="single" w:space="0" w:color="' +
                    options.borderColor +
                    '" w:sz="' +
                    options.borderSize +
                    '"/> \
				<w:right w:val="single" w:space="0" w:color="' +
                    options.borderColor +
                    '" w:sz="' +
                    options.borderSize +
                    '"/> \
				<w:insideH w:val="single" w:space="0" w:color="' +
                    options.borderColor +
                    '" w:sz="' +
                    options.borderSize +
                    '"/> \
				<w:insideV w:val="single" w:space="0" w:color="' +
                    options.borderColor +
                    '" w:sz="' +
                    options.borderSize +
                    '"/> \
				</w:tblBorders>	\
			</w:tblPr>'
            );
        }
    };

    this.insertRow = function() {
        builder.push('<w:tr><w:tc>');
    };

    this.nextColumn = function() {
        builder.push('</w:tc><w:tc>');
    };

    this.nextRow = function() {
        builder.push('</w:tc></w:tr><w:tr><w:tc>');
    };

    this.endTable = function() {
        builder.push('</w:tc></w:tr></w:tbl>');
    };

    this.insertText = function(text) {
        var p =
            '<w:p>' +
            (alignment
                ? '<w:pPr><w:jc w:val="' + alignment + '"/></w:pPr>'
                : '') +
            '<w:r> \
				<w:rPr>' +
            (size ? '<w:sz w:val="' + size + '"/>' : '') +
            (bold ? '<w:b/>' : '') +
            (italic ? '<w:i/>' : '') +
            (underline ? '<w:u w:val="single"/>' : '') +
            (font
                ? '<w:rFonts w:hAnsi="' + font + '" w:ascii="' + font + '"/>'
                : '') +
            '</w:rPr> \
				<w:t>[CONTENT]</w:t> \
			</w:r> \
		</w:p>';

        builder.push(p.replace('[CONTENT]', text));
    };

    this.insertSection = function(options) {
        const startSection =
            '<w:p> \
                <w:pPr> \
                    <w:sectPr>';

        const endSection =
            '      </w:sectPr> \
                </w:pPr> \
             </w:p>';

        builder.push(startSection);

        if (options && options.orientation) {
            builder.push(
                `        <w:pgSz w:orient="${options.orientation}" />`
            );
        }

        if (options && options.type) {
            builder.push(`        <w:type w:val="${options.type}" />`);
        }

        builder.push(endSection);
    };

    this.insertRaw = function(xml) {
        builder.push(xml);
    };

    this.getExternalDocxRawXml = function(docxData) {
        var zip = new JSZip(docxData);

        var xml = utf8ArrayToString(
            zip.file('word/document.xml')._data.getContent()
        );
        xml = xml.substring(xml.indexOf('<w:body>') + 8);
        xml = xml.substring(0, xml.indexOf('</w:body>'));

        var relsXml = utf8ArrayToString(
            zip.file('word/_rels/document.xml.rels')._data.getContent()
        );
        var replacements = null;

        while (relsXml.indexOf('<Relationship') != -1) {
            relsXml = relsXml.substring(relsXml.indexOf('<Relationship') + 13);
            relsXml = relsXml.substring(relsXml.indexOf('Id="') + 4);
            var id = relsXml.substring(0, relsXml.indexOf('"'));
            relsXml = relsXml.substring(relsXml.indexOf('Type="') + 6);
            var type = relsXml.substring(0, relsXml.indexOf('"'));
            relsXml = relsXml.substring(relsXml.indexOf('Target="') + 8);
            var target = relsXml.substring(0, relsXml.indexOf('"'));

            var filename =
                target.indexOf('/') != -1
                    ? target.substring(target.lastIndexOf('/') + 1)
                    : target;
            var zipPath = target.startsWith('../')
                ? target.substring(3)
                : 'word/' + target;

            var newId = systemXmlRelIds[filename];
            var newTarget = target;

            if (!newId) {
                var hrtime = process.hrtime();
                var rand = hrtime[0] + '' + hrtime[1];
                newId = id + '_' + rand;
                newTarget = target.split('/');
                newTarget[newTarget.length - 1] =
                    rand + '_' + newTarget[newTarget.length - 1];
                newTarget = newTarget.join('/');
            }

            rels.push({
                id: id,
                newId: newId,
                data: zip.file(zipPath)._data.getContent(),
                zipPath: zipPath,
                filename: filename,
                type: type,
                target: target,
                newTarget: newTarget
            });

            replacements = replacements || {};
            replacements[id] = newId;
        }

        if (replacements) xml = replaceRIds(xml, replacements);

        return xml;
    };

    this.insertDocxSync = function(data) {
        var xml = this.getExternalDocxRawXml(data);
        this.insertRaw(xml);
    };

    this.render = function(template, generatedType) {
        if (!generatedType) {
            generatedType = 'base64';
        }

        var zip = new JSZip(template);
        var filesToSave = {};

        if (rels.length > 0) {
            var relsXmlBuilder = [];

            for (var i = 0; i < rels.length; i++) {
                var rel = rels[i];
                var saveTo = rel.newTarget.startsWith('../')
                    ? rel.newTarget.substring(3)
                    : 'word/' + rel.newTarget;

                if (rel.target != rel.newTarget) {
                    zip.file(saveTo, rel.data);
                    relsXmlBuilder.push(
                        '<Relationship Id="' +
                            rel.newId +
                            '" Type="' +
                            rel.type +
                            '" Target="' +
                            rel.newTarget +
                            '"/>'
                    );
                } else if (rel.filename.endsWith('.xml')) {
                    var zipFile = zip.file(rel.zipPath);

                    if (
                        (filesToSave[saveTo] || zipFile) &&
                        !rel.target.startsWith('theme/')
                    ) {
                        var xml = utf8ArrayToString(rel.data).substring(1);
                        xml = xml.substring(xml.indexOf('<'));
                        xml = xml.substring(xml.indexOf('>') + 1);

                        var closingTag = xml.substring(xml.lastIndexOf('</'));

                        var mergedXml =
                            filesToSave[saveTo] ||
                            utf8ArrayToString(zipFile._data.getContent());
                        mergedXml = mergedXml.replace(closingTag, xml);
                        filesToSave[saveTo] = mergedXml;
                    } else filesToSave[saveTo] = utf8ArrayToString(rel.data);
                } else console.log('Cannot merge file ' + filename);
            }

            if (relsXmlBuilder.length > 0) {
                var relsXml = utf8ArrayToString(
                    zip.file('word/_rels/document.xml.rels')._data.getContent()
                );
                relsXmlBuilder.push('</Relationships>');
                relsXml = relsXml.replace(
                    '</Relationships>',
                    relsXmlBuilder.join('')
                );
                zip.file('word/_rels/document.xml.rels', relsXml);
            }

            for (var path in filesToSave) {
                zip.file(path, filesToSave[path]);
            }
        }

        var doc = new Docxtemplater().loadZip(zip);

        doc.setData({
            body: normalizeBody(body.join('')),
            header: header.join(''),
            footer: footer.join('')
        });
        doc.render();

        return doc.getZip().generate({ type: generatedType });
    };
}
