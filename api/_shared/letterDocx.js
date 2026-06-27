/*
 * Shared: build a branded King EPCM letter as a Word (.docx) document, with the same
 * letterhead (logo + title/name, gold+navy rule) and footer as the PDF letters.
 * Used by employment-letter and salary-letter. No external dependencies.
 *
 * buildSimpleLetterDocx({
 *   docTitle, name, dateStr, salutation,
 *   paragraphs: [ { text, bold, italic, indent(twips), after(twips) } ],
 *   closing, signer: { name, title, phone, email }, logo: Buffer|null
 * }) -> Buffer (.docx)
 */
const NAVY = "14294D", GOLD = "E5A823", MUTE = "5B6675";
const CONTENT_W = 10080; // page 12240 - margins 1080*2 (twips)

function xesc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function P(text, o) {
  o = o || {};
  const rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' + (o.bold ? '<w:b/>' : '') + (o.italic ? '<w:i/>' : '') + (o.color ? '<w:color w:val="' + o.color + '"/>' : '') + '<w:sz w:val="' + (o.sz || 22) + '"/></w:rPr>';
  const ind = o.indent ? '<w:ind w:left="' + (o.indent === true ? 360 : o.indent) + '"/>' : '';
  const ppr = '<w:pPr><w:spacing w:after="' + (o.after == null ? 200 : o.after) + '" w:line="259" w:lineRule="auto"/>' + (o.align ? '<w:jc w:val="' + o.align + '"/>' : '') + ind + '</w:pPr>';
  const t = String(text == null ? "" : text);
  if (t === "") return '<w:p>' + ppr + '</w:p>';
  const runs = t.split("\n").map(function (ln, i) { return (i ? '<w:br/>' : '') + '<w:t xml:space="preserve">' + xesc(ln) + '</w:t>'; }).join("");
  return '<w:p>' + ppr + '<w:r>' + rpr + runs + '</w:r></w:p>';
}

function buildSimpleLetterDocx(opts) {
  opts = opts || {};
  const logo = opts.logo || null;
  const sg = opts.signer || {};

  // ---- body ----
  const body = [];
  if (opts.dateStr) body.push(P(opts.dateStr, { after: 220 }));
  body.push(P("RE:    " + (opts.docTitle || ""), { bold: true, after: 0 }));
  if (opts.name) body.push(P("          " + opts.name, { bold: true, after: 200 }));
  if (opts.salutation) body.push(P(opts.salutation, { after: 160 }));
  (opts.paragraphs || []).forEach(function (p) { body.push(P(p.text, { bold: p.bold, italic: p.italic, indent: p.indent, after: (p.after == null ? 200 : p.after) })); });
  body.push(P(opts.closing || "Sincerely,", { after: 60 }));
  body.push(P("", { after: 560 })); // room to sign
  if (sg.name) body.push(P(sg.name, { bold: true, after: 0 }));
  if (sg.title) body.push(P(sg.title, { italic: true, after: 0 }));
  if (sg.phone) body.push(P("T: " + sg.phone, { after: 0 }));
  if (sg.email) body.push(P("E: " + sg.email, { after: 0 }));

  // ---- header (letterhead): logo left, title + name right, gold + navy rules ----
  const png = pngSize(logo);
  const logoCx = 1463040, logoCy = logo ? Math.max(180000, Math.round(logoCx * png.h / png.w)) : 520000;
  const logoRun = logo
    ? '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + logoCx + '" cy="' + logoCy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="11" name="logo"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="11" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + logoCx + '" cy="' + logoCy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    : '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:color w:val="' + GOLD + '"/><w:sz w:val="40"/></w:rPr><w:t>KING EPCM</w:t></w:r>';
  const noTblBorders = '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>';
  const hdrLeftW = 4600, hdrRightW = CONTENT_W - hdrLeftW;
  const headerTbl = '<w:tbl><w:tblPr><w:tblW w:w="' + CONTENT_W + '" w:type="dxa"/>' + noTblBorders + '<w:tblLayout w:type="fixed"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="' + hdrLeftW + '"/><w:gridCol w:w="' + hdrRightW + '"/></w:tblGrid><w:tr>'
    + '<w:tc><w:tcPr><w:tcW w:w="' + hdrLeftW + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' + logoRun + '</w:p></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="' + hdrRightW + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
    + P(opts.docTitle || "", { align: "right", bold: true, color: NAVY, sz: 20, after: 0 })
    + (opts.name ? P(opts.name, { align: "right", color: MUTE, sz: 18, after: 0 }) : "")
    + '</w:tc></w:tr></w:tbl>';
  const headerXml = xmlDecl() + '<w:hdr ' + nsAttrs() + '>' + headerTbl
    + '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="22" w:space="1" w:color="' + GOLD + '"/></w:pBdr><w:spacing w:after="0"/></w:pPr></w:p>'
    + '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="' + NAVY + '"/></w:pBdr><w:spacing w:after="0"/></w:pPr></w:p></w:hdr>';
  const footerXml = xmlDecl() + '<w:ftr ' + nsAttrs() + '>'
    + '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="' + GOLD + '"/></w:pBdr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:color w:val="' + MUTE + '"/><w:sz w:val="15"/></w:rPr>'
    + '<w:t xml:space="preserve">King EPCM   ·   3780 14th Avenue, Unit 211, Markham, ON  L3R 9Y5   ·   416-342-3001   ·   KingEPCM.com</w:t></w:r></w:p></w:ftr>';

  const documentXml = xmlDecl()
    + '<w:document ' + nsAttrs() + '><w:body>' + body.join("")
    + '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:footerReference w:type="default" r:id="rIdFtr"/>'
    + '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="2040" w:right="1080" w:bottom="1080" w:left="1080" w:header="540" w:footer="480" w:gutter="0"/>'
    + '<w:cols w:space="720"/></w:sectPr></w:body></w:document>';

  const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Default Extension="png" ContentType="image/png"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    + '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    + '</Types>';
  const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '</Relationships>';
  const docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    + '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    + '</Relationships>';
  const headerRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + (logo ? '<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' : '')
    + '</Relationships>';

  const files = [
    { name: "[Content_Types].xml", data: Buffer.from(contentTypes, "utf8") },
    { name: "_rels/.rels", data: Buffer.from(rootRels, "utf8") },
    { name: "word/document.xml", data: Buffer.from(documentXml, "utf8") },
    { name: "word/_rels/document.xml.rels", data: Buffer.from(docRels, "utf8") },
    { name: "word/header1.xml", data: Buffer.from(headerXml, "utf8") },
    { name: "word/_rels/header1.xml.rels", data: Buffer.from(headerRels, "utf8") },
    { name: "word/footer1.xml", data: Buffer.from(footerXml, "utf8") }
  ];
  if (logo) files.push({ name: "word/media/logo.png", data: logo });
  return zipStore(files);
}

function xmlDecl() { return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'; }
function nsAttrs() {
  return 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    + 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    + 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
}
function pngSize(buf) {
  try { if (buf && buf.length >= 24 && buf.readUInt32BE(0) === 0x89504e47) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) }; } catch (e) {}
  return { w: 600, h: 200 };
}
function crc32(buf) {
  let crc = ~0;
  for (let i = 0; i < buf.length; i++) { crc ^= buf[i]; for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1)); }
  return (~crc) >>> 0;
}
function zipStore(files) {
  const local = [], central = []; let offset = 0;
  files.forEach(function (f) {
    const name = Buffer.from(f.name, "utf8");
    const data = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data);
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0, 6); lh.writeUInt16LE(0, 8);
    lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0, 12); lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(name.length, 26); lh.writeUInt16LE(0, 28);
    local.push(lh, name, data);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0, 8);
    ch.writeUInt16LE(0, 10); ch.writeUInt16LE(0, 12); ch.writeUInt16LE(0, 14); ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(name.length, 28);
    ch.writeUInt16LE(0, 30); ch.writeUInt16LE(0, 32); ch.writeUInt16LE(0, 34); ch.writeUInt16LE(0, 36);
    ch.writeUInt32LE(0, 38); ch.writeUInt32LE(offset, 42);
    central.push(ch, name);
    offset += lh.length + name.length + data.length;
  });
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(0, 4); eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(offset, 16); eocd.writeUInt16LE(0, 20);
  return Buffer.concat([Buffer.concat(local), cd, eocd]);
}

module.exports = { buildSimpleLetterDocx: buildSimpleLetterDocx };
