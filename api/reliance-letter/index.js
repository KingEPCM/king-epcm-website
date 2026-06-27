/*
 * POST /api/reliance-letter  (licensed engineers / admin only)
 * Generates a branded King EPCM Reliance Letter and returns BOTH a PDF and a Word (.docx)
 * for download. The two formats share one content model (composeLetter) so the wording is
 * always identical between them.
 *   - "bank"      : lender/bank reliance letter (RBC style)
 *   - "municipal" : municipality/region reliance letter under O.Reg 153/04 (Durham/Toronto style)
 *   - "other"     : owner / private reliance letter
 * No payment/SharePoint integration — the files are returned to the browser to save.
 *
 * Gated to the SWA roles "engineer" or "admin".
 * App settings: none beyond the function runtime (pure document generation).
 */
const SIGNER = { name: "Yu Tao (Tony) Wang, P. Eng", title: "President, Principal Engineer", peo: "100228476", coa: "100538219", phone: "416-342-3001", email: "Twang@KingEPCM.com" };

module.exports = async function (context, req) {
 try {
  const roles = principalRoles(req);
  if (roles.indexOf("engineer") === -1 && roles.indexOf("admin") === -1) {
    context.res = json(403, { ok: false, error: "Engineers only" }); return;
  }
  const b = req.body || {};
  const type = /muni/i.test(b.relType) ? "municipal" : /other|owner|private/i.test(b.relType) ? "other" : "bank";
  const reports = (Array.isArray(b.reports) ? b.reports : []).filter(function (r) { return r && (r.type || r.date); });
  const site = String(b.siteAddress || "").trim();

  if (!site) { context.res = json(400, { ok: false, error: "Site / property address is required" }); return; }
  if (!reports.length) { context.res = json(400, { ok: false, error: "Add at least one report" }); return; }
  if ((type === "bank" || type === "other") && !String(b.recipientOrg || "").trim()) { context.res = json(400, { ok: false, error: "Recipient organisation is required" }); return; }
  if (type === "municipal" && !String(b.authorityName || "").trim()) { context.res = json(400, { ok: false, error: "Municipality / authority name is required" }); return; }

  const dateStr = plainDate(b.date ? parseDate(b.date) : new Date());
  const model = composeLetter(type, b, reports, site);
  const logo = loadLogo();

  const pdf = await buildLetterPdf(model, site, dateStr, logo);
  const docx = buildLetterDocx(model, site, dateStr, logo);

  const who = type === "municipal" ? (b.authorityName || "") : (b.recipientOrg || "");
  const base = sanitize("DRAFT - Reliance Letter - " + (who ? who + " - " : "") + site + " - " + shortToday());
  context.res = json(200, {
    ok: true,
    filename: base + ".pdf", pdfBase64: pdf.toString("base64"),
    docxFilename: base + ".docx", docxBase64: docx.toString("base64")
  });
 } catch (e) {
  context.log.error(e);
  context.res = json(500, { ok: false, error: String((e && e.message) || e) });
 }
};

function loadLogo() {
  try { return require("fs").readFileSync(require("path").join(__dirname, "logo.png")); }
  catch (e) { return null; }
}

/* ============================================================================
 * Shared content model — the single source of truth for the letter wording.
 * Both the PDF and the DOCX renderers consume this, so they can never diverge.
 * ==========================================================================*/
function composeLetter(type, b, reports, site) {
  const sg = { name: t(b.signerName, SIGNER.name), title: t(b.signerTitle, SIGNER.title), phone: t(b.phone, SIGNER.phone), email: t(b.email, SIGNER.email) };
  const basis = /eng/i.test(b.warrantyBasis) ? "engineering" : "esa";
  const qp = t(b.qpName, "");
  const reportLines = reports.map(function (r) {
    const parts = [t(r.type, "Report") + " at " + site, fmtDate(r.date)];
    if (qp) parts.push(qp + (basis === "esa" ? " (QP)" : ""));
    parts.push("King EPCM");
    return parts.filter(Boolean).join(", ");
  });

  if (type === "bank" || type === "other") {
    const isBank = type === "bank";
    const org = t(b.recipientOrg, "the recipient");
    const toLines = [b.recipientName, b.recipientOrg].filter(Boolean)
      .concat(String(b.recipientAddress || "").split(/\r?\n/))
      .filter(function (x) { return String(x).trim(); });
    const before = [];
    if (basis === "engineering") {
      before.push({ text: "King EPCM warrants that the report(s) identified herein were prepared by, or under the direct supervision of, a Professional Engineer licensed in the Province of Ontario, in accordance with the generally accepted standards of the engineering profession applicable at the time and place the services were performed." });
    } else if (isBank) {
      before.push({ text: "King EPCM warrants that the work completed in the reports identified herein was completed by or under the supervision of a Qualified Person per the meaning of Sections 5 and 6, as applicable, of Ontario Regulation 153/04. King EPCM represents and warrants that the report complies with all sections of the Ontario Regulation 153/04, including the insurance provisions contained therein." });
    } else {
      before.push({ text: "King EPCM warrants that the work completed in the report(s) identified herein was completed by or under the supervision of a Qualified Person per the meaning of Sections 5 and 6, as applicable, of Ontario Regulation 153/04." });
    }
    if (isBank) {
      before.push({ text: "King EPCM shall provide proof of insurance and maintain Professional Liability (also known as Errors & Omissions) insurance coverage of " + money(b.plPerClaim, "$2,000,000") + " per claim and " + money(b.plAggregate, "$4,000,000") + " aggregate, in addition to Commercial General Liability insurance coverage of " + money(b.cglPerClaim, "$5,000,000") + " per claim and " + money(b.cglAggregate, "$5,000,000") + " aggregate." });
    }
    if (isBank) {
      before.push({ text: "King EPCM agrees that the report(s) listed herein may be relied upon by " + org + " to the same extent as the original client, except that our potential liability to " + org + " arising out of the report is limited to the amount of professional liability insurance coverage maintained (to a maximum of " + t(b.liabilityCap, "$1 million or the loan amount, whichever is less") + "), regardless of any limitation on liability agreed to by our client." });
    } else {
      before.push({ text: "King EPCM agrees that the report(s) listed herein may be relied upon by " + org + " to the same extent as our original client, subject always to the qualifications and limitations contained in the report(s). " + org + " has no greater rights than those of our original client under our agreement and the report(s)." });
    }
    const after = [{ text: "Should you have any questions, please do not hesitate to contact our office.", gap: 26 }];
    return { kind: "corp", heading: { text: "RELIANCE LETTER", big: true }, showTo: true, toLines: toLines, re: "RE:  Letter of Reliance – " + site, before: before, reportLines: reportLines, after: after, sg: sg };
  }

  const authority = t(b.authorityName, "the Authority");
  const shortName = t(b.authorityShort, "the Authority");
  const client = t(b.clientName, "the Owner");
  const showTo = !!(String(b.authorityAddress || "").trim() || String(b.authorityContact || "").trim());
  const toLines = [authority, b.authorityContact].filter(function (x) { return String(x || "").trim(); })
    .concat(String(b.authorityAddress || "").split(/\r?\n/))
    .filter(function (x) { return String(x).trim(); });
  const before = [
    { text: "At the request of " + client + " and for other good and valuable consideration, KING EPCM represents and warrants to " + authority + " (“" + shortName + "”) that the work completed in the reports identified herein was completed by or under the supervision of a Qualified Person per the meaning of Sections 5 and 6, as applicable, of Ontario Regulation 153/04." },
    { text: "If applicable, KING EPCM also represents and warrants that it meets the MECP requirements in relation to the preparation and supervision of a risk assessment." },
    { text: "KING EPCM agrees that " + shortName + " and its Peer Reviewers may rely upon the reports listed herein, including the representations, assumptions, findings, and recommendations contained in the reports:" }
  ];
  const after = [
    { text: "KING EPCM further agrees that in the case of any inconsistency between this Reliance Agreement and any limitations within any reports provided to " + authority + ", this Reliance Agreement takes priority over any such limitations." },
    { text: "KING EPCM further agrees that it will promptly notify " + shortName + " upon receipt of notice by the Ministry of the Environment, Conservation and Parks (MECP) that the MECP intends to audit any of the reports listed herein and, if so, to provide " + shortName + " with written confirmation of the results of the audit, including that any Record of Site Condition was approved by the MECP under Ontario Regulation 153/04." },
    { text: "KING EPCM represents and warrants that it complies with Ontario Regulation 153/04 and, specifically, the insurance provisions contained therein. KING EPCM shall provide " + shortName + " with proof of insurance and maintain Professional Liability insurance coverage of " + money(b.mplPerClaim, "$2,000,000") + " per claim and " + money(b.mplAggregate, "$2,000,000") + " aggregate." },
    { text: client + " agrees that it shall be responsible to indemnify and save " + shortName + " harmless from any and all claims, demands, causes of action, costs, including defending against any legal proceedings or other damages howsoever arising from " + shortName + "’s direct or indirect reliance upon the representations, findings, assumptions and conclusions contained in the reports prepared by KING EPCM listed herein, save and except any damages, claims, demands, actions or causes of action arising out of or as a result of the negligent actions of " + shortName + ", its agents or employees.", gap: 22 }
  ];
  return { kind: "muni", heading: { text: authority + " Reliance Letter", big: false }, showTo: showTo, toLines: toLines, re: null, before: before, reportLines: reportLines, after: after, client: client };
}

/* ============================================================================
 * PDF renderer (pdfkit)
 * ==========================================================================*/
function buildLetterPdf(model, site, dateStr, logo) {
  return new Promise(function (resolve, reject) {
    try {
      const PDFDocument = require("pdfkit");
      const doc = new PDFDocument({ size: "LETTER", margins: { top: 90, left: 64, right: 64, bottom: 84 }, bufferPages: true });
      const chunks = [];
      doc.on("data", function (d) { chunks.push(d); });
      doc.on("end", function () { resolve(Buffer.concat(chunks)); });

      const NAVY = "#14294D", GOLD = "#E5A823", INK = "#222", MUTE = "#5B6675";
      const W = doc.page.width, H = doc.page.height, L = 64, R = W - 64, CW = R - L;
      const DOCTYPE = "Letter of Reliance";

      function drawFooter() {
        doc.page.margins.bottom = 0;
        const y = H - 52;
        doc.moveTo(L, y).lineTo(R, y).lineWidth(2.2).strokeColor(GOLD).stroke();
        doc.moveTo(L, y + 3.2).lineTo(R, y + 3.2).lineWidth(0.8).strokeColor(NAVY).stroke();
        doc.fillColor(MUTE).font("Times-Roman").fontSize(8)
          .text("King EPCM   ·   3780 14th Avenue, Unit 211, Markham, ON  L3R 9Y5   ·   416-342-3001   ·   KingEPCM.com", L, y + 13, { width: CW, align: "center", lineBreak: false });
      }
      let headerRuleY = null;
      function drawLetterhead() {
        let drew = false;
        if (logo) { try { doc.image(logo, L, 44, { width: 128 }); drew = true; } catch (e) {} }
        if (!drew) {
          doc.fillColor(GOLD).font("Times-Bold").fontSize(22).text("KING EPCM", L, 46);
          doc.fillColor(NAVY).font("Times-Bold").fontSize(8).text("Flexible. Dependable. On-site Engineering.", L, 74);
        }
        const rx = L + CW * 0.40, rw = CW * 0.60;
        doc.fillColor(MUTE).font("Times-Roman").fontSize(10).text(dateStr, rx, 48, { width: rw, align: "right" });
        doc.fillColor(NAVY).font("Times-Bold").fontSize(11).text(DOCTYPE, rx, 64, { width: rw, align: "right" });
        doc.fillColor(MUTE).font("Times-Roman").fontSize(9).text(site, rx, 80, { width: rw, align: "right" });
        if (headerRuleY == null) headerRuleY = Math.max(96, doc.y + 4);
        doc.moveTo(L, headerRuleY).lineTo(R, headerRuleY).lineWidth(2.5).strokeColor(GOLD).stroke();
        doc.moveTo(L, headerRuleY + 4.5).lineTo(R, headerRuleY + 4.5).lineWidth(0.8).strokeColor(NAVY).stroke();
      }
      drawLetterhead();
      const bodyTop = headerRuleY + 18;
      doc.page.margins.top = bodyTop;
      if (doc.options && doc.options.margins) doc.options.margins.top = bodyTop;
      doc.y = bodyTop;

      function para(text, opts) {
        opts = opts || {};
        doc.fillColor(INK).font("Times-Roman").fontSize(10.5).text(text, L, doc.y, { width: CW, align: opts.align || "left", lineGap: 2, paragraphGap: opts.gap == null ? 11 : opts.gap });
      }
      function heading(text) { doc.fillColor(NAVY).font("Times-Bold").fontSize(13).text(text, L, doc.y, { width: CW }); doc.moveDown(0.4); }
      function ensure(h) { if (doc.y + h > H - doc.page.margins.bottom) doc.addPage(); }
      function renderReportLines(lines) {
        doc.fillColor(NAVY).font("Times-Bold").fontSize(10.5).text("Report(s) relied upon:", L, doc.y, { width: CW });
        doc.moveDown(0.3);
        lines.forEach(function (line) {
          ensure(34);
          doc.fillColor(INK).font("Times-Roman").fontSize(10.5).text("•  " + line, L + 6, doc.y, { width: CW - 6, lineGap: 2, paragraphGap: 6 });
        });
        doc.moveDown(0.6);
      }

      // ---- body from the shared model ----
      if (model.heading.big) { doc.fillColor(NAVY).font("Times-Bold").fontSize(15).text(model.heading.text, L, doc.y, { width: CW }); doc.moveDown(0.6); }
      else { heading(model.heading.text); }

      if (model.showTo && model.toLines.length) {
        doc.fillColor(NAVY).font("Times-Bold").fontSize(10.5).text("To:", L, doc.y); doc.moveDown(0.2);
        doc.fillColor(INK).font("Times-Roman").fontSize(10.5).text(model.toLines.join("\n"), L, doc.y, { width: CW, lineGap: 1.5 });
        doc.moveDown(0.8);
      }
      if (model.re) para(model.re, { gap: 12 });
      model.before.forEach(function (p) { para(p.text, { gap: p.gap }); });
      renderReportLines(model.reportLines);
      model.after.forEach(function (p) { para(p.text, { gap: p.gap }); });

      if (model.kind === "corp") signature(doc, model.sg, NAVY, INK, MUTE, L, CW);
      else municipalSignature(doc, model.client, NAVY, INK, MUTE, L, CW);

      function drawWatermark() {
        doc.save();
        doc.rotate(-45, { origin: [W / 2, H / 2] });
        doc.font("Times-Bold").fontSize(130).fillColor("#9aa0a6").fillOpacity(0.13);
        var tw = doc.widthOfString("DRAFT");
        doc.text("DRAFT", W / 2 - tw / 2, H / 2 - 70, { lineBreak: false });
        doc.fillOpacity(1).restore();
      }
      var range = doc.bufferedPageRange();
      for (var pi = range.start; pi < range.start + range.count; pi++) { doc.switchToPage(pi); if (pi > range.start) drawLetterhead(); drawWatermark(); drawFooter(); }
      doc.end();
    } catch (e) { reject(e); }
  });
}

function signature(doc, sg, NAVY, INK, MUTE, L, CW) {
  const R = L + CW, LINE = "#555";
  if (doc.y + 150 > doc.page.height - doc.page.margins.bottom) doc.addPage();
  doc.fillColor(INK).font("Times-Roman").fontSize(10.5).text("Yours very truly,", L, doc.y); doc.moveDown(0.3);
  doc.fillColor(NAVY).font("Times-Bold").fontSize(11).text("King EPCM", L, doc.y);
  doc.moveDown(2.8);
  const y = doc.y, sigW = CW * 0.50, gap = CW * 0.08, dateX = L + sigW + gap;
  doc.moveTo(L, y).lineTo(L + sigW, y).lineWidth(0.9).strokeColor(LINE).stroke();
  doc.moveTo(dateX, y).lineTo(R, y).lineWidth(0.9).strokeColor(LINE).stroke();
  doc.fillColor(MUTE).font("Times-Roman").fontSize(8).text("Date", dateX, y + 4, { width: R - dateX });
  doc.y = y + 12;
  doc.fillColor(INK).font("Times-Bold").fontSize(11).text(sg.name, L, doc.y);
  if (sg.title) doc.fillColor(INK).font("Times-Roman").fontSize(10).text(sg.title, L, doc.y);
  doc.fillColor(INK).font("Times-Roman").fontSize(10).text("King EPCM", L, doc.y);
  var lines = [];
  if (sg.phone) lines.push("Tel:     " + sg.phone);
  if (sg.email) lines.push("Email:  " + sg.email);
  if (lines.length) { doc.moveDown(0.35); doc.fillColor(MUTE).font("Times-Roman").fontSize(9.5).text(lines.join("\n"), L, doc.y, { width: CW, lineGap: 3.5 }); }
}
function municipalSignature(doc, client, NAVY, INK, MUTE, L, CW) {
  const R = L + CW, LINE = "#555";
  function ensureSpace(h) { if (doc.y + h > doc.page.height - doc.page.margins.bottom) doc.addPage(); }
  doc.moveDown(0.6);
  doc.fillColor(INK).font("Times-Roman").fontSize(10.5).text("IN WITNESS WHEREOF the parties have executed this Reliance Agreement as of the date written below.", L, doc.y, { width: CW, lineGap: 2 });
  doc.moveDown(1.4);
  function sigDateBlock(caption) {
    ensureSpace(72);
    doc.fillColor(NAVY).font("Times-Bold").fontSize(10).text(caption, L, doc.y, { width: CW });
    doc.moveDown(2.0);
    const y = doc.y, sigW = CW * 0.56, gap = CW * 0.08, dateX = L + sigW + gap;
    doc.moveTo(L, y).lineTo(L + sigW, y).lineWidth(0.9).strokeColor(LINE).stroke();
    doc.moveTo(dateX, y).lineTo(R, y).lineWidth(0.9).strokeColor(LINE).stroke();
    doc.fillColor(MUTE).font("Times-Roman").fontSize(8).text("Signature", L, y + 4, { width: sigW });
    doc.fillColor(MUTE).font("Times-Roman").fontSize(8).text("Date", dateX, y + 4, { width: R - dateX });
    doc.y = y + 24;
    doc.moveDown(1.1);
  }
  sigDateBlock("Signed by a Qualified Person");
  sigDateBlock("Signed by the person authorized to bind the Consulting Firm (King EPCM)");
  doc.moveDown(0.2);
  ensureSpace(54);
  doc.fillColor(NAVY).font("Times-Bold").fontSize(11).text("Property Owner, or Authorized Officer", L, doc.y, { width: CW });
  const ry = doc.y + 4;
  doc.moveTo(L, ry).lineTo(R, ry).lineWidth(0.8).strokeColor(LINE).stroke();
  doc.y = ry + 12;
  function fieldRow(label, value) {
    ensureSpace(30);
    const y = doc.y, labW = CW * 0.34, lineX = L + labW;
    if (label) doc.fillColor(INK).font("Times-Roman").fontSize(10).text(label, L, y + 1, { width: labW - 8 });
    if (value) doc.fillColor(INK).font("Times-Bold").fontSize(10).text(value, lineX, y + 1, { width: R - lineX });
    else { const ly = y + 12; doc.moveTo(lineX, ly).lineTo(R, ly).lineWidth(0.8).strokeColor(LINE).stroke(); }
    doc.y = y + 26;
  }
  fieldRow("Name (please print):");
  fieldRow("Signature:");
  fieldRow("Title of Authorized Officer:");
  fieldRow("Company name:", client);
  fieldRow("Address:");
  fieldRow("", null);
  fieldRow("Telephone:");
  fieldRow("Date:");
}

/* ============================================================================
 * DOCX renderer (hand-built Open XML, zipped with no external dependency)
 * ==========================================================================*/
function buildLetterDocx(model, site, dateStr, logo) {
  const NAVY = "14294D", GOLD = "E5A823", MUTE = "5B6675", LINE = "555555";
  const CONTENT_W = 10080; // page 12240 - margins 1080*2 (twips)

  function xesc(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  // Paragraph. opts: bold, color, sz(half-pts), align, after(twips)
  function P(text, opts) {
    opts = opts || {};
    const rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>' + (opts.bold ? '<w:b/>' : '') + (opts.color ? '<w:color w:val="' + opts.color + '"/>' : '') + '<w:sz w:val="' + (opts.sz || 21) + '"/></w:rPr>';
    const ppr = '<w:pPr><w:spacing w:after="' + (opts.after == null ? 180 : opts.after) + '" w:line="259" w:lineRule="auto"/>' + (opts.align ? '<w:jc w:val="' + opts.align + '"/>' : '') + '</w:pPr>';
    const t = String(text == null ? "" : text);
    if (t === "") return '<w:p>' + ppr + '</w:p>';
    const runs = t.split("\n").map(function (ln, i) { return (i ? '<w:br/>' : '') + '<w:t xml:space="preserve">' + xesc(ln) + '</w:t>'; }).join("");
    return '<w:p>' + ppr + '<w:r>' + rpr + runs + '</w:r></w:p>';
  }
  function ruleP(color, sz, after) {
    return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="' + sz + '" w:space="1" w:color="' + color + '"/></w:pBdr><w:spacing w:after="' + (after == null ? 160 : after) + '"/></w:pPr></w:p>';
  }
  const noTblBorders = '<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>';
  // Signature + date ruled line (3 columns: signature | gap | date)
  function sigDateTable(leftCap, rightCap) {
    const lw = 5400, mw = 480, rw = CONTENT_W - lw - mw;
    function bcell(w, bordered) {
      return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (bordered ? '<w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="' + LINE + '"/></w:tcBorders>' : '') + '</w:tcPr><w:p><w:pPr><w:spacing w:before="260" w:after="0"/></w:pPr></w:p></w:tc>';
    }
    function ccell(w, cap) {
      return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' + P(cap || "", { sz: 16, color: MUTE, after: 120 }) + '</w:tc>';
    }
    return '<w:tbl><w:tblPr><w:tblW w:w="' + CONTENT_W + '" w:type="dxa"/>' + noTblBorders + '<w:tblLayout w:type="fixed"/></w:tblPr>'
      + '<w:tblGrid><w:gridCol w:w="' + lw + '"/><w:gridCol w:w="' + mw + '"/><w:gridCol w:w="' + rw + '"/></w:tblGrid>'
      + '<w:tr>' + bcell(lw, true) + bcell(mw, false) + bcell(rw, true) + '</w:tr>'
      + '<w:tr>' + ccell(lw, leftCap) + ccell(mw, "") + ccell(rw, rightCap) + '</w:tr></w:tbl>';
  }
  // Field rows: [label, value]; blank value => ruled line.
  function fieldRows(rows) {
    const labW = 3400, valW = CONTENT_W - labW;
    const trs = rows.map(function (r) {
      const label = r[0], value = r[1];
      const lc = '<w:tc><w:tcPr><w:tcW w:w="' + labW + '" w:type="dxa"/></w:tcPr>' + P(label || "", { sz: 20, after: 60 }) + '</w:tc>';
      let vc;
      if (value) vc = '<w:tc><w:tcPr><w:tcW w:w="' + valW + '" w:type="dxa"/></w:tcPr>' + P(value, { bold: true, sz: 20, after: 60 }) + '</w:tc>';
      else vc = '<w:tc><w:tcPr><w:tcW w:w="' + valW + '" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="' + LINE + '"/></w:tcBorders></w:tcPr><w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr></w:p></w:tc>';
      return '<w:tr>' + lc + vc + '</w:tr>';
    }).join("");
    return '<w:tbl><w:tblPr><w:tblW w:w="' + CONTENT_W + '" w:type="dxa"/>' + noTblBorders + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="' + labW + '"/><w:gridCol w:w="' + valW + '"/></w:tblGrid>' + trs + '</w:tbl>';
  }

  // ---- body from the shared model ----
  const body = [];
  body.push(P(model.heading.text, { bold: true, color: NAVY, sz: model.heading.big ? 30 : 26, after: 220 }));
  if (model.showTo && model.toLines.length) {
    body.push(P("To:", { bold: true, color: NAVY, after: 40 }));
    body.push(P(model.toLines.join("\n"), { after: 220 }));
  }
  if (model.re) body.push(P(model.re, { after: 240 }));
  model.before.forEach(function (p) { body.push(P(p.text, { after: 240 })); });
  body.push(P("Report(s) relied upon:", { bold: true, color: NAVY, after: 60 }));
  model.reportLines.forEach(function (ln) { body.push(P("•  " + ln, { after: 80 })); });
  body.push(P("", { after: 80 }));
  model.after.forEach(function (p) { body.push(P(p.text, { after: p.gap ? Math.round(p.gap * 18) : 240 })); });

  if (model.kind === "corp") {
    const sg = model.sg;
    body.push(P("Yours very truly,", { after: 60 }));
    body.push(P("King EPCM", { bold: true, color: NAVY, sz: 22, after: 700 }));
    body.push(sigDateTable("", "Date"));
    body.push(P(sg.name, { bold: true, sz: 22, after: 0 }));
    if (sg.title) body.push(P(sg.title, { sz: 20, after: 0 }));
    body.push(P("King EPCM", { sz: 20, after: 60 }));
    const lines = [];
    if (sg.phone) lines.push("Tel:     " + sg.phone);
    if (sg.email) lines.push("Email:  " + sg.email);
    if (lines.length) body.push(P(lines.join("\n"), { sz: 19, color: MUTE, after: 0 }));
  } else {
    body.push(P("IN WITNESS WHEREOF the parties have executed this Reliance Agreement as of the date written below.", { after: 320 }));
    body.push(P("Signed by a Qualified Person", { bold: true, color: NAVY, sz: 20, after: 220 }));
    body.push(sigDateTable("Signature", "Date"));
    body.push(P("", { after: 140 }));
    body.push(P("Signed by the person authorized to bind the Consulting Firm (King EPCM)", { bold: true, color: NAVY, sz: 20, after: 220 }));
    body.push(sigDateTable("Signature", "Date"));
    body.push(P("", { after: 220 }));
    body.push(P("Property Owner, or Authorized Officer", { bold: true, color: NAVY, sz: 22, after: 40 }));
    body.push(ruleP(LINE, 6, 140));
    body.push(fieldRows([
      ["Name (please print):", ""], ["Signature:", ""], ["Title of Authorized Officer:", ""],
      ["Company name:", model.client], ["Address:", ""], ["", ""], ["Telephone:", ""], ["Date:", ""]
    ]));
  }

  // ---- header (letterhead): logo left, date/doctype/site right, gold+navy rules ----
  const png = pngSize(logo);
  const logoCx = 1463040; // 1.6in
  const logoCy = logo ? Math.max(180000, Math.round(logoCx * png.h / png.w)) : 520000;
  const logoRun = logo
    ? '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + logoCx + '" cy="' + logoCy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="11" name="logo"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="11" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + logoCx + '" cy="' + logoCy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    : '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:color w:val="' + GOLD + '"/><w:sz w:val="40"/></w:rPr><w:t>KING EPCM</w:t></w:r>';
  const hdrLeftW = 4600, hdrRightW = CONTENT_W - hdrLeftW;
  const headerTbl = '<w:tbl><w:tblPr><w:tblW w:w="' + CONTENT_W + '" w:type="dxa"/>' + noTblBorders + '<w:tblLayout w:type="fixed"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="' + hdrLeftW + '"/><w:gridCol w:w="' + hdrRightW + '"/></w:tblGrid><w:tr>'
    + '<w:tc><w:tcPr><w:tcW w:w="' + hdrLeftW + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' + logoRun + '</w:p></w:tc>'
    + '<w:tc><w:tcPr><w:tcW w:w="' + hdrRightW + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
    + P(dateStr, { align: "right", sz: 20, color: MUTE, after: 0 })
    + P("Letter of Reliance", { align: "right", bold: true, color: NAVY, sz: 22, after: 0 })
    + P(site, { align: "right", sz: 18, color: MUTE, after: 0 })
    + '</w:tc></w:tr></w:tbl>';
  const headerXml = xmlDecl()
    + '<w:hdr ' + nsAttrs() + '>' + headerTbl
    + '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="22" w:space="1" w:color="' + GOLD + '"/></w:pBdr><w:spacing w:after="0"/></w:pPr></w:p>'
    + '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="' + NAVY + '"/></w:pBdr><w:spacing w:after="0"/></w:pPr></w:p>'
    + '</w:hdr>';
  const footerXml = xmlDecl()
    + '<w:ftr ' + nsAttrs() + '>'
    + '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="22" w:space="1" w:color="' + GOLD + '"/><w:bottom w:val="single" w:sz="6" w:space="1" w:color="' + NAVY + '"/></w:pBdr><w:spacing w:after="0" w:line="40" w:lineRule="exact"/></w:pPr></w:p>'
    + '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="90" w:after="0"/></w:pPr>'
    + '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:color w:val="' + MUTE + '"/><w:sz w:val="15"/></w:rPr>'
    + '<w:t xml:space="preserve">King EPCM   ·   3780 14th Avenue, Unit 211, Markham, ON  L3R 9Y5   ·   416-342-3001   ·   KingEPCM.com</w:t></w:r></w:p>'
    + '</w:ftr>';

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

/* ---------- minimal STORED-zip writer (no dependency) ---------- */
function crc32(buf) {
  let crc = ~0;
  for (let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for (let j = 0; j < 8; j++) crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
  }
  return (~crc) >>> 0;
}
function zipStore(files) {
  const local = [], central = [];
  let offset = 0;
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

/* ---------- helpers ---------- */
function principalRoles(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return [];
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    return Array.isArray(p.userRoles) ? p.userRoles : [];
  } catch (e) { return []; }
}
function t(v, d) { const s = String(v == null ? "" : v).trim(); return s || d; }
function money(v, d) {
  const s = String(v == null ? "" : v).trim();
  if (!s) return d;
  if (/^\$|million|aggregate/i.test(s)) return s;
  const n = Number(s.replace(/[^0-9.]/g, ""));
  return isNaN(n) ? s : "$" + n.toLocaleString("en-CA");
}
function parseDate(s) { const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(s); }
function fmtDate(s) {
  if (!s) return "";
  const d = parseDate(s);
  if (isNaN(d)) return String(s);
  return d.toLocaleDateString("en-CA", { year: "numeric", month: "long", day: "numeric" });
}
function plainDate(d) {
  try {
    const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    return months[d.getMonth()] + " " + d.getDate() + ", " + d.getFullYear();
  } catch (e) { return new Date().toDateString(); }
}
function sanitize(s) { return String(s || "letter").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 120); }
function shortToday() { const d = new Date(); return String(d.getFullYear()).slice(-2) + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0"); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }

// Exposed for local preview/testing only (harmless in production).
module.exports.composeLetter = composeLetter;
module.exports.buildLetterPdf = buildLetterPdf;
module.exports.buildLetterDocx = buildLetterDocx;
module.exports.loadLogo = loadLogo;
