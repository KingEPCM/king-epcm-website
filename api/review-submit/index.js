/*
 * POST /api/review-submit  — receives a completed performance-review form and:
 *   1) stores a copy in Azure Table Storage (always, if storage configured)
 *   2) generates a branded PDF
 *   3) files the PDF into the staff member's KRA / Development Review folder (Graph)
 *   4) emails a copy to HR (Graph sendMail)
 * Each step is best-effort and independently configured, so the form still works
 * (stores) even before the SharePoint-write / email permissions are added.
 *
 * Body (JSON), built by the front end so every form reuses this one endpoint:
 *   { form, title, quarter, date, subjectName?, sections: [ { heading, rows:[{label,value}] } ] }
 *   - subjectName: only for manager forms (Form C) — the team member being assessed.
 *
 * App settings:
 *   NEWS_STORAGE_CONNECTION                         (store copies; reuses the news storage)
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT  (Graph app-only)
 *   REVIEW_SITE_PATH, REVIEW_BASE_PATH, REVIEW_FOLDER_MAP, REVIEW_SUBFOLDERS  (same as my-reviews)
 *   Graph APP permissions: Sites.ReadWrite.All (file the PDF), Mail.Send (email), User.Read.All (name lookup)
 *   Sends AS the person completing the form (Graph Mail.Send "send as any user").
 *   REVIEW_NOTIFY_FROM   optional fallback sender if a submitter address isn't available
 *   REVIEW_NOTIFY_TO     comma-separated recipients (default hr@kingepcm.com)
 */
const { TableClient } = require("@azure/data-tables");
const PDFDocument = require("pdfkit");

const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.REVIEW_SITE_PATH;
const BASE_PATH = process.env.REVIEW_BASE_PATH;
const LIBRARY = process.env.REVIEW_LIBRARY || "";
const CONN = process.env.NEWS_STORAGE_CONNECTION;
const NOTIFY_FROM = process.env.REVIEW_NOTIFY_FROM || "";
const NOTIFY_TO = process.env.REVIEW_NOTIFY_TO || NOTIFY_FROM || "hr@kingepcm.com";
let FOLDER_MAP = {}; try { FOLDER_MAP = JSON.parse(process.env.REVIEW_FOLDER_MAP || "{}"); } catch (e) {}
const SUB_RE = new RegExp("(" + (process.env.REVIEW_SUBFOLDERS || "kra,development review,performance review,review,appraisal,kpi")
  .split(",").map(function (s) { return s.trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }).filter(Boolean).join("|") + ")", "i");
let _driveId;

module.exports = async function (context, req) {
  const p = principal(req);
  if (!p || !p.email) { context.res = json(401, { error: "Not signed in" }); return; }
  const body = req.body || {};
  if (!body.sections || !body.sections.length) { context.res = json(400, { error: "Empty submission" }); return; }

  const isManagerForm = /^C$/i.test(body.form || "");   // restricted to managers/admins
  const isSubjectForm = /^[CD]$/i.test(body.form || ""); // routed by a typed subject name (C, D)
  const isGrowthForm = /^D$/i.test(body.form || "");     // Form D — copy staff + the person completing it
  if (isManagerForm && !(p.roles.indexOf("manager") > -1 || p.roles.indexOf("admin") > -1)) {
    context.res = json(403, { error: "Manager only" }); return;
  }

  const status = { ok: false, stored: false, filed: false, emailed: false, notes: [] };
  let token = null, myName = "";
  try { token = await appToken(); } catch (e) { context.log.error(e); }

  // Resolve the signed-in person's real name (for self forms + filing).
  if (token) { try { var me = await lookupUser(token, p.email); myName = (me && me.displayName) || ""; } catch (e) {} }
  if (!myName) myName = p.name || p.email;
  const subjectName = (isSubjectForm && body.subjectName) ? String(body.subjectName).trim() : myName;

  // Shortened export name, no version: e.g. "Form B - Jessica Fung - Q2 2026".
  const shortTitle = body.form ? ("Form " + String(body.form).trim().toUpperCase()) : (body.title || "Review");
  const fileBase = sanitize(shortTitle + " - " + subjectName + " - " + (body.quarter || body.date || today()));

  // 1) Store a copy in Table Storage.
  if (CONN) {
    try {
      const t = TableClient.fromConnectionString(CONN, "reviewforms");
      try { await t.createTable(); } catch (e) {}
      await t.createEntity({
        partitionKey: "form", rowKey: String(Date.now()) + String(Math.floor(Math.random() * 1000)),
        form: body.form || "", title: body.title || "", version: String(body.version || ""), byEmail: p.email, byName: myName,
        subjectName: subjectName, quarter: body.quarter || "", date: body.date || today(),
        data: JSON.stringify(body.sections).slice(0, 60000), createdAt: new Date().toISOString()
      });
      status.stored = true;
    } catch (e) { context.log.error(e); status.notes.push("store failed"); }
  } else status.notes.push("storage not configured");

  // 2) Generate the PDF.
  let pdf = null;
  try { pdf = await buildPdf(body, subjectName); } catch (e) { context.log.error(e); status.notes.push("pdf failed"); }

  // 3) File the PDF into the staff member's review folder.
  if (pdf && token && SITE_PATH && BASE_PATH) {
    try {
      const driveId = await resolveDrive(token);
      const staff = await findStaffFolder(token, driveId, subjectName, isSubjectForm ? "" : p.email);
      if (staff) {
        const wantKra = /^A$/i.test(body.form || "");
        const sub = await ensureSubfolder(token, driveId, trim(BASE_PATH) + "/" + staff, wantKra ? "KRA" : "Development Review");
        await uploadPdf(token, driveId, trim(BASE_PATH) + "/" + staff + "/" + sub + "/" + fileBase + ".pdf", pdf);
        status.filed = true;
      } else status.notes.push("staff folder not found");
    } catch (e) { context.log.error(e); status.notes.push("file failed"); }
  }

  // 4) Email a copy. HR by default; Form D also copies the staff member (resolved
  //    from the typed name) and the person completing it (the manager in the review).
  let recipients = NOTIFY_TO ? NOTIFY_TO.split(",").map(function (s) { return s.trim(); }).filter(Boolean) : [];
  if (isGrowthForm) {
    if (p.email) recipients.push(p.email);
    if (isSubjectForm && subjectName && subjectName.toLowerCase() !== String(myName).toLowerCase() && token) {
      try { const su = await findUserByName(token, subjectName); const se = su && (su.mail || su.userPrincipalName); if (se) recipients.push(se); }
      catch (e) { context.log.error(e); status.notes.push("staff email lookup failed"); }
    }
  }
  recipients = uniqEmails(recipients);
  // Send AS the person completing the form (their mailbox), so HR receives it from them
  // and replies go straight back. Falls back to REVIEW_NOTIFY_FROM if no submitter address.
  const notifyFrom = p.email || NOTIFY_FROM;
  if (pdf && token && notifyFrom && recipients.length) {
    try {
      const subject = (body.title || "Performance review") + " — " + subjectName + (body.quarter ? " (" + body.quarter + ")" : "");
      const html = "<p>A copy of the completed performance-review form is attached.</p><ul>" +
        "<li><b>Form:</b> " + esc(body.title || body.form) + (body.version ? " (v" + esc(body.version) + ")" : "") + "</li>" +
        "<li><b>Staff member:</b> " + esc(subjectName) + "</li>" +
        "<li><b>Completed by:</b> " + esc(myName) + " (" + esc(p.email) + ")</li>" +
        "<li><b>Quarter:</b> " + esc(body.quarter || "—") + "</li></ul>";
      await sendMail(token, notifyFrom, recipients.join(","), subject, html, pdf, fileBase + ".pdf");
      status.emailed = true;
    } catch (e) { context.log.error(e); status.notes.push("email failed"); }
  }

  status.ok = status.stored || status.filed || status.emailed;
  context.res = json(status.ok ? 200 : 502, status);
};

/* ---------- PDF — branded, mimics the Word KRA template ---------- */
function loadLogo() {
  try { return require("fs").readFileSync(require("path").join(__dirname, "logo.png")); } catch (e) { return null; }
}
function buildPdf(body, name) {
  return new Promise(function (resolve, reject) {
    try {
      const logo = loadLogo();
      const doc = new PDFDocument({ size: "A4", margin: 50, bufferPages: true });
      const chunks = [];
      doc.on("data", function (d) { chunks.push(d); });
      doc.on("end", function () { resolve(Buffer.concat(chunks)); });

      const NAVY = "#14294D", GOLD = "#E5A823", GOLD_DARK = "#C8901A",
        INK = "#1F2A37", MUTED = "#5B6675", FAINT = "#9AA4B2",
        LINE = "#D8DEE7", SHADE = "#F1F4F8", HILITE = "#FBF2DA";
      const L = 50, R = 545, W = R - L, BOTTOM = 800;
      const titleText = body.title || "Performance review";
      const intro = (body.intro || "").trim();
      const logoH = logo ? 150 * 352 / 982 : 0;

      function ensure(h) { if (doc.y + h > BOTTOM) { doc.addPage(); slimHeader(); } }

      // Full brand band (page 1)
      function topBand() {
        var topY = 42, logoBottom = topY;
        if (logo) { try { doc.image(logo, L, topY, { width: 150 }); logoBottom = topY + logoH; } catch (e) {} }
        doc.font("Helvetica-Oblique").fontSize(8.5).fillColor(MUTED)
          .text("Flexible. Dependable. On-site Engineering.", L, topY + 2, { width: W, align: "right" });
        doc.font("Helvetica").fontSize(8).fillColor(MUTED)
          .text("HR & Payroll · Development Review & KRA", L, doc.y + 1, { width: W, align: "right" });
        if (body.version) doc.font("Helvetica").fontSize(7.5).fillColor(FAINT)
          .text("Form version " + body.version, L, doc.y + 1, { width: W, align: "right" });
        var y = Math.max(logoBottom, doc.y) + 9;
        doc.moveTo(L, y).lineTo(R, y).lineWidth(1.4).strokeColor(NAVY).stroke();
        return y + 12;
      }
      // Slim running header (pages 2+)
      function slimHeader() {
        var y = 40;
        if (logo) { try { doc.image(logo, L, y, { width: 92 }); } catch (e) {} }
        doc.font("Helvetica-Bold").fontSize(9).fillColor(NAVY).text(titleText, L, y + 8, { width: W, align: "right" });
        var yy = y + 34;
        doc.moveTo(L, yy).lineTo(R, yy).lineWidth(1).strokeColor(LINE).stroke();
        doc.y = yy + 12;
      }
      function titleBlock(startY) {
        doc.font("Helvetica-Bold").fontSize(15).fillColor(NAVY).text(titleText, L, startY, { width: W });
        var y = doc.y + 3;
        doc.moveTo(L, y).lineTo(R, y).lineWidth(2).strokeColor(GOLD).stroke();
        var yy = y + 8;
        if (intro) { doc.font("Helvetica-Oblique").fontSize(9.5).fillColor(MUTED).text(intro, L, yy, { width: W }); yy = doc.y + 4; }
        doc.y = yy + 6;
      }
      function infoTable() {
        var rows = [["Name", name], ["Quarter", body.quarter || "—"], ["Date", body.date || today()]];
        var labelW = 120, rowH = 22;
        rows.forEach(function (r) {
          ensure(rowH); var y = doc.y;
          doc.rect(L, y, labelW, rowH).fillAndStroke(SHADE, LINE);
          doc.rect(L + labelW, y, W - labelW, rowH).lineWidth(1).strokeColor(LINE).stroke();
          doc.fillColor(NAVY).font("Helvetica-Bold").fontSize(9.5).text(r[0], L + 8, y + 6, { width: labelW - 12 });
          doc.fillColor(INK).font("Helvetica").fontSize(9.5).text(String(r[1]), L + labelW + 8, y + 6, { width: W - labelW - 14 });
          doc.y = y + rowH;
        });
        doc.y += 12;
      }
      function scaleBox(scale) {
        if (!scale || !scale.length) return;
        ensure(30);
        doc.font("Helvetica-Bold").fontSize(10.5).fillColor(NAVY).text("Rating scale — how to read the scores", L, doc.y, { width: W });
        var y = doc.y + 2; doc.moveTo(L, y).lineTo(R, y).lineWidth(1).strokeColor(GOLD).stroke(); doc.y = y + 6;
        var numW = 26;
        scale.forEach(function (s) {
          doc.font("Helvetica").fontSize(9);
          var dh = doc.heightOfString(s.desc, { width: W - numW - 16 });
          var rowH = Math.max(20, dh + 8);
          ensure(rowH); var yy = doc.y;
          doc.rect(L, yy, numW, rowH).fillAndStroke(NAVY, NAVY);
          doc.rect(L + numW, yy, W - numW, rowH).lineWidth(0.8).strokeColor(LINE).stroke();
          doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(10).text(String(s.n), L, yy + (rowH - 11) / 2, { width: numW, align: "center" });
          doc.fillColor(INK).font("Helvetica").fontSize(9).text(s.desc, L + numW + 8, yy + (rowH - dh) / 2, { width: W - numW - 16 });
          doc.y = yy + rowH;
        });
        doc.y += 12;
      }
      function sectionTitle(t) {
        ensure(34); doc.moveDown(0.2);
        doc.font("Helvetica-Bold").fontSize(11.5).fillColor(NAVY).text(t, L, doc.y, { width: W });
        var y = doc.y + 2; doc.moveTo(L, y).lineTo(R, y).lineWidth(1).strokeColor(GOLD).stroke();
        doc.y = y + 8;
      }
      function scoreRow(label, value, max) {
        max = max || 5;
        var labelW = W * 0.56;
        doc.font("Helvetica-Bold").fontSize(9);
        var lblH = doc.heightOfString(label, { width: labelW - 14 });
        var rowH = Math.max(30, lblH + 12);
        ensure(rowH); var y = doc.y, scaleX = L + labelW, cellW = (W - labelW) / max;
        doc.lineWidth(0.8).strokeColor(LINE);
        doc.rect(L, y, labelW, rowH).stroke();
        for (var i = 0; i < max; i++) doc.rect(scaleX + i * cellW, y, cellW, rowH).stroke();
        doc.fillColor(INK).font("Helvetica-Bold").fontSize(9).text(label, L + 8, y + (rowH - lblH) / 2, { width: labelW - 14 });
        var sel = parseInt(value, 10);
        for (var n = 1; n <= max; n++) {
          var cx = scaleX + (n - 1) * cellW + cellW / 2, cy = y + rowH / 2, isSel = (n === sel);
          doc.font(isSel ? "Helvetica-Bold" : "Helvetica").fontSize(10).fillColor(isSel ? NAVY : MUTED)
            .text(String(n), scaleX + (n - 1) * cellW, cy - 6, { width: cellW, align: "center" });
          if (isSel) doc.lineWidth(1.4).strokeColor(GOLD_DARK).circle(cx, cy, 9).stroke();
        }
        doc.y = y + rowH;
      }
      function optionRow(label, value, options) {
        options = options || [];
        var labelW = W * 0.40, optX = L + labelW, optW = W - labelW, txtX = optX + 22, txtW = optW - 32;
        doc.font("Helvetica").fontSize(9);
        var heights = options.map(function (o) { return Math.max(15, doc.heightOfString(o, { width: txtW }) + 5); });
        var bodyH = heights.reduce(function (a, b) { return a + b; }, 0) + 10;
        doc.font("Helvetica-Bold").fontSize(9);
        var lblH = doc.heightOfString(label, { width: labelW - 14 });
        var rowH = Math.max(bodyH, lblH + 12, 26);
        ensure(rowH); var y = doc.y;
        doc.lineWidth(0.8).strokeColor(LINE);
        doc.rect(L, y, labelW, rowH).stroke();
        doc.rect(optX, y, optW, rowH).stroke();
        doc.fillColor(INK).font("Helvetica-Bold").fontSize(9).text(label, L + 8, y + 6, { width: labelW - 14 });
        var oy = y + 6;
        options.forEach(function (o, i) {
          var isSel = (o === value);
          if (isSel) doc.rect(optX + 4, oy - 1, optW - 8, heights[i]).fillColor(HILITE).fill();
          // marker drawn as a vector dot (Unicode bullets aren't in the standard PDF font)
          var mx = optX + 11, my = oy + 8;
          if (isSel) doc.circle(mx, my, 3.2).fillColor(NAVY).fill();
          else doc.circle(mx, my, 3).lineWidth(0.8).strokeColor(MUTED).stroke();
          doc.font(isSel ? "Helvetica-Bold" : "Helvetica").fontSize(9).fillColor(isSel ? NAVY : MUTED)
            .text(o, txtX, oy + 2, { width: txtW });
          oy += heights[i];
        });
        doc.y = y + rowH;
      }
      function avgRow(value, count) {
        ensure(24); var y = doc.y, h = 22;
        doc.rect(L, y, W, h).fillAndStroke(SHADE, GOLD);
        doc.fillColor(NAVY).font("Helvetica-Bold").fontSize(9.5).text("Section average", L + 8, y + 6, { width: W / 2 });
        doc.fillColor(GOLD_DARK).font("Helvetica-Bold").fontSize(11)
          .text(String(value) + (count ? "   (avg of " + count + ")" : ""), L, y + 5, { width: W - 10, align: "right" });
        doc.y = y + h + 10;
      }
      function textRow(label, value) {
        value = (value == null || value === "") ? "—" : String(value);
        var labelW = W * 0.30, ansX = L + labelW, ansW = W - labelW;
        doc.font("Helvetica").fontSize(9.5);
        var ansH = doc.heightOfString(value, { width: ansW - 16 });
        doc.font("Helvetica-Bold").fontSize(9.5);
        var lblH = doc.heightOfString(label, { width: labelW - 14 });
        var rowH = Math.max(ansH, lblH) + 14;
        ensure(Math.min(rowH, BOTTOM - 110)); var y = doc.y;
        doc.rect(L, y, labelW, rowH).fillAndStroke(SHADE, LINE);
        doc.rect(ansX, y, ansW, rowH).lineWidth(0.8).strokeColor(LINE).stroke();
        doc.fillColor(NAVY).font("Helvetica-Bold").fontSize(9.5).text(label, L + 8, y + 7, { width: labelW - 14 });
        doc.fillColor(INK).font("Helvetica").fontSize(9.5).text(value, ansX + 8, y + 7, { width: ansW - 16 });
        doc.y = y + rowH;
      }

      var hasScore = (body.sections || []).some(function (sec) { return (sec.rows || []).some(function (r) { return r.type === "score"; }); });
      titleBlock(topBand());
      infoTable();
      if (hasScore) scaleBox(body.scale);
      (body.sections || []).forEach(function (sec) {
        sectionTitle(sec.heading || "");
        (sec.rows || []).forEach(function (r) {
          if (r.type === "score") scoreRow(r.label, r.value, r.max || 5);
          else if (r.type === "option") optionRow(r.label, r.value, r.options || []);
          else if (r.type === "avg") avgRow(r.value, r.count);
          else textRow(r.label, r.value);
        });
        doc.y += 6;
      });

      // Footer on every page. Drop the bottom margin first so writing into the
      // footer band doesn't make PDFKit auto-insert blank pages.
      var range = doc.bufferedPageRange();
      for (var i = 0; i < range.count; i++) {
        doc.switchToPage(range.start + i);
        doc.page.margins.bottom = 0;
        var fy = doc.page.height - 30;
        doc.font("Helvetica").fontSize(7.5).fillColor(FAINT)
          .text("King EPCM · Confidential — Development Review & KRA", L, fy, { width: W, align: "left", lineBreak: false });
        doc.text("Page " + (i + 1) + " of " + range.count, L, fy, { width: W, align: "right", lineBreak: false });
      }
      doc.end();
    } catch (e) { reject(e); }
  });
}

/* ---------- Graph ---------- */
async function appToken() {
  const b = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: b });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function lookupUser(token, email) {
  return gget(token, "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(email) + "?$select=displayName,givenName,surname");
}
// Resolve a user (email) from a typed display name. Tries an exact displayName
// match first, then a fuzzy $search. Needs User.Read.All (already required).
async function findUserByName(token, name) {
  const n = String(name || "").trim();
  if (!n) return null;
  try {
    const r = await gget(token, "https://graph.microsoft.com/v1.0/users?$filter=" +
      encodeURIComponent("displayName eq '" + n.replace(/'/g, "''") + "'") + "&$select=displayName,mail,userPrincipalName&$top=1");
    if (r.value && r.value.length) return r.value[0];
  } catch (e) {}
  try {
    const r = await fetch("https://graph.microsoft.com/v1.0/users?$search=" +
      encodeURIComponent('"displayName:' + n + '"') + "&$select=displayName,mail,userPrincipalName&$top=5",
      { headers: { Authorization: "Bearer " + token, ConsistencyLevel: "eventual" } });
    if (r.ok) { const j = await r.json(); if (j.value && j.value.length) return j.value[0]; }
  } catch (e) {}
  return null;
}
async function resolveDrive(token) {
  if (_driveId) return _driveId;
  const site = await gget(token, "https://graph.microsoft.com/v1.0/sites/" + SITE_PATH + "?$select=id");
  if (LIBRARY) {
    const ds = await gget(token, "https://graph.microsoft.com/v1.0/sites/" + site.id + "/drives?$select=id,name");
    const d = (ds.value || []).find(function (x) { return (x.name || "").toLowerCase() === LIBRARY.toLowerCase(); });
    if (d) { _driveId = d.id; return _driveId; }
  }
  const drive = await gget(token, "https://graph.microsoft.com/v1.0/sites/" + site.id + "/drive?$select=id");
  _driveId = drive.id; return _driveId;
}
async function findStaffFolder(token, driveId, name, email) {
  if (email && FOLDER_MAP[email.toLowerCase()]) return FOLDER_MAP[email.toLowerCase()];
  const kids = (await listChildren(token, driveId, trim(BASE_PATH))).filter(function (i) { return i.folder; });
  const want = toks(name);
  let best = null, bestScore = 0;
  kids.forEach(function (f) {
    const ft = toks(deNum(f.name));
    const shared = want.filter(function (t) { return ft.indexOf(t) > -1; }).length;
    if (shared >= 2 && shared > bestScore) { bestScore = shared; best = f.name; }
  });
  return best;
}
async function ensureSubfolder(token, driveId, parentPath, name) {
  const kids = await listChildren(token, driveId, parentPath).catch(function () { return []; });
  const existing = kids.find(function (i) { return i.folder && SUB_RE.test(deNum(i.name)) && deNum(i.name).toLowerCase().indexOf(name.toLowerCase().split(" ")[0]) > -1; });
  if (existing) return existing.name;
  const anyReview = kids.find(function (i) { return i.folder && SUB_RE.test(deNum(i.name)); });
  if (anyReview) return anyReview.name;
  // none exists → create it
  const enc = parentPath.split("/").map(encodeURIComponent).join("/");
  await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + enc + ":/children", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ name: name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" })
  }).catch(function () {});
  return name;
}
async function uploadPdf(token, driveId, path, buffer) {
  const enc = path.split("/").map(encodeURIComponent).join("/");
  const r = await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + enc + ":/content", {
    method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": "application/pdf" }, body: buffer
  });
  if (!r.ok) throw new Error("upload " + r.status);
}
async function sendMail(token, from, to, subject, html, pdf, filename) {
  const message = {
    subject: subject, body: { contentType: "HTML", content: html },
    toRecipients: to.split(",").map(function (e) { return { emailAddress: { address: e.trim() } }; }),
    attachments: [{ "@odata.type": "#microsoft.graph.fileAttachment", name: filename, contentType: "application/pdf", contentBytes: pdf.toString("base64") }]
  };
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: message, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status);
}
async function listChildren(token, driveId, path) {
  const enc = path.split("/").map(encodeURIComponent).join("/");
  const d = await gget(token, "https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + enc + ":/children?$top=200");
  return d.value || [];
}
async function gget(token, url) {
  const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) throw new Error("graph " + r.status);
  return r.json();
}

/* ---------- helpers ---------- */
function principal(req) {
  try {
    const h = req.headers["x-ms-client-principal"];
    if (!h) return null;
    const o = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    const claims = o.claims || [];
    const cv = function (types) { const c = claims.find(function (x) { return types.indexOf((x.typ || "").toLowerCase()) > -1 && x.val; }); return c ? c.val : ""; };
    let email = (o.userDetails && o.userDetails.indexOf("@") > -1) ? o.userDetails : cv(["email", "preferred_username", "upn"]);
    const name = cv(["name"]);
    return { email: email, name: (name && name.indexOf("@") === -1) ? name : "", roles: o.userRoles || [] };
  } catch (e) { return null; }
}
function deNum(s) { return String(s || "").replace(/^\s*\d+\s*[.)\-:_]*\s*/, ""); }
function toks(s) { return String(s || "").toLowerCase().normalize("NFKD").replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter(function (t) { return t.length > 1; }); }
function sanitize(s) { return String(s || "form").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 120); }
function today() { return new Date().toLocaleDateString("en-CA"); }
function esc(s) { return String(s == null ? "" : s).replace(/[&<>]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]; }); }
function trim(s) { return String(s || "").replace(/^\/+|\/+$/g, ""); }
function uniqEmails(arr) { const seen = {}, out = []; (arr || []).forEach(function (e) { e = String(e || "").trim(); const k = e.toLowerCase(); if (e && !seen[k]) { seen[k] = 1; out.push(e); } }); return out; }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
