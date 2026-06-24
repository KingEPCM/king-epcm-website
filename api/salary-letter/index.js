/*
 * POST /api/salary-letter  (HR / admin only)
 * Generates a branded King EPCM letter for either a SALARY INCREASE or a PROMOTION as a PDF.
 * Returns the PDF for immediate download AND files a dated copy into the employee's HR folder
 * (the same Employee Records folders that hold KRA / review files), matched by name (best-effort).
 *
 * Body (JSON):
 *   { kind: "raise"|"promotion", name, effectiveDate,
 *     salaryBasis:"year"|"hour", prevAmount, newAmount,         // raise (and optional on promotion)
 *     prevPosition, newPosition,                                 // promotion
 *     reason, signedByName, signedByTitle, signedByPhone, signedByEmail }
 *
 * Gated to the SWA roles "hr" or "admin". Reuses the reviews/Graph settings (see employment-letter).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.REVIEW_SITE_PATH;
const BASE_PATH = process.env.REVIEW_BASE_PATH;
const LIBRARY = process.env.REVIEW_LIBRARY || "";
const SUBFOLDER = process.env.EMPLOYMENT_LETTER_SUBFOLDER || ""; // blank = staff folder root
let _driveId;

module.exports = async function (context, req) {
 try {
  const roles = principalRoles(req);
  if (roles.indexOf("hr") === -1 && roles.indexOf("admin") === -1) {
    context.res = json(403, { ok: false, error: "HR only" }); return;
  }
  const b = req.body || {};
  const kind = /promo/i.test(b.kind) ? "promotion" : "raise";
  const name = (b.name || "").trim();
  if (!name) { context.res = json(400, { ok: false, error: "Employee name is required" }); return; }
  if (kind === "raise" && (!parseNum(b.prevAmount) || !parseNum(b.newAmount))) {
    context.res = json(400, { ok: false, error: "Previous and new amounts are required for a salary increase" }); return;
  }
  if (kind === "promotion" && !(b.newPosition || "").trim()) {
    context.res = json(400, { ok: false, error: "New position is required for a promotion" }); return;
  }

  const docTitle = kind === "promotion" ? "Promotion Letter" : "Salary Adjustment Letter";
  const pdf = await buildLetter(kind, docTitle, b, loadLogo());

  const base = sanitize(docTitle + " - " + name + " - " + shortToday());
  const out = { ok: true, filename: base + ".pdf", pdfBase64: pdf.toString("base64"), filed: false, filedTo: null, notes: [] };

  if (CLIENT_ID && CLIENT_SECRET && SITE_PATH && BASE_PATH) {
    try {
      const token = await appToken();
      const driveId = await resolveDrive(token);
      const staff = await findStaffFolder(token, driveId, name);
      if (staff) {
        let folderPath = trim(BASE_PATH) + "/" + staff;
        if (SUBFOLDER) {
          const sub = await ensureSubfolder(token, driveId, folderPath, SUBFOLDER).catch(function () { return null; });
          if (sub) folderPath += "/" + sub;
        }
        await uploadPdf(token, driveId, folderPath + "/" + base + ".pdf", pdf);
        out.filed = true; out.filedTo = staff + (SUBFOLDER ? " / " + SUBFOLDER : "");
      } else out.notes.push("no folder match");
    } catch (e) { context.log.error(e); out.notes.push("file failed"); }
  } else out.notes.push("filing not configured");

  context.res = json(200, out);
 } catch (e) {
  context.log.error(e);
  context.res = json(500, { ok: false, error: String((e && e.message) || e) });
 }
};

function loadLogo() {
  try { return require("fs").readFileSync(require("path").join(__dirname, "logo.png")); }
  catch (e) { return null; }
}

function buildLetter(kind, docTitle, b, logo) {
  return new Promise(function (resolve, reject) {
    try {
      const PDFDocument = require("pdfkit");
      const doc = new PDFDocument({ size: "LETTER", margins: { top: 50, left: 64, right: 64, bottom: 96 } });
      const chunks = [];
      doc.on("data", function (d) { chunks.push(d); });
      doc.on("end", function () { resolve(Buffer.concat(chunks)); });

      const NAVY = "#14294D", GOLD = "#E5A823", INK = "#222", GREY = "#666";
      const W = doc.page.width, H = doc.page.height, L = 64, R = W - 64, CW = R - L;
      const name = (b.name || "").trim(), fn = firstName(name);
      const eff = b.effectiveDate ? ordinalDate(parseDate(b.effectiveDate)) : "";
      const basis = String(b.salaryBasis || "year").toLowerCase() === "hour" ? "hour" : "year";
      const prev = parseNum(b.prevAmount), next = parseNum(b.newAmount);
      const hasPay = prev && next;
      const reason = (b.reason || "").trim();

      // ---- Letterhead ----
      let drew = false;
      if (logo) { try { doc.image(logo, L, 44, { width: 165 }); drew = true; } catch (e) {} }
      if (!drew) {
        doc.fillColor(GOLD).font("Helvetica-Bold").fontSize(26).text("KING EPCM", L, 44);
        doc.fillColor(NAVY).font("Helvetica-Bold").fontSize(8).text("Flexible. Dependable. On-site Engineering.", L, 80);
      }
      doc.fillColor(NAVY).font("Times-Roman").fontSize(10).text(docTitle, L, 61, { width: CW, align: "right" });
      if (name) doc.fillColor(GREY).text(name, L, 75, { width: CW, align: "right" });
      doc.moveTo(L, 120).lineTo(R, 120).lineWidth(2.5).strokeColor(GOLD).stroke();
      doc.moveTo(L, 124.5).lineTo(R, 124.5).lineWidth(0.8).strokeColor(NAVY).stroke();

      // ---- Date ----
      doc.fillColor(INK).font("Times-Roman").fontSize(12).text(ordinalDate(new Date()), L, 144);
      doc.moveDown(1);

      // ---- RE + salutation (addressed to the employee) ----
      doc.font("Times-Bold").fontSize(12).fillColor(INK).text("RE:    " + docTitle, L);
      if (name) doc.text("          " + name, L);
      doc.moveDown(0.8);
      doc.font("Times-Roman").fontSize(12).fillColor(INK).text("Dear " + fn + ",", { paragraphGap: 12 });

      const basisWord = basis === "hour" ? "hourly rate of pay" : "annual salary";
      const per = basis === "hour" ? " per hour" : "";
      const pct = pctText(prev, next);

      if (kind === "promotion") {
        const prevPos = (b.prevPosition || "").trim(), newPos = (b.newPosition || "").trim();
        let p1 = "On behalf of King EPCM, I am pleased to confirm your promotion to the position of " + newPos +
          (prevPos ? ", from your previous role as " + prevPos : "") + (eff ? ", effective " + eff : "") + ".";
        doc.text(p1, { paragraphGap: 12, lineGap: 3 });
        if (hasPay) {
          doc.text("In connection with this promotion, your " + basisWord + " will increase from " + money(prev) + per + " to " + money(next) + per +
            (pct ? ", an increase of " + pct + "%" : "") + (eff ? ", effective " + eff : "") + ". All amounts are before taxes and applicable deductions.", { paragraphGap: 12, lineGap: 3 });
        }
        if (reason) doc.text(reason, { paragraphGap: 12, lineGap: 3 });
        doc.text("This promotion reflects our appreciation of your contributions and the value you bring to the team. Congratulations, and thank you for your continued commitment to King EPCM.", { paragraphGap: 12, lineGap: 3 });
      } else {
        doc.text("On behalf of King EPCM, I am pleased to inform you that your " + basisWord + " will be adjusted" + (eff ? ", effective " + eff : "") + " as follows:", { paragraphGap: 10, lineGap: 3 });
        // Simple figures block
        doc.font("Times-Roman").fontSize(12);
        doc.text("Previous " + basisWord + ":     " + money(prev) + per, L + 24, doc.y, { lineGap: 2 });
        doc.text("New " + basisWord + ":            " + money(next) + per, L + 24, doc.y, { lineGap: 2 });
        if (pct) doc.font("Times-Bold").text("Increase:                       " + pct + "%", L + 24, doc.y, { lineGap: 2 });
        doc.font("Times-Roman").moveDown(0.6);
        doc.text("All amounts are stated before taxes and applicable deductions." + (reason ? " " + reason : ""), L, doc.y, { paragraphGap: 12, lineGap: 3 });
        doc.text("Thank you for your continued contributions and commitment to King EPCM. We look forward to your ongoing success with the team.", { paragraphGap: 12, lineGap: 3 });
      }

      doc.text("Should you have any questions, please feel free to contact the undersigned.", { paragraphGap: 26, lineGap: 3 });

      // ---- Signed block ----
      doc.text("Sincerely,");
      doc.moveDown(3.2);
      const sName = (b.signedByName || "").trim() || "Angela Shi";
      const sTitle = (b.signedByTitle || "").trim() || "Operation Manager";
      const sPhone = (b.signedByPhone || "").trim() || "416-342-3001 x109";
      const sEmail = (b.signedByEmail || "").trim() || "AShi@KingEPCM.com";
      doc.font("Times-Roman").fontSize(12).fillColor(INK).text(sName, L);
      doc.font("Times-Italic").fontSize(12).text(sTitle, L);
      doc.font("Times-Roman").fontSize(12).text("T: " + sPhone, L);
      doc.text("E: " + sEmail, L);

      // ---- Navy footer band ----
      doc.page.margins.bottom = 0;
      var bandH = 46;
      doc.rect(0, H - bandH, W, bandH).fill(NAVY);
      doc.fillColor("#FFFFFF").font("Times-Roman").fontSize(9)
        .text("King EPCM   ·   3780 14th Avenue, Unit 211, Markham, ON  L3R 9Y5", 0, H - bandH + 13, { width: W, align: "center", lineBreak: false });
      doc.fillColor(GOLD).font("Times-Roman").fontSize(9)
        .text("www.KingEPCM.com", 0, H - bandH + 26, { width: W, align: "center", lineBreak: false });

      doc.end();
    } catch (e) { reject(e); }
  });
}

/* ---------- Graph: file the PDF into the staff member's HR folder ---------- */
async function appToken() {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
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
async function findStaffFolder(token, driveId, name) {
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
  const existing = kids.find(function (i) { return i.folder && deNum(i.name).toLowerCase() === name.toLowerCase(); });
  if (existing) return existing.name;
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
function deNum(s) { return String(s || "").replace(/^\s*\d+\s*[.)\-:_]*\s*/, ""); }
function toks(s) { return String(s || "").toLowerCase().normalize("NFKD").replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter(function (t) { return t.length > 1; }); }
function trim(s) { return String(s || "").replace(/^\/+|\/+$/g, ""); }
function shortToday() { var d = new Date(); return String(d.getFullYear()).slice(-2) + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0"); }

/* ---------- helpers ---------- */
function parseNum(v) { const n = Number(String(v == null ? "" : v).replace(/[^0-9.]/g, "")); return isNaN(n) ? 0 : n; }
function money(n) { return "$" + Number(n || 0).toLocaleString("en-CA", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function pctText(prev, next) {
  if (!prev || !next) return "";
  const p = ((next - prev) / prev) * 100;
  if (!isFinite(p)) return "";
  const r = Math.round(p * 10) / 10;
  return (r % 1 === 0 ? String(r) : r.toFixed(1));
}
function parseDate(s) { const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(s); }
function ordinalDate(d) {
  try {
    var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var day = d.getDate(), t = day % 10, h = day % 100;
    var suf = (t === 1 && h !== 11) ? "st" : (t === 2 && h !== 12) ? "nd" : (t === 3 && h !== 13) ? "rd" : "th";
    return months[d.getMonth()] + " " + day + suf + " " + d.getFullYear();
  } catch (e) { return new Date().toDateString(); }
}
function firstName(n) { return String(n || "").trim().split(/\s+/)[0] || n; }
function principalRoles(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return [];
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    return Array.isArray(p.userRoles) ? p.userRoles : [];
  } catch (e) { return []; }
}
function sanitize(s) { return String(s || "letter").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 120); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
