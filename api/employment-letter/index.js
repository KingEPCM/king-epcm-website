/*
 * POST /api/employment-letter  (HR / admin only)
 * Generates a branded King EPCM employment / income-verification letter as a PDF.
 * HR fills the form; the function returns the PDF for immediate download AND files a
 * dated copy into the employee's HR folder (the same Employee Records folders that hold
 * KRA / review files), matched by name. Filing is best-effort. No payroll integration.
 *
 * Body (JSON): { name, position, employmentType, startDate, salaryAmount,
 *                salaryBasis: "year"|"hour", signedByName, signedByTitle, ... }
 *
 * Gated to the SWA roles "hr" or "admin".
 * App settings (reused from the reviews setup): AAD_CLIENT_ID, AAD_CLIENT_SECRET,
 *   GRAPH_TENANT, REVIEW_SITE_PATH, REVIEW_BASE_PATH, REVIEW_LIBRARY (optional),
 *   EMPLOYMENT_LETTER_SUBFOLDER (optional — a subfolder inside the staff folder).
 * Graph APPLICATION permission to file the PDF: Sites.ReadWrite.All (admin consent).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.REVIEW_SITE_PATH;
const BASE_PATH = process.env.REVIEW_BASE_PATH;
const LIBRARY = process.env.REVIEW_LIBRARY || "";
const SUBFOLDER = process.env.EMPLOYMENT_LETTER_SUBFOLDER || ""; // blank = staff folder root
const { buildSimpleLetterDocx } = require("../_shared/letterDocx");
let _driveId;

module.exports = async function (context, req) {
 try {
  const roles = principalRoles(req);
  if (roles.indexOf("hr") === -1 && roles.indexOf("admin") === -1) {
    context.res = json(403, { ok: false, error: "HR only" }); return;
  }
  const b = req.body || {};
  const name = (b.name || "").trim();
  if (!name) { context.res = json(400, { ok: false, error: "Employee name is required" }); return; }

  const logo = loadLogo();
  const model = letterModel(b);
  const pdf = await buildPdf(model, logo);
  const docx = buildSimpleLetterDocx({ docTitle: model.docTitle, name: model.name, dateStr: ordinalDate(new Date()), salutation: model.salutation, paragraphs: model.paragraphs, closing: model.closing, signer: model.signer, logo: logo });
  const base = sanitize("Employment Letter - " + name + " - " + shortToday());
  const out = { ok: true, filename: base + ".pdf", pdfBase64: pdf.toString("base64"), docxFilename: base + ".docx", docxBase64: docx.toString("base64"), filed: false, filedTo: null, notes: [] };

  // Best-effort: file a dated copy into the staff member's HR folder (name match).
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

// Read the letterhead logo bundled alongside this function (no network/auth dependency —
// the intranet is sign-in protected, so an HTTP fetch of /assets would hit the login page).
function loadLogo() {
  try { return require("fs").readFileSync(require("path").join(__dirname, "logo.png")); }
  catch (e) { return null; }
}

// Build the letter content once; both the PDF and the Word renderer consume this model.
function letterModel(b) {
  const name = (b.name || "").trim(), fn = firstName(name);
  const position = (b.position || "").trim(), etype = (b.employmentType || "").trim();
  const start = b.startDate ? ordinalDate(parseDate(b.startDate)) : "";
  const dob = b.dateOfBirth ? isoDate(b.dateOfBirth) : "";
  const pay = payPhrase(b.salaryAmount, b.salaryBasis);
  const ohip = !!b.ohip, addr = (b.employeeAddress || "").trim();
  const paras = [];
  if (ohip) {
    let p1 = "This letter confirms that " + name + (dob ? " (date of birth: " + dob + ")" : "") + " is employed with King EPCM on a full-time basis" + (position ? " in the position of " + position : "") + ".";
    if (start) p1 += " " + fn + " has been employed with King EPCM since " + start + ".";
    paras.push({ text: p1 });
    paras.push({ text: "King EPCM intends to employ " + fn + " on a full-time basis for a minimum of six (6) months." });
    if (pay) paras.push({ text: name + "'s current " + pay + "." });
    if (addr) paras.push({ text: fn + "'s residential address is " + addr + "." });
    paras.push({ text: "This letter is provided to confirm employment, including for the purpose of Ontario Health Insurance Plan (OHIP) eligibility. Should you have any questions or inquiries regarding this information, please feel free to contact the undersigned at any time." });
  } else {
    let p1 = "This letter confirms that " + name + (dob ? " (date of birth: " + dob + ")" : "") + " is employed with King EPCM";
    if (position) p1 += " in the position of " + position;
    if (etype) p1 += " on a " + etype.toLowerCase() + " basis";
    p1 += ".";
    if (start) p1 += " " + fn + " has been employed with King EPCM since " + start + ".";
    paras.push({ text: p1 });
    if (pay) paras.push({ text: name + "'s current " + pay + "." });
    paras.push({ text: "Should you have any questions or inquiries regarding this information, please feel free to contact the undersigned at any time." });
  }
  return {
    docTitle: "Employment Confirmation Letter", name: name,
    salutation: "To whom it may concern,", closing: "Regards,", paragraphs: paras,
    signer: {
      name: (b.signedByName || "").trim() || "Angela Shi",
      title: (b.signedByTitle || "").trim() || "Operation Manager",
      phone: (b.signedByPhone || "").trim() || "416-342-3001 x109",
      email: (b.signedByEmail || "").trim() || "AShi@KingEPCM.com"
    }
  };
}

// Render the model to a branded PDF (King EPCM letterhead, gold rule, navy footer).
function buildPdf(model, logo) {
  return new Promise(function (resolve, reject) {
    try {
      const PDFDocument = require("pdfkit");
      const doc = new PDFDocument({ size: "LETTER", margins: { top: 50, left: 64, right: 64, bottom: 96 } });
      const chunks = [];
      doc.on("data", function (d) { chunks.push(d); });
      doc.on("end", function () { resolve(Buffer.concat(chunks)); });

      const NAVY = "#14294D", GOLD = "#E5A823", INK = "#222", GREY = "#666";
      const W = doc.page.width, H = doc.page.height, L = 64, R = W - 64, CW = R - L;
      const title = model.docTitle, name = model.name;

      let drew = false;
      if (logo) { try { doc.image(logo, L, 44, { width: 165 }); drew = true; } catch (e) {} }
      if (!drew) {
        doc.fillColor(GOLD).font("Times-Bold").fontSize(26).text("KING EPCM", L, 44);
        doc.fillColor(NAVY).font("Times-Bold").fontSize(8).text("Flexible. Dependable. On-site Engineering.", L, 80);
      }
      doc.fillColor(NAVY).font("Times-Roman").fontSize(10).text(title, L, 61, { width: CW, align: "right" });
      if (name) doc.fillColor(GREY).text(name, L, 75, { width: CW, align: "right" });
      doc.moveTo(L, 120).lineTo(R, 120).lineWidth(2.5).strokeColor(GOLD).stroke();
      doc.moveTo(L, 124.5).lineTo(R, 124.5).lineWidth(0.8).strokeColor(NAVY).stroke();

      doc.fillColor(INK).font("Times-Roman").fontSize(12).text(ordinalDate(new Date()), L, 144);
      doc.moveDown(1);
      doc.font("Times-Bold").fontSize(12).fillColor(INK).text("RE:    " + title, L);
      if (name) doc.text("          " + name, L);
      doc.moveDown(0.8);
      doc.font("Times-Roman").fontSize(12).fillColor(INK).text(model.salutation, { paragraphGap: 12 });

      model.paragraphs.forEach(function (p) {
        doc.font(p.bold ? "Times-Bold" : "Times-Roman").fontSize(12).fillColor(INK)
          .text(p.text, L + (p.indent ? 24 : 0), doc.y, { width: CW - (p.indent ? 24 : 0), paragraphGap: 12, lineGap: 3 });
      });

      doc.font("Times-Roman").fontSize(12).fillColor(INK).text(model.closing || "Sincerely,", L, doc.y);
      doc.moveDown(3.2);
      const sg = model.signer || {};
      doc.font("Times-Roman").fontSize(12).fillColor(INK).text(sg.name, L);
      if (sg.title) doc.font("Times-Italic").fontSize(12).text(sg.title, L);
      doc.font("Times-Roman").fontSize(12).text("T: " + sg.phone, L);
      doc.text("E: " + sg.email, L);

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
// Best-match the staff folder by name tokens (handles leading "000." numbering + order).
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
function isoToday() { return new Date().toISOString().slice(0, 10); }
function shortToday() { var d = new Date(); return String(d.getFullYear()).slice(-2) + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0"); }

/* ---------- helpers ---------- */
function payPhrase(amount, basis) {
  const num = Number(String(amount == null ? "" : amount).replace(/[^0-9.]/g, ""));
  if (!num) return "";
  const money = "$" + num.toLocaleString("en-CA", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return (String(basis).toLowerCase() === "hour")
    ? "hourly rate of pay is " + money + " per hour before taxes and benefits"
    : "gross annual salary is " + money + " before taxes and benefits";
}
function parseDate(s) {
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(s);
}
// "April 2nd 2026" — matches the reference letter's date style.
function ordinalDate(d) {
  try {
    var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    var day = d.getDate(), t = day % 10, h = day % 100;
    var suf = (t === 1 && h !== 11) ? "st" : (t === 2 && h !== 12) ? "nd" : (t === 3 && h !== 13) ? "rd" : "th";
    return months[d.getMonth()] + " " + day + suf + " " + d.getFullYear();
  } catch (e) { return new Date().toDateString(); }
}
function isoDate(s) {
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return s;
  var d = new Date(s); return isNaN(d) ? String(s) : d.toISOString().slice(0, 10);
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
