/*
 * POST /api/onboarding-submit  — public, anonymous endpoint for the Project
 * Onboarding Form on the marketing site. On each submission it:
 *   1) builds a branded PDF of the answers
 *   2) emails it to the sales team (Microsoft Graph, app-only)
 *   3) posts it into Teamwork: finds the project whose name matches the project
 *      address (falls back to an intake project), creates a notebook with the
 *      answers, and attaches the PDF to that notebook.
 * Every step is best-effort and independently configured, so the form keeps
 * working even if mail or Teamwork isn't set up yet.
 *
 * Body (JSON), built by project-onboarding.html:
 *   { website: "<honeypot, must be empty>", fields: [ { label, value } ] }
 *
 * App settings (Static Web App / Function App -> Configuration):
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET   app registration with Graph Mail.Send (application)
 *   GRAPH_TENANT                       directory (tenant) id  (defaults to King EPCM's)
 *   ONBOARD_MAIL_FROM                  mailbox the email is sent from, e.g. info@kingepcm.com
 *   ONBOARD_MAIL_TO                    recipients, comma-separated (default sales@kingepcm.com)
 *   TEAMWORK_DOMAIN                    Teamwork subdomain, e.g. "kingepcm"
 *   TEAMWORK_API_KEY                   Teamwork API token
 *   TEAMWORK_INTAKE_PROJECT_ID         fallback project id when no address match is found
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const MAIL_FROM = process.env.ONBOARD_MAIL_FROM || "";
const MAIL_TO = process.env.ONBOARD_MAIL_TO || "sales@kingepcm.com";
const TW_DOMAIN = process.env.TEAMWORK_DOMAIN || "";
const TW_KEY = process.env.TEAMWORK_API_KEY || "";
const TW_INTAKE = process.env.TEAMWORK_INTAKE_PROJECT_ID || "";

module.exports = async function (context, req) {
 try {
  const body = req.body || {};

  // Anti-spam: silently accept honeypot hits without doing anything.
  if (body.website) { context.res = json(200, { ok: true }); return; }

  const fields = Array.isArray(body.fields) ? body.fields : [];
  const filled = fields.filter(function (f) { return f && f.value && String(f.value).trim(); });
  if (!filled.length) { context.res = json(400, { ok: false, error: "Empty submission" }); return; }
  if (fields.some(function (f) { return String(f.value || "").length > 5000; })) {
    context.res = json(413, { ok: false, error: "Submission too large" }); return;
  }

  const owner = valueOf(fields, /owner/i);
  const address = valueOf(fields, /project address/i);
  const fileBase = sanitize("Project Onboarding - " + (owner || address || today()));
  const notebookName = sanitize("Project Contact - " + (owner || address || today())); // Teamwork notebook title
  const status = { ok: false, emailed: false, postedToTeamwork: false, matchedProject: null, notes: [] };

  // 1) PDF (pull the live logo so it always matches the site; falls back to a wordmark)
  let pdf = null;
  try {
    let logo = null;
    try { logo = await fetchLogo(req); } catch (e) { context.log.error(e); }
    pdf = await buildPdf(fields, owner, address, logo);
  } catch (e) { context.log.error(e); status.notes.push("pdf failed"); }

  // 2) Email sales
  if (CLIENT_ID && CLIENT_SECRET && MAIL_FROM) {
    try {
      const token = await graphToken();
      const subject = "New project onboarding" + (owner ? " — " + owner : "") + " (kingepcm.com)";
      await sendMail(token, MAIL_FROM, MAIL_TO, subject, mailHtml(fields), pdf, fileBase + ".pdf");
      status.emailed = true;
    } catch (e) { context.log.error(e); status.notes.push("email failed"); }
  } else status.notes.push("mail not configured");

  // 3) Teamwork: match a project by address, create a notebook, attach the PDF.
  if (TW_DOMAIN && TW_KEY) {
    try {
      const projects = await twProjects();
      let proj = bestProject(projects, address, owner);
      if (!proj && TW_INTAKE) proj = { id: TW_INTAKE, name: "(intake)" };
      if (proj) {
        // Notify everyone on the project about the new notebook.
        let notify = "";
        try { notify = (await twProjectPeople(proj.id)).join(","); }
        catch (e) { context.log.error(e); status.notes.push("people lookup failed"); }
        const notebookId = await twCreateNotebook(proj.id, notebookName, notebookHtml(fields, owner, address), notify);
        if (notebookId && pdf) {
          const ref = await twUploadPending(pdf, fileBase + ".pdf").catch(function (e) { context.log.error(e); return null; });
          if (ref) await twAttachToNotebook(notebookId, ref).catch(function (e) { context.log.error(e); status.notes.push("pdf attach failed"); });
          else status.notes.push("pdf upload failed");
        }
        status.postedToTeamwork = !!notebookId;
        status.matchedProject = proj.name;
      } else status.notes.push("no project match and no intake project set");
    } catch (e) { context.log.error(e); status.notes.push("teamwork failed"); }
  } else status.notes.push("teamwork not configured");

  // Success for the client as long as we captured it somewhere (email or Teamwork).
  status.ok = status.emailed || status.postedToTeamwork;
  context.res = json(status.ok ? 200 : 502, status);
 } catch (e) {
  context.log.error(e);
  context.res = json(500, { ok: false, error: String((e && e.message) || e) });
 }
};

/* ---------- PDF ---------- */
// Pull the site's white logo so the PDF header matches the brand. Uses the same
// host the form was served from (works on preview + production); ONBOARD_LOGO_URL overrides.
async function fetchLogo(req) {
  const h = req.headers || {};
  const host = h["x-forwarded-host"] || h["host"] || "kingepcm.com";
  const proto = h["x-forwarded-proto"] || "https";
  const url = process.env.ONBOARD_LOGO_URL || (proto + "://" + host + "/assets/logo-white.png");
  const r = await fetch(url);
  if (!r.ok) throw new Error("logo " + r.status);
  return Buffer.from(await r.arrayBuffer());
}

function buildPdf(fields, owner, address, logo) {
  return new Promise(function (resolve, reject) {
    try {
      const PDFDocument = require("pdfkit"); // lazy: a missing lib disables the PDF, not the endpoint
      const doc = new PDFDocument({ size: "A4", margins: { top: 50, left: 50, right: 50, bottom: 56 } });
      const chunks = [];
      doc.on("data", function (d) { chunks.push(d); });
      doc.on("end", function () { resolve(Buffer.concat(chunks)); });

      const NAVY = "#14294D", GOLD = "#E5A823", INK = "#1F2A37", MUTE = "#5B6675", LINE = "#E3E8EF";
      const W = doc.page.width, L = 50, R = W - 50, CW = R - L;

      // Header band
      doc.rect(0, 0, W, 96).fill(NAVY);
      let drew = false;
      if (logo) { try { doc.image(logo, L, 29, { height: 40 }); drew = true; } catch (e) {} }
      if (!drew) doc.fillColor("#FFFFFF").font("Helvetica-Bold").fontSize(22).text("KING EPCM", L, 35);
      doc.fillColor(GOLD).font("Helvetica").fontSize(9.5).text("kingepcm.com", L, 44, { width: CW, align: "right" });

      // Title block
      doc.fillColor(NAVY).font("Helvetica-Bold").fontSize(20).text("Project Onboarding Form", L, 122);
      doc.fillColor(MUTE).font("Helvetica").fontSize(9.5).text("Submitted " + new Date().toLocaleString("en-CA"), L, doc.y + 2);
      let ry = doc.y + 8;
      doc.moveTo(L, ry).lineTo(R, ry).lineWidth(2).strokeColor(GOLD).stroke();
      doc.y = ry + 18;

      const GROUPS = [
        { title: "Owner & Project", labels: ["Legal Owner Name", "Project Address"] },
        { title: "Billing", labels: ["Billing Contact Name", "Billing Contact Phone", "Billing Contact Email", "Billing Mailing Address"] },
        { title: "Project Contacts", labels: ["Primary Site Contact", "Additional contacts needed", "Additional Contacts"] }
      ];
      const used = {};
      function ensure(h) { if (doc.y + h > doc.page.height - doc.page.margins.bottom) doc.addPage(); }
      function section(title, rows) {
        ensure(64);
        doc.fillColor(NAVY).font("Helvetica-Bold").fontSize(12.5).text(title, L, doc.y);
        const yy = doc.y + 3;
        doc.moveTo(L, yy).lineTo(R, yy).lineWidth(0.8).strokeColor(GOLD).stroke();
        doc.y = yy + 10;
        rows.forEach(function (r) { row(r.label, r.value); });
        doc.y += 8;
      }
      function row(label, value) {
        const val = String(value == null || value === "" ? "—" : value);
        ensure(46);
        doc.fillColor(MUTE).font("Helvetica-Bold").fontSize(8).text(String(label).toUpperCase(), L, doc.y, { width: CW, characterSpacing: 0.4 });
        doc.fillColor(INK).font("Helvetica").fontSize(10.5).text(val, L, doc.y + 2, { width: CW });
        const yy = doc.y + 8;
        doc.moveTo(L, yy).lineTo(R, yy).lineWidth(0.5).strokeColor(LINE).stroke();
        doc.y = yy + 10;
      }

      GROUPS.forEach(function (g) {
        const rows = [];
        g.labels.forEach(function (l) { const f = (fields || []).find(function (x) { return x.label === l; }); if (f) { rows.push(f); used[l] = 1; } });
        if (rows.length) section(g.title, rows);
      });
      const leftover = (fields || []).filter(function (f) { return !used[f.label]; });
      if (leftover.length) section("Additional Information", leftover);

      doc.end();
    } catch (e) { reject(e); }
  });
}

/* ---------- Microsoft Graph (app-only) ---------- */
async function graphToken() {
  const b = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: b });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function sendMail(token, from, to, subject, html, pdf, filename) {
  const message = {
    subject: subject, body: { contentType: "HTML", content: html },
    toRecipients: to.split(",").map(function (e) { return { emailAddress: { address: e.trim() } }; })
  };
  if (pdf) message.attachments = [{ "@odata.type": "#microsoft.graph.fileAttachment", name: filename, contentType: "application/pdf", contentBytes: pdf.toString("base64") }];
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: message, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status + " " + (await r.text().catch(function () { return ""; })));
}

/* ---------- Teamwork (Projects API v1) ---------- */
function twBase() {
  // Accept "kingepcm", "kingepcm.teamwork.com", or a full URL — normalize to the host.
  var d = String(TW_DOMAIN || "").trim().replace(/^https?:\/\//i, "").replace(/\/.*$/, "").replace(/\.teamwork\.com$/i, "");
  return "https://" + d + ".teamwork.com";
}
function twAuth() { return "Basic " + Buffer.from(TW_KEY + ":x").toString("base64"); }
async function twProjects() {
  const r = await fetch(twBase() + "/projects.json?pageSize=500&status=ACTIVE", { headers: { Authorization: twAuth() } });
  if (!r.ok) throw new Error("tw projects " + r.status);
  const j = await r.json();
  return (j.projects || []).map(function (p) { return { id: p.id, name: p.name || "" }; });
}
// Generic words that don't help tell two addresses apart (street types + directions).
const GENERIC = {
  st: "street", str: "street", street: "street", rd: "road", road: "road", ave: "avenue", av: "avenue", avenue: "avenue",
  dr: "drive", drive: "drive", blvd: "boulevard", boulevard: "boulevard", cres: "crescent", crescent: "crescent",
  ct: "court", crt: "court", court: "court", ln: "lane", lane: "lane", hwy: "highway", highway: "highway",
  trl: "trail", trail: "trail", ter: "terrace", terr: "terrace", terrace: "terrace", pkwy: "parkway", parkway: "parkway",
  sdrd: "sideroad", sideroad: "sideroad", sderd: "sideroad", conc: "concession", concession: "concession",
  pl: "place", place: "place", cir: "circle", circle: "circle", grv: "grove", grove: "grove", way: "way",
  line: "line", sideline: "sideline", n: "north", s: "south", e: "east", w: "west"
};
function isGeneric(t) { return Object.prototype.hasOwnProperty.call(GENERIC, t) || ["north", "south", "east", "west"].indexOf(t) > -1; }
// Tokenize and canonicalize abbreviations, preserving order.
function normToks(s) { return toks(s).map(function (t) { return GENERIC[t] || t; }); }
// The distinctive street-name word(s): the name that follows EACH street number
// (so "16 Beaver St/11 Credit St" yields both "beaver" and "credit"), skipping a
// leading street-type and stopping before the city. Named projects (no number)
// contribute their whole name.
function coreTokens(s) {
  const t = normToks(s), out = [];
  if (t.some(isNum)) {
    for (let i = 0; i < t.length; i++) {
      if (!isNum(t[i])) continue;
      let j = i + 1;
      while (j < t.length && isGeneric(t[j])) j++;                 // skip "St" in "St Regis", etc.
      while (j < t.length && !isNum(t[j]) && !isGeneric(t[j])) { out.push(t[j]); j++; } // capture the street name
    }
  } else {
    for (let i = 0; i < t.length; i++) if (!isGeneric(t[i])) out.push(t[i]);
  }
  return out;
}
// Fuzzy match the entered address to a project name. Keys on the street number +
// the distinctive street-name word, so abbreviations and a missing city still match,
// but a different street (or building) on the same road does not.
function bestProject(projects, address, owner) {
  const aAll = normToks(address);
  if (!aAll.length) return null;
  const aNums = aAll.filter(isNum);
  const aCore = coreTokens(address);
  let best = null, bestScore = 0;
  (projects || []).forEach(function (p) {
    const pAll = normToks(p.name);
    const pNums = pAll.filter(isNum);
    const pCore = coreTokens(p.name);
    const numMatch = aNums.some(function (n) { return pNums.indexOf(n) > -1; });
    const coreShared = aCore.filter(function (t) { return pCore.indexOf(t) > -1; }).length;
    const allShared = aAll.filter(function (t) { return pAll.indexOf(t) > -1; }).length;
    // Addresses (have a number): need the number AND a street-name word to match.
    // Named projects (no number): need a strong name overlap.
    const ok = aNums.length ? (numMatch && coreShared >= 1) : (coreShared >= 1 && allShared >= 2);
    if (ok) {
      const score = coreShared * 3 + allShared + (numMatch ? 2 : 0);
      if (score > bestScore) { bestScore = score; best = p; }
    }
  });
  return best;
}
async function twCreateNotebook(projectId, name, contentHtml, notify) {
  const r = await fetch(twBase() + "/projects/" + projectId + "/notebooks.json", {
    method: "POST", headers: { Authorization: twAuth(), "Content-Type": "application/json" },
    body: JSON.stringify({ notebook: { name: name.slice(0, 200), type: "HTML", content: contentHtml, "notify": notify || "", description: "Submitted via the website onboarding form" } })
  });
  if (!r.ok) throw new Error("tw notebook " + r.status + " " + (await r.text().catch(function () { return ""; })));
  const j = await r.json().catch(function () { return {}; });
  return j.notebookId || (j.notebook && j.notebook.id) || null;
}
// Everyone on the project (their user ids) — used to notify them of the new notebook.
async function twProjectPeople(projectId) {
  const r = await fetch(twBase() + "/projects/" + projectId + "/people.json", { headers: { Authorization: twAuth() } });
  if (!r.ok) throw new Error("tw people " + r.status);
  const j = await r.json().catch(function () { return {}; });
  return (j.people || []).map(function (x) { return x.id; }).filter(Boolean);
}
async function twUploadPending(pdf, filename) {
  const fd = new FormData();
  fd.append("file", new Blob([pdf], { type: "application/pdf" }), filename);
  const r = await fetch(twBase() + "/pendingfiles.json", { method: "POST", headers: { Authorization: twAuth() }, body: fd });
  if (!r.ok) throw new Error("tw pendingfiles " + r.status);
  const j = await r.json().catch(function () { return {}; });
  return (j.pendingFile && j.pendingFile.ref) || j.ref || j.REF || null;
}
async function twAttachToNotebook(notebookId, ref) {
  const r = await fetch(twBase() + "/notebooks/" + notebookId + "/comments.json", {
    method: "POST", headers: { Authorization: twAuth(), "Content-Type": "application/json" },
    body: JSON.stringify({ comment: { body: "Completed onboarding form (PDF attached).", "notify": "", "pendingFileAttachments": ref } })
  });
  if (!r.ok) throw new Error("tw notebook comment " + r.status);
}

/* ---------- content ---------- */
function rowsHtml(fields) {
  return (fields || []).map(function (f) {
    const v = String(f.value == null ? "" : f.value).trim();
    return "<tr>" +
      "<td style=\"padding:8px 12px;border:1px solid #e2e2e2;background:#f6f7f9;font-weight:600;color:#14294D;vertical-align:top;width:34%\">" + esc(f.label) + "</td>" +
      "<td style=\"padding:8px 12px;border:1px solid #e2e2e2;color:#222;white-space:pre-wrap\">" + (v ? esc(v) : "—") + "</td></tr>";
  }).join("");
}
function mailHtml(fields) {
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222\">" +
    "<h2 style=\"color:#14294D;margin:0 0 4px\">Project Onboarding Form</h2>" +
    "<p style=\"color:#555;margin:0 0 14px\">Submitted from kingepcm.com on " + esc(new Date().toLocaleString("en-CA")) + ".</p>" +
    "<table style=\"border-collapse:collapse;width:100%;max-width:680px\">" + rowsHtml(fields) + "</table></div>";
}
function notebookHtml(fields, owner, address) {
  return "<p>Submitted via the website project onboarding form on " + esc(new Date().toLocaleString("en-CA")) + ". The completed PDF is attached to this notebook.</p>" +
    "<table style=\"border-collapse:collapse;width:100%\">" + rowsHtml(fields) + "</table>";
}

/* ---------- helpers ---------- */
function valueOf(fields, re) { const f = (fields || []).find(function (x) { return re.test(x.label || ""); }); return f ? String(f.value || "").trim() : ""; }
function toks(s) { return String(s || "").toLowerCase().normalize("NFKD").replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter(function (t) { return t.length > 1; }); }
function isNum(t) { return /^\d+$/.test(t); }
function sanitize(s) { return String(s || "form").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 120); }
function today() { return new Date().toLocaleDateString("en-CA"); }
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
