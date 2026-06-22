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
  const status = { ok: false, emailed: false, postedToTeamwork: false, matchedProject: null, notes: [] };

  // 1) PDF
  let pdf = null;
  try { pdf = await buildPdf(fields, owner, address); }
  catch (e) { context.log.error(e); status.notes.push("pdf failed"); }

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
        const notebookId = await twCreateNotebook(proj.id, fileBase, notebookHtml(fields, owner, address));
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
function buildPdf(fields, owner, address) {
  return new Promise(function (resolve, reject) {
    try {
      const PDFDocument = require("pdfkit"); // lazy: a missing lib disables the PDF, not the endpoint
      const doc = new PDFDocument({ size: "A4", margin: 50 });
      const chunks = [];
      doc.on("data", function (d) { chunks.push(d); });
      doc.on("end", function () { resolve(Buffer.concat(chunks)); });
      doc.fillColor("#14294D").font("Helvetica-Bold").fontSize(18).text("King EPCM");
      doc.fillColor("#C8901A").fontSize(13).text("Project Onboarding Form");
      doc.moveDown(0.3).fillColor("#555").font("Helvetica").fontSize(10)
        .text("Submitted from kingepcm.com on " + new Date().toLocaleString("en-CA"));
      doc.moveDown(0.6);
      (fields || []).forEach(function (f) {
        if (f.label === "Additional contacts needed") return; // captured implicitly
        doc.fillColor("#14294D").font("Helvetica-Bold").fontSize(10).text((f.label || "") + ":");
        doc.fillColor("#222").font("Helvetica").fontSize(10).text(String(f.value || "—"), { indent: 14 });
        doc.moveDown(0.25);
      });
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
// Find the project whose name matches the project address (street number + street word).
function bestProject(projects, address, owner) {
  const at = toks(address);
  const nums = at.filter(isNum);
  if (!at.length) return null;
  let best = null, bestScore = 0;
  (projects || []).forEach(function (p) {
    const pt = toks(p.name);
    const shared = at.filter(function (t) { return pt.indexOf(t) > -1; });
    const numMatch = nums.some(function (n) { return pt.indexOf(n) > -1; });
    const alphaShared = shared.filter(function (t) { return !isNum(t); }).length;
    // Confident only when the street number matches AND a street word matches.
    if (numMatch && alphaShared >= 1) {
      const score = shared.length + 2;
      if (score > bestScore) { bestScore = score; best = p; }
    }
  });
  return best;
}
async function twCreateNotebook(projectId, name, contentHtml) {
  const r = await fetch(twBase() + "/projects/" + projectId + "/notebooks.json", {
    method: "POST", headers: { Authorization: twAuth(), "Content-Type": "application/json" },
    body: JSON.stringify({ notebook: { name: name.slice(0, 200), type: "HTML", content: contentHtml, "notify": "", description: "Submitted via the website onboarding form" } })
  });
  if (!r.ok) throw new Error("tw notebook " + r.status + " " + (await r.text().catch(function () { return ""; })));
  const j = await r.json().catch(function () { return {}; });
  return j.notebookId || (j.notebook && j.notebook.id) || null;
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
