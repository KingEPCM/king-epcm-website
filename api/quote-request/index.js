/*
 * POST /api/quote-request — public, anonymous endpoint for the "Request a Quote" form
 * on the marketing site. Emails the details to the sales team (Microsoft Graph, app-only),
 * with the visitor set as reply-to and any uploaded documents attached. Honeypot anti-spam.
 *
 * Body (JSON), built by contact.html:
 *   { website (honeypot, must be empty), name, company, email, phone, service, location,
 *     message, attachments:[{name,contentType,base64}] }
 *
 * App settings (Static Web App -> Configuration):
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET   app registration with Graph Mail.Send (application)
 *   GRAPH_TENANT                       directory (tenant) id (defaults to King EPCM's)
 *   FORM_MAIL_FROM                     mailbox the email is sent from (falls back to ONBOARD_MAIL_FROM)
 *   QUOTE_MAIL_TO                      recipients, comma-separated (default sales@kingepcm.com)
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const MAIL_FROM = process.env.FORM_MAIL_FROM || process.env.ONBOARD_MAIL_FROM || "";
const MAIL_TO = process.env.QUOTE_MAIL_TO || "sales@kingepcm.com";
const MAX_FILE = 4 * 1024 * 1024;   // 4 MB per file
const MAX_EMAIL = 3 * 1024 * 1024;  // ~3 MB total attached (Graph simple-send limit)

module.exports = async function (context, req) {
 try {
  const b = req.body || {};
  if (b.website) { context.res = json(200, { ok: true }); return; } // honeypot: silently accept
  if (!CLIENT_ID || !CLIENT_SECRET || !MAIL_FROM) { context.res = json(501, { ok: false, error: "Email not configured" }); return; }

  const name = String(b.name || "").trim();
  const email = String(b.email || "").trim();
  const message = String(b.message || "").trim();
  if (!name || !email || !message) { context.res = json(400, { ok: false, error: "Name, email, and project details are required." }); return; }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { context.res = json(400, { ok: false, error: "Please enter a valid email address." }); return; }

  const atts = (Array.isArray(b.attachments) ? b.attachments : []).filter(function (a) { return a && a.base64; });
  let totalBytes = 0;
  for (const a of atts) {
    const sz = Buffer.from(a.base64, "base64").length;
    if (sz > MAX_FILE) { context.res = json(413, { ok: false, error: "A file is larger than 4 MB. Please email large files to sales@KingEPCM.com." }); return; }
    totalBytes += sz;
  }
  if (totalBytes > MAX_EMAIL) { context.res = json(413, { ok: false, error: "Attachments are too large to send — please email them to sales@KingEPCM.com." }); return; }

  const token = await graphToken();
  const attachments = atts.map(function (a) {
    return { "@odata.type": "#microsoft.graph.fileAttachment", name: a.name || "attachment", contentType: a.contentType || "application/octet-stream", contentBytes: a.base64 };
  });
  const subject = "Quote request — " + name + (b.service ? " — " + String(b.service).trim() : "") + " (kingepcm.com)";
  await sendMail(token, MAIL_FROM, MAIL_TO, email, subject, quoteHtml(b), attachments);
  context.res = json(200, { ok: true });
 } catch (e) {
  context.log.error(e);
  context.res = json(502, { ok: false, error: "Could not send your request" });
 }
};

/* ---------- Microsoft Graph (app-only) ---------- */
async function graphToken() {
  const b = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: b });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function sendMail(token, from, to, replyTo, subject, html, attachments) {
  const message = {
    subject: subject, body: { contentType: "HTML", content: html },
    toRecipients: to.split(",").map(function (e) { return { emailAddress: { address: e.trim() } }; })
  };
  if (replyTo) message.replyTo = [{ emailAddress: { address: replyTo } }];
  if (attachments && attachments.length) message.attachments = attachments;
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: message, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status + " " + (await r.text().catch(function () { return ""; })));
}

/* ---------- content ---------- */
function quoteHtml(b) {
  const rows = [
    ["Name", b.name], ["Company", b.company], ["Email", b.email], ["Phone", b.phone],
    ["Service needed", b.service], ["Project location", b.location], ["Project details", b.message]
  ].filter(function (r) { return String(r[1] == null ? "" : r[1]).trim(); });
  const body = rows.map(function (r) {
    return "<tr><td style=\"padding:8px 12px;border:1px solid #e2e2e2;background:#f6f7f9;font-weight:600;color:#14294D;vertical-align:top;width:32%\">" + esc(r[0]) +
      "</td><td style=\"padding:8px 12px;border:1px solid #e2e2e2;color:#222;white-space:pre-wrap\">" + esc(r[1]) + "</td></tr>";
  }).join("");
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222\">" +
    "<h2 style=\"color:#14294D;margin:0 0 4px\">New quote request</h2>" +
    "<p style=\"color:#555;margin:0 0 14px\">Submitted from kingepcm.com on " + esc(new Date().toLocaleString("en-CA")) + ". Reply to this email to reach the client directly.</p>" +
    "<table style=\"border-collapse:collapse;width:100%;max-width:680px\">" + body + "</table></div>";
}
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
