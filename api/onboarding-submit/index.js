/*
 * POST /api/onboarding-submit  — public, anonymous endpoint for the Project
 * Onboarding Form on the marketing site. It validates the submission and emails
 * it to the sales team via Microsoft Graph (app-only, client-credentials).
 *
 * Body (JSON), built by project-onboarding.html:
 *   { website: "<honeypot, must be empty>", fields: [ { label, value } ] }
 *
 * App settings (Static Web App / Function App → Configuration):
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET   app registration with Graph Mail.Send (application)
 *   GRAPH_TENANT                       directory (tenant) id  (defaults to King EPCM's)
 *   ONBOARD_MAIL_FROM                  mailbox the email is sent from, e.g. info@kingepcm.com
 *   ONBOARD_MAIL_TO                    recipients, comma-separated (default sales@kingepcm.com)
 *
 * No third party: data goes straight from the form to your mailbox. If the mail
 * settings aren't configured yet, the function returns 200 but flags notSent so
 * the page can fall back to "email us" without looking broken.
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const MAIL_FROM = process.env.ONBOARD_MAIL_FROM || "";
const MAIL_TO = process.env.ONBOARD_MAIL_TO || "sales@kingepcm.com";

module.exports = async function (context, req) {
  const body = req.body || {};

  // Anti-spam: silently accept honeypot hits without emailing (don't tip off bots).
  if (body.website) { context.res = json(200, { ok: true }); return; }

  const fields = Array.isArray(body.fields) ? body.fields : [];
  const filled = fields.filter(function (f) { return f && f.value && String(f.value).trim(); });
  if (!filled.length) { context.res = json(400, { ok: false, error: "Empty submission" }); return; }

  // Basic size guard against abuse.
  const tooBig = fields.some(function (f) { return String(f.value || "").length > 5000; });
  if (tooBig) { context.res = json(413, { ok: false, error: "Submission too large" }); return; }

  if (!CLIENT_ID || !CLIENT_SECRET || !MAIL_FROM) {
    context.log.warn("onboarding-submit: mail not configured");
    context.res = json(200, { ok: false, notSent: true, error: "Mail not configured" });
    return;
  }

  try {
    const token = await appToken();
    const ownerRow = filled.find(function (f) { return /owner/i.test(f.label); });
    const owner = ownerRow ? String(ownerRow.value).trim() : "";
    const subject = "New project onboarding" + (owner ? " — " + owner : "") + " (kingepcm.com)";
    const html = buildHtml(fields);
    await sendMail(token, MAIL_FROM, MAIL_TO, subject, html);
    context.res = json(200, { ok: true });
  } catch (e) {
    context.log.error(e);
    context.res = json(502, { ok: false, error: "Send failed" });
  }
};

function buildHtml(fields) {
  const rows = fields.map(function (f) {
    const v = String(f.value == null ? "" : f.value).trim();
    return "<tr>" +
      "<td style=\"padding:8px 12px;border:1px solid #e2e2e2;background:#f6f7f9;font-weight:600;color:#14294D;vertical-align:top;width:34%\">" + esc(f.label) + "</td>" +
      "<td style=\"padding:8px 12px;border:1px solid #e2e2e2;color:#222;white-space:pre-wrap\">" + (v ? esc(v) : "<span style=\"color:#999\">—</span>") + "</td>" +
      "</tr>";
  }).join("");
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222\">" +
    "<h2 style=\"color:#14294D;margin:0 0 4px\">Project Onboarding Form</h2>" +
    "<p style=\"color:#555;margin:0 0 14px\">Submitted from kingepcm.com on " + esc(new Date().toLocaleString("en-CA")) + ".</p>" +
    "<table style=\"border-collapse:collapse;width:100%;max-width:680px\">" + rows + "</table>" +
    "</div>";
}

/* ---------- Microsoft Graph (app-only) ---------- */
async function appToken() {
  const b = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: b });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function sendMail(token, from, to, subject, html) {
  const message = {
    subject: subject,
    body: { contentType: "HTML", content: html },
    toRecipients: to.split(",").map(function (e) { return { emailAddress: { address: e.trim() } }; })
  };
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: message, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status + " " + (await r.text().catch(function () { return ""; })));
}

/* ---------- helpers ---------- */
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
