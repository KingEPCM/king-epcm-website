/*
 * POST /api/pay-details — public, anonymous endpoint for the "Pay by Credit Card" form
 * on the marketing site. Emails the payer's project/invoice details to Accounting
 * (Microsoft Graph, app-only) with the payer set as reply-to, then the browser
 * redirects the visitor to the secure Clover payment page. Honeypot anti-spam.
 *
 * Body (JSON), built by pay.html:
 *   { website (honeypot), first_name, last_name, company, phone, email, invoice }
 *
 * App settings (Static Web App -> Configuration):
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET   app registration with Graph Mail.Send (application)
 *   GRAPH_TENANT                       directory (tenant) id (defaults to King EPCM's)
 *   FORM_MAIL_FROM                     mailbox the email is sent from (falls back to ONBOARD_MAIL_FROM)
 *   PAY_MAIL_TO                        recipients, comma-separated (default accounting@kingepcm.com)
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const MAIL_FROM = process.env.FORM_MAIL_FROM || process.env.ONBOARD_MAIL_FROM || "";
const MAIL_TO = process.env.PAY_MAIL_TO || "accounting@kingepcm.com";

module.exports = async function (context, req) {
 try {
  const b = req.body || {};
  if (b.website) { context.res = json(200, { ok: true }); return; } // honeypot: silently accept
  if (!CLIENT_ID || !CLIENT_SECRET || !MAIL_FROM) { context.res = json(501, { ok: false, error: "Email not configured" }); return; }

  const first = String(b.first_name || "").trim();
  const last = String(b.last_name || "").trim();
  const phone = String(b.phone || "").trim();
  const email = String(b.email || "").trim();
  const invoice = String(b.invoice || "").trim();
  if (!first || !last || !phone || !email || !invoice) { context.res = json(400, { ok: false, error: "Please complete all required fields." }); return; }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { context.res = json(400, { ok: false, error: "Please enter a valid email address." }); return; }

  const token = await graphToken();
  const subject = "Credit card payment details — " + first + " " + last + " (kingepcm.com)";
  await sendMail(token, MAIL_FROM, MAIL_TO, email, subject, payHtml(b), null);
  context.res = json(200, { ok: true });
 } catch (e) {
  context.log.error(e);
  context.res = json(502, { ok: false, error: "Could not submit your details" });
 }
};

/* ---------- Microsoft Graph (app-only) ---------- */
async function graphToken() {
  const b = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: b });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function sendMail(token, from, to, replyTo, subject, html) {
  const message = {
    subject: subject, body: { contentType: "HTML", content: html },
    toRecipients: to.split(",").map(function (e) { return { emailAddress: { address: e.trim() } }; })
  };
  if (replyTo) message.replyTo = [{ emailAddress: { address: replyTo } }];
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: message, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status + " " + (await r.text().catch(function () { return ""; })));
}

/* ---------- content ---------- */
function payHtml(b) {
  const rows = [
    ["Name", (String(b.first_name || "").trim() + " " + String(b.last_name || "").trim()).trim()],
    ["Company", b.company], ["Phone", b.phone], ["Email", b.email],
    ["Invoice / Estimate / Project", b.invoice]
  ].filter(function (r) { return String(r[1] == null ? "" : r[1]).trim(); });
  const body = rows.map(function (r) {
    return "<tr><td style=\"padding:8px 12px;border:1px solid #e2e2e2;background:#f6f7f9;font-weight:600;color:#14294D;vertical-align:top;width:34%\">" + esc(r[0]) +
      "</td><td style=\"padding:8px 12px;border:1px solid #e2e2e2;color:#222;white-space:pre-wrap\">" + esc(r[1]) + "</td></tr>";
  }).join("");
  return "<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222\">" +
    "<h2 style=\"color:#14294D;margin:0 0 4px\">Credit card payment — project details</h2>" +
    "<p style=\"color:#555;margin:0 0 14px\">Submitted from kingepcm.com on " + esc(new Date().toLocaleString("en-CA")) + ". The payer was then redirected to Clover to enter their card. Reply to this email to reach them directly.</p>" +
    "<table style=\"border-collapse:collapse;width:100%;max-width:680px\">" + body + "</table></div>";
}
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
