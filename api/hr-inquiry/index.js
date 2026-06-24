/*
 * POST /api/hr-inquiry — sends a general HR enquiry to HR automatically (no mailto).
 * Sent AS the person submitting it (Graph Mail.Send "send as any user"), so it reaches
 * HR from the employee and replies go straight back to them. Any attachment is included.
 *
 * Body (JSON): { topic?, message, attachments?:[{name,contentType,base64}] }
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional),
 *   HR_NOTIFY_TO (default "hr@kingepcm.com").
 * Graph APPLICATION permission: Mail.Send (admin consent).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const TO = process.env.HR_NOTIFY_TO || "hr@kingepcm.com";
const MAX_FILE = 4 * 1024 * 1024;   // 4 MB per attachment
const MAX_EMAIL = 3 * 1024 * 1024;  // ~3 MB total attached to the email

module.exports = async function (context, req) {
 try {
  const p = principal(req);
  if (!p.email) { context.res = json(401, { ok: false, error: "Not signed in" }); return; }
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(501, { ok: false, error: "Email not configured" }); return; }

  const b = req.body || {};
  const topic = String(b.topic || "").trim() || "General question";
  const message = String(b.message || "").trim();
  if (!message) { context.res = json(400, { ok: false, error: "A message is required" }); return; }

  // Optional attachments — validate sizes (emailed only).
  const atts = (Array.isArray(b.attachments) ? b.attachments : []).filter(function (a) { return a && a.base64; });
  let totalBytes = 0;
  for (const a of atts) {
    const sz = Buffer.from(a.base64, "base64").length;
    if (sz > MAX_FILE) { context.res = json(413, { ok: false, error: "An attachment is larger than 4 MB. Please reduce its size." }); return; }
    totalBytes += sz;
  }
  if (totalBytes > MAX_EMAIL) { context.res = json(413, { ok: false, error: "Attachments are too large to email — keep the total under 3 MB." }); return; }

  const token = await appToken();
  const attachments = atts.map(function (a) {
    return { "@odata.type": "#microsoft.graph.fileAttachment", name: a.name || "attachment", contentType: a.contentType || "application/octet-stream", contentBytes: a.base64 };
  });

  const subject = "HR enquiry: " + topic + " — " + (p.name || p.email);
  await sendMail(token, p.email, TO, p.email, subject, inquiryHtml(topic, message, p, attachments.length), attachments);

  context.res = json(200, { ok: true });
 } catch (e) {
  context.log.error(e);
  context.res = json(502, { ok: false, error: "Could not send your enquiry" });
 }
};

/* ---------- Email (Graph Mail.Send) ---------- */
async function sendMail(token, from, to, replyTo, subject, html, attachments) {
  const msg = { subject: subject, body: { contentType: "HTML", content: html }, toRecipients: String(to).split(",").map(function (e) { return { emailAddress: { address: e.trim() } }; }) };
  if (replyTo) msg.replyTo = [{ emailAddress: { address: replyTo } }];
  if (attachments && attachments.length) msg.attachments = attachments;
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: msg, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status + " " + (await r.text().catch(function () { return ""; })));
}
function inquiryHtml(topic, message, p, attachedCount) {
  const body = "<p><b>Topic:</b> " + esc(topic) + "</p>" +
    "<p><b>Message:</b><br>" + esc(message).replace(/\r?\n/g, "<br>") + "</p>" +
    "<p><b>Attachments:</b> " + (attachedCount ? attachedCount + " attached to this email" : "none") + "</p>" +
    "<p style='color:#555;font-size:12px'>Submitted via the King EPCM staff intranet.</p>";
  return "<div style='font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222'>" +
    "<h2 style='color:#14294D;margin:0 0 6px'>General HR enquiry</h2>" +
    "<p style='color:#555;margin:0 0 14px'>From <b>" + esc(p.name || "") + "</b> &lt;" + esc(p.email) + "&gt; on " + esc(new Date().toLocaleString("en-CA")) + ".</p>" + body + "</div>";
}
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

/* ---------- Graph ---------- */
async function appToken() {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}

/* ---------- helpers ---------- */
function principal(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return {};
    const o = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    const claims = o.claims || [];
    const cv = function (t) { const c = claims.find(function (x) { return t.indexOf((x.typ || "").toLowerCase()) > -1 && x.val; }); return c ? c.val : ""; };
    const email = (o.userDetails && o.userDetails.indexOf("@") > -1) ? o.userDetails : cv(["email", "preferred_username", "upn"]);
    const name = cv(["name"]);
    return { email: email, name: (name && name.indexOf("@") === -1) ? name : "" };
  } catch (e) { return {}; }
}
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
