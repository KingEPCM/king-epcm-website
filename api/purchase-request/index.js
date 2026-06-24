/*
 * POST /api/purchase-request — submits a purchase request and EMAILS it automatically
 * (no mailto). Sent AS the person submitting it (Graph Mail.Send "send as any user"),
 * so it reaches the approvers from the requester and replies go straight back to them.
 * Recipients: Tony Wang + Accounting. Any attachments (quotes/screenshots) are attached
 * to the email. Every request is logged to Table Storage so the requester can see their
 * own history and avoid duplicates.
 *
 * Body (JSON): { project, vendor, items:[{description,price}], total?, comments?,
 *                attachments?:[{name,contentType,base64}] }
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional),
 *   PURCHASE_NOTIFY_TO (default "twang@kingepcm.com,accounting@kingepcm.com"),
 *   NEWS_STORAGE_CONNECTION (history log).
 * Graph APPLICATION permission: Mail.Send (admin consent).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const TO = process.env.PURCHASE_NOTIFY_TO || "twang@kingepcm.com,accounting@kingepcm.com";
const CONN = process.env.NEWS_STORAGE_CONNECTION;
const MAX_FILE = 4 * 1024 * 1024;   // 4 MB per attachment
const MAX_EMAIL = 3 * 1024 * 1024;  // ~3 MB total attached to the email
const { TableClient } = require("@azure/data-tables");

module.exports = async function (context, req) {
 try {
  const p = principal(req);
  if (!p.email) { context.res = json(401, { ok: false, error: "Not signed in" }); return; }
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(501, { ok: false, error: "Email not configured" }); return; }

  const b = req.body || {};
  const project = String(b.project || "").trim();
  const vendor = String(b.vendor || "").trim();
  const comments = String(b.comments || "").trim();
  const items = (Array.isArray(b.items) ? b.items : []).filter(function (it) { return it && (it.description || it.quantity || it.unitPrice); });
  if (!project) { context.res = json(400, { ok: false, error: "Project / purpose is required" }); return; }
  if (!vendor) { context.res = json(400, { ok: false, error: "Where you're buying from is required" }); return; }
  if (!items.length) { context.res = json(400, { ok: false, error: "Add at least one item" }); return; }
  for (const it of items) {
    if (!it.description || !(num(it.quantity) > 0) || !(num(it.unitPrice) > 0)) { context.res = json(400, { ok: false, error: "Every item needs a description, quantity, and unit price" }); return; }
  }

  // Optional attachments — validate sizes (emailed; not archived).
  const atts = (Array.isArray(b.attachments) ? b.attachments : []).filter(function (a) { return a && a.base64; });
  let totalBytes = 0;
  for (const a of atts) {
    const sz = Buffer.from(a.base64, "base64").length;
    if (sz > MAX_FILE) { context.res = json(413, { ok: false, error: "An attachment is larger than 4 MB. Please reduce its size." }); return; }
    totalBytes += sz;
  }
  if (totalBytes > MAX_EMAIL) { context.res = json(413, { ok: false, error: "Attachments are too large to email — keep the total under 3 MB." }); return; }

  const token = await appToken();
  const fromAddr = p.email; // send as the requester

  const attachments = atts.map(function (a) {
    return { "@odata.type": "#microsoft.graph.fileAttachment", name: a.name || "attachment", contentType: a.contentType || "application/octet-stream", contentBytes: a.base64 };
  });

  const total = "$" + money(items.reduce(function (s, it) { return s + lineTotal(it); }, 0));

  const subject = "Purchase request — " + project + " — " + (p.name || p.email);
  await sendMail(token, fromAddr, TO, p.email, subject, requestHtml(project, vendor, items, total, comments, p, attachments.length), attachments);

  const recSum = items.map(function (it, i) { return (i + 1) + ". " + it.description + " — " + num(it.quantity) + " × $" + money(it.unitPrice) + " = $" + money(lineTotal(it)); }).join("; ");
  await recordRequest(context, p, project, vendor, total, recSum, comments, items.length, attachments.length);

  context.res = json(200, { ok: true });
 } catch (e) {
  context.log.error(e);
  context.res = json(502, { ok: false, error: "Could not submit the request" });
 }
};

// Log a lightweight record so the requester can see their own history (avoid duplicates).
async function recordRequest(context, p, project, vendor, total, summary, comments, count, attachmentCount) {
  if (!CONN) return;
  try {
    const t = TableClient.fromConnectionString(CONN, "purchaserequests");
    try { await t.createTable(); } catch (e) {}
    await t.createEntity({
      partitionKey: keyOf(p.email), rowKey: String(Date.now()) + String(Math.floor(Math.random() * 1000)),
      project: project, vendor: vendor, total: String(total || ""), summary: String(summary || "").slice(0, 500),
      comments: String(comments || "").slice(0, 500), itemCount: count || 0, attachmentCount: attachmentCount || 0,
      byName: p.name || "", byEmail: p.email, date: isoToday(), createdAt: new Date().toISOString()
    });
  } catch (e) { context.log.error(e); }
}
function money(a) { var n = Number(String(a == null ? "" : a).replace(/[^0-9.]/g, "")); return (isNaN(n) ? 0 : n).toFixed(2); }
function num(v) { var n = parseFloat(String(v == null ? "" : v).replace(/[^0-9.]/g, "")); return isNaN(n) ? 0 : n; }
function lineTotal(it) { return num(it.quantity) * num(it.unitPrice); }
function keyOf(email) { return String(email || "").toLowerCase().replace(/[\\/#?\t\n\r]/g, "_"); }

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
function wrapHtml(title, p, inner) {
  return "<div style='font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222'>" +
    "<h2 style='color:#14294D;margin:0 0 6px'>" + esc(title) + "</h2>" +
    "<p style='color:#555;margin:0 0 14px'>Submitted by <b>" + esc(p.name || "") + "</b> &lt;" + esc(p.email) + "&gt; on " + esc(new Date().toLocaleString("en-CA")) + ".</p>" + inner + "</div>";
}
function th(s) { return "<th style='border:1px solid #e2e2e2;padding:6px 10px;background:#f6f7f9;text-align:left'>" + s + "</th>"; }
function td(s) { return "<td style='border:1px solid #e2e2e2;padding:6px 10px'>" + s + "</td>"; }
function requestHtml(project, vendor, items, total, comments, p, attachedCount) {
  const rows = items.map(function (it, i) { return "<tr>" + td(i + 1) + td(esc(it.description || "")) + td(num(it.quantity)) + td("$" + money(it.unitPrice)) + td("$" + money(lineTotal(it))) + "</tr>"; }).join("");
  const tbl = "<table style='border-collapse:collapse'><tr>" + th("#") + th("Item") + th("Qty") + th("Unit price") + th("Total") + "</tr>" + rows + "</table>";
  let body = "<p><b>Project / purpose:</b> " + esc(project) + "<br><b>Buying from:</b> " + esc(vendor) + "</p>" +
    tbl + "<p><b>Estimated total:</b> " + esc(total) + "</p>" +
    (comments ? "<p><b>Comments:</b> " + esc(comments) + "</p>" : "") +
    "<p><b>Attachments:</b> " + (attachedCount ? attachedCount + " attached to this email" : "none") + "</p>" +
    "<p style='color:#555;font-size:12px'>Submitted via the King EPCM staff intranet for approval.</p>";
  return wrapHtml("Purchase request", p, body);
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
function isoToday() { return new Date().toISOString().slice(0, 10); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
