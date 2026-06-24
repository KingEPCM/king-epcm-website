/*
 * POST /api/reimbursement-submit  — submits a reimbursement claim and EMAILS it to the
 * right department automatically (no mailto): mileage & expense → Accounting; PEO/PEng
 * professional fees → HR (for review/approval). Expense & professional receipts are
 * attached to the email and (if SharePoint is configured) also archived with a link.
 * Every claim is logged for the staff member's own history.
 *
 * Sends AS the signed-in submitter's mailbox (Graph Mail.Send "send as any user"),
 * so claims reach Accounting/HR from the employee and replies go back to them.
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT,
 *   REIMBURSE_FROM (optional fallback sender if a submitter address isn't available),
 *   REIMBURSE_TO_ACCOUNTING (default accounting@kingepcm.com),
 *   REIMBURSE_TO_HR (default hr@kingepcm.com),
 *   REVIEW_SITE_PATH + REIMBURSE_BASE_PATH (optional archive), NEWS_STORAGE_CONNECTION (history).
 * Graph APPLICATION permissions: Mail.Send (always), Sites.ReadWrite.All (archive, optional).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.REVIEW_SITE_PATH;
const LIBRARY = process.env.REVIEW_LIBRARY || "";
const BASE_PATH = process.env.REIMBURSE_BASE_PATH;
const CONN = process.env.NEWS_STORAGE_CONNECTION; // reuse the news storage account for the claim log
const FROM = process.env.REIMBURSE_FROM || process.env.REVIEW_NOTIFY_FROM || "";
const TO_ACCOUNTING = process.env.REIMBURSE_TO_ACCOUNTING || "accounting@kingepcm.com";
const TO_HR = process.env.REIMBURSE_TO_HR || "hr@kingepcm.com";
const MAX_FILE = 4 * 1024 * 1024;   // 4 MB per receipt
const MAX_EMAIL = 3 * 1024 * 1024;  // attach receipts to the email only under ~3 MB total
const { TableClient } = require("@azure/data-tables");
let _driveId;

module.exports = async function (context, req) {
 try {
  const p = principal(req);
  if (!p.email) { context.res = json(401, { ok: false, error: "Not signed in" }); return; }
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(501, { ok: false, error: "Email not configured" }); return; }
  const b = req.body || {};
  const type = /mileage/i.test(b.type) ? "mileage" : /prof/i.test(b.type) ? "professional" : "expense";
  const isProf = type === "professional";
  const dept = isProf ? "HR" : "Accounting";
  const to = isProf ? TO_HR : TO_ACCOUNTING;
  // Send the claim AS the person who submitted it (their mailbox), so it reaches
  // Accounting/HR from the employee and replies go straight back to them. Falls back
  // to REIMBURSE_FROM only if a submitter address somehow isn't available.
  const fromAddr = p.email || FROM;

  const token = await appToken();

  // ----- Mileage → Accounting (summary email, no receipts) -----
  if (type === "mileage") {
    const trips = (Array.isArray(b.trips) ? b.trips : []).filter(function (t) { return t && (t.date || t.from || t.to || t.km || t.note); });
    if (!trips.length) { context.res = json(400, { ok: false, error: "No trips" }); return; }
    await sendMail(token, fromAddr, to, p.email, "Mileage reimbursement — " + (p.name || p.email), mileageHtml(trips, b.total, b.notes, p), []);
    const recSum = trips.map(function (t, i) { return (i + 1) + ". " + (t.date || "(no date)") + " " + (t.from || "?") + "→" + (t.to || "?") + " " + (t.km || "0") + "km"; }).join("; ");
    await recordClaim(context, p, "mileage", b.total || "", recSum, "", trips.length);
    context.res = json(200, { ok: true, dept: dept });
    return;
  }

  // ----- Expense → Accounting / Professional → HR (receipts attached + optional archive) -----
  const items = Array.isArray(b.items) ? b.items : [];
  if (!items.length) { context.res = json(400, { ok: false, error: "No items" }); return; }
  let totalBytes = 0;
  for (const it of items) {
    if (!it.description || !it.amount || !it.receipt || !it.receipt.base64 || (!isProf && !it.project)) {
      context.res = json(400, { ok: false, error: isProf ? "Every item needs a description, amount, and receipt." : "Every item needs a project, description, amount, and receipt." }); return;
    }
    const sz = Buffer.from(it.receipt.base64, "base64").length;
    if (sz > MAX_FILE) { context.res = json(413, { ok: false, error: "A receipt is larger than 4 MB. Please reduce its size." }); return; }
    totalBytes += sz;
  }

  function attName(it, n) { const ext = extOf(it.receipt.name); const tag = isProf ? "PEO-PEng" : it.project; return sanitize("Item " + n + " - " + tag + " - " + it.description).slice(0, 90) + ext; }

  // Optional SharePoint archive → folder link.
  let folderUrl = "";
  if (SITE_PATH && BASE_PATH) {
    try {
      const driveId = await resolveDrive(token);
      const claimName = sanitize(shortToday() + " - " + (p.name || p.email) + " - " + (isProf ? "Professional Fees" : "Expense Claim"));
      const folder = await createFolder(token, driveId, trim(BASE_PATH), claimName);
      if (folder && folder.id) {
        let n = 0;
        for (const it of items) { n++; await uploadFile(token, driveId, folder.id, attName(it, n), Buffer.from(it.receipt.base64, "base64"), it.receipt.contentType || "application/octet-stream"); }
        await uploadFile(token, driveId, folder.id, "Claim summary.txt", Buffer.from(buildSummary(items, b.total, b.notes, p, isProf), "utf8"), "text/plain");
        folderUrl = folder.webUrl || "";
      }
    } catch (e) { context.log.error(e); }
  }

  // Attach receipts to the email when small enough; otherwise rely on the archive link.
  let attachments = [];
  if (totalBytes <= MAX_EMAIL) {
    let n = 0;
    attachments = items.map(function (it) { n++; return { "@odata.type": "#microsoft.graph.fileAttachment", name: attName(it, n), contentType: it.receipt.contentType || "application/octet-stream", contentBytes: it.receipt.base64 }; });
  }
  if (!attachments.length && !folderUrl) { context.res = json(413, { ok: false, error: "Receipts are too large to email. Please reduce their size." }); return; }

  const subject = (isProf ? "PEO/PEng fee reimbursement" : "Expense reimbursement") + " — " + (p.name || p.email);
  await sendMail(token, fromAddr, to, p.email, subject, itemsHtml(items, b.total, b.notes, p, isProf, folderUrl, attachments.length > 0), attachments);

  const recSum = items.map(function (it, i) { return (i + 1) + ". " + (isProf ? "" : "[" + it.project + "] ") + it.description + " $" + money(it.amount); }).join("; ");
  await recordClaim(context, p, isProf ? "professional" : "expense", b.total || "", recSum, folderUrl, items.length);

  context.res = json(200, { ok: true, dept: dept, folderUrl: folderUrl });
 } catch (e) {
  context.log.error(e);
  context.res = json(502, { ok: false, error: "Could not submit the claim" });
 }
};

// Save a lightweight record of the claim so the staff member can see their own history
// (and avoid duplicate submissions). Best-effort — reuses the news storage account.
async function recordClaim(context, p, type, total, summary, folderUrl, count) {
  if (!CONN) return;
  try {
    const t = TableClient.fromConnectionString(CONN, "reimbursements");
    try { await t.createTable(); } catch (e) {}
    await t.createEntity({
      partitionKey: keyOf(p.email), rowKey: String(Date.now()) + String(Math.floor(Math.random() * 1000)),
      type: type, total: String(total || ""), summary: String(summary || "").slice(0, 500),
      folderUrl: folderUrl || "", itemCount: count || 0, byName: p.name || "", byEmail: p.email,
      date: shortToday(), createdAt: new Date().toISOString()
    });
  } catch (e) { context.log.error(e); }
}
function money(a) { var n = Number(String(a == null ? "" : a).replace(/[^0-9.]/g, "")); return (isNaN(n) ? 0 : n).toFixed(2); }
function keyOf(email) { return String(email || "").toLowerCase().replace(/[\\/#?\t\n\r]/g, "_"); }

function buildSummary(items, total, notes, p, isProf) {
  const lines = [];
  lines.push(isProf ? "Professional fees (PEO/PEng) reimbursement" : "Expense reimbursement claim");
  lines.push("Submitted by: " + (p.name || "") + " <" + p.email + ">");
  lines.push("Date: " + isoToday());
  lines.push("");
  items.forEach(function (it, i) {
    lines.push((i + 1) + ". " + (isProf ? "" : "[" + it.project + "] ") + it.description + " — $" + money(it.amount));
  });
  lines.push("");
  lines.push("Total: " + (total || ""));
  if (notes) lines.push("Notes: " + notes);
  return lines.join("\r\n");
}

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
function mileageHtml(trips, total, notes, p) {
  const rows = trips.map(function (t, i) { return "<tr>" + td(i + 1) + td(esc(t.date || "")) + td(esc((t.from || "") + " → " + (t.to || ""))) + td(esc((t.km || "") + " km")) + td(esc(t.note || "")) + "</tr>"; }).join("");
  const tbl = "<table style='border-collapse:collapse'><tr>" + th("#") + th("Date") + th("Route") + th("Distance") + th("Note") + "</tr>" + rows + "</table>";
  return wrapHtml("Mileage reimbursement", p, tbl + "<p><b>Total distance:</b> " + esc(total || "") + " (Accounting applies the per-km rate.)</p>" + (notes ? "<p><b>Notes:</b> " + esc(notes) + "</p>" : ""));
}
function itemsHtml(items, total, notes, p, isProf, folderUrl, attached) {
  const rows = items.map(function (it, i) { return "<tr>" + td(i + 1) + (isProf ? "" : td(esc(it.project || ""))) + td(esc(it.description || "")) + td("$" + money(it.amount)) + "</tr>"; }).join("");
  const head = "<tr>" + th("#") + (isProf ? "" : th("Project")) + th("Description") + th("Amount") + "</tr>";
  let body = "<table style='border-collapse:collapse'>" + head + rows + "</table><p><b>Total:</b> " + esc(total || "") + "</p>" + (notes ? "<p><b>Notes:</b> " + esc(notes) + "</p>" : "");
  body += "<p><b>Receipts:</b> " + (attached ? "attached to this email" : "see the linked folder") + (folderUrl ? " · <a href='" + esc(folderUrl) + "'>" + esc(folderUrl) + "</a>" : "") + "</p>";
  return wrapHtml(isProf ? "PEO/PEng professional-fee reimbursement" : "Expense reimbursement", p, body);
}
function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

/* ---------- Graph ---------- */
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
async function createFolder(token, driveId, parentPath, name) {
  const enc = parentPath.split("/").map(encodeURIComponent).join("/");
  const r = await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + enc + ":/children", {
    method: "POST", headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ name: name, folder: {}, "@microsoft.graph.conflictBehavior": "rename" })
  });
  if (!r.ok) throw new Error("folder " + r.status + " " + (await r.text().catch(function () { return ""; })));
  return r.json();
}
async function uploadFile(token, driveId, parentItemId, filename, buffer, contentType) {
  const r = await fetch("https://graph.microsoft.com/v1.0/drives/" + driveId + "/items/" + parentItemId + ":/" + encodeURIComponent(filename) + ":/content", {
    method: "PUT", headers: { Authorization: "Bearer " + token, "Content-Type": contentType }, body: buffer
  });
  if (!r.ok) throw new Error("upload " + r.status);
}
async function gget(token, url) {
  const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) throw new Error("graph " + r.status);
  return r.json();
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
function extOf(n) { const m = String(n || "").match(/\.[a-z0-9]{1,5}$/i); return m ? m[0].toLowerCase() : ""; }
function sanitize(s) { return String(s || "claim").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 120); }
function isoToday() { return new Date().toISOString().slice(0, 10); }
function shortToday() { var d = new Date(); return String(d.getFullYear()).slice(-2) + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0"); }
function trim(s) { return String(s || "").replace(/^\/+|\/+$/g, ""); }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
