/*
 * POST /api/project-kickoff — sends a client project kick-off email.
 * Sent AS the signed-in person (Graph Mail.Send "send as any user") so it leaves their
 * own mailbox, lands in their Sent Items, and the client's replies come straight back to them.
 *
 * The page builds the email body (so the live Preview is exactly what's sent) and posts it here.
 * Body (JSON): { to, cc?, subject, bodyHtml, bodyText? }
 *   - to / cc: comma-separated email addresses (to is required)
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional).
 * Graph APPLICATION permission: Mail.Send (admin consent) — already used by /api/hr-inquiry.
 */
const fs = require("fs");
const path = require("path");
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const MAX_BODY = 700 * 1024; // ~700 KB of HTML is plenty for a text email
const LOGO_CID = "kelogo";   // signature logo, referenced as <img src="cid:kelogo">

// Load the inline signature logo once (base64). Optional — email still sends if it's missing.
let LOGO_B64 = "";
try { LOGO_B64 = fs.readFileSync(path.join(__dirname, "logo.png")).toString("base64"); } catch (e) { LOGO_B64 = ""; }

module.exports = async function (context, req) {
 try {
  const p = principal(req);
  if (!p.email) { context.res = json(401, { ok: false, error: "Not signed in" }); return; }
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(501, { ok: false, error: "Email not configured" }); return; }

  const b = req.body || {};
  const to = cleanList(b.to);
  const cc = cleanList(b.cc);
  const subject = String(b.subject || "").trim();
  const bodyHtml = String(b.bodyHtml || "");

  if (!to.length) { context.res = json(400, { ok: false, error: "A valid client email address is required." }); return; }
  if (!subject) { context.res = json(400, { ok: false, error: "A subject line is required." }); return; }
  if (!bodyHtml.trim()) { context.res = json(400, { ok: false, error: "The email body is empty." }); return; }
  if (Buffer.byteLength(bodyHtml, "utf8") > MAX_BODY) { context.res = json(413, { ok: false, error: "The email is too large to send." }); return; }

  // Attach the King EPCM logo inline only if the body references it.
  const attachments = [];
  if (LOGO_B64 && bodyHtml.indexOf("cid:" + LOGO_CID) > -1) {
    attachments.push({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: "king-epcm.png",
      contentType: "image/png",
      isInline: true,
      contentId: LOGO_CID,
      contentBytes: LOGO_B64
    });
  }

  const token = await appToken();
  await sendMail(token, p.email, to, cc, p.email, subject, bodyHtml, attachments);

  context.res = json(200, { ok: true, sentFrom: p.email });
 } catch (e) {
  context.log.error(e);
  context.res = json(502, { ok: false, error: "Could not send the email. Please try again." });
 }
};

/* ---------- Email (Graph Mail.Send) ---------- */
async function sendMail(token, from, to, cc, replyTo, subject, html, attachments) {
  const msg = {
    subject: subject,
    body: { contentType: "HTML", content: html },
    toRecipients: to.map(addr),
    replyTo: replyTo ? [addr(replyTo)] : undefined
  };
  if (cc && cc.length) msg.ccRecipients = cc.map(addr);
  if (attachments && attachments.length) msg.attachments = attachments;
  const r = await fetch("https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(from) + "/sendMail", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({ message: msg, saveToSentItems: true })
  });
  if (!r.ok) throw new Error("sendMail " + r.status + " " + (await r.text().catch(function () { return ""; })));
}
function addr(e) { return { emailAddress: { address: String(e).trim() } }; }

/* ---------- Graph token ---------- */
async function appToken() {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}

/* ---------- helpers ---------- */
function cleanList(v) {
  if (!v) return [];
  return String(v).split(/[,;]+/).map(function (s) { return s.trim(); })
    .filter(function (s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); });
}
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
