/*
 * GET /api/my-mail  → { unread: <number> }
 * The signed-in person's Inbox unread count, read server-side with the app's own
 * identity (app-only Graph) so it works on mobile (no in-browser Microsoft sign-in).
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional).
 * Graph APPLICATION permission required (admin consent): Mail.Read
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;

module.exports = async function (context, req) {
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(501, { error: "Not configured" }); return; }
  const email = principalEmail(req);
  if (!email) { context.res = json(401, { error: "Not signed in" }); return; }
  try {
    const token = await appToken();
    const j = await gget(token, "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(email) + "/mailFolders/Inbox?$select=unreadItemCount");
    const n = j && typeof j.unreadItemCount === "number" ? j.unreadItemCount : null;
    context.res = json(200, { unread: n }, { "Cache-Control": "no-store" });
  } catch (e) { context.log.error(e); context.res = json(502, { error: "Could not load mail" }); }
};

async function appToken() {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function gget(token, url) {
  const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) throw new Error("graph " + r.status);
  return r.json();
}
function principalEmail(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return "";
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    if (p.userDetails && p.userDetails.indexOf("@") > -1) return p.userDetails;
    const c = (p.claims || []).find(function (x) { return ["email", "preferred_username", "upn", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress"].indexOf((x.typ || "").toLowerCase()) > -1 && x.val; });
    return c ? c.val : "";
  } catch (e) { return ""; }
}
function json(status, body, extra) { return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body }; }
