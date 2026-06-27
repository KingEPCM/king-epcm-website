/*
 * GET /api/staff-directory — a list of current staff for name pickers on the forms.
 * Source: Microsoft Entra (Graph app-only). Returns only real people:
 *   - userType eq 'Member'         (no external guests)
 *   - accountEnabled eq true       (shared mailboxes / rooms have sign-in blocked → excluded)
 *   - a mailbox (mail present)
 * Response: { configured, staff: [ { name, email, title } ] }  sorted by name.
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional).
 * Graph APP permission: User.Read.All (already granted for the review tools).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;

let _cache = null, _cacheTs = 0;
const TTL = 10 * 60 * 1000; // 10 minutes

module.exports = async function (context, req) {
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(200, { configured: false, staff: [] }); return; }
  try {
    if (_cache && Date.now() - _cacheTs < TTL) { context.res = json(200, { configured: true, staff: _cache }); return; }
    const token = await appToken();
    const staff = await loadStaff(token);
    _cache = staff; _cacheTs = Date.now();
    context.res = json(200, { configured: true, staff: staff });
  } catch (e) {
    context.log.error(e);
    context.res = json(200, { configured: true, staff: _cache || [], error: "lookup failed" });
  }
};

async function loadStaff(token) {
  const seen = {}, out = [];
  let url = "https://graph.microsoft.com/v1.0/users?$select=displayName,mail,jobTitle,userType,accountEnabled" +
    "&$filter=" + encodeURIComponent("accountEnabled eq true and userType eq 'Member'") + "&$top=999";
  for (let guard = 0; url && guard < 20; guard++) {
    const d = await gget(token, url);
    (d.value || []).forEach(function (u) {
      const mail = (u.mail || "").trim();
      if (!mail) return;                                  // no mailbox → skip
      const key = mail.toLowerCase();
      if (seen[key]) return; seen[key] = 1;
      out.push({ name: (u.displayName || mail).trim(), email: mail, title: (u.jobTitle || "").trim() });
    });
    url = d["@odata.nextLink"] || "";
  }
  out.sort(function (a, b) { return a.name.localeCompare(b.name); });
  return out;
}

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
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
