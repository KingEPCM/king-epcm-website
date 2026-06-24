/*
 * GET /api/my-calendar?start=<ISO>&end=<ISO>  → { events: [ {subject,start,end,isAllDay,webLink,__src} ] }
 * The signed-in person's events plus the company group calendar, read server-side
 * (app-only Graph) so the calendar works on mobile. Defaults to a 14-day window.
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional),
 *   COMPANY_CAL_EMAIL (optional, default group@kingepcm.com).
 * Graph APPLICATION permissions (admin consent): Calendars.Read (personal events),
 *   Group.Read.All (company group calendar — best-effort; degrades to personal only).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const COMPANY_CAL_EMAIL = process.env.COMPANY_CAL_EMAIL || "group@kingepcm.com";
const TZ = 'Eastern Standard Time';
let _companyGroupId; // cached across warm invocations

module.exports = async function (context, req) {
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(501, { error: "Not configured" }); return; }
  const email = principalEmail(req);
  if (!email) { context.res = json(401, { error: "Not signed in" }); return; }

  // Window: explicit start/end (e.g. the month grid) or a default 14-day range.
  let start, end;
  try { start = req.query && req.query.start ? new Date(req.query.start) : startOfToday(); } catch (e) { start = startOfToday(); }
  try { end = req.query && req.query.end ? new Date(req.query.end) : new Date(start.getTime() + 14 * 86400000); } catch (e) { end = new Date(start.getTime() + 14 * 86400000); }
  if (isNaN(start)) start = startOfToday();
  if (isNaN(end)) end = new Date(start.getTime() + 14 * 86400000);

  try {
    const token = await appToken();
    const qs = "startDateTime=" + start.toISOString() + "&endDateTime=" + end.toISOString() +
      "&$orderby=start/dateTime&$top=400&$select=subject,start,end,isAllDay,webLink";

    // Personal events.
    const mine = await gcal(token, "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(email) + "/calendarView?" + qs)
      .then(function (j) { return (j.value || []).map(function (e) { e.__src = "me"; return e; }); })
      .catch(function (e) { context.log.error(e); return []; });

    // Company group calendar (best-effort).
    let company = [];
    try {
      const gid = await companyGroupId(token);
      if (gid) {
        const j = await gcal(token, "https://graph.microsoft.com/v1.0/groups/" + gid + "/calendarView?" + qs);
        company = (j.value || []).map(function (e) { e.__src = "company"; return e; });
      }
    } catch (e) { context.log.error(e); }

    context.res = json(200, { events: mine.concat(company) }, { "Cache-Control": "no-store" });
  } catch (e) { context.log.error(e); context.res = json(502, { error: "Could not load calendar" }); }
};

async function companyGroupId(token) {
  if (_companyGroupId) return _companyGroupId;
  const j = await gget(token, "https://graph.microsoft.com/v1.0/groups?$filter=mail eq '" + COMPANY_CAL_EMAIL + "'&$select=id");
  const g = (j.value || [])[0];
  if (g) { _companyGroupId = g.id; return _companyGroupId; }
  return null;
}
function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
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
// Calendar GET with the Eastern-time preference so times render correctly.
async function gcal(token, url) {
  const r = await fetch(url, { headers: { Authorization: "Bearer " + token, Prefer: 'outlook.timezone="' + TZ + '"' } });
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
