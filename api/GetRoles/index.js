/*
 * POST /api/GetRoles — Azure Static Web Apps "rolesSource" function.
 *
 * Static Web Apps calls this automatically each time a user logs in (when
 * staticwebapp.config.json sets  "auth": { "rolesSource": "/api/GetRoles" }).
 * It returns the user's custom roles, so role membership is driven by Microsoft 365 /
 * Entra GROUP MEMBERSHIP instead of the one-by-one invitation system.
 *
 * How a person gets a role: add them to the matching Entra security group. By default we
 * match on the group's display name (case-insensitive):
 *     name contains "engineer"  -> engineer
 *     name contains "manager"   -> manager
 *     name contains "hr" / "human resource" -> hr
 *     name contains "admin"     -> admin   (e.g. "Intranet Admins")
 * You can override/extend this with the ROLE_GROUP_MAP app setting (JSON), keyed by group
 * id OR lowercased group name, e.g.  {"<group-guid>":"engineer","field crew":"engineer"}.
 *
 * Bootstrap admins (INTRANET_BOOTSTRAP_ADMINS, comma-separated emails) ALWAYS get "admin"
 * so you can't lock yourself out — defaults to ashi@kingepcm.com.
 *
 * App settings: AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT (optional),
 *   ROLE_GROUP_MAP (optional), INTRANET_BOOTSTRAP_ADMINS (optional).
 * Graph APPLICATION permission required (admin consent): GroupMember.Read.All
 *   (or Directory.Read.All) — plus the User.Read.All already granted.
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const BOOTSTRAP = (process.env.INTRANET_BOOTSTRAP_ADMINS || "ashi@kingepcm.com")
  .toLowerCase().split(/[,;\s]+/).filter(Boolean);
// Known group-id → role mappings (extend/override with the ROLE_GROUP_MAP app setting).
const DEFAULT_MAP = {
  "d87a76c1-731d-43c9-8d9b-e972d30cc985": "engineer",  // Engineers group
  "594e4070-71af-4726-ad32-1dd3794c27e1": "manager",   // Managers group
  "9302fdb2-c0b2-4070-9f11-9dfc8e5b9d45": "field",     // Field group
  "74f8e8da-1256-4ec7-aa2e-2d4217b2d9c9": "admin"      // Admin team group
};
let MAP = lowerKeys(DEFAULT_MAP);
try { MAP = Object.assign(MAP, lowerKeys(JSON.parse(process.env.ROLE_GROUP_MAP || "{}"))); } catch (e) {}

module.exports = async function (context, req) {
  const roles = [];
  try {
    const body = req.body || {};
    const claims = body.claims || [];
    const email = (body.userDetails && body.userDetails.indexOf("@") > -1 ? body.userDetails : "") ||
      claimVal(claims, ["preferred_username", "upn", "email", "emails", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress"]);
    const oid = claimVal(claims, ["http://schemas.microsoft.com/identity/claims/objectidentifier", "oid", "sub"]);

    // Bootstrap admins always get admin (lock-out safety net).
    if (email && BOOTSTRAP.indexOf(email.toLowerCase()) > -1) addRole(roles, "admin");

    const who = oid || email; // Graph accepts the object id or the userPrincipalName/email
    if (who && CLIENT_ID && CLIENT_SECRET) {
      const token = await appToken();
      const groups = await memberGroups(token, who);
      groups.forEach(function (g) {
        const r = MAP[String(g.id)] || MAP[(g.displayName || "").toLowerCase()] || keywordRole(g.displayName);
        if (r) addRole(roles, r);
      });
    }
  } catch (e) {
    context.log.error("GetRoles: " + (e && e.message || e));
    // Never fail the login — return whatever we have (at least bootstrap admin).
  }
  context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: { roles: roles } };
};

// Default name-based mapping (used when a group isn't in ROLE_GROUP_MAP).
function keywordRole(name) {
  const n = (name || "").toLowerCase();
  if (n.indexOf("admin") > -1) return "admin";
  if (n.indexOf("engineer") > -1) return "engineer";
  if (n.indexOf("manager") > -1) return "manager";
  if (n.indexOf("field") > -1) return "field";
  if (/(^|[^a-z])hr([^a-z]|$)/.test(n) || n.indexOf("human resource") > -1) return "hr";
  return "";
}

// The user's group memberships (transitive groups), app-only Graph.
async function memberGroups(token, who) {
  const out = [];
  let url = "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(who) +
    "/transitiveMemberOf/microsoft.graph.group?$select=id,displayName&$top=200";
  for (let g = 0; url && g < 10; g++) {
    const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) {
      // Fall back to the direct (non-transitive) membership endpoint if transitive isn't permitted.
      if (g === 0) { return directGroups(token, who); }
      throw new Error("graph memberOf " + r.status);
    }
    const j = await r.json();
    (j.value || []).forEach(function (x) { if (x.id) out.push({ id: x.id, displayName: x.displayName || "" }); });
    url = j["@odata.nextLink"] || "";
  }
  return out;
}
async function directGroups(token, who) {
  const out = [];
  let url = "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(who) + "/memberOf?$select=id,displayName&$top=200";
  for (let g = 0; url && g < 10; g++) {
    const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (!r.ok) throw new Error("graph memberOf " + r.status);
    const j = await r.json();
    (j.value || []).forEach(function (x) { if (x.id) out.push({ id: x.id, displayName: x.displayName || "" }); });
    url = j["@odata.nextLink"] || "";
  }
  return out;
}

async function appToken() {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}

/* ---------- helpers ---------- */
function addRole(arr, r) { if (r && arr.indexOf(r) === -1) arr.push(r); }
function claimVal(claims, types) {
  const want = types.map(function (t) { return t.toLowerCase(); });
  const c = (claims || []).find(function (x) { return want.indexOf((x.typ || "").toLowerCase()) > -1 && x.val; });
  return c ? c.val : "";
}
function lowerKeys(o) { const out = {}; Object.keys(o || {}).forEach(function (k) { out[k.toLowerCase()] = o[k]; }); return out; }
