/*
 * GET /api/my-reviews
 * Lists the SIGNED-IN staff member's KRA / development-review files and returns
 * short-lived download links — WITHOUT giving staff any SharePoint access.
 *
 * Folder layout (configurable):
 *   <site library> / <REVIEW_BASE_PATH> / <Staff Name folder> / [KRA] [Development Review] / files
 * e.g. REVIEW_BASE_PATH = "03. HR & Payroll/03. Employee Records"
 *
 * The function uses the app's OWN identity (app-only Graph), finds the folder that
 * best matches the signed-in person's name (handles legal vs preferred names), and
 * returns only that person's files. A staff member can never see another's reviews,
 * and gets no edit/browse access in SharePoint.
 *
 * App settings:
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET  (reuse the intranet app registration; it needs
 *                                      Graph APPLICATION permission Sites.Read.All + admin consent)
 *   GRAPH_TENANT       (optional; defaults to the King EPCM tenant)
 *   REVIEW_SITE_PATH   e.g. "kingepcm.sharepoint.com:/sites/ManagementTeam"
 *   REVIEW_BASE_PATH   e.g. "03. HR & Payroll/03. Employee Records"
 *   REVIEW_LIBRARY     (optional) document-library name if not the site default
 *   REVIEW_FOLDER_MAP  (optional) JSON of exact overrides, e.g.
 *                      {"ashi@kingepcm.com":"Angela Shi","mjohnson@kingepcm.com":"Michael Johnson"}
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.REVIEW_SITE_PATH;
const BASE_PATH = process.env.REVIEW_BASE_PATH;
const LIBRARY = process.env.REVIEW_LIBRARY || "";
let FOLDER_MAP = {};
try { FOLDER_MAP = JSON.parse(process.env.REVIEW_FOLDER_MAP || "{}"); } catch (e) { FOLDER_MAP = {}; }
// Which subfolders (inside each staff folder) hold review files. Override with the
// app setting REVIEW_SUBFOLDERS (comma-separated keywords) if your naming differs.
const REVIEW_SUBFOLDER_RE = new RegExp(
  "(" + (process.env.REVIEW_SUBFOLDERS || "kra,development review,performance review,review,appraisal,kpi")
    .split(",").map(function (s) { return s.trim().replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }).filter(Boolean).join("|") + ")", "i");

let _driveId; // cached across warm invocations

module.exports = async function (context, req) {
  if (!CLIENT_ID || !CLIENT_SECRET || !SITE_PATH || !BASE_PATH) {
    context.res = json(501, { error: "Reviews not configured" }); return;
  }
  const info = principalInfo(req);
  if (!info.display && !info.email) { context.res = json(401, { error: "Not signed in" }); return; }

  try {
    const token = await appToken();

    // The SWA sign-in token often carries only the email (no name claims), so look
    // up the person's real name from Microsoft by email. Needs Graph application
    // permission User.Read.All (admin consent). Falls back silently if unavailable.
    if (info.email && (!info.given || !info.surname)) {
      const u = await lookupUser(token, info.email).catch(function () { return null; });
      if (u) {
        info.display = u.displayName || info.display;
        info.given = u.givenName || info.given;
        info.surname = u.surname || info.surname;
      }
    }

    const driveId = await resolveDrive(token);

    // 1) Find the staff member's folder inside the Employee Records base folder.
    let folderName = FOLDER_MAP[(info.email || "").toLowerCase()];
    let staffFolders = null;
    if (!folderName) {
      try { staffFolders = (await listChildren(token, driveId, trim(BASE_PATH))).filter(function (i) { return i.folder; }); }
      catch (e) { if (String(e.message).indexOf("404") > -1) { context.res = json(501, { error: "Base folder not found" }); return; } throw e; }
      // Admin-only diagnostic: /api/my-reviews?debug=1 → shows the parsed name + folder names.
      if (req.query && req.query.debug && isAdminReq(req)) {
        context.res = json(200, { you: info, folders: staffFolders.map(function (f) { return f.name; }) });
        return;
      }
      const match = matchStaffFolder(staffFolders, info);
      folderName = match && match.name;
    }
    if (!folderName) { context.res = json(200, { name: info.display, files: [], notFound: true }); return; }

    // 2) List that folder + one level of subfolders (KRA / Development Review).
    let topItems;
    try { topItems = await listChildren(token, driveId, trim(BASE_PATH) + "/" + folderName); }
    catch (e) { if (String(e.message).indexOf("404") > -1) { context.res = json(200, { name: info.display, files: [], notFound: true }); return; } throw e; }

    // Only pull files from the KRA / Development Review subfolders — never the other
    // documents in the staff folder (contracts, etc.).
    const files = [];
    for (const it of topItems) {
      if (it.folder && REVIEW_SUBFOLDER_RE.test(deNum(it.name))) {
        await collectFiles(token, driveId, it.id, it.name, 1, files); // include one nested level (e.g. by year)
      }
    }
    files.sort(function (a, b) { return String(b.modified).localeCompare(String(a.modified)); });
    context.res = json(200, { name: folderName, files: files }, { "Cache-Control": "no-store" });
  } catch (err) {
    context.log.error(err);
    context.res = json(502, { error: "Could not load reviews" });
  }
};

/* ---------- name matching (handles leading numbers + legal vs preferred) ---------- */
function matchStaffFolder(folders, info) {
  const gname = norm(info.given), sname = norm(info.surname);
  const uTokens = uniq([].concat(toks(info.display), toks(info.given), toks(info.surname), emailToks(info.email)))
    .filter(function (t) { return !/^\d+$/.test(t); }); // ignore pure-number tokens
  let best = null, bestScore = 0, fullMatches = 0;
  folders.forEach(function (f) {
    const ft = toks(deNum(f.name)); // strip leading "01." style numbering before matching
    if (!ft.length) return;
    const has = function (t) { return t && ft.indexOf(t) > -1; };
    const shared = uTokens.filter(has).length;
    const hasSur = has(sname);
    const hasGiven = has(gname);
    const initial = gname && ft.some(function (t) { return t !== sname && t[0] === gname[0]; });
    const exact = norm(deNum(f.name)) === norm(info.display);
    const both = hasSur && hasGiven;                 // given + surname both present (any order)
    const strong = exact || both || (hasSur && initial) || (shared >= 2);
    if (!strong) return;
    if (both || exact) fullMatches++;
    const score = shared + (hasSur ? 2 : 0) + (hasGiven ? 2 : 0) + (exact ? 5 : 0);
    if (score > bestScore) { bestScore = score; best = f; }
  });
  // Only treat as ambiguous if two folders both fully matched given+surname (real duplicate).
  if (fullMatches > 1) return null;
  return best;
}
function deNum(s) { return String(s || "").replace(/^\s*\d+\s*[.)\-:_]*\s*/, ""); }
function norm(s) { return String(s || "").toLowerCase().normalize("NFKD").replace(/[^a-z0-9 ]/g, " ").replace(/\s+/g, " ").trim(); }
function toks(s) { return norm(s).split(" ").filter(function (t) { return t.length > 1; }); }
function emailToks(e) { return String(e || "").split("@")[0].split(/[^a-z0-9]+/i).filter(function (t) { return t.length > 1; }).map(function (t) { return t.toLowerCase(); }); }
function uniq(a) { return a.filter(function (v, i) { return a.indexOf(v) === i; }); }

function principalInfo(req) {
  const out = { display: "", given: "", surname: "", email: "" };
  try {
    const h = req.headers["x-ms-client-principal"];
    if (!h) return out;
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    const claims = p.claims || [];
    const nameClaim = claimVal(claims, ["name"]);
    if (nameClaim && nameClaim.indexOf("@") === -1) out.display = nameClaim;
    out.given = claimVal(claims, ["given_name", "givenname", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/givenname"]);
    out.surname = claimVal(claims, ["family_name", "surname", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/surname"]);
    out.email = (p.userDetails && p.userDetails.indexOf("@") > -1 ? p.userDetails : "") ||
      claimVal(claims, ["email", "preferred_username", "upn", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress"]);
    if ((!out.given || !out.surname) && out.display) {
      const parts = out.display.split(/\s+/);
      if (parts.length >= 2) { out.given = out.given || parts[0]; out.surname = out.surname || parts[parts.length - 1]; }
    }
    if (!out.display && out.given && out.surname) out.display = out.given + " " + out.surname;
  } catch (e) {}
  return out;
}
function claimVal(claims, types) {
  const c = (claims || []).find(function (x) { return types.indexOf((x.typ || "").toLowerCase()) > -1 && x.val; });
  return c ? c.val : "";
}
function isAdminReq(req) {
  try {
    const h = req.headers["x-ms-client-principal"];
    if (!h) return false;
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    return Array.isArray(p.userRoles) && p.userRoles.indexOf("admin") > -1;
  } catch (e) { return false; }
}

/* ---------- Graph helpers ---------- */
function mapFile(it, folder) {
  return { name: it.name, folder: folder || "", size: it.size || 0, modified: it.lastModifiedDateTime || "", download: it["@microsoft.graph.downloadUrl"] || "" };
}
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
    const drives = await gget(token, "https://graph.microsoft.com/v1.0/sites/" + site.id + "/drives?$select=id,name");
    const d = (drives.value || []).find(function (x) { return (x.name || "").toLowerCase() === LIBRARY.toLowerCase(); });
    if (d) { _driveId = d.id; return _driveId; }
  }
  const drive = await gget(token, "https://graph.microsoft.com/v1.0/sites/" + site.id + "/drive?$select=id");
  _driveId = drive.id;
  return _driveId;
}
async function listChildren(token, driveId, path) {
  const enc = path.split("/").map(encodeURIComponent).join("/");
  const data = await gget(token, "https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + enc + ":/children?$top=200");
  return data.value || [];
}
async function listChildrenById(token, driveId, itemId) {
  const data = await gget(token, "https://graph.microsoft.com/v1.0/drives/" + driveId + "/items/" + itemId + "/children?$top=200");
  return data.value || [];
}
// Gather files from a folder (and up to `depth` more nested levels), tagging each
// with its containing folder so the page can show "KRA" / "Development Review".
async function collectFiles(token, driveId, folderId, label, depth, out) {
  const kids = await listChildrenById(token, driveId, folderId).catch(function () { return []; });
  for (const k of kids) {
    if (k.file) out.push(mapFile(k, label));
    else if (k.folder && depth > 0) await collectFiles(token, driveId, k.id, label + " / " + k.name, depth - 1, out);
  }
}
async function lookupUser(token, email) {
  return gget(token, "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(email) + "?$select=displayName,givenName,surname");
}
async function gget(token, url) {
  const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
  if (!r.ok) throw new Error("graph " + r.status);
  return r.json();
}
function trim(s) { return String(s || "").replace(/^\/+|\/+$/g, ""); }
function json(status, body, extra) { return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body }; }
