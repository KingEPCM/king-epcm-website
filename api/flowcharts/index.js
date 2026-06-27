/*
 * GET /api/flowcharts — a LIVE list of process flowcharts, read from SharePoint.
 * Searches the document library under BASE_PATH for PDFs whose name contains
 * "flowchart", groups them by their top-level department folder, and returns:
 *   { configured, categories: [ { category, items: [ { title, url } ] } ] }
 * The front end falls back to the static procedures.json if this isn't configured.
 *
 * App settings (all optional — sensible defaults for King EPCM):
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT      (Graph app-only)
 *   FLOW_SITE_PATH   Graph site path (default kingepcm.sharepoint.com root site)
 *   FLOW_LIBRARY     drive/library name (default: the site's default documents library)
 *   FLOW_BASE_PATH   folder to scope to (default "03. Assets & Templates")
 *   FLOW_DEPT_MAP    JSON map of raw dept folder → friendly category name
 * Graph APP permission: Sites.Read.All (or the Sites.ReadWrite.All already granted).
 */
const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.FLOW_SITE_PATH || "kingepcm.sharepoint.com";
const LIBRARY = process.env.FLOW_LIBRARY || "";
const BASE_PATH = process.env.FLOW_BASE_PATH || "03. Assets & Templates";

const DEFAULT_DEPT_MAP = {
  "04. civil": "Civil & Municipal Design",
  "05. geohyd": "Geotechnical & Hydrogeology",
  "07. field": "Survey & Field Inspections",
  "02. natural heritage environments": "Natural Heritage Environmental",
  "09. indirect deliverables": "Office & Operations"
};
let DEPT_MAP = DEFAULT_DEPT_MAP;
try { DEPT_MAP = Object.assign({}, DEFAULT_DEPT_MAP, lowerKeys(JSON.parse(process.env.FLOW_DEPT_MAP || "{}"))); } catch (e) {}
// Preferred display order = the order categories appear in the map, then anything else A–Z.
const ORDER = Object.keys(DEFAULT_DEPT_MAP).map(function (k) { return DEFAULT_DEPT_MAP[k]; });

let _driveId, _cache = null, _cacheTs = 0;
const TTL = 30 * 60 * 1000;

module.exports = async function (context, req) {
  const debug = req && req.query && (req.query.debug === "1" || req.query.debug === "true");
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(200, { configured: false, categories: [], debug: debug ? { reason: "Graph client id/secret not set" } : undefined }); return; }
  try {
    if (!debug && _cache && Date.now() - _cacheTs < TTL) { context.res = json(200, { configured: true, categories: _cache }); return; }
    const token = await appToken();
    const driveId = await resolveDrive(token);
    const dbg = { hits: 0, matched: 0, scopedOut: 0, sampleMatched: [], sampleScoped: [] };
    const items = await searchFlowcharts(token, driveId, dbg);
    const categories = group(items);
    _cache = categories; _cacheTs = Date.now();
    context.res = json(200, {
      configured: true, categories: categories,
      debug: debug ? Object.assign({ driveId: driveId, basePath: BASE_PATH, categories: categories.length, items: items.length }, dbg) : undefined
    });
  } catch (e) {
    context.log.error(e);
    // Signal "not live" so the front end falls back to the static list gracefully.
    context.res = json(200, { configured: false, categories: _cache || [], error: String((e && e.message) || e) });
  }
};

// One cheap drive-wide search, then scope + group by the folder path found in each
// file's webUrl (search results don't populate parentReference.path, but the webUrl
// does contain the full path). Far fewer Graph calls than walking the tree.
async function searchFlowcharts(token, driveId, dbg) {
  return searchAndScope(token, driveId, dbg, "flowchart", function (name) {
    return /\.pdf$/i.test(name) && name.toLowerCase().indexOf("flowchart") > -1;
  }, function (it, dept) { return { name: it.name, url: it.webUrl, dept: dept }; });
}

async function searchAndScope(token, driveId, dbg, q, keep, make) {
  const seen = {}, out = [];
  let url = "https://graph.microsoft.com/v1.0/drives/" + driveId + "/root/search(q='" + q + "')?$select=name,webUrl&$top=200";
  for (let g = 0; url && g < 25; g++) {
    const d = await gget(token, url);
    (d.value || []).forEach(function (it) {
      if (dbg) dbg.hits++;
      const name = it.name || "";
      if (!keep(name)) return;
      const dept = deptFromUrl(it.webUrl);
      if (!dept) { if (dbg) { dbg.scopedOut++; if (dbg.sampleScoped.length < 6) dbg.sampleScoped.push(it.webUrl || name); } return; }
      const key = (it.webUrl || name).toLowerCase();
      if (seen[key]) return; seen[key] = 1;
      if (dbg) { dbg.matched++; if (dbg.sampleMatched.length < 10) dbg.sampleMatched.push(dept + " / " + name); }
      out.push(make(it, dept));
    });
    url = d["@odata.nextLink"] || "";
  }
  return out;
}
// Department = the folder segment right after BASE_PATH in the file's webUrl.
function deptFromUrl(u) {
  const s = safeDecode(u || "");
  const i = s.toLowerCase().indexOf(BASE_PATH.toLowerCase());
  if (i < 0) return null;                                   // not under Assets & Templates
  const rest = s.slice(i + BASE_PATH.length).replace(/^\/+/, "");
  const seg = (rest.split("/")[0] || "").trim();
  if (!seg || /\.[a-z0-9]{2,5}$/i.test(seg)) return null;   // file sat directly in the base folder
  return seg;
}

function group(items) {
  const byCat = {};
  items.forEach(function (it) {
    const cat = catName(it.dept);
    (byCat[cat] = byCat[cat] || []).push({ title: titleFrom(it.name), url: it.url });
  });
  return Object.keys(byCat)
    .sort(function (a, b) { return ord(a) - ord(b) || a.localeCompare(b); })
    .map(function (c) {
      const list = byCat[c].sort(function (x, y) { return x.title.localeCompare(y.title); });
      return { category: c, items: list };
    });
}
function ord(cat) { var i = ORDER.indexOf(cat); return i < 0 ? 999 : i; }

function deptFrom(path) {
  const p = safeDecode(path);
  const i = p.toLowerCase().indexOf(BASE_PATH.toLowerCase());
  if (i < 0) return "";
  const rest = p.slice(i + BASE_PATH.length).replace(/^\/+/, "");
  return (rest.split("/")[0] || "").trim();
}
function catName(seg) {
  if (!seg) return "Other";
  const m = DEPT_MAP[seg.toLowerCase()];
  if (m) return m;
  const cleaned = seg.replace(/^\s*\d+\s*[.)\-]*\s*/, "").trim();
  return cleaned || seg;
}
function titleFrom(name) {
  let t = name.replace(/\.pdf$/i, "");
  t = t.replace(/flow\s*chart/ig, " ")
    .replace(/\bfor\b/ig, " ").replace(/\bof\b/ig, " ")
    .replace(/\(\s*v?\.?\s*\d+\s*\)/ig, " ").replace(/\bv\.?\s*\d+\b/ig, " ")
    .replace(/[-–_]+/g, " ").replace(/\s+/g, " ").trim();
  if (!t) t = name.replace(/\.pdf$/i, "").trim();
  return t.charAt(0).toUpperCase() + t.slice(1);
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
async function appToken() {
  const body = new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const r = await fetch("https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: body });
  if (!r.ok) throw new Error("token " + r.status);
  return (await r.json()).access_token;
}
async function gget(token, url) {
  for (let attempt = 0; ; attempt++) {
    const r = await fetch(url, { headers: { Authorization: "Bearer " + token } });
    if (r.status === 429 && attempt < 4) {
      const ra = parseInt(r.headers.get("Retry-After") || "2", 10);
      await sleep((ra > 0 ? ra : 2) * 1000);
      continue;
    }
    if (!r.ok) throw new Error("graph " + r.status);
    return r.json();
  }
}
function sleep(ms) { return new Promise(function (res) { setTimeout(res, ms); }); }
function lowerKeys(o) { const out = {}; Object.keys(o || {}).forEach(function (k) { out[k.toLowerCase()] = o[k]; }); return out; }
function safeDecode(s) { try { return decodeURIComponent(String(s || "")); } catch (e) { return String(s || ""); } }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
