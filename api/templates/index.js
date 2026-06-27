/*
 * GET /api/templates — a LIVE list of document templates, read from SharePoint.
 * Searches the document library under BASE_PATH for files whose name contains
 * "template" (Word / Excel / PowerPoint / PDF), groups them by their top-level
 * department folder, and returns:
 *   { configured, categories: [ { category, items: [ { title, url, note } ] } ] }
 * The front end keeps the curated intranet generator links and falls back to the
 * static resources.json templates if this isn't configured.
 *
 * App settings (all optional — sensible defaults for King EPCM):
 *   AAD_CLIENT_ID, AAD_CLIENT_SECRET, GRAPH_TENANT      (Graph app-only)
 *   FLOW_SITE_PATH   Graph site path (default kingepcm.sharepoint.com root site)
 *   FLOW_LIBRARY     drive/library name (default: the site's default documents library)
 *   FLOW_BASE_PATH   folder to scope to (default "03. Assets & Templates")
 *   FLOW_DEPT_MAP    JSON map of raw dept folder → friendly category name
 * Graph APP permission: Sites.Read.All (or the Sites.ReadWrite.All already granted).
 */
const zlib = require("zlib");
let TableClient = null;
try { TableClient = require("@azure/data-tables").TableClient; } catch (e) { /* durable cache optional */ }

const TENANT = process.env.GRAPH_TENANT || "52e591ce-74f5-4c10-9dc3-b1020c47dc24";
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SITE_PATH = process.env.FLOW_SITE_PATH || "kingepcm.sharepoint.com";
const LIBRARY = process.env.FLOW_LIBRARY || "";
const BASE_PATH = process.env.FLOW_BASE_PATH || "03. Assets & Templates";

// Per-request timeout so a slow Microsoft Graph call fails fast (falls back to cache / curated list).
const FETCH_TIMEOUT = +(process.env.GRAPH_FETCH_TIMEOUT_MS || 12000);
// Durable cache (survives cold starts). Reuses the news storage / functions storage connection.
const CACHE_CONN = process.env.QBT_CACHE_CONNECTION || process.env.NEWS_STORAGE_CONNECTION || process.env.AzureWebJobsStorage || "";
const CACHE_TABLE = "qbtcache"; // shared small cache table; row keyed by partition below

const DEFAULT_DEPT_MAP = {
  "04. civil": "Civil & Municipal Design",
  "05. geohyd": "Geotechnical & Hydrogeology",
  "07. field": "Survey & Field Inspections",
  "02. natural heritage environments": "Natural Heritage Environmental",
  "09. indirect deliverables": "Office & Operations"
};
let DEPT_MAP = DEFAULT_DEPT_MAP;
try { DEPT_MAP = Object.assign({}, DEFAULT_DEPT_MAP, lowerKeys(JSON.parse(process.env.FLOW_DEPT_MAP || "{}"))); } catch (e) {}
const ORDER = Object.keys(DEFAULT_DEPT_MAP).map(function (k) { return DEFAULT_DEPT_MAP[k]; });

// File types we treat as templates.
const TYPE = { docx: "Word", doc: "Word", dotx: "Word", xlsx: "Excel", xls: "Excel", xltx: "Excel", pptx: "PowerPoint", ppt: "PowerPoint", potx: "PowerPoint", pdf: "PDF" };

let _driveId, _baseId, _cache = null, _cacheTs = 0;
const TTL = 30 * 60 * 1000;

module.exports = async function (context, req) {
  const debug = req && req.query && (req.query.debug === "1" || req.query.debug === "true");
  if (!CLIENT_ID || !CLIENT_SECRET) { context.res = json(200, { configured: false, categories: [], debug: debug ? { reason: "Graph client id/secret not set" } : undefined }); return; }
  try {
    if (!debug && _cache && Date.now() - _cacheTs < TTL) { context.res = json(200, { configured: true, categories: _cache }); return; }
    // Durable cache — a single fast read instead of the slow Graph search after a cold start.
    if (!debug) {
      const dc = await durableRead();
      if (dc && dc.length) { _cache = dc; _cacheTs = Date.now(); context.res = json(200, { configured: true, categories: dc }); return; }
    }
    const token = await appToken();
    const driveId = await resolveDrive(token);
    const dbg = { hits: 0, matched: 0, scopedOut: 0, sampleMatched: [], sampleScoped: [] };
    const items = await searchTemplates(token, driveId, dbg);
    const categories = group(items);
    _cache = categories; _cacheTs = Date.now();
    if (!debug && categories.length) await durableWrite(categories);
    context.res = json(200, {
      configured: true, categories: categories,
      debug: debug ? Object.assign({ driveId: driveId, basePath: BASE_PATH, via: _baseId ? "folder-walk" : "drive-search", baseFolderId: _baseId || "", categories: categories.length, items: items.length }, dbg) : undefined
    });
  } catch (e) {
    context.log.error(e);
    context.res = json(200, { configured: false, categories: _cache || [], error: String((e && e.message) || e) });
  }
};

// Templates live under "03. Assets & Templates". We ENUMERATE that folder tree directly rather than
// doing a drive-wide search for the (very common) word "template" — the drive-wide search returns a
// huge result set that can exceed the function timeout ("Backend call failure"), and folder-scoped
// SEARCH returns nothing on this SharePoint. Walking the tree is scoped, correct, and bounded.
// Falls back to the old drive-wide search only if the base folder can't be resolved.
function keepTemplate(name) {
  const ext = (name.split(".").pop() || "").toLowerCase();
  return !!TYPE[ext] && name.toLowerCase().indexOf("template") > -1;
}
function makeTemplate(it, dept) {
  const ext = ((it.name || "").split(".").pop() || "").toLowerCase();
  return { name: it.name, url: it.webUrl, note: TYPE[ext], dept: dept };
}
async function searchTemplates(token, driveId, dbg) {
  const baseId = await resolveBaseFolder(token, driveId);
  if (baseId) {
    const files = await walkFolder(token, driveId, baseId, dbg);
    const seen = {}, out = [];
    files.forEach(function (it) {
      const name = it.name || "";
      if (!keepTemplate(name)) return;
      const key = (it.webUrl || name).toLowerCase();
      if (seen[key]) return; seen[key] = 1;
      // Department comes from the walk path (the top-level folder), since the webUrl of Office files
      // (…/_layouts/15/Doc.aspx?…) doesn't include the folder path. Fall back to the webUrl if needed.
      const dept = it.__dept || deptFromUrl(it.webUrl) || "";
      if (dbg && dbg.matched != null) { dbg.matched++; if (dbg.sampleMatched.length < 10) dbg.sampleMatched.push(dept + " / " + name); }
      out.push(makeTemplate(it, dept));
    });
    return out;
  }
  // Fallback: drive-wide search (slower) if we couldn't resolve the folder.
  return searchAndScope(token, driveId, dbg, "template", keepTemplate, makeTemplate);
}

// Breadth-first walk of a folder's descendants, processing each level in PARALLEL batches so a large
// tree completes quickly (sequential walking was slow). Returns file items (name + webUrl). Bounded by
// a folder-request budget so it always finishes well within the function timeout.
async function walkFolder(token, driveId, rootId, dbg) {
  const out = []; let level = [{ id: rootId, dept: "" }]; let calls = 0, truncated = false;
  const MAX = 600, CONC = 10;
  while (level.length) {
    if (calls >= MAX) { truncated = true; break; }
    const next = [];
    for (let i = 0; i < level.length; i += CONC) {
      if (calls >= MAX) { truncated = true; break; }
      const batch = level.slice(i, i + CONC);
      calls += batch.length;
      const lists = await Promise.all(batch.map(function (node) {
        return listChildren(token, driveId, node.id).catch(function () { return []; });
      }));
      lists.forEach(function (children, idx) {
        const parentDept = batch[idx].dept;
        children.forEach(function (it) {
          if (dbg) dbg.hits++;
          if (it.folder) {
            if (it.id) next.push({ id: it.id, dept: parentDept || it.name }); // top-level folder name = department
          } else {
            it.__dept = parentDept; // file inherits its branch's department
            out.push(it);
          }
        });
      });
    }
    level = next;
  }
  if (dbg) { dbg.folderCalls = calls; dbg.truncated = truncated; }
  return out;
}
// List every child of a folder (following pagination).
async function listChildren(token, driveId, id) {
  const acc = [];
  let url = "https://graph.microsoft.com/v1.0/drives/" + driveId + "/items/" + id + "/children?$select=id,name,webUrl,folder,file&$top=200";
  while (url) {
    const d = await gget(token, url);
    (d.value || []).forEach(function (it) { acc.push(it); });
    url = d["@odata.nextLink"] || "";
  }
  return acc;
}

async function searchAndScope(token, driveId, dbg, q, keep, make) {
  const seen = {}, out = [];
  // Drive-wide search, then scope by webUrl. (Folder-scoped search via /items/{id}/search returns
  // nothing on this SharePoint — its search index doesn't honour that scope — so we search the drive.)
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
function deptFromUrl(u) {
  const s = safeDecode(u || "");
  const i = s.toLowerCase().indexOf(BASE_PATH.toLowerCase());
  if (i < 0) return null;
  const rest = s.slice(i + BASE_PATH.length).replace(/^\/+/, "");
  const seg = (rest.split("/")[0] || "").trim();
  if (!seg || /\.[a-z0-9]{2,5}$/i.test(seg)) return null;
  return seg;
}

function group(items) {
  const byCat = {};
  items.forEach(function (it) {
    const cat = catName(it.dept);
    (byCat[cat] = byCat[cat] || []).push({ title: titleFrom(it.name), url: it.url, note: it.note });
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
  let t = name.replace(/\.[a-z0-9]+$/i, "");
  t = t.replace(/template/ig, " ")
    .replace(/\bfor\b/ig, " ").replace(/\bof\b/ig, " ")
    .replace(/\(\s*v?\.?\s*\d+\s*\)/ig, " ").replace(/\bv\.?\s*\d+\b/ig, " ")
    .replace(/[-–_]+/g, " ").replace(/\s+/g, " ").trim();
  if (!t) t = name.replace(/\.[a-z0-9]+$/i, "").trim();
  return t.charAt(0).toUpperCase() + t.slice(1);
}

// Resolve the item id of the BASE_PATH folder. Best-effort: returns "" on failure (caller falls back).
async function resolveBaseFolder(token, driveId) {
  if (_baseId !== undefined) return _baseId;
  try {
    const enc = BASE_PATH.split("/").map(function (s) { return encodeURIComponent(s); }).join("/");
    const it = await gget(token, "https://graph.microsoft.com/v1.0/drives/" + driveId + "/root:/" + enc + "?$select=id");
    _baseId = (it && it.id) || "";
  } catch (e) { _baseId = ""; }
  return _baseId;
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
    const ac = new AbortController();
    const timer = setTimeout(function () { ac.abort(); }, FETCH_TIMEOUT);
    let r;
    try { r = await fetch(url, { headers: { Authorization: "Bearer " + token }, signal: ac.signal }); }
    catch (e) { if (e && (e.name === "AbortError" || /abort/i.test(String(e.message || e)))) throw new Error("graph timed out"); throw e; }
    finally { clearTimeout(timer); }
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

/* ---------- Durable cache (Table Storage, gzip + chunked) for the grouped categories ---------- */
function cacheClient() {
  if (!TableClient || !CACHE_CONN) return null;
  try { return TableClient.fromConnectionString(CACHE_CONN, CACHE_TABLE); } catch (e) { return null; }
}
async function durableRead() {
  const c = cacheClient(); if (!c) return null;
  try {
    const e = await c.getEntity("templates", "v1");
    if (!e || !e.ts || (Date.now() - Number(e.ts)) > TTL) return null;
    let b64 = ""; const n = Number(e.chunks || 0);
    for (let i = 0; i < n; i++) b64 += e["c" + i] || "";
    if (!b64) return null;
    return JSON.parse(zlib.gunzipSync(Buffer.from(b64, "base64")).toString("utf8"));
  } catch (e) { return null; }
}
async function durableWrite(categories) {
  const c = cacheClient(); if (!c) return;
  try { await c.createTable(); } catch (e) { /* exists */ }
  try {
    const b64 = zlib.gzipSync(Buffer.from(JSON.stringify(categories), "utf8")).toString("base64");
    const CH = 30000, ent = { partitionKey: "templates", rowKey: "v1", ts: Date.now() };
    let n = 0;
    for (let p = 0; p < b64.length; p += CH) { ent["c" + n] = b64.slice(p, p + CH); n++; if (n > 30) return; }
    ent.chunks = n;
    await c.upsertEntity(ent, "Replace");
  } catch (e) { /* best-effort */ }
}
function lowerKeys(o) { const out = {}; Object.keys(o || {}).forEach(function (k) { out[k.toLowerCase()] = o[k]; }); return out; }
function safeDecode(s) { try { return decodeURIComponent(String(s || "")); } catch (e) { return String(s || ""); } }
function json(status, body) { return { status: status, headers: { "Content-Type": "application/json" }, body: body }; }
