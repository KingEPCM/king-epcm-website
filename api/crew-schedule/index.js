/*
 * GET /api/crew-schedule?start=<ISO>&end=<ISO>
 *   → { configured:true,
 *       shifts:[ {id,userId,name,title,start,end,allDay,notes,location} ],
 *       off:   [ {id,userId,name,date,seconds} ] }
 *
 * Reads the QuickBooks Time (TSheets) REST API:
 *   - schedule_events            → field-crew shifts
 *   - time_off_request_entries   → who's off (PTO / time off)
 *   - users / supplemental_data  → maps user_id → person's name
 *
 * App settings (Azure portal → Configuration):
 *   QBT_ACCESS_TOKEN            (required) Long-lived token from QuickBooks Time:
 *                               Feature Add-ons → API → "Add a new application", then extend the
 *                               token's expiry in the web UI so it doesn't need refreshing.
 *   QBT_SCHEDULE_CALENDAR_IDS   (optional) CSV of schedule-calendar ids to include.
 *                               Default: all active schedule calendars.
 *   QBT_BASE                    (optional) API base. Default https://rest.tsheets.com/api/v1
 *
 * No Microsoft Graph / Azure permissions are needed. Degrades gracefully:
 * returns 501 when not configured; per-section failures log and return empty lists.
 */
const zlib = require("zlib");
let TableClient = null;
try { TableClient = require("@azure/data-tables").TableClient; } catch (e) { /* durable cache optional */ }

const BASE = (process.env.QBT_BASE || "https://rest.tsheets.com/api/v1").replace(/\/+$/, "");
const TOKEN = process.env.QBT_ACCESS_TOKEN;
const CAL_IDS = (process.env.QBT_SCHEDULE_CALENDAR_IDS || "").trim();

// Per-request timeout so a slow/hung QuickBooks call fails fast instead of hanging the page.
const FETCH_TIMEOUT = +(process.env.QBT_FETCH_TIMEOUT_MS || 12000);

// In-memory caches shared across warm invocations. The all-history time-off fetch is the
// slow part (the endpoint has no date filter), so we cache it and just re-filter the cached
// entries to the requested window on each call.
const TOF_TTL = +(process.env.QBT_TIMEOFF_TTL_MS || 600000); // 10 min
const CAL_TTL = 600000; // 10 min — schedule calendars rarely change
let _tof = null, _tofTs = 0;
let _calIds = null, _calTs = 0;
let _jc = null, _jcTs = 0; const JC_TTL = 600000; // cached jobcode id -> name (time-off types)

// Durable cache for the time-off bundle (survives cold starts, unlike the in-memory cache above).
// Stored gzip+base64 in Table Storage, chunked across properties. Entirely optional & best-effort.
const CACHE_CONN = process.env.QBT_CACHE_CONNECTION || process.env.NEWS_STORAGE_CONNECTION || process.env.AzureWebJobsStorage || "";
const CACHE_TABLE = "qbtcache";

module.exports = async function (context, req) {
  if (!TOKEN) { context.res = json(501, { configured: false, error: "QuickBooks Time not configured" }); return; }
  if (!principalEmail(req)) { context.res = json(401, { configured: true, error: "Not signed in" }); return; }

  let start, end;
  try { start = req.query && req.query.start ? new Date(req.query.start) : startOfToday(); } catch (e) { start = startOfToday(); }
  try { end = req.query && req.query.end ? new Date(req.query.end) : new Date(start.getTime() + 14 * 86400000); } catch (e) { end = new Date(start.getTime() + 14 * 86400000); }
  if (isNaN(start)) start = startOfToday();
  if (isNaN(end)) end = new Date(start.getTime() + 14 * 86400000);

  const warnings = [];
  const userMap = {};     // user_id -> display/preferred name, from supplemental_data
  const userActive = {};  // user_id -> false if the staff member is archived / has left
  const jobcodeName = {}; // jobcode_id -> name (the "type" of time off, e.g. Vacation / Sick)
  function absorbUsers(supp) {
    if (supp && supp.users) Object.keys(supp.users).forEach(function (id) {
      const u = supp.users[id]; userMap[id] = nameOf(u); userActive[id] = (u.active !== false);
    });
  }
  function absorbJobcodes(supp) {
    if (supp && supp.jobcodes) Object.keys(supp.jobcodes).forEach(function (id) { jobcodeName[id] = supp.jobcodes[id].name || ""; });
  }

  // ---- crew shifts and who's-off load in parallel; each is guarded independently ----
  let shifts = [], off = [];
  async function loadShifts() {
  try {
    let calIds = CAL_IDS;
    if (!calIds) {
      if (_calIds != null && (Date.now() - _calTs) < CAL_TTL) { calIds = _calIds; }
      else {
        const cals = await qbAll("schedule_calendars", {}, "schedule_calendars", absorbUsers, context);
        calIds = cals.map(function (c) { return c.id; }).join(","); _calIds = calIds; _calTs = Date.now();
      }
    }
    if (calIds) {
      // schedule_events defaults to active, published (non-draft) events.
      const evs = await qbAll("schedule_events", {
        schedule_calendar_ids: calIds,
        start: isoNoMs(start),
        end: isoNoMs(end)
      }, "schedule_events", absorbUsers, context);
      shifts = evs.map(function (e) {
        const uid = String(e.user_id || (Array.isArray(e.assigned_user_ids) && e.assigned_user_ids[0]) || "");
        return { id: e.id, userId: uid, name: userMap[uid] || "", title: e.title || "", start: e.start, end: e.end, allDay: !!e.all_day, notes: e.notes || "", location: e.location || "" };
      });
    } else {
      warnings.push("shifts: no schedule calendars found");
    }
  } catch (e) { const m = String(e && e.message || e); context.log.error("crew-schedule shifts: " + m); warnings.push("shifts: " + m); }
  }

  // ---- who's off — approved time off in the window ----
  // time_off_requests has no date filter, so fetch requests and keep the entries
  // (from supplemental_data) whose date lands in our window and whose request is approved.
  async function loadOff() {
  try {
    var tof;
    if (_tof && (Date.now() - _tofTs) < TOF_TTL) {
      tof = _tof; // reuse the in-memory cache (warm instance)
    } else {
      tof = await durableReadTof(); // survives cold starts — a single fast read instead of the slow all-history fetch
      if (tof) { _tof = tof; _tofTs = Date.now(); }
      else {
        const supEntries = {}, reqStatus = {}, users = {}, jobcodes = {};
        function absorbTOR(supp) {
          if (supp && supp.users) Object.keys(supp.users).forEach(function (id) { var u = supp.users[id]; users[id] = { name: nameOf(u), active: u.active !== false }; });
          if (supp && supp.jobcodes) Object.keys(supp.jobcodes).forEach(function (id) { jobcodes[id] = supp.jobcodes[id].name || ""; });
          if (supp && supp.time_off_request_entries) Object.keys(supp.time_off_request_entries).forEach(function (id) { supEntries[id] = supp.time_off_request_entries[id]; });
        }
        const reqs = await qbAll("time_off_requests", {}, "time_off_requests", absorbTOR, context);
        reqs.forEach(function (r) { reqStatus[String(r.id)] = String(r.status || ""); });
        tof = { supEntries: supEntries, reqStatus: reqStatus, users: users, jobcodes: jobcodes };
        _tof = tof; _tofTs = Date.now();
        await durableWriteTof(tof); // populate the durable cache for the next cold start
      }
    }
    // seed the shared maps from the (possibly cached) time-off data
    Object.keys(tof.users).forEach(function (id) { if (userMap[id] === undefined) userMap[id] = tof.users[id].name; if (userActive[id] === undefined) userActive[id] = tof.users[id].active; });
    Object.keys(tof.jobcodes).forEach(function (id) { if (!jobcodeName[id]) jobcodeName[id] = tof.jobcodes[id]; });
    const lo = ymd(start), hi = ymd(end);
    Object.keys(tof.supEntries).forEach(function (id) {
      const t = tof.supEntries[id];
      if (!t || t.active === false || !t.date) return;
      if (t.date < lo || t.date > hi) return; // window guard (YYYY-MM-DD compares lexically)
      const st = (t.time_off_request_id != null && tof.reqStatus[String(t.time_off_request_id)]) || t.status || "";
      if (st && !/approv/i.test(st)) return; // keep approved (or unknown) only
      const uid = String(t.user_id || ""), jid = String(t.jobcode_id || "");
      off.push({ id: t.id, userId: uid, name: userMap[uid] || "", date: t.date, seconds: t.seconds || t.duration || 0, type: jobcodeName[jid] || "", jobcodeId: jid });
    });
  } catch (e) { const m = String(e && e.message || e); context.log.error("crew-schedule time-off: " + m); warnings.push("timeoff: " + m); }
  }
  await Promise.all([loadShifts(), loadOff()]);

  // ---- resolve any names not already provided via supplemental_data ----
  const missing = uniq([].concat(shifts, off).map(function (x) { return x.userId; }).filter(function (id) { return id && !userMap[id]; }));
  if (missing.length) {
    try {
      const us = await qbAll("users", { ids: missing.join(","), active: "both" }, "users", function () {}, context);
      us.forEach(function (u) { userMap[String(u.id)] = nameOf(u); userActive[String(u.id)] = (u.active !== false); });
    } catch (e) { context.log.error("crew-schedule users: " + (e && e.message || e)); }
  }

  // ---- resolve time-off types: if any entry's jobcode name is still unknown, pull the full
  //      jobcode list (cached) and map id -> name (covers codes not in supplemental_data) ----
  if (off.some(function (o) { return o.jobcodeId && !jobcodeName[o.jobcodeId]; })) {
    const all = await allJobcodes(context);
    Object.keys(all).forEach(function (id) { if (!jobcodeName[id]) jobcodeName[id] = all[id]; });
  }

  // Use the QuickBooks Time name (display/preferred name).
  shifts.forEach(function (s) { s.name = userMap[s.userId] || s.name || (s.userId ? "User " + s.userId : "Unassigned"); });
  off.forEach(function (o) { o.name = userMap[o.userId] || o.name || (o.userId ? "User " + o.userId : ""); o.type = jobcodeName[o.jobcodeId] || o.type || ""; delete o.jobcodeId; });

  // Exclude staff who have left (archived / inactive in QuickBooks Time).
  shifts = shifts.filter(function (s) { return userActive[s.userId] !== false; });
  off = off.filter(function (o) { return userActive[o.userId] !== false; });

  context.res = json(200, { configured: true, shifts: shifts, off: off, warnings: warnings }, { "Cache-Control": "no-store" });
};

// Full jobcode id -> name map (cached). Time-off "types" are jobcode names. Note: QuickBooks
// Time defaults /jobcodes to type=regular and OMITS pto codes, so we must pass type=all.
async function allJobcodes(context) {
  if (_jc && Date.now() - _jcTs < JC_TTL) return _jc;
  const map = {};
  try {
    const jcs = await qbAll("jobcodes", { type: "all", active: "both" }, "jobcodes", function () {}, context);
    jcs.forEach(function (j) { map[String(j.id)] = j.name || ""; });
    _jc = map; _jcTs = Date.now();
  } catch (e) { if (context) context.log.error("crew-schedule jobcodes: " + (e && e.message || e)); }
  return _jc || map;
}

// A single fetch with a hard timeout (so the request can't hang indefinitely).
async function qbFetch(url) {
  const ac = new AbortController();
  const t = setTimeout(function () { ac.abort(); }, FETCH_TIMEOUT);
  try {
    return await fetch(url, { headers: { Authorization: "Bearer " + TOKEN }, signal: ac.signal });
  } catch (e) {
    if (e && (e.name === "AbortError" || /abort/i.test(String(e.message || e)))) throw new Error("timed out after " + FETCH_TIMEOUT + "ms");
    throw e;
  } finally { clearTimeout(t); }
}

// Fetch every page of a QuickBooks Time list endpoint. Returns an array of objects.
async function qbAll(endpoint, params, resultsKey, onSupp, context) {
  const out = []; let page = 1, guard = 0;
  while (guard++ < 25) {
    const qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]); }).join("&");
    const url = BASE + "/" + endpoint + "?" + qs + "&page=" + page;
    const r = await qbFetch(url);
    if (!r.ok) { const t = await safeText(r); throw new Error(endpoint + " " + r.status + " " + t.slice(0, 200)); }
    const j = await r.json();
    if (onSupp) onSupp(j.supplemental_data);
    const res = (j.results && j.results[resultsKey]) || {};
    Object.keys(res).forEach(function (id) { out.push(res[id]); });
    if (!j.more) break;
    page++;
  }
  return out;
}

/* ---------- Durable cache (Table Storage, gzip + chunked) for the time-off bundle ---------- */
function cacheClient() {
  if (!TableClient || !CACHE_CONN) return null;
  try { return TableClient.fromConnectionString(CACHE_CONN, CACHE_TABLE); } catch (e) { return null; }
}
async function durableReadTof() {
  const c = cacheClient(); if (!c) return null;
  try {
    const e = await c.getEntity("timeoff", "v1");
    if (!e || !e.ts || (Date.now() - Number(e.ts)) > TOF_TTL) return null;
    let b64 = ""; const n = Number(e.chunks || 0);
    for (let i = 0; i < n; i++) b64 += e["c" + i] || "";
    if (!b64) return null;
    return JSON.parse(zlib.gunzipSync(Buffer.from(b64, "base64")).toString("utf8"));
  } catch (e) { return null; }
}
async function durableWriteTof(tof) {
  const c = cacheClient(); if (!c) return;
  try { await c.createTable(); } catch (e) { /* exists */ }
  try {
    const b64 = zlib.gzipSync(Buffer.from(JSON.stringify(tof), "utf8")).toString("base64");
    const CH = 30000, ent = { partitionKey: "timeoff", rowKey: "v1", ts: Date.now() };
    let n = 0;
    for (let p = 0; p < b64.length; p += CH) { ent["c" + n] = b64.slice(p, p + CH); n++; if (n > 30) return; } // too big — skip (stay on in-memory)
    ent.chunks = n;
    await c.upsertEntity(ent, "Replace");
  } catch (e) { /* best-effort */ }
}

function nameOf(u) {
  if (!u) return "";
  const dn = (u.display_name || "").trim();
  if (dn) return dn;
  const n = ((u.first_name || "") + " " + (u.last_name || "")).trim();
  return n || (u.username || "");
}
function safeText(r) { return r.text().then(function (t) { return t; }).catch(function () { return ""; }); }
function uniq(a) { const s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
function startOfToday() { const d = new Date(); d.setHours(0, 0, 0, 0); return d; }
function ymd(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
// QuickBooks Time wants ISO-8601 without milliseconds, e.g. 2026-06-24T00:00:00+00:00
function isoNoMs(d) { return d.toISOString().replace(/\.\d{3}Z$/, "+00:00"); }
function principalEmail(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return "";
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    if (p.userDetails && p.userDetails.indexOf("@") > -1) return p.userDetails;
    const c = (p.claims || []).find(function (x) { return ["email", "preferred_username", "upn"].indexOf((x.typ || "").toLowerCase()) > -1 && x.val; });
    return c ? c.val : (p.userId || "");
  } catch (e) { return ""; }
}
function json(status, body, extra) { return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body }; }
