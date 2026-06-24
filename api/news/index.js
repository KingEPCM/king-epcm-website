/*
 * Company News API (Azure Static Web Apps managed function)
 *
 *   GET    /api/news   → list all posts (any signed-in user)              [newest first]
 *   POST   /api/news   → create (no id) or update (with id) a post        [admin role only]
 *   DELETE /api/news?id=…  → delete a post                                [admin role only]
 *
 * Posts are stored in Azure Table Storage. "Admin" = the signed-in user has the
 * Static Web Apps "admin" role (assigned via SWA role management). The role is read
 * from the platform-provided x-ms-client-principal header, so it can't be spoofed
 * from the browser.
 *
 * App setting required: NEWS_STORAGE_CONNECTION = an Azure Storage account
 * connection string. If unset, GET returns 501 and the page falls back to news.json.
 */
const { TableClient } = require("@azure/data-tables");

const CONN = process.env.NEWS_STORAGE_CONNECTION;
const TABLE = "news";
const PK = "news";

module.exports = async function (context, req) {
  const method = (req.method || "GET").toUpperCase();

  if (!CONN) { context.res = json(501, { error: "News storage not configured" }); return; }
  const client = TableClient.fromConnectionString(CONN, TABLE);
  try { await client.createTable(); } catch (e) { /* table already exists */ }

  try {
    if (method === "GET") {
      const items = [];
      for await (const e of client.listEntities()) {
        items.push({ id: e.rowKey, date: e.dateText || "", title: e.title || "", body: e.body || "", createdAt: e.createdAt || "", pinned: e.pinned === true });
      }
      // Pinned first, then newest first.
      items.sort(function (a, b) {
        if (!!a.pinned !== !!b.pinned) return a.pinned ? -1 : 1;
        return String(b.createdAt).localeCompare(String(a.createdAt));
      });
      context.res = json(200, items);
      return;
    }

    // Everything below changes data → admins only.
    const p = principal(req);
    if (!isAdmin(p)) { context.res = json(403, { error: "Admins only" }); return; }

    if (method === "POST") {
      const b = req.body || {};
      const title = String(b.title || "").trim();
      const body = String(b.body || "").trim();
      if (!title || !body) { context.res = json(400, { error: "Title and message are required" }); return; }
      const dateText = String(b.date || "").trim() || formatToday();
      const pinned = b.pinned === true || b.pinned === "true";
      const now = new Date().toISOString();
      let id = String(b.id || "").trim();
      if (id) {
        const existing = await client.getEntity(PK, id).catch(function () { return null; });
        const createdAt = (existing && existing.createdAt) || now;
        await client.upsertEntity({ partitionKey: PK, rowKey: id, title: title, body: body, dateText: dateText, pinned: pinned, createdAt: createdAt, updatedAt: now, author: emailOf(p) }, "Replace");
      } else {
        id = String(Date.now()) + String(Math.floor(Math.random() * 1000));
        await client.createEntity({ partitionKey: PK, rowKey: id, title: title, body: body, dateText: dateText, pinned: pinned, createdAt: now, author: emailOf(p) });
      }
      context.res = json(200, { ok: true, id: id });
      return;
    }

    if (method === "DELETE") {
      const id = String((req.query && req.query.id) || (req.body && req.body.id) || "").trim();
      if (!id) { context.res = json(400, { error: "id required" }); return; }
      await client.deleteEntity(PK, id).catch(function () {});
      context.res = json(200, { ok: true });
      return;
    }

    context.res = json(405, { error: "Method not allowed" });
  } catch (err) {
    context.log.error(err);
    context.res = json(502, { error: "News storage error" });
  }
};

function principal(req) {
  try {
    const h = req.headers["x-ms-client-principal"];
    if (!h) return null;
    return JSON.parse(Buffer.from(h, "base64").toString("utf8"));
  } catch (e) { return null; }
}
function isAdmin(p) {
  return !!(p && Array.isArray(p.userRoles) && p.userRoles.indexOf("admin") > -1);
}
function emailOf(p) {
  if (!p) return "";
  if (p.userDetails && String(p.userDetails).includes("@")) return p.userDetails;
  const c = (p.claims || []).find(function (x) { return /email|preferred_username|upn/i.test(x.typ || "") && String(x.val).includes("@"); });
  return c ? c.val : "";
}
function formatToday() {
  return new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}
function json(status, body, extra) {
  return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body };
}
