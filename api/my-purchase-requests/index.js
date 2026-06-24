/*
 * GET /api/my-purchase-requests  → { requests: [ {project,vendor,total,summary,comments,date,createdAt,itemCount,attachmentCount} ] }
 * The signed-in person's own submitted purchase requests (newest first), so they can
 * check what they've already requested and avoid duplicates. Reads the
 * "purchaserequests" table written by /api/purchase-request.
 */
const { TableClient } = require("@azure/data-tables");
const CONN = process.env.NEWS_STORAGE_CONNECTION;

module.exports = async function (context, req) {
  const email = principalEmail(req);
  if (!email) { context.res = json(401, { error: "Not signed in" }); return; }
  if (!CONN) { context.res = json(200, { requests: [] }, { "Cache-Control": "no-store" }); return; }
  try {
    const t = TableClient.fromConnectionString(CONN, "purchaserequests");
    const pk = keyOf(email);
    const out = [];
    const iter = t.listEntities({ queryOptions: { filter: "PartitionKey eq '" + pk.replace(/'/g, "''") + "'" } });
    for await (const e of iter) {
      out.push({ project: e.project || "", vendor: e.vendor || "", total: e.total || "", summary: e.summary || "", comments: e.comments || "", date: e.date || "", createdAt: e.createdAt || "", itemCount: e.itemCount || 0, attachmentCount: e.attachmentCount || 0 });
    }
    out.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
    context.res = json(200, { requests: out.slice(0, 30) }, { "Cache-Control": "no-store" });
  } catch (e) {
    context.log.error(e);
    context.res = json(200, { requests: [] }, { "Cache-Control": "no-store" });
  }
};

function principalEmail(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return "";
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    if (p.userDetails && p.userDetails.indexOf("@") > -1) return p.userDetails;
    const c = (p.claims || []).find(function (x) { return ["email", "preferred_username", "upn", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress"].indexOf((x.typ || "").toLowerCase()) > -1 && x.val; });
    return c ? c.val : "";
  } catch (e) { return ""; }
}
function keyOf(email) { return String(email || "").toLowerCase().replace(/[\\/#?\t\n\r]/g, "_"); }
function json(status, body, extra) { return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body }; }
