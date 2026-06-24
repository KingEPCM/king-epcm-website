/*
 * GET /api/my-reimbursements  → { claims: [ {type,total,summary,folderUrl,date,createdAt,itemCount} ] }
 * The signed-in person's own submitted reimbursement claims (newest first), so they can
 * check what they've already sent and avoid duplicates. Reads the "reimbursements" table
 * written by /api/reimbursement-submit.
 */
const { TableClient } = require("@azure/data-tables");
const CONN = process.env.NEWS_STORAGE_CONNECTION;

module.exports = async function (context, req) {
  const email = principalEmail(req);
  if (!email) { context.res = json(401, { error: "Not signed in" }); return; }
  if (!CONN) { context.res = json(200, { claims: [] }, { "Cache-Control": "no-store" }); return; }
  try {
    const t = TableClient.fromConnectionString(CONN, "reimbursements");
    const pk = keyOf(email);
    const out = [];
    const iter = t.listEntities({ queryOptions: { filter: "PartitionKey eq '" + pk.replace(/'/g, "''") + "'" } });
    for await (const e of iter) {
      out.push({ type: e.type || "", total: e.total || "", summary: e.summary || "", folderUrl: e.folderUrl || "", date: e.date || "", createdAt: e.createdAt || "", itemCount: e.itemCount || 0 });
    }
    out.sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
    context.res = json(200, { claims: out.slice(0, 30) }, { "Cache-Control": "no-store" });
  } catch (e) {
    context.log.error(e);
    context.res = json(200, { claims: [] }, { "Cache-Control": "no-store" });
  }
};

function principalEmail(req) {
  try {
    const h = req.headers["x-ms-client-principal"]; if (!h) return "";
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    if (p.userDetails && p.userDetails.indexOf("@") > -1) return p.userDetails;
    const c = (p.claims || []).find(function (x) { return ["email", "preferred_username", "upn"].indexOf((x.typ || "").toLowerCase()) > -1 && x.val; });
    return c ? c.val : "";
  } catch (e) { return ""; }
}
function keyOf(email) { return String(email || "").toLowerCase().replace(/[\\/#?\t\n\r]/g, "_"); }
function json(status, body, extra) { return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body }; }
