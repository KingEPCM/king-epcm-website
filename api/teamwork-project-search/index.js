/*
 * GET /api/teamwork-project-search?q=<text>
 * Type-ahead search of ACTIVE Teamwork projects by name. At King EPCM the project
 * NAME is the site address, so this lets the generators/forms fill an address field
 * from a real project. Manual entry still works — this only offers suggestions.
 *
 * Response: { projects: [ { id, name, company } ] }   (name = address)
 * App settings: TEAMWORK_DOMAIN, TEAMWORK_API_KEY (already used by the other functions).
 * Degrades gracefully: returns an empty list (not an error) when not configured.
 */
module.exports = async function (context, req) {
  const domain = process.env.TEAMWORK_DOMAIN;
  const apiKey = process.env.TEAMWORK_API_KEY;
  const q = ((req.query && req.query.q) || "").trim();

  if (!domain || !apiKey || q.length < 2) { context.res = json(200, { projects: [] }); return; }

  const base = `https://${domain}.teamwork.com`;
  const headers = {
    Authorization: "Basic " + Buffer.from(apiKey + ":x").toString("base64"),
    Accept: "application/json"
  };

  try {
    const url = `${base}/projects/api/v3/projects.json?searchTerm=${encodeURIComponent(q)}` +
      `&status=active&pageSize=25&include=companies`;
    const r = await fetch(url, { headers });
    if (!r.ok) throw new Error("search " + r.status);
    const d = await r.json();

    // map sideloaded companies (v3 "included")
    const companies = (d.included && d.included.companies) || {};
    const ql = q.toLowerCase();
    const items = (d.projects || [])
      .filter(p => p && p.name && p.name.toLowerCase().includes(ql))
      .slice(0, 15)
      .map(p => ({
        id: p.id,
        name: p.name,
        company: companyName(p, companies)
      }));

    context.res = json(200, { projects: items }, { "Cache-Control": "private, max-age=120" });
  } catch (e) {
    context.log.error(e);
    context.res = json(502, { projects: [], error: "search failed" });
  }
};

function companyName(p, companies) {
  const cid = p.companyId || (p.company && p.company.id);
  if (cid && companies[cid] && companies[cid].name) return companies[cid].name;
  if (p.company && p.company.name) return p.company.name;
  return "";
}
function json(status, body, extra) {
  return { status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body };
}
