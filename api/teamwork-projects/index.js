/*
 * GET /api/teamwork-projects
 * Returns the SIGNED-IN staff member's Teamwork data:
 *   { projects: [...], tasks: [...] }
 *   - projects: their active projects (soonest due first)
 *   - tasks: their incomplete tasks that are OVERDUE or DUE SOON (next 14 days)
 *
 * Identifies the user from the Azure SWA "x-ms-client-principal" header (their M365
 * email), matches the Teamwork person with that email, and returns their items.
 * (Assumes Teamwork email == Microsoft 365 email — true at King EPCM.)
 *
 * App settings required: TEAMWORK_DOMAIN (e.g. "kingepcm"), TEAMWORK_API_KEY.
 * If not configured / no signed-in user, returns 501 so the UI shows sample data.
 */
const DUE_SOON_DAYS = 14;

module.exports = async function (context, req) {
  const domain = process.env.TEAMWORK_DOMAIN;
  const apiKey = process.env.TEAMWORK_API_KEY;
  if (!domain || !apiKey) { context.res = json(501, { error: "Teamwork not configured" }); return; }

  const email = userEmail(req);
  if (!email) { context.res = json(501, { error: "No signed-in user" }); return; }

  const base = `https://${domain}.teamwork.com`;
  const headers = {
    Authorization: "Basic " + Buffer.from(apiKey + ":x").toString("base64"),
    Accept: "application/json"
  };

  try {
    // 1) Resolve the Teamwork person by email.
    const pRes = await fetch(`${base}/projects/api/v3/people.json?searchTerm=${encodeURIComponent(email)}&pageSize=50`, { headers });
    if (!pRes.ok) throw new Error("people lookup " + pRes.status);
    const pData = await pRes.json();
    const people = pData.people || [];
    const me = people.find(p => (p.email || "").toLowerCase() === email.toLowerCase()) || people[0];
    if (!me) { context.res = json(200, { projects: [], tasks: [] }); return; }

    // 2) Their active projects + their due/overdue tasks (in parallel).
    const [projects, tasks] = await Promise.all([
      getProjects(base, headers, me.id),
      getDueTasks(base, headers, me.id)
    ]);

    context.res = json(200, { projects, tasks }, { "Cache-Control": "private, max-age=300" });
  } catch (err) {
    context.log.error(err);
    context.res = json(502, { error: "Could not reach Teamwork" });
  }
};

async function getProjects(base, headers, personId) {
  const r = await fetch(`${base}/people/${personId}/projects.json?status=active`, { headers });
  if (!r.ok) throw new Error("projects " + r.status);
  const d = await r.json();
  const raw = (d.projects || []).filter(p => {
    var st = String(p.status || p.subStatus || p["sub-status"] || "").toLowerCase();
    if (st === "completed") return false;
    if (p.completedAt || p.completedOn || p["completed-on"] || p.completed === true || p.completed === "true") return false;
    return true;
  });
  const out = raw.map(p => ({
    name: p.name,
    category: (p.category && p.category.name) || p["category-name"] || "",
    due: p.endDate || p["end-date"] || p.endDateTime || null,
    status: String(p.subStatus || p["sub-status"] || p.status || "").toLowerCase(),
    url: `${base}/app/projects/${p.id}`
  }));
  out.sort((a, b) => new Date(a.due || "2999") - new Date(b.due || "2999"));
  return out;
}

async function getDueTasks(base, headers, personId) {
  const url = `${base}/projects/api/v3/tasks.json?responsiblePartyIds=${personId}` +
    `&includeCompletedTasks=false&includeOverdueTasks=true&include=projects,tasklists&pageSize=250`;
  const r = await fetch(url, { headers });
  if (!r.ok) throw new Error("tasks " + r.status);
  const d = await r.json();
  const raw = d.tasks || [];
  const projById = (d.included && d.included.projects) || {};
  const tlById = (d.included && d.included.tasklists) || {};

  const now = new Date();
  const horizon = new Date(now.getTime() + DUE_SOON_DAYS * 86400000);

  const out = [];
  for (const t of raw) {
    const aids = t.assigneeUserIds || [];                 // safety: only this person's tasks
    if (aids.length && aids.indexOf(personId) === -1) continue;
    const due = t.dueDate || t.dueAt || null;
    if (!due) continue;                                   // only dated tasks
    if (String(t.status).toLowerCase() === "completed" || t.progress === 100) continue;
    const dueDate = new Date(due);
    if (dueDate > horizon) continue;                      // not yet relevant
    out.push({
      name: t.name,
      project: projectName(t, projById, tlById),
      due: due,
      overdue: dueDate < now,
      url: (t.meta && t.meta.webLink) || `${base}/app/tasks/${t.id}`
    });
  }
  out.sort((a, b) => new Date(a.due) - new Date(b.due));   // soonest/most overdue first
  return out;
}

function projectName(t, projById, tlById) {
  let pid = t.projectId || (t.project && t.project.id);
  if (!pid && t.tasklistId && tlById[t.tasklistId]) {
    const tl = tlById[t.tasklistId];
    pid = tl.projectId || (tl.project && tl.project.id);
  }
  const p = pid && (projById[pid] || projById[String(pid)]);
  return (p && p.name) || "";
}

function userEmail(req) {
  try {
    const h = req.headers["x-ms-client-principal"];
    if (!h) return null;
    const p = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    if (p.userDetails && p.userDetails.includes("@")) return p.userDetails;
    const claims = p.claims || [];
    const c = claims.find(x => /email|preferred_username|upn/i.test(x.typ || "") && String(x.val).includes("@"));
    return c ? c.val : null;
  } catch (e) { return null; }
}
function json(status, body, extra) {
  return { status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body };
}
