/*
 * GET /api/my-tasks — the signed-in staff member's FULL open task list from Teamwork
 * (for the "My Work" page). Unlike /api/teamwork-projects (which is limited to the
 * dashboard's due-soon preview), this returns every incomplete task assigned to the
 * person, dated or not, sorted overdue → due soon → later → no date.
 *
 * Response: { configured, person, tasks: [ { name, project, due, overdue, dueSoon, priority, url } ] }
 *   - configured:false → UI shows sample data / a friendly note.
 *
 * Identifies the user from the Azure SWA "x-ms-client-principal" header (their M365
 * email) and matches the Teamwork person with that email (Teamwork email == M365 email).
 * App settings: TEAMWORK_DOMAIN (e.g. "kingepcm"), TEAMWORK_API_KEY.
 */
const DUE_SOON_DAYS = 14;

module.exports = async function (context, req) {
  const domain = process.env.TEAMWORK_DOMAIN;
  const apiKey = process.env.TEAMWORK_API_KEY;
  if (!domain || !apiKey) { context.res = json(200, { configured: false, tasks: [] }); return; }

  const email = userEmail(req);
  if (!email) { context.res = json(200, { configured: false, tasks: [] }); return; }

  const base = "https://" + domain + ".teamwork.com";
  const headers = { Authorization: "Basic " + Buffer.from(apiKey + ":x").toString("base64"), Accept: "application/json" };

  try {
    const pRes = await fetch(base + "/projects/api/v3/people.json?searchTerm=" + encodeURIComponent(email) + "&pageSize=50", { headers });
    if (!pRes.ok) throw new Error("people lookup " + pRes.status);
    const pData = await pRes.json();
    const people = pData.people || [];
    const me = people.find(function (p) { return (p.email || "").toLowerCase() === email.toLowerCase(); }) || people[0];
    if (!me) { context.res = json(200, { configured: true, person: null, tasks: [] }); return; }

    const tasks = await getOpenTasks(base, headers, me.id);
    context.res = json(200, { configured: true, person: { name: (me.firstName || "") + (me.lastName ? " " + me.lastName : "") || me.email, id: me.id }, tasks: tasks }, { "Cache-Control": "private, max-age=300" });
  } catch (err) {
    context.log.error(err);
    context.res = json(200, { configured: false, tasks: [], error: "Could not reach Teamwork" });
  }
};

async function getOpenTasks(base, headers, personId) {
  const now = new Date();
  const soon = new Date(now.getTime() + DUE_SOON_DAYS * 86400000);
  const out = [];
  let page = 1;
  for (let guard = 0; guard < 12; guard++) {  // up to 12 pages (3000 tasks) — plenty
    const url = base + "/projects/api/v3/tasks.json?responsiblePartyIds=" + personId +
      "&includeCompletedTasks=false&includeOverdueTasks=true&include=projects,tasklists" +
      "&pageSize=250&page=" + page;
    const r = await fetch(url, { headers });
    if (!r.ok) throw new Error("tasks " + r.status);
    const d = await r.json();
    const raw = d.tasks || [];
    const projById = (d.included && d.included.projects) || {};
    const tlById = (d.included && d.included.tasklists) || {};
    raw.forEach(function (t) {
      const aids = t.assigneeUserIds || [];
      if (aids.length && aids.indexOf(personId) === -1) return;     // only this person's tasks
      if (String(t.status).toLowerCase() === "completed" || t.progress === 100) return;
      const due = t.dueDate || t.dueAt || null;
      const dueDate = due ? new Date(due) : null;
      out.push({
        name: t.name,
        project: projectName(t, projById, tlById),
        due: due,
        overdue: !!(dueDate && dueDate < now),
        dueSoon: !!(dueDate && dueDate >= now && dueDate <= soon),
        priority: (t.priority || "").toLowerCase(),
        url: (t.meta && t.meta.webLink) || (base + "/app/tasks/" + t.id)
      });
    });
    if (raw.length < 250) break;     // last page
    page++;
  }
  // Sort: overdue (soonest first) → dated upcoming → undated.
  out.sort(function (a, b) {
    const ad = a.due ? new Date(a.due).getTime() : Infinity;
    const bd = b.due ? new Date(b.due).getTime() : Infinity;
    return ad - bd;
  });
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
    const c = claims.find(function (x) { return /email|preferred_username|upn/i.test(x.typ || "") && String(x.val).includes("@"); });
    return c ? c.val : null;
  } catch (e) { return null; }
}
function json(status, body, extra) {
  return { status: status, headers: Object.assign({ "Content-Type": "application/json" }, extra || {}), body: body };
}
