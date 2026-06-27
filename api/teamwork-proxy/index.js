/*
 * POST /api/teamwork-proxy — backend for the embeddable Teamwork forms
 * (forms/new_project_form.html, forms/add_tasklists_form.html) when they run
 * on the intranet instead of inside Cowork.
 *
 * Request:  { tool: "mcp__…__twprojects-<name>", args: {...} }
 * Response: { content: [ { text: "<RESPONSE>" } ], isError: bool }
 *   - list tools  → text is the raw Teamwork JSON ({people}/{projects}/{tasklists}/…)
 *   - create/clone/update/add tools → text contains the new record id ("… ID 12345")
 *
 * ADMIN ONLY. Creating Teamwork projects is privileged, so every call requires the
 * "admin" role (Azure SWA x-ms-client-principal). Reuses TEAMWORK_DOMAIN / TEAMWORK_API_KEY.
 *
 * NOTE: Teamwork's write endpoints (project copy-from-template, custom-field values,
 * milestones) vary by account/version. The mappings below follow Teamwork's documented
 * REST API; on any mismatch the proxy passes Teamwork's own error text back to the form
 * log so it can be pinpointed and adjusted. Read endpoints are straightforward.
 */
const DOMAIN = process.env.TEAMWORK_DOMAIN;
const API_KEY = process.env.TEAMWORK_API_KEY;

module.exports = async function (context, req) {
  const p = principal(req);
  const roles = (p && p.roles) || [];
  if (roles.indexOf("admin") === -1) { context.res = pack("Forbidden — this tool is for admins only.", true, 403); return; }
  if (!DOMAIN || !API_KEY) { context.res = pack("Teamwork is not configured (TEAMWORK_DOMAIN / TEAMWORK_API_KEY).", true); return; }

  const body = req.body || {};
  const tool = String(body.tool || "").replace(/^.*twprojects-/, "");   // strip MCP server prefix
  const args = body.args || {};
  if (!tool) { context.res = pack("No tool specified.", true); return; }

  try {
    const text = await handle(tool, args, context);
    context.res = pack(text, false);
  } catch (e) {
    context.log.error(e);
    context.res = pack(String((e && e.message) || e), true);
  }
};

const BASE = function () { return "https://" + DOMAIN + ".teamwork.com"; };
const V3 = function (path) { return BASE() + "/projects/api/v3/" + path; };

async function handle(tool, a, context) {
  switch (tool) {
    /* ---------- reads: return Teamwork's JSON verbatim (v3 keys match the forms) ---------- */
    case "list_users":
      return await raw("GET", V3("people.json?pageSize=" + num(a.page_size, 200)));
    case "list_project_categories":
      return await raw("GET", V3("projectcategories.json?pageSize=200"));
    case "list_tasklists":
      return await raw("GET", V3("tasklists.json?pageSize=" + num(a.page_size, 500) +
        (a.projectId || a.project_id ? "&projectIds=" + (a.projectId || a.project_id) : "")));
    case "list_companies":
      return await raw("GET", V3("companies.json?pageSize=200" + (a.search_term ? "&searchTerm=" + encodeURIComponent(a.search_term) : "")));
    case "list_projects":
      return await raw("GET", V3("projects.json?pageSize=" + num(a.page_size, 250) + "&status=active"));
    case "list_tasklist_templates":
      // v1 — all account task-list templates (id, name, description). Owner-company admin key required.
      return await raw("GET", BASE() + "/tasklists/templates.json");
    case "list_project_members":
      // v1 — people currently on a project (Teamwork auto-adds everyone on create).
      return await raw("GET", BASE() + "/projects/" + a.project_id + "/people.json");
    case "get_project":
      // v1 — single project incl. "active-pages" (enabled features/tabs) + integrations + category.
      return await raw("GET", BASE() + "/projects/" + a.project_id + ".json");
    case "list_workflows": {
      // v3 — account workflows. project_ids scopes to a project's workflow; only_default = the default.
      let q = "pageSize=50";
      if (a.project_ids) q += "&projectIds=" + a.project_ids;
      if (a.only_default) q += "&onlyDefaultWorkflow=true";
      return await raw("GET", V3("workflows.json?" + q));
    }

    /* ---------- writes: perform the action, return "… ID <id>" ---------- */
    case "create_company": {
      const r = await tw("POST", V3("companies.json"), { company: { name: a.name } });
      return "Created company ID " + pickId(r, ["company"]);
    }
    case "clone_project": {
      // Teamwork has no project "copy/clone" endpoint (only files, notebooks and task lists can
      // be copied), so we create the project here and the caller recreates the template's task
      // lists afterwards. Company, description, owner, category and dates are applied via update_project.
      const r = await tw("POST", BASE() + "/projects.json", { project: { name: a.name } });
      return "Created project ID " + pickId(r, ["project"]);
    }
    case "update_project": {
      // v1 Update Project (PUT /projects/{id}.json). v3 PATCH is not allowed here (405).
      const proj = {};
      if (a.owned_id) proj.projectOwnerId = String(a.owned_id);
      // v1 expects the hyphenated "category-id"; send the camelCase alias too for safety.
      if (a.category_id) { proj["category-id"] = String(a.category_id); proj.categoryId = String(a.category_id); }
      if (a.company_id) proj.companyId = String(a.company_id);
      if (a.description != null && a.description !== "") proj.description = a.description;
      if (a.start_at) proj.startDate = String(a.start_at).replace(/-/g, "");
      if (a.end_at) proj.endDate = String(a.end_at).replace(/-/g, "");
      if (Array.isArray(a.custom_fields) && a.custom_fields.length)
        proj.customFields = a.custom_fields.map(function (c) { return { customFieldId: Number(c.id), value: String(c.value) }; });
      // feature/tab flags (use-tasks, use-billing, …) to mirror a template's enabled tabs.
      if (a.features && typeof a.features === "object") Object.assign(proj, a.features);
      await tw("PUT", BASE() + "/projects/" + a.id + ".json", { project: proj });
      return "Updated project ID " + a.id;
    }
    case "add_project_member": {
      // v3 Add people to a project — flat { userIds: [...] }.
      const ids = (a.user_ids || []).map(Number).filter(function (n) { return n > 0; });
      await tw("PUT", V3("projects/" + a.project_id + "/people.json"), { userIds: ids });
      return "Added members to project ID " + a.project_id;
    }
    case "set_project_members": {
      // v1 add/remove people on a project (merge) — used to strip the auto-added "everyone".
      const add = (a.add || []).map(Number).filter(function (n) { return n > 0; }).join(",");
      const remove = (a.remove || []).map(Number).filter(function (n) { return n > 0; }).join(",");
      const body = {};
      if (add) body.add = { userIdList: add };
      if (remove) body.remove = { userIdList: remove };
      await tw("PUT", BASE() + "/projects/" + a.project_id + "/people.json", body);
      return "Set members on project ID " + a.project_id;
    }
    case "create_custom_field_value": {
      // Set a project custom field via the v1 Update Project endpoint.
      await tw("PUT", BASE() + "/projects/" + a.entity_id + ".json", { project: { customFields: [{ customFieldId: Number(a.custom_field_id), value: String(a.value) }] } });
      return "Set custom field " + a.custom_field_id;
    }
    case "create_tasklist": {
      // v1 Create Task List (POST /projects/{id}/tasklists.json) with { "todo-list": {...} }.
      // When template_id is given, the list is instantiated from a task-list template —
      // bringing its preset tasks and assigned personnel.
      const tl = {};
      if (a.name) tl.name = a.name;
      if (a.description) tl.description = a.description;
      if (a.template_id) tl["todo-list-template-id"] = Number(a.template_id);
      if (a.template_start_date) tl["todo-list-template-start-date"] = String(a.template_start_date).replace(/-/g, "");
      if (a.assignments && typeof a.assignments === "object") tl["todo-list-template-assignments"] = a.assignments;
      const r = await tw("POST", BASE() + "/projects/" + a.project_id + "/tasklists.json", { "todo-list": tl });
      const id = r.idHeader || (r.json && (r.json.TASKLISTID || r.json.tasklistId || r.json.id)) || (((r.text || "").match(/"TASKLISTID"\s*:\s*"?(\d+)/i) || [])[1]) || "?";
      return "Created task list ID " + id;
    }
    case "copy_tasklist": {
      // v1 Copy a task list (with its tasks) to another project: PUT /tasklist/{id}/copy.json
      await tw("PUT", BASE() + "/tasklist/" + a.tasklist_id + "/copy.json", { projectId: Number(a.project_id), includeCompletedTasks: a.include_completed ? 1 : 0 });
      return "Copied task list " + a.tasklist_id + " to project " + a.project_id;
    }
    case "update_tasklist": {
      // v1 rename / set description on a task list: PUT /tasklists/{id}.json
      const tl = {};
      if (a.name != null) tl.name = a.name;
      if (a.description != null) tl.description = a.description;
      await tw("PUT", BASE() + "/tasklists/" + a.tasklist_id + ".json", { "todo-list": tl });
      return "Updated task list " + a.tasklist_id;
    }
    case "delete_tasklist": {
      // v1 delete a task list (used to remove a duplicate created from an empty template).
      await tw("DELETE", BASE() + "/tasklists/" + a.tasklist_id + ".json");
      return "Deleted task list " + a.tasklist_id;
    }
    case "apply_workflow": {
      // v3 apply a workflow (board) to a project.
      await tw("POST", V3("projects/" + a.project_id + "/workflows.json"), { workflow: { id: Number(a.workflow_id) } });
      return "Applied workflow " + a.workflow_id + " to project " + a.project_id;
    }
    case "create_task": {
      const task = { name: a.name };
      if (a.priority) task.priority = a.priority;
      const r = await tw("POST", V3("tasklists/" + a.tasklist_id + "/tasks.json"), { task: task });
      return "Created task ID " + pickId(r, ["task"]);
    }
    case "create_milestone": {
      // v1 milestone endpoint (reliable for responsible-party + tasklist linkage).
      const uids = (a.assignees && a.assignees.user_ids) || [];
      const ms = {
        title: a.name,
        deadline: String(a.due_date || "").replace(/-/g, ""),
        "responsible-party-ids": uids.join(",")
      };
      if (a.tasklist_ids && a.tasklist_ids.length) ms.tasklistIds = a.tasklist_ids.join(",");
      const r = await tw("POST", BASE() + "/projects/" + a.project_id + "/milestones.json", { milestone: ms });
      return "Created milestone ID " + pickId(r, ["milestone"]);
    }
    default:
      throw new Error("Unsupported tool: " + tool);
  }
}

/* ---------- Teamwork HTTP ---------- */
function authHeaders() {
  return { Authorization: "Basic " + Buffer.from(API_KEY + ":x").toString("base64"), Accept: "application/json" };
}
// Returns the response body text (used for list_* pass-through).
async function raw(method, url) {
  const r = await fetch(url, { method: method, headers: authHeaders() });
  const t = await r.text();
  if (!r.ok) throw new Error("Teamwork " + r.status + ": " + t.slice(0, 300));
  return t;
}
// Returns parsed JSON (plus the raw text + any "id" header) for write actions.
async function tw(method, url, bodyObj) {
  const r = await fetch(url, { method: method, headers: Object.assign({ "Content-Type": "application/json" }, authHeaders()), body: JSON.stringify(bodyObj) });
  const t = await r.text();
  if (!r.ok) throw new Error("Teamwork " + r.status + ": " + t.slice(0, 300));
  let j = null; try { j = JSON.parse(t); } catch (e) {}
  return { json: j, text: t, idHeader: r.headers.get("id") || r.headers.get("Id") };
}
// Find the new record id from a Teamwork write response (varies v1/v3).
function pickId(r, keys) {
  if (r.idHeader && /^\d+$/.test(r.idHeader)) return r.idHeader;
  const j = r.json || {};
  for (const k of keys || []) { if (j[k] && (j[k].id != null)) return String(j[k].id); }
  if (j.id != null) return String(j.id);
  const m = (r.text || "").match(/"id"\s*:\s*"?(\d+)/i);
  return m ? m[1] : "?";
}

/* ---------- helpers ---------- */
function num(v, d) { v = parseInt(v, 10); return isFinite(v) && v > 0 ? v : d; }
function isoDate(s) { s = String(s || "").replace(/-/g, ""); return s.length === 8 ? s.slice(0, 4) + "-" + s.slice(4, 6) + "-" + s.slice(6, 8) : s; }
function principal(req) {
  try {
    const h = req.headers["x-ms-client-principal"];
    if (!h) return null;
    const o = JSON.parse(Buffer.from(h, "base64").toString("utf8"));
    return { roles: o.userRoles || [] };
  } catch (e) { return null; }
}
function pack(text, isError, status) {
  return { status: status || 200, headers: { "Content-Type": "application/json" }, body: { content: [{ text: text }], isError: !!isError } };
}
