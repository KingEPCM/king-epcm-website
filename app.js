/* King EPCM Intranet — shared front-end logic.
   - Deployed on Azure Static Web Apps: real M365 user + the signed-in person's Teamwork projects & tasks.
   - Local preview: falls back to sample data so the UI still renders. */
(function () {
  "use strict";

  /* ---------- Signed-in user (Azure SWA /.auth/me) ---------- */
  function initUser() {
    var nameEl = document.querySelector(".userbox .who");
    var avEl = document.querySelector(".userbox .avatar");
    function apply(name, initialsSrc) {
      if (!name) return;
      if (nameEl) nameEl.textContent = name;
      if (avEl) avEl.textContent = initials(initialsSrc || name);
    }
    // Primary: the display name from the sign-in token claims. This needs no
    // Microsoft Graph call, so the name + initials show even on mobile.
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var p = d && d.clientPrincipal; if (!p) return;
      var nm = nameFromClaims(p);
      apply(nm.display, nm.full || nm.display);
    }).catch(function () {});
    // Refinement: the Microsoft 365 display name from our server-side endpoint
    // (works on mobile — no in-browser Microsoft sign-in needed).
    fetch("/api/my-profile").then(function (r) { return r.ok ? r.json() : null; }).then(function (me) {
      if (!me || !me.displayName) return;
      apply(me.displayName, me.givenName && me.surname ? me.givenName + " " + me.surname : me.displayName);
    }).catch(function () {});
  }
  function initials(s) {
    var parts = String(s).replace(/@.*/, "").split(/[ ._-]+/).filter(Boolean);
    return (((parts[0] || "")[0] || "") + ((parts[1] || "")[0] || "")).toUpperCase() || "?";
  }
  // Pull a usable name out of the SWA clientPrincipal claims (no Graph needed).
  function nameFromClaims(p) {
    var claims = (p && p.claims) || [];
    function get(types) {
      for (var i = 0; i < claims.length; i++) {
        var t = (claims[i].typ || "").toLowerCase();
        if (types.indexOf(t) > -1 && claims[i].val) return claims[i].val;
      }
      return "";
    }
    var display = get(["name", "http://schemas.microsoft.com/identity/claims/displayname"]);
    var given = get(["given_name", "givenname", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/givenname"]);
    var family = get(["family_name", "surname", "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/surname"]);
    var full = given && family ? given + " " + family : "";
    if (!display) display = full;
    var ud = p && p.userDetails;
    if (!display && ud && ud.indexOf("@") === -1) display = ud; // never show a raw email
    return { display: display, given: given, family: family, full: full };
  }

  // Dashboard greeting: "Welcome back, <preferred first name>"
  function initWelcome() {
    var el = document.getElementById("welcome");
    if (!el) return;
    var gotName = false; // once we have the preferred (display) name, don't overwrite it
    function setFirst(name) {
      var first = String(name || "").replace(/@.*/, "").split(/[ ._-]+/).filter(Boolean)[0];
      if (!first) return false;
      el.textContent = "Welcome back, " + first.charAt(0).toUpperCase() + first.slice(1);
      return true;
    }
    // Preferred display name from the sign-in token (works on mobile, no Graph).
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var p = d && d.clientPrincipal; if (!p) return;
      if (setFirst(nameFromClaims(p).display)) gotName = true;
    }).catch(function () {});
    // Fallback only (if the token had no usable name): use the M365 DISPLAY name —
    // the preferred name — never the legal given name.
    fetch("/api/my-profile").then(function (r) { return r.ok ? r.json() : null; }).then(function (me) {
      if (gotName || !me || !me.displayName) return;
      setFirst(me.displayName);
    }).catch(function () {});
  }

  /* ---------- Mobile sidebar ---------- */
  function initNav() {
    var h = document.getElementById("hamb"), s = document.getElementById("sidebar");
    if (h && s) h.addEventListener("click", function () { s.classList.toggle("open"); });
    var ub = document.getElementById("userMenuBtn"), um = document.getElementById("userMenu");
    if (ub && um) {
      ub.addEventListener("click", function (e) {
        e.stopPropagation();
        var open = um.classList.toggle("open");
        ub.setAttribute("aria-expanded", open ? "true" : "false");
      });
      document.addEventListener("click", function (e) { if (!um.contains(e.target)) um.classList.remove("open"); });
    }
  }

  // Teamwork sends due/end dates as a calendar date at UTC midnight, e.g.
  // "2026-06-19" or "2026-06-19T00:00:00Z". Pull the Y-M-D out of whatever form it's
  // in and build a LOCAL date, so Toronto shows the right day instead of the day before.
  function parseDue(s) {
    var m = String(s).match(/(\d{4})-(\d{2})-(\d{2})/);
    return m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(s);
  }
  function fmtDate(d) {
    if (!d) return "No due date";
    try { return parseDue(d).toLocaleDateString("en-CA", { year: "numeric", month: "short", day: "numeric" }); }
    catch (e) { return d; }
  }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  // Lightweight per-session cache (sessionStorage) so moving between pages shows
  // data instantly while a fresh copy loads in the background. ~10 min freshness.
  function cacheGet(key) {
    try { var raw = sessionStorage.getItem("ic_" + key); if (!raw) return null; var o = JSON.parse(raw); return (o && o.t && Date.now() - o.t < 600000) ? o.d : null; }
    catch (e) { return null; }
  }
  function cacheSet(key, data) {
    try { sessionStorage.setItem("ic_" + key, JSON.stringify({ t: Date.now(), d: data })); } catch (e) {}
  }

  /* ---------- Teamwork: projects + due/overdue tasks (one fetch) ---------- */
  var SAMPLE_PROJECTS = [
    { name: "203 Church Street, Markham", category: "11. Re-zoning & Subdivisions", due: "2026-06-26", status: "late", url: "#" },
    { name: "240 Industrial Pkwy S, Aurora", category: "04. Geotechnical & Septics", due: "2026-07-02", status: "current", url: "#" },
    { name: "1 Kingscross Dr, King", category: "04. Geotechnical & Septics", due: "2026-07-15", status: "current", url: "#" },
    { name: "280 Hopkins Street, Whitby", category: "11. Re-zoning & Subdivisions", due: "2026-08-04", status: "upcoming", url: "#" }
  ];
  var SAMPLE_TASKS = [
    { name: "Submit Phase II ESA report", project: "203 Church Street, Markham", due: "2026-06-17", overdue: true, url: "#" },
    { name: "Follow up on building permit comments", project: "1 Kingscross Dr, King", due: "2026-06-22", overdue: false, url: "#" },
    { name: "Book borehole drilling", project: "240 Industrial Pkwy S, Aurora", due: "2026-06-27", overdue: false, url: "#" }
  ];
  var _allProjects = [];

  function projStatus(status, due) {
    var s = String(status || "").toLowerCase();
    if (s.indexOf("complet") > -1) return { cls: "ok", label: "Completed" };
    if (s.indexOf("late") > -1) return { cls: "late", label: "Overdue" };
    if (due) {
      var today = new Date(); today.setHours(0, 0, 0, 0);
      var days = (parseDue(due) - today) / 86400000;
      if (days < 0) return { cls: "late", label: "Overdue" };
      if (days <= 10) return { cls: "warn", label: "Due soon" };
    }
    return { cls: "ok", label: "On track" };
  }
  function renderProjects(el, items) {
    if (!el) return;
    if (!items.length) { el.innerHTML = '<div class="note" style="margin:0">No projects to show.</div>'; return; }
    el.innerHTML = items.map(function (p) {
      var si = projStatus(p.status, p.due);
      var href = p.url && p.url !== "#" ? ' href="' + p.url + '" target="_blank" rel="noopener"' : ' href="projects.html"';
      var meta = [p.category, "Due " + fmtDate(p.due)].filter(Boolean).join(" · ");
      return '<a class="proj" style="display:block"' + href + '>' +
        '<div class="row1"><span class="pname">' + esc(p.name) + '</span>' +
        '<span class="pill ' + si.cls + '">' + si.label + '</span></div>' +
        '<div class="meta">' + esc(meta) + '</div></a>';
    }).join("");
  }
  // Date-only-aware: a task is overdue only if its due date is before today (local).
  // Due *today* counts as "due soon", not overdue. Avoids the UTC-midnight off-by-one.
  function isOverdue(due) {
    if (!due) return false;
    var today = new Date(); today.setHours(0, 0, 0, 0);
    return parseDue(due) < today;
  }
  function taskItemHtml(t) {
    var href = t.url && t.url !== "#" ? ' href="' + t.url + '" target="_blank" rel="noopener"' : ' href="projects.html"';
    var meta = [t.project, "Due " + fmtDate(t.due)].filter(Boolean).join(" · ");
    return '<a class="proj" style="display:block"' + href + '>' +
      '<div class="row1"><span class="pname">' + esc(t.name) + '</span></div>' +
      '<div class="meta">' + esc(meta) + '</div></a>';
  }
  // Group tasks into Overdue / Due today / Due later, each a collapsible section.
  // Dashboard "My tasks": only Overdue (expanded) and Due today (collapsed). Date-only buckets.
  function renderTasks(el, items) {
    if (!el) return;
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var overdue = [], dueToday = [];
    (items || []).forEach(function (t) {
      if (!t.due) return;
      var d = parseDue(t.due);
      if (d < today) overdue.push(t);
      else if (d.getTime() === today.getTime()) dueToday.push(t);
    });
    var groups = [
      { label: "Overdue", items: overdue, cls: "late", open: true },
      { label: "Due today", items: dueToday, cls: "warn", open: false }
    ].filter(function (g) { return g.items.length; });
    if (!groups.length) {
      el.innerHTML = '<div class="note" style="margin:0">🎉 Nothing overdue or due today.</div>';
      return;
    }
    el.innerHTML = groups.map(function (g) {
      return '<details class="proc-group task-group"' + (g.open ? " open" : "") + '>' +
        '<summary><span class="task-dot ' + g.cls + '"></span>' + g.label +
        '<span class="proc-count">' + g.items.length + '</span></summary>' +
        '<div class="proc-group-body">' + g.items.map(taskItemHtml).join("") + '</div></details>';
    }).join("");
  }

  /* ---------- My tasks (from /api/my-tasks) — drives BOTH the My Work full list and the dashboard preview ---------- */
  function initMyWork() {
    var fullEl = document.getElementById("my-task-list");   // My Work: full grouped list
    var dashEl = document.getElementById("task-list");      // Dashboard: Overdue + Due today preview
    if (!fullEl && !dashEl) return;
    var whoEl = document.getElementById("mw-who");

    // My Work full view: Overdue / Due today / Due soon / Later / No due date, all collapsed.
    function renderFull(tasks) {
      if (!tasks.length) { fullEl.innerHTML = '<div class="note" style="margin:0">🎉 No open tasks assigned to you in Teamwork right now.</div>'; return; }
      var today = new Date(); today.setHours(0, 0, 0, 0);
      var soonLimit = new Date(today.getTime() + 14 * 86400000);
      function bucket(t) {
        if (!t.due) return "none";
        var d = parseDue(t.due);
        if (d < today) return "overdue";
        if (d.getTime() === today.getTime()) return "today";
        if (d <= soonLimit) return "soon";
        return "later";
      }
      var groups = [
        { label: "Overdue", items: tasks.filter(function (t) { return bucket(t) === "overdue"; }), open: false },
        { label: "Due today", items: tasks.filter(function (t) { return bucket(t) === "today"; }), open: false },
        { label: "Due soon", items: tasks.filter(function (t) { return bucket(t) === "soon"; }), open: false },
        { label: "Later", items: tasks.filter(function (t) { return bucket(t) === "later"; }), open: false },
        { label: "No due date", items: tasks.filter(function (t) { return bucket(t) === "none"; }), open: false }
      ].filter(function (g) { return g.items.length; });
      fullEl.innerHTML = groups.map(function (g) {
        return '<details class="proc-group"' + (g.open ? " open" : "") + '>' +
          '<summary>' + g.label + '<span class="proc-count">' + g.items.length + '</span></summary>' +
          '<div class="proc-group-body">' + g.items.map(taskItemHtml).join("") + '</div></details>';
      }).join("");
    }
    function render(d) {
      var tasks = (d && d.tasks) || [];
      if (whoEl && d && d.person && d.person.name) whoEl.textContent = d.person.name;
      if (fullEl) renderFull(tasks);
      if (dashEl) renderTasks(dashEl, tasks);   // dashboard: Overdue (open) + Due today (collapsed)
    }
    var cached = cacheGet("mytasks");
    if (cached) render(cached);
    fetch("/api/my-tasks").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      if (d && d.configured) { cacheSet("mytasks", d); render(d); }
      else if (!cached) {
        if (fullEl) fullEl.innerHTML = '<div class="note" style="margin:0">Your Teamwork tasks will appear here once Teamwork is connected. <a href="https://kingepcm.teamwork.com" target="_blank" rel="noopener">Open Teamwork →</a></div>';
        if (dashEl) renderTasks(dashEl, SAMPLE_TASKS);
      }
    }).catch(function () {
      if (!cached) {
        if (fullEl) fullEl.innerHTML = '<div class="note" style="margin:0">Couldn\'t load your tasks right now. <a href="https://kingepcm.teamwork.com" target="_blank" rel="noopener">Open Teamwork →</a></div>';
        if (dashEl) renderTasks(dashEl, SAMPLE_TASKS);
      }
    });
  }

  /* ---------- My Work: reveal the admin-only project-tool cards (kick-off card is shown to all) ---------- */
  function initProjectTools() {
    var sec = document.getElementById("proj-tools");
    if (!sec) return;
    var adminCards = sec.querySelectorAll(".proj-admin");
    if (!adminCards.length) return;
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("admin") === -1) return;        // admin cards only
      adminCards.forEach(function (c) { c.hidden = false; });
    }).catch(function () {});
  }

  /* ---------- Project tool pages (new-project.html / new-scope-of-work.html): gate + load form ---------- */
  function initProjectToolPage() {
    var content = document.getElementById("pt-content");
    var frame = document.getElementById("pt-frame");
    if (!content || !frame) return;                      // not on a tool page
    var gate = document.getElementById("pt-gate");
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("admin") === -1) { if (gate) gate.hidden = false; return; } // admins only
      content.hidden = false;
      if (!frame.getAttribute("src")) frame.setAttribute("src", frame.getAttribute("data-src")); // lazy-load
    }).catch(function () { if (gate) gate.hidden = false; });
  }

  /* ---------- Project kick-off email (project-kickoff.html): compose → preview → send as the user ---------- */
  function initKickoff() {
    var panel = document.getElementById("kf-panel");
    if (!panel) return;                                   // not on the kick-off page
    var $ = function (id) { return document.getElementById(id); };
    var ONBOARD_URL = "https://kingepcm.com/project-onboarding";
    var SIG_KEY = "kf-signature";
    var LOGO_DISPLAY_URL = "assets/logo-full.png";   // shown in the browser preview; email uses cid:kelogo
    var previewBuilt = false, fieldsChanged = true, subjectEdited = false;

    // Pre-fill sender details + signature from the M365 profile.
    fetch("/api/my-profile").then(function (r) { return r.ok ? r.json() : null; }).then(function (me) {
      if (!me) return;
      var full = me.displayName || ((me.givenName || "") + " " + (me.surname || "")).trim();
      var bphone = (me.businessPhones && me.businessPhones[0]) || "";
      var ext = "";
      var m = bphone.match(/(?:ext\.?|x)\s*(\d+)/i);
      if (m) { ext = m[1]; bphone = bphone.replace(/\s*(?:ext\.?|x)\s*\d+/i, "").trim(); }
      var phone = bphone || me.mobilePhone || "416-342-3001";
      if ($("kf-from")) $("kf-from").textContent = me.mail || "your mailbox";
      if (!$("kf-sender-name").value) $("kf-sender-name").value = full;   // point of contact = full display name
      if (!$("kf-phone").value) $("kf-phone").value = phone;
      if (!$("kf-ext").value && ext) $("kf-ext").value = ext;
      // Signature: use a saved one if present, otherwise build a branded default from the profile.
      var saved = "";
      try { saved = window.localStorage.getItem(SIG_KEY) || ""; } catch (e) {}
      if (!$("kf-signature").value) $("kf-signature").value = saved || defaultSignature(me, ext);
      fieldsChanged = true;
    }).catch(function () {});

    // Default signature, matching the King EPCM house style (the logo image is added automatically).
    function defaultSignature(me, ext) {
      var lines = ["Sincerely,", me.displayName || ""];
      if (me.jobTitle) lines.push(me.jobTitle);
      lines.push("3780 14th Ave Unit 211");
      lines.push("Markham, ON, L3R 9Y5");
      lines.push("Office: 416-342-3001" + (ext ? "x" + ext : ""));
      if (me.mail) lines.push("Email: " + me.mail);
      return lines.filter(Boolean).join("\n");
    }

    // Subject = "Project kick-off - <project name>", auto-synced until the user edits it.
    function syncSubject() {
      if (subjectEdited) return;
      var pn = $("kf-project-name").value.trim();
      $("kf-subject").value = "Project kick-off" + (pn ? " - " + pn : "");
    }
    $("kf-subject").addEventListener("input", function () { subjectEdited = true; });
    $("kf-project-name").addEventListener("input", syncSubject);
    syncSubject();

    // Deliverables: add/remove rows (numbered).
    var delivRows = $("kf-deliv-rows");
    function renumberDeliv() {
      Array.prototype.forEach.call(delivRows.children, function (row, i) {
        var n = row.querySelector(".kf-num"); if (n) n.textContent = (i + 1) + ".";
      });
    }
    function addDeliv(val, focus) {
      var row = document.createElement("div"); row.className = "kf-deliv";
      row.innerHTML = '<span class="kf-num"></span>' +
        '<input type="text" data-k="del" placeholder="e.g. Phase II ESA report" autocomplete="off">' +
        '<button type="button" class="rb-del" title="Remove deliverable" aria-label="Remove deliverable">×</button>';
      row.querySelector(".rb-del").onclick = function () {
        row.remove();
        if (!delivRows.children.length) addDeliv("");
        renumberDeliv();
      };
      if (val) row.querySelector("input").value = val;
      delivRows.appendChild(row);
      renumberDeliv();
      if (focus) row.querySelector("input").focus();
    }
    if (!delivRows.children.length) addDeliv("");
    $("kf-add-deliv").addEventListener("click", function () { addDeliv("", true); });

    function sigToHtml(text) {
      var lines = String(text || "").split(/\r?\n/);
      var out = lines.map(function (ln) {
        if (!ln.trim()) return "<div style='height:6px'></div>";
        var safe = esc(ln).replace(/([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/g,
          '<a href="mailto:$1" style="color:#2a6df4">$1</a>');
        return "<div>" + safe + "</div>";
      }).join("");
      // King EPCM logo — referenced by cid; the server attaches it inline (preview swaps to a local URL).
      var logo = "<div style='margin-top:10px'><img data-ke-logo=\"1\" src=\"cid:kelogo\" alt=\"King EPCM\" style=\"height:54px;width:auto;border:0\"></div>";
      return "<div style='margin-top:18px;font-size:13px;color:#333;line-height:1.5'>" + out + logo + "</div>";
    }

    // Canonical email body — used for BOTH the live preview and what gets sent.
    function buildBodyHtml() {
      var clientName = $("kf-client-name").value.trim() || "there";
      var sender = $("kf-sender-name").value.trim() || "";
      var estimate = $("kf-estimate").value.trim();
      var weeks = $("kf-weeks").value.trim();
      var phone = $("kf-phone").value.trim();
      var ext = $("kf-ext").value.trim();
      var dels = Array.prototype.map.call(delivRows.querySelectorAll('input[data-k="del"]'), function (i) { return i.value.trim(); }).filter(Boolean);

      var delHtml = dels.length
        ? "<ol style='margin:0 0 14px;padding-left:22px'>" + dels.map(function (d) { return "<li style='margin:0 0 4px'>" + esc(d) + "</li>"; }).join("") + "</ol>"
        : "<ol style='margin:0 0 14px;padding-left:22px'><li style='color:#999'>[ add your deliverables ]</li></ol>";

      var callLine = "call me at " + (esc(phone) || "416-342-3001") + (ext ? " Ext " + esc(ext) : "");

      return "" +
        "<div style=\"font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.55\">" +
        "<p>Hi " + esc(clientName) + ",</p>" +
        "<p>My name is " + (esc(sender) || "[your name]") + " of King EPCM. I am the King EPCM Project Owner on your file and will also be the primary point of contact on this file for you.</p>" +
        "<p>We&rsquo;ve had an internal project kick-off already, and my understanding is that based off the Estimate # " + (esc(estimate) || "[estimate #]") + ", our primary deliverables are:</p>" +
        delHtml +
        "<p>We are trying to wrap up the projects on-hand, and once we start working on this project in-full, it will be approx. " + (esc(weeks) || "[ ]") + " weeks until draft completion.<br>" +
        "Before we start on the project, our dispatch team will reach out to schedule any site visits if required. In order for us to get the project completed smoothly, please fill out the digital form in the link below, and let us know all the necessary info:</p>" +
        "<p><a href=\"" + ONBOARD_URL + "\" style=\"color:#14294D;font-weight:600\">King EPCM Project Onboarding Form</a></p>" +
        "<p>Finally, if you have any questions, I am your primary point of contact. Feel free to reply directly back to this email chain or " + callLine + ".</p>" +
        "<p>Alternatively, below are some direct access:</p>" +
        "<ul style='margin:0 0 14px;padding-left:22px'>" +
        "<li style='margin:0 0 4px'>Dispatch for site visits &ndash; Mahmoud Seddigh &ndash; <a href=\"mailto:mseddigh@kingepcm.com\">mseddigh@kingepcm.com</a> - 416-342-3001 Ext 209</li>" +
        "<li style='margin:0 0 4px'>Direct supervisor &ndash; Tony Wang &ndash; <a href=\"mailto:Twang@kingepcm.com\">Twang@kingepcm.com</a> &ndash; 416-342-3001 Ext 108</li>" +
        "<li style='margin:0 0 4px'>Sales - <a href=\"mailto:sales@kingepcm.com\">Sales@kingepcm.com</a> &ndash; 416-342-3001 Ext 108</li>" +
        "</ul>" +
        sigToHtml($("kf-signature").value) +
        "</div>";
    }

    // Render the preview from the fields, and show the logo via a local URL (email uses cid).
    function refreshPreview() {
      $("kf-preview").innerHTML = buildBodyHtml();
      Array.prototype.forEach.call($("kf-preview").querySelectorAll('img[data-ke-logo]'), function (im) { im.setAttribute("src", LOGO_DISPLAY_URL); });
      $("kf-pv-to").textContent = $("kf-client-email").value.trim() || "—";
      $("kf-pv-subject").textContent = $("kf-subject").value.trim() || "—";
      var cc = $("kf-cc").value.trim();
      $("kf-pv-cc-row").hidden = !cc;
      $("kf-pv-cc").textContent = cc;
    }
    // Any change in Compose marks the preview stale (so it rebuilds, replacing manual edits).
    $("kf-compose").addEventListener("input", function () { fieldsChanged = true; });

    // The HTML actually sent: take the (possibly edited) preview, or build fresh; restore the logo to cid.
    function finalBodyHtml() {
      var html = previewBuilt ? $("kf-preview").innerHTML : buildBodyHtml();
      var tmp = document.createElement("div"); tmp.innerHTML = html;
      Array.prototype.forEach.call(tmp.querySelectorAll('img[data-ke-logo]'), function (im) { im.setAttribute("src", "cid:kelogo"); });
      return tmp.innerHTML;
    }

    // Tabs (Compose / Preview)
    function showTab(which) {
      var preview = which === "preview";
      panel.querySelectorAll(".cal-vbtn").forEach(function (b) { b.classList.toggle("active", b.getAttribute("data-kf") === which); });
      $("kf-compose").hidden = preview;
      $("kf-preview-wrap").hidden = !preview;
      if (preview && (!previewBuilt || fieldsChanged)) { refreshPreview(); previewBuilt = true; fieldsChanged = false; }
    }
    panel.querySelectorAll(".cal-vbtn").forEach(function (b) {
      b.addEventListener("click", function () { showTab(b.getAttribute("data-kf")); });
    });
    if ($("kf-preview-btn")) $("kf-preview-btn").addEventListener("click", function () { showTab("preview"); window.scrollTo({ top: 0, behavior: "smooth" }); });

    // Save the signature locally for next time.
    $("kf-signature").addEventListener("change", function () {
      try { window.localStorage.setItem(SIG_KEY, $("kf-signature").value); } catch (e) {}
    });

    // Send
    var status = $("kf-status");
    function emailsValid(v) { return String(v).split(/[,;]+/).map(function (s) { return s.trim(); }).filter(Boolean).every(function (s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }); }
    $("kf-send").addEventListener("click", function () {
      var to = $("kf-client-email").value.trim();
      var subject = $("kf-subject").value.trim();
      if (!to || !emailsValid(to)) { status.style.color = "#d6336c"; status.textContent = "Enter a valid client email address."; return; }
      var cc = $("kf-cc").value.trim();
      if (cc && !emailsValid(cc)) { status.style.color = "#d6336c"; status.textContent = "One of the CC addresses looks invalid."; return; }
      if (!subject) { status.style.color = "#d6336c"; status.textContent = "Add a subject line."; return; }
      if (!$("kf-client-name").value.trim() && !confirm("No client name entered — the email will say “Hi there,”. Send anyway?")) return;

      var btn = $("kf-send");
      btn.disabled = true; status.style.color = "var(--muted)"; status.textContent = "Sending…";
      fetch("/api/project-kickoff", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ to: to, cc: cc, subject: subject, bodyHtml: finalBodyHtml() })
      }).then(function (r) { return r.json().then(function (d) { return { ok: r.ok, d: d }; }); })
        .then(function (res) {
          btn.disabled = false;
          if (res.ok && res.d && res.d.ok) {
            status.style.color = "#198754";
            status.textContent = "✓ Sent to " + to + (res.d.sentFrom ? " from " + res.d.sentFrom : "") + ". A copy is in your Sent Items.";
          } else {
            status.style.color = "#d6336c";
            status.textContent = (res.d && res.d.error) || "Couldn't send. Please try again.";
          }
        }).catch(function () { btn.disabled = false; status.style.color = "#d6336c"; status.textContent = "Couldn't send right now. Please try again."; });
    });
  }

  // Projects list (dashboard + My Work). Tasks come from initMyWork (/api/my-tasks) so both pages agree.
  function initTeamwork() {
    var projEl = document.getElementById("project-list");
    if (!projEl) return;
    var projLimit = parseInt(projEl.getAttribute("data-limit") || "0", 10);

    function render(data) {
      var projects = data.projects || (Array.isArray(data) ? data : []);
      applyProjects(projEl, projects, projLimit);
    }
    var cached = cacheGet("tw");
    if (cached) render(cached); // instant from cache
    fetch("/api/teamwork-projects")
      .then(function (r) { return r.ok ? r.json() : Promise.reject(); })
      .then(function (data) { cacheSet("tw", data); render(data); })
      .catch(function () {
        if (cached) return; // keep showing cached data on a transient failure
        applyProjects(projEl, SAMPLE_PROJECTS, projLimit);
      });
  }
  // Project list: show 5 by default with a Show all / Show fewer toggle, and live search. Used on
  // both the dashboard and My Work.
  function applyProjects(el, items, limit) {
    if (!el) return;
    _allProjects = items || [];
    var DEFAULT = limit || 5;
    var search = document.getElementById("project-search");
    var expanded = false;

    function draw() {
      var q = (search && search.value.toLowerCase().trim()) || "";
      var filtered = !q ? _allProjects : _allProjects.filter(function (p) {
        return (p.name + " " + (p.category || "")).toLowerCase().indexOf(q) > -1;
      });
      if (q && !filtered.length) {
        el.innerHTML = '<div class="note" style="margin:0">No projects match “' + esc(q) + '”.</div>';
        return;
      }
      var showAll = !!q || expanded;                       // searching always shows all matches
      renderProjects(el, showAll ? filtered : filtered.slice(0, DEFAULT));
      if (!q && filtered.length > DEFAULT) {
        var btn = document.createElement("button");
        btn.type = "button"; btn.className = "show-more";
        btn.textContent = expanded ? "Show fewer ▲" : ("Show all " + filtered.length + " ▼");
        btn.onclick = function () { expanded = !expanded; draw(); };
        el.appendChild(btn);
      }
    }
    if (search) search.oninput = function () { expanded = false; draw(); };
    draw();
  }

  /* ---------- Company news (Azure Table Storage via /api/news; admins can post) ---------- */
  function initNews() {
    var el = document.getElementById("news-list");
    if (!el) return;
    var limit = parseInt(el.getAttribute("data-limit") || "0", 10);
    var adminEl = document.getElementById("news-admin"); // present only on the full News page
    var items = [], isAdmin = false;

    function byId(id) { for (var i = 0; i < items.length; i++) if (String(items[i].id) === String(id)) return items[i]; return null; }
    function render() {
      var list = limit ? items.slice(0, limit) : items;
      if (!list.length) { el.innerHTML = '<div class="note" style="margin:0">No news yet.</div>'; return; }
      el.innerHTML = list.map(function (n) {
        var editable = isAdmin && adminEl && String(n.id).indexOf("static-") !== 0;
        var flag = n.pinned ? '<span class="news-pinflag">📌 Pinned</span>' : "";
        var actions = editable
          ? '<div class="news-actions"><button type="button" class="news-pinbtn" data-id="' + esc(n.id) + '">' + (n.pinned ? "Unpin" : "Pin to top") + "</button>" +
            '<button type="button" class="news-edit" data-id="' + esc(n.id) + '">Edit</button>' +
            '<button type="button" class="news-del" data-id="' + esc(n.id) + '">Delete</button></div>' : "";
        return '<div class="news-item' + (n.pinned ? " pinned" : "") + '"><div class="date">' + esc(n.date) + flag + '</div>' +
          '<div class="nt">' + esc(n.title) + '</div><div class="nx">' + esc(n.body) + '</div>' + actions + "</div>";
      }).join("");
      if (isAdmin && adminEl) {
        el.querySelectorAll(".news-edit").forEach(function (b) { b.onclick = function () { var n = byId(b.getAttribute("data-id")); if (n) fillForm(n); }; });
        el.querySelectorAll(".news-del").forEach(function (b) { b.onclick = function () { delPost(b.getAttribute("data-id")); }; });
        el.querySelectorAll(".news-pinbtn").forEach(function (b) { b.onclick = function () { togglePin(b.getAttribute("data-id")); }; });
      }
    }
    function load() {
      var cached = cacheGet("news");
      if (cached) { items = cached; render(); } // instant from cache
      fetch("/api/news").then(function (r) { return r.ok ? r.json() : Promise.reject(); })
        .then(function (data) { if (!Array.isArray(data)) throw 0; items = data; cacheSet("news", data); render(); })
        .catch(function () { // storage not set up yet → show the static seed file (read-only)
          if (cached) return;
          fetch("news.json").then(function (r) { return r.json(); })
            .then(function (seed) { items = (seed || []).map(function (n, i) { return { id: "static-" + i, date: n.date, title: n.title, body: n.body }; }); render(); })
            .catch(function () {});
        });
    }
    function delPost(id) {
      if (!window.confirm("Delete this news post?")) return;
      fetch("/api/news?id=" + encodeURIComponent(id), { method: "DELETE" })
        .then(function (r) { if (r.ok) load(); else alert("Couldn't delete — admins only."); });
    }
    function togglePin(id) {
      var n = byId(id); if (!n) return;
      var payload = { id: n.id, title: n.title, body: n.body, date: n.date, pinned: !n.pinned };
      fetch("/api/news", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
        .then(function (r) { if (r.ok) load(); else alert("Couldn't update — admins only."); });
    }

    // ---- Admin compose / edit form (News page only) ----
    var fillForm = function () {};
    if (adminEl) {
      var form = document.getElementById("news-form");
      var idEl = document.getElementById("news-id"), titleEl = document.getElementById("news-title"),
        bodyEl = document.getElementById("news-body"), dateEl = document.getElementById("news-date"),
        statusEl = document.getElementById("news-status"), cancelEl = document.getElementById("news-cancel"),
        heading = document.getElementById("news-form-heading"), submitEl = document.getElementById("news-submit"),
        pinnedEl = document.getElementById("news-pinned");
      function reset() { idEl.value = ""; titleEl.value = ""; bodyEl.value = ""; dateEl.value = ""; if (pinnedEl) pinnedEl.checked = false; heading.textContent = "Post an update"; submitEl.textContent = "Post update"; cancelEl.hidden = true; statusEl.textContent = ""; }
      fillForm = function (n) {
        idEl.value = n.id; titleEl.value = n.title; bodyEl.value = n.body; dateEl.value = n.date || ""; if (pinnedEl) pinnedEl.checked = !!n.pinned;
        heading.textContent = "Edit update"; submitEl.textContent = "Save changes"; cancelEl.hidden = false;
        adminEl.scrollIntoView({ behavior: "smooth", block: "start" });
      };
      cancelEl.onclick = reset;
      form.onsubmit = function (e) {
        e.preventDefault();
        var payload = { title: titleEl.value.trim(), body: bodyEl.value.trim(), pinned: !!(pinnedEl && pinnedEl.checked) };
        if (idEl.value) payload.id = idEl.value;
        if (dateEl.value.trim()) payload.date = dateEl.value.trim();
        if (!payload.title || !payload.body) { statusEl.textContent = "Title and message are required."; return; }
        submitEl.disabled = true; statusEl.textContent = "Saving…";
        fetch("/api/news", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
          .then(function (r) { return r.ok ? r.json() : Promise.reject(r); })
          .then(function () { submitEl.disabled = false; reset(); statusEl.textContent = "Saved."; load(); setTimeout(function () { statusEl.textContent = ""; }, 2500); })
          .catch(function () { submitEl.disabled = false; statusEl.textContent = "Couldn't save — you need the admin role, or storage isn't set up yet."; });
      };
      // Reveal the form only for admins.
      fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
        var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
        isAdmin = roles.indexOf("admin") > -1;
        if (isAdmin) { adminEl.hidden = false; render(); }
      }).catch(function () {});
    }

    load();
  }
  var _allProcs = [];
  // Flat list with department tags (used for the dashboard preview)
  function renderProcRows(el, items) {
    if (!items.length) { el.innerHTML = '<div class="note" style="margin:0">No flowcharts yet.</div>'; return; }
    el.innerHTML = items.map(function (p) {
      return '<a class="proc-row" href="' + esc(p.url) + '" target="_blank" rel="noopener">' +
        '<span class="proc-name">📄 ' + esc(p.title) + '</span>' +
        '<span class="proc-dept">' + esc(p.dept) + '</span></a>';
    }).join("");
  }
  // Collapsible category sections, sorted A–Z, items A–Z (used for the full page)
  function renderProcGrouped(el, items, open) {
    if (!items.length) { el.innerHTML = '<div class="note" style="margin:0">No matching flowcharts.</div>'; return; }
    var groups = {};
    items.forEach(function (p) { (groups[p.dept] = groups[p.dept] || []).push(p); });
    var cats = Object.keys(groups).sort(function (a, b) { return a.localeCompare(b); });
    el.innerHTML = cats.map(function (cat) {
      var list = groups[cat].slice().sort(function (a, b) { return a.title.localeCompare(b.title); });
      return '<details class="proc-group"' + (open ? " open" : "") + '>' +
        '<summary>' + esc(cat) + '<span class="proc-count">' + list.length + '</span></summary>' +
        '<div class="proc-group-body">' + list.map(function (p) {
          return '<a class="proc-row" href="' + esc(p.url) + '" target="_blank" rel="noopener">' +
            '<span class="proc-name">📄 ' + esc(p.title) + '</span></a>';
        }).join("") + '</div></details>';
    }).join("");
  }
  function initProcedures() {
    var el = document.getElementById("proc-list");
    if (!el) return;
    var limit = parseInt(el.getAttribute("data-limit") || "0", 10);
    function render(cats) {
      var all = [];
      (cats || []).forEach(function (c) {
        (c.items || []).forEach(function (i) { all.push({ title: i.title, dept: c.category, url: i.url }); });
      });
      _allProcs = all;
      var cnt = document.getElementById("count-sops");
      if (cnt) cnt.textContent = all.length;
      var s = document.getElementById("proc-search");
      if (!s) { // dashboard preview: short alphabetical flat list
        var preview = all.slice().sort(function (a, b) { return a.title.localeCompare(b.title); });
        renderProcRows(el, limit ? preview.slice(0, limit) : preview);
        return;
      }
      renderProcGrouped(el, all, false); // full page: collapsed category sections
      s.oninput = function () {
        var q = s.value.toLowerCase().trim();
        var f = !q ? _allProcs : _allProcs.filter(function (p) {
          return (p.title + " " + p.dept).toLowerCase().indexOf(q) > -1;
        });
        renderProcGrouped(el, f, !!q); // auto-expand sections while searching
      };
    }
    // Live list from SharePoint (/api/flowcharts). Show the cached copy instantly,
    // then refresh in the background. Fall back to the bundled procedures.json only
    // if there's no cache and the live source is unconfigured/empty/fails.
    var cached = cacheGet("flowcharts");
    if (cached && cached.length) render(cached);
    function fallback() { if (cached && cached.length) return; fetch("procedures.json").then(function (r) { return r.json(); }).then(render).catch(function () {}); }
    fetch("/api/flowcharts").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      if (d && d.configured && d.categories && d.categories.length) { cacheSet("flowcharts", d.categories); render(d.categories); }
      else fallback();
    }).catch(fallback);
  }

  /* ---------- Resources: templates, vendors, links (resources.json) ---------- */
  function resLinkHtml(items, icon) {
    if (!items || !items.length) return '<div class="note" style="margin:0">Nothing here yet — your team can add items in resources.json.</div>';
    return items.map(function (i) {
      var note = i.note ? ' <span style="color:var(--muted);font-size:.8rem">(' + esc(i.note) + ')</span>' : "";
      var label = '<span class="fic">' + icon + '</span> ' + esc(i.title) + note;
      // External links (http/https) open in a new tab; internal pages — the generators &
      // builders — stay in the same tab so they don't pop a new window each time.
      var ext = /^(https?:)?\/\//i.test(i.url || "");
      var tgt = ext ? ' target="_blank" rel="noopener"' : "";
      return i.url
        ? '<a href="' + esc(i.url) + '"' + tgt + ">" + label + "</a>"
        : '<span class="res-static">' + label + "</span>";
    }).join("");
  }
  function resVendorHtml(items) {
    if (!items || !items.length) return '<li><span class="hname">No vendors listed yet.</span></li>';
    return items.map(function (v) {
      var name = esc(v.name) + (v.service ? ' <span class="vendor-svc">— ' + esc(v.service) + "</span>" : "");
      var contact = v.url
        ? '<a href="' + esc(v.url) + '" target="_blank" rel="noopener">' + esc(v.contact || "Visit site") + "</a>"
        : esc(v.contact || "");
      return '<li><span class="hname">' + name + '</span><span class="hdate">' + contact + "</span></li>";
    }).join("");
  }
  function resCount(id, arr) {
    var el = document.getElementById(id);
    if (el) el.textContent = (arr && arr.length) || 0;
  }
  var _allTemplates = [];
  // Grouped (collapsible) template list, by the "cat" field, A–Z. Expands while searching.
  function renderTemplates(items, expand) {
    var t = document.getElementById("templates-list");
    if (!t) return;
    if (!_allTemplates.length) { t.innerHTML = resLinkHtml(_allTemplates, "📄"); return; }
    if (!items || !items.length) {
      t.innerHTML = '<div class="note" style="margin:0">No templates match your search.</div>';
      return;
    }
    var groups = {};
    items.forEach(function (i) { var c = i.cat || "Other"; (groups[c] = groups[c] || []).push(i); });
    // "Generators & tools" (intranet tools) pinned first, then the rest A–Z.
    var cats = Object.keys(groups).sort(function (a, b) {
      var ga = a === "Generators & tools" ? 0 : 1, gb = b === "Generators & tools" ? 0 : 1;
      return ga - gb || a.localeCompare(b);
    });
    t.innerHTML = cats.map(function (cat) {
      return '<details class="proc-group"' + (expand ? " open" : "") + '>' +
        '<summary>' + esc(cat) + '<span class="proc-count">' + groups[cat].length + '</span></summary>' +
        '<div class="proc-group-body">' + resLinkHtml(groups[cat], "📄") + '</div></details>';
    }).join("");
  }
  function initResources() {
    var t = document.getElementById("templates-list");
    var v = document.getElementById("vendors-list");
    var k = document.getElementById("links-list");
    if (!t && !v && !k) return;
    fetch("resources.json").then(function (r) { return r.json(); }).then(function (d) {
      if (k) k.innerHTML = resLinkHtml(d.links, "🔗");
      if (v) v.innerHTML = resVendorHtml(d.vendors);
      resCount("count-vendors", d.vendors);
      resCount("count-links", d.links);
      var curated = d.templates || [];

      function applyTemplates(list) {
        _allTemplates = list;
        renderTemplates(_allTemplates, false);
        resCount("count-templates", _allTemplates);
        syncResTabSelect();
        var ts = document.getElementById("tpl-search");
        if (ts) ts.oninput = function () {
          var q = ts.value.toLowerCase().trim();
          var f = !q ? _allTemplates : _allTemplates.filter(function (i) {
            return ((i.title || "") + " " + (i.note || "") + " " + (i.cat || "")).toLowerCase().indexOf(q) > -1;
          });
          renderTemplates(f, !!q);
        };
      }

      // Live templates from SharePoint (/api/templates), plus the curated intranet
      // generator/tool links (the ones that point at an .html page). Show the cached
      // copy instantly, refresh in the background, and fall back to the full curated
      // list only if there's no cache and the live source is unconfigured/empty/fails.
      var tools = curated.filter(function (x) { return /\.html?($|[?#])/i.test(x.url || ""); });
      function flatten(cats) {
        var live = [];
        (cats || []).forEach(function (c) {
          (c.items || []).forEach(function (i) { live.push({ title: i.title, url: i.url, note: i.note, cat: c.category }); });
        });
        return live;
      }
      var cachedLive = cacheGet("tpl_live");
      // Show something immediately — the cached live list if we have it, otherwise the curated list —
      // so the Templates tab is never blank while the (sometimes slow) SharePoint fetch runs.
      applyTemplates(cachedLive && cachedLive.length ? tools.concat(flatten(cachedLive)) : curated);
      // Refresh from SharePoint in the background, with a timeout so a slow/hung API can't leave it blank.
      var tplAc = new AbortController();
      var tplTimer = setTimeout(function () { tplAc.abort(); }, 25000);
      fetch("/api/templates", { signal: tplAc.signal })
        .then(function (r) { clearTimeout(tplTimer); return r.ok ? r.json() : null; })
        .then(function (td) {
          if (td && td.configured && td.categories && td.categories.length) {
            cacheSet("tpl_live", td.categories);
            applyTemplates(tools.concat(flatten(td.categories)));
          } // else: keep whatever is already on screen (cached or curated)
        })
        .catch(function () { clearTimeout(tplTimer); /* keep what's already shown */ });
    }).catch(function () {});
  }
  // Keep the mobile section dropdown labels in sync with the live counts.
  function syncResTabSelect() {
    var sel = document.getElementById("resTabSelect");
    if (!sel) return;
    var map = { sops: "count-sops", templates: "count-templates", vendors: "count-vendors", links: "count-links" };
    Array.prototype.forEach.call(sel.options, function (o) {
      var base = o.getAttribute("data-label");
      if (!base) { base = o.textContent.replace(/\s*\(\d+\)\s*$/, ""); o.setAttribute("data-label", base); }
      var c = document.getElementById(map[o.value]);
      var n = c && c.textContent ? c.textContent : "";
      o.textContent = n ? (base + " (" + n + ")") : base;
    });
  }
  // Tabbed Resources page: show one section at a time (keeps the page short).
  function initResTabs() {
    var btns = document.querySelectorAll(".res-tab-btn");
    var sel = document.getElementById("resTabSelect");
    if (!btns.length && !sel) return;
    var valid = ["sops", "templates", "vendors", "links"];
    function show(name) {
      document.querySelectorAll(".res-tab-btn").forEach(function (b) {
        b.classList.toggle("active", b.getAttribute("data-tab") === name);
      });
      document.querySelectorAll(".res-tab").forEach(function (p) {
        p.classList.toggle("active", p.id === "tab-" + name);
      });
      if (sel && sel.value !== name) sel.value = name;
    }
    btns.forEach(function (b) {
      b.addEventListener("click", function () {
        var n = b.getAttribute("data-tab");
        show(n);
        try { history.replaceState(null, "", "#" + n); } catch (e) {}
      });
    });
    if (sel) sel.addEventListener("change", function () {
      show(sel.value);
      try { history.replaceState(null, "", "#" + sel.value); } catch (e) {}
    });
    syncResTabSelect();
    var hash = (location.hash || "").replace("#", "");
    if (valid.indexOf(hash) > -1) show(hash);
  }

  /* ---------- Unread email count (server-side via /api/my-mail) ---------- */
  function initMail() {
    var pill = document.getElementById("mailPill");
    var txt = document.getElementById("mailText");
    if (!pill) return; // pill always links to the mailbox even if the count can't load
    fetch("/api/my-mail").then(function (r) { return r.ok ? r.json() : null; }).then(function (j) {
      if (!j || typeof j.unread !== "number" || !txt) return;
      txt.textContent = j.unread + " unread";
      if (j.unread > 0) pill.classList.add("has-unread"); else pill.classList.remove("has-unread");
    }).catch(function () {});
  }

  /* ---------- Calendar: my events + company group calendar ---------- */
  // All-day company-calendar entries whose title matches this are shown as "time off"
  // in the Who's-off panel and tinted in the month grid. Adjust the words as needed.
  var TIMEOFF_RE = /\b(vacation|annual leave|on leave|leave|pto|o\.?o\.?o\.?|out of office|time off|day off|away|sick|parental|maternity|paternity)\b/i;

  // Events come from our server-side endpoint (mobile-safe): the signed-in person's
  // calendar + the company group calendar, already merged and tagged with __src.
  function fetchCalendar(start, end) {
    var u = "/api/my-calendar";
    if (start && end) u += "?start=" + encodeURIComponent(start.toISOString()) + "&end=" + encodeURIComponent(end.toISOString());
    return fetch(u).then(function (r) { return r.ok ? r.json() : { events: [] }; }).then(function (j) { return (j && j.events) || []; });
  }

  // Field crew shifts + who's off, from QuickBooks Time (server-side /api/crew-schedule).
  var CREW_LINK = "https://workforce.intuit.com";
  function fetchCrew(start, end) {
    var u = "/api/crew-schedule";
    if (start && end) u += "?start=" + encodeURIComponent(start.toISOString()) + "&end=" + encodeURIComponent(end.toISOString());
    // Retry a couple of times with a per-try timeout — the first call after the server has been
    // idle can be slow (cold start), so this rides through it instead of showing a false error.
    function attempt(triesLeft) {
      var ac = new AbortController();
      var timer = setTimeout(function () { ac.abort(); }, 22000);
      return fetch(u, { signal: ac.signal })
        .then(function (r) { clearTimeout(timer); if (!r.ok) throw new Error("http " + r.status); return r.json(); })
        .then(function (j) { return j || { configured: false, shifts: [], off: [], __failed: true }; })
        .catch(function () {
          clearTimeout(timer);
          if (triesLeft > 0) return new Promise(function (res) { setTimeout(res, 1500); }).then(function () { return attempt(triesLeft - 1); });
          return { configured: false, shifts: [], off: [], __failed: true }; // genuinely couldn't load (not the same as "not configured")
        });
    }
    return attempt(2);
  }
  // Represent a crew shift / a day off as calendar "events" so the grids can render them.
  function crewShiftToEvent(s) {
    var subj = s.name + (s.title ? " — " + s.title : "");
    return { subject: subj, start: { dateTime: s.start }, end: { dateTime: s.end }, isAllDay: !!s.allDay, __src: "crew", webLink: CREW_LINK, person: s.name || "", title: s.title || "", location: s.location || "", notes: s.notes || "" };
  }
  function crewOffToEvent(o) {
    var nx = new Date(parseYmd(o.date).getTime() + 86400000); // all-day span = the date through next midnight (end exclusive)
    return { subject: (o.name || "Someone") + " — " + (o.type || "off"), start: { dateTime: o.date + "T00:00:00" }, end: { dateTime: ymdAttr(nx) + "T00:00:00" }, isAllDay: true, __src: "off", webLink: CREW_LINK, person: o.name || "", offType: o.type || "" };
  }
  function parseYmd(s) { var m = String(s).match(/(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : new Date(s); }
  function startOfDay(d) { var x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
  function sameDay(a, b) { return startOfDay(a).getTime() === startOfDay(b).getTime(); }
  function ymdAttr(d) { return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function uniqStr(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  function setText(id, t) { var e = document.getElementById(id); if (e) e.textContent = t; }

  // ---------- Detail popup (shared) ----------
  function fmtWhen(e) {
    var s = e.start && e.start.dateTime ? new Date(e.start.dateTime) : null;
    var en = e.end && e.end.dateTime ? new Date(e.end.dateTime) : null;
    if (!s) return "";
    var dOpt = { weekday: "long", month: "long", day: "numeric", year: "numeric" };
    if (e.isAllDay) {
      var last = en ? new Date(en.getTime() - 86400000) : s;
      if (last.toDateString() !== s.toDateString())
        return s.toLocaleDateString("en-CA", dOpt) + " – " + last.toLocaleDateString("en-CA", dOpt) + " · All day";
      return s.toLocaleDateString("en-CA", dOpt) + " · All day";
    }
    var t = s.toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" });
    if (en) t += " – " + en.toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" });
    return s.toLocaleDateString("en-CA", dOpt) + " · " + t;
  }
  function srcLabel(src) { return src === "company" ? "Company calendar" : src === "crew" ? "Field crew" : src === "off" ? "Time off" : "Your calendar"; }
  function openDetail(e) {
    var m = document.getElementById("cal-modal");
    if (!m) { if (e.webLink) window.open(e.webLink, "_blank", "noopener"); return; }
    var rows = [];
    if (e.person) rows.push('<div class="cd-row"><span>Who</span><b>' + esc(e.person) + "</b></div>");
    if (e.offType) rows.push('<div class="cd-row"><span>Type</span><b>' + esc(e.offType) + "</b></div>");
    if (e.location) rows.push('<div class="cd-row"><span>Location</span><b>' + esc(e.location) + "</b></div>");
    if (e.notes) rows.push('<div class="cd-row"><span>Notes</span><b>' + esc(e.notes) + "</b></div>");
    var openLbl = (e.__src === "crew" || e.__src === "off") ? "Open in QuickBooks Time" : "Open in Outlook";
    m.querySelector(".cd-tag").className = "cd-tag cal-tag " + (srcMeta(e).cls);
    m.querySelector(".cd-tag").textContent = srcLabel(e.__src);
    m.querySelector(".cd-title").textContent = e.subject || "(no title)";
    m.querySelector(".cd-when").textContent = fmtWhen(e);
    m.querySelector(".cd-body").innerHTML = rows.join("") || '<div class="note" style="margin:0">No extra details.</div>';
    m.querySelector(".cd-foot").innerHTML = e.webLink ? '<a class="btn-outline" href="' + e.webLink + '" target="_blank" rel="noopener">' + openLbl + " →</a>" : "";
    m.classList.add("show"); m.setAttribute("aria-hidden", "false");
  }
  function closeDetail() { var m = document.getElementById("cal-modal"); if (m) { m.classList.remove("show"); m.setAttribute("aria-hidden", "true"); } }
  function wireModalClose() {
    var modal = document.getElementById("cal-modal");
    if (!modal || modal._wired) return; modal._wired = true;
    modal.addEventListener("click", function (e) { if (e.target === modal || (e.target.closest && e.target.closest(".cd-close"))) closeDetail(); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeDetail(); });
  }

  // ---------- Event registry: chips/rows reference events by index, clicks open the popup ----------
  var evReg = [];
  function regEv(e) { evReg.push(e); return evReg.length - 1; }
  function srcMeta(e) {
    var off = e.__src === "off" || (e.isAllDay && e.__src === "company" && TIMEOFF_RE.test(e.subject || ""));
    return { off: off, cls: off ? "off" : (e.__src || "me"), tag: e.__src === "company" ? "Company" : e.__src === "crew" ? "Crew" : off ? "Off" : "You" };
  }
  function evRow(e) {
    var s = new Date(e.start.dateTime);
    var time = e.isAllDay ? "All day" : s.toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" });
    var meta = srcMeta(e);
    return '<button type="button" class="cal-event ' + meta.cls + '" data-ev="' + regEv(e) + '">' +
      '<span class="cal-time">' + esc(time) + '</span><span class="cal-title">' + esc(e.subject || "(no title)") +
      '</span><span class="cal-tag">' + meta.tag + '</span></button>';
  }
  function evChip(e) {
    var meta = srcMeta(e);
    var t = e.isAllDay ? "" : new Date(e.start.dateTime).toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" }) + " ";
    return '<button type="button" class="cal-chip ' + meta.cls + '" data-ev="' + regEv(e) + '" title="' + esc(t + (e.subject || "")) + '">' + esc(t + (e.subject || "(no title)")) + "</button>";
  }
  function wireEvClicks(container) {
    if (!container || container._wired) return; container._wired = true;
    container.addEventListener("click", function (ev) {
      var b = ev.target.closest ? ev.target.closest("[data-ev]") : null;
      if (b && container.contains(b)) { ev.preventDefault(); var i = +b.getAttribute("data-ev"); if (evReg[i]) openDetail(evReg[i]); }
    });
  }
  function bucketByDay(events) {
    var byDay = {};
    (events || []).forEach(function (e) {
      if (!e.start || !e.start.dateTime) return;
      var s = new Date(e.start.dateTime);
      var endD = e.end && e.end.dateTime ? new Date(e.end.dateTime) : new Date(s.getTime() + 3600000);
      var d = new Date(s); d.setHours(0, 0, 0, 0);
      var last = new Date(e.isAllDay ? endD.getTime() - 1 : endD.getTime()); last.setHours(0, 0, 0, 0);
      for (var g = 0; d <= last && g < 60; g++) { (byDay[d.toDateString()] = byDay[d.toDateString()] || []).push(e); d = new Date(d.getTime() + 86400000); }
    });
    Object.keys(byDay).forEach(function (k) { byDay[k].sort(function (a, b) { return (a.isAllDay ? 0 : 1) - (b.isAllDay ? 0 : 1) || new Date(a.start.dateTime) - new Date(b.start.dateTime); }); });
    return byDay;
  }
  function dayEvents(events, day) {
    var d0 = startOfDay(day), d1 = new Date(d0.getTime() + 86400000);
    return (events || []).filter(function (e) {
      if (!e.start || !e.start.dateTime) return false;
      var s = new Date(e.start.dateTime), en = e.end && e.end.dateTime ? new Date(e.end.dateTime) : new Date(s.getTime() + 3600000);
      return s < d1 && en > d0;
    }).sort(function (a, b) { return (a.isAllDay ? 0 : 1) - (b.isAllDay ? 0 : 1) || new Date(a.start.dateTime) - new Date(b.start.dateTime); });
  }
  function overlapsDay(s, en, day) { var d0 = startOfDay(day), d1 = new Date(d0.getTime() + 86400000); en = en || new Date(s.getTime() + 3600000); return s < d1 && en > d0; }

  function initCalendar() {
    var elSched = document.getElementById("sched-view");
    if (elSched) { initSchedulePage(elSched); return; }
    var elView = document.getElementById("cal-view");
    if (elView) { initCalendarPage(elView); return; }
    var elToday = document.getElementById("calendar-today");
    var elSites = document.getElementById("site-visits");
    if (elToday || elSites) { // dashboard: today's events + today's site visits (loaded independently)
      var calEvents = null, crewData = null;
      function renderDash() {
        evReg.length = 0;
        if (elToday && calEvents !== null) {
          var ev = calEvents.concat(((crewData && crewData.shifts) || []).map(crewShiftToEvent));
          var t = dayEvents(ev, new Date());
          elToday.innerHTML = t.length ? t.map(evRow).join("") : '<div class="note" style="margin:0">No events today.</div>';
        }
        if (elSites && crewData !== null) renderCrewPanel(elSites, crewData, new Date(), "No site visits scheduled today.");
        wireEvClicks(elToday); wireEvClicks(elSites); wireModalClose();
      }
      fetchCalendar().then(function (ev) { calEvents = ev; renderDash(); }).catch(function () { calEvents = []; renderDash(); });
      fetchCrew().then(function (c) { crewData = c || { configured: false, shifts: [], off: [] }; renderDash(); }).catch(function () { crewData = { configured: false, shifts: [], off: [] }; renderDash(); });
    }
  }

  // Dedicated field-crew schedule page: Week / 2 Weeks / Month of crew shifts + time off only.
  function initSchedulePage(elView) {
    var view = "week", selected = startOfDay(new Date()), cache = {}, loadSeq = 0, lastCrew = null;
    var label = document.getElementById("sched-label");
    wireEvClicks(elView);
    function span() { return view === "fortnight" ? 14 : 7; }
    function rangeFor() {
      if (view === "month") { var first = new Date(selected.getFullYear(), selected.getMonth(), 1), g = new Date(first); g.setDate(1 - first.getDay()); return [g, new Date(g.getTime() + 42 * 86400000)]; }
      var w = new Date(selected); w.setDate(w.getDate() - w.getDay()); w.setHours(0, 0, 0, 0); return [w, new Date(w.getTime() + span() * 86400000)];
    }
    function setLabel() {
      if (!label) return;
      if (view === "month") { label.textContent = selected.toLocaleDateString("en-CA", { month: "long", year: "numeric" }); return; }
      var r = rangeFor(), last = new Date(r[1].getTime() - 86400000);
      label.textContent = r[0].toLocaleDateString("en-CA", { month: "short", day: "numeric" }) + " – " + last.toLocaleDateString("en-CA", { month: "short", day: "numeric", year: "numeric" });
    }
    function render() {
      evReg.length = 0; setLabel();
      if (lastCrew === null) { elView.innerHTML = '<div class="note" style="margin:0">Loading the schedule…</div>'; return; }
      if (lastCrew.__failed) { elView.innerHTML = '<div class="note" style="margin:0">Couldn\'t load the schedule right now — it can be slow to wake up. <a href="#" id="sched-retry">Try again</a>.</div>'; var rb = document.getElementById("sched-retry"); if (rb) rb.onclick = function (ev) { ev.preventDefault(); cache = {}; load(); }; return; }
      if (!lastCrew.configured) { elView.innerHTML = '<div class="note" style="margin:0">The team schedule isn\'t connected yet. Once QuickBooks Time is linked, shifts will appear here.</div>'; return; }
      var ev = (lastCrew.shifts || []).map(crewShiftToEvent).concat((lastCrew.off || []).map(crewOffToEvent));
      if (view === "month") { renderMonth(elView, selected, ev); return; }
      // Week / 2 weeks: stacked day-by-day list (readable on mobile and desktop).
      var w = new Date(selected); w.setDate(w.getDate() - w.getDay()); w.setHours(0, 0, 0, 0);
      renderDayList(elView, w, view === "fortnight" ? 14 : 7, selected, ev, "No site visits");
    }
    function load() {
      var r = rangeFor(), key = view + ":" + ymdAttr(r[0]), seq = ++loadSeq;
      if (cache[key]) { lastCrew = cache[key]; render(); return; }
      lastCrew = null; render();
      fetchCrew(r[0], r[1]).then(function (crew) { if (seq !== loadSeq) return; lastCrew = crew || { configured: false, shifts: [], off: [] }; cache[key] = lastCrew; render(); });
    }
    function shift(dir) {
      if (view === "month") selected = startOfDay(new Date(selected.getFullYear(), selected.getMonth() + dir, Math.min(selected.getDate(), 28)));
      else selected = startOfDay(new Date(selected.getTime() + dir * span() * 86400000));
      load();
    }
    var prev = document.getElementById("sched-prev"), next = document.getElementById("sched-next"), todayBtn = document.getElementById("sched-today");
    if (prev) prev.onclick = function () { shift(-1); };
    if (next) next.onclick = function () { shift(1); };
    if (todayBtn) todayBtn.onclick = function () { selected = startOfDay(new Date()); load(); };
    document.querySelectorAll(".cal-vbtn").forEach(function (b) {
      b.addEventListener("click", function () { view = b.getAttribute("data-view") || "week"; document.querySelectorAll(".cal-vbtn").forEach(function (x) { x.classList.toggle("active", x === b); }); load(); });
    });
    var modal = document.getElementById("cal-modal");
    if (modal) {
      modal.addEventListener("click", function (e) { if (e.target === modal || (e.target.closest && e.target.closest(".cd-close"))) closeDetail(); });
      document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeDetail(); });
    }
    load();
  }

  // Day / Week / Month controller with a selected date; the crew & who's-off panels follow it.
  function initCalendarPage(elView) {
    var view = "month", selected = startOfDay(new Date()), lastCal = null, lastCrew = null, cache = {}, loadSeq = 0;
    var elCrew = document.getElementById("crew-list"), elOff = document.getElementById("timeoff-list");
    var label = document.getElementById("cal-label");
    wireEvClicks(elView); wireEvClicks(elCrew); wireEvClicks(elOff);

    // Which calendars to show (persisted). Toggled via the legend.
    var FILT_KEY = "kepcm.calFilters";
    var visible = (function () { var d = { me: true, company: true, crew: true, off: true }; try { var j = JSON.parse(localStorage.getItem(FILT_KEY)); if (j) d = Object.assign(d, j); } catch (e) {} return d; })();
    function srcKey(e) { var s = e.__src || "me"; return (s === "company" || s === "crew" || s === "off") ? s : "me"; }
    function visibleEvents() { return combined().filter(function (e) { return visible[srcKey(e)] !== false; }); }
    document.querySelectorAll("#cal-legend .cal-legend-item").forEach(function (b) {
      var s = b.getAttribute("data-src");
      b.classList.toggle("hid", visible[s] === false);
      b.setAttribute("aria-pressed", visible[s] !== false ? "true" : "false");
      b.addEventListener("click", function () {
        visible[s] = (visible[s] === false);
        b.classList.toggle("hid", visible[s] === false);
        b.setAttribute("aria-pressed", visible[s] !== false ? "true" : "false");
        try { localStorage.setItem(FILT_KEY, JSON.stringify(visible)); } catch (e) {}
        render();
      });
    });

    function rangeFor() {
      if (view === "day") { var a = new Date(selected); return [a, new Date(a.getTime() + 86400000)]; }
      if (view === "week") { var w = new Date(selected); w.setDate(w.getDate() - w.getDay()); return [w, new Date(w.getTime() + 7 * 86400000)]; }
      var first = new Date(selected.getFullYear(), selected.getMonth(), 1), g = new Date(first); g.setDate(1 - first.getDay());
      return [g, new Date(g.getTime() + 42 * 86400000)];
    }
    function setLabel() {
      if (!label) return;
      if (view === "day") label.textContent = selected.toLocaleDateString("en-CA", { weekday: "long", month: "long", day: "numeric", year: "numeric" });
      else if (view === "week") { var r = rangeFor(), last = new Date(r[1].getTime() - 86400000); label.textContent = r[0].toLocaleDateString("en-CA", { month: "short", day: "numeric" }) + " – " + last.toLocaleDateString("en-CA", { month: "short", day: "numeric", year: "numeric" }); }
      else label.textContent = selected.toLocaleDateString("en-CA", { month: "long", year: "numeric" });
    }
    function combined() {
      var ev = (lastCal || []).slice();
      if (lastCrew) ev = ev.concat((lastCrew.shifts || []).map(crewShiftToEvent)).concat((lastCrew.off || []).map(crewOffToEvent));
      return ev;
    }
    function render() {
      evReg.length = 0; setLabel();
      if (lastCal === null) { elView.innerHTML = '<div class="note" style="margin:0">Loading your calendar…</div>'; }
      else {
        var ev = visibleEvents();
        if (view === "day") renderDay(elView, selected, ev);
        else if (view === "week") renderWeek(elView, selected, ev);
        else renderMonth(elView, selected, ev);
      }
      renderPanels();
    }
    function renderPanels() {
      var dtxt = sameDay(selected, new Date()) ? "today" : selected.toLocaleDateString("en-CA", { weekday: "short", month: "short", day: "numeric" });
      setText("crew-date-label", dtxt); setText("off-date-label", dtxt);
      if (lastCrew === null) {
        var ld = '<div class="note" style="margin:0">Loading…</div>';
        if (elCrew) elCrew.innerHTML = ld; if (elOff) elOff.innerHTML = ld;
        return;
      }
      if (elCrew) renderCrewPanel(elCrew, lastCrew, selected);
      if (elOff) renderOffPanel(elOff, lastCrew, selected, combined());
    }
    // Calendar and crew load independently and cache per range, so the grid shows as soon as
    // Outlook returns and the QuickBooks panels fill in when they're ready (and instantly on revisit).
    function load() {
      var r = rangeFor(), key = view + ":" + ymdAttr(r[0]), seq = ++loadSeq;
      cache[key] = cache[key] || {}; var c = cache[key];
      lastCal = c.cal || null; lastCrew = c.crew || null;
      render();
      if (!c.cal) fetchCalendar(r[0], r[1]).then(function (cal) { if (seq !== loadSeq) return; c.cal = cal; lastCal = cal; render(); })
        .catch(function () { if (seq === loadSeq && !lastCal) elView.innerHTML = '<div class="note" style="margin:0">Could not load the calendar.</div>'; });
      if (!c.crew) fetchCrew(r[0], r[1]).then(function (crew) { if (seq !== loadSeq) return; c.crew = crew || { configured: false, shifts: [], off: [] }; lastCrew = c.crew; render(); });
    }
    function shift(dir) {
      if (view === "day") selected = startOfDay(new Date(selected.getTime() + dir * 86400000));
      else if (view === "week") selected = startOfDay(new Date(selected.getTime() + dir * 7 * 86400000));
      else selected = startOfDay(new Date(selected.getFullYear(), selected.getMonth() + dir, Math.min(selected.getDate(), 28)));
      load();
    }
    var prev = document.getElementById("cal-prev"), next = document.getElementById("cal-next"), todayBtn = document.getElementById("cal-today");
    if (prev) prev.onclick = function () { shift(-1); };
    if (next) next.onclick = function () { shift(1); };
    if (todayBtn) todayBtn.onclick = function () { selected = startOfDay(new Date()); load(); };
    document.querySelectorAll(".cal-vbtn").forEach(function (b) {
      b.addEventListener("click", function () {
        view = b.getAttribute("data-view") || "month";
        document.querySelectorAll(".cal-vbtn").forEach(function (x) { x.classList.toggle("active", x === b); });
        load();
      });
    });
    // Click a day cell (not an event) to select it — panels follow; "+N more" jumps to Day view.
    elView.addEventListener("click", function (ev) {
      if (ev.target.closest("[data-ev]")) return;
      var cell = ev.target.closest ? ev.target.closest("[data-day]") : null;
      if (!cell || !elView.contains(cell)) return;
      selected = startOfDay(parseYmd(cell.getAttribute("data-day")));
      if (ev.target.closest(".cal-more")) { view = "day"; document.querySelectorAll(".cal-vbtn").forEach(function (x) { x.classList.toggle("active", x.getAttribute("data-view") === "day"); }); load(); return; }
      render();
    });
    var modal = document.getElementById("cal-modal");
    if (modal) {
      modal.addEventListener("click", function (e) { if (e.target === modal || (e.target.closest && e.target.closest(".cd-close"))) closeDetail(); });
      document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeDetail(); });
    }
    load();
  }

  var DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  // Sun–Sat day grid of `weeksN` rows starting at `gridStart`. Chips open the detail popup.
  // dimMonth (a month index) dims out-of-month cells; pass null to keep all bright.
  function renderGridDays(el, gridStart, weeksN, dimMonth, selected, events, maxChips) {
    var byDay = bucketByDay(events);
    var todayKey = new Date().toDateString(), selKey = selected ? selected.toDateString() : "";
    var mc = maxChips || 3;
    var html = '<div class="cal-grid"><div class="cal-grid-head">' + DOW.map(function (d) { return "<div>" + d + "</div>"; }).join("") + '</div><div class="cal-grid-body">';
    for (var i = 0; i < weeksN * 7; i++) {
      var day = new Date(gridStart.getTime() + i * 86400000), k = day.toDateString();
      var evs = byDay[k] || [];
      var cls = "cal-cell" + (dimMonth != null && day.getMonth() !== dimMonth ? " out" : "") + (k === todayKey ? " today" : "") + (selKey && k === selKey ? " sel" : "");
      var chips = evs.slice(0, mc).map(evChip).join("");
      var more = evs.length > mc ? '<span class="cal-more">+' + (evs.length - mc) + " more</span>" : "";
      html += '<div class="' + cls + '" data-day="' + ymdAttr(day) + '"><div class="cal-cell-date">' + day.getDate() + "</div>" + chips + more + "</div>";
    }
    el.innerHTML = html + "</div></div>";
  }
  function renderMonth(el, selected, events) {
    var cursor = new Date(selected.getFullYear(), selected.getMonth(), 1), g = new Date(cursor); g.setDate(1 - cursor.getDay());
    renderGridDays(el, g, 6, selected.getMonth(), selected, events, 3);
  }
  // Two-week grid (the week of `selected` + the next), taller cells so more shifts show.
  function renderFortnight(el, selected, events) {
    var w = new Date(selected); w.setDate(w.getDate() - w.getDay()); w.setHours(0, 0, 0, 0);
    renderGridDays(el, w, 2, null, selected, events, 8);
  }
  // Stacked day-by-day list (mobile-friendly): each day is a section with its events as rows.
  function renderDayList(el, gridStart, numDays, selected, events, emptyMsg) {
    var todayKey = new Date().toDateString(), selKey = selected ? selected.toDateString() : "";
    var html = "";
    for (var i = 0; i < numDays; i++) {
      var day = new Date(gridStart.getTime() + i * 86400000), k = day.toDateString();
      var evs = dayEvents(events, day);
      var cls = "cal-day cal-stack-day" + (k === todayKey ? " today" : "") + (selKey && k === selKey ? " sel" : "");
      html += '<div class="' + cls + '" data-day="' + ymdAttr(day) + '">' +
        '<div class="cal-date">' + day.toLocaleDateString("en-CA", { weekday: "short", month: "short", day: "numeric" }) + (k === todayKey ? " · Today" : "") + "</div>" +
        (evs.length ? evs.map(evRow).join("") : '<div class="cal-empty-day">' + (emptyMsg || "Nothing scheduled") + "</div>") + "</div>";
    }
    el.innerHTML = html;
  }
  function fmtHour(h) { var ap = h < 12 ? "AM" : "PM"; var hh = h % 12; if (hh === 0) hh = 12; return hh + " " + ap; }
  // Lay out overlapping events into columns, then let each expand right to fill free space.
  // Sets .left and .width (0..1) on each item. Items have {_s,_e} in minutes.
  function layoutColumns(items) {
    items.sort(function (a, b) { return a._s - b._s || b._e - a._e; });
    var cluster = [], clusterMaxEnd = -1;
    function finalize(cl) {
      var colEnds = [];
      cl.forEach(function (e) {
        var placed = false;
        for (var i = 0; i < colEnds.length; i++) { if (e._s >= colEnds[i]) { e.col = i; colEnds[i] = e._e; placed = true; break; } }
        if (!placed) { e.col = colEnds.length; colEnds.push(e._e); }
      });
      var ncol = colEnds.length;
      cl.forEach(function (e) {
        var span = 1;
        for (var c = e.col + 1; c < ncol; c++) {
          var hit = cl.some(function (o) { return o !== e && o.col === c && o._s < e._e && o._e > e._s; });
          if (hit) break; span++;
        }
        e.left = e.col / ncol; e.width = span / ncol;
      });
    }
    items.forEach(function (e) {
      if (cluster.length && e._s >= clusterMaxEnd) { finalize(cluster); cluster = []; clusterMaxEnd = -1; }
      cluster.push(e); clusterMaxEnd = Math.max(clusterMaxEnd, e._e);
    });
    if (cluster.length) finalize(cluster);
    return items;
  }
  // Time-grid: hours down the Y axis, `numDays` day columns from `gridStart`, events
  // positioned & sized by time (overlaps split side-by-side), all-day items in a top band.
  function renderTimeGrid(el, gridStart, numDays, selected, events) {
    var todayKey = new Date().toDateString(), selKey = selected.toDateString();
    var days = []; for (var i = 0; i < numDays; i++) days.push(new Date(gridStart.getTime() + i * 86400000));
    var timedByDay = [], alldayByDay = [], anyAllday = false, minM = 8 * 60, maxM = 18 * 60;
    days.forEach(function (day) {
      var d0 = startOfDay(day), d1 = new Date(d0.getTime() + 86400000), timed = [], allday = [];
      (events || []).forEach(function (e) {
        if (!e.start || !e.start.dateTime) return;
        var s = new Date(e.start.dateTime), en = e.end && e.end.dateTime ? new Date(e.end.dateTime) : new Date(s.getTime() + 3600000);
        if (s >= d1 || en <= d0) return;
        if (e.isAllDay) { allday.push(e); return; }
        var sm = Math.max(0, (s - d0) / 60000), em = Math.min(1440, (en - d0) / 60000); if (em <= sm) em = sm + 30;
        timed.push({ e: e, _s: sm, _e: em });
        minM = Math.min(minM, Math.floor(sm / 60) * 60); maxM = Math.max(maxM, Math.ceil(em / 60) * 60);
      });
      if (allday.length) anyAllday = true;
      timedByDay.push(timed); alldayByDay.push(allday);
    });
    minM = Math.max(0, minM); maxM = Math.min(1440, maxM);
    var hourH = 46, startH = Math.floor(minM / 60), endH = Math.ceil(maxM / 60), gridH = (endH - startH) * hourH;

    var head = '<div class="cw-row cw-head"><div class="cw-gutter"></div>';
    days.forEach(function (day) { var k = day.toDateString(); head += '<div class="cw-dayhead' + (k === todayKey ? " today" : "") + (k === selKey ? " sel" : "") + '" data-day="' + ymdAttr(day) + '">' + DOW[day.getDay()] + " " + day.getDate() + "</div>"; });
    head += "</div>";

    var band = "";
    if (anyAllday) {
      band = '<div class="cw-row cw-allday"><div class="cw-gutter">all-day</div>';
      days.forEach(function (day, di) { band += '<div class="cw-allcell" data-day="' + ymdAttr(day) + '">' + alldayByDay[di].map(evChip).join("") + "</div>"; });
      band += "</div>";
    }

    var gutter = '<div class="cw-gutter cw-times">';
    for (var h = startH; h < endH; h++) gutter += '<div class="cw-hour" style="height:' + hourH + 'px">' + fmtHour(h) + "</div>";
    gutter += "</div>";

    var cols = "";
    days.forEach(function (day, di) {
      var k = day.toDateString();
      var col = '<div class="cw-col' + (k === selKey ? " sel" : "") + (k === todayKey ? " today" : "") + '" data-day="' + ymdAttr(day) + '" style="height:' + gridH + 'px">';
      for (var h = startH; h < endH; h++) col += '<div class="cw-line" style="top:' + ((h - startH) * hourH) + 'px"></div>';
      var ev = layoutColumns(timedByDay[di].slice());
      ev.forEach(function (it) {
        var top = (it._s - startH * 60) / 60 * hourH, hgt = Math.max(20, (it._e - it._s) / 60 * hourH - 2);
        var left = it.left * 100, wpct = it.width * 100, meta = srcMeta(it.e);
        var st = new Date(it.e.start.dateTime).toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" });
        col += '<button type="button" class="cw-ev ' + meta.cls + '" data-ev="' + regEv(it.e) + '" style="top:' + top + "px;height:" + hgt + "px;left:calc(" + left + "% + 1px);width:calc(" + wpct + '% - 2px)" title="' + esc(st + " " + (it.e.subject || "")) + '"><span class="cw-evt">' + esc(st) + "</span> " + esc(it.e.subject || "(no title)") + "</button>";
      });
      col += "</div>"; cols += col;
    });

    el.innerHTML = '<div class="cal-week2' + (numDays > 1 ? " cw-multi" : " cw-single") + '">' + head + band + '<div class="cw-row cw-grid">' + gutter + cols + "</div></div>";
  }
  function renderWeek(el, selected, events) {
    var w = new Date(selected); w.setDate(w.getDate() - w.getDay()); w.setHours(0, 0, 0, 0);
    renderTimeGrid(el, w, 7, selected, events);
  }
  // Day view: same time-grid, a single full-width day column.
  function renderDay(el, selected, events) {
    renderTimeGrid(el, startOfDay(selected), 1, selected, events);
  }

  // Field crew shifts for the selected day (QuickBooks Time). Each row opens the detail popup.
  function renderCrewPanel(el, crew, selected, emptyMsg) {
    if (!el) return;
    if (crew.__failed) { el.innerHTML = '<div class="note" style="margin:0">Couldn\'t load site visits right now — it can be slow to wake up. Refresh to try again.</div>'; return; }
    if (!crew.configured) { el.innerHTML = '<div class="note" style="margin:0">The team schedule isn\'t connected yet. Once QuickBooks Time is linked, shifts will appear here.</div>'; return; }
    var shifts = (crew.shifts || []).filter(function (s) { return s.start && overlapsDay(new Date(s.start), s.end ? new Date(s.end) : null, selected); })
      .sort(function (a, b) { return new Date(a.start) - new Date(b.start); });
    if (!shifts.length) { el.innerHTML = '<div class="note" style="margin:0">' + (emptyMsg || "No crew shifts on this day.") + "</div>"; return; }
    el.innerHTML = shifts.map(function (s) {
      var st = new Date(s.start);
      var time = s.allDay ? "All day" : st.toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" }) + (s.end ? "–" + new Date(s.end).toLocaleTimeString("en-CA", { hour: "numeric", minute: "2-digit" }) : "");
      var label = (s.title || "Shift") + (s.location ? " · " + s.location : "");
      return '<button type="button" class="cal-event crew" data-ev="' + regEv(crewShiftToEvent(s)) + '"><span class="cal-time">' + esc(time) + '</span><span class="cal-title">' + esc(s.name + " — " + label) + '</span><span class="cal-tag">Crew</span></button>';
    }).join("");
  }
  // Who's off on the selected day. Uses QuickBooks Time; falls back to company-calendar matching.
  function renderOffPanel(el, crew, selected, events) {
    if (!el) return;
    var rows = [];
    if (crew.configured) {
      var selKey = ymdAttr(selected);
      (crew.off || []).forEach(function (o) { if (o.date === selKey) rows.push({ name: o.name || "Someone", type: o.type || "", ev: crewOffToEvent(o) }); });
    } else {
      dayEvents(events, selected).forEach(function (e) { if (e.isAllDay && e.__src === "company" && TIMEOFF_RE.test(e.subject || "")) rows.push({ name: e.subject || "Time off", type: "", ev: e }); });
    }
    var seen = {}, uniq = [];
    rows.forEach(function (r) { var k = r.name + "|" + r.type; if (!seen[k]) { seen[k] = 1; uniq.push(r); } });
    if (!uniq.length) { el.innerHTML = '<div class="note" style="margin:0">No one is off on this day.</div>'; return; }
    el.innerHTML = '<ul class="holiday-list">' + uniq.map(function (r) {
      return '<li data-ev="' + regEv(r.ev) + '" class="off-row"><span class="hname">' + esc(r.name) + '</span><span class="hdate">' + esc(r.type || "Off") + '</span></li>';
    }).join("") + "</ul>";
  }
  /* ---------- KRA & quarterly development review ----------
     Files live in a SharePoint folder staff can't access directly. The /api/my-reviews
     function reads the signed-in person's folder server-side (app identity) and returns
     download links, so staff can view/download here but not edit or browse SharePoint. */
  function fmtBytes(n) {
    n = +n || 0;
    if (n >= 1048576) return (n / 1048576).toFixed(1) + " MB";
    if (n >= 1024) return Math.round(n / 1024) + " KB";
    return n + " B";
  }
  function initReviews() {
    var el = document.getElementById("review-files");
    if (!el) return;
    function render(d) {
      var files = (d && d.files) || [];
      if (d && d.notFound) {
        el.innerHTML = '<div class="note" style="margin:0">We couldn\'t find a review folder under your name yet. Email <a href="mailto:hr@kingepcm.com?subject=KRA%20%2F%20development%20review">hr@kingepcm.com</a> to set it up.</div>';
        return;
      }
      if (!files.length) {
        el.innerHTML = '<div class="note" style="margin:0">No review documents yet. New reviews will appear here once added.</div>';
        return;
      }
      el.innerHTML = '<ul class="review-list">' + files.map(function (f) {
        var meta = [f.folder, fmtDate(f.modified), fmtBytes(f.size)].filter(Boolean).join(" · ");
        var dl = f.download
          ? '<a class="review-dl" href="' + esc(f.download) + '" download rel="noopener">⬇ Download</a>'
          : '<span class="review-dl" style="color:var(--muted)">Unavailable</span>';
        return '<li class="review-row"><span class="review-file"><span class="review-ic">📄</span>' +
          '<span><span class="review-name">' + esc(f.name) + '</span>' +
          '<span class="review-meta">' + esc(meta) + '</span></span></span>' + dl + '</li>';
      }).join("") + "</ul>";
    }
    var cached = cacheGet("reviews");
    if (cached) render(cached); // instant from cache (download links stay valid ~1h)
    fetch("/api/my-reviews")
      .then(function (r) {
        if (r.status === 501) throw "notset";
        return r.ok ? r.json() : Promise.reject();
      })
      .then(function (d) { cacheSet("reviews", d); render(d); })
      .catch(function (e) {
        if (cached) return;
        el.innerHTML = e === "notset"
          ? '<div class="note" style="margin:0">Review documents aren\'t connected yet. <a href="mailto:it@kingepcm.com?subject=Reviews%20setup">IT can finish setup</a>.</div>'
          : '<div class="note" style="margin:0">Couldn\'t load your reviews right now. Try again later, or email <a href="mailto:hr@kingepcm.com">hr@kingepcm.com</a>.</div>';
      });
  }

  /* ---------- Reveal manager-only links (Form C) for manager/admin roles ---------- */
  function initManagerLinks() {
    var link = document.getElementById("form-c-link");      // manager-only (Form C)
    if (!link) return;
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("manager") > -1 || roles.indexOf("admin") > -1) link.hidden = false;
    }).catch(function () {});
  }

  /* ---------- HR Portal: inject an HR/admin-only nav item on every page ---------- */
  function initHrPortalNav() {
    var sidebar = document.getElementById("sidebar");
    if (!sidebar || document.getElementById("hr-portal-nav")) return;
    var menu = sidebar.querySelector(".navgroup"); // first group = "Menu"
    if (!menu) return;
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("hr") === -1 && roles.indexOf("admin") === -1) return;
      var hint = document.getElementById("hr-portal-hint");
      if (hint) hint.hidden = false;
      if (document.getElementById("hr-portal-nav")) return;
      var a = document.createElement("a");
      a.className = "item";
      a.id = "hr-portal-nav";
      a.href = "hr-portal.html";
      a.innerHTML = '<span class="ic">🗝️</span> HR Portal';
      var page = (location.pathname.split("/").pop() || "").toLowerCase();
      if (page === "hr-portal.html") a.classList.add("active");
      var hrItem = menu.querySelector('a[href="hr.html"]');
      if (hrItem && hrItem.nextSibling) menu.insertBefore(a, hrItem.nextSibling);
      else menu.appendChild(a);
    }).catch(function () {});
  }

  /* ---------- HR Portal page: gate content to hr/admin ---------- */
  function initHrPortal() {
    var content = document.getElementById("hrp-content");
    if (!content) return;
    var gate = document.getElementById("hrp-gate");
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("hr") > -1 || roles.indexOf("admin") > -1) content.hidden = false;
      else if (gate) gate.hidden = false;
    }).catch(function () { if (gate) gate.hidden = false; });
  }

  /* ---------- HR: generate a branded employment-verification letter (hr/admin) ---------- */
  function initEmploymentLetter() {
    var form = document.getElementById("elForm");
    if (!form) return;
    var gate = document.getElementById("el-gate");
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("hr") > -1 || roles.indexOf("admin") > -1) form.hidden = false;
      else if (gate) gate.hidden = false;
    }).catch(function () { if (gate) gate.hidden = false; });

    var statusEl = document.getElementById("el-status");
    var submitEl = document.getElementById("el-submit");
    var last = null, lastUrl = null;
    function val(id) { var el = document.getElementById(id); return el ? (el.value || "") : ""; }
    function collect() {
      var name = val("el-name").trim();
      if (!name) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Employee legal name is required."; return null; }
      var ohip = !!(document.getElementById("el-ohip") && document.getElementById("el-ohip").checked);
      if (ohip && (!val("el-position").trim() || !val("el-start"))) { statusEl.style.color = "#C0392B"; statusEl.textContent = "OHIP letters must include the position and start date."; return null; }
      return {
        name: name, position: val("el-position").trim(), dateOfBirth: val("el-dob"), employmentType: val("el-etype"),
        startDate: val("el-start"), salaryAmount: val("el-salary").trim(), salaryBasis: val("el-basis"),
        ohip: ohip, employeeAddress: val("el-address").trim(),
        signedByName: val("el-signname").trim(), signedByTitle: val("el-signtitle").trim(),
        signedByPhone: val("el-signphone").trim(), signedByEmail: val("el-signemail").trim()
      };
    }
    function request(busy) {
      if (submitEl.disabled) return Promise.resolve(null); // a request is already in flight
      var payload = collect(); if (!payload) return Promise.resolve(null);
      submitEl.disabled = true; statusEl.style.color = ""; statusEl.textContent = busy;
      return fetch("/api/employment-letter", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
        .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
        .then(function (res) {
          submitEl.disabled = false;
          if (res.ok && res.j && res.j.ok) { last = res.j; return res.j; }
          statusEl.style.color = "#C0392B"; statusEl.textContent = (res.j && res.j.error === "HR only") ? "You need the HR role to generate letters." : ((res.j && res.j.error) || "Couldn't generate the letter. Please try again."); return null;
        }).catch(function () { submitEl.disabled = false; statusEl.style.color = "#C0392B"; statusEl.textContent = "Couldn't generate the letter. Please try again."; return null; });
    }
    function filedMsg(j) { return "Letter generated." + (j.filed ? " Filed to the staff's HR folder (" + (j.filedTo || "matched folder") + ")." : " (Downloaded only — couldn't auto-match a staff folder to file it.)"); }
    function showPreview(j) {
      if (lastUrl) { try { URL.revokeObjectURL(lastUrl); } catch (e) {} }
      lastUrl = b64ToBlobUrl(j.pdfBase64, "application/pdf");
      var fr = document.getElementById("el-preview-frame"); if (fr) fr.src = lastUrl;
      var wrap = document.getElementById("el-preview-wrap"); if (wrap) { wrap.hidden = false; if (wrap.scrollIntoView) wrap.scrollIntoView({ behavior: "smooth", block: "start" }); }
    }
    var prev = document.getElementById("el-preview");
    if (prev) prev.addEventListener("click", function () { request("Generating preview…").then(function (j) { if (!j) return; showPreview(j); statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "Preview ready below — review, then download."; }); });
    var docxBtn = document.getElementById("el-submit-docx");
    if (docxBtn) docxBtn.addEventListener("click", function () { request("Generating Word document…").then(function (j) { if (!j) return; downloadBase64File(j.docxBase64, j.docxFilename || "employment-letter.docx", DOCX_MIME); statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = filedMsg(j); }); });
    var pPdf = document.getElementById("el-preview-pdf");
    if (pPdf) pPdf.addEventListener("click", function () { if (last) downloadBase64Pdf(last.pdfBase64, last.filename || "employment-letter.pdf"); });
    var pDocx = document.getElementById("el-preview-docx");
    if (pDocx) pDocx.addEventListener("click", function () { if (last) downloadBase64File(last.docxBase64, last.docxFilename || "employment-letter.docx", DOCX_MIME); });
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      request("Generating…").then(function (j) { if (!j) return; downloadBase64Pdf(j.pdfBase64, j.filename || "employment-letter.pdf"); statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = filedMsg(j); });
    });
  }
  function b64ToBlobUrl(b64, mime) {
    var bin = atob(b64), len = bin.length, arr = new Uint8Array(len);
    for (var i = 0; i < len; i++) arr[i] = bin.charCodeAt(i);
    return URL.createObjectURL(new Blob([arr], { type: mime || "application/pdf" }));
  }
  function downloadBase64Pdf(b64, filename) { downloadBase64File(b64, filename, "application/pdf"); }
  function downloadBase64File(b64, filename, mime) {
    try {
      var bin = atob(b64), len = bin.length, arr = new Uint8Array(len);
      for (var i = 0; i < len; i++) arr[i] = bin.charCodeAt(i);
      var url = URL.createObjectURL(new Blob([arr], { type: mime || "application/octet-stream" }));
      var a = document.createElement("a"); a.href = url; a.download = filename; document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
    } catch (e) {}
  }
  var DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  /* ---------- Staff name pickers — choose from the Microsoft directory ----------
     Turns the name fields on the forms into a type-to-choose list of current staff
     (members with a mailbox; shared mailboxes excluded). Falls back to plain text
     if the directory isn't available, so a brand-new hire can still be typed in. */
  function initStaffPickers() {
    var singles = [].slice.call(document.querySelectorAll('#el-name, #sal-name, #rf-name-input'));
    var pairFirsts = [].slice.call(document.querySelectorAll('[data-meta="subject-first"]'));
    if (!singles.length && !pairFirsts.length) return;

    fetch("/api/staff-directory").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var staff = (d && d.staff) || [];
      if (!staff.length) return;
      var dl = document.getElementById("kepcm-staff");
      if (!dl) { dl = document.createElement("datalist"); dl.id = "kepcm-staff"; document.body.appendChild(dl); }
      dl.innerHTML = staff.map(function (s) { return '<option value="' + esc(s.name) + '">' + (s.title ? esc(s.title) : "") + "</option>"; }).join("");
      var byName = {}; staff.forEach(function (s) { byName[s.name.toLowerCase()] = s; });

      // Single full-name fields → attach the shared list.
      singles.forEach(function (inp) {
        inp.setAttribute("list", "kepcm-staff");
        inp.setAttribute("autocomplete", "off");
        inp.setAttribute("placeholder", "Start typing to choose a staff name…");
      });

      // First/last pairs → add one picker above them that fills both fields.
      pairFirsts.forEach(function (first) {
        var pair = first.closest(".rf-namepair"); if (!pair || pair._picked) return;
        var last = pair.querySelector('[data-meta="subject-last"]'); if (!last) return;
        pair._picked = true;
        var pick = document.createElement("input");
        pick.className = "rf-input";
        pick.setAttribute("list", "kepcm-staff");
        pick.setAttribute("autocomplete", "off");
        pick.placeholder = "Start typing to choose a staff name…";
        pick.style.marginBottom = "8px";
        pair.parentNode.insertBefore(pick, pair);
        pick.addEventListener("change", function () {
          var m = byName[(pick.value || "").trim().toLowerCase()];
          if (m) { var parts = m.name.split(/\s+/); first.value = parts.shift(); last.value = parts.join(" "); }
        });
      });
    }).catch(function () {});
  }

  /* ---------- Performance-review forms (A–D) — generic collect + submit ---------- */
  function initReviewForm() {
    var form = document.getElementById("reviewForm");
    if (!form) return;
    // Manager-only forms (Form C): hide the form for non-managers with a friendly note.
    if (form.hasAttribute("data-manager-only")) {
      fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
        var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
        if (roles.indexOf("manager") === -1 && roles.indexOf("admin") === -1) {
          var gate = document.getElementById("rf-gate");
          if (gate) gate.hidden = false;
          form.hidden = true;
        }
      }).catch(function () {});
    }
    // A scored row is "numeric" (1–5, counts toward the average) unless its scale
    // declares its own data-options (a single-choice picker, e.g. autonomy rung).
    function isNumericScore(row) { return !row.querySelector(".rf-scale[data-options]"); }
    // Build the score scales: 1–5 by default, or custom options from data-options.
    form.querySelectorAll(".rf-scale").forEach(function (sc) {
      var nm = sc.getAttribute("data-name");
      var opts = sc.getAttribute("data-options");
      var html = "";
      if (opts) {
        sc.classList.add("rf-scale-opts");
        opts.split("|").forEach(function (o) {
          o = o.trim();
          html += '<label class="rf-opt"><input type="radio" name="' + nm + '" value="' + esc(o) + '"><span>' + esc(o) + "</span></label>";
        });
      } else {
        for (var i = 1; i <= 5; i++) html += '<label><input type="radio" name="' + nm + '" value="' + i + '"><span>' + i + "</span></label>";
      }
      sc.innerHTML = html;
    });
    // Live per-section average for any section with 1–5 numeric scores.
    form.querySelectorAll(".rf-section").forEach(function (sec) {
      var hasNumeric = false;
      sec.querySelectorAll(".rf-score").forEach(function (r) { if (isNumericScore(r)) hasNumeric = true; });
      if (hasNumeric) {
        var avg = document.createElement("div");
        avg.className = "rf-avg";
        avg.innerHTML = 'Section average <strong>—</strong>';
        sec.appendChild(avg);
        sec._avgEl = avg;
      }
    });
    function recompute() {
      form.querySelectorAll(".rf-section").forEach(function (sec) {
        if (!sec._avgEl) return;
        var total = 0, sum = 0, n = 0;
        sec.querySelectorAll(".rf-score").forEach(function (r) {
          if (!isNumericScore(r)) return;
          total++;
          var c = r.querySelector("input:checked"); if (c) { sum += +c.value; n++; }
        });
        sec._avgEl.querySelector("strong").textContent = n ? (sum / n).toFixed(1) + (n < total ? " (" + n + " of " + total + ")" : "") : "—";
      });
    }
    form.addEventListener("change", recompute);

    // Show the signed-in name (banner span and/or a Name input on the form).
    var nameEl = document.getElementById("rf-name");
    var nameInput = document.getElementById("rf-name-input");
    if (nameEl || nameInput) fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var p = d && d.clientPrincipal; if (!p) return; var n = nameFromClaims(p).display; if (!n) return;
      if (nameEl) nameEl.textContent = n;
      if (nameInput && !nameInput.value) nameInput.value = n;
    }).catch(function () {});

    var statusEl = document.getElementById("rf-status");
    var submitEl = document.getElementById("rf-submit");
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var quarter = (form.querySelector('[data-meta="quarter"]') || {}).value || "";
      // Subject = the staff member a manager/growth form is about. Either a single
      // [data-meta="subject"] field, or a required first + last name pair (cleaner matching).
      var subjFirst = form.querySelector('[data-meta="subject-first"]');
      var subjLast = form.querySelector('[data-meta="subject-last"]');
      var subjSingle = form.querySelector('[data-meta="subject"]');
      var subjectName = "";
      if (subjFirst || subjLast) {
        var fv = ((subjFirst && subjFirst.value) || "").trim();
        var lv = ((subjLast && subjLast.value) || "").trim();
        if (!fv || !lv) {
          statusEl.style.color = "#C0392B"; statusEl.textContent = "Please enter the staff member's first and last name before submitting.";
          (fv ? subjLast : subjFirst).focus(); return;
        }
        subjectName = fv + " " + lv;
      } else if (subjSingle) {
        subjectName = (subjSingle.value || "").trim();
        if (!subjectName) {
          statusEl.style.color = "#C0392B"; statusEl.textContent = "Please enter the team member's name before submitting.";
          subjSingle.focus(); return;
        }
      }
      var sections = [], missing = 0;
      form.querySelectorAll(".rf-section").forEach(function (sec) {
        var heading = sec.getAttribute("data-section"); if (!heading) return;
        var rows = [], numRows = [];
        sec.querySelectorAll("[data-field]").forEach(function (f) {
          var label = f.getAttribute("data-label") || "";
          var val = "", row;
          if (f.classList.contains("rf-score")) {
            var checked = f.querySelector("input:checked");
            val = checked ? checked.value : "";
            var scale = f.querySelector(".rf-scale");
            var optAttr = scale && scale.getAttribute("data-options");
            if (optAttr) {
              // single-choice picker (e.g. autonomy rung) — keep the option list for the PDF
              row = { label: label, value: val, type: "option", options: optAttr.split("|").map(function (o) { return o.trim(); }) };
            } else {
              row = { label: label, value: val, type: "score", max: 5 };
              numRows.push(f); if (!val) missing++;
            }
          } else {
            val = (f.value || "").trim();
            row = { label: label, value: val, type: "text" };
          }
          rows.push(row);
        });
        // Add the auto-calculated average for numeric (1–5) scored sections only.
        if (numRows.length) {
          var sum = 0; numRows.forEach(function (r) { var c = r.querySelector("input:checked"); if (c) sum += +c.value; });
          rows.push({ label: "Section average", value: (sum / numRows.length).toFixed(1), type: "avg", count: numRows.length });
        }
        if (rows.length) sections.push({ heading: heading, rows: rows });
      });
      if (missing) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Please score every item (1–5) before submitting."; return; }
      // Faithful instruction line for the PDF: the page's own lead paragraph,
      // minus the trailing "Completing as …" personalisation.
      var intro = "";
      var leadP = document.querySelector(".pagehead p");
      if (leadP) { intro = (leadP.textContent || "").replace(/\s*Completing as.*$/i, "").replace(/\s+/g, " ").trim(); }
      // 1–5 rating criteria, pulled from the form's scoring guide, so the PDF
      // records what each score meant for future reference.
      var scale = [];
      document.querySelectorAll(".rf-guide table.rf-table tr").forEach(function (tr) {
        var tds = tr.querySelectorAll("td");
        if (tds.length >= 2) {
          var k = (tds[0].textContent || "").trim();
          if (/^[1-5]$/.test(k) && !scale.some(function (s) { return s.n === k; })) {
            scale.push({ n: k, desc: (tds[1].textContent || "").replace(/\s+/g, " ").trim() });
          }
        }
      });
      var payload = { form: form.getAttribute("data-form") || "", title: form.getAttribute("data-title") || "Review", version: form.getAttribute("data-version") || "", intro: intro, scale: scale, quarter: quarter, date: new Date().toLocaleDateString("en-CA"), sections: sections };
      if (subjectName) payload.subjectName = subjectName;
      submitEl.disabled = true; statusEl.style.color = ""; statusEl.textContent = "Submitting…";
      fetch("/api/review-submit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
        .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
        .then(function (res) {
          submitEl.disabled = false;
          if (res.ok && res.j && res.j.ok) {
            form.querySelectorAll('input[type="radio"]').forEach(function (x) { x.checked = false; });
            form.querySelectorAll("textarea,.rf-input").forEach(function (x) { if (x.tagName === "TEXTAREA" || x.getAttribute("data-meta")) x.value = ""; });
            statusEl.style.color = "var(--gold-dark)";
            statusEl.textContent = "Submitted — thank you." + (res.j.filed ? " Filed to the review folder." : "") + (res.j.emailed ? " A copy was emailed." : "");
          } else {
            statusEl.style.color = "#C0392B";
            statusEl.textContent = "Couldn't submit. Please try again or email hr@kingepcm.com.";
          }
        })
        .catch(function () { submitEl.disabled = false; statusEl.style.color = "#C0392B"; statusEl.textContent = "Couldn't submit. Please try again later."; });
    });
  }

  /* ---------- My profile (Microsoft 365 directory info) ---------- */
  function initProfile() {
    var el = document.getElementById("profile-info");
    if (!el) return;
    fetch("/api/my-profile").then(function (r) { return r.ok ? r.json() : null; })
      .then(function (me) {
        if (!me || !me.displayName) { el.innerHTML = '<div class="note" style="margin:0">Could not load your profile right now.</div>'; return; }
        var phone = me.mobilePhone || (me.businessPhones && me.businessPhones[0]) || "";
        var rows = [
          ["Name", me.displayName], ["Job title", me.jobTitle], ["Department", me.department],
          ["Email", me.mail], ["Phone", phone], ["Office", me.officeLocation], ["Employee ID", me.employeeId]
        ].filter(function (r) { return r[1]; });
        el.innerHTML = rows.map(function (r) {
          return '<li><span class="hname">' + esc(r[0]) + '</span><span class="hdate">' + esc(r[1]) + "</span></li>";
        }).join("");
      }).catch(function () { el.innerHTML = '<div class="note" style="margin:0">Could not load your profile right now. Try again shortly.</div>'; });
  }

  /* ---------- Reimbursement claim history (expense.html + profile.html) ---------- */
  var _profClaimedThisYear = 0; // PEO/PEng fees claimed in the current calendar year
  function loadClaimHistory() {
    var el = document.getElementById("reimb-history");
    var allowEl = document.getElementById("prof-allowance");
    if (!el && !allowEl) return;
    fetch("/api/my-reimbursements").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var c = (d && d.claims) || [];
      // Professional fees claimed this calendar year (for the $500 allowance display).
      var yr = String(new Date().getFullYear()), profSum = 0;
      c.forEach(function (x) {
        if (x.type === "professional" && String(x.createdAt || "").slice(0, 4) === yr) {
          var nval = parseFloat(String(x.total || "").replace(/[^0-9.]/g, "")); if (!isNaN(nval)) profSum += nval;
        }
      });
      _profClaimedThisYear = profSum;
      if (allowEl) {
        var rem = Math.max(0, 500 - profSum);
        allowEl.innerHTML = "Claimed this year: <strong>$" + profSum.toFixed(2) + "</strong> of $500 · Remaining: <strong>$" + rem.toFixed(2) + "</strong>";
      }
      if (el) {
        if (!c.length) { el.innerHTML = '<div class="note" style="margin:0">No claims submitted yet.</div>'; return; }
        el.innerHTML = '<ul class="review-list">' + c.map(function (x) {
          var when = x.date || String(x.createdAt || "").slice(0, 10);
          var tl = x.type === "mileage" ? "🚗 Mileage" : x.type === "professional" ? "📜 PEO/PEng" : "🧾 Expense";
          var label = tl + (x.total ? " · " + x.total : "");
          var dl = x.folderUrl ? '<a class="review-dl" href="' + esc(x.folderUrl) + '" target="_blank" rel="noopener">Open</a>' : '<span class="review-dl" style="color:var(--muted)">Emailed</span>';
          return '<li class="review-row"><span class="review-file"><span class="review-ic">📄</span>' +
            '<span><span class="review-name">' + esc(label) + '</span>' +
            '<span class="review-meta">' + esc(when) + (x.summary ? ' · ' + esc(x.summary) : '') + '</span></span></span>' + dl + '</li>';
        }).join("") + "</ul>";
      }
    }).catch(function () { if (el) el.innerHTML = '<div class="note" style="margin:0">Couldn\'t load your claim history right now.</div>'; });
  }

  /* ---------- Expense / mileage reimbursement (expense.html) ---------- */
  function initReimbursement() {
    var mTab = document.getElementById("rb-tab-mileage");
    var eTab = document.getElementById("rb-tab-expense");
    if (!mTab && !eTab) return;
    var mPanel = document.getElementById("rb-mileage"), ePanel = document.getElementById("rb-expense");
    var mRows = document.getElementById("mileage-rows"), eRows = document.getElementById("expense-rows");

    function showTab(which) {
      if (mPanel) mPanel.hidden = which !== "mileage";
      if (ePanel) ePanel.hidden = which !== "expense";
      if (pPanel) pPanel.hidden = which !== "prof";
      if (mTab) mTab.classList.toggle("active", which === "mileage");
      if (eTab) eTab.classList.toggle("active", which === "expense");
      if (pTab) pTab.classList.toggle("active", which === "prof");
      if (which === "mileage" && mRows && !mRows.children.length) addTrip();
      if (which === "expense" && eRows && !eRows.children.length) addItem();
      if (which === "prof" && pRows && !pRows.children.length) addProf();
    }
    if (mTab) mTab.onclick = function () { showTab("mileage"); };
    if (eTab) eTab.onclick = function () { showTab("expense"); };

    function addTrip() {
      var card = document.createElement("div"); card.className = "rb-trip";
      card.innerHTML =
        '<div class="rb-row mileage-a">' +
          '<input type="date" data-k="date" aria-label="Trip date">' +
          '<input type="text" data-k="from" placeholder="From" autocomplete="off">' +
          '<input type="text" data-k="to" placeholder="To" autocomplete="off">' +
          '<input type="number" data-k="km" placeholder="km" min="0" step="0.1" inputmode="decimal">' +
        '</div>' +
        '<div class="rb-row mileage-b">' +
          '<input type="text" data-k="note" placeholder="Note (optional)" autocomplete="off">' +
          '<button type="button" class="rb-del" title="Remove trip">×</button>' +
        '</div>';
      card.querySelector(".rb-del").onclick = function () { card.remove(); recalcMileage(); };
      card.querySelector('[data-k="km"]').addEventListener("input", recalcMileage);
      mRows.appendChild(card);
    }
    function recalcMileage() {
      var t = 0; mRows.querySelectorAll('[data-k="km"]').forEach(function (i) { var v = parseFloat(i.value); if (!isNaN(v)) t += v; });
      document.getElementById("mileage-total").textContent = (Math.round(t * 10) / 10) + " km";
    }
    var addTripBtn = document.getElementById("add-trip"); if (addTripBtn) addTripBtn.onclick = addTrip;

    function addItem() {
      var card = document.createElement("div"); card.className = "rb-trip";
      card.innerHTML =
        '<div class="rb-row expense-a">' +
          '<input type="text" data-k="project" placeholder="Project / purpose" autocomplete="off">' +
          '<input type="text" data-k="desc" placeholder="Description" autocomplete="off">' +
          '<input type="number" data-k="amt" placeholder="0.00" min="0" step="0.01" inputmode="decimal">' +
        '</div>' +
        '<div class="rb-row expense-b">' +
          '<input type="file" data-k="receipt" accept=".pdf,.jpg,.jpeg,.png,.heic,.webp" aria-label="Receipt for this item">' +
          '<button type="button" class="rb-del" title="Remove item">×</button>' +
        '</div>';
      card.querySelector(".rb-del").onclick = function () { card.remove(); recalcExpense(); };
      card.querySelector('[data-k="amt"]').addEventListener("input", recalcExpense);
      eRows.appendChild(card);
    }
    function fileToB64(file) {
      return new Promise(function (resolve, reject) {
        var fr = new FileReader();
        fr.onload = function () { var s = String(fr.result || ""); var i = s.indexOf(","); resolve(i >= 0 ? s.slice(i + 1) : s); };
        fr.onerror = reject; fr.readAsDataURL(file);
      });
    }
    function recalcExpense() {
      var t = 0; eRows.querySelectorAll('[data-k="amt"]').forEach(function (i) { var v = parseFloat(i.value); if (!isNaN(v)) t += v; });
      document.getElementById("expense-total").textContent = "$" + t.toFixed(2);
    }
    var addItemBtn = document.getElementById("add-item"); if (addItemBtn) addItemBtn.onclick = addItem;

    function rowVal(r, k) { var el = r.querySelector('[data-k="' + k + '"]'); return el ? (el.value || "").trim() : ""; }

    // ----- PEO / PEng professional fees -----
    var pPanel = document.getElementById("rb-prof"), pRows = document.getElementById("prof-rows"), pTab = document.getElementById("rb-tab-prof");
    if (pTab) pTab.onclick = function () { showTab("prof"); };
    function addProf() {
      var card = document.createElement("div"); card.className = "rb-trip";
      card.innerHTML =
        '<div class="rb-row prof-a">' +
          '<input type="text" data-k="desc" placeholder="Description (e.g. P.Eng licence fee)" autocomplete="off">' +
          '<input type="number" data-k="amt" placeholder="0.00" min="0" step="0.01" inputmode="decimal">' +
        '</div>' +
        '<div class="rb-row prof-b">' +
          '<input type="file" data-k="receipt" accept=".pdf,.jpg,.jpeg,.png,.heic,.webp" aria-label="Receipt">' +
          '<button type="button" class="rb-del" title="Remove item">×</button>' +
        '</div>';
      card.querySelector(".rb-del").onclick = function () { card.remove(); recalcProf(); };
      card.querySelector('[data-k="amt"]').addEventListener("input", recalcProf);
      pRows.appendChild(card);
    }
    function recalcProf() {
      var t = 0; pRows.querySelectorAll('[data-k="amt"]').forEach(function (i) { var v = parseFloat(i.value); if (!isNaN(v)) t += v; });
      document.getElementById("prof-total").textContent = "$" + t.toFixed(2);
    }
    var addProfBtn = document.getElementById("add-prof"); if (addProfBtn) addProfBtn.onclick = addProf;

    var sp = document.getElementById("send-prof");
    if (sp) sp.onclick = function () {
      var status = document.getElementById("prof-status");
      var cards = Array.prototype.slice.call(pRows.querySelectorAll(".rb-trip"));
      if (!cards.length) { status.style.color = "#C0392B"; status.textContent = "Add at least one item."; return; }
      var sum = 0;
      for (var i = 0; i < cards.length; i++) {
        var r = cards[i], fileEl = r.querySelector('[data-k="receipt"]'), file = fileEl && fileEl.files && fileEl.files[0];
        if (!rowVal(r, "desc") || !rowVal(r, "amt") || !file) { status.style.color = "#C0392B"; status.textContent = "Item " + (i + 1) + ": add a description, amount, and receipt."; return; }
        if (file.size > 4 * 1024 * 1024) { status.style.color = "#C0392B"; status.textContent = "Item " + (i + 1) + ": receipt is over 4 MB — please reduce its size."; return; }
        sum += parseFloat(rowVal(r, "amt")) || 0;
      }
      if (_profClaimedThisYear + sum > 500) {
        if (!window.confirm("This brings your PEO/PEng claims to $" + (_profClaimedThisYear + sum).toFixed(2) + " this year, over the $500 annual limit. Submit anyway?")) return;
      }
      var notes = (document.getElementById("prof-notes").value || "").trim();
      var total = document.getElementById("prof-total").textContent;
      sp.disabled = true; status.style.color = ""; status.textContent = "Uploading receipts…";
      Promise.all(cards.map(function (r) {
        var file = r.querySelector('[data-k="receipt"]').files[0];
        return fileToB64(file).then(function (b64) {
          return { description: rowVal(r, "desc"), amount: rowVal(r, "amt"), receipt: { name: file.name, contentType: file.type || "application/octet-stream", base64: b64 } };
        });
      })).then(function (items) {
        return fetch("/api/reimbursement-submit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ type: "professional", items: items, total: total, notes: notes }) })
          .then(function (rr) { return rr.json().then(function (j) { return { code: rr.status, ok: rr.ok, j: j }; }); });
      }).then(function (res) {
        sp.disabled = false;
        if (res.ok && res.j && res.j.ok) {
          status.style.color = "var(--gold-dark)"; status.textContent = "Sent to " + (res.j.dept || "HR") + " for review. You'll hear back once it's approved.";
          pRows.innerHTML = ""; addProf(); recalcProf(); loadClaimHistory();
        } else if (res.code === 501) {
          status.style.color = "#C0392B"; status.innerHTML = "Reimbursements aren't fully set up yet — please email <a href=\"mailto:hr@kingepcm.com\">hr@kingepcm.com</a>.";
        } else {
          status.style.color = "#C0392B"; status.textContent = (res.j && res.j.error) || "Couldn't submit. Please try again.";
        }
      }).catch(function () { sp.disabled = false; status.style.color = "#C0392B"; status.textContent = "Couldn't submit. Please try again."; });
    };

    var sm = document.getElementById("send-mileage");
    if (sm) sm.onclick = function () {
      var status = document.getElementById("mileage-status"), trips = [];
      mRows.querySelectorAll(".rb-trip").forEach(function (r) {
        var date = rowVal(r, "date"), from = rowVal(r, "from"), to = rowVal(r, "to"), km = rowVal(r, "km"), note = rowVal(r, "note");
        if (date || from || to || km || note) trips.push({ date: date, from: from, to: to, km: km, note: note });
      });
      if (!trips.length) { status.style.color = "#C0392B"; status.textContent = "Add at least one trip."; return; }
      var notes = (document.getElementById("mileage-notes").value || "").trim();
      var total = document.getElementById("mileage-total").textContent;
      sm.disabled = true; status.style.color = ""; status.textContent = "Sending…";
      fetch("/api/reimbursement-submit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ type: "mileage", trips: trips, total: total, notes: notes }) })
        .then(function (rr) { return rr.json().then(function (j) { return { code: rr.status, ok: rr.ok, j: j }; }); })
        .then(function (res) {
          sm.disabled = false;
          if (res.ok && res.j && res.j.ok) {
            status.style.color = "var(--gold-dark)"; status.textContent = "Sent to " + (res.j.dept || "Accounting") + ". You'll be reimbursed with your next pay run.";
            mRows.innerHTML = ""; addTrip(); recalcMileage(); loadClaimHistory();
          } else if (res.code === 501) {
            status.style.color = "#C0392B"; status.innerHTML = "Reimbursements aren't fully set up yet — please email <a href=\"mailto:accounting@kingepcm.com\">accounting@kingepcm.com</a>.";
          } else { status.style.color = "#C0392B"; status.textContent = (res.j && res.j.error) || "Couldn't send. Please try again."; }
        })
        .catch(function () { sm.disabled = false; status.style.color = "#C0392B"; status.textContent = "Couldn't send. Please try again."; });
    };

    var se = document.getElementById("send-expense");
    if (se) se.onclick = function () {
      var status = document.getElementById("expense-status");
      var cards = Array.prototype.slice.call(eRows.querySelectorAll(".rb-trip"));
      if (!cards.length) { status.style.color = "#C0392B"; status.textContent = "Add at least one item."; return; }
      // Every field — project, description, amount, and a receipt — is required per item.
      for (var i = 0; i < cards.length; i++) {
        var r = cards[i], fileEl = r.querySelector('[data-k="receipt"]'), file = fileEl && fileEl.files && fileEl.files[0];
        if (!rowVal(r, "project") || !rowVal(r, "desc") || !rowVal(r, "amt") || !file) {
          status.style.color = "#C0392B"; status.textContent = "Item " + (i + 1) + ": add the project, description, amount, and a receipt."; return;
        }
        if (file.size > 4 * 1024 * 1024) { status.style.color = "#C0392B"; status.textContent = "Item " + (i + 1) + ": receipt is over 4 MB — please reduce its size."; return; }
      }
      var notes = (document.getElementById("expense-notes").value || "").trim();
      var total = document.getElementById("expense-total").textContent;
      se.disabled = true; status.style.color = ""; status.textContent = "Uploading receipts…";
      Promise.all(cards.map(function (r) {
        var file = r.querySelector('[data-k="receipt"]').files[0];
        return fileToB64(file).then(function (b64) {
          return { project: rowVal(r, "project"), description: rowVal(r, "desc"), amount: rowVal(r, "amt"), receipt: { name: file.name, contentType: file.type || "application/octet-stream", base64: b64 } };
        });
      })).then(function (items) {
        return fetch("/api/reimbursement-submit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ items: items, total: total, notes: notes }) })
          .then(function (rr) { return rr.json().then(function (j) { return { code: rr.status, ok: rr.ok, j: j }; }); });
      }).then(function (res) {
        se.disabled = false;
        if (res.ok && res.j && res.j.ok) {
          status.style.color = "var(--gold-dark)"; status.textContent = "Sent to " + (res.j.dept || "Accounting") + " with your receipts. You'll be reimbursed with your next pay run.";
          eRows.innerHTML = ""; addItem(); recalcExpense(); loadClaimHistory();
        } else if (res.code === 501) {
          status.style.color = "#C0392B"; status.innerHTML = "Reimbursements aren't fully set up yet — please email <a href=\"mailto:accounting@kingepcm.com\">accounting@kingepcm.com</a>.";
        } else {
          status.style.color = "#C0392B"; status.textContent = (res.j && res.j.error) || "Couldn't submit the claim. Please try again.";
        }
      }).catch(function () { se.disabled = false; status.style.color = "#C0392B"; status.textContent = "Couldn't submit the claim. Please try again."; });
    };
  }

  /* ---------- Purchase request history (purchase-request.html) ---------- */
  function loadPurchaseHistory() {
    var el = document.getElementById("purchase-history");
    if (!el) return;
    fetch("/api/my-purchase-requests").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var c = (d && d.requests) || [];
      if (!c.length) { el.innerHTML = '<div class="note" style="margin:0">No purchase requests submitted yet.</div>'; return; }
      el.innerHTML = '<ul class="review-list">' + c.map(function (x) {
        var when = x.date || String(x.createdAt || "").slice(0, 10);
        var head = "🛒 " + esc(x.project || "Purchase") + (x.total ? " · " + esc(x.total) : "");
        var bits = [x.vendor ? "From " + x.vendor : "", x.summary || ""].filter(Boolean).join(" · ");
        var meta = [when, bits].filter(Boolean).join(" · ") + (x.attachmentCount ? " · 📎 " + x.attachmentCount : "");
        return '<li class="review-row"><span class="review-file"><span class="review-ic">📄</span>' +
          '<span><span class="review-name">' + head + '</span>' +
          '<span class="review-meta">' + esc(meta) + '</span></span></span>' +
          '<span class="review-dl" style="color:var(--muted)">Sent</span></li>';
      }).join("") + "</ul>";
    }).catch(function () { el.innerHTML = '<div class="note" style="margin:0">Couldn\'t load your purchase requests right now.</div>'; });
  }

  /* ---------- Purchase request form (purchase-request.html) ---------- */
  function initPurchaseRequest() {
    var sendBtn = document.getElementById("send-purchase");
    if (!sendBtn) return;
    var rows = document.getElementById("purchase-rows");
    var totalEl = document.getElementById("purchase-total");
    var statusEl = document.getElementById("purchase-status");

    function recalc() {
      var grand = 0;
      rows.querySelectorAll(".rb-trip").forEach(function (card) {
        var q = parseFloat((card.querySelector('[data-k="qty"]') || {}).value);
        var u = parseFloat((card.querySelector('[data-k="unit"]') || {}).value);
        var line = (isNaN(q) ? 0 : q) * (isNaN(u) ? 0 : u);
        var lineEl = card.querySelector('[data-k="line"]');
        if (lineEl) lineEl.textContent = "= $" + line.toFixed(2);
        grand += line;
      });
      totalEl.textContent = "$" + grand.toFixed(2);
    }
    function addItem() {
      var card = document.createElement("div"); card.className = "rb-trip";
      card.innerHTML =
        '<div class="rb-row purchase-a">' +
          '<input type="text" data-k="desc" placeholder="Item / description" autocomplete="off">' +
          '<input type="number" data-k="qty" placeholder="Qty" min="0" step="1" inputmode="numeric" value="1">' +
          '<input type="number" data-k="unit" placeholder="Unit price" min="0" step="0.01" inputmode="decimal">' +
          '<span class="pr-line" data-k="line">= $0.00</span>' +
          '<button type="button" class="rb-del" title="Remove item">×</button>' +
        '</div>';
      card.querySelector(".rb-del").onclick = function () { card.remove(); recalc(); };
      card.querySelector('[data-k="qty"]').addEventListener("input", recalc);
      card.querySelector('[data-k="unit"]').addEventListener("input", recalc);
      rows.appendChild(card);
    }
    var addBtn = document.getElementById("add-purchase-item"); if (addBtn) addBtn.onclick = addItem;
    if (!rows.children.length) addItem();

    function rowVal(r, k) { var el = r.querySelector('[data-k="' + k + '"]'); return el ? (el.value || "").trim() : ""; }
    function fileToB64(file) {
      return new Promise(function (resolve, reject) {
        var fr = new FileReader();
        fr.onload = function () { var s = String(fr.result || ""); var i = s.indexOf(","); resolve(i >= 0 ? s.slice(i + 1) : s); };
        fr.onerror = reject; fr.readAsDataURL(file);
      });
    }

    sendBtn.onclick = function () {
      var project = (document.getElementById("pr-project").value || "").trim();
      var vendor = (document.getElementById("pr-vendor").value || "").trim();
      var comments = (document.getElementById("purchase-comments").value || "").trim();
      if (!project) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter the project or purpose."; return; }
      if (!vendor) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter where you're buying from."; return; }
      var cards = Array.prototype.slice.call(rows.querySelectorAll(".rb-trip"));
      var items = [];
      for (var i = 0; i < cards.length; i++) {
        var d = rowVal(cards[i], "desc"), q = rowVal(cards[i], "qty"), u = rowVal(cards[i], "unit");
        if (!d && !q && !u) continue;
        var qn = parseFloat(q), un = parseFloat(u);
        if (!d || !(qn > 0) || !(un > 0)) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Item " + (i + 1) + ": add a description, quantity, and unit price."; return; }
        items.push({ description: d, quantity: q, unitPrice: u, price: (qn * un).toFixed(2) });
      }
      if (!items.length) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Add at least one item."; return; }

      var fileInput = document.getElementById("purchase-files");
      var files = fileInput && fileInput.files ? Array.prototype.slice.call(fileInput.files) : [];
      var totalSize = 0;
      for (var f = 0; f < files.length; f++) {
        if (files[f].size > 4 * 1024 * 1024) { statusEl.style.color = "#C0392B"; statusEl.textContent = '"' + files[f].name + '" is over 4 MB — please reduce its size.'; return; }
        totalSize += files[f].size;
      }
      if (totalSize > 3 * 1024 * 1024) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Attachments total over 3 MB — please remove or shrink some."; return; }

      var total = totalEl.textContent;
      sendBtn.disabled = true; statusEl.style.color = ""; statusEl.textContent = files.length ? "Uploading…" : "Sending…";
      Promise.all(files.map(function (file) {
        return fileToB64(file).then(function (b64) { return { name: file.name, contentType: file.type || "application/octet-stream", base64: b64 }; });
      })).then(function (attachments) {
        return fetch("/api/purchase-request", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ project: project, vendor: vendor, items: items, total: total, comments: comments, attachments: attachments }) })
          .then(function (rr) { return rr.json().then(function (j) { return { code: rr.status, ok: rr.ok, j: j }; }); });
      }).then(function (res) {
        sendBtn.disabled = false;
        if (res.ok && res.j && res.j.ok) {
          statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "Sent to Tony Wang and Accounting for approval. You'll hear back once it's reviewed.";
          rows.innerHTML = ""; addItem(); recalc();
          document.getElementById("pr-project").value = ""; document.getElementById("pr-vendor").value = "";
          document.getElementById("purchase-comments").value = ""; if (fileInput) fileInput.value = "";
          loadPurchaseHistory();
        } else if (res.code === 501) {
          statusEl.style.color = "#C0392B"; statusEl.innerHTML = "Purchase requests aren't fully set up yet — please email <a href=\"mailto:twang@kingepcm.com\">twang@kingepcm.com</a>.";
        } else {
          statusEl.style.color = "#C0392B"; statusEl.textContent = (res.j && res.j.error) || "Couldn't submit the request. Please try again.";
        }
      }).catch(function () { sendBtn.disabled = false; statusEl.style.color = "#C0392B"; statusEl.textContent = "Couldn't submit the request. Please try again."; });
    };
  }

  /* ---------- General HR enquiry (hr.html) ---------- */
  function initHrInquiry() {
    var btn = document.getElementById("send-hr-inquiry");
    if (!btn) return;
    var statusEl = document.getElementById("hr-inquiry-status");
    function fileToB64(file) {
      return new Promise(function (resolve, reject) {
        var fr = new FileReader();
        fr.onload = function () { var s = String(fr.result || ""); var i = s.indexOf(","); resolve(i >= 0 ? s.slice(i + 1) : s); };
        fr.onerror = reject; fr.readAsDataURL(file);
      });
    }
    btn.onclick = function () {
      var topic = (document.getElementById("hr-topic").value || "").trim();
      var message = (document.getElementById("hr-message").value || "").trim();
      if (!message) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Please type your message."; return; }
      var fileInput = document.getElementById("hr-files");
      var files = fileInput && fileInput.files ? Array.prototype.slice.call(fileInput.files) : [];
      var totalSize = 0;
      for (var i = 0; i < files.length; i++) {
        if (files[i].size > 4 * 1024 * 1024) { statusEl.style.color = "#C0392B"; statusEl.textContent = '"' + files[i].name + '" is over 4 MB — please reduce its size.'; return; }
        totalSize += files[i].size;
      }
      if (totalSize > 3 * 1024 * 1024) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Attachments total over 3 MB — please remove or shrink some."; return; }
      btn.disabled = true; statusEl.style.color = ""; statusEl.textContent = files.length ? "Uploading…" : "Sending…";
      Promise.all(files.map(function (file) {
        return fileToB64(file).then(function (b64) { return { name: file.name, contentType: file.type || "application/octet-stream", base64: b64 }; });
      })).then(function (attachments) {
        return fetch("/api/hr-inquiry", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ topic: topic, message: message, attachments: attachments }) })
          .then(function (rr) { return rr.json().then(function (j) { return { code: rr.status, ok: rr.ok, j: j }; }); });
      }).then(function (res) {
        btn.disabled = false;
        if (res.ok && res.j && res.j.ok) {
          statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "Sent to HR. They'll reply by email.";
          document.getElementById("hr-message").value = ""; if (fileInput) fileInput.value = ""; document.getElementById("hr-topic").selectedIndex = 0;
        } else if (res.code === 501) {
          statusEl.style.color = "#C0392B"; statusEl.innerHTML = "HR enquiries aren't fully set up yet — please email <a href=\"mailto:hr@kingepcm.com\">hr@kingepcm.com</a>.";
        } else {
          statusEl.style.color = "#C0392B"; statusEl.textContent = (res.j && res.j.error) || "Couldn't send. Please try again.";
        }
      }).catch(function () { btn.disabled = false; statusEl.style.color = "#C0392B"; statusEl.textContent = "Couldn't send. Please try again."; });
    };
  }

  /* ---------- Highlight the next upcoming statutory holiday (hr.html) ---------- */
  function initHolidays() {
    var ul = document.getElementById("holiday-list");
    if (!ul) return;
    var today = new Date(); today.setHours(0, 0, 0, 0);
    var next = null, nextTime = Infinity;
    ul.querySelectorAll("li[data-date]").forEach(function (li) {
      var d = parseDue(li.getAttribute("data-date"));
      if (d >= today && d.getTime() < nextTime) { nextTime = d.getTime(); next = li; }
    });
    // Collapse the panel by default on mobile/tablet (stays open on desktop).
    var det = document.getElementById("holidays");
    if (det && det.tagName.toLowerCase() === "details" && window.innerWidth <= 1024) det.open = false;
    if (!next) return;
    next.classList.add("next-holiday");
    var nameEl = next.querySelector(".hname");
    if (nameEl && !next.querySelector(".next-badge")) {
      var b = document.createElement("span"); b.className = "next-badge"; b.textContent = "Next";
      nameEl.appendChild(b);
    }
  }

  /* ---------- Reliance letter generator (reliance-letter.html, engineers/admin) ---------- */
  function initRelianceLetter() {
    var form = document.getElementById("rlForm");
    if (!form) return;
    var gate = document.getElementById("rl-gate");
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("engineer") > -1 || roles.indexOf("admin") > -1) form.hidden = false;
      else if (gate) gate.hidden = false;
    }).catch(function () { if (gate) gate.hidden = false; });

    var typeEl = document.getElementById("rl-type");
    function show(sel, on) { document.querySelectorAll(sel).forEach(function (el) { el.hidden = !on; }); }
    function applyType() {
      var t = typeEl.value; // bank | municipal | other
      show(".rl-recipient", t === "bank" || t === "other"); // To: block (lender or other party)
      show(".rl-bank", t === "bank");                        // bank insurance fields
      show(".rl-municipal", t === "municipal");              // authority fields + municipal insurance
      show(".rl-insurance", t !== "other");                  // hide the whole insurance section for "other"
      show(".rl-warranty", t === "bank" || t === "other");   // report basis selector (municipal is always 153/04)
    }
    typeEl.addEventListener("change", applyType);
    applyType();

    // Default the letter date to today.
    var dateEl = document.getElementById("rl-date");
    if (dateEl && !dateEl.value) dateEl.value = new Date().toISOString().slice(0, 10);

    // The signatory + QP default to the signed-in engineer (their M365 profile). Editable.
    fetch("/api/my-profile").then(function (r) { return r.ok ? r.json() : null; }).then(function (me) {
      if (!me) return;
      var name = me.displayName || "";
      var phone = me.mobilePhone || (me.businessPhones && me.businessPhones[0]) || "";
      function setIf(id, v) { var el = document.getElementById(id); if (el && v && !el.value) el.value = v; }
      setIf("rl-qp", name);
      setIf("rl-signer-name", name);
      setIf("rl-signer-title", me.jobTitle || "");
      setIf("rl-phone", phone);
      setIf("rl-email", me.mail || "");
    }).catch(function () {});

    // Repeatable report rows.
    var rows = document.getElementById("rl-reports");
    function addReport() {
      var row = document.createElement("div");
      row.className = "rb-row rl-report-row";
      row.style.gridTemplateColumns = "1.8fr 1fr 28px";
      row.style.alignItems = "center";
      row.innerHTML =
        '<input type="text" data-k="rtype" placeholder="e.g. Phase I Environmental Site Assessment" autocomplete="off">' +
        '<input type="date" data-k="rdate" aria-label="Report date">' +
        '<button type="button" class="rb-del" title="Remove report">×</button>';
      row.querySelector(".rb-del").onclick = function () { row.remove(); };
      rows.appendChild(row);
    }
    var addBtn = document.getElementById("rl-add-report"); if (addBtn) addBtn.onclick = addReport;
    if (!rows.children.length) addReport();

    function val(id) { var el = document.getElementById(id); return el ? (el.value || "").trim() : ""; }
    var statusEl = document.getElementById("rl-status");
    var submitEl = document.getElementById("rl-submit");
    var previewEl = document.getElementById("rl-preview");
    var lastPdf = null, lastUrl = null;

    // Validate the form and build the request payload (returns null + sets the status on failure).
    function collect() {
      var type = typeEl.value;
      var site = val("rl-site");
      if (!site) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter the site / property address."; return null; }
      if (type === "bank" && !val("rl-recipient-org")) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter the lender / bank name."; return null; }
      if (type === "municipal" && !val("rl-authority")) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter the municipality / authority name."; return null; }
      var reports = [];
      rows.querySelectorAll(".rl-report-row").forEach(function (r) {
        var ty = (r.querySelector('[data-k="rtype"]').value || "").trim();
        var dt = (r.querySelector('[data-k="rdate"]').value || "").trim();
        if (ty || dt) reports.push({ type: ty, date: dt });
      });
      if (!reports.length || !reports[0].type) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Add at least one report (with a description)."; return null; }
      return {
        relType: type, warrantyBasis: val("rl-basis"), date: val("rl-date"), siteAddress: site, clientName: val("rl-client"),
        recipientName: val("rl-recipient-name"), recipientOrg: val("rl-recipient-org"), recipientAddress: val("rl-recipient-address"),
        authorityName: val("rl-authority"), authorityShort: val("rl-authority-short"), authorityContact: val("rl-authority-contact"), authorityAddress: val("rl-authority-address"),
        qpName: val("rl-qp"), reports: reports,
        plPerClaim: val("rl-pl-claim"), plAggregate: val("rl-pl-agg"), cglPerClaim: val("rl-cgl-claim"), cglAggregate: val("rl-cgl-agg"), liabilityCap: val("rl-cap"),
        mplPerClaim: val("rl-mpl-claim"), mplAggregate: val("rl-mpl-agg"),
        signerName: val("rl-signer-name"), signerTitle: val("rl-signer-title"), phone: val("rl-phone"), email: val("rl-email")
      };
    }
    // POST the payload and resolve with the result JSON (or null after showing an error).
    function requestPdf(payload, busyText) {
      submitEl.disabled = true; if (previewEl) previewEl.disabled = true;
      statusEl.style.color = ""; statusEl.textContent = busyText;
      return fetch("/api/reliance-letter", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
        .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, code: r.status, j: j }; }); })
        .then(function (res) {
          submitEl.disabled = false; if (previewEl) previewEl.disabled = false;
          if (res.ok && res.j && res.j.ok) return res.j;
          if (res.code === 403) statusEl.textContent = "You need the engineer role to generate reliance letters.";
          else statusEl.textContent = (res.j && res.j.error) || "Couldn't generate the letter. Please try again.";
          statusEl.style.color = "#C0392B"; return null;
        })
        .catch(function () { submitEl.disabled = false; if (previewEl) previewEl.disabled = false; statusEl.style.color = "#C0392B"; statusEl.textContent = "Couldn't generate the letter. Please try again."; return null; });
    }
    function b64ToBlobUrl(b64) {
      var bin = atob(b64), len = bin.length, arr = new Uint8Array(len);
      for (var i = 0; i < len; i++) arr[i] = bin.charCodeAt(i);
      return URL.createObjectURL(new Blob([arr], { type: "application/pdf" }));
    }
    function showPreview(b64) {
      if (lastUrl) { try { URL.revokeObjectURL(lastUrl); } catch (e) {} }
      lastUrl = b64ToBlobUrl(b64);
      var fr = document.getElementById("rl-preview-frame"); if (fr) fr.src = lastUrl;
      var wrap = document.getElementById("rl-preview-wrap"); if (wrap) { wrap.hidden = false; if (wrap.scrollIntoView) wrap.scrollIntoView({ behavior: "smooth", block: "start" }); }
    }

    if (previewEl) previewEl.addEventListener("click", function () {
      var p = collect(); if (!p) return;
      requestPdf(p, "Generating preview…").then(function (j) {
        if (!j) return;
        lastPdf = j; showPreview(j.pdfBase64);
        statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "Preview ready below — review, then download.";
      });
    });
    function saveDocx(j) { if (j && j.docxBase64) downloadBase64File(j.docxBase64, j.docxFilename || "Reliance Letter.docx", DOCX_MIME); }
    var dlBtn = document.getElementById("rl-preview-download");
    if (dlBtn) dlBtn.addEventListener("click", function () { if (lastPdf) downloadBase64Pdf(lastPdf.pdfBase64, lastPdf.filename || "Reliance Letter.pdf"); });
    var dlDocxBtn = document.getElementById("rl-preview-docx");
    if (dlDocxBtn) dlDocxBtn.addEventListener("click", function () { if (lastPdf) saveDocx(lastPdf); });

    // PDF download (form submit)
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      var p = collect(); if (!p) return;
      requestPdf(p, "Generating PDF…").then(function (j) {
        if (!j) return;
        lastPdf = j; downloadBase64Pdf(j.pdfBase64, j.filename || "Reliance Letter.pdf");
        statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "PDF generated — check your downloads.";
      });
    });
    // Word (.docx) download
    var docxBtn = document.getElementById("rl-submit-docx");
    if (docxBtn) docxBtn.addEventListener("click", function () {
      var p = collect(); if (!p) return;
      requestPdf(p, "Generating Word document…").then(function (j) {
        if (!j) return;
        lastPdf = j; saveDocx(j);
        statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "Word document generated — check your downloads.";
      });
    });
  }

  /* ---------- HR: salary increase / promotion letters (hr/admin) ---------- */
  function initSalaryLetter() {
    var form = document.getElementById("salForm");
    if (!form) return;
    var gate = document.getElementById("sal-gate");
    fetch("/.auth/me").then(function (r) { return r.ok ? r.json() : null; }).then(function (d) {
      var roles = (d && d.clientPrincipal && d.clientPrincipal.userRoles) || [];
      if (roles.indexOf("hr") > -1 || roles.indexOf("admin") > -1) form.hidden = false;
      else if (gate) gate.hidden = false;
    }).catch(function () { if (gate) gate.hidden = false; });

    var statusEl = document.getElementById("sal-status");
    var submitEl = document.getElementById("sal-submit");
    var pctEl = document.getElementById("sal-pct");
    function val(id) { var el = document.getElementById(id); return el ? (el.value || "") : ""; }
    function num(id) { var n = Number(val(id).replace(/[^0-9.]/g, "")); return isNaN(n) ? 0 : n; }
    function money(n) { return "$" + Number(n || 0).toLocaleString("en-CA", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

    function applyKind() {
      var promo = form.getAttribute("data-kind") === "promotion";
      form.querySelectorAll(".sal-promo-only").forEach(function (el) { el.hidden = !promo; });
      // salary required for a raise, optional for a promotion
      form.querySelectorAll(".sal-pay-req").forEach(function (el) { el.style.display = promo ? "none" : ""; });
      var lbl = document.getElementById("sal-pay-label");
      if (lbl) lbl.firstChild.nodeValue = promo ? "Salary change (optional) " : "Salary change ";
    }
    form.querySelectorAll(".cal-vbtn").forEach(function (b) {
      b.addEventListener("click", function () {
        form.setAttribute("data-kind", b.getAttribute("data-kind") || "raise");
        form.querySelectorAll(".cal-vbtn").forEach(function (x) { x.classList.toggle("active", x === b); });
        applyKind();
      });
    });
    applyKind();

    function updatePct() {
      var prev = num("sal-prev"), next = num("sal-new");
      var pb = val("sal-prevbasis") === "hour" ? "hour" : "year", nb = val("sal-newbasis") === "hour" ? "hour" : "year";
      if (!prev || !next) { pctEl.style.display = "none"; return; }
      var fromTxt = money(prev) + (pb === "hour" ? "/hr" : "/yr"), toTxt = money(next) + (nb === "hour" ? "/hr" : "/yr");
      pctEl.style.display = "block";
      if (pb === nb) {
        var p = (next - prev) / prev * 100, r = Math.round(p * 10) / 10;
        pctEl.style.color = p >= 0 ? "var(--gold-dark)" : "#C0392B";
        pctEl.textContent = (p >= 0 ? "Increase" : "Change") + ": " + (r % 1 === 0 ? r : r.toFixed(1)) + "%   (" + fromTxt + " → " + toTxt + ")";
      } else {
        pctEl.style.color = "var(--muted)";
        pctEl.textContent = fromTxt + " → " + toTxt + "  (basis change — no percentage)";
      }
    }
    ["sal-prev", "sal-new", "sal-prevbasis", "sal-newbasis"].forEach(function (id) { var el = document.getElementById(id); if (el) el.addEventListener(id.indexOf("basis") > -1 ? "change" : "input", updatePct); });

    var last = null, lastUrl = null;
    function collect() {
      var kind = form.getAttribute("data-kind") === "promotion" ? "promotion" : "raise";
      var name = val("sal-name").trim();
      if (!name) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Employee legal name is required."; return null; }
      if (kind === "raise" && (!num("sal-prev") || !num("sal-new"))) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter the previous and new amounts."; return null; }
      if (kind === "promotion" && !val("sal-newpos").trim()) { statusEl.style.color = "#C0392B"; statusEl.textContent = "Enter the new position."; return null; }
      return {
        kind: kind, name: name, effectiveDate: val("sal-eff"),
        prevPosition: val("sal-prevpos").trim(), newPosition: val("sal-newpos").trim(),
        prevBasis: val("sal-prevbasis"), newBasis: val("sal-newbasis"), prevAmount: val("sal-prev").trim(), newAmount: val("sal-new").trim(),
        reason: val("sal-reason").trim(),
        signedByName: val("sal-signname").trim(), signedByTitle: val("sal-signtitle").trim(),
        signedByPhone: val("sal-signphone").trim(), signedByEmail: val("sal-signemail").trim()
      };
    }
    function request(busy) {
      if (submitEl.disabled) return Promise.resolve(null); // a request is already in flight
      var payload = collect(); if (!payload) return Promise.resolve(null);
      submitEl.disabled = true; statusEl.style.color = ""; statusEl.textContent = busy;
      return fetch("/api/salary-letter", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) })
        .then(function (r) { return r.json().then(function (j) { return { ok: r.ok, j: j }; }); })
        .then(function (res) {
          submitEl.disabled = false;
          if (res.ok && res.j && res.j.ok) { last = res.j; return res.j; }
          statusEl.style.color = "#C0392B"; statusEl.textContent = (res.j && res.j.error === "HR only") ? "You need the HR role to generate letters." : ((res.j && res.j.error) || "Couldn't generate the letter. Please try again."); return null;
        }).catch(function () { submitEl.disabled = false; statusEl.style.color = "#C0392B"; statusEl.textContent = "Couldn't generate the letter. Please try again."; return null; });
    }
    function filedMsg(j) { return "Letter generated." + (j.filed ? " Filed to the staff's HR folder (" + (j.filedTo || "matched folder") + ")." : " (Downloaded only — couldn't auto-match a staff folder to file it.)"); }
    function showPreview(j) {
      if (lastUrl) { try { URL.revokeObjectURL(lastUrl); } catch (e) {} }
      lastUrl = b64ToBlobUrl(j.pdfBase64, "application/pdf");
      var fr = document.getElementById("sal-preview-frame"); if (fr) fr.src = lastUrl;
      var wrap = document.getElementById("sal-preview-wrap"); if (wrap) { wrap.hidden = false; if (wrap.scrollIntoView) wrap.scrollIntoView({ behavior: "smooth", block: "start" }); }
    }
    var prev = document.getElementById("sal-preview");
    if (prev) prev.addEventListener("click", function () { request("Generating preview…").then(function (j) { if (!j) return; showPreview(j); statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = "Preview ready below — review, then download."; }); });
    var docxBtn = document.getElementById("sal-submit-docx");
    if (docxBtn) docxBtn.addEventListener("click", function () { request("Generating Word document…").then(function (j) { if (!j) return; downloadBase64File(j.docxBase64, j.docxFilename || "letter.docx", DOCX_MIME); statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = filedMsg(j); }); });
    var pPdf = document.getElementById("sal-preview-pdf");
    if (pPdf) pPdf.addEventListener("click", function () { if (last) downloadBase64Pdf(last.pdfBase64, last.filename || "letter.pdf"); });
    var pDocx = document.getElementById("sal-preview-docx");
    if (pDocx) pDocx.addEventListener("click", function () { if (last) downloadBase64File(last.docxBase64, last.docxFilename || "letter.docx", DOCX_MIME); });
    form.addEventListener("submit", function (e) {
      e.preventDefault();
      request("Generating…").then(function (j) { if (!j) return; downloadBase64Pdf(j.pdfBase64, j.filename || "letter.pdf"); statusEl.style.color = "var(--gold-dark)"; statusEl.textContent = filedMsg(j); });
    });
  }

  /* ---------- Teamwork project → address type-ahead (any input[data-twproject]) ---------- */
  function initProjectAddressSearch() {
    if (window._twPickerInit) return; window._twPickerInit = true;
    // Lazy-attach on first focus so it also works for React-rendered inputs (bearing capacity).
    document.addEventListener("focusin", function (e) {
      var inp = e.target;
      if (inp && inp.tagName === "INPUT" && inp.getAttribute && inp.getAttribute("data-twproject") != null && !inp._twAttached) {
        inp._twAttached = true;
        try { attachTwProjectSearch(inp); } catch (err) {}
      }
    });
  }
  function attachTwProjectSearch(inp) {
    inp.setAttribute("autocomplete", "off");
    function escHtml(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
    function escAttr(s) { return escHtml(s).replace(/"/g, "&quot;"); }
    // Dropdown lives on <body> (position:fixed) so we never re-parent the input — safe for React.
    var dd = document.createElement("div");
    dd.style.cssText = "position:fixed;z-index:9999;background:#fff;border:1px solid #d9dee6;border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.16);max-height:260px;overflow:auto;display:none;font-family:inherit";
    document.body.appendChild(dd);
    var valueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
    var timer, lastQ = "", items = [], suppress = false;
    function place() { var r = inp.getBoundingClientRect(); dd.style.left = r.left + "px"; dd.style.top = (r.bottom + 3) + "px"; dd.style.width = r.width + "px"; }
    function close() { dd.style.display = "none"; }
    function fill(v) { suppress = true; valueSetter.call(inp, v); inp.dispatchEvent(new Event("input", { bubbles: true })); inp.dispatchEvent(new Event("change", { bubbles: true })); close(); }
    function render() {
      if (!items.length) { close(); return; }
      var head = '<div style="padding:6px 11px;font-size:11px;color:#9aa3b2;border-bottom:1px solid #eef1f5">Teamwork projects — or keep typing any address</div>';
      dd.innerHTML = head + items.map(function (p) {
        return '<div class="tw-opt" data-name="' + escAttr(p.name) + '" style="padding:8px 11px;cursor:pointer;font-size:.9rem;border-bottom:1px solid #f1f4f8">' +
          escHtml(p.name) + (p.company ? ' <span style="color:#8a93a3;font-size:.8rem">· ' + escHtml(p.company) + '</span>' : '') + '</div>';
      }).join("");
      place(); dd.style.display = "block";
    }
    function run(q) {
      if (q === lastQ) { render(); return; } lastQ = q;
      fetch("/api/teamwork-project-search?q=" + encodeURIComponent(q))
        .then(function (r) { return r.ok ? r.json() : null; })
        .then(function (j) { items = (j && j.projects) || []; render(); })
        .catch(close);
    }
    inp.addEventListener("input", function () { if (suppress) { suppress = false; return; } var q = inp.value.trim(); clearTimeout(timer); if (q.length < 2) { close(); return; } timer = setTimeout(function () { run(q); }, 250); });
    inp.addEventListener("blur", function () { setTimeout(close, 160); });
    inp.addEventListener("focus", function () { if (items.length && inp.value.trim().length >= 2) { place(); dd.style.display = "block"; } });
    window.addEventListener("scroll", function () { if (dd.style.display !== "none") place(); }, true);
    window.addEventListener("resize", function () { if (dd.style.display !== "none") place(); });
    dd.addEventListener("mouseover", function (e) { var o = e.target.closest && e.target.closest(".tw-opt"); if (o) o.style.background = "#f3f6fb"; });
    dd.addEventListener("mouseout", function (e) { var o = e.target.closest && e.target.closest(".tw-opt"); if (o) o.style.background = ""; });
    dd.addEventListener("mousedown", function (e) { var o = e.target.closest && e.target.closest(".tw-opt"); if (!o) return; e.preventDefault(); fill(o.getAttribute("data-name")); });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initUser(); initWelcome(); initNav(); initTeamwork(); initMyWork(); initProjectTools(); initProjectToolPage(); initKickoff(); initNews(); initProcedures(); initResources(); initResTabs(); initMail(); initCalendar(); initProfile(); initReviews(); initManagerLinks(); initHrPortalNav(); initHrPortal(); initStaffPickers(); initReviewForm(); initEmploymentLetter(); initSalaryLetter(); initReimbursement(); loadClaimHistory(); initPurchaseRequest(); loadPurchaseHistory(); initHrInquiry(); initHolidays(); initRelianceLetter(); initProjectAddressSearch();
  });
})();
