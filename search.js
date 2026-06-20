/* King EPCM — lightweight client-side site search.
   No backend needed: a curated index routes common terms to the right page. */
(function () {
  var INDEX = [
    { u: "index.html", t: "Home", d: "Civil, environmental, geotechnical & permitting engineering across Ontario.",
      k: "home king epcm overview engineering land development ontario markham gta golden horseshoe" },
    { u: "about.html", t: "About King EPCM", d: "Established 2019 — flexible, dependable, on-site engineering.",
      k: "about company history story 2019 flexible dependable on-site mission who we are markham stouffville" },
    { u: "services.html", t: "Services", d: "All seven engineering disciplines.",
      k: "services disciplines all what we do offerings list" },
    { u: "chemical-environmental.html", t: "Chemical Environmental", d: "Phase I & II ESA, contamination, remediation, excess soil.",
      k: "chemical environmental esa phase i phase ii environmental site assessment contamination remediation soil groundwater testing well contractor o.reg 153 406 excess soil management record of site condition rsc designated substance survey dss odour vapour intrusion qualified person brat mecp" },
    { u: "natural-heritage-environmental.html", t: "Natural Heritage Environmental", d: "Tree inventory, natural heritage evaluation, ecological surveys.",
      k: "natural heritage environmental tree inventory arborist report planting plan protection eis environmental impact study ecological biodiversity nhe oak ridges moraine mds omafra nutrient management boating capacity species habitat" },
    { u: "civil-municipal-design.html", t: "Civil & Municipal Design", d: "Grading, drainage, servicing, stormwater, septic, floodplain.",
      k: "civil municipal design grading drainage servicing stormwater swm septic system engineered floodplain modelling fluvial geomorphology erosion sediment control esc functional servicing report fsr lot grading subdivision site plan culvert sanitary sewer water phosphorus lid salt management ifc" },
    { u: "survey-inspections.html", t: "Survey & Inspections", d: "Topographic surveys, asphalt & concrete QA/QC, inspections.",
      k: "survey inspections topographic surveying asphalt concrete qa qc quality control marshall test density compaction proctor rebar formwork slump cctv borehole drilling monitoring well property condition assessment pca structural pull-out cadastral boundary drone bathymetry stockpile" },
    { u: "geotechnical-hydrogeology.html", t: "Geotechnical & Hydrogeology", d: "Soil investigation, bearing capacity, foundations, groundwater.",
      k: "geotechnical hydrogeology soil investigation bearing capacity spt astm excavation testing groundwater foundation review slope stability caisson pile retaining wall dewatering permeameter settlement vibration monitoring water balance well supply pttw esar mecp d-5-4 d-5-5 section 59 borehole testpit" },
    { u: "permit-project-management.html", t: "Permit & Project Management", d: "Approvals, agency coordination, project delivery.",
      k: "permit project management approvals agency coordination septic permit site alteration eca mecp pttw zoning bylaw amendment noise study traffic impact planning justification reliance letter archaeological assessment severance consent conservation authority pre-application consultation private locates" },
    { u: "mining-aggregate.html", t: "Mining & Aggregate", d: "Open-pit & underground engineering across Canada.",
      k: "mining aggregate open pit underground quarry ni 43-101 resource reserve feasibility pea prefeasibility bankable orebody core logging geology cashflow wacc irr paste fill ventsim drill blast cable bolt aggregate resources act ara limestone sand gravel canada miners p.geo" },
    { u: "team.html", t: "Our Team", d: "Engineers, designers, and field experts.",
      k: "team people staff leadership engineers designers field experts tony wang angela shi amir samadi ebrahim partners p.eng professionals" },
    { u: "careers.html", t: "Careers", d: "Join our engineering team.",
      k: "careers jobs hiring employment join work apply resume open positions opportunities engineer" },
    { u: "contact.html", t: "Contact & Request a Quote", d: "Reach us or request a project quote.",
      k: "contact request quote rfq phone email address location map directions get in touch sales info reach us google review 3780 14th avenue markham" },
    { u: "pay.html", t: "Pay Invoice / Retainer", d: "Pay an invoice or retainer (e-transfer, cheque, card).",
      k: "pay payment invoice retainer bill billing e-transfer cheque credit card visa mastercard clover accounting secure" }
  ];

  function norm(s) { return (s || "").toLowerCase().replace(/[^a-z0-9& ]+/g, " "); }

  function search(q) {
    var terms = norm(q).split(/\s+/).filter(Boolean);
    if (!terms.length) return [];
    var scored = INDEX.map(function (e) {
      var title = norm(e.t), keys = norm(e.k), desc = norm(e.d);
      var score = 0;
      terms.forEach(function (term) {
        if (title.indexOf(term) === 0) score += 12;        // title prefix
        else if (title.indexOf(term) > -1) score += 8;     // title contains
        if ((" " + keys + " ").indexOf(" " + term) > -1) score += 5; // keyword word-start
        else if (keys.indexOf(term) > -1) score += 3;      // keyword substring
        if (desc.indexOf(term) > -1) score += 1;
      });
      return { e: e, score: score };
    }).filter(function (r) { return r.score > 0; });
    scored.sort(function (a, b) { return b.score - a.score; });
    return scored.slice(0, 8).map(function (r) { return r.e; });
  }

  document.addEventListener("DOMContentLoaded", function () {
    var toggle = document.getElementById("searchToggle");
    var modal = document.getElementById("searchModal");
    var input = document.getElementById("searchInput");
    var results = document.getElementById("searchResults");
    var closeBtn = document.getElementById("searchClose");
    var navLinks = document.getElementById("navLinks");
    if (!toggle || !modal || !input || !results) return;

    var sel = -1, current = [];

    function open() {
      modal.classList.add("open");
      if (navLinks) navLinks.classList.remove("open"); // close mobile menu if open
      document.body.style.overflow = "hidden";
      setTimeout(function () { input.focus(); }, 30);
    }
    function close() {
      modal.classList.remove("open");
      document.body.style.overflow = "";
      input.value = ""; results.innerHTML = ""; sel = -1; current = [];
    }
    function render(items, q) {
      current = items; sel = -1;
      if (!q.trim()) { results.innerHTML = ""; return; }
      if (!items.length) {
        results.innerHTML = '<li class="search-empty">No matches. Try a service name, e.g. “septic”, “ESA”, or “survey”.</li>';
        return;
      }
      results.innerHTML = items.map(function (e, i) {
        return '<li><a href="' + e.u + '" data-i="' + i + '">' +
          '<span class="r-title">' + e.t + '</span>' +
          '<span class="r-desc">' + e.d + '</span></a></li>';
      }).join("");
    }
    function highlight() {
      var as = results.querySelectorAll("a");
      as.forEach(function (a, i) { a.classList.toggle("sel", i === sel); });
      if (sel > -1 && as[sel]) as[sel].scrollIntoView({ block: "nearest" });
    }

    toggle.addEventListener("click", function (e) { e.preventDefault(); open(); });
    if (closeBtn) closeBtn.addEventListener("click", close);
    modal.addEventListener("click", function (e) { if (e.target === modal) close(); });

    input.addEventListener("input", function () { render(search(input.value), input.value); });

    input.addEventListener("keydown", function (e) {
      var as = results.querySelectorAll("a");
      if (e.key === "ArrowDown") { e.preventDefault(); sel = Math.min(sel + 1, as.length - 1); highlight(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); sel = Math.max(sel - 1, 0); highlight(); }
      else if (e.key === "Enter") {
        e.preventDefault();
        var target = (sel > -1 && current[sel]) ? current[sel] : current[0];
        if (target) window.location.href = target.u;
      } else if (e.key === "Escape") { close(); }
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && modal.classList.contains("open")) close();
    });
  });

  if (typeof module !== "undefined" && module.exports) module.exports = { search: search, INDEX: INDEX };
})();
