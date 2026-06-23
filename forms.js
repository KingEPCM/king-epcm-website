/* King EPCM public site — Quote request & Pay-by-card forms.
   Submits to the site's Azure Functions (no third-party), shows inline status,
   attaches uploaded files, and (for pay) redirects to Clover on success. */
(function () {
  "use strict";

  function val(id) { var el = document.getElementById(id); return el ? (el.value || "").trim() : ""; }
  function setStatus(el, kind, msg) {
    if (!el) return;
    el.textContent = msg;
    el.style.color = kind === "err" ? "#C0392B" : (kind === "ok" ? "#1d7a44" : "");
  }
  function fileToB64(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { var s = String(fr.result || ""); var i = s.indexOf(","); resolve(i >= 0 ? s.slice(i + 1) : s); };
      fr.onerror = reject; fr.readAsDataURL(file);
    });
  }

  /* ---------- Request a Quote (contact.html) ---------- */
  var qf = document.getElementById("quoteForm");
  if (qf) {
    var qStatus = document.getElementById("quoteStatus");
    var qBtn = qf.querySelector("button[type=submit]");
    qf.addEventListener("submit", function (e) {
      e.preventDefault();
      var name = val("name"), email = val("email"), message = val("message");
      if (!name || !email || !message) { setStatus(qStatus, "err", "Please fill in your name, email, and project details."); return; }

      var fileInput = document.getElementById("upload");
      var files = fileInput && fileInput.files ? Array.prototype.slice.call(fileInput.files) : [];
      var total = 0;
      for (var i = 0; i < files.length; i++) {
        if (files[i].size > 4 * 1024 * 1024) { setStatus(qStatus, "err", '"' + files[i].name + '" is over 4 MB — please email large files to sales@KingEPCM.com.'); return; }
        total += files[i].size;
      }
      if (total > 3 * 1024 * 1024) { setStatus(qStatus, "err", "Attachments total over 3 MB — please email them to sales@KingEPCM.com."); return; }

      if (qBtn) qBtn.disabled = true;
      setStatus(qStatus, "", files.length ? "Uploading…" : "Sending…");
      Promise.all(files.map(function (f) {
        return fileToB64(f).then(function (b64) { return { name: f.name, contentType: f.type || "application/octet-stream", base64: b64 }; });
      })).then(function (attachments) {
        return fetch("/api/quote-request", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            website: val("company_url"), name: name, company: val("company"), email: email, phone: val("phone"),
            service: val("service"), location: val("location"), message: message, attachments: attachments
          })
        }).then(function (r) { return r.json().then(function (j) { return { ok: r.ok, code: r.status, j: j }; }); });
      }).then(function (res) {
        if (qBtn) qBtn.disabled = false;
        if (res.ok && res.j && res.j.ok) {
          qf.reset();
          setStatus(qStatus, "ok", "Thank you — your request has been sent. We'll reply within 3 business days.");
        } else if (res.code === 501) {
          setStatus(qStatus, "err", "Our form isn't fully set up yet — please email sales@KingEPCM.com.");
        } else {
          setStatus(qStatus, "err", (res.j && res.j.error) || "Couldn't send your request. Please try again, or email sales@KingEPCM.com.");
        }
      }).catch(function () {
        if (qBtn) qBtn.disabled = false;
        setStatus(qStatus, "err", "Couldn't send your request. Please try again, or email sales@KingEPCM.com.");
      });
    });
  }

  /* ---------- Pay by Credit Card (pay.html) ---------- */
  var pf = document.getElementById("payForm");
  if (pf) {
    var pStatus = document.getElementById("payStatus");
    var pBtn = pf.querySelector("button[type=submit]");
    pf.addEventListener("submit", function (e) {
      e.preventDefault();
      var first = val("firstname"), last = val("lastname"), phone = val("phone"), email = val("email"), invoice = val("invoice");
      if (!first || !last || !phone || !email || !invoice) { setStatus(pStatus, "err", "Please complete all required fields."); return; }
      var clover = pf.getAttribute("data-clover") || "";
      if (pBtn) pBtn.disabled = true;
      setStatus(pStatus, "", "Saving your details…");
      fetch("/api/pay-details", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          website: val("company_url"), first_name: first, last_name: last,
          company: val("company"), phone: phone, email: email, invoice: invoice
        })
      }).then(function (r) { return r.json().then(function (j) { return { ok: r.ok, code: r.status, j: j }; }); })
        .then(function (res) {
          if (res.ok && res.j && res.j.ok) {
            setStatus(pStatus, "ok", "Redirecting you to our secure payment page…");
            if (clover) { window.location.href = clover; }
          } else {
            if (pBtn) pBtn.disabled = false;
            if (res.code === 501) setStatus(pStatus, "err", "Our payment form isn't fully set up yet — please email accounting@KingEPCM.com.");
            else setStatus(pStatus, "err", (res.j && res.j.error) || "Couldn't submit your details. Please try again, or email accounting@KingEPCM.com.");
          }
        }).catch(function () {
          if (pBtn) pBtn.disabled = false;
          setStatus(pStatus, "err", "Couldn't submit your details. Please try again, or email accounting@KingEPCM.com.");
        });
    });
  }
})();
