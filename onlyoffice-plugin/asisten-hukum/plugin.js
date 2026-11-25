// Plugin ONLYOFFICE "Asisten Hukum": asisten AI di dalam editor Word. Pilih teks -> klik kanan -> "Asisten AI"
// (jelaskan / tulis ulang / sederhanakan / analisis risiko / dasar hukum / terjemahkan), atau pakai panel ini.
// Hasil "tulis ulang" & "terjemahkan" bisa menggantikan teks terpilih langsung di dokumen. Semua permintaan
// lewat tautan bertoken ke server (sama seperti plugin Panel CLM); kunci AI tidak pernah ada di browser.
(function (window) {
  var opts = {}, history = [], busy = false, lastSel = "", engines = null;
  var ACTIONS = [
    { id: "sel_ask", text: "Jelaskan teks ini" }, { id: "sel_rewrite", text: "Tulis ulang lebih formal" },
    { id: "sel_simplify", text: "Sederhanakan" }, { id: "sel_risk", text: "Analisis risiko" },
    { id: "sel_law", text: "Cek dasar hukum" }, { id: "sel_en", text: "Terjemahkan ke Inggris" }
  ];
  function $(id) { return document.getElementById(id); }
  function err(t) { var m = $("msg"); m.className = t ? "err" : ""; m.textContent = t || ""; }

  function api(path, body) {
    return fetch(opts.panelUrl + path + "?t=" + encodeURIComponent(opts.token), body === undefined ? {} : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
      .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error(j.error || ("HTTP " + r.status)); return j; }); });
  }

  function getSelection(cb) {
    try {
      window.Asc.plugin.executeMethod("GetSelectedText", [{ Numbering: false, Math: false, TableCellSeparator: "\n", ParaSeparator: "\n", TabSymbol: "\t" }], function (t) { cb(String(t || "").trim()); });
    } catch (e) { cb(""); }
  }
  function getDocText(cb) {
    window.Asc.plugin.callCommand(function () {
      var doc = Api.GetDocument(), n = 0, out = [];
      try { n = doc.GetElementsCount(); } catch (e) { n = 0; }
      for (var i = 0; i < n && out.join("\n").length < 24000; i++) {
        try { var el = doc.GetElement(i); var t = el.GetText ? el.GetText() : ""; if (t) out.push(t); } catch (e2) {}
      }
      return out.join("\n");
    }, false, false, function (t) { cb(String(t || "")); });
  }

  function addTurn(q) {
    var log = $("log"), dq = document.createElement("div"); dq.className = "q"; dq.textContent = q; log.appendChild(dq);
    var da = document.createElement("div"); da.className = "a"; da.textContent = "Menganalisis…"; log.appendChild(da);
    log.scrollTop = log.scrollHeight; return da;
  }
  function replaceSelection(text) {
    try { window.Asc.plugin.executeMethod("ReplaceTextSmart", [[text]]); return; } catch (e) {}
    try { window.Asc.plugin.executeMethod("PasteText", [text]); } catch (e2) { err("Gagal mengganti teks. Salin manual dari panel."); }
  }

  function run(action, label, question, selection) {
    if (busy) return; busy = true; err("");
    var da = addTurn(label || question);
    // Ada teks terpilih -> fokus ke teks itu saja (jangan kirim seluruh dokumen). Tanpa pilihan -> pakai isi dokumen sebagai konteks.
    (selection ? function (cb) { cb(""); } : getDocText)(function (ctx) {
      api("/ask", { action: action || undefined, question: question || "", selection: selection || "", context: ctx, history: history.slice(-8), engine: $("engine").value })
        .then(function (d) {
          da.textContent = d.answer || "Tidak ada jawaban.";
          history.push({ q: label || question, a: d.answer || "" });
          if (d.answer && selection && (action === "sel_rewrite" || action === "sel_en")) {
            var b = document.createElement("button"); b.textContent = "Ganti teks terpilih di dokumen";
            b.onclick = function () { replaceSelection(d.answer.trim()); b.disabled = true; b.textContent = "Teks diganti"; };
            da.appendChild(document.createElement("br")); da.appendChild(b);
          }
          if (d.sources && d.sources.length) { var s = document.createElement("div"); s.className = "hint"; s.textContent = "Rujukan: " + d.sources.join(", "); da.appendChild(s); }
        })
        .catch(function (e) { da.textContent = "Gagal: " + e.message; })
        .then(function () { busy = false; $("log").scrollTop = $("log").scrollHeight; });
    });
  }
  function runOnSelection(action, label) {
    getSelection(function (t) {
      if (!t) { err("Pilih teks di dokumen dulu."); return; }
      lastSel = t; var box = $("sel"); box.style.display = "block"; box.textContent = "Teks terpilih: " + t;
      run(action, label, "", t);
    });
  }

  function initUi() {
    var d = $("doc").querySelectorAll("button");
    for (var i = 0; i < d.length; i++) d[i].onclick = (function (b) { return function () { run(b.getAttribute("data-a"), b.textContent, "", ""); }; })(d[i]);
    var s = $("selbtns").querySelectorAll("button");
    for (var j = 0; j < s.length; j++) s[j].onclick = (function (b) { return function () { runOnSelection(b.getAttribute("data-a"), b.textContent); }; })(s[j]);
    $("send").onclick = function () {
      var q = $("q").value.trim(); if (!q) return; $("q").value = "";
      getSelection(function (t) { run(undefined, q, q, t); });
    };
    $("q").onkeydown = function (e) { if (e.key === "Enter") $("send").onclick(); };
    api("/status").then(function (st) {
      engines = st; var sel = $("engine"); sel.innerHTML = "";
      [["gemini", "Gemini (cloud)", st.gemini], ["local", "Model hukum lokal", st.local]].forEach(function (e) {
        var o = document.createElement("option"); o.value = e[0]; o.textContent = "Mesin: " + e[1] + (e[2] ? "" : " — belum diatur"); o.disabled = !e[2]; sel.appendChild(o);
      });
      sel.value = st.gemini ? "gemini" : (st.local ? "local" : "gemini");
    }).catch(function () {});
  }

  // Menu klik kanan: muncul hanya bila ada teks terpilih.
  function attachContextMenu() {
    window.Asc.plugin.attachEvent("onContextMenuShow", function (options) {
      if (!options || options.type !== "Selection") return;
      try {
        window.Asc.plugin.executeMethod("AddContextMenuItem", [{ guid: window.Asc.plugin.guid, items: [{ id: "ah_menu", text: "Asisten AI", items: ACTIONS.map(function (a) { return { id: "ah_" + a.id, text: a.text }; }) }] }]);
      } catch (e) {}
    });
    ACTIONS.forEach(function (a) {
      try { window.Asc.plugin.attachContextMenuClickEvent("ah_" + a.id, function () { runOnSelection(a.id, a.text); }); } catch (e) {}
    });
  }

  window.Asc.plugin.init = function () {
    opts = (window.Asc.plugin.info && window.Asc.plugin.info.options) || {};
    if (!opts.panelUrl || !opts.token) { err("Plugin belum dikonfigurasi."); return; }
    initUi(); attachContextMenu();
  };
  window.Asc.plugin.button = function () { };
})(window);
