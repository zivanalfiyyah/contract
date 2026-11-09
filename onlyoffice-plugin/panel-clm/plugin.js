// Plugin ONLYOFFICE "Panel CLM": data pihak & kontrak (diedit langsung di dokumen), bahasa, kop, logo, TTD,
// TTD digital, meterai, dan rangkap. Aksi tata letak/bahasa dikirim ke server (tautan bertoken), lalu halaman
// utama diminta memuat ulang editor lewat postMessage.
(function (window) {
  var opts = {}, state = null, vars = {}, busy = false;
  function $(id) { return document.getElementById(id); }
  function say(t, ok) { var m = $("msg"); m.className = ok ? "ok" : "err"; m.textContent = t; if (ok) setTimeout(function () { m.className = ""; m.textContent = ""; }, 5000); }

  function api(path, body) {
    return fetch(opts.panelUrl + path + (path.indexOf("?") >= 0 ? "&" : "?") + "t=" + encodeURIComponent(opts.token), body === undefined ? {} : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
      .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw new Error(j.error || ("HTTP " + r.status)); return j; }); });
  }
  function reloadHost() { try { window.top.postMessage({ clm: "reload" }, "*"); } catch (e) { try { window.parent.postMessage({ clm: "reload" }, "*"); } catch (e2) {} } }
  function act(action, body, doing) {
    if (busy) return; busy = true; say(doing + "…", true);
    api("/" + action, body).then(function () { say("Selesai. Memuat ulang dokumen…", true); reloadHost(); })
      .catch(function (e) { say("Gagal: " + e.message, false); }).then(function () { busy = false; });
  }

  // ---- Data (variabel) — dibaca & ditulis langsung di dokumen yang sedang terbuka ----
  function readVars() {
    window.Asc.plugin.callCommand(function () {
      var out = {}, doc = Api.GetDocument(), all = [];
      try { all = doc.GetAllContentControls() || []; } catch (e) { all = []; }
      for (var i = 0; i < all.length; i++) {
        var tag = ""; try { tag = all[i].GetTag() || ""; } catch (e1) {}
        if (tag.indexOf("var:") !== 0) continue;
        var key = tag.slice(4); if (key.indexOf("@en") >= 0) continue;
        if (out[key] === undefined) { var t = ""; try { t = all[i].GetRange().GetText() || ""; } catch (e2) {} out[key] = t; }
      }
      return out;
    }, false, true, function (res) { vars = res || {}; renderVars(); });
  }
  function renderVars() {
    var box = $("vars"); box.innerHTML = ""; var keys = Object.keys(vars).sort();
    if (!keys.length) { box.innerHTML = '<div class="hint">Tidak ada data yang bisa diubah di dokumen ini.</div>'; $("saveVars").style.display = "none"; return; }
    keys.forEach(function (k) {
      var l = document.createElement("label"); l.textContent = k; box.appendChild(l);
      var i = document.createElement("input"); i.type = "text"; i.value = vars[k]; i.setAttribute("data-k", k); box.appendChild(i);
    });
    $("saveVars").style.display = "";
  }
  function saveVars() {
    var changed = {}; var ins = document.querySelectorAll("#vars input");
    for (var i = 0; i < ins.length; i++) { var k = ins[i].getAttribute("data-k"); if (ins[i].value !== vars[k]) changed[k] = ins[i].value; }
    if (!Object.keys(changed).length) { say("Tidak ada yang berubah.", true); return; }
    window.Asc.scope.changed = changed;
    window.Asc.plugin.callCommand(function () {
      var ch = Asc.scope.changed, doc = Api.GetDocument(), all = [], n = 0;
      try { all = doc.GetAllContentControls() || []; } catch (e) { all = []; }
      for (var i = 0; i < all.length; i++) {
        var tag = ""; try { tag = all[i].GetTag() || ""; } catch (e1) {}
        if (tag.indexOf("var:") !== 0) continue;
        var key = tag.slice(4).replace("@en", "");
        if (ch[key] === undefined) continue;
        try { all[i].RemoveAllElements(); all[i].AddText(String(ch[key])); n++; } catch (e2) {}
      }
      return n;
    }, false, true, function (n) { say(n ? "Data diterapkan (" + n + " tempat). Tersimpan otomatis & disinkronkan ke data kontrak." : "Tidak ada data yang cocok di dokumen.", !!n); readVars(); });
  }

  // ---- Bahasa / tata letak ----
  function renderLang() {
    var cur = state.lang === "bilingual" ? "bilingual-" + (state.layout === "stacked" ? "stacked" : "side") : state.lang;
    var bs = $("lang").querySelectorAll("button");
    for (var i = 0; i < bs.length; i++) {
      bs[i].className = bs[i].getAttribute("data-m") === cur ? "on" : "";
      bs[i].onclick = (function (b) { return function () {
        var m = b.getAttribute("data-m");
        if (m === cur) return;
        if (m.indexOf("bilingual") === 0) act("language", { mode: "bilingual", layout: m.split("-")[1] }, "Mengubah bahasa");
        else act("language", { mode: m }, "Mengubah bahasa");
      }; })(bs[i]);
    }
  }
  function chk(label, key, val) {
    var d = document.createElement("div"); d.className = "chk";
    var c = document.createElement("input"); c.type = "checkbox"; c.checked = !!val;
    c.onchange = function () { var b = {}; b[key] = c.checked; act("layout", b, "Mengubah tata letak"); };
    var s = document.createElement("span"); s.textContent = label; d.appendChild(c); d.appendChild(s); return d;
  }
  function renderLayout() {
    var L = { kop: state.kop, logo: state.logo, ttd: state.ttd, digital: state.digital, meterai: state.meterai, copies: state.copies }, box = $("layout"); box.innerHTML = "";
    box.appendChild(chk("Kop surat", "kop", L.kop));
    box.appendChild(chk("Logo pada kop", "logo", L.logo));
    box.appendChild(chk("Blok tanda tangan", "ttd", L.ttd));
    box.appendChild(chk("TTD digital", "digital", L.digital));
    box.appendChild(chk("Meterai", "meterai", L.meterai));
    var r = document.createElement("div"); r.className = "row"; r.style.marginTop = "6px";
    [1, 2].forEach(function (n) {
      var b = document.createElement("button"); b.textContent = n + " rangkap"; b.className = L.copies === n ? "on" : "";
      b.onclick = function () { if (L.copies !== n) act("layout", { copies: n }, "Mengubah rangkap"); }; r.appendChild(b);
    });
    box.appendChild(r);
  }

  function loadState() {
    api("").then(function (s) { if (!s.available) { say("Panel hanya untuk dokumen Word.", false); return; } state = s; renderLang(); renderLayout(); })
      .catch(function (e) { say("Gagal memuat: " + e.message, false); });
  }

  window.Asc.plugin.init = function () {
    opts = (window.Asc.plugin.info && window.Asc.plugin.info.options) || {};
    if (!opts.panelUrl) { say("Plugin belum dikonfigurasi.", false); return; }
    $("saveVars").onclick = saveVars;
    loadState(); readVars();
  };
  window.Asc.plugin.button = function () { };
})(window);
