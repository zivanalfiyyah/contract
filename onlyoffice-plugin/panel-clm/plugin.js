// Plugin ONLYOFFICE "Panel CLM": data pihak & kontrak (padanan modal "Edit Data Pihak & Kontrak" di aplikasi utama),
// bahasa, kop, logo, TTD, TTD digital, meterai, dan rangkap. Aksi tata letak/bahasa dikirim ke server (tautan bertoken),
// lalu halaman utama diminta memuat ulang editor lewat postMessage.
//
// Data pihak & kontrak:
//  - Bentuk form (kelompok, label, tipe input, dropdown, placeholder, field terkunci) DIKIRIM SERVER (party-fields.ts)
//    sehingga selalu sama dengan modal aplikasi utama — plugin ini tidak menyimpan daftar field sendiri.
//  - Simpan  : panel -> database (server memvalidasi ulang) -> placeholder Word yang bersangkutan diperbarui di editor
//              -> aplikasi utama diminta memuat ulang datanya.
//  - Sinkron : tiap ±30 dtk (dan saat panel dibuka): perubahan di aplikasi utama didorong ke placeholder Word, dan isian
//              yang diketik langsung di placeholder ditarik ke database (hanya setelah isinya berhenti berubah).
(function (window) {
  var POLL_MS = 30000;
  var opts = {}, state = null, vars = {}, busy = false;
  var party = null;          // { view, editable, reason }
  var loaded = {};           // nilai terakhir dari server: key -> string
  var dirty = {};            // key -> true bila diubah pengguna di form
  var logoNow = { url: "", key: "", mimeType: "" };
  var prevDoc = null;        // isi variabel dokumen pada sinkron sebelumnya (untuk mendeteksi "sudah berhenti diketik")
  var syncing = false, uploading = false, timer = null;
  var LOGO_KEYS = ["party2LogoUrl", "party2LogoKey", "party2LogoMimeType"];

  function $(id) { return document.getElementById(id); }
  function el(tag, cls, text) { var e = document.createElement(tag); if (cls) e.className = cls; if (text !== undefined) e.textContent = text; return e; }
  function say(t, ok) { var m = $("msg"); m.className = ok ? "ok" : "err"; m.textContent = t; if (ok) setTimeout(function () { m.className = ""; m.textContent = ""; }, 5000); }
  // Pesan di bilah simpan (selalu terlihat walau formulir sedang di-scroll).
  function sayParty(t, ok) {
    var m = $("partyMsg"); if (!m) return;
    m.className = "pmsg " + (ok ? "ok" : "err"); m.textContent = t;
    if (ok) setTimeout(function () { if (m.textContent === t) { m.className = "pmsg"; m.textContent = ""; } }, 6000);
  }

  function api(path, body) {
    return fetch(opts.panelUrl + path + (path.indexOf("?") >= 0 ? "&" : "?") + "t=" + encodeURIComponent(opts.token), body === undefined ? {} : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })
      .then(function (r) {
        return r.json().then(function (j) {
          if (!r.ok) { var e = new Error(j.error || ("HTTP " + r.status)); e.data = j; e.status = r.status; throw e; }
          return j;
        });
      });
  }
  function postHost(msg) { try { window.top.postMessage(msg, "*"); } catch (e) { try { window.parent.postMessage(msg, "*"); } catch (e2) {} } }
  function reloadHost() { postHost({ clm: "reload" }); }
  // Data pihak berubah di database: aplikasi utama diminta mengambil ulang kontraknya (tanpa memuat ulang editor).
  function notifyHost() { postHost({ clm: "party-updated" }); }
  function act(action, body, doing) {
    if (busy) return; busy = true; say(doing + "…", true);
    api("/" + action, body).then(function () { say("Selesai. Memuat ulang dokumen…", true); reloadHost(); })
      .catch(function (e) { say("Gagal: " + e.message, false); }).then(function () { busy = false; });
  }

  // ---- Dokumen: baca / tulis variabel (content control bertag "var:<Nama>") ----
  function readDocVars(cb) {
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
    }, false, true, function (res) { cb(res || {}); });
  }
  // push = { Kunci: { id: "nilai Indonesia", en: "nilai Inggris" } } -> tulis ke placeholder yang ADA di dokumen
  // (varian "@en" memakai nilai Inggris: tanggal diterjemahkan). Yang isinya sudah sama dilewati.
  function applyToDoc(push, cb) {
    if (!push || !Object.keys(push).length) { cb(0); return; }
    window.Asc.scope.push = push;
    window.Asc.plugin.callCommand(function () {
      var p = Asc.scope.push, doc = Api.GetDocument(), all = [], n = 0;
      try { all = doc.GetAllContentControls() || []; } catch (e) { all = []; }
      for (var i = 0; i < all.length; i++) {
        var tag = ""; try { tag = all[i].GetTag() || ""; } catch (e1) {}
        if (tag.indexOf("var:") !== 0) continue;
        var key = tag.slice(4), en = false;
        if (key.length > 3 && key.slice(-3) === "@en") { en = true; key = key.slice(0, -3); }
        var item = p[key]; if (!item) continue;
        var val = String(en ? item.en : item.id);
        var cur = ""; try { cur = all[i].GetRange().GetText() || ""; } catch (e2) {}
        if (cur === val) continue;
        try { all[i].RemoveAllElements(); all[i].AddText(val); n++; } catch (e3) {}
      }
      return n;
    }, false, true, function (n) { cb(n || 0); });
  }

  // ---- Variabel lain (di luar field formulir) — perilaku lama dipertahankan ----
  function partyWordKey(k) {
    var ks = (party && party.view && party.view.wordKeys) || [];
    for (var i = 0; i < ks.length; i++) if (ks[i] === k) return true;
    return false;
  }
  function readVars() { readDocVars(function (res) { vars = res || {}; renderVars(); }); }
  function renderVars() {
    var box = $("vars"); box.innerHTML = "";
    var keys = Object.keys(vars).filter(function (k) { return !partyWordKey(k); }).sort();
    if (!keys.length) {
      var all = Object.keys(vars).length;
      box.innerHTML = '<div class="hint">' + (all ? "Semua data di dokumen ini sudah diatur lewat formulir di atas." : "Tidak ada data yang bisa diubah di dokumen ini.") + "</div>";
      $("saveVars").style.display = "none"; return;
    }
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

  // ---- Data pihak & kontrak (formulir) ----
  function valuesOf(view) {
    var out = {};
    view.groups.forEach(function (g) { g.fields.forEach(function (f) { out[f.key] = f.value || ""; }); });
    out.party2LogoUrl = view.logo.url || ""; out.party2LogoKey = view.logo.key || ""; out.party2LogoMimeType = view.logo.mimeType || "";
    return out;
  }
  function anyDirty() { return Object.keys(dirty).length > 0; }
  function inputOf(key) { return document.querySelector('#party [data-k="' + key + '"]'); }

  function updateBar() {
    var bar = $("partyBar"); if (!party) { bar.style.display = "none"; return; }
    bar.style.display = "";
    $("partySave").disabled = !(party.editable && anyDirty()) || busy || uploading;
    $("partyReset").disabled = !anyDirty() || busy;
  }
  function setBanner(text) {
    var b = $("partyBanner"); b.textContent = text || ""; b.style.display = text ? "block" : "none";
  }
  function absUrl(u) { try { return new URL(u, opts.panelUrl).href; } catch (e) { return u; } }

  function onEdit(key, inp) {
    var same = (inp.value || "").trim() === (loaded[key] || "").trim();
    if (same) delete dirty[key]; else dirty[key] = true;
    inp.classList.toggle("dirty", !same); inp.classList.remove("invalid");
    updateBar();
  }

  function fieldEl(f, ro) {
    var w = el("div", "f" + (f.size === "narrow" ? " narrow" : ""));
    var l = el("label", "", f.label);
    if (f.required) l.appendChild(el("span", "req", " *"));
    w.appendChild(l);
    var inp;
    if (f.type === "select") {
      inp = el("select");
      (f.options || []).forEach(function (o) { var op = el("option", "", o); op.value = o; inp.appendChild(op); });
      inp.value = f.value;
    } else {
      inp = el("input"); inp.type = f.type === "date" ? "date" : "text"; inp.value = f.value;
      if (f.placeholder) { inp.placeholder = f.placeholder; inp.title = f.placeholder; }
    }
    inp.setAttribute("data-k", f.key); inp.disabled = ro;
    inp.oninput = inp.onchange = function () { onEdit(f.key, inp); };
    w.appendChild(inp);
    return w;
  }
  function lockedEl(lk) {
    var w = el("div", "f");
    var l = el("label", "", lk.label + " "); l.appendChild(el("span", "lockb", "🔒 terkunci")); l.title = lk.note; w.appendChild(l);
    var v = el("div", "locked", lk.value || "-"); v.title = lk.note; w.appendChild(v);
    w.appendChild(el("div", "hint", lk.note));
    return w;
  }
  function logoEl(lg, ro) {
    var w = el("div", "f");
    var l = el("label", "", lg.label); l.title = lg.hint; w.appendChild(l);
    var holder = el("div", "logo");
    var prev = el("img", "logoimg"); prev.alt = lg.label; prev.style.display = logoNow.url ? "" : "none"; if (logoNow.url) prev.src = absUrl(logoNow.url);
    holder.appendChild(prev);
    var file = el("input"); file.type = "file"; file.accept = "image/png,image/jpeg"; file.disabled = ro;
    file.onchange = function () { var f = file.files && file.files[0]; if (f) uploadLogo(f, file, prev, del); };
    holder.appendChild(file);
    var del = el("button", "", "Hapus"); del.type = "button"; del.title = "Hapus logo Pihak Kedua"; del.disabled = ro; del.style.display = logoNow.url ? "" : "none";
    del.onclick = function () { setLogo({ url: "", key: "", mimeType: "" }); prev.style.display = "none"; del.style.display = "none"; file.value = ""; };
    holder.appendChild(del);
    w.appendChild(holder);
    w.appendChild(el("div", "hint", lg.hint));
    return w;
  }
  function setLogo(next) {
    logoNow = next;
    var d = (logoNow.url !== (loaded.party2LogoUrl || "")) || (logoNow.key !== (loaded.party2LogoKey || "")) || (logoNow.mimeType !== (loaded.party2LogoMimeType || ""));
    LOGO_KEYS.forEach(function (k) { if (d) dirty[k] = true; else delete dirty[k]; });
    updateBar();
  }
  function uploadLogo(file, fileInp, prev, del) {
    if (!/^image\/(png|jpe?g)$/.test(file.type)) { sayParty("Logo harus berformat PNG atau JPEG", false); fileInp.value = ""; return; }
    if (file.size > 5 * 1024 * 1024) { sayParty("Logo terlalu besar (maks 5 MB)", false); fileInp.value = ""; return; }
    uploading = true; updateBar(); sayParty("Mengunggah logo…", true);
    var rd = new FileReader();
    rd.onerror = function () { uploading = false; updateBar(); sayParty("Gagal membaca berkas logo", false); };
    rd.onload = function () {
      api("/party-logo", { dataUrl: String(rd.result), fileName: file.name })
        .then(function (r) {
          setLogo({ url: r.url, key: r.key, mimeType: r.mimeType });
          prev.src = absUrl(r.url); prev.style.display = ""; del.style.display = "";
          sayParty('Logo terunggah — klik "Simpan Perubahan".', true);
        })
        .catch(function (e) { fileInp.value = ""; sayParty("Gagal mengunggah logo: " + e.message, false); })
        .then(function () { uploading = false; updateBar(); });
    };
    rd.readAsDataURL(file);
  }

  function renderParty() {
    var v = party.view, box = $("party"); box.innerHTML = "";
    var ro = !party.editable;
    setBanner(ro ? (party.reason || "Data tidak dapat diubah.") : "");
    v.groups.forEach(function (g) {
      var sec = el("div", "grp"); sec.appendChild(el("div", "grpt", g.title));
      var i = 0;
      while (i < g.fields.length) {
        var f = g.fields[i];
        if (f.row) {
          var r = el("div", "row2");
          while (i < g.fields.length && g.fields[i].row === f.row) { r.appendChild(fieldEl(g.fields[i], ro)); i++; }
          sec.appendChild(r);
        } else { sec.appendChild(fieldEl(f, ro)); i++; }
      }
      if (g.id === "kontrak") sec.appendChild(lockedEl(v.locked));
      if (g.id === "pihak2") sec.appendChild(logoEl(v.logo, ro));
      box.appendChild(sec);
    });
    updateBar();
  }
  // Muat tampilan dari server -> form dan "nilai terakhir dari server" disegarkan, perubahan belum disimpan dibuang.
  function applyServerView(view, editable, reason) {
    party = { view: view, editable: editable !== false, reason: reason };
    loaded = valuesOf(view); dirty = {};
    logoNow = { url: view.logo.url || "", key: view.logo.key || "", mimeType: view.logo.mimeType || "" };
    renderParty(); renderVars();
  }

  function collect() {
    var fields = {}, base = {};
    Object.keys(dirty).forEach(function (k) {
      if (LOGO_KEYS.indexOf(k) >= 0) return;
      var inp = inputOf(k); if (!inp) return;
      fields[k] = inp.value; base[k] = loaded[k] || "";
    });
    if (LOGO_KEYS.some(function (k) { return dirty[k]; })) {
      fields.party2LogoUrl = logoNow.url; fields.party2LogoKey = logoNow.key; fields.party2LogoMimeType = logoNow.mimeType;
      LOGO_KEYS.forEach(function (k) { base[k] = loaded[k] || ""; });
    }
    return { fields: fields, base: base };
  }
  // Validasi yang sama dengan modal (server memeriksa ulang): judul & nama Pihak Kedua wajib, tanggal tidak boleh kosong.
  function validateLocal(fields) {
    var rules = [
      ["title", "Judul dokumen tidak boleh kosong"], ["party2Name", "Nama Pihak Kedua tidak boleh kosong"],
      ["startDate", "Tanggal tidak boleh kosong"], ["endDate", "Tanggal berakhir kontrak tidak boleh kosong"]
    ];
    for (var i = 0; i < rules.length; i++) {
      var k = rules[i][0];
      if (fields[k] !== undefined && !String(fields[k]).trim()) {
        var inp = inputOf(k); if (inp) { inp.classList.add("invalid"); inp.focus(); }
        sayParty(rules[i][1], false); return false;
      }
    }
    return true;
  }

  function saveParty() {
    if (busy || !party || !party.editable) return;
    var c = collect();
    if (!Object.keys(c.fields).length) { sayParty("Tidak ada yang berubah.", true); return; }
    if (!validateLocal(c.fields)) return;
    busy = true; updateBar(); sayParty("Menyimpan…", true);
    api("/party", c).then(function (res) {
      applyServerView(res.view, true);
      if (!res.changed || !res.changed.length) { sayParty("Tidak ada yang berubah.", true); return; }
      return new Promise(function (resolve) { applyToDoc(res.push || {}, resolve); }).then(function (n) {
        prevDoc = null; // pembacaan berikutnya jadi patokan baru
        var msg = "Tersimpan ke database & aplikasi utama" + (n ? "; " + n + " placeholder di dokumen diperbarui." : ".");
        if (res.changed.indexOf("party2Name") >= 0 && state && state.ttd) msg += " Catatan: nama di blok tanda tangan tidak berubah otomatis — bangun ulang lewat bagian “Kop & tanda tangan” bila perlu.";
        sayParty(msg, true); say(msg, true);
        notifyHost(); readVars();
      });
    }).catch(function (e) {
      if (e.status === 409 && e.data && e.data.view) { applyServerView(e.data.view, true); notifyHost(); }
      sayParty("Gagal: " + e.message, false);
    }).then(function () { busy = false; updateBar(); });
  }
  function resetParty() { if (party) { applyServerView(party.view, party.editable, party.reason); sayParty("Perubahan dibatalkan.", true); } }

  function loadParty() {
    return api("/party").then(function (res) {
      if (!res.available) { setBanner("Data pihak & kontrak tidak tersedia untuk dokumen ini."); $("party").innerHTML = ""; return; }
      applyServerView(res.view, res.editable, res.reason);
    }).catch(function (e) {
      $("party").innerHTML = '<div class="hint">Gagal memuat data pihak &amp; kontrak: ' + String(e.message).replace(/[<>&]/g, "") + "</div>";
    });
  }

  // ---- Sinkron berkala dua arah (database <-> placeholder Word) ----
  function tick(force) {
    if (!party || busy || syncing || uploading) return;
    if (document.hidden && !force) return;
    if (anyDirty()) return; // jangan mengganggu pengguna yang sedang mengisi formulir
    syncing = true;
    readDocVars(function (vals) {
      // "Stabil" = isinya sama dengan pembacaan sebelumnya: sudah berhenti diketik, aman ditarik ke database.
      var stable = [];
      if (prevDoc) Object.keys(vals).forEach(function (k) { if (prevDoc[k] === vals[k]) stable.push(k); });
      api("/party-sync", { docValues: vals, pullKeys: stable }).then(function (res) {
        var viewChanged = res.view && party.view && res.view.rev !== party.view.rev;
        var roChanged = (res.editable === false) !== !party.editable;
        if ((viewChanged || roChanged) && !anyDirty()) applyServerView(res.view, res.editable !== false, res.reason);
        return new Promise(function (resolve) { applyToDoc(res.push || {}, resolve); }).then(function (n) {
          prevDoc = vals;
          Object.keys(res.push || {}).forEach(function (k) { prevDoc[k] = res.push[k].id; });
          if (res.pulled && res.pulled.length) { notifyHost(); sayParty("Perubahan di dokumen disalin ke data kontrak (" + res.pulled.length + ").", true); }
          else if (n) sayParty("Dokumen diperbarui mengikuti data kontrak (" + n + " placeholder).", true);
          if (n || (res.pulled && res.pulled.length)) readVars();
        });
      }).catch(function () { /* sinkron berkala diam-diam; coba lagi di putaran berikut */ }).then(function () { syncing = false; });
    });
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
    $("partySave").onclick = saveParty;
    $("partyReset").onclick = resetParty;
    loadState();
    loadParty().then(function () { readVars(); tick(true); });
    if (timer) clearInterval(timer);
    timer = setInterval(function () { tick(false); }, POLL_MS);
    window.addEventListener("focus", function () { tick(false); });
    document.addEventListener("visibilitychange", function () { if (!document.hidden) tick(false); });
  };
  window.Asc.plugin.button = function () { };
})(window);
