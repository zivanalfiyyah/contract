// Plugin ONLYOFFICE "Sisipkan Pasal": memilih pasal dari Library Klausul dan menyisipkannya sebagai
// content control bertag "clause:<id>" (dikenali sistem: pasal wajib/terlindungi dikunci server).
(function (window) {
  var clauses = [], values = {}, opts = {};

  function $(id) { return document.getElementById(id); }
  function say(text, ok) { var m = $("msg"); m.className = ok ? "ok" : "err"; m.textContent = text; if (ok) setTimeout(function () { m.className = ""; m.textContent = ""; }, 4000); }

  function render() {
    var q = ($("q").value || "").toLowerCase(), cat = $("cat").value, list = $("list");
    list.innerHTML = "";
    var shown = clauses.filter(function (c) { return (!cat || c.category === cat) && (!q || c.title.toLowerCase().indexOf(q) >= 0 || (c.category || "").toLowerCase().indexOf(q) >= 0); });
    if (!shown.length) { list.innerHTML = '<div id="empty">Tidak ada pasal yang cocok.</div>'; return; }
    shown.forEach(function (c) {
      var d = document.createElement("div"); d.className = "item";
      var b = document.createElement("b"); b.textContent = c.title; d.appendChild(b);
      var s = document.createElement("span"); s.className = "badge"; s.textContent = c.category || "Umum"; d.appendChild(s);
      if (c.mandatory) { var m = document.createElement("span"); m.className = "badge lock"; m.textContent = "Wajib"; d.appendChild(m); }
      if (c.protected) { var p = document.createElement("span"); p.className = "badge lock"; p.textContent = "Terlindungi"; d.appendChild(p); }
      var btn = document.createElement("button"); btn.textContent = "Sisipkan"; btn.onclick = function () { insert(c); }; d.appendChild(btn);
      list.appendChild(d);
    });
  }

  function load() {
    if (!opts.listUrl) { $("list").innerHTML = '<div id="empty">Plugin belum dikonfigurasi.</div>'; return; }
    fetch(opts.listUrl).then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); }).then(function (d) {
      clauses = d.clauses || []; values = d.values || {};
      var cats = {}; clauses.forEach(function (c) { cats[c.category || "Umum"] = 1; });
      var sel = $("cat"); Object.keys(cats).sort().forEach(function (k) { var o = document.createElement("option"); o.value = k; o.textContent = k; sel.appendChild(o); });
      render();
    }).catch(function (e) { $("list").innerHTML = ""; say("Gagal memuat daftar pasal: " + e.message, false); });
  }

  function insert(c) {
    var payload = { id: c.id, title: c.title, blocks: c.blocks, values: values };
    window.Asc.scope.clause = payload;
    window.Asc.plugin.callCommand(function () {
      // Berjalan di dalam editor (ES5, hanya Asc.scope & Api yang tersedia).
      var cl = Asc.scope.clause, doc = Api.GetDocument();
      function addRuns(par, runs) {
        for (var i = 0; i < runs.length; i++) {
          var r = runs[i];
          if (r.v) {
            var val = (cl.values && cl.values[r.v] !== undefined) ? String(cl.values[r.v]) : "{{" + r.v + "}}";
            try {
              var s = Api.CreateInlineLvlSdt();
              s.SetTag("var:" + r.v);
              s.AddText(val);
              par.AddInlineLvlSdt(s);
              continue;
            } catch (e) { /* jatuh ke teks biasa */ }
            var rr = Api.CreateRun(); rr.AddText(val); par.AddElement(rr);
            continue;
          }
          var parts = String(r.x).split("\n");
          for (var j = 0; j < parts.length; j++) {
            var run = Api.CreateRun();
            if (r.b) run.SetBold(true);
            if (r.i) run.SetItalic(true);
            if (r.u) run.SetUnderline(true);
            if (j > 0) run.AddLineBreak();
            run.AddText(parts[j]);
            par.AddElement(run);
          }
        }
      }
      function makePar(b) {
        var p = Api.CreateParagraph();
        if (b.t === "li") {
          p.SetIndLeft(720 * (b.level + 1));
          p.SetIndFirstLine(-360);
          var m = Api.CreateRun(); m.AddText(b.ordered ? (b.n + ". ") : "• "); p.AddElement(m);
        }
        if (b.jc === "center") p.SetJc("center"); else if (b.jc === "right") p.SetJc("right"); else if (b.jc === "both") p.SetJc("both");
        addRuns(p, b.runs || []);
        return p;
      }
      var sdt = Api.CreateBlockLvlSdt();
      try { sdt.SetTag("clause:" + cl.id); } catch (e1) { }
      try { sdt.SetAlias("Pasal: " + cl.title); } catch (e2) { }
      var content = sdt.GetContent();
      var head = content.GetElement(0);
      var hr = Api.CreateRun(); hr.SetBold(true); hr.AddText(cl.title); head.AddElement(hr);
      for (var k = 0; k < cl.blocks.length; k++) {
        var b = cl.blocks[k];
        if (b.t === "table") {
          var rows = b.rows, cols = 1;
          for (var a = 0; a < rows.length; a++) if (rows[a].length > cols) cols = rows[a].length;
          var tbl = Api.CreateTable(cols, rows.length);
          try { tbl.SetWidth("percent", 100); } catch (e3) { }
          try { tbl.SetTableBorderTop("single", 4, 0, 100, 116, 139); tbl.SetTableBorderBottom("single", 4, 0, 100, 116, 139); tbl.SetTableBorderLeft("single", 4, 0, 100, 116, 139); tbl.SetTableBorderRight("single", 4, 0, 100, 116, 139); tbl.SetTableBorderInsideH("single", 4, 0, 100, 116, 139); tbl.SetTableBorderInsideV("single", 4, 0, 100, 116, 139); } catch (e4) { }
          for (var ri = 0; ri < rows.length; ri++) {
            for (var ci = 0; ci < rows[ri].length; ci++) {
              var cell = tbl.GetRow(ri).GetCell(ci), cc = cell.GetContent(), paras = rows[ri][ci];
              for (var pi = 0; pi < paras.length; pi++) {
                var pp = makePar(paras[pi]);
                cc.Push(pp);
                if (pi === 0) cc.RemoveElement(0); // buang paragraf kosong bawaan sel
              }
            }
          }
          content.Push(tbl);
        } else {
          content.Push(makePar(b));
        }
      }
      var placed = false;
      try { placed = doc.InsertContent([sdt]) !== false; } catch (e5) { placed = false; }
      if (!placed) doc.Push(sdt);
    }, false, true, function () {
      say("Pasal “" + c.title + "” ditambahkan.", true);
    });
  }

  window.Asc.plugin.init = function () {
    opts = (window.Asc.plugin.info && window.Asc.plugin.info.options) || {};
    $("q").oninput = render; $("cat").onchange = render;
    load();
  };
  window.Asc.plugin.button = function () { };
})(window);
