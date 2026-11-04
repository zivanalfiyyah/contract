// Pembersih HTML hasil copy-paste (Word, Google Docs, web) — dipakai SEMUA
// editor dokumen supaya perilaku paste-nya sama.
//
// Dua sumbu pilihan:
//  - target  "rich": editor template (contentEditable HTML) — tabel/list/heading
//                    tetap berbentuk tabel/list/heading.
//            "flat": editor .docx upload — penyimpanan hanya mengenal <p> berisi
//                    teks + format karakter, jadi tabel diratakan jadi paragraf
//                    dan list diberi PENANDA (<span data-list-marker>) yang saat
//                    simpan diubah jadi list Word sungguhan (lihat docx-patch.ts).
//  - keep    "merge":  ikuti format dokumen tujuan (Word: Merge Formatting) —
//                      struktur & penekanan dipertahankan; font/ukuran/warna sumber
//                      dibuang.
//            "source": pertahankan format sumber (Word: Keep Source Formatting) —
//                      font, ukuran, dan warna ikut (latar sel/teks hanya utk "rich").
// "Keep Text Only" tidak lewat sini (cukup teks polos dari clipboard).

export type PasteMode = "rich" | "flat";
export type PasteKeep = "merge" | "source";

interface Ctx { mode: PasteMode; source: boolean; listBase: string; listN: number }

const BLOCK_TAGS = new Set([
  "p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "li", "table", "thead", "tbody", "tfoot",
  "tr", "td", "th", "blockquote", "hr", "pre", "section", "article", "header", "footer", "center",
]);
const SKIP_TAGS = new Set(["style", "script", "meta", "link", "title", "xml", "head", "noscript", "template"]);

// Gaya sel tabel — sama dgn TABLE_TD_STYLE di toolbar dokumen (App.tsx).
const TD_STYLE = "border:1px solid #64748b;padding:6px 8px;";
const NBSP4 = "    ";

const escapeText = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const escapeAttr = (s: string) => escapeText(s).replace(/"/g, "&quot;");

/** Ambil hanya bagian antara StartFragment/EndFragment (pembungkus clipboard Windows/Word). */
function sliceFragment(html: string): string {
  const a = html.indexOf("<!--StartFragment-->");
  const b = html.indexOf("<!--EndFragment-->");
  if (a === -1 || b <= a) return html;
  const fragment = html.slice(a + "<!--StartFragment-->".length, b);
  // Blok <style> (gaya Word: Heading, Body Text, ...) letaknya SEBELUM StartFragment di <head>.
  // Kalau ikut terpotong, semua format yang berasal dari gaya hilang — bawa ke depan fragmen.
  const styles = (html.slice(0, a).match(/<style[\s\S]*?<\/style>/gi) || []).join("");
  return styles ? `<html><head>${styles}</head><body>${fragment}</body></html>` : fragment;
}

function removeComments(root: Node) {
  const doc = root.ownerDocument!;
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_COMMENT);
  const found: Node[] = [];
  while (walker.nextNode()) found.push(walker.currentNode);
  found.forEach((n) => n.parentNode?.removeChild(n));
}

function hasBlockDescendant(el: Element): boolean {
  return Array.from(el.querySelectorAll("*")).some((c) => BLOCK_TAGS.has(c.tagName.toLowerCase()));
}

// ---------------------------------------------------------------------------
// List Word ("MsoListParagraph" + mso-list) -> <ul>/<ol>/<li> sungguhan.
// ---------------------------------------------------------------------------
const isWordListP = (el: Element) => el.tagName === "P" && /mso-list\s*:\s*l\d+/i.test(el.getAttribute("style") || "");

/**
 * Jenis penanda list Word:
 *  - "ol"   : angka biasa ("1." / "2)") -> <ol>
 *  - "ul"   : simbol bullet (·, o, §, ...) -> <ul>
 *  - "text" : penomoran yang TIDAK bisa ditiru <ol> apa adanya — bertingkat ("1.1"),
 *             huruf/romawi ("I.", "a)"), "(1)", dsb. Paragrafnya dibiarkan biasa dgn
 *             teks penandanya tetap ada, daripada nomornya berubah jadi "1." / "•".
 */
function classifyMarker(m: string): "ul" | "ol" | "text" {
  if (!m) return "ul";
  if (/^[0-9]+[.)]$/.test(m)) return "ol";
  if (/[a-z0-9]/i.test(m) && m.length > 1) return "text";
  return "ul";
}

function convertWordLists(root: HTMLElement) {
  const doc = root.ownerDocument;
  const listPs = Array.from(root.querySelectorAll("p")).filter(isWordListP);
  if (!listPs.length) return;
  const parents = new Set<Element>(listPs.map((p) => p.parentElement!).filter(Boolean));
  parents.forEach((parent) => {
    let stack: { level: number; list: HTMLElement }[] = [];
    for (const el of Array.from(parent.children)) {
      if (!isWordListP(el)) { stack = []; continue; }
      const style = el.getAttribute("style") || "";
      const level = parseInt(/level(\d+)/i.exec(style)?.[1] || "1", 10);
      const markerEl = Array.from(el.querySelectorAll("span")).find((s) => /mso-list\s*:\s*ignore/i.test(s.getAttribute("style") || ""));
      const marker = (markerEl?.textContent || "").replace(/ /g, " ").trim();
      const kind = classifyMarker(marker);
      if (kind === "text") { stack = []; continue; } // dibiarkan: teks penanda tetap di paragraf
      const tag = kind;
      while (stack.length && stack[stack.length - 1].level > level) stack.pop();
      let top = stack[stack.length - 1];
      const makeList = () => {
        const list = doc.createElement(tag);
        const host = stack.length ? stack[stack.length - 1].list.lastElementChild : null;
        if (host) host.appendChild(list); else parent.insertBefore(list, el);
        return list;
      };
      // List bernomor yang dimulai BUKAN dari 1 (mis. potongan "3." "4.") tidak bisa
      // ditiru <ol> (akan jadi "1." "2.") -> dibiarkan sebagai teks biasa.
      const startsNew = !top || top.level < level || top.list.tagName.toLowerCase() !== tag;
      if (startsNew && tag === "ol" && parseInt(marker, 10) !== 1) { stack = []; continue; }
      markerEl?.remove();
      if (!top || top.level < level) {
        top = { level, list: makeList() };
        stack.push(top);
      } else if (top.list.tagName.toLowerCase() !== tag) {
        stack.pop();
        top = { level, list: makeList() };
        stack.push(top);
      }
      const li = doc.createElement("li");
      li.innerHTML = el.innerHTML;
      top.list.appendChild(li);
      el.remove();
    }
  });
}

// ---------------------------------------------------------------------------
// Format karakter
// ---------------------------------------------------------------------------
function charFormat(el: HTMLElement) {
  const tag = el.tagName.toLowerCase();
  const st = el.style;
  const fw = (st.fontWeight || "").toLowerCase();
  const fwNum = parseInt(fw, 10);
  let bold = tag === "b" || tag === "strong";
  // Google Docs membungkus SELURUH isi dgn <b style="font-weight:normal"> — itu bukan tebal.
  if (fw === "normal" || (!isNaN(fwNum) && fwNum < 600)) bold = false;
  else if (fw === "bold" || fw === "bolder" || (!isNaN(fwNum) && fwNum >= 600)) bold = true;
  let italic = tag === "i" || tag === "em";
  const fs = (st.fontStyle || "").toLowerCase();
  if (fs === "normal") italic = false;
  else if (fs === "italic" || fs === "oblique") italic = true;
  const deco = `${st.textDecoration || ""} ${st.textDecorationLine || ""}`.toLowerCase();
  let underline = tag === "u" || deco.includes("underline");
  const strike = tag === "s" || tag === "strike" || tag === "del" || deco.includes("line-through");
  if (tag === "a") underline = false; // garis bawah tautan datang dari <a> itu sendiri
  return { bold, italic, underline, strike };
}

// ---- Gaya sumber (hanya mode "source") -------------------------------------
const NAMED_COLORS_IGNORED = new Set(["windowtext", "auto", "inherit", "initial", "unset", "currentcolor", "transparent", "black", "window"]);

// Nama warna CSS yang umum — diselesaikan tanpa DOM supaya hasilnya pasti.
const BASIC_COLORS: Record<string, string> = {
  red: "ff0000", blue: "0000ff", green: "008000", yellow: "ffff00", orange: "ffa500", purple: "800080",
  gray: "808080", grey: "808080", silver: "c0c0c0", maroon: "800000", navy: "000080", teal: "008080",
  olive: "808000", lime: "00ff00", aqua: "00ffff", cyan: "00ffff", fuchsia: "ff00ff", magenta: "ff00ff",
  white: "ffffff", brown: "a52a2a", pink: "ffc0cb", darkred: "8b0000", darkblue: "00008b", darkgreen: "006400",
  gold: "ffd700", crimson: "dc143c", indigo: "4b0082", violet: "ee82ee",
};

/** Hex 6 digit (tanpa #) dari "#rgb" / "#rrggbb" / "rgb(a)(...)"; "" kalau bukan. TIDAK rekursif. */
function toHex(s: string): string {
  let m = /^#([0-9a-f]{3})$/.exec(s);
  if (m) return m[1].split("").map((c) => c + c).join("");
  m = /^#([0-9a-f]{6})$/.exec(s);
  if (m) return m[1];
  m = /^rgba?\(\s*(\d{1,3})\s*[, ]\s*(\d{1,3})\s*[, ]\s*(\d{1,3})(?:\s*[,/]\s*([\d.]+%?))?\s*\)$/.exec(s);
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return "";
    return [m[1], m[2], m[3]].map((n) => Math.min(255, parseInt(n, 10)).toString(16).padStart(2, "0")).join("");
  }
  return "";
}

/** Warna CSS aman -> "#rrggbb" ("" kalau tak dikenal / sama dgn warna bawaan hitam). */
function safeColor(v: string): string {
  const s = (v || "").trim().toLowerCase();
  if (!s || NAMED_COLORS_IGNORED.has(s)) return "";
  let hex = toHex(s);
  if (!hex && /^[a-z]{3,20}$/.test(s)) {
    hex = BASIC_COLORS[s] || "";
    if (!hex) {
      // Nama lain: biarkan browser menyelesaikannya lewat elemen sementara (sekali, tanpa rekursi).
      try {
        const probe = document.createElement("span");
        probe.style.color = s;
        if (probe.style.color) {
          document.body.appendChild(probe);
          hex = toHex(getComputedStyle(probe).color.trim().toLowerCase());
          probe.remove();
        }
      } catch { hex = ""; }
    }
  }
  if (!hex || hex === "000000") return "";
  return `#${hex}`;
}

function safeFontFamily(v: string): string {
  const s = (v || "").replace(/["]/g, "'").trim();
  if (!s || s.length > 120 || !/^[A-Za-z0-9 ,'\-_.]+$/.test(s)) return "";
  return s;
}

function safeFontSizePt(v: string): string {
  const m = /^(\d+(?:\.\d+)?)(pt|px)$/i.exec((v || "").trim());
  if (!m) return "";
  let n = parseFloat(m[1]);
  if (m[2].toLowerCase() === "px") n = n * 0.75;
  if (!(n >= 4 && n <= 96)) return "";
  return `${Math.round(n * 2) / 2}pt`;
}

function sourceStyle(el: HTMLElement, ctx: Ctx): string {
  if (!ctx.source) return "";
  const st = el.style;
  const parts: string[] = [];
  const face = safeFontFamily(st.fontFamily || el.getAttribute("face") || "");
  if (face) parts.push(`font-family:${face}`);
  const size = safeFontSizePt(st.fontSize);
  if (size) parts.push(`font-size:${size}`);
  const color = safeColor(st.color || el.getAttribute("color") || "");
  if (color) parts.push(`color:${color}`);
  // Latar/sorot hanya utk editor template; penyimpanan .docx belum menulisnya.
  if (ctx.mode === "rich") {
    const bg = safeColor(st.backgroundColor || "");
    if (bg && bg !== "#ffffff") parts.push(`background-color:${bg}`);
  }
  return parts.length ? parts.join(";") + ";" : "";
}

/**
 * Bungkus isi dgn format karakter milik elemen `el` (tebal/miring/garis bawah/coret
 * + gaya sumber utk mode "source"). Dipakai utk elemen inline DAN elemen blok
 * (paragraf/judul/butir list): di Word, tebal judul dan font isi sering menempel
 * di paragrafnya lewat gaya (class), bukan di <span> di dalamnya.
 */
function decorate(inner: string, el: HTMLElement, ctx: Ctx): string {
  if (!inner.replace(/<[^>]*>|&nbsp;|\s/g, "")) return inner; // jangan bungkus yg kosong
  let out = inner;
  const style = sourceStyle(el, ctx);
  if (style) out = `<span style="${escapeAttr(style)}">${out}</span>`;
  const f = charFormat(el);
  if (f.strike) out = `<s>${out}</s>`;
  if (f.underline) out = `<u>${out}</u>`;
  if (f.italic) out = `<i>${out}</i>`;
  if (f.bold) out = `<b>${out}</b>`;
  return out;
}

/** Isi inline (teks + tebal/miring/dll) dari satu node. */
function inlineOf(node: Node, ctx: Ctx): string {
  if (node.nodeType === Node.TEXT_NODE) {
    return escapeText((node as Text).data.replace(/[ \t\r\n\f]+/g, " "));
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return "";
  const el = node as HTMLElement;
  const tag = el.tagName.toLowerCase();
  if (SKIP_TAGS.has(tag)) return "";
  if (tag === "br") return "<br>";
  if (tag === "img") {
    const src = el.getAttribute("src") || "";
    // Word menempel gambar sebagai file:///... (tak bisa dibuka) — hanya data:/http(s) yang dipakai.
    if (ctx.mode === "rich" && /^(data:image\/|https?:\/\/)/i.test(src)) {
      return `<img src="${escapeAttr(src)}" alt="${escapeAttr(el.getAttribute("alt") || "")}" style="max-width:100%;height:auto;" />`;
    }
    return "";
  }
  const kids = Array.from(el.childNodes).map((c) => inlineOf(c, ctx)).join("");
  if (BLOCK_TAGS.has(tag)) {
    // Blok di dalam konteks inline (mis. judul <h1> di dalam <li>): format gayanya tetap dibawa;
    // judul tebal secara bawaan kecuali gayanya menyatakan lain.
    const isHeading = /^h[1-6]$/.test(tag);
    let body = decorate(kids, el, ctx);
    if (isHeading && !el.style.fontWeight && kids.replace(/<[^>]*>|&nbsp;|\s/g, "")) body = `<b>${body}</b>`;
    return body + (tag === "p" || tag === "div" || tag === "li" || isHeading ? " " : "");
  }
  let out = decorate(kids, el, ctx);
  if (tag === "a" && ctx.mode === "rich" && kids.replace(/<[^>]*>|&nbsp;|\s/g, "")) {
    const href = el.getAttribute("href") || "";
    if (/^(https?:|mailto:|tel:)/i.test(href)) {
      out = `<a href="${escapeAttr(href)}" target="_blank" rel="noopener noreferrer">${out}</a>`;
    }
  }
  return out;
}

const isBlank = (html: string) => !html.replace(/<(?!img\b)[^>]*>|&nbsp;|\s/g, "");

/** Buang spasi biasa di awal/akhir isi paragraf, juga yang terselip tepat di balik tag pembuka/penutup (mis. "<b> Judul</b> "). */
const trimInline = (html: string) =>
  html.replace(/^((?:<[^>]*>)*) +/, "$1").replace(/ +((?:<\/[^>]*>)*)$/, "$1");

function alignOf(el: HTMLElement): string {
  const a = (el.style.textAlign || el.getAttribute("align") || "").toLowerCase();
  return a === "center" || a === "right" || a === "justify" ? a : "";
}

function para(inner: string, align = ""): string {
  const content = isBlank(inner) ? "<br>" : trimInline(inner);
  return `<p${align ? ` style="text-align:${align};"` : ""}>${content}</p>`;
}

// ---------------------------------------------------------------------------
// Blok
// ---------------------------------------------------------------------------
function blocksOf(container: Node, ctx: Ctx): string {
  let out = "";
  let buf = "";
  const flush = () => {
    if (!isBlank(buf)) out += para(buf);
    buf = "";
  };
  container.childNodes.forEach((child) => {
    if (child.nodeType !== Node.ELEMENT_NODE) { buf += inlineOf(child, ctx); return; }
    const el = child as HTMLElement;
    const tag = el.tagName.toLowerCase();
    if (SKIP_TAGS.has(tag)) return;
    // Pembungkus inline yg berisi blok (mis. <b> raksasa Google Docs) = wadah, bukan teks.
    if (BLOCK_TAGS.has(tag) || hasBlockDescendant(el)) {
      flush();
      out += blockEl(el, ctx);
    } else {
      buf += inlineOf(el, ctx);
    }
  });
  flush();
  return out;
}

function blockEl(el: HTMLElement, ctx: Ctx): string {
  const tag = el.tagName.toLowerCase();
  const inner = () => Array.from(el.childNodes).map((c) => inlineOf(c, ctx)).join("");
  if (tag === "p" || tag === "pre" || tag === "li") return para(decorate(inner(), el, ctx), alignOf(el));
  if (tag === "div" || tag === "section" || tag === "article" || tag === "header" || tag === "footer" || tag === "center") {
    if (hasBlockDescendant(el)) return blocksOf(el, ctx);
    return para(decorate(inner(), el, ctx), alignOf(el) || (tag === "center" ? "center" : ""));
  }
  if (/^h[1-6]$/.test(tag)) {
    const html = decorate(inner(), el, ctx);
    if (isBlank(html)) return "";
    if (ctx.mode === "flat") return para(charFormat(el).bold ? html : `<b>${html}</b>`, alignOf(el));
    const h = tag === "h4" || tag === "h5" || tag === "h6" ? "h4" : "h3";
    return `<${h}>${html}</${h}>`;
  }
  if (tag === "ul" || tag === "ol") return listOut(el, ctx, 0);
  if (tag === "table") return tableOut(el as HTMLTableElement, ctx);
  if (tag === "blockquote") {
    const html = blocksOf(el, ctx);
    return ctx.mode === "rich" ? `<blockquote>${html}</blockquote>` : html;
  }
  if (tag === "hr") return ctx.mode === "rich" ? "<hr>" : "";
  // thead/tbody/tr/td/th liar di luar tabel, dsb.
  return blocksOf(el, ctx);
}

/** Isi inline sebuah <li> TANPA sub-list di dalamnya + daftar sub-list-nya. */
function splitLi(li: HTMLElement, ctx: Ctx): { inline: string; nested: HTMLElement[] } {
  let inline = "";
  const nested: HTMLElement[] = [];
  li.childNodes.forEach((c) => {
    if (c.nodeType === Node.ELEMENT_NODE) {
      const t = (c as HTMLElement).tagName.toLowerCase();
      if (t === "ul" || t === "ol") { nested.push(c as HTMLElement); return; }
      if ((t === "p" || t === "div") && inline && !isBlank(inline)) inline += "<br>";
    }
    inline += inlineOf(c, ctx);
  });
  return { inline: decorate(trimInline(inline.replace(/^(<br>)+|(<br>)+$/g, "")), li, ctx), nested };
}

/** Jenis nomor list HTML: "1" | "a" | "A" | "i" | "I" (dari atribut type atau CSS list-style-type). */
function listNumberType(list: HTMLElement): "1" | "a" | "A" | "i" | "I" {
  const t = (list.getAttribute("type") || "").trim();
  if (t === "a" || t === "A" || t === "i" || t === "I") return t;
  const ls = (list.style.listStyleType || "").toLowerCase();
  if (ls === "lower-alpha" || ls === "lower-latin") return "a";
  if (ls === "upper-alpha" || ls === "upper-latin") return "A";
  if (ls === "lower-roman") return "i";
  if (ls === "upper-roman") return "I";
  return "1";
}

function numberLabel(n: number, t: string): string {
  if (t === "1") return String(n);
  if (t === "a" || t === "A") {
    let s = "";
    let x = n;
    while (x > 0) { x--; s = String.fromCharCode(97 + (x % 26)) + s; x = Math.floor(x / 26); }
    return t === "A" ? s.toUpperCase() : s;
  }
  const map: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let x = n;
  let s = "";
  for (const [v, r] of map) while (x >= v) { s += r; x -= v; }
  return t === "I" ? s.toUpperCase() : s;
}

/**
 * Butir sebuah <ul>/<ol>. Daftar bersarang yang menempel LANGSUNG di <ul>/<ol>
 * (tanpa <li> pembungkus — dihasilkan LibreOffice dkk) dianggap anak butir sebelumnya,
 * bukan dibuang.
 */
function listEntries(list: HTMLElement): { li?: HTMLElement; nested: HTMLElement[] }[] {
  const out: { li?: HTMLElement; nested: HTMLElement[] }[] = [];
  for (const c of Array.from(list.children) as HTMLElement[]) {
    const t = c.tagName.toLowerCase();
    if (t === "li") out.push({ li: c, nested: [] });
    else if (t === "ul" || t === "ol") {
      if (out.length) out[out.length - 1].nested.push(c);
      else out.push({ nested: [c] });
    }
  }
  return out;
}

function listOut(list: HTMLElement, ctx: Ctx, level: number): string {
  const ordered = list.tagName.toLowerCase() === "ol";
  const numType = ordered ? listNumberType(list) : "1";
  const entries = listEntries(list);
  const emptyParts = (): { inline: string; nested: HTMLElement[] } => ({ inline: "", nested: [] });
  if (ctx.mode === "rich") {
    const lis = entries.map(({ li, nested }) => {
      const parts = li ? splitLi(li, ctx) : emptyParts();
      const all = [...parts.nested, ...nested];
      return `<li${li ? "" : ' style="list-style:none;"'}>${parts.inline}${all.map((n) => listOut(n, ctx, level + 1)).join("")}</li>`;
    }).join("");
    let attrs = "";
    if (ordered) {
      if (numType !== "1") attrs += ` type="${numType}"`;
      const st = parseInt(list.getAttribute("start") || "", 10);
      if (st > 1) attrs += ` start="${st}"`;
    }
    const tag = ordered ? "ol" : "ul";
    return `<${tag}${attrs}>${lis}</${tag}>`;
  }
  // flat: tiap butir = paragraf. Bullet & angka desimal diberi PENANDA list: teks
  // penanda ("• " / "1. ") tampil di editor; saat simpan, penanda di awal paragraf
  // dibuang dan diganti list Word sungguhan (numPr). Penomoran huruf/romawi ("I.", "a.")
  // belum bisa jadi list Word asli -> ditulis sbg teks biasa (tanpa penanda).
  const id = `${ctx.listBase}-${++ctx.listN}`;
  const realList = !ordered || numType === "1";
  let n = parseInt(list.getAttribute("start") || "1", 10);
  if (isNaN(n)) n = 1;
  let out = "";
  for (const { li, nested } of entries) {
    const parts = li ? splitLi(li, ctx) : emptyParts();
    if (li) {
      const label = ordered ? `${numberLabel(n, numType)}.` : "•";
      const prefix = `${NBSP4.repeat(level)}${label} `;
      n++;
      out += realList
        ? `<p><span data-list-marker="${ordered ? "decimal" : "bullet"}|${level}|${id}">${prefix}</span>${parts.inline}</p>`
        : `<p>${prefix}${parts.inline}</p>`;
    }
    out += [...parts.nested, ...nested].map((x) => listOut(x, ctx, level + 1)).join("");
  }
  return out;
}

function cellHtml(cell: HTMLElement, ctx: Ctx): string {
  const inner = blocksOf(cell, ctx);
  if (isBlank(inner)) return "&nbsp;";
  if ((inner.match(/<p[\s>]/g) || []).length === 1 && /^<p>[\s\S]*<\/p>$/.test(inner)) {
    return inner.slice(3, -4);
  }
  // Beberapa paragraf dlm satu sel -> dipisah <br> (tanpa <p> bersarang di <td>).
  return inner.replace(/<\/p><p[^>]*>/g, "<br>").replace(/^<p[^>]*>|<\/p>$/g, "");
}

function tableOut(table: HTMLTableElement, ctx: Ctx): string {
  const rows = Array.from(table.rows);
  if (ctx.mode === "flat") {
    return rows.map((row) => {
      const cells = Array.from(row.cells).map((c) =>
        Array.from(c.childNodes).map((n) => inlineOf(n, ctx)).join("").replace(/\s+/g, " ").trim());
      return para(cells.filter((x) => x).join(NBSP4));
    }).join("");
  }
  const trs = rows.map((row) => {
    const tds = Array.from(row.cells).map((cell) => {
      const tag = cell.tagName.toLowerCase() === "th" ? "th" : "td";
      const num = (name: string) => {
        const v = parseInt(cell.getAttribute(name) || "", 10);
        return v > 1 ? ` ${name}="${v}"` : "";
      };
      const bg = ctx.source ? safeColor(cell.style.backgroundColor || "") : "";
      const style = TD_STYLE + (bg && bg !== "#ffffff" ? `background-color:${bg};` : "");
      return `<${tag}${num("colspan")}${num("rowspan")} style="${style}">${cellHtml(cell, ctx)}</${tag}>`;
    }).join("");
    return `<tr>${tds}</tr>`;
  }).join("");
  return `<table style="border-collapse:collapse;width:100%;"><tbody>${trs}</tbody></table>`;
}

// ---------------------------------------------------------------------------
// Gaya dari blok <style> (class Word seperti p.MsoBodyText, h1, span.Judul)
//
// Word/LibreOffice mengirim format gaya (Heading, Body Text, ...) lewat daftar
// aturan CSS + class, BUKAN langsung di tiap elemen. Tanpa langkah ini tebal
// judul, rata teks, dan font isi yang berasal dari gaya ikut hilang. Aturan yang
// cocok disalin ke atribut style tiap elemen (style langsung tetap menang),
// hanya utk properti yang dipakai pembersih ini.
// ---------------------------------------------------------------------------
const INLINE_PROPS = new Set([
  "font-weight", "font-style", "font-family", "font-size", "color", "text-align",
  "background-color", "text-decoration", "text-decoration-line",
]);

function specificity(sel: string): number {
  const ids = (sel.match(/#[\w-]+/g) || []).length;
  const cls = (sel.match(/\.[\w-]+|\[[^\]]+\]|:[\w-]+/g) || []).length;
  const tags = (sel.match(/(^|[\s>+~])[a-z][\w-]*/gi) || []).length;
  return ids * 10000 + cls * 100 + tags;
}

function inlineStyleSheets(doc: Document) {
  interface Rule { sel: string; rank: number; style: CSSStyleDeclaration }
  const raw: { sel: string; spec: number; order: number; style: CSSStyleDeclaration }[] = [];
  let order = 0;
  for (const sheet of Array.from(doc.styleSheets)) {
    let list: CSSRuleList;
    try { list = sheet.cssRules; } catch { continue; }
    for (const r of Array.from(list)) {
      const sr = r as CSSStyleRule;
      if (typeof sr.selectorText !== "string" || !sr.style) continue; // @page, @font-face, @media, ...
      for (const part of sr.selectorText.split(",")) {
        const sel = part.trim();
        if (sel) raw.push({ sel, spec: specificity(sel), order: order++, style: sr.style });
      }
    }
  }
  if (!raw.length || !doc.body) return;
  raw.sort((a, b) => a.spec - b.spec || a.order - b.order);
  // Indeks per kunci (class terakhir / nama tag) supaya tiap elemen hanya mencoba aturan yang mungkin cocok.
  const byKey = new Map<string, Rule[]>();
  raw.forEach((r, rank) => {
    const compound = r.sel.split(/[\s>+~]+/).pop() || "";
    const cls = /\.([\w-]+)/.exec(compound);
    const tag = /^[a-z][\w-]*/i.exec(compound);
    const key = cls ? `.${cls[1]}` : tag ? tag[0].toLowerCase() : "*";
    const arr = byKey.get(key) || [];
    arr.push({ sel: r.sel, rank, style: r.style });
    byKey.set(key, arr);
  });
  doc.body.querySelectorAll<HTMLElement>("*").forEach((el) => {
    if (!el.style) return;
    const cand: Rule[] = [];
    const push = (k: string) => { const a = byKey.get(k); if (a) cand.push(...a); };
    push(el.tagName.toLowerCase());
    push("*");
    el.classList.forEach((c) => push(`.${c}`));
    if (!cand.length) return;
    cand.sort((a, b) => a.rank - b.rank);
    const add: Record<string, string> = {};
    for (const r of cand) {
      let hit = false;
      try { hit = el.matches(r.sel); } catch { hit = false; }
      if (!hit) continue;
      for (let i = 0; i < r.style.length; i++) {
        const p = r.style[i];
        if (INLINE_PROPS.has(p)) add[p] = r.style.getPropertyValue(p);
      }
    }
    for (const p in add) if (!el.style.getPropertyValue(p)) el.style.setProperty(p, add[p]);
  });
}

/**
 * Bersihkan HTML clipboard jadi HTML ringkas siap disisipkan (insertHTML).
 * Mengembalikan "" kalau tidak ada isi berguna (pemanggil jatuh ke teks polos).
 * Kalau hasilnya SATU paragraf polos, yang dikembalikan hanya isi inline-nya
 * supaya menyatu dgn paragraf tempat kursor berada (perilaku Word), bukan
 * memecah paragraf.
 */
export function cleanPastedHtml(html: string, mode: PasteMode = "rich", keep: PasteKeep = "merge"): string {
  if (!html || !/<[a-z][\s\S]*>/i.test(html)) return "";
  const doc = new DOMParser().parseFromString(sliceFragment(html), "text/html");
  const body = doc.body;
  if (!body) return "";
  inlineStyleSheets(doc); // SEBELUM <style> dibuang
  body.querySelectorAll("style,script,meta,link,title,xml,noscript,template").forEach((n) => n.remove());
  removeComments(body);
  convertWordLists(body);
  const ctx: Ctx = { mode, source: keep === "source", listBase: `L${Date.now().toString(36)}`, listN: 0 };
  const out = blocksOf(body, ctx);
  if (isBlank(out)) return "";
  const m = /^<p>([\s\S]*)<\/p>$/.exec(out);
  if (m && (out.match(/<p[\s>]/g) || []).length === 1) return m[1];
  return out;
}
