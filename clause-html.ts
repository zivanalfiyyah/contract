import { parseFragment } from "parse5";

// ---------------------------------------------------------------------------
// Isi pasal (HTML dari editor rich-text, atau teks biasa lama) -> XML Word.
// Mendukung tag yang diizinkan editor pasal (RICH_ALLOWED di App.tsx):
//   p, div, br, b/strong, i/em, u, s, span(style), font(color,size), h3, h4,
//   ul/ol/li (bersarang), table/thead/tbody/tr/td/th (colspan, rowspan), blockquote, hr, a
// Gambar (<img>) belum didukung: dilewati (lihat `skipped` pada hasil).
// Variabel {{Key}} di mana pun (juga di dalam tabel/daftar) menjadi content control sebaris.
// ---------------------------------------------------------------------------

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export const looksLikeHtml = (s: string) => /<[a-z][\s\S]*>/i.test(s || "");

/** Teks lama (markdown ringan) -> HTML, sama dengan mdToHtmlForEditor di App.tsx. */
export function contentToHtml(value: string): string {
  if (!value) return "";
  if (looksLikeHtml(value)) return value;
  const h = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return value.split(/\n\s*\n/).map((para) => `<p>${h(para).replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>").replace(/\n/g, "<br>")}</p>`).join("");
}

interface Fmt { b?: boolean; i?: boolean; u?: boolean; s?: boolean; color?: string; sz?: number; font?: string }

export interface NumberingUse { numId: number; ordered: boolean; start: number }

export interface Ctx {
  /** Pembuat content control variabel; mengembalikan XML sebaris. */
  variable: (key: string, rPr: string) => string;
  /** Daftar instance penomoran yang dibutuhkan (diteruskan ke numbering.xml). */
  numbering: NumberingUse[];
  skipped: string[];
  defaultJc: "both" | "left";
}

const NAMED_COLORS: Record<string, string> = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
  orange: "FFA500", purple: "800080", gray: "808080", grey: "808080", navy: "000080", maroon: "800000",
};
const FONT_SIZE_PT = [8, 10, 12, 14, 18, 24, 36]; // <font size="1..7">

function parseColor(v: string | undefined): string | undefined {
  if (!v) return undefined;
  const s = v.trim().toLowerCase();
  let m = /^#([0-9a-f]{6})$/.exec(s); if (m) return m[1].toUpperCase();
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s); if (m) return (m[1] + m[1] + m[2] + m[2] + m[3] + m[3]).toUpperCase();
  m = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/.exec(s);
  if (m) return [m[1], m[2], m[3]].map((n) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, "0")).join("").toUpperCase();
  return NAMED_COLORS[s];
}

function parseStyle(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of (style || "").split(";")) {
    const i = part.indexOf(":");
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

function sizeHalfPts(v: string | undefined): number | undefined {
  if (!v) return undefined;
  const m = /^([\d.]+)\s*(px|pt)?$/i.exec(v.trim());
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!n) return undefined;
  const pt = (m[2] || "px").toLowerCase() === "px" ? n * 0.75 : n;
  return Math.max(2, Math.min(400, Math.round(pt * 2)));
}

function rPrXml(f: Fmt): string {
  // urutan elemen mengikuti skema OOXML: rFonts, b, i, strike, color, sz, szCs, u
  const parts: string[] = [];
  if (f.font) parts.push(`<w:rFonts w:ascii="${esc(f.font)}" w:hAnsi="${esc(f.font)}" w:cs="${esc(f.font)}"/>`);
  if (f.b) parts.push("<w:b/>");
  if (f.i) parts.push("<w:i/>");
  if (f.s) parts.push("<w:strike/>");
  if (f.color) parts.push(`<w:color w:val="${f.color}"/>`);
  if (f.sz) parts.push(`<w:sz w:val="${f.sz}"/><w:szCs w:val="${f.sz}"/>`);
  if (f.u) parts.push('<w:u w:val="single"/>');
  return parts.length ? `<w:rPr>${parts.join("")}</w:rPr>` : "";
}

interface PNode { nodeName: string; tagName?: string; value?: string; attrs?: { name: string; value: string }[]; childNodes?: PNode[] }
const attr = (n: PNode, name: string) => n.attrs?.find((a) => a.name === name)?.value;
const kids = (n: PNode) => n.childNodes || [];

function fmtFor(n: PNode, f: Fmt): Fmt {
  const t = n.nodeName;
  const out: Fmt = { ...f };
  if (t === "b" || t === "strong") out.b = true;
  if (t === "i" || t === "em") out.i = true;
  if (t === "u") out.u = true;
  if (t === "s" || t === "strike" || t === "del") out.s = true;
  if (t === "font") {
    const c = parseColor(attr(n, "color")); if (c) out.color = c;
    const sz = Number(attr(n, "size")); if (sz >= 1 && sz <= 7) out.sz = FONT_SIZE_PT[sz - 1] * 2;
  }
  const st = parseStyle(attr(n, "style"));
  if (st["font-weight"] && (/^(bold|[6-9]00)$/.test(st["font-weight"]))) out.b = true;
  if (st["font-style"] === "italic") out.i = true;
  if ((st["text-decoration"] || st["text-decoration-line"] || "").includes("underline")) out.u = true;
  if ((st["text-decoration"] || st["text-decoration-line"] || "").includes("line-through")) out.s = true;
  const c = parseColor(st["color"]); if (c) out.color = c;
  const sz = sizeHalfPts(st["font-size"]); if (sz) out.sz = sz;
  if (st["font-family"]) out.font = st["font-family"].split(",")[0].replace(/["']/g, "").trim() || out.font;
  return out;
}

function jcOf(n: PNode, dflt: string): string {
  const st = parseStyle(attr(n, "style"));
  const a = (st["text-align"] || attr(n, "align") || "").toLowerCase();
  return a === "center" ? "center" : a === "right" ? "right" : a === "justify" ? "both" : a === "left" ? "left" : dflt;
}

const INLINE_TAGS = new Set(["b", "strong", "i", "em", "u", "s", "strike", "del", "span", "font", "a", "br", "sub", "sup", "small", "mark", "code", "#text"]);

interface ParaOpts { jc: string; numId?: number; ilvl?: number; indentTwips?: number; heading?: boolean; keepNext?: boolean; border?: boolean }

function paraXml(runs: string, o: ParaOpts): string {
  const pPr: string[] = [];
  if (o.keepNext) pPr.push("<w:keepNext/>");
  if (o.numId) pPr.push(`<w:numPr><w:ilvl w:val="${o.ilvl || 0}"/><w:numId w:val="${o.numId}"/></w:numPr>`);
  if (o.border) pPr.push('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="808080"/></w:pBdr>');
  if (o.indentTwips) pPr.push(`<w:ind w:left="${o.indentTwips}"/>`);
  // Baris yang diakhiri <br> pada paragraf rata kiri-kanan akan melar di Word; pakai rata kiri utk paragraf ber-<br>.
  pPr.push(`<w:jc w:val="${o.jc === "both" && runs.includes("<w:br/>") ? "left" : o.jc}"/>`);
  return `<w:p><w:pPr>${pPr.join("")}</w:pPr>${runs}</w:p>`;
}

function textRuns(text: string, f: Fmt, ctx: Ctx): string {
  if (!text) return "";
  const rPr = rPrXml(f);
  const out: string[] = [];
  const re = /\{\{([^}]+)\}\}/g;
  let last = 0, m: RegExpExecArray | null;
  const plain = (s: string) => (s ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(s)}</w:t></w:r>` : "");
  while ((m = re.exec(text))) {
    out.push(plain(text.slice(last, m.index)));
    out.push(ctx.variable(m[1].trim(), rPr));
    last = m.index + m[0].length;
  }
  out.push(plain(text.slice(last)));
  return out.join("");
}

/** Run sebaris dari daftar node inline (teks, b, i, span, br, ...). */
function inlineRuns(nodes: PNode[], f: Fmt, ctx: Ctx, state: { lastSpace: boolean }): string {
  const out: string[] = [];
  for (const n of nodes) {
    if (n.nodeName === "#text") {
      let t = (n.value || "").replace(/\s+/g, " ");
      if (state.lastSpace) t = t.replace(/^ /, "");
      if (!t) continue;
      state.lastSpace = t.endsWith(" ");
      out.push(textRuns(t, f, ctx));
    } else if (n.nodeName === "br") {
      out.push("<w:r><w:br/></w:r>");
      state.lastSpace = true;
    } else if (n.nodeName === "img") {
      ctx.skipped.push("img");
    } else if (n.nodeName === "#comment") {
      // abaikan
    } else {
      out.push(inlineRuns(kids(n), fmtFor(n, f), ctx, state));
    }
  }
  return out.join("");
}

function trimEnd(runs: string): string {
  return runs.replace(/(<w:t xml:space="preserve">)([^<]*?) (<\/w:t><\/w:r>)$/, "$1$2$3");
}

/** Ubah daftar node blok menjadi XML (paragraf + tabel). */
function blocks(nodes: PNode[], f: Fmt, ctx: Ctx, o: { jc: string; listDepth: number; numId?: number; indent?: number }): string {
  const out: string[] = [];
  let buf: PNode[] = [];
  const flush = () => {
    if (!buf.length) return;
    const runs = trimEnd(inlineRuns(buf, f, ctx, { lastSpace: true }));
    buf = [];
    if (runs.trim()) out.push(paraXml(runs, { jc: o.jc, indentTwips: o.indent }));
  };
  for (const n of nodes) {
    const t = n.nodeName;
    if (INLINE_TAGS.has(t) || t === "img" || t === "#comment") { buf.push(n); continue; }
    flush();
    if (t === "p" || t === "div" || t === "h3" || t === "h4" || t === "h1" || t === "h2" || t === "blockquote") {
      const heading = t[0] === "h";
      const ff = fmtFor(n, heading ? { ...f, b: true, sz: f.sz || (t === "h3" ? 28 : 26) } : f);
      const hasBlockChild = kids(n).some((k) => !INLINE_TAGS.has(k.nodeName) && k.nodeName !== "#comment" && k.nodeName !== "img");
      const jc = jcOf(n, o.jc);
      const indent = t === "blockquote" ? (o.indent || 0) + 720 : o.indent;
      if (hasBlockChild) out.push(blocks(kids(n), ff, ctx, { ...o, jc, indent }));
      else {
        const runs = trimEnd(inlineRuns(kids(n), ff, ctx, { lastSpace: true }));
        out.push(paraXml(runs, { jc, indentTwips: indent, keepNext: heading }));
      }
    } else if (t === "ul" || t === "ol") {
      out.push(listXml(n, f, ctx, o));
    } else if (t === "table") {
      out.push(tableXml(n, f, ctx, o.jc));
    } else if (t === "hr") {
      out.push(paraXml("", { jc: o.jc, border: true }));
    } else {
      // tag tak dikenal: perlakukan isinya sebagai blok
      out.push(blocks(kids(n), fmtFor(n, f), ctx, o));
    }
  }
  flush();
  return out.join("");
}

function listXml(n: PNode, f: Fmt, ctx: Ctx, o: { jc: string; listDepth: number; indent?: number }): string {
  const ordered = n.nodeName === "ol";
  const startAttr = Number(attr(n, "start"));
  const numId = ctx.numbering.length + 1;
  ctx.numbering.push({ numId, ordered, start: startAttr > 0 ? startAttr : 1 });
  const ilvl = Math.min(o.listDepth, 8);
  const out: string[] = [];
  for (const li of kids(n)) {
    if (li.nodeName !== "li") continue;
    const ff = fmtFor(li, f);
    // isi li: bagian inline sebelum sublist -> paragraf bernomor; sublist/blok -> setelahnya
    const inlineNodes: PNode[] = [], rest: PNode[] = [];
    let seenBlock = false;
    for (const c of kids(li)) {
      const isInline = INLINE_TAGS.has(c.nodeName) || c.nodeName === "img" || c.nodeName === "#comment";
      if (!seenBlock && isInline) inlineNodes.push(c);
      else if (!seenBlock && (c.nodeName === "p" || c.nodeName === "div") && !kids(c).some((k) => !INLINE_TAGS.has(k.nodeName))) {
        // <li><p>teks</p>...</li>: p pertama jadi paragraf bernomor
        inlineNodes.push(...kids(c)); seenBlock = true;
      } else { seenBlock = true; rest.push(c); }
    }
    const runs = trimEnd(inlineRuns(inlineNodes, ff, ctx, { lastSpace: true }));
    out.push(paraXml(runs, { jc: o.jc, numId, ilvl }));
    if (rest.length) out.push(blocks(rest, ff, ctx, { jc: o.jc, listDepth: o.listDepth + 1, indent: o.indent }));
  }
  return out.join("");
}

function tableXml(n: PNode, f: Fmt, ctx: Ctx, jc: string): string {
  // kumpulkan baris (thead/tbody/tfoot boleh ada atau tidak)
  const rows: PNode[] = [];
  const collect = (x: PNode) => { for (const c of kids(x)) { if (c.nodeName === "tr") rows.push(c); else if (["thead", "tbody", "tfoot"].includes(c.nodeName)) collect(c); } };
  collect(n);
  if (!rows.length) return "";
  // grid kolom + rowspan
  const occupied: Record<string, boolean> = {};
  interface Cell { node: PNode; r: number; c: number; colspan: number; rowspan: number; th: boolean }
  const cells: Cell[] = [];
  let maxCols = 0;
  rows.forEach((tr, r) => {
    let c = 0;
    for (const td of kids(tr)) {
      if (td.nodeName !== "td" && td.nodeName !== "th") continue;
      while (occupied[`${r},${c}`]) c++;
      const colspan = Math.max(1, Math.min(50, Number(attr(td, "colspan")) || 1));
      const rowspan = Math.max(1, Math.min(rows.length - r, Number(attr(td, "rowspan")) || 1));
      for (let rr = r; rr < r + rowspan; rr++) for (let cc = c; cc < c + colspan; cc++) occupied[`${rr},${cc}`] = true;
      cells.push({ node: td, r, c, colspan, rowspan, th: td.nodeName === "th" });
      c += colspan;
    }
    maxCols = Math.max(maxCols, c);
  });
  maxCols = Math.max(1, maxCols);
  const colW = Math.floor(9000 / maxCols);
  const border = (side: string) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="64748B"/>`;
  const tblPr = `<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(border).join("")}</w:tblBorders><w:tblLayout w:type="autofit"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>`;
  const grid = `<w:tblGrid>${Array.from({ length: maxCols }, () => `<w:gridCol w:w="${colW}"/>`).join("")}</w:tblGrid>`;
  const trs: string[] = [];
  rows.forEach((_, r) => {
    const tcs: string[] = [];
    let col = 0;
    const rowCells = cells.filter((x) => x.r === r).sort((a, b) => a.c - b.c);
    // sel yang "tertutup" rowspan dari baris atas: tulis sel lanjutan vMerge
    const continued = Object.keys(occupied).filter((k) => k.startsWith(r + ",")).map((k) => Number(k.split(",")[1])).filter((cc) => !rowCells.some((x) => cc >= x.c && cc < x.c + x.colspan)).sort((a, b) => a - b);
    const all: { c: number; xml: string }[] = [];
    for (const cell of rowCells) {
      const ff = fmtFor(cell.node, cell.th ? { ...f, b: true } : f); // style/warna pada <td> ikut berlaku
      const cj = cell.th ? jcOf(cell.node, "center") : jcOf(cell.node, "left");
      let inner = blocks(kids(cell.node), ff, ctx, { jc: cj, listDepth: 0 });
      if (!inner) inner = paraXml("", { jc: cj });
      else if (!/<\/w:p>$/.test(inner)) inner += paraXml("", { jc: cj }); // sel wajib diakhiri paragraf
      const tcPr = `<w:tcPr><w:tcW w:w="${colW * cell.colspan}" w:type="dxa"/>${cell.colspan > 1 ? `<w:gridSpan w:val="${cell.colspan}"/>` : ""}${cell.rowspan > 1 ? '<w:vMerge w:val="restart"/>' : ""}${cell.th ? '<w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/>' : ""}</w:tcPr>`;
      all.push({ c: cell.c, xml: `<w:tc>${tcPr}${inner}</w:tc>` });
    }
    for (const cc of continued) all.push({ c: cc, xml: `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>` });
    all.sort((a, b) => a.c - b.c);
    void col;
    trs.push(`<w:tr>${all.map((x) => x.xml).join("")}</w:tr>`);
    void tcs;
  });
  // paragraf kosong setelah tabel supaya blok berikutnya tidak menempel
  return `<w:tbl>${tblPr}${grid}${trs.join("")}</w:tbl>${paraXml("", { jc })}`;
}

/** Konversi isi pasal (HTML atau teks lama) -> XML blok Word. */
export function clauseContentToWordXml(content: string, ctx: Ctx): string {
  const frag = parseFragment(contentToHtml(content)) as unknown as PNode;
  return blocks(kids(frag), {}, ctx, { jc: ctx.defaultJc, listDepth: 0 });
}

/** numbering.xml untuk semua daftar yang dipakai (bullet & angka, 9 tingkat). */
export function numberingXml(uses: NumberingUse[]): string {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const BUL = ["•", "o", "▪"];
  const lvl = (i: number, ordered: boolean, start: number) => {
    const fmt = ordered ? (["decimal", "lowerLetter", "lowerRoman"][i % 3]) : "bullet";
    const text = ordered ? `%${i + 1}.` : BUL[i % 3];
    const font = !ordered ? '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>' : "";
    const textSafe = !ordered && i % 3 === 0 ? "" : text;
    const rpr = !ordered && i % 3 === 0 ? font : "";
    return `<w:lvl w:ilvl="${i}"><w:start w:val="${i === 0 ? start : 1}"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${textSafe}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>${rpr}</w:lvl>`;
  };
  const abs = uses.map((u, k) => `<w:abstractNum w:abstractNumId="${k}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => lvl(i, u.ordered, u.start)).join("")}</w:abstractNum>`).join("");
  const nums = uses.map((u, k) => `<w:num w:numId="${u.numId}"><w:abstractNumId w:val="${k}"/></w:num>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}">${abs}${nums}</w:numbering>`;
}

// ---------------------------------------------------------------------------
// Bentuk ringkas (JSON) untuk plugin editor Word: paragraf, daftar, dan tabel.
// Plugin menyusunnya kembali lewat API editor. Variabel {{Key}} menjadi run { v: "Key" }.
// ---------------------------------------------------------------------------
export interface PlugRun { x: string; b?: boolean; i?: boolean; u?: boolean; v?: string }
export type PlugBlock =
  | { t: "p"; runs: PlugRun[]; jc?: string }
  | { t: "li"; runs: PlugRun[]; ordered: boolean; level: number; n: number }
  | { t: "table"; rows: PlugBlock[][][] }; // baris -> sel -> paragraf

function plugRuns(nodes: PNode[], f: Fmt, out: PlugRun[]): void {
  for (const n of nodes) {
    if (n.nodeName === "#text") {
      const t = (n.value || "").replace(/\s+/g, " ");
      if (!t) continue;
      const re = /\{\{([^}]+)\}\}/g;
      let last = 0, m: RegExpExecArray | null;
      while ((m = re.exec(t))) {
        if (m.index > last) out.push({ x: t.slice(last, m.index), b: f.b, i: f.i, u: f.u });
        out.push({ x: "", v: m[1].trim(), b: f.b, i: f.i, u: f.u });
        last = m.index + m[0].length;
      }
      if (last < t.length) out.push({ x: t.slice(last), b: f.b, i: f.i, u: f.u });
    } else if (n.nodeName === "br") {
      out.push({ x: "\n" });
    } else if (n.nodeName !== "#comment" && n.nodeName !== "img") {
      plugRuns(kids(n), fmtFor(n, f), out);
    }
  }
}
const trimRuns = (r: PlugRun[]): PlugRun[] => {
  const o = r.map((x) => ({ ...x }));
  while (o.length && !o[0].v && !o[0].x.trim()) o.shift();
  while (o.length && !o[o.length - 1].v && !o[o.length - 1].x.trim()) o.pop();
  if (o.length && !o[0].v) o[0].x = o[0].x.replace(/^\s+/, "");
  if (o.length && !o[o.length - 1].v) o[o.length - 1].x = o[o.length - 1].x.replace(/\s+$/, "");
  return o;
};

function plugBlocks(nodes: PNode[], f: Fmt, level: number, out: PlugBlock[]): void {
  let buf: PNode[] = [];
  const flush = () => {
    if (!buf.length) return;
    const runs: PlugRun[] = [];
    plugRuns(buf, f, runs);
    buf = [];
    const t = trimRuns(runs);
    if (t.length) out.push({ t: "p", runs: t });
  };
  for (const n of nodes) {
    const t = n.nodeName;
    if (INLINE_TAGS.has(t) || t === "img" || t === "#comment") { buf.push(n); continue; }
    flush();
    if (["p", "div", "h1", "h2", "h3", "h4", "blockquote"].includes(t)) {
      const ff = fmtFor(n, t[0] === "h" ? { ...f, b: true } : f);
      const hasBlock = kids(n).some((k) => !INLINE_TAGS.has(k.nodeName) && k.nodeName !== "#comment" && k.nodeName !== "img");
      if (hasBlock) plugBlocks(kids(n), ff, level, out);
      else {
        const runs: PlugRun[] = [];
        plugRuns(kids(n), ff, runs);
        const tr = trimRuns(runs);
        out.push({ t: "p", runs: tr.length ? tr : [{ x: "" }], jc: jcOf(n, "") || undefined });
      }
    } else if (t === "ul" || t === "ol") {
      let i = 0;
      for (const li of kids(n)) {
        if (li.nodeName !== "li") continue;
        i++;
        const inline: PNode[] = [], nested: PNode[] = [];
        for (const k of kids(li)) (k.nodeName === "ul" || k.nodeName === "ol" ? nested : inline).push(k);
        const runs: PlugRun[] = [];
        plugRuns(inline, f, runs);
        out.push({ t: "li", runs: trimRuns(runs), ordered: t === "ol", level, n: i });
        if (nested.length) plugBlocks(nested, f, level + 1, out);
      }
    } else if (t === "table") {
      const rows: PNode[] = [];
      const collect = (x: PNode) => { for (const c of kids(x)) { if (c.nodeName === "tr") rows.push(c); else if (["thead", "tbody", "tfoot"].includes(c.nodeName)) collect(c); } };
      collect(n);
      const tbl: PlugBlock[][][] = rows.map((tr) =>
        kids(tr).filter((c) => c.nodeName === "td" || c.nodeName === "th").map((td) => {
          const cell: PlugBlock[] = [];
          plugBlocks(kids(td), fmtFor(td, td.nodeName === "th" ? { ...f, b: true } : f), 0, cell);
          return cell.length ? cell : [{ t: "p", runs: [{ x: "" }] } as PlugBlock];
        }));
      if (tbl.length && tbl.some((r) => r.length)) out.push({ t: "table", rows: tbl });
    } else {
      plugBlocks(kids(n), fmtFor(n, f), level, out);
    }
  }
  flush();
}

/** Isi pasal -> blok ringkas untuk plugin. */
export function clauseToPluginBlocks(content: string): PlugBlock[] {
  const frag = parseFragment(contentToHtml(content)) as unknown as PNode;
  const out: PlugBlock[] = [];
  plugBlocks(kids(frag), {}, 0, out);
  return out;
}
