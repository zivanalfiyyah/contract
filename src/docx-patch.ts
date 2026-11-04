// ---------------------------------------------------------------------------
// DOCX "edit in place" untuk Document Workspace (dokumen hasil Upload).
//
// Kenapa bukan konversi DOCX -> HTML -> DOCX?
//   Konversi bolak-balik (mammoth, html-docx, dst) selalu kehilangan kop/
//   header, logo, footer, section/page break, style tabel, dan tata letak
//   area tanda tangan — persis masalah "dokumen jadi kumpulan paragraf" yang
//   mau dihindari. Di sini berkas .docx ASLI tetap jadi sumber kebenaran:
//
//   1. Tampilan: docx-preview merender word/document.xml apa adanya (header,
//      footer, tabel, gambar, page break).
//   2. Edit: setiap <p> hasil render dipasangkan (by text) ke <w:p> di XML.
//      Hanya paragraf yang berhasil dipasangkan yang bisa diedit.
//   3. Simpan: untuk tiap paragraf yang berubah, dihitung diff karakter antara
//      teks lama & baru, lalu diff itu diterapkan langsung ke node <w:t> di
//      XML. Run (w:r) & properti paragraf (w:pPr) tidak disentuh, jadi bold/
//      italic/warna/ukuran font/alignment/numbering tetap; teks yang disisip
//      mengikuti format karakter di posisi sisip (perilaku sama seperti Word).
//   4. Seluruh part lain di dalam zip (header, footer, media, styles,
//      numbering, settings) disalin byte-per-byte ke berkas versi baru.
//
// Batasan tahap awal (sengaja): tidak bisa menambah paragraf/baris tabel
// baru dan header/footer tidak bisa diedit dari workspace (diproteksi supaya
// kop/logo tidak rusak). Untuk perubahan struktur, gunakan "Unggah Revisi".
// ---------------------------------------------------------------------------
import JSZip from "jszip";
import { diffChars } from "diff";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

export interface DocxModel {
  zip: JSZip;
  xml: XMLDocument;
  /** <w:p> di body dokumen, urut dokumen (termasuk yang di dalam tabel). */
  paragraphs: Element[];
  /**
   * word/numbering.xml (definisi list). null = berkas tidak ada. Dibaca supaya
   * list BARU (hasil paste) bisa ditambahkan tanpa merusak list yang sudah ada.
   */
  numbering?: XMLDocument | null;
  /** numbering.xml ada tapi tidak bisa dibaca -> list baru TIDAK dibuat (jangan menimpanya). */
  numberingBroken?: boolean;
  /** numbering.xml dibuat dari nol (perlu didaftarkan di [Content_Types] & rels saat simpan). */
  numberingIsNew?: boolean;
  /** numbering.xml diubah -> harus ditulis ulang saat simpan. */
  numberingDirty?: boolean;
  /** Cache definisi list yang dibuat dalam sesi ini. */
  numState?: { bullet?: number; decimalAbs?: number };
}

type Seg =
  | { kind: "t"; el: Element; start: number; len: number }
  | { kind: "tab" | "br"; el: Element; start: number; len: 1 };

function isW(el: Element, local: string) {
  return el.namespaceURI === W_NS && el.localName === local;
}

/** true kalau el berada di dalam textbox/drawing (paragraf "melayang"). */
function insideTextbox(el: Element, stopAt: Element | null): boolean {
  let cur: Element | null = el.parentElement;
  while (cur && cur !== stopAt) {
    if (cur.namespaceURI === W_NS && cur.localName === "txbxContent") return true;
    cur = cur.parentElement;
  }
  return false;
}

/** Segmen teks sebuah <w:p> berurutan: w:t, w:tab ("\t"), w:br/w:cr ("\n"). */
function segmentsOf(p: Element): Seg[] {
  const segs: Seg[] = [];
  let pos = 0;
  const walk = (node: Element) => {
    for (const child of Array.from(node.children)) {
      if (child.namespaceURI === W_NS) {
        const ln = child.localName;
        // Teks yang terhapus (tracked change), instruksi field, & isi
        // textbox bukan bagian teks paragraf yang tampil — lewati.
        if (ln === "del" || ln === "instrText" || ln === "delText" || ln === "txbxContent" || ln === "pPr" || ln === "rPr") continue;
        if (ln === "t") {
          const len = (child.textContent || "").length;
          segs.push({ kind: "t", el: child, start: pos, len });
          pos += len;
          continue;
        }
        if (ln === "tab" && child.parentElement && isW(child.parentElement, "r")) {
          segs.push({ kind: "tab", el: child, start: pos, len: 1 });
          pos += 1;
          continue;
        }
        if ((ln === "br" || ln === "cr") && child.getAttributeNS(W_NS, "type") !== "page") {
          segs.push({ kind: "br", el: child, start: pos, len: 1 });
          pos += 1;
          continue;
        }
      }
      walk(child);
    }
  };
  walk(p);
  return segs;
}

export function paragraphText(p: Element): string {
  return segmentsOf(p)
    .map((s) => (s.kind === "t" ? s.el.textContent || "" : s.kind === "tab" ? "\t" : "\n"))
    .join("");
}

export async function loadDocx(buf: ArrayBuffer): Promise<DocxModel> {
  const zip = await JSZip.loadAsync(buf);
  const entry = zip.file("word/document.xml");
  if (!entry) throw new Error("Berkas .docx tidak valid (word/document.xml tidak ditemukan).");
  const xml = new DOMParser().parseFromString(await entry.async("string"), "application/xml");
  if (xml.getElementsByTagName("parsererror").length) throw new Error("word/document.xml tidak bisa dibaca.");
  const body = xml.getElementsByTagNameNS(W_NS, "body")[0];
  if (!body) throw new Error("Berkas .docx tidak punya body dokumen.");
  const paragraphs = Array.from(body.getElementsByTagNameNS(W_NS, "p")).filter((p) => !insideTextbox(p, body));
  let numbering: XMLDocument | null = null;
  let numberingBroken = false;
  const numEntry = zip.file("word/numbering.xml");
  if (numEntry) {
    const doc = new DOMParser().parseFromString(await numEntry.async("string"), "application/xml");
    if (doc.getElementsByTagName("parsererror").length || !doc.documentElement) numberingBroken = true;
    else numbering = doc;
  }
  return { zip, xml, paragraphs, numbering, numberingBroken };
}

// ---- Normalisasi & ekstraksi teks dari DOM hasil render --------------------

/** Samakan karakter yang dirender beda tapi maknanya sama (1:1, panjang tetap). */
export function normChars(s: string): string {
  return s.replace(/ /g, " ").replace(/​/g, "");
}
function matchKey(s: string): string {
  return normChars(s).replace(/\s+/g, " ").trim();
}

// Atribut karakter yang bisa diubah dari toolbar workspace (bitmask).
export const ATTR_B = 1, ATTR_I = 2, ATTR_U = 4;

function styleAttrs(el: Element | null, host: Element): number {
  if (!el) return 0;
  const cs = getComputedStyle(el);
  let a = 0;
  const w = parseInt(cs.fontWeight, 10);
  if ((isNaN(w) ? cs.fontWeight === "bold" : w >= 600)) a |= ATTR_B;
  if (cs.fontStyle === "italic" || cs.fontStyle === "oblique") a |= ATTR_I;
  // text-decoration tidak diwariskan di computed style -> telusuri leluhur.
  for (let cur: Element | null = el; cur; cur = cur.parentElement) {
    if (getComputedStyle(cur).textDecorationLine.includes("underline")) { a |= ATTR_U; break; }
    if (cur === host) break;
  }
  return a;
}

/** Nama keluarga font pertama dari computed style (tanpa kutip/fallback list). */
function firstFontFamily(cs: CSSStyleDeclaration): string {
  return (cs.fontFamily || "").split(",")[0].trim().replace(/^["']|["']$/g, "") || "Calibri";
}

/** px -> half-point OOXML (1px = 0.75pt = 1.5 half-point), dibulatkan. */
function pxToHalfPt(px: string): number {
  const n = parseFloat(px);
  return Math.max(2, Math.round((isNaN(n) ? 16 : n) * 1.5));
}

/** Font-family & ukuran (half-point) aktif pada sebuah node, dari computed style. */
function styleFontInfo(el: Element | null, host: Element): { font: string; sizeHp: number } {
  const cs = getComputedStyle((el as HTMLElement) || (host as HTMLElement));
  return { font: firstFontFamily(cs), sizeHp: pxToHalfPt(cs.fontSize) };
}

/** Warna teks aktif pada sebuah node sebagai hex 6 digit huruf besar tanpa "#" ("" kalau tak terbaca). */
function styleColor(el: Element | null, host: Element): string {
  const c = getComputedStyle((el as HTMLElement) || (host as HTMLElement)).color || "";
  const m = /^rgba?\(\s*(\d{1,3})\s*[, ]\s*(\d{1,3})\s*[, ]\s*(\d{1,3})/.exec(c.trim());
  if (!m) return "";
  return [m[1], m[2], m[3]].map((n) => Math.min(255, parseInt(n, 10)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/**
 * Gaya karakter yang diwarisi sebuah <p> POLOS (tanpa class gaya dokumen) di tempat
 * yang sama dgn `el`. Paragraf baru hasil paste tidak membawa class gaya docx-preview,
 * jadi teks tanpa format eksplisit terbaca bergaya "bawaan aplikasi" — bukan gaya
 * dokumen. Nilai ini dipakai utk mengenalinya supaya TIDAK ditulis ke .docx.
 */
export function defaultCharStyle(el: HTMLElement): { font: string; size: number; color: string } | undefined {
  const parent = el.parentElement;
  if (!parent) return undefined;
  const probe = document.createElement("p");
  probe.textContent = "x";
  parent.appendChild(probe);
  try {
    const rc = renderedChars(probe);
    return { font: rc.fonts[0], size: rc.sizes[0], color: rc.colors[0] };
  } finally {
    parent.removeChild(probe);
  }
}

/** Rasio line-height thd font-size ("normal"/tanpa unit eksplisit -> dianggap 1). */
export function renderedLineSpacing(el: HTMLElement): number {
  const cs = getComputedStyle(el);
  const lh = cs.lineHeight;
  if (!lh || lh === "normal") return 1;
  const px = parseFloat(lh);
  const fs = parseFloat(cs.fontSize) || 16;
  if (isNaN(px) || !fs) return 1;
  const ratio = px / fs;
  return ratio > 0.5 && ratio < 4 ? ratio : 1;
}

/**
 * Teks + atribut per karakter sebuah <p> hasil render.
 * <br> -> "\n", tab-stop docx-preview -> "\t", nbsp -> spasi.
 */
export function renderedChars(el: Element): { text: string; attrs: number[]; fonts: string[]; sizes: number[]; colors: string[] } {
  let text = "";
  const attrs: number[] = [];
  const fonts: string[] = [];
  const sizes: number[] = [];
  const colors: string[] = [];
  const push = (s: string, a: number, fi: { font: string; sizeHp: number }, color: string) => {
    for (const ch of s) {
      if (ch === "\u200b" || ch === "\ufeff") continue;
      text += ch === "\u00a0" ? " " : ch;
      attrs.push(a);
      fonts.push(fi.font);
      sizes.push(fi.sizeHp);
      colors.push(color);
    }
  };
  const walk = (n: Node) => {
    if (n.nodeType === Node.TEXT_NODE) { push((n as Text).data, styleAttrs(n.parentElement, el), styleFontInfo(n.parentElement, el), styleColor(n.parentElement, el)); return; }
    if (n.nodeType !== Node.ELEMENT_NODE) return;
    const e = n as HTMLElement;
    if (e.tagName === "BR") {
      const a = attrs.length ? attrs[attrs.length - 1] : styleAttrs(e.parentElement, el);
      const fi = fonts.length ? { font: fonts[fonts.length - 1], sizeHp: sizes[sizes.length - 1] } : styleFontInfo(e.parentElement, el);
      const col = colors.length ? colors[colors.length - 1] : styleColor(e.parentElement, el);
      push("\n", a, fi, col);
      return;
    }
    if (e.classList.contains("docx-tab-stop")) { push("\t", styleAttrs(e, el), styleFontInfo(e, el), styleColor(e, el)); return; }
    e.childNodes.forEach(walk);
  };
  el.childNodes.forEach(walk);
  return { text, attrs, fonts, sizes, colors };
}

/** Teks sebuah <p> hasil render (lihat renderedChars). */
export function renderedText(el: Element): string {
  return renderedChars(el).text;
}

export function renderedAlign(el: HTMLElement): string {
  const a = getComputedStyle(el).textAlign;
  return a === "start" ? "left" : a === "end" ? "right" : a;
}

export interface ParagraphBinding {
  el: HTMLElement; index: number; original: string; attrs: number[]; align: string;
  fonts: string[]; sizes: number[]; lineSpacing: number; colors: string[];
}

/**
 * Pasangkan <p> hasil render (di luar header/footer) dengan <w:p> di XML,
 * greedy berdasarkan teks. Paragraf yang tidak yakin cocok TIDAK diikat
 * (tetap tampil, tapi tidak bisa diedit) — lebih aman daripada salah tulis.
 */
export function bindRenderedParagraphs(container: HTMLElement, model: DocxModel): ParagraphBinding[] {
  const rendered = Array.from(container.querySelectorAll("p")).filter(
    (p) => !p.closest("header") && !p.closest("footer"),
  ) as HTMLElement[];
  const xmlKeys = model.paragraphs.map((p) => matchKey(paragraphText(p)));
  const bindings: ParagraphBinding[] = [];
  let j = 0;
  const LOOKAHEAD = 40;
  for (const el of rendered) {
    const rc = renderedChars(el);
    const key = matchKey(rc.text);
    let found = -1;
    for (let k = j; k < Math.min(xmlKeys.length, j + LOOKAHEAD); k++) {
      if (xmlKeys[k] === key) { found = k; break; }
    }
    if (found === -1) continue;
    bindings.push({
      el, index: found, original: rc.text, attrs: rc.attrs, align: renderedAlign(el),
      fonts: rc.fonts, sizes: rc.sizes, lineSpacing: renderedLineSpacing(el), colors: rc.colors,
    });
    j = found + 1;
  }
  return bindings;
}

// ---- Terapkan perubahan teks ke XML ----------------------------------------

function setText(t: Element, text: string) {
  t.textContent = text;
  t.setAttributeNS(XML_NS, "xml:space", "preserve");
}

/** Buat node run-child untuk teks yang boleh mengandung \n / \t. */
function nodesForText(xml: XMLDocument, text: string): Element[] {
  const out: Element[] = [];
  let buf = "";
  const flush = () => {
    if (!buf) return;
    const t = xml.createElementNS(W_NS, "w:t");
    setText(t, buf);
    out.push(t);
    buf = "";
  };
  for (const ch of text) {
    if (ch === "\n") { flush(); out.push(xml.createElementNS(W_NS, "w:br")); }
    else if (ch === "\t") { flush(); out.push(xml.createElementNS(W_NS, "w:tab")); }
    else buf += ch;
  }
  flush();
  return out;
}

/** Hapus teks [from, to) dari paragraf. */
function deleteRange(p: Element, from: number, to: number) {
  if (to <= from) return;
  for (const s of segmentsOf(p)) {
    const a = Math.max(from, s.start);
    const b = Math.min(to, s.start + s.len);
    if (b <= a) continue;
    if (s.kind === "t") {
      const txt = s.el.textContent || "";
      setText(s.el, txt.slice(0, a - s.start) + txt.slice(b - s.start));
    } else {
      s.el.parentNode?.removeChild(s.el);
    }
  }
}

/** Sisipkan teks pada posisi pos, mewarisi format run di posisi itu. */
function insertAt(xml: XMLDocument, p: Element, pos: number, text: string) {
  if (!text) return;
  const segs = segmentsOf(p);
  // Prioritas: segmen w:t yang memuat pos (atau berakhir tepat di pos —
  // format karakter sebelumnya, seperti Word), lalu segmen w:t mana pun
  // sesudahnya, lalu run apa pun.
  let target: Seg | undefined =
    segs.find((s) => s.kind === "t" && pos > s.start && pos <= s.start + s.len) ||
    segs.find((s) => s.kind === "t" && pos === s.start);
  if (!target) {
    // Posisi di batas tab/br atau paragraf tanpa w:t sama sekali.
    const before = [...segs].reverse().find((s) => s.start + s.len <= pos);
    const after = segs.find((s) => s.start >= pos);
    const anchor = before || after;
    if (anchor) {
      const nodes = nodesForText(xml, text);
      const ref = anchor.el;
      if (before) {
        let cursor: Node = ref;
        for (const n of nodes) { cursor.parentNode!.insertBefore(n, cursor.nextSibling); cursor = n; }
      } else {
        for (const n of nodes) ref.parentNode!.insertBefore(n, ref);
      }
      return;
    }
    // Paragraf kosong: buat run baru, salin format "mark run" paragraf bila ada.
    const r = xml.createElementNS(W_NS, "w:r");
    const pPr = Array.from(p.children).find((c) => isW(c, "pPr"));
    const markRPr = pPr ? Array.from(pPr.children).find((c) => isW(c, "rPr")) : undefined;
    if (markRPr) r.appendChild(markRPr.cloneNode(true));
    for (const n of nodesForText(xml, text)) r.appendChild(n);
    p.appendChild(r);
    return;
  }
  const t = target.el;
  const txt = t.textContent || "";
  const off = pos - target.start;
  if (!/[\n\t]/.test(text)) {
    setText(t, txt.slice(0, off) + text + txt.slice(off));
    return;
  }
  // Teks berisi baris baru/tab: pecah w:t jadi [sebelum][...sisipan...][sesudah]
  // di dalam run yang sama.
  setText(t, txt.slice(0, off));
  let cursor: Node = t;
  const tail = txt.slice(off);
  const nodes = nodesForText(xml, text);
  if (tail) {
    const tt = xml.createElementNS(W_NS, "w:t");
    setText(tt, tail);
    nodes.push(tt);
  }
  for (const n of nodes) { cursor.parentNode!.insertBefore(n, cursor.nextSibling); cursor = n; }
}

/**
 * Ganti rentang [start, end) paragraf dengan `text` sebagai SATU penggantian:
 * teks baru menempati run yang memuat karakter PERTAMA yang diganti, jadi
 * mewarisi formatnya (tebal/miring/font) — beda dengan diff per karakter yang
 * bisa memecah kata pengganti ke run tetangga. Dipakai "terima usulan markup".
 */
export function replaceTextRange(xml: XMLDocument, p: Element, start: number, end: number, text: string) {
  const segs = segmentsOf(p);
  const first = segs.find((s) => s.kind === "t" && start >= s.start && start < s.start + s.len);
  if (!first || end <= start || /[\n\t]/.test(text)) {
    deleteRange(p, start, end);
    insertAt(xml, p, start, text);
    return;
  }
  const firstEnd = Math.min(end, first.start + first.len);
  const txt = first.el.textContent || "";
  deleteRange(p, firstEnd, end); // bagian yang menjulur ke segmen lain
  setText(first.el, txt.slice(0, start - first.start) + text + txt.slice(firstEnd - first.start));
}

/**
 * Terapkan teks baru ke satu <w:p>. Perubahan dihitung dengan diff karakter
 * sehingga teks yang tidak berubah (dan run/format-nya) tetap persis sama.
 */
export function applyParagraphEdit(xml: XMLDocument, p: Element, oldText: string, newText: string) {
  if (oldText === newText) return;
  const current = normChars(paragraphText(p));
  // Offset diff dihitung terhadap teks yang dibaca dari XML. Kalau teks
  // render (oldText) sedikit berbeda dari XML (mis. spasi), pakai XML sbg dasar.
  const base = current === normChars(oldText) ? oldText : current;
  const ops: { pos: number; del: number; ins: string }[] = [];
  let pos = 0;
  for (const part of diffChars(normChars(base), newText)) {
    if (part.added) ops.push({ pos, del: 0, ins: part.value });
    else if (part.removed) { ops.push({ pos, del: part.value.length, ins: "" }); pos += part.value.length; }
    else pos += part.value.length;
  }
  // Gabungkan hapus+sisip yang bersebelahan (penggantian kata) supaya sisipan
  // mewarisi format teks yang diganti.
  const merged: typeof ops = [];
  for (const op of ops) {
    const last = merged[merged.length - 1];
    if (last && last.ins === "" && op.del === 0 && op.pos === last.pos + last.del) {
      last.ins = op.ins;
    } else merged.push({ ...op });
  }
  for (const op of merged.reverse()) {
    if (op.del) deleteRange(p, op.pos, op.pos + op.del);
    if (op.ins) {
      // Setelah deleteRange, sisipkan di awal rentang yang dihapus. Kalau
      // seluruh w:t di rentang itu habis, insertAt akan mencari run tetangga.
      insertAt(xml, p, op.pos, op.ins);
    }
  }
}

export interface DocxEdit { index: number; oldText: string; newText: string }

/** Hasilkan berkas .docx baru (Blob) dari model + daftar editan paragraf. */
export async function buildEditedDocx(model: DocxModel, edits: DocxEdit[]): Promise<Blob> {
  for (const e of edits) {
    const p = model.paragraphs[e.index];
    if (p) applyParagraphEdit(model.xml, p, e.oldText, e.newText);
  }
  const serialized = new XMLSerializer().serializeToString(model.xml);
  model.zip.file("word/document.xml", serialized.startsWith("<?xml") ? serialized : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${serialized}`);
  return model.zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}


// ---------------------------------------------------------------------------
// Editor lengkap: format karakter, perataan, paragraf baru/hapus, baris tabel
// ---------------------------------------------------------------------------

// Urutan anak w:rPr / w:pPr sesuai skema OOXML — Word menolak/merusak berkas
// bila urutannya salah, jadi elemen baru disisipkan di posisi yang benar.
const RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"];
const PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];

function childW(el: Element, local: string): Element | undefined {
  return Array.from(el.children).find((c) => isW(c, local));
}

function setOrderedChild(xml: XMLDocument, parent: Element, local: string, order: string[], val?: string): Element {
  let el = childW(parent, local);
  if (!el) {
    el = xml.createElementNS(W_NS, `w:${local}`);
    const idx = order.indexOf(local);
    const before = Array.from(parent.children).find((c) => c.namespaceURI === W_NS && order.indexOf(c.localName) > idx);
    parent.insertBefore(el, before || null);
  }
  if (val === undefined) el.removeAttributeNS(W_NS, "val");
  else el.setAttributeNS(W_NS, "w:val", val);
  return el;
}

function ensureRPr(xml: XMLDocument, run: Element): Element {
  let rPr = childW(run, "rPr");
  if (!rPr) {
    rPr = xml.createElementNS(W_NS, "w:rPr");
    run.insertBefore(rPr, run.firstChild);
  }
  return rPr;
}
function ensurePPr(xml: XMLDocument, p: Element): Element {
  let pPr = childW(p, "pPr");
  if (!pPr) {
    pPr = xml.createElementNS(W_NS, "w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  return pPr;
}

function runOf(el: Element): Element | null {
  const r = el.parentElement;
  return r && isW(r, "r") ? r : null;
}

/** Pastikan ada batas run tepat di posisi pos (pecah w:t / w:r bila perlu). */
function splitRunAt(xml: XMLDocument, p: Element, pos: number) {
  let segs = segmentsOf(p);
  const inside = segs.find((s) => s.kind === "t" && pos > s.start && pos < s.start + s.len);
  if (inside) {
    const txt = inside.el.textContent || "";
    const off = pos - inside.start;
    setText(inside.el, txt.slice(0, off));
    const t2 = xml.createElementNS(W_NS, "w:t");
    setText(t2, txt.slice(off));
    inside.el.parentNode!.insertBefore(t2, inside.el.nextSibling);
    segs = segmentsOf(p);
  }
  const at = segs.find((s) => s.start === pos);
  if (!at) return;
  const run = runOf(at.el);
  if (!run) return;
  const hasContentBefore = (() => {
    for (let c = at.el.previousElementSibling; c; c = c.previousElementSibling) if (!isW(c, "rPr")) return true;
    return false;
  })();
  if (!hasContentBefore) return;
  const r2 = xml.createElementNS(W_NS, "w:r");
  const rPr = childW(run, "rPr");
  if (rPr) r2.appendChild(rPr.cloneNode(true));
  let c: Element | null = at.el;
  while (c) { const next: Element | null = c.nextElementSibling; r2.appendChild(c); c = next; }
  run.parentNode!.insertBefore(r2, run.nextSibling);
}

/** Set tebal/miring/garis bawah eksplisit pada rentang [s, e) paragraf. */
export function setRangeFormat(xml: XMLDocument, p: Element, s: number, e: number, attr: number, mask: number) {
  if (e <= s || !mask) return;
  splitRunAt(xml, p, s);
  splitRunAt(xml, p, e);
  const runs = new Set<Element>();
  for (const seg of segmentsOf(p)) {
    if (seg.start >= s && seg.start + seg.len <= e) { const r = runOf(seg.el); if (r) runs.add(r); }
  }
  for (const r of runs) {
    const rPr = ensureRPr(xml, r);
    if (mask & ATTR_B) { const v = attr & ATTR_B ? undefined : "0"; setOrderedChild(xml, rPr, "b", RPR_ORDER, v); setOrderedChild(xml, rPr, "bCs", RPR_ORDER, v); }
    if (mask & ATTR_I) { const v = attr & ATTR_I ? undefined : "0"; setOrderedChild(xml, rPr, "i", RPR_ORDER, v); setOrderedChild(xml, rPr, "iCs", RPR_ORDER, v); }
    if (mask & ATTR_U) setOrderedChild(xml, rPr, "u", RPR_ORDER, attr & ATTR_U ? "single" : "none");
  }
}

const JC: Record<string, string> = { left: "left", center: "center", right: "right", justify: "both" };
export function setParagraphAlign(xml: XMLDocument, p: Element, align: string) {
  const v = JC[align];
  if (!v) return;
  setOrderedChild(xml, ensurePPr(xml, p), "jc", PPR_ORDER, v);
}

/** Set line-spacing (rasio thd 1 baris) di pPr, mis. 1 = single, 1.5, 2 = double. */
export function setParagraphLineSpacing(xml: XMLDocument, p: Element, ratio: number) {
  const pPr = ensurePPr(xml, p);
  let el = childW(pPr, "spacing");
  if (!el) {
    el = xml.createElementNS(W_NS, "w:spacing");
    const idx = PPR_ORDER.indexOf("spacing");
    const before = Array.from(pPr.children).find((c) => c.namespaceURI === W_NS && PPR_ORDER.indexOf(c.localName) > idx);
    pPr.insertBefore(el, before || null);
  }
  el.setAttributeNS(W_NS, "w:line", String(Math.round(ratio * 240)));
  el.setAttributeNS(W_NS, "w:lineRule", "auto");
}

/** Set rFonts (ascii/hAnsi/cs) pada sebuah <w:rPr>. */
function setRFonts(xml: XMLDocument, rPr: Element, name: string) {
  let el = childW(rPr, "rFonts");
  if (!el) {
    el = xml.createElementNS(W_NS, "w:rFonts");
    const idx = RPR_ORDER.indexOf("rFonts");
    const before = Array.from(rPr.children).find((c) => c.namespaceURI === W_NS && RPR_ORDER.indexOf(c.localName) > idx);
    rPr.insertBefore(el, before || null);
  }
  for (const attr of ["ascii", "hAnsi", "cs"]) el.setAttributeNS(W_NS, `w:${attr}`, name);
}

/** Set font-family eksplisit pada rentang [s, e) paragraf. */
export function setRangeFont(xml: XMLDocument, p: Element, s: number, e: number, font: string) {
  if (e <= s) return;
  splitRunAt(xml, p, s);
  splitRunAt(xml, p, e);
  const runs = new Set<Element>();
  for (const seg of segmentsOf(p)) {
    if (seg.start >= s && seg.start + seg.len <= e) { const r = runOf(seg.el); if (r) runs.add(r); }
  }
  for (const r of runs) setRFonts(xml, ensureRPr(xml, r), font);
}

/** Set ukuran font eksplisit (half-point) pada rentang [s, e) paragraf. */
export function setRangeSize(xml: XMLDocument, p: Element, s: number, e: number, sizeHp: number) {
  if (e <= s) return;
  splitRunAt(xml, p, s);
  splitRunAt(xml, p, e);
  const runs = new Set<Element>();
  for (const seg of segmentsOf(p)) {
    if (seg.start >= s && seg.start + seg.len <= e) { const r = runOf(seg.el); if (r) runs.add(r); }
  }
  const v = String(Math.round(sizeHp));
  for (const r of runs) {
    const rPr = ensureRPr(xml, r);
    setOrderedChild(xml, rPr, "sz", RPR_ORDER, v);
    setOrderedChild(xml, rPr, "szCs", RPR_ORDER, v);
  }
}

/** Set warna teks eksplisit (hex 6 digit tanpa "#") pada rentang [s, e) paragraf. */
export function setRangeColor(xml: XMLDocument, p: Element, s: number, e: number, hex: string) {
  if (e <= s || !/^[0-9A-Fa-f]{6}$/.test(hex)) return;
  splitRunAt(xml, p, s);
  splitRunAt(xml, p, e);
  const runs = new Set<Element>();
  for (const seg of segmentsOf(p)) {
    if (seg.start >= s && seg.start + seg.len <= e) { const r = runOf(seg.el); if (r) runs.add(r); }
  }
  for (const r of runs) {
    const el = setOrderedChild(xml, ensureRPr(xml, r), "color", RPR_ORDER, hex.toUpperCase());
    // Warna tema (themeColor/Shade/Tint) menimpa w:val — buang supaya warna eksplisit yang berlaku.
    for (const a of ["themeColor", "themeShade", "themeTint"]) el.removeAttributeNS(W_NS, a);
  }
}

/**
 * Terapkan format karakter: bandingkan atribut lama vs baru, dengan
 * penyelarasan teks lama->baru lewat diff. Hanya karakter yang atributnya
 * BERUBAH (atau karakter sisipan yang atributnya beda dari tempat sisipnya)
 * yang diberi format eksplisit — sisanya tidak disentuh.
 */
function applyAttrChanges(xml: XMLDocument, p: Element, oldText: string, newText: string, oldAttrs: number[], newAttrs: number[]) {
  const want: (number | null)[] = new Array(newText.length).fill(null); // null = biarkan
  let io = 0, inew = 0;
  for (const part of diffChars(oldText, newText)) {
    const n = part.value.length;
    if (part.removed) { io += n; continue; }
    if (part.added) {
      const inherit = io > 0 ? oldAttrs[io - 1] ?? 0 : oldAttrs[0] ?? 0;
      for (let k = 0; k < n; k++) if ((newAttrs[inew + k] ?? inherit) !== inherit) want[inew + k] = newAttrs[inew + k];
      inew += n;
      continue;
    }
    for (let k = 0; k < n; k++) if ((newAttrs[inew + k] ?? 0) !== (oldAttrs[io + k] ?? 0)) want[inew + k] = newAttrs[inew + k];
    io += n; inew += n;
  }
  // Kelompokkan jadi rentang berurutan dgn nilai sama; mask = bit yang beda.
  let k = 0;
  while (k < want.length) {
    if (want[k] === null) { k++; continue; }
    const v = want[k]!;
    let e = k + 1;
    while (e < want.length && want[e] === v) e++;
    setRangeFormat(xml, p, k, e, v, ATTR_B | ATTR_I | ATTR_U);
    k = e;
  }
}

/**
 * Terapkan perubahan font-family: sama seperti applyAttrChanges, tapi
 * membandingkan string (bukan bitmask). Hanya rentang yang nilainya BERUBAH
 * dari baseline lama yang diberi rFonts eksplisit.
 */
function applyFontChanges(xml: XMLDocument, p: Element, oldText: string, newText: string, oldFonts: string[], newFonts: string[]) {
  const want: (string | null)[] = new Array(newText.length).fill(null);
  let io = 0, inew = 0;
  for (const part of diffChars(oldText, newText)) {
    const n = part.value.length;
    if (part.removed) { io += n; continue; }
    if (part.added) {
      const inherit = io > 0 ? oldFonts[io - 1] : oldFonts[0];
      for (let k = 0; k < n; k++) if ((newFonts[inew + k] ?? inherit) !== inherit) want[inew + k] = newFonts[inew + k];
      inew += n;
      continue;
    }
    for (let k = 0; k < n; k++) if (newFonts[inew + k] !== oldFonts[io + k]) want[inew + k] = newFonts[inew + k];
    io += n; inew += n;
  }
  let k = 0;
  while (k < want.length) {
    if (want[k] === null) { k++; continue; }
    const v = want[k]!;
    let e = k + 1;
    while (e < want.length && want[e] === v) e++;
    setRangeFont(xml, p, k, e, v);
    k = e;
  }
}

/** Terapkan perubahan ukuran font (half-point), pola sama dgn applyFontChanges. */
function applySizeChanges(xml: XMLDocument, p: Element, oldText: string, newText: string, oldSizes: number[], newSizes: number[]) {
  const want: (number | null)[] = new Array(newText.length).fill(null);
  let io = 0, inew = 0;
  for (const part of diffChars(oldText, newText)) {
    const n = part.value.length;
    if (part.removed) { io += n; continue; }
    if (part.added) {
      const inherit = io > 0 ? oldSizes[io - 1] : oldSizes[0];
      for (let k = 0; k < n; k++) if ((newSizes[inew + k] ?? inherit) !== inherit) want[inew + k] = newSizes[inew + k];
      inew += n;
      continue;
    }
    for (let k = 0; k < n; k++) if (newSizes[inew + k] !== oldSizes[io + k]) want[inew + k] = newSizes[inew + k];
    io += n; inew += n;
  }
  let k = 0;
  while (k < want.length) {
    if (want[k] === null) { k++; continue; }
    const v = want[k]!;
    let e = k + 1;
    while (e < want.length && want[e] === v) e++;
    setRangeSize(xml, p, k, e, v);
    k = e;
  }
}

/** Terapkan perubahan warna teks, pola sama dgn applyFontChanges (hanya yang BERUBAH yang ditulis). */
function applyColorChanges(xml: XMLDocument, p: Element, oldText: string, newText: string, oldColors: string[], newColors: string[]) {
  const want: (string | null)[] = new Array(newText.length).fill(null);
  let io = 0, inew = 0;
  for (const part of diffChars(oldText, newText)) {
    const n = part.value.length;
    if (part.removed) { io += n; continue; }
    if (part.added) {
      const inherit = io > 0 ? oldColors[io - 1] : oldColors[0];
      for (let k = 0; k < n; k++) if ((newColors[inew + k] ?? inherit) !== inherit) want[inew + k] = newColors[inew + k];
      inew += n;
      continue;
    }
    for (let k = 0; k < n; k++) if (newColors[inew + k] !== oldColors[io + k]) want[inew + k] = newColors[inew + k];
    io += n; inew += n;
  }
  let k = 0;
  while (k < want.length) {
    if (want[k] === null) { k++; continue; }
    const v = want[k]!;
    let e = k + 1;
    while (e < want.length && want[e] === v) e++;
    if (v) setRangeColor(xml, p, k, e, v);
    k = e;
  }
}

// ---------------------------------------------------------------------------
// List Word sungguhan (word/numbering.xml + w:numPr)
//
// Paste list di editor upload menghasilkan paragraf berpenanda (lihat
// paste-clean.ts, mode "flat"). Saat simpan, penanda dibuang dan paragraf diberi
// numPr yang menunjuk ke definisi list di numbering.xml. Definisi dibuat sendiri
// (bullet + desimal, 9 level) dan ditambahkan di AKHIR numbering.xml tanpa
// menyentuh definisi yang sudah ada. List bernomor baru selalu mulai dari 1
// (satu w:num per list + startOverride).
// ---------------------------------------------------------------------------
export interface ListSpec { kind: "bullet" | "decimal"; level: number; id: string; prefixLen: number }

const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

function listsSupported(model: DocxModel): boolean {
  return !model.numberingBroken;
}

function ensureNumbering(model: DocxModel): XMLDocument {
  if (!model.numbering) {
    model.numbering = new DOMParser().parseFromString(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W_NS}"/>`, "application/xml");
    model.numberingIsNew = true;
  }
  return model.numbering;
}

function maxAttrOf(root: Element, local: string, attr: string): number {
  let max = -1;
  for (const c of Array.from(root.children)) {
    if (!isW(c, local)) continue;
    const v = parseInt(c.getAttributeNS(W_NS, attr) || c.getAttribute(`w:${attr}`) || "", 10);
    if (!isNaN(v) && v > max) max = v;
  }
  return max;
}

function lvlXml(kind: "bullet" | "decimal", i: number): string {
  const left = 720 * (i + 1);
  const ind = `<w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr>`;
  if (kind === "bullet") {
    const ch = ["•", "◦", "▪"][i % 3];
    return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${ch}"/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
  }
  // Semua level angka desimal: yang ditempel hanya list angka ("1.") — di level bersarang pun
  // tetap tampil "1." seperti di sumbernya (bukan "a." ala list bertingkat Word).
  const fmt = "decimal";
  return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/>${ind}</w:lvl>`;
}

function importXml(xml: XMLDocument, fragment: string): Element {
  const parsed = new DOMParser().parseFromString(fragment, "application/xml");
  return xml.importNode(parsed.documentElement, true) as Element;
}

function addAbstractNum(model: DocxModel, kind: "bullet" | "decimal"): number {
  const xml = ensureNumbering(model);
  const root = xml.documentElement;
  const id = maxAttrOf(root, "abstractNum", "abstractNumId") + 1;
  const lvls = Array.from({ length: 9 }, (_, i) => lvlXml(kind, i)).join("");
  const node = importXml(xml, `<w:abstractNum xmlns:w="${W_NS}" w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`);
  // Skema: semua w:abstractNum harus sebelum w:num.
  const firstNum = Array.from(root.children).find((c) => isW(c, "num"));
  root.insertBefore(node, firstNum || null);
  model.numberingDirty = true;
  return id;
}

function addNum(model: DocxModel, abstractId: number, restart: boolean): number {
  const xml = ensureNumbering(model);
  const root = xml.documentElement;
  // numId 0 dicadangkan Word utk "tanpa penomoran" — id pertama harus 1.
  const id = Math.max(1, maxAttrOf(root, "num", "numId") + 1);
  const ov = restart
    ? Array.from({ length: 9 }, (_, i) => `<w:lvlOverride w:ilvl="${i}"><w:startOverride w:val="1"/></w:lvlOverride>`).join("")
    : "";
  const node = importXml(xml, `<w:num xmlns:w="${W_NS}" w:numId="${id}"><w:abstractNumId w:val="${abstractId}"/>${ov}</w:num>`);
  // w:numIdMacAtCleanup (bila ada) harus tetap yang terakhir.
  const tail = Array.from(root.children).find((c) => isW(c, "numIdMacAtCleanup"));
  root.insertBefore(node, tail || null);
  model.numberingDirty = true;
  return id;
}

/** numId untuk sebuah list baru; null kalau numbering.xml tak bisa dipakai. */
function numIdFor(model: DocxModel, spec: ListSpec, cache: Map<string, number>): number | null {
  if (!listsSupported(model)) return null;
  model.numState = model.numState || {};
  if (spec.kind === "bullet") {
    if (model.numState.bullet === undefined) model.numState.bullet = addNum(model, addAbstractNum(model, "bullet"), false);
    return model.numState.bullet;
  }
  const cached = cache.get(spec.id);
  if (cached !== undefined) return cached;
  if (model.numState.decimalAbs === undefined) model.numState.decimalAbs = addAbstractNum(model, "decimal");
  const n = addNum(model, model.numState.decimalAbs, true);
  cache.set(spec.id, n);
  return n;
}

/** Jadikan paragraf butir list (numPr); indentasi mengikuti definisi list, bukan w:ind lama. */
export function setParagraphList(xml: XMLDocument, p: Element, numId: number, level: number) {
  const pPr = ensurePPr(xml, p);
  for (const tag of ["numPr", "ind"]) { const e = childW(pPr, tag); if (e) pPr.removeChild(e); }
  const numPr = xml.createElementNS(W_NS, "w:numPr");
  const il = xml.createElementNS(W_NS, "w:ilvl");
  il.setAttributeNS(W_NS, "w:val", String(Math.max(0, Math.min(8, level))));
  const ni = xml.createElementNS(W_NS, "w:numId");
  ni.setAttributeNS(W_NS, "w:val", String(numId));
  numPr.appendChild(il);
  numPr.appendChild(ni);
  const idx = PPR_ORDER.indexOf("numPr");
  const before = Array.from(pPr.children).find((c) => c.namespaceURI === W_NS && PPR_ORDER.indexOf(c.localName) > idx);
  pPr.insertBefore(numPr, before || null);
}

/** Daftarkan numbering.xml yang baru dibuat di [Content_Types].xml dan word/_rels/document.xml.rels. */
async function registerNumberingPart(zip: JSZip) {
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    const ct = await ctFile.async("string");
    if (!/PartName="\/word\/numbering\.xml"/.test(ct)) {
      zip.file("[Content_Types].xml", ct.replace("</Types>", `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CT}"/></Types>`));
    }
  }
  const relsPath = "word/_rels/document.xml.rels";
  const relsFile = zip.file(relsPath);
  const rels = relsFile ? await relsFile.async("string") : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  if (/Target="numbering\.xml"/.test(rels)) return;
  let max = 0;
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, parseInt(m[1], 10));
  zip.file(relsPath, rels.replace("</Relationships>", `<Relationship Id="rId${max + 1}" Type="${NUMBERING_REL}" Target="numbering.xml"/></Relationships>`));
}

function paragraphIsRemovable(p: Element): boolean {
  if (p.getElementsByTagNameNS(W_NS, "sectPr").length) return false;
  for (const tag of ["drawing", "pict", "object", "fldChar", "fldSimple"]) if (p.getElementsByTagNameNS(W_NS, tag).length) return false;
  const parent = p.parentElement;
  if (parent && isW(parent, "tc")) {
    const ps = Array.from(parent.children).filter((c) => isW(c, "p"));
    if (ps.length <= 1) return false; // sel tabel wajib punya minimal 1 paragraf
  }
  return true;
}

/** Paragraf baru berdasarkan paragraf acuan (gaya paragraf & format run). */
function makeParagraph(
  xml: XMLDocument, anchor: Element, text: string, attrs: number[], baseAttr: number,
  align?: string, fonts?: string[], baseFont?: string, sizes?: number[], baseSize?: number, lineSpacing?: number,
  colors?: string[], baseColor?: string, list?: { numId: number; level: number },
): Element {
  const p = xml.createElementNS(W_NS, "w:p");
  const aPPr = childW(anchor, "pPr");
  if (aPPr) {
    const pPr = aPPr.cloneNode(true) as Element;
    const sect = childW(pPr, "sectPr");
    if (sect) pPr.removeChild(sect); // section break hanya milik paragraf asal
    p.appendChild(pPr);
  }
  // Format run diambil dari run berteks terakhir paragraf acuan.
  const textRuns = Array.from(anchor.getElementsByTagNameNS(W_NS, "r")).filter((r) => r.getElementsByTagNameNS(W_NS, "t").length);
  const baseRun = textRuns[textRuns.length - 1];
  const r = xml.createElementNS(W_NS, "w:r");
  const rPr = baseRun ? childW(baseRun, "rPr") : (aPPr ? childW(aPPr, "rPr") : undefined);
  if (rPr) r.appendChild(rPr.cloneNode(true));
  for (const n of nodesForText(xml, text)) r.appendChild(n);
  if (!text) { const t = xml.createElementNS(W_NS, "w:t"); setText(t, ""); r.appendChild(t); }
  p.appendChild(r);
  if (align) setParagraphAlign(xml, p, align);
  if (lineSpacing && Math.abs(lineSpacing - 1) > 0.02) setParagraphLineSpacing(xml, p, lineSpacing);
  // Karakter yang formatnya beda dari format dasar diberi format eksplisit.
  let k = 0;
  while (k < attrs.length) {
    if (attrs[k] === baseAttr) { k++; continue; }
    const v = attrs[k];
    let e = k + 1;
    while (e < attrs.length && attrs[e] === v) e++;
    setRangeFormat(xml, p, k, e, v, ATTR_B | ATTR_I | ATTR_U);
    k = e;
  }
  if (fonts && baseFont !== undefined) {
    let k2 = 0;
    while (k2 < fonts.length) {
      if (fonts[k2] === baseFont) { k2++; continue; }
      const v = fonts[k2];
      let e = k2 + 1;
      while (e < fonts.length && fonts[e] === v) e++;
      setRangeFont(xml, p, k2, e, v);
      k2 = e;
    }
  }
  if (sizes && baseSize !== undefined) {
    let k3 = 0;
    while (k3 < sizes.length) {
      if (sizes[k3] === baseSize) { k3++; continue; }
      const v = sizes[k3];
      let e = k3 + 1;
      while (e < sizes.length && sizes[e] === v) e++;
      setRangeSize(xml, p, k3, e, v);
      k3 = e;
    }
  }
  if (colors && baseColor !== undefined) {
    let k4 = 0;
    while (k4 < colors.length) {
      if (colors[k4] === baseColor || !colors[k4]) { k4++; continue; }
      const v = colors[k4];
      let e = k4 + 1;
      while (e < colors.length && colors[e] === v) e++;
      setRangeColor(xml, p, k4, e, v);
      k4 = e;
    }
  }
  if (list) setParagraphList(xml, p, list.numId, list.level);
  return p;
}

export interface EditedParagraph {
  index: number; oldText: string; newText: string; oldAttrs: number[]; newAttrs: number[]; oldAlign: string; newAlign: string;
  oldFonts: string[]; newFonts: string[]; oldSizes: number[]; newSizes: number[]; oldLineSpacing: number; newLineSpacing: number;
  oldColors?: string[]; newColors?: string[];
}
export interface NewParagraph {
  text: string; attrs: number[]; align?: string; fonts: string[]; sizes: number[]; lineSpacing?: number;
  colors?: string[];
  /** Butir list hasil paste: awal `prefixLen` karakter adalah penanda ("• "/"1. ") yang dibuang saat simpan. */
  list?: ListSpec;
  /** Gaya "tanpa format" di tempat paragraf ini (lihat defaultCharStyle) — karakter bergaya ini mewarisi gaya paragraf acuan. */
  inherit?: { font: string; size: number; color: string };
}
export interface DocxEditPlan {
  edits: EditedParagraph[];
  deleted: number[];
  /** Paragraf baru; afterIndex -1 = sebelum paragraf terikat pertama. */
  inserts: { afterIndex: number; anchorIndex: number; baseAttr: number; baseFont: string; baseSize: number; baseColor?: string; paragraphs: NewParagraph[] }[];
}

/** Terapkan rencana edit ke model (in-place). Elemen w:p lama tetap sama. */
export function applyPlan(model: DocxModel, plan: DocxEditPlan) {
  const { xml, paragraphs } = model;
  for (const e of plan.edits) {
    const p = paragraphs[e.index];
    if (!p) continue;
    applyParagraphEdit(xml, p, e.oldText, e.newText);
    applyAttrChanges(xml, p, e.oldText, e.newText, e.oldAttrs, e.newAttrs);
    applyFontChanges(xml, p, e.oldText, e.newText, e.oldFonts, e.newFonts);
    applySizeChanges(xml, p, e.oldText, e.newText, e.oldSizes, e.newSizes);
    if (e.oldColors && e.newColors) applyColorChanges(xml, p, e.oldText, e.newText, e.oldColors, e.newColors);
    if (e.newAlign && e.newAlign !== e.oldAlign) setParagraphAlign(xml, p, e.newAlign);
    if (e.newLineSpacing && Math.abs(e.newLineSpacing - e.oldLineSpacing) > 0.02) setParagraphLineSpacing(xml, p, e.newLineSpacing);
  }
  const listCache = new Map<string, number>();
  for (const ins of plan.inserts) {
    const anchor = paragraphs[ins.anchorIndex];
    if (!anchor) continue;
    // Butir list: buang penanda teks di awal & ganti dgn numPr — hanya kalau
    // definisi list bisa dibuat; kalau tidak, penanda dibiarkan sbg teks biasa.
    const build = (np: NewParagraph): Element => {
      let { text, attrs, fonts, sizes } = np;
      let colors = np.colors;
      let list: { numId: number; level: number } | undefined;
      if (np.inherit) {
        const inh = np.inherit;
        fonts = fonts.map((f) => (f === inh.font ? ins.baseFont : f));
        sizes = sizes.map((s) => (s === inh.size ? ins.baseSize : s));
        if (colors && ins.baseColor !== undefined) colors = colors.map((c) => (c === inh.color ? ins.baseColor! : c));
      }
      const spec = np.list;
      if (spec && spec.prefixLen > 0 && text.length >= spec.prefixLen) {
        const numId = numIdFor(model, spec, listCache);
        if (numId !== null) {
          const n = spec.prefixLen;
          text = text.slice(n); attrs = attrs.slice(n); fonts = fonts.slice(n); sizes = sizes.slice(n);
          colors = colors?.slice(n);
          list = { numId, level: spec.level };
        }
      }
      return makeParagraph(xml, anchor, text, attrs, ins.baseAttr, np.align, fonts, ins.baseFont, sizes, ins.baseSize, np.lineSpacing, colors, ins.baseColor, list);
    };
    if (ins.afterIndex === -1) {
      for (const np of ins.paragraphs) anchor.parentNode!.insertBefore(build(np), anchor);
    } else {
      let cursor: Element = paragraphs[ins.afterIndex];
      for (const np of ins.paragraphs) {
        const el = build(np);
        cursor.parentNode!.insertBefore(el, cursor.nextSibling);
        cursor = el;
      }
    }
  }
  for (const idx of plan.deleted) {
    const p = paragraphs[idx];
    if (!p || !p.parentNode) continue;
    if (paragraphIsRemovable(p)) p.parentNode.removeChild(p);
    else applyParagraphEdit(xml, p, paragraphText(p), "");
  }
}

export async function serializeDocx(model: DocxModel): Promise<ArrayBuffer> {
  const serialized = new XMLSerializer().serializeToString(model.xml);
  model.zip.file("word/document.xml", serialized.startsWith("<?xml") ? serialized : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${serialized}`);
  // numbering.xml hanya ditulis ulang bila ada list baru; kalau tidak, tetap byte-per-byte seperti aslinya.
  if (model.numbering && model.numberingDirty) {
    const ns = new XMLSerializer().serializeToString(model.numbering);
    model.zip.file("word/numbering.xml", ns.startsWith("<?xml") ? ns : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${ns}`);
    if (model.numberingIsNew) await registerNumberingPart(model.zip);
  }
  return model.zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

function rowOf(p: Element): Element | null {
  for (let c = p.parentElement; c; c = c.parentElement) if (isW(c, "tr")) return c;
  return null;
}

/** Tambah baris tabel (salinan kosong baris milik paragraf p) di bawahnya. */
export function addTableRowAfter(p: Element | undefined): boolean {
  const tr = p && rowOf(p);
  if (!tr) return false;
  const clone = tr.cloneNode(true) as Element;
  for (const tc of Array.from(clone.getElementsByTagNameNS(W_NS, "tc"))) {
    const ps = Array.from(tc.children).filter((c) => isW(c, "p"));
    ps.slice(1).forEach((x) => tc.removeChild(x));
    const first = ps[0];
    if (first) {
      for (const t of Array.from(first.getElementsByTagNameNS(W_NS, "t"))) setText(t, "");
      for (const tag of ["br", "tab", "drawing", "pict"]) for (const x of Array.from(first.getElementsByTagNameNS(W_NS, tag))) x.parentNode?.removeChild(x);
    }
  }
  tr.parentNode!.insertBefore(clone, tr.nextSibling);
  return true;
}

export function removeTableRow(p: Element | undefined): boolean {
  const tr = p && rowOf(p);
  if (!tr) return false;
  const rows = Array.from(tr.parentElement!.children).filter((c) => isW(c, "tr"));
  if (rows.length <= 1) return false;
  tr.parentNode!.removeChild(tr);
  return true;
}
