import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { acceptAllChanges, originalTextOf } from "./review-changes.js";

// PEMISAHAN MARKUP INTERNAL ↔ EKSTERNAL (satu master file, riwayat utuh).
//
//  - MASTER = berkas dokumen utama di server. Menyimpan SEMUA markup (Track Changes + komentar) internal
//    maupun eksternal dari awal sampai akhir. Tidak pernah dikirim ke pihak eksternal.
//  - SALINAN EKSTERNAL (buildExternalCopy) = turunan master untuk satu ronde review eksternal: semua perubahan
//    diterima, semua komentar dibuang, dan bagian berkas yang menyimpan komentar/daftar penulis dihapus
//    (komentar & Track Changes tersimpan DI DALAM .docx, jadi menyembunyikannya di layar saja tidak cukup).
//  - Setelah eksternal selesai, mergeExternalMarkup() memindahkan markup eksternal (sisipan, coretan, komentar,
//    paragraf baru) ke MASTER tanpa menyentuh markup internal yang sudah ada.
//
// Modul ini murni (tanpa akses DB/berkas) supaya bisa diuji: lihat scripts/review-isolation-test.ts.
// Bila ada perubahan yang belum didukung penggabung otomatis (perubahan format, struktur tabel, paragraf
// dipecah/digabung, markup di header/footer), hasilnya { ok: false } dan master TIDAK diubah — pemanggil
// menyimpan berkas eksternal apa adanya agar tidak ada yang hilang.

const W_COMMENT_PARTS = /^word\/(comments|commentsExtended|commentsIds|commentsExtensible|people)\.xml$/;
const W_TRACK_PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
const LEAK_TAGS = /<w:(ins|del|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|trPrChange|tcPrChange|numberingChange|moveFrom|moveTo|commentReference|commentRangeStart|commentRangeEnd|cellIns|cellDel|cellMerge)\b/;
const UNSUPPORTED_TAGS = /<w:(rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|trPrChange|tcPrChange|tblPrExChange|numberingChange|moveFrom|moveTo|moveFromRangeStart|moveToRangeStart|cellIns|cellDel|cellMerge)\b/;
const REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

// ---------------------------------------------------------------------------------------------------------------------
// 1. Salinan eksternal
// ---------------------------------------------------------------------------------------------------------------------

/** Gagal bila berkas masih memuat jejak internal (Track Changes, komentar, daftar penulis). */
export async function assertNoInternalTraces(docx: Buffer): Promise<void> {
  const zip = await JSZip.loadAsync(docx);
  for (const name of Object.keys(zip.files)) {
    if (W_COMMENT_PARTS.test(name)) throw new Error(`Salinan eksternal masih memuat ${name}`);
    if (W_TRACK_PARTS.test(name)) {
      const xml = await zip.file(name)!.async("string");
      if (LEAK_TAGS.test(xml)) throw new Error(`Salinan eksternal masih memuat markup di ${name}`);
    }
  }
}

/**
 * Salinan untuk pihak eksternal: terima semua perubahan, buang komentar, hapus bagian berkas komentar & daftar
 * penulis, kosongkan "terakhir diubah oleh". Setelah dibuat, diperiksa ulang (fail-closed): bila masih ada jejak,
 * melempar error dan TIDAK PERNAH mengembalikan berkas.
 */
export async function buildExternalCopy(master: Buffer): Promise<Buffer> {
  const accepted = await acceptAllChanges(master);
  const zip = await JSZip.loadAsync(accepted);
  const dropped = Object.keys(zip.files).filter((n) => W_COMMENT_PARTS.test(n));
  for (const n of dropped) zip.remove(n);
  if (dropped.length) {
    const relsName = "word/_rels/document.xml.rels";
    const rels = await zip.file(relsName)?.async("string");
    if (rels) zip.file(relsName, rels.replace(/<Relationship\b[^>]*\bTarget="(?:comments|commentsExtended|commentsIds|commentsExtensible|people)\.xml"[^>]*\/>/g, ""));
    const ct = await zip.file("[Content_Types].xml")?.async("string");
    if (ct) zip.file("[Content_Types].xml", ct.replace(/<Override\b[^>]*\bPartName="\/word\/(?:comments|commentsExtended|commentsIds|commentsExtensible|people)\.xml"[^>]*\/>/g, ""));
  }
  const core = await zip.file("docProps/core.xml")?.async("string");
  if (core) zip.file("docProps/core.xml", core.replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, "<cp:lastModifiedBy></cp:lastModifiedBy>"));
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await assertNoInternalTraces(out);
  return out;
}

// ---------------------------------------------------------------------------------------------------------------------
// 2. Penggabungan markup eksternal ke master
// ---------------------------------------------------------------------------------------------------------------------

export interface MergeStats { inserted: number; deleted: number; comments: number; paragraphsAdded: number; paragraphsRemoved: number }
export type MergeResult =
  | { ok: true; buf: Buffer; stats: MergeStats; reason?: undefined }
  | { ok: false; reason: string; buf?: undefined; stats?: undefined };

class Unsupported extends Error {}
const unsupported = (msg: string): never => { throw new Unsupported(msg); };

const attrOf = (e: Element, name: string) => e.getAttribute(`w:${name}`) || "";
const kids = (e: Element): Element[] => Array.from(e.childNodes).filter((n) => n.nodeType === 1) as Element[];
const byName = (e: Element, name: string) => kids(e).find((k) => k.nodeName === name);
const allOf = (doc: Document | Element, tag: string) => Array.from(doc.getElementsByTagName(tag)) as Element[];

/** Telusuri isi paragraf MILIKNYA saja: tidak masuk paragraf bersarang (kotak teks) dan tidak masuk ke dalam w:r. */
function walkOwn(p: Element, fn: (n: Element) => boolean | void): void {
  const rec = (parent: Element, top: boolean) => {
    for (const c of kids(parent)) {
      if (top && c.nodeName === "w:pPr") continue;
      if (c.nodeName === "w:p") continue;
      const descend = fn(c);
      if (descend !== false && c.nodeName !== "w:r") rec(c, false);
    }
  };
  rec(p, true);
}
function flagsOf(n: Element, p: Element): { inIns: boolean; inDel: boolean } {
  let inIns = false, inDel = false;
  for (let a = n.parentNode as Element | null; a && a !== p; a = a.parentNode as Element | null) {
    if (a.nodeName === "w:ins") inIns = true;
    else if (a.nodeName === "w:del") inDel = true;
  }
  return { inIns, inDel };
}
const isCommentRefRun = (r: Element) => kids(r).some((k) => k.nodeName === "w:commentReference");
/** Satuan teks run: tiap karakter (titik kode) + tiap tab. Isi non-teks (gambar, pemisah baris, field) tidak dihitung. */
function runUnits(r: Element): string[] {
  const out: string[] = [];
  for (const c of kids(r)) {
    if (c.nodeName === "w:t" || c.nodeName === "w:delText") for (const ch of Array.from(c.textContent || "")) out.push(ch);
    else if (c.nodeName === "w:tab") out.push("\t");
  }
  return out;
}
const paraMark = (p: Element, tag: "w:ins" | "w:del") => {
  const ppr = byName(p, "w:pPr"); const rpr = ppr && byName(ppr, "w:rPr");
  return (rpr && byName(rpr, tag)) || null;
};
/** Sama dengan penentu di acceptAllChanges: paragraf yang tanda paragrafnya dihapus & isinya habis ikut hilang. */
function paraTextNoDel(p: Element): string {
  let out = "";
  const walk = (n: Node) => {
    if (n.nodeType !== 1) return;
    const e = n as Element;
    if (e.nodeName === "w:t") out += e.textContent || "";
    else if (e.nodeName === "w:delText") return;
    else if (e.nodeName === "w:tab") out += "\t";
    else for (const c of Array.from(e.childNodes)) walk(c);
  };
  walk(p);
  return out;
}
const removedByAccept = (p: Element) => !!paraMark(p, "w:del") && !paraTextNoDel(p).replace(/\s+/g, "");

interface ParaInfo { baseUnits: string[]; hasRuns: boolean; allIns: boolean; remainingBase: number }
function infoOfReturned(p: Element): ParaInfo {
  const baseUnits: string[] = []; let hasRuns = false, allIns = true, remainingBase = 0;
  walkOwn(p, (n) => {
    if (n.nodeName !== "w:r" || isCommentRefRun(n)) return true;
    hasRuns = true;
    const f = flagsOf(n, p);
    if (!f.inIns) { allIns = false; for (const u of runUnits(n)) { baseUnits.push(u); if (!f.inDel) remainingBase++; } }
    return false;
  });
  return { baseUnits, hasRuns, allIns: hasRuns && allIns, remainingBase };
}
function acceptedUnitsOfMaster(p: Element): string[] {
  const out: string[] = [];
  walkOwn(p, (n) => {
    if (n.nodeName !== "w:r" || isCommentRefRun(n)) return true;
    if (!flagsOf(n, p).inDel) for (const u of runUnits(n)) out.push(u);
    return false;
  });
  return out;
}

type Op =
  | { kind: "ins"; at: number; node: Element }
  | { kind: "del"; from: number; to: number; node: Element }
  | { kind: "cstart" | "cend"; at: number; id: string }
  | { kind: "cref"; at: number; id: string; node: Element };

/** Markup eksternal sebuah paragraf, dinyatakan dalam urutan teks "terima semua" (indeks satuan teks). */
function extractOps(p: Element): Op[] {
  const ops: Op[] = []; let idx = 0;
  walkOwn(p, (n) => {
    if (n.nodeName === "w:ins") { ops.push({ kind: "ins", at: idx, node: n }); return false; }
    if (n.nodeName === "w:del") {
      const units = allOf(n, "w:r").reduce((s, r) => s + runUnits(r).length, 0);
      if (units === 0) unsupported("pihak eksternal menghapus objek non-teks (gambar/field)");
      ops.push({ kind: "del", from: idx, to: idx + units, node: n });
      idx += units; return false;
    }
    if (n.nodeName === "w:commentRangeStart") { ops.push({ kind: "cstart", at: idx, id: attrOf(n, "id") }); return false; }
    if (n.nodeName === "w:commentRangeEnd") { ops.push({ kind: "cend", at: idx, id: attrOf(n, "id") }); return false; }
    if (n.nodeName === "w:r") {
      if (isCommentRefRun(n)) { ops.push({ kind: "cref", at: idx, id: attrOf(byName(n, "w:commentReference")!, "id"), node: n }); return false; }
      idx += runUnits(n).length; return false;
    }
    return true;
  });
  return ops;
}

interface Ctx {
  doc: Document;
  nextTrackId: () => string;
  commentMap: Map<string, string>;
  mapComment: (oldId: string) => string;
}

function setText(doc: Document, e: Element, s: string) {
  while (e.firstChild) e.removeChild(e.firstChild);
  e.appendChild(doc.createTextNode(s));
  e.setAttribute("xml:space", "preserve");
}
function splitRun(doc: Document, r: Element, offset: number): Element {
  const right = doc.createElement("w:r");
  for (const a of Array.from(r.attributes)) right.setAttribute(a.name, a.value);
  const rpr = byName(r, "w:rPr");
  if (rpr) right.appendChild(rpr.cloneNode(true));
  let c = 0, rightStarted = false;
  for (const child of kids(r)) {
    if (child.nodeName === "w:rPr") continue;
    if (child.nodeName === "w:t" || child.nodeName === "w:delText") {
      const u = Array.from(child.textContent || ""); const end = c + u.length;
      if (end <= offset && !(u.length === 0 && c >= offset)) { /* kiri */ }
      else if (c >= offset) { r.removeChild(child); right.appendChild(child); rightStarted = true; }
      else {
        const cut = offset - c; const rt = child.cloneNode(false) as Element;
        setText(doc, child, u.slice(0, cut).join("")); setText(doc, rt, u.slice(cut).join(""));
        right.appendChild(rt); rightStarted = true;
      }
      c = end;
    } else if (child.nodeName === "w:tab") {
      if (c + 1 > offset) { r.removeChild(child); right.appendChild(child); rightStarted = true; }
      c += 1;
    } else if (rightStarted) { r.removeChild(child); right.appendChild(child); }
  }
  r.parentNode!.insertBefore(right, r.nextSibling);
  return right;
}
function wrapDeleted(ctx: Ctx, run: Element, author: string, date: string) {
  const d = ctx.doc.createElement("w:del");
  d.setAttribute("w:id", ctx.nextTrackId()); d.setAttribute("w:author", author); d.setAttribute("w:date", date);
  run.parentNode!.insertBefore(d, run);
  d.appendChild(run);
  for (const t of kids(run).filter((k) => k.nodeName === "w:t")) {
    const dt = ctx.doc.createElement("w:delText");
    setText(ctx.doc, dt, t.textContent || "");
    run.replaceChild(dt, t);
  }
}
/** Sisipkan `node` tepat sebelum `anchor`. Bila anchor ada di dalam sisipan internal (w:ins), sisipan itu dipecah dulu. */
function insertBefore(ctx: Ctx, p: Element, node: Node, anchor: Element | null) {
  if (!anchor) { p.appendChild(node); return; }
  let parent = anchor.parentNode as Element; let ref: Node = anchor;
  if (parent.nodeName === "w:ins" || parent.nodeName === "w:del") {
    const first = kids(parent)[0];
    if (first === anchor) { ref = parent; }
    else {
      const right = parent.cloneNode(false) as Element;
      right.setAttribute("w:id", ctx.nextTrackId());
      const after: Node[] = []; for (let s: Node | null = anchor; s; s = s.nextSibling) after.push(s);
      for (const s of after) right.appendChild(s);
      parent.parentNode!.insertBefore(right, parent.nextSibling);
      ref = right;
    }
    parent = ref.parentNode as Element;
  }
  parent.insertBefore(node, ref);
}
function renumberTracks(ctx: Ctx, root: Element) {
  for (const t of [...allOf(root, "w:ins"), ...allOf(root, "w:del")]) t.setAttribute("w:id", ctx.nextTrackId());
}
function remapComments(ctx: Ctx, root: Element) {
  for (const tag of ["w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"]) {
    for (const e of allOf(root, tag)) e.setAttribute("w:id", ctx.mapComment(attrOf(e, "id")));
  }
}
function ensureParaMark(ctx: Ctx, p: Element, tag: "w:ins" | "w:del", author: string, date: string) {
  const doc = ctx.doc;
  let ppr = byName(p, "w:pPr");
  if (!ppr) { ppr = doc.createElement("w:pPr"); p.insertBefore(ppr, p.firstChild); }
  let rpr = byName(ppr, "w:rPr");
  if (!rpr) {
    rpr = doc.createElement("w:rPr");
    const after = kids(ppr).find((k) => k.nodeName === "w:sectPr" || k.nodeName === "w:pPrChange");
    ppr.insertBefore(rpr, after || null);
  }
  if (byName(rpr, tag)) return;
  const m = doc.createElement(tag);
  m.setAttribute("w:id", ctx.nextTrackId()); m.setAttribute("w:author", author); m.setAttribute("w:date", date);
  rpr.insertBefore(m, rpr.firstChild);
}

function applyOps(ctx: Ctx, mp: Element, ops: Op[], counters: MergeStats) {
  const doc = ctx.doc;
  const runs = () => {
    const out: { el: Element; start: number; units: number }[] = []; let pos = 0;
    walkOwn(mp, (n) => {
      if (n.nodeName !== "w:r") return true;
      if (isCommentRefRun(n) || flagsOf(n, mp).inDel) return false;
      const u = runUnits(n).length; out.push({ el: n, start: pos, units: u }); pos += u; return false;
    });
    return { list: out, total: pos };
  };
  const positions = new Set<number>();
  for (const o of ops) { if (o.kind === "del") { positions.add(o.from); positions.add(o.to); } else positions.add(o.at); }
  for (const pos of [...positions].sort((a, b) => a - b)) {
    const hit = runs().list.find((r) => r.start < pos && pos < r.start + r.units);
    if (hit) splitRun(doc, hit.el, pos - hit.start);
  }
  const { list, total } = runs();
  const anchorAt = (pos: number): Element | null => (pos >= total ? null : list.find((r) => r.start === pos)?.el ?? null);
  const delTargets: { run: Element; author: string; date: string }[] = [];
  for (const o of ops) {
    if (o.kind === "del") {
      const author = attrOf(o.node, "author"), date = attrOf(o.node, "date");
      for (const r of list) if (r.units > 0 && r.start >= o.from && r.start + r.units <= o.to) delTargets.push({ run: r.el, author, date });
      counters.deleted++;
    }
  }
  for (const o of ops) {
    if (o.kind === "del") continue;
    const anchor = anchorAt(o.at);
    if (o.kind === "ins") {
      const clone = doc.importNode(o.node, true) as Element;
      renumberTracks(ctx, clone); remapComments(ctx, clone);
      insertBefore(ctx, mp, clone, anchor); counters.inserted++;
    } else if (o.kind === "cstart" || o.kind === "cend") {
      const m = doc.createElement(o.kind === "cstart" ? "w:commentRangeStart" : "w:commentRangeEnd");
      m.setAttribute("w:id", ctx.mapComment(o.id));
      insertBefore(ctx, mp, m, anchor);
    } else {
      const clone = doc.importNode((o as Extract<Op, { kind: "cref" }>).node, true) as Element;
      remapComments(ctx, clone);
      insertBefore(ctx, mp, clone, anchor);
    }
  }
  for (const t of delTargets) wrapDeleted(ctx, t.run, t.author, t.date);
}

const maxIdIn = (xml: string, re: RegExp) => { let m = 0; for (const x of xml.matchAll(re)) m = Math.max(m, Number(x[1])); return m; };
const norm = (s: string) => s.replace(/\s+/g, "");
async function plainAcceptedText(docx: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(await acceptAllChanges(docx));
  const xml = await zip.file("word/document.xml")!.async("string");
  const d = new DOMParser().parseFromString(xml, "text/xml");
  return norm(allOf(d, "w:p").map((p) => paraTextNoDel(p)).join("\n"));
}

/**
 * Gabungkan markup eksternal ke MASTER.
 *  - master   : dokumen dengan semua markup (awal ronde eksternal).
 *  - baseline : salinan eksternal yang dikirim (buildExternalCopy(master)).
 *  - returned : berkas dari editor setelah eksternal memberi markup (baseline + Track Changes/komentar eksternal).
 * Hasilnya divalidasi: (a) "terima semua" hasil gabungan = "terima semua" berkas eksternal, (b) tampilan "Original"
 * hasil gabungan = master, (c) markup internal tidak berkurang. Gagal → { ok: false }, master tidak diubah.
 */
export async function mergeExternalMarkup(input: { master: Buffer; baseline: Buffer; returned: Buffer }): Promise<MergeResult> {
  try {
    return await doMerge(input);
  } catch (err: any) {
    if (err instanceof Unsupported) return { ok: false, reason: err.message };
    return { ok: false, reason: `kesalahan tak terduga saat menggabungkan: ${err?.message || err}` };
  }
}

async function doMerge({ master, baseline, returned }: { master: Buffer; baseline: Buffer; returned: Buffer }): Promise<MergeResult> {
  if ((await plainAcceptedText(baseline)) !== (await plainAcceptedText(master))) unsupported("salinan yang dikirim tidak cocok dengan master");
  const mz = await JSZip.loadAsync(master);
  const rz = await JSZip.loadAsync(returned);
  const mXml = await mz.file("word/document.xml")?.async("string");
  const rXml = await rz.file("word/document.xml")?.async("string");
  if (!mXml || !rXml) return unsupported("berkas bukan dokumen Word yang valid");

  for (const name of Object.keys(rz.files)) {
    if (!W_TRACK_PARTS.test(name)) continue;
    const xml = name === "word/document.xml" ? rXml : await rz.file(name)!.async("string");
    if (UNSUPPORTED_TAGS.test(xml)) unsupported("ada perubahan format/tabel/pemindahan teks (belum didukung penggabung otomatis)");
    if (name !== "word/document.xml" && /<w:(ins|del|commentReference)\b/.test(xml)) unsupported("ada markup di header/footer/catatan kaki");
  }

  const mDoc = new DOMParser().parseFromString(mXml, "text/xml");
  const rDoc = new DOMParser().parseFromString(rXml, "text/xml");
  for (const t of [...allOf(rDoc, "w:ins"), ...allOf(rDoc, "w:del")]) {
    if ((t.parentNode as Element).nodeName === "w:trPr") unsupported("ada baris tabel yang disisipkan/dihapus");
  }

  const masterParas = allOf(mDoc, "w:p");
  const sentParas = masterParas.filter((p) => !removedByAccept(p));
  const sentText = sentParas.map((p) => acceptedUnitsOfMaster(p).join(""));
  const returnedParas = allOf(rDoc, "w:p");

  let nextTrack = Math.max(maxIdIn(mXml, /\bw:id="(\d+)"/g), maxIdIn(rXml, /\bw:id="(\d+)"/g)) + 1000;
  const mcXmlOld = await mz.file("word/comments.xml")?.async("string");
  const rcXml = await rz.file("word/comments.xml")?.async("string");
  let nextComment = Math.max(mcXmlOld ? maxIdIn(mcXmlOld, /<w:comment\b[^>]*\bw:id="(\d+)"/g) : -1, maxIdIn(mXml, /<w:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*\bw:id="(\d+)"/g)) + 1;
  const ctx: Ctx = {
    doc: mDoc, commentMap: new Map(),
    nextTrackId: () => String(nextTrack++),
    mapComment: (oldId) => { let n = ctx.commentMap.get(oldId); if (n === undefined) { n = String(nextComment++); ctx.commentMap.set(oldId, n); } return n; },
  };
  const stats: MergeStats = { inserted: 0, deleted: 0, comments: 0, paragraphsAdded: 0, paragraphsRemoved: 0 };

  let s = 0;
  let lastMaster: Element | null = null, lastReturned: Element | null = null, lastInserted: Element | null = null;
  for (const rp of returnedParas) {
    const info = infoOfReturned(rp);
    const base = info.baseUnits.join("");
    const markDel = paraMark(rp, "w:del"), markIns = paraMark(rp, "w:ins");
    const isAligned = s < sentParas.length && sentText[s] === base;
    if (isAligned) {
      const mp = sentParas[s++];
      if (markDel && info.remainingBase > 0) unsupported("paragraf digabung/dipecah oleh pihak eksternal");
      const ops = extractOps(rp);
      if (ops.length) applyOps(ctx, mp, ops, stats);
      if (markDel) { ensureParaMark(ctx, mp, "w:del", attrOf(markDel, "author"), attrOf(markDel, "date")); stats.paragraphsRemoved++; }
      lastMaster = mp; lastReturned = rp; lastInserted = null;
    } else if (base === "" && (info.allIns || (!info.hasRuns && markIns))) {
      if (!lastMaster || !lastReturned || lastReturned.parentNode !== rp.parentNode) unsupported("paragraf baru di lokasi yang belum didukung");
      const clone = mDoc.importNode(rp, true) as Element;
      renumberTracks(ctx, clone); remapComments(ctx, clone);
      const firstIns = allOf(clone, "w:ins").find((e) => (e.parentNode as Element).nodeName !== "w:rPr");
      const author = markIns ? attrOf(markIns, "author") : firstIns ? attrOf(firstIns, "author") : "";
      const date = markIns ? attrOf(markIns, "date") : firstIns ? attrOf(firstIns, "date") : "";
      if (!author) unsupported("paragraf baru tanpa penulis");
      ensureParaMark(ctx, clone, "w:ins", author, date);
      const at = lastInserted || lastMaster;
      at.parentNode!.insertBefore(clone, at.nextSibling);
      lastInserted = clone; stats.paragraphsAdded++;
    } else {
      unsupported("teks berubah di luar Track Changes, atau paragraf dipecah/digabung");
    }
  }
  if (s !== sentParas.length) unsupported("paragraf dokumen tidak selaras dengan salinan yang dikirim");

  // Komentar eksternal → comments.xml master.
  if (ctx.commentMap.size) {
    if (!rcXml) unsupported("komentar eksternal tidak ditemukan di berkas");
    const rcDoc = new DOMParser().parseFromString(rcXml!, "text/xml");
    let mcDoc: Document;
    if (mcXmlOld) mcDoc = new DOMParser().parseFromString(mcXmlOld, "text/xml");
    else {
      mcDoc = new DOMParser().parseFromString(rcXml!, "text/xml");
      for (const c of allOf(mcDoc, "w:comment")) c.parentNode!.removeChild(c);
    }
    const root = mcDoc.documentElement;
    for (const c of allOf(rcDoc, "w:comment")) {
      const mapped = ctx.commentMap.get(attrOf(c, "id"));
      if (mapped === undefined) continue;
      const clone = mcDoc.importNode(c, true) as Element;
      clone.setAttribute("w:id", mapped);
      for (const p of allOf(clone, "w:p")) { p.removeAttribute("w14:paraId"); p.removeAttribute("w14:textId"); }
      root.appendChild(clone); stats.comments++;
    }
    mz.file("word/comments.xml", new XMLSerializer().serializeToString(mcDoc));
    if (!mcXmlOld) {
      const relsName = "word/_rels/document.xml.rels";
      let rels = await mz.file(relsName)?.async("string");
      if (!rels) unsupported("berkas relasi dokumen tidak ditemukan");
      const rid = "rIdClm" + Date.now();
      rels = rels!.replace("</Relationships>", `<Relationship Id="${rid}" Type="${REL_COMMENTS}" Target="comments.xml"/></Relationships>`);
      mz.file(relsName, rels);
      const ct = await mz.file("[Content_Types].xml")!.async("string");
      mz.file("[Content_Types].xml", ct.replace("</Types>", `<Override PartName="/word/comments.xml" ContentType="${CT_COMMENTS}"/></Types>`));
    }
  }

  mz.file("word/document.xml", new XMLSerializer().serializeToString(mDoc));
  const merged = await mz.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

  // Verifikasi — hasil gabungan harus konsisten dengan ketiga berkas sumbernya.
  if ((await plainAcceptedText(merged)) !== (await plainAcceptedText(returned))) unsupported("verifikasi gagal: teks akhir hasil gabungan berbeda dari berkas eksternal");
  if ((await originalTextOf(merged)) !== (await originalTextOf(master))) unsupported("verifikasi gagal: tampilan Original berubah");
  const countTracks = (x: string) => (x.match(/<w:(ins|del)\b/g) || []).length;
  if (countTracks(new XMLSerializer().serializeToString(mDoc)) < countTracks(mXml)) unsupported("verifikasi gagal: markup internal berkurang");
  return { ok: true, buf: merged, stats };
}
