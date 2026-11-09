import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

// Riwayat perubahan ala "Track Changes" Word. Dokumen hasil tinjauan (versi coretan) menyimpan
// w:ins / w:del / komentar lengkap dengan NAMA PENGUBAH dan WAKTU; fungsi di sini membaca riwayat itu
// dan membuat versi bersih (semua perubahan diterima, tanpa komentar).

export interface TrackedChange {
  id: string;
  kind: "insert" | "delete" | "format";
  author: string;
  date: string;
  text: string;
  /** Lokasi: judul pasal terdekat (atau awal paragraf). */
  where: string;
  /** Kutipan paragraf tempat perubahan terjadi. */
  context: string;
}
export interface TrackedComment {
  id: string;
  author: string;
  date: string;
  text: string;
  where: string;
}
export interface ChangeHistory { changes: TrackedChange[]; comments: TrackedComment[] }

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function attr(el: Element, name: string): string {
  return el.getAttribute(`w:${name}`) || el.getAttributeNS(W, name) || "";
}
function kids(el: Element, name: string): Element[] {
  return Array.from(el.childNodes).filter((n) => n.nodeType === 1 && (n as Element).nodeName === name) as Element[];
}
function paraText(p: Element, includeDeleted = true): string {
  let out = "";
  const walk = (n: Node, inDel: boolean) => {
    if (n.nodeType !== 1) return;
    const e = n as Element;
    if (e.nodeName === "w:t") out += e.textContent || "";
    else if (e.nodeName === "w:delText") { if (includeDeleted) out += e.textContent || ""; }
    else if (e.nodeName === "w:tab") out += "\t";
    else for (const c of Array.from(e.childNodes)) walk(c, inDel || e.nodeName === "w:del");
  };
  walk(p, false);
  return out;
}
function elText(el: Element): string {
  let out = "";
  const walk = (n: Node) => {
    if (n.nodeType !== 1) return;
    const e = n as Element;
    if (e.nodeName === "w:t" || e.nodeName === "w:delText") out += e.textContent || "";
    else if (e.nodeName === "w:tab") out += "\t";
    else for (const c of Array.from(e.childNodes)) walk(c);
  };
  walk(el);
  return out;
}
function clip(s: string, n: number): string {
  const t = s.replace(/\s+/g, " ").trim();
  return t.length > n ? t.slice(0, n - 1) + "…" : t;
}

/** Judul tempat sebuah elemen berada: paragraf pertama pasal (content control clause:*) atau awal paragrafnya sendiri. */
function whereOf(el: Element): string {
  let n: Node | null = el;
  let para: Element | null = null;
  while (n) {
    if (n.nodeType === 1) {
      const e = n as Element;
      if (e.nodeName === "w:p" && !para) para = e;
      if (e.nodeName === "w:sdt") {
        const pr = kids(e, "w:sdtPr")[0];
        const tag = pr ? kids(pr, "w:tag")[0] : undefined;
        if (tag && /^clause:/.test(attr(tag, "val"))) {
          const content = kids(e, "w:sdtContent")[0];
          const first = content ? (Array.from(content.getElementsByTagName("w:p"))[0] as Element | undefined) : undefined;
          if (first) return clip(paraText(first, false), 70);
        }
      }
    }
    n = n.parentNode;
  }
  return para ? clip(paraText(para, false), 70) : "";
}
function paragraphOf(el: Element): Element | null {
  let n: Node | null = el;
  while (n) { if (n.nodeType === 1 && (n as Element).nodeName === "w:p") return n as Element; n = n.parentNode; }
  return null;
}

export async function readChangeHistory(docx: Buffer): Promise<ChangeHistory> {
  const zip = await JSZip.loadAsync(docx);
  const docXml = await zip.file("word/document.xml")?.async("string");
  if (!docXml) return { changes: [], comments: [] };
  const doc = new DOMParser().parseFromString(docXml, "text/xml");
  const changes: TrackedChange[] = [];
  const mark = (tag: "w:ins" | "w:del" | "w:rPrChange" | "w:pPrChange", kind: TrackedChange["kind"]) => {
    for (const el of Array.from(doc.getElementsByTagName(tag)) as Element[]) {
      const parent = el.parentNode as Element | null;
      // penanda paragraf (di dalam w:rPr milik w:pPr) tidak punya teks: lewati, ikut paragrafnya
      if ((tag === "w:ins" || tag === "w:del") && parent && parent.nodeName === "w:rPr") continue;
      const text = kind === "format" ? "" : elText(el);
      if (kind !== "format" && !text.trim()) continue;
      const p = paragraphOf(el);
      changes.push({
        id: attr(el, "id"), kind, author: attr(el, "author") || "(tanpa nama)", date: attr(el, "date"),
        text: clip(text, 400), where: whereOf(el), context: p ? clip(paraText(p, true), 160) : "",
      });
    }
  };
  mark("w:ins", "insert");
  mark("w:del", "delete");
  mark("w:rPrChange", "format");
  mark("w:pPrChange", "format");
  // Gabungkan potongan bersebelahan (pengubah, jenis & paragraf sama) supaya daftar ringkas.
  const merged: TrackedChange[] = [];
  for (const c of changes) {
    const last = merged[merged.length - 1];
    if (last && last.kind === c.kind && last.author === c.author && last.context === c.context && last.kind !== "format" && last.date.slice(0, 16) === c.date.slice(0, 16)) {
      last.text = clip(last.text + " " + c.text, 400);
    } else merged.push({ ...c });
  }
  merged.sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  const comments: TrackedComment[] = [];
  const cXml = await zip.file("word/comments.xml")?.async("string");
  if (cXml) {
    const cdoc = new DOMParser().parseFromString(cXml, "text/xml");
    const refs = new Map<string, string>();
    for (const r of Array.from(doc.getElementsByTagName("w:commentReference")) as Element[]) refs.set(attr(r, "id"), whereOf(r));
    for (const c of Array.from(cdoc.getElementsByTagName("w:comment")) as Element[]) {
      const id = attr(c, "id");
      comments.push({ id, author: attr(c, "author") || "(tanpa nama)", date: attr(c, "date"), text: clip(elText(c), 600), where: refs.get(id) || "" });
    }
    comments.sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  }
  return { changes: merged, comments };
}

/** Versi bersih: terima semua perubahan (sisipan tetap, hapusan dibuang) dan buang komentar. */
export async function acceptAllChanges(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx);
  const names = Object.keys(zip.files).filter((n) => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n));
  for (const name of names) {
    const xml = await zip.file(name)!.async("string");
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const remove = (el: Node) => el.parentNode?.removeChild(el);
    const unwrap = (el: Element) => {
      const parent = el.parentNode!;
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
    };
    // paragraf yang tanda paragrafnya dihapus: bila isinya kosong setelah hapusan, buang paragrafnya
    for (const p of Array.from(doc.getElementsByTagName("w:p")) as Element[]) {
      const ppr = kids(p, "w:pPr")[0];
      const rpr = ppr ? kids(ppr, "w:rPr")[0] : undefined;
      const delMark = rpr ? kids(rpr, "w:del")[0] : undefined;
      if (delMark) {
        const rest = paraText(p, false).replace(/\s+/g, "");
        if (!rest) { remove(p); continue; }
      }
    }
    for (const e of Array.from(doc.getElementsByTagName("w:del")) as Element[]) {
      if (e.parentNode && (e.parentNode as Element).nodeName === "w:rPr") remove(e); else remove(e);
    }
    for (const e of Array.from(doc.getElementsByTagName("w:ins")) as Element[]) {
      const parent = e.parentNode as Element | null;
      if (parent && parent.nodeName === "w:rPr") remove(e); else unwrap(e);
    }
    for (const tag of ["w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange", "w:trPrChange", "w:tcPrChange", "w:numberingChange", "w:moveFromRangeStart", "w:moveFromRangeEnd", "w:moveToRangeStart", "w:moveToRangeEnd", "w:commentRangeStart", "w:commentRangeEnd"]) {
      for (const e of Array.from(doc.getElementsByTagName(tag)) as Element[]) remove(e);
    }
    for (const r of Array.from(doc.getElementsByTagName("w:commentReference")) as Element[]) {
      const run = r.parentNode as Element;
      if (run && run.nodeName === "w:r") remove(run); else remove(r);
    }
    zip.file(name, new XMLSerializer().serializeToString(doc));
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Ada perubahan teks yang belum diterima/ditolak? (komentar saja tidak dihitung) */
export async function hasTrackedChanges(docx: Buffer): Promise<boolean> {
  const zip = await JSZip.loadAsync(docx);
  const xml = await zip.file("word/document.xml")?.async("string");
  return !!xml && /<w:(ins|del|rPrChange|pPrChange)\b/.test(xml);
}

/** Teks tampilan "Original": semua perubahan ditolak (sisipan dibuang, hapusan kembali). Dipakai untuk menjaga peninjau tidak mengubah teks di luar Track Changes. */
export async function originalTextOf(docx: Buffer): Promise<string> {
  const zip = await JSZip.loadAsync(docx);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (!xml) return "";
  const doc = new DOMParser().parseFromString(xml, "text/xml");
  const out: string[] = [];
  for (const p of Array.from(doc.getElementsByTagName("w:p")) as Element[]) {
    // paragraf yang seluruhnya sisipan (tanda paragraf bertanda w:ins) tidak ada di dokumen asli
    const ppr = kids(p, "w:pPr")[0];
    const rpr = ppr ? kids(ppr, "w:rPr")[0] : undefined;
    const insMark = !!(rpr && kids(rpr, "w:ins")[0]);
    let t = "";
    const walk = (n: Node, inIns: boolean) => {
      if (n.nodeType !== 1) return;
      const e = n as Element;
      if (e.nodeName === "w:ins") return; // sisipan: tidak ada di asli
      if (e.nodeName === "w:t" || e.nodeName === "w:delText") t += e.textContent || "";
      else if (e.nodeName === "w:tab") t += "\t";
      else for (const c of Array.from(e.childNodes)) walk(c, inIns);
    };
    walk(p, false);
    if (insMark && !t) continue;
    out.push(t);
  }
  return out.join("").replace(/\s+/g, ""); // spasi/baris baru diabaikan: memecah/menggabung paragraf tidak dianggap mengubah teks
}
