import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { clauseContentToWordXml, numberingXml, type Ctx } from "./clause-html.js";

// ---------------------------------------------------------------------------
// PURWARUPA — template berbasis .docx dengan "content control" Word (w:sdt).
//
//  • Pasal      = content control tingkat blok, tag  "clause:<id>"
//      - wajib      -> w:lock = sdtLocked         (tidak bisa dihapus, isi boleh diedit)
//      - terlindungi-> w:lock = sdtContentLocked  (tidak bisa dihapus & isi tidak bisa diedit)
//  • Variabel   = content control sebaris, tag  "var:<Key>"  (isinya boleh diganti kapan saja,
//                 juga oleh sistem lewat setVariables() walau dokumen sudah banyak diedit)
//
// Kunci di dalam file hanyalah lapis pertama (editor menolak). Lapis penentu ada di
// server: checkTemplateEdit() membandingkan hasil edit dgn versi sebelumnya dan
// menolak perubahan pasal terlindungi / penghapusan pasal wajib oleh non-legal.
// Kebijakan (wajib/terlindungi) SELALU diambil dari database, bukan dari isi berkas.
// ---------------------------------------------------------------------------

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const CLAUSE_PREFIX = "clause:";
const VAR_PREFIX = "var:";

export interface TemplateClauseInput {
  id: string;
  title: string;
  /** Teks pasal; paragraf dipisah baris kosong; variabel ditulis {{Key}}. */
  content: string;
  isMandatory: boolean;
  isProtected: boolean;
  /** Padanan bahasa Inggris (dokumen bilingual). */
  titleEn?: string;
  contentEn?: string;
}

export interface ClausePolicy { isMandatory: boolean; isProtected: boolean }
export type PolicyMap = Record<string, ClausePolicy>;

export interface Violation {
  clauseId: string;
  kind: "protected-changed" | "protected-removed" | "mandatory-removed";
  message: string;
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

let sdtCounter = 1000;
const nextId = () => ++sdtCounter;

function run(text: string, bold = false): string {
  if (!text) return "";
  return `<w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}

function varSdt(key: string, value: string, rPr = ""): string {
  const r = value ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(value)}</w:t></w:r>` : "";
  return `<w:sdt><w:sdtPr><w:alias w:val="${esc(key)}"/><w:tag w:val="${VAR_PREFIX}${esc(key)}"/><w:id w:val="${nextId()}"/><w:text/></w:sdtPr><w:sdtContent>${r}</w:sdtContent></w:sdt>`;
}

const clauseInner = (title: string, content: string, ctx: Ctx) =>
  `<w:p><w:pPr><w:keepNext/></w:pPr>${run(title, true)}</w:p>${clauseContentToWordXml(content, ctx)}`;

/** Isi sel tabel harus berakhir dengan paragraf. */
const cellXml = (x: string) => (x && /<\/w:p>$/.test(x) ? x : x + "<w:p/>");

/** Dua bahasa: berdampingan (tabel 2 kolom) atau atas-bawah. */
function pairXml(idXml: string, enXml: string, layout: "side" | "stacked"): string {
  if (!enXml) return idXml;
  if (layout === "stacked") return idXml + enXml;
  const w = 4513;
  return `<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>${NO_BORDERS}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="${w}"/><w:gridCol w:w="${w}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${cellXml(idXml)}</w:tc><w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${cellXml(enXml)}</w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`;
}

function clauseSdt(c: TemplateClauseInput, ctx: Ctx, bi?: { layout: "side" | "stacked" }): string {
  const lock = c.isProtected ? "sdtContentLocked" : c.isMandatory ? "sdtLocked" : "";
  const idPart = clauseInner(c.title, c.content, ctx);
  const enPart = bi && (c.titleEn || c.contentEn) ? clauseInner(c.titleEn || c.title, c.contentEn || "", ctx) : "";
  const inner = bi ? pairXml(idPart, enPart, bi.layout) : idPart;
  return `<w:sdt><w:sdtPr><w:alias w:val="Pasal: ${esc(c.title)}"/><w:tag w:val="${CLAUSE_PREFIX}${esc(c.id)}"/><w:id w:val="${nextId()}"/>${lock ? `<w:lock w:val="${lock}"/>` : ""}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
}

const LAYOUT_PREFIX = "layout:";
let layoutId = 7000;
const layoutSdt = (kind: "kop" | "ttd", alias: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${LAYOUT_PREFIX}${kind}"/><w:id w:val="${layoutId++}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

// ---- Gambar (media) ---------------------------------------------------------
export interface MediaItem { rid: string; ext: "png" | "jpeg"; bytes: Buffer }
/** Kumpulan gambar yang dipakai satu potongan dokumen; id relasi berawalan `prefix`. */
export class Media {
  items: MediaItem[] = [];
  constructor(public prefix: string) {}
  add(bytes: Buffer, ext: "png" | "jpeg"): string {
    const rid = `${this.prefix}${this.items.length + 1}`;
    this.items.push({ rid, ext, bytes });
    return rid;
  }
}
const mediaPath = (m: MediaItem) => `word/media/clm-${m.rid}.${m.ext}`;

/** Gambar sebaris (EMU). */
function drawingRun(rid: string, cx: number, cy: number, name: string): string {
  const id = ++layoutId;
  return `<w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="${esc(name)}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
/** Gambar mengambang (di depan teks), diletakkan di tengah kolom dan digeser vertikal `offY` EMU dari paragrafnya. */
function anchorRun(rid: string, cx: number, cy: number, name: string, offY: number): string {
  const id = ++layoutId;
  return `<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${offY}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="${esc(name)}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${esc(name)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}
/** Skala (lebar,tinggi px) agar muat dalam kotak maksimum (EMU), tidak diperbesar. */
function fitEmu(w: number, h: number, maxW: number, maxH: number): { cx: number; cy: number } {
  let cx = Math.max(w, 1) * 9525, cy = Math.max(h, 1) * 9525;
  const k = Math.min(1, maxW / cx, maxH / cy);
  return { cx: Math.round(cx * k), cy: Math.round(cy * k) };
}

/** Tulis gambar ke zip + daftarkan relasi & tipe konten. Mengembalikan rels/types baru. */
function embedMedia(zip: JSZip, rels: string, types: string, items: MediaItem[]): { rels: string; types: string } {
  for (const m of items) {
    zip.file(mediaPath(m), m.bytes);
    rels = rels.replace(new RegExp(`<Relationship [^>]*Id="${m.rid}"[^>]*/>`, "g"), "");
    rels = rels.replace("</Relationships>", `<Relationship Id="${m.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/clm-${m.rid}.${m.ext}"/></Relationships>`);
    if (!new RegExp(`<Default [^>]*Extension="${m.ext}"`, "i").test(types)) types = types.replace(/(<Types[^>]*>)/, `$1<Default Extension="${m.ext}" ContentType="image/${m.ext}"/>`);
  }
  return { rels, types };
}
/** Buang gambar lama berawalan `prefix` (relasi + berkas) saat blok diganti. */
function dropMedia(zip: JSZip, rels: string, prefix: string): string {
  for (const m of rels.matchAll(new RegExp(`<Relationship [^>]*Id="(${prefix}\\d+)"[^>]*/>`, "g"))) {
    for (const ext of ["png", "jpeg"]) zip.remove(`word/media/clm-${m[1]}.${ext}`);
  }
  return rels.replace(new RegExp(`<Relationship [^>]*Id="${prefix}\\d+"[^>]*/>`, "g"), "");
}

export interface LetterheadInput {
  name: string;
  address?: string;
  /** Baris kontak (telp/fax/mobile) dan situs/email, sudah dirangkai. */
  contact?: string;
  web?: string;
  registration?: string;
  logo?: { bytes: Buffer; mimeType: string };
}
export interface SignatureInput { role: string; name: string }

/** Ukuran gambar PNG/JPEG (px) dibaca dari header berkas; null bila format lain. */
export function imageInfo(buf: Buffer, mime: string): { ext: "png" | "jpeg"; width: number; height: number } | null {
  try {
    if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) return { ext: "png", width: buf.readUInt32BE(16), height: buf.readUInt32BE(20) };
    if (buf[0] === 0xff && buf[1] === 0xd8) {
      let i = 2;
      while (i + 9 < buf.length) {
        if (buf[i] !== 0xff) { i++; continue; }
        const m = buf[i + 1];
        if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { ext: "jpeg", height: buf.readUInt16BE(i + 5), width: buf.readUInt16BE(i + 7) };
        i += 2 + buf.readUInt16BE(i + 2);
      }
    }
  } catch { /* format tak dikenali */ }
  void mime;
  return null;
}

const NO_BORDERS = `<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>`;
const cp = (text: string, o: { bold?: boolean; size?: number; jc?: string; after?: number; color?: string } = {}) =>
  `<w:p><w:pPr><w:spacing w:after="${o.after ?? 0}" w:line="240" w:lineRule="auto"/><w:jc w:val="${o.jc || "center"}"/></w:pPr><w:r><w:rPr>${o.bold ? "<w:b/>" : ""}${o.color ? `<w:color w:val="${o.color}"/>` : ""}${o.size ? `<w:sz w:val="${o.size}"/>` : ""}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`;

/** Kop surat: blok di awal dokumen (halaman 1) — logo kiri, identitas perusahaan di tengah, garis bawah. */
export function kopBlockXml(l: LetterheadInput, logo: { ext: "png" | "jpeg"; width: number; height: number } | null, media: Media): string {
  const lines = [
    cp(l.name.toUpperCase(), { bold: true, size: 32 }),
    ...(l.address ? [cp(l.address, { size: 20, color: "444444" })] : []),
    ...(l.contact ? [cp(l.contact, { size: 20, color: "444444" })] : []),
    ...(l.web ? [cp(l.web, { size: 20, color: "444444" })] : []),
    ...(l.registration ? [cp(l.registration, { size: 20, color: "666666" })] : []),
  ].join("");
  let cell1 = "";
  let cols = `<w:gridCol w:w="9026"/>`;
  if (logo && l.logo) {
    const rid = media.add(l.logo.bytes, logo.ext);
    const k = Math.min(1800000 / Math.max(logo.width, 1), 900000 / Math.max(logo.height, 1));
    const cx = Math.round(logo.width * k), cy = Math.round(logo.height * k);
    cell1 = `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>${drawingRun(rid, cx, cy, "Logo")}</w:p></w:tc>`;
    cols = `<w:gridCol w:w="1700"/><w:gridCol w:w="7326"/>`;
  }
  const cell2 = `<w:tc><w:tcPr><w:tcW w:w="${cell1 ? 7326 : 9026}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>${lines}</w:tc>`;
  const rule = `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="475569"/></w:pBdr><w:spacing w:before="60" w:after="240"/></w:pPr></w:p>`;
  return layoutSdt("kop", "Kop surat", `<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>${NO_BORDERS}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${cols}</w:tblGrid><w:tr>${cell1}${cell2}</w:tr></w:tbl>${rule}`);
}

export interface SignatureOptions { copies: 1 | 2; meterai: boolean }
/** Gambar tanda tangan digital Pihak Pertama. */
export interface DigitalSignatureInput { bytes: Buffer }

/** Satu tabel tanda tangan: peran, ruang tanda tangan (+ kotak meterai di pihak `meteraiIdx`, + gambar TTD digital di pihak pertama), ( nama ). */
function sigTable(sigs: SignatureInput[], meteraiIdx: number, digital: { rid: string; cx: number; cy: number } | null): string {
  const n = Math.max(sigs.length, 1);
  const w = Math.floor(9026 / n);
  const row = (inner: (s: SignatureInput, i: number) => string, h?: number) =>
    `<w:tr><w:trPr><w:cantSplit/>${h ? `<w:trHeight w:val="${h}" w:hRule="atLeast"/>` : ""}</w:trPr>${sigs.map((s, i) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>${inner(s, i)}</w:tc>`).join("")}</w:tr>`;
  const boxP = (extraRun: string) => `<w:p><w:pPr><w:keepNext/><w:pBdr><w:top w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/><w:left w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/><w:bottom w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/><w:right w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/></w:pBdr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="${Math.max(Math.floor(w / 2) - 700, 0)}" w:right="${Math.max(Math.floor(w / 2) - 700, 0)}"/><w:jc w:val="center"/></w:pPr>${extraRun}<w:r><w:rPr><w:sz w:val="16"/><w:color w:val="64748B"/></w:rPr><w:t>METERAI</w:t></w:r><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="64748B"/></w:rPr><w:br/><w:t>Rp10.000</w:t></w:r></w:p>`;
  // Tanda tangan digital: bila kolom ini juga bermeterai, TTD menimpa kotak meterai (seperti tanda tangan basah);
  // bila tidak, TTD berdiri sendiri di tengah ruang tanda tangan.
  const cell = (i: number) => {
    const sig = digital && i === 0 ? digital : null;
    if (i === meteraiIdx) return boxP(sig ? anchorRun(sig.rid, sig.cx, sig.cy, "Tanda tangan digital", -Math.round(sig.cy * 0.35)) : "");
    if (sig) return `<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr>${drawingRun(sig.rid, sig.cx, sig.cy, "Tanda tangan digital")}</w:p>`;
    return "";
  };
  const empty = `<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr></w:p>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>${NO_BORDERS}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${sigs.map(() => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>` +
    row((s) => `<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${esc(s.role.toUpperCase())}</w:t></w:r></w:p>`) +
    row((_s, i) => cell(i) || empty, 1300) +
    row((s) => `<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">( ${esc(s.name)} )</w:t></w:r></w:p>`) +
    `</w:tbl>`;
}

const capt = (t: string) => `<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="20"/><w:color w:val="475569"/></w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`;

/**
 * Blok tanda tangan di akhir dokumen. Tunggal: satu blok, meterai di Pihak Pertama.
 * Dua rangkap: dua blok (tiap rangkap di halaman sendiri) — rangkap 1 dipegang Pihak Pertama
 * (meterai di kolomnya), rangkap 2 dipegang Pihak Kedua. Tanpa meterai: kotak meterai tidak dibuat.
 * TTD digital (bila ada) dipasang di kolom Pihak Pertama pada tiap rangkap.
 */
// Pemisah rangkap: paragraf kecil dengan "mulai di halaman baru" (pageBreakBefore). Dulu memakai <w:br type=page/>
// di paragraf sendiri; bila blok tanda tangan rangkap 1 tepat mengisi halaman, paragraf itu terdorong ke halaman
// berikut lalu pindah halaman lagi -> muncul halaman KOSONG di tengah. pageBreakBefore tidak memakan ruang.
const PAGE_BREAK_P = `<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>`;
const isRangkapBreak = (x: Node): boolean =>
  x.nodeName === "w:p" && (
    Array.from((x as Element).getElementsByTagName("w:pageBreakBefore")).length > 0 ||
    Array.from((x as Element).getElementsByTagName("w:br")).some((b) => (b as Element).getAttribute("w:type") === "page"));
export function signatureBlockXml(sigs: SignatureInput[], o: SignatureOptions, media: Media, digital?: DigitalSignatureInput): string {
  const holder = (i: number) => sigs[i]?.role || (i === 0 ? "Pihak Pertama" : "Pihak Kedua");
  let dg: { rid: string; cx: number; cy: number } | null = null;
  if (digital) {
    const info = imageInfo(digital.bytes, "");
    if (info) {
      const colW = Math.floor(9026 / Math.max(sigs.length, 1)) * 635; // twips -> EMU
      const { cx, cy } = fitEmu(info.width, info.height, colW * 0.8, 760000);
      dg = { rid: media.add(digital.bytes, info.ext), cx, cy };
    }
  }
  const t = (idx: number) => sigTable(sigs, idx, dg);
  const sdtInner = (() => {
    if (o.copies === 2 && sigs.length > 1) {
      return capt(`Rangkap 1 dari 2 — dipegang ${holder(0)}${o.meterai ? " (bermeterai)" : ""}`) + t(o.meterai ? 0 : -1) +
        PAGE_BREAK_P +
        capt(`Rangkap 2 dari 2 — dipegang ${holder(1)}${o.meterai ? " (bermeterai)" : ""}`) + t(o.meterai ? 1 : -1);
    }
    return capt(`Rangkap tunggal${o.meterai ? " — bermeterai pada " + holder(0) : ""}`) + t(o.meterai ? 0 : -1);
  })();
  return layoutSdt("ttd", "Tanda tangan", sdtInner + `<w:p/>`);
}

export interface LayoutPatch {
  /** undefined = biarkan; null = hapus blok; objek = pasang/ganti blok. */
  kop?: { letterhead: LetterheadInput } | null;
  ttd?: { signatures: SignatureInput[]; options: SignatureOptions; digital?: DigitalSignatureInput } | null;
}


// ---- Rangkap 2: ulangi tulisan yang satu halaman dengan blok tanda tangan -------------------------------
// Dokumen fisik dua rangkap: halaman tanda tangan tiap rangkap memuat tulisan yang sama dengan yang tampil
// di halaman tanda tangan rangkap pertama (mis. ayat (2) dan (3) dari pasal terakhir), supaya halaman
// tanda tangan tidak berdiri sendiri. Salinan dibuat PER PARAGRAF (bukan per pasal). Word tidak memberi
// tahu pemisah halaman, jadi posisinya DIPERKIRAKAN dari panjang teks; hasilnya salinan biasa yang boleh
// diubah/dihapus. Daftar bernomor pada salinan diberi nomor lanjutan yang benar (mis. mulai dari "2.").
const LINES_PER_PAGE = 44;
const SIG_LINES = 11;
const CHARS_PER_LINE = 86;
const SIDE_CAPTION = "clm_side";

const isHiddenPara = (p: Element) => {
  const r = Array.from(p.getElementsByTagName("w:r"))[0] as Element | undefined;
  return !!r && r.getElementsByTagName("w:vanish").length > 0;
};

function rowLines(tr: Element): number {
  const cells = Array.from(tr.childNodes).filter((c) => c.nodeName === "w:tc");
  const k = Math.max(cells.length, 1);
  return Math.max(1, ...cells.map((c) => Array.from(c.childNodes).reduce((a, x) => a + (x.nodeName === "w:p" && !isHiddenPara(x as Element) ? Math.max(1, Math.ceil((Array.from((x as Element).getElementsByTagName("w:t")).map((t) => t.textContent || "").join("").length) / (CHARS_PER_LINE / k))) + 0.4 : 0), 0)));
}

function estLines(n: Node): number {
  const el = n as Element;
  const name = el.nodeName;
  if (name === "w:p") {
    if (isHiddenPara(el)) return 0;
    const text = Array.from(el.getElementsByTagName("w:t")).map((x) => x.textContent || "").join("");
    let lines = Math.max(1, Math.ceil(text.length / CHARS_PER_LINE)) + Array.from(el.getElementsByTagName("w:br")).length;
    for (const ext of Array.from(el.getElementsByTagName("wp:extent")) as Element[]) lines += Math.ceil(Number(ext.getAttribute("cy") || 0) / 190000);
    return lines + 0.4;
  }
  if (name === "w:tbl") return Array.from(el.childNodes).filter((c) => c.nodeName === "w:tr").reduce((a, tr) => a + rowLines(tr as Element), 0) + 1;
  if (name === "w:sdt") {
    const c = sdtContent(el);
    return c ? Array.from(c.childNodes).reduce((a, x) => a + estLines(x), 0) : 0;
  }
  return 0;
}

const isLayoutSdt = (n: Node) => n.nodeName === "w:sdt" && sdtTag(n as Element).startsWith(LAYOUT_PREFIX);
const hasContent = (n: Node) => n.nodeName !== "w:p" || (n as Element).getElementsByTagName("w:drawing").length > 0 || Array.from((n as Element).getElementsByTagName("w:t")).some((t) => (t.textContent || "").trim());
const isSideTable = (n: Node) => {
  if (n.nodeName !== "w:tbl") return false;
  const pr = Array.from(n.childNodes).find((c) => c.nodeName === "w:tblPr") as Element | undefined;
  const cap = pr && (Array.from(pr.childNodes).find((c) => c.nodeName === "w:tblCaption") as Element | undefined);
  return cap?.getAttribute("w:val") === SIDE_CAPTION;
};

interface Leaf { el?: Element; lines: number; tbl?: Element; boundary?: boolean }
/** Pecah isi dokumen jadi satuan terkecil: paragraf, baris tabel dwibahasa, atau tabel; blok tata letak = batas. */
function leavesOf(n: Node, out: Leaf[]): void {
  if (n.nodeName === "w:p") out.push({ el: n as Element, lines: estLines(n) });
  else if (n.nodeName === "w:tbl") {
    if (isSideTable(n)) for (const tr of Array.from(n.childNodes).filter((c) => c.nodeName === "w:tr")) out.push({ el: tr as Element, tbl: n as Element, lines: rowLines(tr as Element) });
    else out.push({ el: n as Element, lines: estLines(n) });
  } else if (n.nodeName === "w:sdt") {
    if (isLayoutSdt(n)) out.push({ lines: 0, boundary: true });
    else { const c = sdtContent(n as Element); if (c) for (const k of Array.from(c.childNodes)) leavesOf(k, out); }
  }
}

const TAIL_NUM_BASE = 9000;   // num salinan rangkap 2: 9000..19999 (dibersihkan bila tak terpakai)
const LANG_NUM_BASE = 20000;  // num daftar paragraf Inggris: >= 20000

/** Penyunting numbering.xml: membuat daftar baru yang memakai bentuk yang sama tapi nomornya dimulai sendiri. */
class NumCtx {
  changed = false;
  private cache = new Map<string, string>();
  private constructor(private nd: Document) {}
  static async load(zip: JSZip): Promise<NumCtx | null> {
    const e = zip.file("word/numbering.xml");
    return e ? new NumCtx(new DOMParser().parseFromString(await e.async("string"), "application/xml") as unknown as Document) : null;
  }
  private nums() { return Array.from(this.nd.getElementsByTagName("w:num")) as Element[]; }
  private nextId(lang: boolean): number {
    const lo = lang ? LANG_NUM_BASE : TAIL_NUM_BASE, hi = lang ? Infinity : LANG_NUM_BASE;
    return Math.max(lo - 1, ...this.nums().map((n) => Number(n.getAttribute("w:numId") || 0)).filter((i) => i >= lo && i < hi)) + 1;
  }
  /** Daftar baru dari `numId`; butir pertama bernomor (nomor awal + extra). Null bila daftar asal tak ditemukan. */
  derive(numId: string, lang: boolean, extra = 0, ilvl = 0): string | null {
    const key = `${lang}:${numId}:${ilvl}:${extra}`;
    const hit = this.cache.get(key);
    if (hit) return hit;
    const src = this.nums().find((n) => n.getAttribute("w:numId") === numId);
    const abs = src && attrNum(src, "w:abstractNumId");
    if (!abs) return null;
    const absEl = (Array.from(this.nd.getElementsByTagName("w:abstractNum")) as Element[]).find((x) => x.getAttribute("w:abstractNumId") === abs);
    if (!absEl) return null;
    // daftar baru = salinan penuh bentuk lama (abstractNum sendiri) supaya nomornya benar-benar terpisah di Word/ONLYOFFICE/LibreOffice
    const copy = absEl.cloneNode(true) as Element;
    const absIds = (Array.from(this.nd.getElementsByTagName("w:abstractNum")) as Element[]).map((x) => Number(x.getAttribute("w:abstractNumId") || 0));
    const newAbs = String(Math.max(...absIds, 0) + 1);
    copy.setAttribute("w:abstractNumId", newAbs);
    for (const x of Array.from(copy.childNodes)) if (x.nodeName === "w:nsid" || x.nodeName === "w:tmpl") copy.removeChild(x);
    const lvl = (Array.from(copy.getElementsByTagName("w:lvl")) as Element[]).find((l) => l.getAttribute("w:ilvl") === String(ilvl));
    const startEl = lvl && (Array.from(lvl.childNodes).find((x) => x.nodeName === "w:start") as Element | undefined);
    if (startEl) startEl.setAttribute("w:val", String(Number(startEl.getAttribute("w:val") || 1) + extra));
    const lastAbs = (Array.from(this.nd.getElementsByTagName("w:abstractNum")) as Element[]).pop()!;
    lastAbs.parentNode!.insertBefore(copy, lastAbs.nextSibling);
    const id = String(this.nextId(lang));
    const num = this.nd.createElementNS(W_NS, "w:num"); num.setAttribute("w:numId", id);
    const a2 = this.nd.createElementNS(W_NS, "w:abstractNumId"); a2.setAttribute("w:val", newAbs); num.appendChild(a2);
    const root = this.nd.documentElement;
    root.insertBefore(num, Array.from(root.childNodes).find((c) => c.nodeName === "w:numIdMacAtCleanup") || null);
    this.cache.set(key, id); this.changed = true;
    return id;
  }
  /** Hapus daftar turunan (rentang tertentu) yang tak dipakai paragraf mana pun di dokumen. */
  prune(doc: Document, lang: boolean): void {
    const used = new Set((Array.from(doc.getElementsByTagName("w:numId")) as Element[]).map((e) => e.getAttribute("w:val")));
    const lo = lang ? LANG_NUM_BASE : TAIL_NUM_BASE, hi = lang ? Infinity : LANG_NUM_BASE;
    for (const n of this.nums()) {
      const id = Number(n.getAttribute("w:numId") || 0);
      if (id >= lo && id < hi && !used.has(String(id))) {
        const abs = attrNum(n, "w:abstractNumId");
        n.parentNode?.removeChild(n); this.changed = true;
        if (abs && !this.nums().some((x) => attrNum(x, "w:abstractNumId") === abs)) {
          const ae = (Array.from(this.nd.getElementsByTagName("w:abstractNum")) as Element[]).find((x) => x.getAttribute("w:abstractNumId") === abs);
          ae?.parentNode?.removeChild(ae);
        }
      }
    }
  }
  save(zip: JSZip): void { if (this.changed) zip.file("word/numbering.xml", new XMLSerializer().serializeToString(this.nd as any)); }
}

const attrNum = (e: Element | null | undefined, tag: string): string | null => {
  const x = e && (Array.from(e.getElementsByTagName(tag))[0] as Element | undefined);
  return x ? x.getAttribute("w:val") : null;
};
const numPrOf = (p: Element): { numId: string; ilvl: number } | null => {
  const np = p.getElementsByTagName("w:numPr")[0] as Element | undefined;
  const id = attrNum(np, "w:numId");
  return np && id && id !== "0" ? { numId: id, ilvl: Number(attrNum(np, "w:ilvl") || 0) } : null;
};

/**
 * Daftar bernomor pada salinan harus melanjutkan nomor aslinya (mis. salinan ayat (2)(3) tetap bernomor 2 dan 3,
 * bukan lanjut ke 4 dan 5): tiap daftar disalin ke num baru yang mulainya diatur (startOverride).
 */
async function renumberClones(zip: JSZip, doc: Document, pairs: { orig: Element; clone: Element }[]): Promise<void> {
  const ctx = await NumCtx.load(zip);
  if (!ctx) return;
  const all = Array.from(doc.getElementsByTagName("w:p")) as Element[];
  const cloned = new Set(pairs.map((x) => x.clone));
  const mapped = new Map<string, string>();
  for (const { orig, clone } of pairs) {
    const np = numPrOf(clone);
    if (!np) continue;
    let target = mapped.get(np.numId);
    if (!target) {
      // jumlah butir sejenis (daftar & tingkat sama) yang muncul sebelum paragraf asli ini
      const idx = all.indexOf(orig);
      let prior = 0;
      for (let i = 0; i < idx; i++) {
        if (cloned.has(all[i])) continue;
        const q = numPrOf(all[i]);
        if (q && q.numId === np.numId && q.ilvl === np.ilvl) prior++;
      }
      const d = ctx.derive(np.numId, false, prior, np.ilvl);
      if (!d) continue;
      target = d; mapped.set(np.numId, d);
    }
    const idEl = clone.getElementsByTagName("w:numId")[0] as Element | undefined;
    if (idEl) idEl.setAttribute("w:val", target);
  }
  ctx.prune(doc, false);
  ctx.save(zip);
}

/** Pasang/ganti salinan "tulisan di halaman tanda tangan" pada rangkap 2 (bila dokumen memuat dua rangkap). */
const normText = (t: string) => t.toLowerCase().replace(/[\s\u00a0]+/g, "");
/** Teks halaman (hasil render sungguhan) yang memuat "Rangkap 1 dari 2" = halaman tanda tangan rangkap pertama. */
export function signaturePageText(pages: string[]): string | null {
  const hit = pages.find((t) => /rangkap\s*1\s*dari\s*2/i.test(t));
  // Teks yang memuat KEDUA rangkap = pembacaan halaman gagal (seluruh dokumen terbaca sebagai satu halaman) -> pakai perkiraan.
  if (hit && /rangkap\s*2\s*dari\s*2/i.test(hit)) return null;
  return hit ?? null;
}
const onPage = (page: string, lf: Leaf): boolean => {
  const texts = lf.tbl
    ? (Array.from(lf.el!.getElementsByTagName("w:p")) as Element[]).map((p) => paraText(p))
    : [Array.from(lf.el!.getElementsByTagName("w:t")).map((x) => x.textContent || "").join("")];
  return texts.every((t) => {
    const n = normText(t);
    if (!n) return true;
    const head = n.slice(0, 24);
    return page.includes(head) && (lf.tbl || n.length <= 24 || page.includes(n.slice(-24)));
  });
};

/** Apakah dokumen memuat blok tanda tangan dua rangkap? */
export async function hasTwoCopies(docx: Buffer): Promise<boolean> {
  const { doc } = await loadBody(docx);
  const body = doc.getElementsByTagName("w:body")[0] as Element;
  const ttd = sdtOf(doc).find((x) => sdtTag(x) === LAYOUT_PREFIX + "ttd" && x.parentNode === body);
  const c = ttd && sdtContent(ttd);
  return !!c && Array.from(c.childNodes).some(isRangkapBreak);
}

async function applyRangkapTail(zip: JSZip, doc: Document, pageText?: string): Promise<void> {
  const body = doc.getElementsByTagName("w:body")[0] as Element;
  const ttd = sdtOf(doc).find((x) => sdtTag(x) === LAYOUT_PREFIX + "ttd" && x.parentNode === body);
  if (!ttd) return;
  const content = sdtContent(ttd);
  if (!content) return;
  for (const x of Array.from(content.childNodes)) if (isLayoutSdt(x) && sdtTag(x as Element) === LAYOUT_PREFIX + "tail") content.removeChild(x);
  let pageBreak = Array.from(content.childNodes).find(isRangkapBreak) as Element | undefined;
  if (pageBreak && pageBreak.getElementsByTagName("w:pageBreakBefore").length === 0) {
    // dokumen lama: ganti paragraf <w:br type=page/> dengan pemisah yang tidak bisa menyisakan halaman kosong
    const fresh = new DOMParser().parseFromString(`<root xmlns:w="${W_NS}">${PAGE_BREAK_P}</root>`, "text/xml").documentElement.firstChild as Element;
    const imported = doc.importNode(fresh, true) as Element;
    content.replaceChild(imported, pageBreak);
    pageBreak = imported;
  }
  if (!pageBreak) { await renumberClones(zip, doc, []); return; } // satu rangkap
  const before = Array.from(body.childNodes).filter((n) => n !== ttd && n.nodeName !== "w:sectPr");
  const total = before.reduce((a, n) => a + estLines(n), 0);
  const lastPage = total % LINES_PER_PAGE;
  const budget = lastPage + SIG_LINES <= LINES_PER_PAGE ? lastPage : 0;
  const leaves: Leaf[] = [];
  for (const n of before) leavesOf(n, leaves);
  const picked: Leaf[] = [];
  let used = 0;
  const page = pageText !== undefined ? normText(pageText) : null;
  for (let i = leaves.length - 1; i >= 0; i--) {
    const lf = leaves[i];
    if (lf.boundary) break;
    if (lf.el!.nodeName === "w:p" && !hasContent(lf.el!)) { if (!picked.length) continue; }
    if (page !== null) { // hasil render sungguhan; tetap dibatasi satu halaman supaya salah baca tidak menyalin seluruh dokumen
      if (!onPage(page, lf) || used + lf.lines > LINES_PER_PAGE - SIG_LINES) break;
      used += lf.lines; picked.unshift(lf); continue;
    }
    if (used + lf.lines > budget) break;
    used += lf.lines; picked.unshift(lf);
  }
  if (!picked.length) { await renumberClones(zip, doc, []); return; }
  const sdt = doc.createElementNS(W_NS, "w:sdt");
  const pr = doc.createElementNS(W_NS, "w:sdtPr");
  const alias = doc.createElementNS(W_NS, "w:alias"); alias.setAttribute("w:val", "Tulisan di halaman tanda tangan (rangkap 2)");
  const tag = doc.createElementNS(W_NS, "w:tag"); tag.setAttribute("w:val", LAYOUT_PREFIX + "tail");
  const id = doc.createElementNS(W_NS, "w:id"); id.setAttribute("w:val", String(layoutId++));
  pr.appendChild(alias); pr.appendChild(tag); pr.appendChild(id);
  const sc = doc.createElementNS(W_NS, "w:sdtContent");
  const pairs: { orig: Element; clone: Element }[] = [];
  const adoptClone = (k: Element): Element => {
    const c = k.cloneNode(true) as Element;
    // bookmark (penanda terjemahan) pada salinan diberi nama/id unik agar tidak ganda
    for (const e of [c, ...(Array.from(c.getElementsByTagName("*")) as Element[])]) {
      if (e.nodeName === "w:bookmarkStart") { e.setAttribute("w:id", String(Number(e.getAttribute("w:id") || 0) + 5000)); e.setAttribute("w:name", (e.getAttribute("w:name") || "") + "_t"); }
      if (e.nodeName === "w:bookmarkEnd") e.setAttribute("w:id", String(Number(e.getAttribute("w:id") || 0) + 5000));
    }
    return c;
  };
  let openTbl: { src: Element; copy: Element } | null = null;
  for (const lf of picked) {
    if (lf.tbl) {
      if (!openTbl || openTbl.src !== lf.tbl) {
        const copy = lf.tbl.cloneNode(false) as Element;
        for (const k of Array.from(lf.tbl.childNodes)) if (k.nodeName === "w:tblPr" || k.nodeName === "w:tblGrid") copy.appendChild(k.cloneNode(true));
        sc.appendChild(copy); openTbl = { src: lf.tbl, copy };
      }
      const rowCopy = adoptClone(lf.el!);
      openTbl.copy.appendChild(rowCopy);
      const origPs = Array.from(lf.el!.getElementsByTagName("w:p")) as Element[];
      (Array.from(rowCopy.getElementsByTagName("w:p")) as Element[]).forEach((cp, i) => pairs.push({ orig: origPs[i], clone: cp }));
    } else {
      openTbl = null;
      const cp = adoptClone(lf.el!);
      sc.appendChild(cp);
      if (cp.nodeName === "w:p") pairs.push({ orig: lf.el!, clone: cp });
      else { const o = Array.from(lf.el!.getElementsByTagName("w:p")) as Element[]; (Array.from(cp.getElementsByTagName("w:p")) as Element[]).forEach((c2, i) => pairs.push({ orig: o[i], clone: c2 })); }
    }
  }
  sdt.appendChild(pr); sdt.appendChild(sc);
  content.insertBefore(sdt, pageBreak.nextSibling);
  await renumberClones(zip, doc, pairs);
}

/** Hitung ulang salinan rangkap 2 (mis. setelah isi/kop berubah). */
export async function refreshRangkapTail(docx: Buffer, pageText?: string): Promise<Buffer> {
  const { zip, doc } = await loadBody(docx);
  await applyRangkapTail(zip, doc, pageText);
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}


// ---- Bahasa dokumen: terjemahan langsung di dalam dokumen Word -------------------------------------------
// Tiap paragraf Indonesia berpasangan dengan paragraf Inggris (penanda bookmark clm_id_N / clm_en_N).
// ONLYOFFICE tidak mendukung teks "tersembunyi", jadi bahasa yang tidak dipilih BENAR-BENAR dikeluarkan dari
// dokumen dan disimpan di server (LangStash), lalu dipasang lagi saat bahasa dipilih. Tidak ada teks yang hilang.
const ID_MARK = "clm_id_";
const EN_MARK = "clm_en_";
export type DocLanguage = "id" | "bilingual" | "en";
export type BilingualLayoutMode = "side" | "stacked";
/** Nomor pasangan -> XML paragraf bahasa yang sedang tidak tampil di dokumen. */
export type LangStash = Record<string, { id?: string; en?: string }>;

const markNo = (p: Element, prefix: string): string | null => {
  for (const b of Array.from(p.getElementsByTagName("w:bookmarkStart")) as Element[]) {
    const n = b.getAttribute("w:name") || "";
    if (n.startsWith(prefix) && /^\d+$/.test(n.slice(prefix.length))) return n.slice(prefix.length);
  }
  return null;
};
const isEnPara = (p: Element) => markNo(p, EN_MARK) !== null;
const isIdPara = (p: Element) => markNo(p, ID_MARK) !== null;
const paraText = (p: Element) => Array.from(p.getElementsByTagName("w:t")).map((x) => x.textContent || "").join("");
function layoutAncestor(n: Node): string[] {
  const out: string[] = [];
  for (let x: Node | null = n.parentNode; x; x = x.parentNode) if (x.nodeName === "w:sdt") { const t = sdtTag(x as Element); if (t.startsWith(LAYOUT_PREFIX)) out.push(t.slice(LAYOUT_PREFIX.length)); }
  return out;
}
const inScope = (p: Element) => { const a = layoutAncestor(p); return !a.includes("kop") && !a.includes("ttd"); };
const scopeParas = (doc: Document) => (Array.from(doc.getElementsByTagName("w:p")) as Element[]).filter(inScope);

function addMark(doc: Document, p: Element, prefix: string, n: string): void {
  const id = String((prefix === ID_MARK ? 100000 : 200000) + Number(n));
  const bs = doc.createElementNS(W_NS, "w:bookmarkStart"); bs.setAttribute("w:id", id); bs.setAttribute("w:name", prefix + n);
  const be = doc.createElementNS(W_NS, "w:bookmarkEnd"); be.setAttribute("w:id", id);
  const pPr = Array.from(p.childNodes).find((x) => x.nodeName === "w:pPr");
  p.insertBefore(bs, pPr ? pPr.nextSibling : p.firstChild);
  p.appendChild(be);
}

// ---- tanggal Indonesia -> Inggris untuk variabel di paragraf Inggris ---------------------------------------
const MONTH_ID = ["januari", "februari", "maret", "april", "mei", "juni", "juli", "agustus", "september", "oktober", "november", "desember"];
const MONTH_EN = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAY_ID: Record<string, string> = { senin: "Monday", selasa: "Tuesday", rabu: "Wednesday", kamis: "Thursday", jumat: "Friday", "jum'at": "Friday", sabtu: "Saturday", minggu: "Sunday" };
/** "Kamis, 5 November 2026" -> "Thursday, November 5, 2026"; nilai lain dibiarkan. */
export function valueToEnglish(v: string): string {
  const m = v.trim().match(/^(?:([A-Za-z']+),\s*)?(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (!m) return v;
  const mi = MONTH_ID.indexOf(m[3].toLowerCase());
  if (mi < 0) return v;
  const date = `${MONTH_EN[mi]} ${Number(m[2])}, ${m[4]}`;
  const day = m[1] ? DAY_ID[m[1].toLowerCase()] : "";
  return day ? `${day}, ${date}` : date;
}

// ---- Pengodean paragraf: teks + format + variabel -> string untuk AI, dan kembali ---------------------------
interface Enc { text: string; fmts: Element[]; base: Element | null; vars: string[] }
const stripRPr = (rp: Element | null | undefined): Element | null => {
  if (!rp) return null;
  const c = rp.cloneNode(true) as Element;
  for (const x of Array.from(c.childNodes)) if (["w:vanish", "w:lang", "w:rStyle"].includes(x.nodeName)) c.removeChild(x);
  return c;
};
const rprOf = (r: Element) => stripRPr(Array.from(r.childNodes).find((x) => x.nodeName === "w:rPr") as Element | undefined);
const runText = (r: Element): string => {
  let s = "";
  for (const k of Array.from(r.childNodes)) {
    if (k.nodeName === "w:t") s += k.textContent || "";
    else if (k.nodeName === "w:tab") s += "\t";
    else if (k.nodeName === "w:br" && (k as Element).getAttribute("w:type") !== "page") s += "\n";
  }
  return s;
};

function encodePara(p: Element): Enc {
  type Seg = { text: string; key: string; rpr: Element | null; v?: string };
  const segs: Seg[] = [];
  const ser = (e: Element | null) => (e ? new XMLSerializer().serializeToString(e as any) : "");
  const addRun = (r: Element) => { const t = runText(r); if (t) { const rp = rprOf(r); segs.push({ text: t, key: ser(rp), rpr: rp }); } };
  const walk = (parent: Element) => {
    for (const k of Array.from(parent.childNodes)) {
      if (k.nodeName === "w:r") addRun(k as Element);
      else if (k.nodeName === "w:hyperlink" || k.nodeName === "w:smartTag") walk(k as Element);
      else if (k.nodeName === "w:sdt") {
        const tag = sdtTag(k as Element);
        if (tag.startsWith(VAR_PREFIX)) {
          const c = sdtContent(k as Element);
          const r0 = c && (Array.from(c.getElementsByTagName("w:r"))[0] as Element | undefined);
          const rp = r0 ? rprOf(r0) : null;
          segs.push({ text: "", key: ser(rp), rpr: rp, v: tag.slice(VAR_PREFIX.length).replace(/@en$/, "") });
        } else { const c = sdtContent(k as Element); if (c) walk(c); }
      }
    }
  };
  walk(p);
  // format dasar = yang teksnya paling banyak; format lain diberi kode f1, f2, ...
  const weight = new Map<string, number>();
  for (const s of segs) if (!s.v) weight.set(s.key, (weight.get(s.key) || 0) + s.text.length);
  const baseKey = [...weight.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "";
  const baseSeg = segs.find((s) => s.key === baseKey);
  const fmts: Element[] = [], codeOf = new Map<string, number>();
  let out = "", open = 0;
  const vars: string[] = [];
  for (const s of segs) {
    if (s.v) { out += `{{${s.v}}}`; vars.push(s.v); continue; }
    if (s.key === baseKey) { out += s.text; continue; }
    let c = codeOf.get(s.key);
    if (!c) { fmts.push(s.rpr || (null as any)); c = fmts.length; codeOf.set(s.key, c); }
    out += `<f${c}>${s.text}</f${c}>`;
    open++;
  }
  return { text: out, fmts, base: baseSeg?.rpr ?? null, vars };
}

const rid = () => String(300000000 + Math.floor(Math.random() * 600000000));
function enRun(doc: Document, rpr: Element | null, text: string): Element {
  const r = doc.createElementNS(W_NS, "w:r");
  const rp = (rpr ? doc.importNode(rpr, true) : doc.createElementNS(W_NS, "w:rPr")) as Element;
  const lang = doc.createElementNS(W_NS, "w:lang"); lang.setAttribute("w:val", "en-US");
  rp.insertBefore(lang, Array.from(rp.childNodes).find((x) => ["w:eastAsianLayout", "w:specVanish", "w:oMath"].includes(x.nodeName)) || null);
  r.appendChild(rp);
  for (const part of text.split(/(\n|\t)/)) {
    if (part === "\n") r.appendChild(doc.createElementNS(W_NS, "w:br"));
    else if (part === "\t") r.appendChild(doc.createElementNS(W_NS, "w:tab"));
    else if (part) { const t = doc.createElementNS(W_NS, "w:t"); t.setAttribute("xml:space", "preserve"); t.appendChild(doc.createTextNode(part)); r.appendChild(t); }
  }
  return r;
}
function enVarSdt(doc: Document, key: string, rpr: Element | null, value: string): Element {
  const sdt = frag(doc, `<w:sdt><w:sdtPr><w:alias w:val="${esc(key)}"/><w:tag w:val="${VAR_PREFIX}${esc(key)}@en"/><w:id w:val="${rid()}"/><w:text/></w:sdtPr><w:sdtContent/></w:sdt>`);
  const c = sdtContent(sdt) as Element;
  c.appendChild(enRun(doc, rpr, value || `{{${key}}}`));
  return sdt;
}

/** String hasil terjemahan AI -> run (format & variabel dikembalikan). null bila tag/variabel dari AI rusak. */
function decodeToRuns(doc: Document, s: string, enc: Enc, values: Record<string, string>, strict: boolean): Element[] | null {
  const parts = s.split(/(<\/?f\d+>|\{\{[^}]+\}\})/);
  const nodes: Element[] = [];
  const usedVars: string[] = [];
  let cur = 0, opens = 0, closes = 0;
  for (const part of parts) {
    if (!part) continue;
    const tag = part.match(/^<(\/?)f(\d+)>$/);
    if (tag) {
      if (strict) { if (tag[1]) { closes++; cur = 0; } else { opens++; cur = Number(tag[2]); if (!enc.fmts[cur - 1]) return null; } }
      continue;
    }
    const v = part.match(/^\{\{([^}]+)\}\}$/);
    if (v && enc.vars.includes(v[1])) {
      usedVars.push(v[1]);
      nodes.push(enVarSdt(doc, v[1], (cur ? enc.fmts[cur - 1] : enc.base) || enc.base, valueToEnglish(values[v[1]] ?? "")));
      continue;
    }
    nodes.push(enRun(doc, (cur ? enc.fmts[cur - 1] : enc.base) || enc.base, part));
  }
  if (strict) {
    if (opens !== closes) return null;
    if ([...usedVars].sort().join("|") !== [...enc.vars].sort().join("|")) return null;
  }
  return nodes;
}

/** Nilai variabel yang sekarang ada di dokumen (tanpa yang "@en"), dilengkapi nilai cadangan. */
function varValues(doc: Document, extra: Record<string, string> = {}): Record<string, string> {
  const out: Record<string, string> = { ...extra };
  const seen = new Set<string>();
  for (const sdt of sdtOf(doc)) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(VAR_PREFIX) || tag.endsWith("@en")) continue;
    const key = tag.slice(VAR_PREFIX.length);
    if (seen.has(key)) continue;
    seen.add(key);
    const c = sdtContent(sdt);
    const t = c ? Array.from(c.getElementsByTagName("w:t")).map((x) => x.textContent || "").join("") : "";
    if (t) out[key] = t;
  }
  return out;
}
function setSdtText(doc: Document, sdt: Element, value: string): boolean {
  const content = sdtContent(sdt);
  if (!content) return false;
  const cur = Array.from(content.getElementsByTagName("w:t")).map((x) => x.textContent || "").join("");
  if (cur === value) return false;
  const firstRun = Array.from(content.getElementsByTagName("w:r"))[0] as Element | undefined;
  const rPr = firstRun ? Array.from(firstRun.childNodes).find((n) => n.nodeName === "w:rPr") : undefined;
  while (content.firstChild) content.removeChild(content.firstChild);
  const r = doc.createElementNS(W_NS, "w:r");
  if (rPr) r.appendChild(rPr.cloneNode(true));
  const t = doc.createElementNS(W_NS, "w:t"); t.setAttribute("xml:space", "preserve"); t.appendChild(doc.createTextNode(value));
  r.appendChild(t); content.appendChild(r);
  return true;
}
/** Samakan variabel di dalam `root`: "K" = nilai K, "K@en" = nilai K versi Inggris (tanggal diterjemahkan). */
function refreshVarsIn(doc: Document, root: Element, vals: Record<string, string>): number {
  let n = 0;
  const list = root.nodeName === "w:sdt" ? [root] : [];
  for (const sdt of [...list, ...(Array.from(root.getElementsByTagName("w:sdt")) as Element[])]) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(VAR_PREFIX)) continue;
    const en = tag.endsWith("@en");
    const key = tag.slice(VAR_PREFIX.length).replace(/@en$/, "");
    if (vals[key] === undefined) continue;
    if (setSdtText(doc, sdt, en ? valueToEnglish(vals[key]) : vals[key])) n++;
  }
  return n;
}
/** Samakan semua variabel bertanda "@en" di dokumen dengan nilai variabel Indonesianya (dipanggil saat sinkron). */
export async function refreshEnglishVars(docx: Buffer, extra: Record<string, string> = {}): Promise<{ docx: Buffer; changed: number }> {
  const { zip, doc } = await loadBody(docx);
  const changed = refreshVarsIn(doc, doc.documentElement, varValues(doc, extra));
  if (!changed) return { docx, changed: 0 };
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return { docx: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), changed };
}

// ---- susunan: kiri-kanan (tabel) / atas-bawah (per pasal) --------------------------------------------------
const frag = (doc: Document, xml: string): Element => {
  const f = new DOMParser().parseFromString(`<w:x xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${xml}</w:x>`, "application/xml") as unknown as Document;
  return doc.importNode(f.documentElement.firstChild as Node, true) as Element;
};

/** Kembalikan tabel kiri-kanan ke paragraf biasa (urutan tetap). */
function unwrapSide(doc: Document): void {
  for (const tbl of (Array.from(doc.getElementsByTagName("w:tbl")) as Element[]).filter(isSideTable)) {
    const parent = tbl.parentNode;
    if (!parent) continue;
    for (const tr of Array.from(tbl.childNodes).filter((c) => c.nodeName === "w:tr"))
      for (const tc of Array.from(tr.childNodes).filter((c) => c.nodeName === "w:tc"))
        for (const k of Array.from(tc.childNodes)) if (k.nodeName !== "w:tcPr") parent.insertBefore(k, tbl);
    parent.removeChild(tbl);
  }
}

/** Bentuk baku: tabel dibuka, tiap paragraf Inggris tepat di bawah pasangan Indonesianya. */
function canonicalize(doc: Document): void {
  unwrapSide(doc);
  const ids = new Map<string, Element>();
  const paras = scopeParas(doc);
  for (const p of paras) { const n = markNo(p, ID_MARK); if (n) ids.set(n, p); }
  for (const p of paras) {
    const n = markNo(p, EN_MARK);
    const idEl = n ? ids.get(n) : undefined;
    if (idEl && idEl.parentNode && idEl.parentNode === p.parentNode && idEl.nextSibling !== p) idEl.parentNode.insertBefore(p, idEl.nextSibling);
  }
}

/** Susun tiap pasangan paragraf (Indonesia, Inggris) menjadi satu baris tabel: kiri Indonesia, kanan Inggris. */
function wrapPairs(doc: Document): void {
  const body = doc.getElementsByTagName("w:body")[0] as Element;
  const parents: Element[] = [body, ...(sdtOf(doc).map((x) => sdtContent(x)).filter((c): c is Element => !!c && Array.from(c.childNodes).some((k) => k.nodeName === "w:p")))];
  const cell = (side: 0 | 1) => `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/><w:tcMar>${side === 0 ? '<w:right w:w="170" w:type="dxa"/>' : '<w:left w:w="170" w:type="dxa"/>'}</w:tcMar></w:tcPr></w:tc>`;
  for (const parent of parents) {
    if (layoutAncestor(parent.firstChild || parent).some((x) => x === "kop" || x === "ttd")) continue;
    const kids = Array.from(parent.childNodes);
    const groups: { id: Element; en: Element }[][] = [];
    let cur: { id: Element; en: Element }[] = [];
    for (let i = 0; i < kids.length; i++) {
      const a = kids[i], b = kids[i + 1];
      if (a.nodeName === "w:p" && b && b.nodeName === "w:p" && isIdPara(a as Element) && isEnPara(b as Element) && markNo(a as Element, ID_MARK) === markNo(b as Element, EN_MARK)) {
        cur.push({ id: a as Element, en: b as Element }); i++;
      } else if (cur.length) { groups.push(cur); cur = []; }
    }
    if (cur.length) groups.push(cur);
    for (const g of groups) {
      const tbl = frag(doc, `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblCaption w:val="${SIDE_CAPTION}"/></w:tblPr><w:tblGrid><w:gridCol w:w="4513"/><w:gridCol w:w="4513"/></w:tblGrid></w:tbl>`);
      parent.insertBefore(tbl, g[0].id);
      for (const pair of g) {
        const tr = frag(doc, `<w:tr>${cell(0)}${cell(1)}</w:tr>`);
        const tcs = Array.from(tr.childNodes).filter((c) => c.nodeName === "w:tc");
        tcs[0].appendChild(pair.id); tcs[1].appendChild(pair.en);
        tbl.appendChild(tr);
      }
    }
  }
}

/** Atas-bawah per pasal: seluruh judul+isi Indonesia dulu, baru seluruh judul+isi Inggris di bawahnya. */
function groupClauseBlocks(doc: Document): void {
  for (const sdt of sdtOf(doc)) {
    if (!sdtTag(sdt).startsWith(CLAUSE_PREFIX)) continue;
    const c = sdtContent(sdt);
    if (!c) continue;
    const kids = Array.from(c.childNodes);
    let run: Element[] = [];
    const flush = (after: Node | null) => {
      const ens = run.filter((p) => isEnPara(p));
      if (ens.length && ens.length < run.length) for (const e of ens) { c.removeChild(e); c.insertBefore(e, after); }
      run = [];
    };
    for (const k of kids) {
      if (k.nodeName === "w:p" && (isIdPara(k as Element) || isEnPara(k as Element))) run.push(k as Element);
      else flush(k);
    }
    flush(null);
  }
}

// ---- API ----------------------------------------------------------------------------------------------------
/** Paragraf isi yang layak diterjemahkan. onlyNew = hanya yang belum punya pasangan Inggris. */
function translatable(doc: Document, onlyNew: boolean): Element[] {
  return scopeParas(doc).filter((p) => {
    if (isEnPara(p) || !/\p{L}/u.test(paraText(p)) || p.getElementsByTagName("w:drawing").length > 0) return false;
    return onlyNew ? !isIdPara(p) : true;
  });
}

/** Teks (dengan penanda format/variabel) tiap paragraf yang perlu diterjemahkan, berurutan. Dokumen harus memuat bahasa Indonesia. */
export async function listTranslatable(docx: Buffer, opts: { onlyNew?: boolean } = {}): Promise<string[]> {
  const { doc } = await loadBody(docx);
  canonicalize(doc);
  return translatable(doc, !!opts.onlyNew).map((p) => encodePara(p).text.replace(/[ ]{2,}/g, " ").trim());
}

/** Pasang terjemahan (selaras dengan listTranslatable). Tanpa onlyNew, terjemahan lama diganti. Hasil: dua bahasa tampil. */
export async function applyTranslations(docx: Buffer, en: string[], opts: { onlyNew?: boolean; values?: Record<string, string> } = {}): Promise<Buffer> {
  const { zip, doc } = await loadBody(docx);
  canonicalize(doc);
  if (!opts.onlyNew) for (const p of scopeParas(doc)) if (isEnPara(p)) p.parentNode?.removeChild(p);
  const paras = translatable(doc, !!opts.onlyNew);
  if (paras.length !== en.length) throw new Error(`Jumlah terjemahan (${en.length}) tidak sama dengan jumlah paragraf (${paras.length}).`);
  const values = varValues(doc, opts.values);
  const numCtx = await NumCtx.load(zip);
  let next = 1 + Math.max(0, ...scopeParas(doc).flatMap((p) => [markNo(p, ID_MARK), markNo(p, EN_MARK)]).filter((x): x is string => !!x).map(Number));
  paras.forEach((p, i) => {
    const text = (en[i] || "").trim();
    if (!text) return;
    const enc = encodePara(p);
    const nodes = decodeToRuns(doc, text, enc, values, true) || decodeToRuns(doc, text, enc, values, false) || [];
    if (!nodes.length) return;
    const n = markNo(p, ID_MARK) || String(next++);
    if (!isIdPara(p)) addMark(doc, p, ID_MARK, n);
    const np = doc.createElementNS(W_NS, "w:p");
    const oldPPr = Array.from(p.childNodes).find((x) => x.nodeName === "w:pPr") as Element | undefined;
    const pPr = (oldPPr ? oldPPr.cloneNode(true) : doc.createElementNS(W_NS, "w:pPr")) as Element;
    for (const x of Array.from(pPr.childNodes)) if (["w:keepNext", "w:pageBreakBefore", "w:sectPr", "w:rPr", "w:pStyle", "w:pBdr"].includes(x.nodeName)) pPr.removeChild(x);
    np.appendChild(pPr);
    mapNumbering(np, p, numCtx);
    for (const x of nodes) np.appendChild(x);
    addMark(doc, np, EN_MARK, n);
    p.parentNode?.insertBefore(np, p.nextSibling);
  });
  numCtx?.save(zip);
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Daftar bernomor pada paragraf `el` dibuat sendiri (nomor mulai dari awal) mengikuti bentuk daftar pasangannya `partner`. */
function mapNumbering(el: Element, partner: Element, ctx: NumCtx | null): void {
  const mine = el.getElementsByTagName("w:numPr")[0] as Element | undefined;
  if (!mine) return;
  const theirs = numPrOf(partner);
  const d = ctx && theirs ? ctx.derive(theirs.numId, true) : null;
  const idEl = mine.getElementsByTagName("w:numId")[0] as Element | undefined;
  if (d && idEl) idEl.setAttribute("w:val", d);
  else mine.parentNode?.removeChild(mine); // tak ada daftar untuk diikuti: jadi paragraf biasa
}

const serP = (p: Element) => new XMLSerializer().serializeToString(p as any);

/**
 * Tampilkan bahasa. id / en: bahasa lain dikeluarkan dari dokumen ke `stash`; bilingual: keduanya dipasang
 * (kiri-kanan = tabel, atas-bawah = per pasal). `values` = nilai variabel cadangan (data sistem).
 */
export async function setLanguageMode(docx: Buffer, mode: DocLanguage, layout: BilingualLayoutMode, stash: LangStash, values: Record<string, string> = {}): Promise<{ docx: Buffer; stash: LangStash }> {
  const { zip, doc } = await loadBody(docx);
  canonicalize(doc);
  const ids = new Map<string, Element>(), ens = new Map<string, Element>();
  for (const p of scopeParas(doc)) {
    const i = markNo(p, ID_MARK), e = markNo(p, EN_MARK);
    if (i) ids.set(i, p);
    if (e) ens.set(e, p);
  }
  const vals = varValues(doc, values);
  const numCtx = await NumCtx.load(zip);
  const parseP = (s: string): Element => frag(doc, s);
  const next: LangStash = {};
  for (const n of new Set([...ids.keys(), ...ens.keys(), ...Object.keys(stash)])) {
    let idEl = ids.get(n), enEl = ens.get(n);
    const st = stash[n] || {};
    if (!idEl && st.id && enEl?.parentNode) { idEl = parseP(st.id); enEl.parentNode.insertBefore(idEl, enEl); mapNumbering(idEl, enEl, numCtx); refreshVarsIn(doc, idEl, vals); }
    if (!enEl && st.en && idEl?.parentNode) { enEl = parseP(st.en); idEl.parentNode.insertBefore(enEl, idEl.nextSibling); mapNumbering(enEl, idEl, numCtx); refreshVarsIn(doc, enEl, vals); }
    const keep: { id?: string; en?: string } = {};
    // sisi yang tidak dipakai disimpan (tidak dibuang); sisi yang sedang tidak punya pasangan dibiarkan tampil
    if (mode === "id" && enEl && idEl) { keep.en = serP(enEl); enEl.parentNode?.removeChild(enEl); }
    else if (mode === "en" && idEl && enEl) { keep.id = serP(idEl); idEl.parentNode?.removeChild(idEl); }
    else {
      // bahasa yang tidak dipilih sudah tersimpan di stash tapi tak bisa dipasang (pasangannya hilang): pertahankan
      if (mode === "id" && !idEl && st.id) keep.id = st.id;
      if (mode === "en" && !enEl && st.en) keep.en = st.en;
    }
    if (keep.id || keep.en) next[n] = keep;
  }
  if (mode === "bilingual") { if (layout === "side") wrapPairs(doc); else groupClauseBlocks(doc); }
  numCtx?.prune(doc, true);
  numCtx?.save(zip);
  await applyRangkapTail(zip, doc); // perkiraan halaman tanda tangan ikut berubah mengikuti bahasa
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return { docx: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), stash: next };
}

/** Keadaan bahasa dokumen: sudah diterjemahkan? mode yang tampil? susunan dwibahasa? ada paragraf baru yang belum diterjemahkan? */
export async function readLanguage(docx: Buffer, stash: LangStash = {}): Promise<{ translated: boolean; mode: DocLanguage; layout: BilingualLayoutMode; hasNew: boolean }> {
  const { doc } = await loadBody(docx);
  const paras = scopeParas(doc);
  const hasId = paras.some(isIdPara), hasEn = paras.some(isEnPara);
  const translated = hasId || hasEn || Object.keys(stash).length > 0;
  const mode: DocLanguage = hasId && hasEn ? "bilingual" : hasEn ? "en" : "id";
  const side = (Array.from(doc.getElementsByTagName("w:tbl")) as Element[]).some(isSideTable);
  const hasNew = translated && mode !== "en" && translatable(doc, true).length > 0;
  return { translated, mode, layout: side ? "side" : "stacked", hasNew };
}

/** Keadaan kop / logo / tanda tangan di dokumen (dibaca dari berkas, bukan dari pengaturan). */
export async function readLayout(docx: Buffer): Promise<{ kop: boolean; logo: boolean; ttd: boolean; digital: boolean }> {
  const { doc } = await loadBody(docx);
  const find = (kind: string) => sdtOf(doc).find((x) => sdtTag(x) === LAYOUT_PREFIX + kind);
  const kop = find("kop"), ttd = find("ttd");
  return {
    kop: !!kop, logo: !!kop && kop.getElementsByTagName("w:drawing").length > 0,
    ttd: !!ttd, digital: !!ttd && ttd.getElementsByTagName("w:drawing").length > 0,
  };
}

/** Pasang / hapus / ganti kop surat dan blok tanda tangan pada dokumen yang sudah ada (isi lain tidak disentuh). */
export async function patchLayout(docx: Buffer, patch: LayoutPatch): Promise<Buffer> {
  const { zip, doc } = await loadBody(docx);
  const body = doc.getElementsByTagName("w:body")[0] as Element;
  const relsEntry = zip.file("word/_rels/document.xml.rels");
  const typesEntry = zip.file("[Content_Types].xml");
  let rels = relsEntry ? await relsEntry.async("string") : "";
  let types = typesEntry ? await typesEntry.async("string") : "";
  const remove = (kind: string) => { for (const x of sdtOf(doc)) if (sdtTag(x) === LAYOUT_PREFIX + kind) x.parentNode?.removeChild(x); };
  const adopt = (xml: string): Node[] => {
    const frag = new DOMParser().parseFromString(`<w:body xmlns:w="${W_NS}">${xml}</w:body>`, "application/xml") as unknown as Document;
    return Array.from(frag.documentElement.childNodes).map((n) => doc.importNode(n, true));
  };
  if (patch.kop !== undefined) {
    remove("kop");
    rels = dropMedia(zip, rels, "rIdClmKop");
    if (patch.kop) {
      const l = patch.kop.letterhead;
      const media = new Media("rIdClmKop");
      const info = l.logo ? imageInfo(l.logo.bytes, l.logo.mimeType) : null;
      const xml = kopBlockXml(l, info, media);
      ({ rels, types } = embedMedia(zip, rels, types, media.items));
      const first = body.firstChild;
      for (const n of adopt(xml)) body.insertBefore(n, first);
    }
  }
  if (patch.ttd !== undefined) {
    remove("ttd");
    rels = dropMedia(zip, rels, "rIdClmTtd");
    if (patch.ttd) {
      const media = new Media("rIdClmTtd");
      const xml = signatureBlockXml(patch.ttd.signatures, patch.ttd.options, media, patch.ttd.digital);
      ({ rels, types } = embedMedia(zip, rels, types, media.items));
      const last = Array.from(body.childNodes).reverse().find((n) => n.nodeName === "w:sectPr") || null;
      for (const n of adopt(xml)) body.insertBefore(n, last);
    }
  }
  if (patch.kop !== undefined || patch.ttd !== undefined) await applyRangkapTail(zip, doc);
  if (relsEntry) zip.file("word/_rels/document.xml.rels", rels);
  if (typesEntry) zip.file("[Content_Types].xml", types);
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

export interface BilingualInput {
  layout: "side" | "stacked";
  docType: string;
  title: string;
  numberLine: string;
  /** Narasi pembuka & kalimat penutup pembuka dalam bahasa Inggris (boleh kosong). */
  opening?: string;
  closing?: string;
}

/** Buat .docx dari daftar pasal. `values` mengisi variabel yang sudah diketahui. */
export async function buildTemplateDocx(opts: {
  title: string;
  /** Jenis dokumen (judul besar, mis. "SURAT PERJANJIAN KERJASAMA"); `title` jadi sub-judul. */
  docType?: string;
  /** Baris nomor dokumen, mis. "NOMOR: 002/CI-...". */
  numberLine?: string;
  clauses: TemplateClauseInput[];
  values?: Record<string, string>;
  /** Narasi pembuka (blok bebas, tidak dikunci): paragraf awal + identifikasi pihak. */
  opening?: string;
  /** Kalimat penutup pembuka ("... dengan ketentuan dan pasal-pasal sebagai berikut:"). */
  closing?: string;
  /** Kop surat di awal dokumen. */
  letterhead?: LetterheadInput;
  /** Kolom tanda tangan di akhir dokumen. */
  signatures?: SignatureInput[];
  signatureOptions?: SignatureOptions;
  digitalSignature?: DigitalSignatureInput;
  /** Dokumen dua bahasa: Indonesia + Inggris berdampingan/atas-bawah. */
  bilingual?: BilingualInput;
  /** Gambar pasal yang sudah dimuat: src -> byte (PNG/JPEG). */
  images?: Record<string, Buffer>;
}): Promise<Buffer> {
  const values = opts.values || {};
  const mediaImg = new Media("rIdClmImg"), mediaKop = new Media("rIdClmKop"), mediaTtd = new Media("rIdClmTtd");
  const imgRid = new Map<string, { rid: string; cx: number; cy: number }>();
  const ctx: Ctx = {
    variable: (key, rPr) => varSdt(key, values[key] ?? `{{${key}}}`, rPr),
    numbering: [], skipped: [], defaultJc: "both",
    image: (src, alt) => {
      const bytes = opts.images?.[src];
      if (!bytes) return null;
      let e = imgRid.get(src);
      if (!e) {
        const info = imageInfo(bytes, "");
        if (!info) return null;
        const { cx, cy } = fitEmu(info.width, info.height, 5400000, 4200000);
        e = { rid: mediaImg.add(bytes, info.ext), cx, cy };
        imgRid.set(src, e);
      }
      return drawingRun(e.rid, e.cx, e.cy, alt || "Gambar");
    },
  };
  const bi = opts.bilingual;
  const layout = bi?.layout || "side";
  const openingXml = opts.opening ? clauseContentToWordXml(opts.opening, ctx) : "";
  const closingXml = opts.closing ? clauseContentToWordXml(opts.closing, ctx) : "";
  const openingEn = bi?.opening ? clauseContentToWordXml(bi.opening, ctx) : "";
  const closingEn = bi?.closing ? clauseContentToWordXml(bi.closing, ctx) : "";
  const clauseXml = opts.clauses.map((c) => clauseSdt(c, ctx, bi ? { layout } : undefined));
  const hasNumbering = ctx.numbering.length > 0;
  const logoInfo = opts.letterhead?.logo ? imageInfo(opts.letterhead.logo.bytes, opts.letterhead.logo.mimeType) : null;
  const kopXml = opts.letterhead ? kopBlockXml(opts.letterhead, logoInfo, mediaKop) : "";
  const sigXml = opts.signatures && opts.signatures.length
    ? signatureBlockXml(opts.signatures, opts.signatureOptions || { copies: 1, meterai: false }, mediaTtd, opts.digitalSignature)
    : "";
  const head = (docType: string | undefined, title: string, numberLine: string | undefined) => docType
    ? `<w:p><w:pPr><w:keepNext/><w:spacing w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${esc(docType)}</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:keepNext/><w:spacing w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${esc(title)}</w:t></w:r></w:p>` +
      (numberLine ? `<w:p><w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${esc(numberLine)}</w:t></w:r></w:p>` : "")
    : `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${esc(title)}</w:t></w:r></w:p>`;
  const headXml = bi
    ? pairXml(head(opts.docType, opts.title, opts.numberLine), head(bi.docType, bi.title, bi.numberLine), layout)
    : head(opts.docType, opts.title, opts.numberLine);
  const body = [
    kopXml,
    headXml,
    bi ? pairXml(openingXml, openingEn, layout) : openingXml,
    bi ? pairXml(closingXml, closingEn, layout) : closingXml,
    ...clauseXml,
    sigXml,
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`,
  ].join("");
  const zip = new JSZip();
  let types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${hasNumbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ""}</Types>`;
  let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${hasNumbering ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : ""}</Relationships>`;
  ({ rels, types } = embedMedia(zip, rels, types, [...mediaImg.items, ...mediaKop.items, ...mediaTtd.items]));
  zip.file("[Content_Types].xml", types);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", rels);
  if (hasNumbering) zip.file("word/numbering.xml", numberingXml(ctx.numbering));
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="id-ID"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`);
  const out = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  return opts.signatures?.length && opts.signatureOptions?.copies === 2 ? refreshRangkapTail(out) : out;
}

// ---- Membaca / mengubah dokumen -------------------------------------------
async function loadBody(docx: Buffer): Promise<{ zip: JSZip; doc: Document }> {
  const zip = await JSZip.loadAsync(docx);
  const entry = zip.file("word/document.xml");
  if (!entry) throw new Error("Bukan berkas .docx yang valid (word/document.xml tidak ada)");
  const doc = new DOMParser().parseFromString(await entry.async("string"), "application/xml") as unknown as Document;
  return { zip, doc };
}

function sdtOf(doc: Document): Element[] {
  return Array.from(doc.getElementsByTagName("w:sdt")) as Element[];
}

function sdtTag(sdt: Element): string {
  const pr = Array.from(sdt.childNodes).find((n) => (n as Element).nodeName === "w:sdtPr") as Element | undefined;
  if (!pr) return "";
  const tag = Array.from(pr.childNodes).find((n) => (n as Element).nodeName === "w:tag") as Element | undefined;
  return tag?.getAttribute("w:val") || "";
}

function sdtContent(sdt: Element): Element | undefined {
  return Array.from(sdt.childNodes).find((n) => (n as Element).nodeName === "w:sdtContent") as Element | undefined;
}

/** Teks sebuah node; variabel di dalamnya dinetralkan jadi {{Key}} supaya nilainya boleh berubah. */
function textOf(node: Node): string {
  let out = "";
  for (const child of Array.from(node.childNodes)) {
    const name = child.nodeName;
    if (name === "w:sdt") {
      const tag = sdtTag(child as Element);
      if (tag.startsWith(VAR_PREFIX)) { out += `{{${tag.slice(VAR_PREFIX.length)}}}`; continue; }
      const c = sdtContent(child as Element);
      if (c) out += textOf(c);
    } else if (name === "w:t") {
      out += child.textContent || "";
    } else if (name === "w:tab") {
      out += "\t";
    } else if (name === "w:br") {
      out += "\n";
    } else if (name === "w:del" || name === "w:delText") {
      // teks yang dihapus lewat Track Changes tidak dihitung
    } else {
      out += textOf(child);
    }
    if (name === "w:p") out += "\n";
  }
  return out;
}

const normalize = (s: string) => s.replace(/\s+/g, " ").trim();

/** Peta id pasal -> teks ternormalisasi (variabel dinetralkan). */
export async function clauseTexts(docx: Buffer): Promise<Record<string, string>> {
  const { doc } = await loadBody(docx);
  const out: Record<string, string> = {};
  for (const sdt of sdtOf(doc)) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(CLAUSE_PREFIX)) continue;
    const c = sdtContent(sdt);
    out[tag.slice(CLAUSE_PREFIX.length)] = c ? normalize(textOf(c)) : "";
  }
  return out;
}

/** Peta id pasal -> judul (dari paragraf pertama pasal). */
async function clauseTitles(docx: Buffer): Promise<Record<string, string>> {
  const { doc } = await loadBody(docx);
  const out: Record<string, string> = {};
  for (const sdt of sdtOf(doc)) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(CLAUSE_PREFIX)) continue;
    const c = sdtContent(sdt);
    const first = c ? Array.from(c.childNodes).find((n) => n.nodeName === "w:p") : undefined;
    out[tag.slice(CLAUSE_PREFIX.length)] = first ? normalize(textOf(first)) : "";
  }
  return out;
}

/** Ada teks selain judul pasal? (judul = paragraf pertama versi sebelumnya) */
function hasBodyText(text: string, title: string | undefined): boolean {
  const rest = title && text.startsWith(title) ? text.slice(title.length) : text;
  return normalize(rest).length > 0;
}

/** Nilai variabel yang sekarang ada di dokumen (kalau satu variabel muncul berkali-kali, yang pertama). */
export async function listVariables(docx: Buffer): Promise<Record<string, string>> {
  const { doc } = await loadBody(docx);
  const out: Record<string, string> = {};
  for (const sdt of sdtOf(doc)) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(VAR_PREFIX)) continue;
    const key = tag.slice(VAR_PREFIX.length);
    if (!(key in out)) {
      const c = sdtContent(sdt);
      const t = c ? Array.from((c as Element).getElementsByTagName("w:t")).map((x) => x.textContent || "").join("") : "";
      out[key] = t;
    }
  }
  return out;
}

/** Ganti isi variabel (semua kemunculannya). Format huruf run pertama dipertahankan. */
export async function setVariables(docx: Buffer, values: Record<string, string>): Promise<{ docx: Buffer; changed: number }> {
  const { zip, doc } = await loadBody(docx);
  let changed = 0;
  for (const sdt of sdtOf(doc)) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(VAR_PREFIX)) continue;
    const key = tag.slice(VAR_PREFIX.length);
    if (!(key in values)) continue;
    const content = sdtContent(sdt);
    if (!content) continue;
    const firstRun = Array.from((content as Element).getElementsByTagName("w:r"))[0] as Element | undefined;
    const rPr = firstRun ? Array.from(firstRun.childNodes).find((n) => n.nodeName === "w:rPr") : undefined;
    while (content.firstChild) content.removeChild(content.firstChild);
    const r = doc.createElementNS(W_NS, "w:r");
    if (rPr) r.appendChild(rPr.cloneNode(true));
    const t = doc.createElementNS(W_NS, "w:t");
    t.setAttribute("xml:space", "preserve");
    t.appendChild(doc.createTextNode(values[key]));
    r.appendChild(t);
    content.appendChild(r);
    changed++;
  }
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return { docx: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), changed };
}

/**
 * Penjaga di server: bandingkan hasil edit dgn versi SEBELUMNYA.
 *  - pasal terlindungi berubah isinya / hilang  -> pelanggaran (kecuali `canOverride`, mis. legal/admin)
 *  - pasal wajib hilang                         -> pelanggaran (kecuali `canOverride`)
 * Variabel di dalam pasal terlindungi boleh berubah (dinetralkan sebelum dibandingkan).
 */
export async function checkTemplateEdit(opts: { baseline: Buffer; edited: Buffer; policy: PolicyMap; canOverride: boolean }): Promise<Violation[]> {
  if (opts.canOverride) return [];
  const before = await clauseTexts(opts.baseline);
  const after = await clauseTexts(opts.edited);
  const titles = await clauseTitles(opts.baseline);
  const violations: Violation[] = [];
  for (const [id, pol] of Object.entries(opts.policy)) {
    if (!(id in before)) continue; // pasal ini memang tidak ada di versi sebelumnya
    const removed = !(id in after);
    if (removed && pol.isProtected) violations.push({ clauseId: id, kind: "protected-removed", message: `Pasal terlindungi "${id}" dihapus; hanya legal yang boleh.` });
    else if (removed && pol.isMandatory) violations.push({ clauseId: id, kind: "mandatory-removed", message: `Pasal wajib "${id}" dihapus.` });
    // Editor tidak selalu menahan penghapusan: isi dihapus tapi kotak pasal tersisa (kosong) juga dianggap dihapus.
    else if (!removed && pol.isMandatory && !hasBodyText(after[id], titles[id])) violations.push({ clauseId: id, kind: "mandatory-removed", message: `Isi pasal wajib "${id}" dikosongkan.` });
    else if (!removed && pol.isProtected && before[id] !== after[id]) violations.push({ clauseId: id, kind: "protected-changed", message: `Isi pasal terlindungi "${id}" diubah; hanya legal yang boleh.` });
  }
  return violations;
}

/**
 * Pasang kunci pasal sesuai kebijakan (dari database) pada dokumen.
 *  - terlindungi -> sdtContentLocked; wajib -> sdtLocked; selain itu tanpa kunci
 *  - `unlockProtected` (utk legal/admin yang boleh mengubah): kunci pasal terlindungi dibuka
 *    (tetap sdtLocked bila pasalnya wajib, supaya tidak terhapus tak sengaja)
 * Dokumen tanpa pasal bertanda dikembalikan apa adanya (tanpa ditulis ulang).
 */
export async function applyLocks(docx: Buffer, policy: PolicyMap, opts: { unlockProtected: boolean }): Promise<Buffer> {
  const { zip, doc } = await loadBody(docx);
  let touched = 0;
  for (const sdt of sdtOf(doc)) {
    const tag = sdtTag(sdt);
    if (!tag.startsWith(CLAUSE_PREFIX)) continue;
    const pol = policy[tag.slice(CLAUSE_PREFIX.length)];
    if (!pol) continue;
    const want = pol.isProtected && !opts.unlockProtected ? "sdtContentLocked" : pol.isMandatory ? "sdtLocked" : "";
    const pr = Array.from(sdt.childNodes).find((n) => n.nodeName === "w:sdtPr") as Element | undefined;
    if (!pr) continue;
    const old = Array.from(pr.childNodes).find((n) => n.nodeName === "w:lock") as Element | undefined;
    const have = old?.getAttribute("w:val") || "";
    if (have === want) continue;
    if (old) pr.removeChild(old);
    if (want) {
      const lock = doc.createElementNS(W_NS, "w:lock");
      lock.setAttribute("w:val", want);
      pr.appendChild(lock);
    }
    touched++;
  }
  if (!touched) return docx;
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Peran yang boleh mengubah pasal terlindungi & menghapus pasal wajib (sama dgn aturan edit perpustakaan pasal). */
export const CLAUSE_OVERRIDE_ROLES = ["admin", "legal", "super_admin"];

/** Perbaikan dokumen lama: versi awal fitur bahasa menyembunyikan teks (w:vanish) yang tidak terbaca di editor/penampil lain. Tampilkan lagi. */
export async function unhideLegacyLanguage(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx);
  const f = zip.file("word/document.xml");
  if (!f) return docx;
  const xml = await f.async("string");
  if (!/clm_(id|en)_\d+/.test(xml) || !/<w:vanish\b/.test(xml)) return docx;
  zip.file("word/document.xml", xml.replace(/<w:vanish\b[^>]*\/>/g, ""));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
