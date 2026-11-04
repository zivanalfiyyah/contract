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

function clauseSdt(c: TemplateClauseInput, ctx: Ctx): string {
  const lock = c.isProtected ? "sdtContentLocked" : c.isMandatory ? "sdtLocked" : "";
  const body = clauseContentToWordXml(c.content, ctx);
  const head = `<w:p><w:pPr><w:keepNext/></w:pPr>${run(c.title, true)}</w:p>`;
  return `<w:sdt><w:sdtPr><w:alias w:val="Pasal: ${esc(c.title)}"/><w:tag w:val="${CLAUSE_PREFIX}${esc(c.id)}"/><w:id w:val="${nextId()}"/>${lock ? `<w:lock w:val="${lock}"/>` : ""}</w:sdtPr><w:sdtContent>${head}${body}</w:sdtContent></w:sdt>`;
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

const LAYOUT_PREFIX = "layout:";
const LOGO_REL = "rIdClmLogo";
let layoutId = 7000;
const layoutSdt = (kind: "kop" | "ttd", alias: string, inner: string) =>
  `<w:sdt><w:sdtPr><w:alias w:val="${alias}"/><w:tag w:val="${LAYOUT_PREFIX}${kind}"/><w:id w:val="${layoutId++}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** Kop surat: blok di awal dokumen (halaman 1) — logo kiri, identitas perusahaan di tengah, garis bawah. */
export function kopBlockXml(l: LetterheadInput, logo: { ext: string; width: number; height: number } | null): string {
  const lines = [
    cp(l.name.toUpperCase(), { bold: true, size: 32 }),
    ...(l.address ? [cp(l.address, { size: 20, color: "444444" })] : []),
    ...(l.contact ? [cp(l.contact, { size: 20, color: "444444" })] : []),
    ...(l.web ? [cp(l.web, { size: 20, color: "444444" })] : []),
    ...(l.registration ? [cp(l.registration, { size: 20, color: "666666" })] : []),
  ].join("");
  let cell1 = "";
  let cols = `<w:gridCol w:w="9026"/>`;
  if (logo) {
    const W0 = Math.min(Math.round((900000 * logo.width) / Math.max(logo.height, 1)), 1800000);
    const H0 = Math.round((W0 * logo.height) / Math.max(logo.width, 1));
    cell1 = `<w:tc><w:tcPr><w:tcW w:w="1700" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${W0}" cy="${H0}"/><wp:docPr id="7001" name="Logo"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="7001" name="logo.${logo.ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${LOGO_REL}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W0}" cy="${H0}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:tc>`;
    cols = `<w:gridCol w:w="1700"/><w:gridCol w:w="7326"/>`;
  }
  const cell2 = `<w:tc><w:tcPr><w:tcW w:w="${logo ? 7326 : 9026}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>${lines}</w:tc>`;
  const rule = `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="475569"/></w:pBdr><w:spacing w:before="60" w:after="240"/></w:pPr></w:p>`;
  return layoutSdt("kop", "Kop surat", `<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>${NO_BORDERS}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${cols}</w:tblGrid><w:tr>${cell1 ? cell1 : ""}${cell2}</w:tr></w:tbl>${rule}`);
}

export interface SignatureOptions { copies: 1 | 2; meterai: boolean }

/** Satu tabel tanda tangan: peran, ruang tanda tangan (+ kotak meterai di pihak `meteraiIdx`), ( nama ). */
function sigTable(sigs: SignatureInput[], meteraiIdx: number): string {
  const n = Math.max(sigs.length, 1);
  const w = Math.floor(9026 / n);
  const row = (inner: (s: SignatureInput, i: number) => string, h?: number) =>
    `<w:tr><w:trPr><w:cantSplit/>${h ? `<w:trHeight w:val="${h}" w:hRule="atLeast"/>` : ""}</w:trPr>${sigs.map((s, i) => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>${inner(s, i)}</w:tc>`).join("")}</w:tr>`;
  const box = `<w:p><w:pPr><w:keepNext/><w:pBdr><w:top w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/><w:left w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/><w:bottom w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/><w:right w:val="dashed" w:sz="6" w:space="4" w:color="64748B"/></w:pBdr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="${Math.max(Math.floor(w / 2) - 700, 0)}" w:right="${Math.max(Math.floor(w / 2) - 700, 0)}"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="64748B"/></w:rPr><w:t>METERAI</w:t></w:r><w:r><w:rPr><w:sz w:val="16"/><w:color w:val="64748B"/></w:rPr><w:br/><w:t>Rp10.000</w:t></w:r></w:p>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/>${NO_BORDERS}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${sigs.map(() => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>` +
    row((s) => `<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${esc(s.role.toUpperCase())}</w:t></w:r></w:p>`) +
    row((_s, i) => (i === meteraiIdx ? box : `<w:p><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr></w:p>`), 1300) +
    row((s) => `<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">( ${esc(s.name)} )</w:t></w:r></w:p>`) +
    `</w:tbl>`;
}

const capt = (t: string) => `<w:p><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="20"/><w:color w:val="475569"/></w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`;

/**
 * Blok tanda tangan di akhir dokumen. Tunggal: satu blok, meterai di Pihak Pertama.
 * Dua rangkap: dua blok (tiap rangkap di halaman sendiri) — rangkap 1 dipegang Pihak Pertama
 * (meterai di kolomnya), rangkap 2 dipegang Pihak Kedua. Tanpa meterai: kotak meterai tidak dibuat.
 */
export function signatureBlockXml(sigs: SignatureInput[], o: SignatureOptions): string {
  const holder = (i: number) => sigs[i]?.role || (i === 0 ? "Pihak Pertama" : "Pihak Kedua");
  const sdtInner = (() => {
    if (o.copies === 2 && sigs.length > 1) {
      return capt(`Rangkap 1 dari 2 — dipegang ${holder(0)}${o.meterai ? " (bermeterai)" : ""}`) + sigTable(sigs, o.meterai ? 0 : -1) +
        `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
        capt(`Rangkap 2 dari 2 — dipegang ${holder(1)}${o.meterai ? " (bermeterai)" : ""}`) + sigTable(sigs, o.meterai ? 1 : -1);
    }
    return capt(`Rangkap tunggal${o.meterai ? " — bermeterai pada " + holder(0) : ""}`) + sigTable(sigs, o.meterai ? 0 : -1);
  })();
  return layoutSdt("ttd", "Tanda tangan", sdtInner + `<w:p/>`);
}

/** Muat logo ke dalam .docx (media + relasi + tipe konten). */
function embedLogo(zip: JSZip, rels: string, types: string, logo: { bytes: Buffer; ext: string }): { rels: string; types: string } {
  zip.file(`word/media/clm-logo.${logo.ext}`, logo.bytes);
  rels = rels.replace(new RegExp(`<Relationship [^>]*Id="${LOGO_REL}"[^>]*/>`, "g"), "");
  rels = rels.replace("</Relationships>", `<Relationship Id="${LOGO_REL}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/clm-logo.${logo.ext}"/></Relationships>`);
  if (!new RegExp(`<Default [^>]*Extension="${logo.ext}"`, "i").test(types)) {
    types = types.replace(/(<Types[^>]*>)/, `$1<Default Extension="${logo.ext}" ContentType="image/${logo.ext}"/>`);
  }
  return { rels, types };
}

export interface LayoutPatch {
  /** undefined = biarkan; null = hapus blok; objek = pasang/ganti blok. */
  kop?: { letterhead: LetterheadInput } | null;
  ttd?: { signatures: SignatureInput[]; options: SignatureOptions } | null;
}

/** Keadaan kop / logo / tanda tangan di dokumen (dibaca dari berkas, bukan dari pengaturan). */
export async function readLayout(docx: Buffer): Promise<{ kop: boolean; logo: boolean; ttd: boolean }> {
  const { doc } = await loadBody(docx);
  const find = (kind: string) => sdtOf(doc).find((x) => sdtTag(x) === LAYOUT_PREFIX + kind);
  const kop = find("kop");
  return { kop: !!kop, logo: !!kop && kop.getElementsByTagName("w:drawing").length > 0, ttd: !!find("ttd") };
}

/** Pasang / hapus / ganti kop surat dan blok tanda tangan pada dokumen yang sudah ada (isi lain tidak disentuh). */
export async function patchLayout(docx: Buffer, patch: LayoutPatch): Promise<Buffer> {
  const { zip, doc } = await loadBody(docx);
  const body = doc.getElementsByTagName("w:body")[0] as Element;
  const remove = (kind: string) => { for (const x of sdtOf(doc)) if (sdtTag(x) === LAYOUT_PREFIX + kind) x.parentNode?.removeChild(x); };
  const adopt = (xml: string): Node[] => {
    const frag = new DOMParser().parseFromString(`<w:body xmlns:w="${W_NS}">${xml}</w:body>`, "application/xml") as unknown as Document;
    return Array.from(frag.documentElement.childNodes).map((n) => doc.importNode(n, true));
  };
  if (patch.kop !== undefined) {
    remove("kop");
    if (patch.kop) {
      const l = patch.kop.letterhead;
      const info = l.logo ? imageInfo(l.logo.bytes, l.logo.mimeType) : null;
      if (info) {
        const relsEntry = zip.file("word/_rels/document.xml.rels");
        const typesEntry = zip.file("[Content_Types].xml");
        if (relsEntry && typesEntry) {
          const r = embedLogo(zip, await relsEntry.async("string"), await typesEntry.async("string"), { bytes: l.logo!.bytes, ext: info.ext });
          zip.file("word/_rels/document.xml.rels", r.rels);
          zip.file("[Content_Types].xml", r.types);
        }
      }
      const first = body.firstChild;
      for (const n of adopt(kopBlockXml(l, info))) body.insertBefore(n, first);
    }
  }
  if (patch.ttd !== undefined) {
    remove("ttd");
    if (patch.ttd) {
      const last = Array.from(body.childNodes).reverse().find((n) => n.nodeName === "w:sectPr") || null;
      for (const n of adopt(signatureBlockXml(patch.ttd.signatures, patch.ttd.options))) body.insertBefore(n, last);
    }
  }
  zip.file("word/document.xml", new XMLSerializer().serializeToString(doc as any));
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
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
  /** Kop surat di header halaman pertama. */
  letterhead?: LetterheadInput;
  /** Kolom tanda tangan di akhir dokumen. */
  signatures?: SignatureInput[];
  signatureOptions?: SignatureOptions;
}): Promise<Buffer> {
  const values = opts.values || {};
  const ctx: Ctx = {
    variable: (key, rPr) => varSdt(key, values[key] ?? `{{${key}}}`, rPr),
    numbering: [], skipped: [], defaultJc: "both",
  };
  const openingXml = opts.opening ? clauseContentToWordXml(opts.opening, ctx) : "";
  const closingXml = opts.closing ? clauseContentToWordXml(opts.closing, ctx) : "";
  const clauseXml = opts.clauses.map((c) => clauseSdt(c, ctx));
  const hasNumbering = ctx.numbering.length > 0;
  const logoInfo = opts.letterhead?.logo ? imageInfo(opts.letterhead.logo.bytes, opts.letterhead.logo.mimeType) : null;
  const kopXml = opts.letterhead ? kopBlockXml(opts.letterhead, logoInfo) : "";
  const body = [
    kopXml,
    ...(opts.docType
      ? [
          `<w:p><w:pPr><w:keepNext/><w:spacing w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${esc(opts.docType)}</w:t></w:r></w:p>`,
          `<w:p><w:pPr><w:keepNext/><w:spacing w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${esc(opts.title)}</w:t></w:r></w:p>`,
          ...(opts.numberLine ? [`<w:p><w:pPr><w:spacing w:after="240"/><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${esc(opts.numberLine)}</w:t></w:r></w:p>`] : []),
        ]
      : [`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">${esc(opts.title)}</w:t></w:r></w:p>`]),
    openingXml,
    closingXml,
    ...clauseXml,
    ...(opts.signatures && opts.signatures.length ? [signatureBlockXml(opts.signatures, opts.signatureOptions || { copies: 1, meterai: false })] : []),
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`,
  ].join("");
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${logoInfo ? `<Default Extension="${logoInfo.ext}" ContentType="image/${logoInfo.ext}"/>` : ""}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${hasNumbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : ""}</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${logoInfo ? `<Relationship Id="${LOGO_REL}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/clm-logo.${logoInfo.ext}"/>` : ""}${hasNumbering ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : ""}</Relationships>`);
  if (logoInfo) zip.file(`word/media/clm-logo.${logoInfo.ext}`, opts.letterhead!.logo!.bytes);
  if (hasNumbering) zip.file("word/numbering.xml", numberingXml(ctx.numbering));
  zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:lang w:val="id-ID"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
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
