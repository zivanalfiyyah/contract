// Generator surat pemutusan / pembatalan kontrak (.docx) — dokumen awal yang
// lalu disunting di editor Word (ONLYOFFICE). Murni (tanpa I/O) supaya mudah diuji.
import JSZip from "jszip";

export type LetterKind = "Terminated" | "Archived";
export interface LetterInput {
  kind: LetterKind;
  contractTitle: string;
  contractNumber: string;
  party1Name: string;
  party2Name: string;
  startDate?: string;
  endDate?: string;
  reasonType: string;
  reason: string;
  effectiveDate: string; // YYYY-MM-DD
  city?: string;
  today?: string; // YYYY-MM-DD (untuk uji)
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
export function tanggalPanjang(iso?: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
  if (!m) return iso || "-";
  return `${Number(m[3])} ${BULAN[Number(m[2]) - 1]} ${m[1]}`;
}

type Para = { text?: string; bold?: boolean; center?: boolean; after?: number; runs?: { text: string; bold?: boolean }[] };
const para = (p: Para) => {
  const runs = p.runs || [{ text: p.text || "", bold: p.bold }];
  const rx = runs.map((r) => `<w:r><w:rPr>${r.bold ? "<w:b/>" : ""}<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`).join("");
  return `<w:p><w:pPr>${p.center ? '<w:jc w:val="center"/>' : '<w:jc w:val="both"/>'}<w:spacing w:after="${p.after ?? 160}" w:line="276" w:lineRule="auto"/></w:pPr>${rx}</w:p>`;
};

export function letterParagraphs(i: LetterInput): Para[] {
  const terminate = i.kind === "Terminated";
  const title = terminate ? "SURAT PEMUTUSAN PERJANJIAN" : "SURAT PEMBATALAN PERJANJIAN";
  const periode = i.startDate && i.endDate ? ` (periode ${tanggalPanjang(i.startDate)} s/d ${tanggalPanjang(i.endDate)})` : "";
  const today = i.today || new Date(Date.now() + 7 * 3600_000).toISOString().slice(0, 10);
  const effect = tanggalPanjang(i.effectiveDate);
  return [
    { text: title, bold: true, center: true, after: 60 },
    { text: "Nomor: ……… / ……… / ………", center: true, after: 300 },
    { text: `${i.city || "Jakarta"}, ${tanggalPanjang(today)}`, after: 200 },
    { runs: [{ text: "Perihal: ", bold: true }, { text: `${terminate ? "Pemutusan" : "Pembatalan"} ${i.contractTitle}` }] },
    { text: "Kepada Yth.", after: 0 },
    { text: i.party2Name, bold: true, after: 0 },
    { text: "di tempat", after: 240 },
    { text: "Dengan hormat," },
    { text: terminate
        ? `Sehubungan dengan ${i.contractTitle} Nomor ${i.contractNumber}${periode} antara ${i.party1Name} dan ${i.party2Name}, dengan ini kami memberitahukan bahwa perjanjian tersebut diakhiri sebelum jangka waktunya berakhir.`
        : `Sehubungan dengan ${i.contractTitle} Nomor ${i.contractNumber}${periode} antara ${i.party1Name} dan ${i.party2Name}, dengan ini kami memberitahukan bahwa perjanjian tersebut dinyatakan dibatalkan dan tidak berlaku lagi.` },
    { runs: [{ text: "Dasar: ", bold: true }, { text: i.reasonType || "-" }] },
    { runs: [{ text: "Keterangan: ", bold: true }, { text: i.reason || "-" }] },
    { text: `${terminate ? "Pengakhiran" : "Pembatalan"} ini berlaku efektif pada tanggal ${effect}. Seluruh hak dan kewajiban yang telah timbul sebelum tanggal efektif tetap diselesaikan sesuai ketentuan perjanjian.` },
    { text: "[Sesuaikan dengan ketentuan pemutusan/pembatalan dan masa pemberitahuan pada klausul perjanjian.]", after: 240 },
    { text: "Demikian kami sampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.", after: 400 },
    { text: "Hormat kami,", after: 0 },
    { text: i.party1Name, bold: true, after: 900 },
    { text: "(Nama & jabatan penandatangan)", after: 0 },
  ];
}

export async function buildTerminationLetter(i: LetterInput): Promise<Buffer> {
  const body = letterParagraphs(i).map(para).join("");
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file("word/document.xml", doc);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
