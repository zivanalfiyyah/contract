// Uji dokumen Word Addendum: urutan blok (pembuka → recital → pasal → penutup → tanda tangan) & variabel tetap bisa dibaca.
// Jalankan: npx tsx scripts/addendum-word-test.ts
import assert from "node:assert/strict";
import JSZip from "jszip";
import { buildTemplateDocx, listVariables, setVariables } from "../template-docx.js";

const recital = 'Bahwa PARA PIHAK telah membuat dan menandatangani Perjanjian Nomor **001/PKS/2026** tanggal **1 Januari 2026** (selanjutnya disebut "Perjanjian").';
const tail = "Demikian Addendum ini dibuat dan ditandatangani oleh PARA PIHAK.";

async function text(buf: Buffer) {
  const z = await JSZip.loadAsync(buf);
  const xml = await z.file("word/document.xml")!.async("string");
  return xml.replace(/<w:p[ >]/g, "\n<w:p ").replace(/<[^>]+>/g, "");
}

(async () => {
  const docx = await buildTemplateDocx({
    title: "Addendum Pertama X", docType: "ADDENDUM / AMANDEMEN", numberLine: "NOMOR: ADD-001/2026",
    opening: "Pada hari ini, kami yang bertandatangan di bawah ini: {{Party2Name}}",
    recital, tail,
    clauses: [
      { id: "auto-jangka-waktu", title: "Perubahan Jangka Waktu Perjanjian", content: "Perjanjian berlaku sampai 2027.", isMandatory: false, isProtected: false },
      { id: "c2", title: "Pasal Lain", content: "Isi pasal lain.", isMandatory: false, isProtected: false },
    ],
    signatures: [{ role: "PIHAK PERTAMA", name: "PT A" }, { role: "PIHAK KEDUA", name: "PT B" }],
    signatureOptions: { copies: 1, meterai: false },
    values: { Party2Name: "PT Maju" },
  });
  const t = await text(docx);
  const pos = (s: string) => { const i = t.indexOf(s); assert.ok(i >= 0, "tidak ada: " + s); return i; };
  const order = ["ADDENDUM / AMANDEMEN", "PT Maju", "001/PKS/2026", "Perubahan Jangka Waktu", "Pasal Lain", "Demikian Addendum", "PIHAK KEDUA"].map(pos);
  assert.deepEqual(order, [...order].sort((a, b) => a - b), "urutan blok dokumen addendum");
  assert.ok(!t.includes("**"), "penanda tebal tidak boleh muncul mentah");
  console.log("  ok  urutan blok & teks");
  const vars = await listVariables(docx);
  assert.equal(vars.Party2Name, "PT Maju");
  const r = await setVariables(docx, { Party2Name: "PT Baru" });
  assert.equal(r.changed, 1);
  assert.ok((await text(r.docx)).includes("PT Baru"));
  console.log("  ok  variabel dapat dibaca & diubah (Panel CLM tetap bekerja)");
  const plain = await buildTemplateDocx({ title: "K", clauses: [{ id: "c", title: "P", content: "x", isMandatory: false, isProtected: false }], opening: "Buka" });
  assert.ok(!(await text(plain)).includes("Demikian"), "kontrak biasa tidak berubah");
  console.log("  ok  kontrak biasa (tanpa recital/tail) tidak berubah");
  console.log("ADDENDUM WORD CHECKS PASSED");
})().catch((e) => { console.error(e); process.exit(1); });
