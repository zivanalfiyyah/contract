// Uji pemisahan markup internal ↔ eksternal. Jalankan: npx tsx scripts/review-isolation-test.ts
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { buildExternalCopy, mergeExternalMarkup, assertNoInternalTraces } from "../review-isolation.js";
import { readChangeHistory, acceptAllChanges, originalTextOf } from "../review-changes.js";

let failed = 0;
const ok = (cond: boolean, msg: string) => { if (!cond) { failed++; console.error("GAGAL:", msg); } else console.log("ok   :", msg); };

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const IN = 'w:author="Budi (Internal)" w:date="2026-11-01T08:00:00Z"';
const EX = 'w:author="Sari (Pihak Eksternal)" w:date="2026-11-10T08:00:00Z"';
const r = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const dr = (t: string) => `<w:r><w:delText xml:space="preserve">${t}</w:delText></w:r>`;

async function docx(body: string, comments?: string): Promise<Buffer> {
  const z = new JSZip();
  z.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${comments ? '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' : ""}</Types>`);
  z.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  z.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${comments ? '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' : ""}</Relationships>`);
  z.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`);
  if (comments) z.file("word/comments.xml", `<?xml version="1.0" encoding="UTF-8"?><w:comments ${NS}>${comments}</w:comments>`);
  return z.generateAsync({ type: "nodebuffer" });
}
async function bodyOf(buf: Buffer, part = "word/document.xml") { return (await (await JSZip.loadAsync(buf)).file(part)?.async("string")) || ""; }
const text = (xml: string) => xml.replace(/<w:delText[^>]*>[^<]*<\/w:delText>/g, "").replace(/<[^>]+>/g, "");

(async () => {
  // MASTER: P1 sisipan+hapusan internal, P2 dengan komentar internal, P3 paragraf dihapus internal, P4 penutup.
  const masterBody =
    `<w:p>${r("Pasal 1 ")}<w:ins w:id="1" ${IN}>${r("pihak pertama ")}</w:ins><w:del w:id="2" ${IN}>${dr("lama ")}</w:del>${r("bayar")}</w:p>` +
    `<w:p><w:commentRangeStart w:id="0"/>${r("Pasal 2 harga seratus juta")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
    `<w:p><w:pPr><w:rPr><w:del w:id="3" ${IN}/></w:rPr></w:pPr><w:del w:id="4" ${IN}>${dr("Pasal dihapus")}</w:del></w:p>` +
    `<w:p>${r("Penutup")}</w:p>`;
  const internalComment = `<w:comment w:id="0" w:author="Budi (Internal)" w:date="2026-11-01T09:00:00Z" w:initials="B"><w:p><w:r><w:t>RAHASIA internal: cek harga</w:t></w:r></w:p></w:comment>`;
  const master = await docx(masterBody, internalComment);

  // 1. Salinan eksternal bersih.
  const copy = await buildExternalCopy(master);
  await assertNoInternalTraces(copy);
  const cz = await JSZip.loadAsync(copy);
  ok(!cz.file("word/comments.xml"), "salinan eksternal tidak memuat comments.xml");
  const cdoc = await bodyOf(copy);
  ok(!/Budi|RAHASIA|w:ins|w:del\b|commentReference/.test(cdoc) && !/RAHASIA/.test(await bodyOf(copy, "[Content_Types].xml")), "salinan tidak membocorkan markup/komentar internal");
  ok(text(cdoc) === "Pasal 1 pihak pertama bayarPasal 2 harga seratus jutaPenutup", "teks salinan = hasil 'terima semua'");
  const hist = await readChangeHistory(copy);
  ok(hist.changes.length === 0 && hist.comments.length === 0, "riwayat markup pada salinan kosong");

  // 2. Eksternal memberi markup di salinan: hapus 'bayar'+sisip, ganti 'seratus', komentar di 'Penutup', paragraf baru.
  const returnedBody =
    `<w:p>${r("Pasal 1 pihak pertama ")}<w:del w:id="21" ${EX}>${dr("bayar")}</w:del><w:ins w:id="22" ${EX}>${r("membayar tunai")}</w:ins></w:p>` +
    `<w:p>${r("Pasal 2 harga ")}<w:del w:id="23" ${EX}>${dr("seratus")}</w:del><w:ins w:id="24" ${EX}>${r("dua ratus")}</w:ins>${r(" juta")}</w:p>` +
    `<w:p><w:commentRangeStart w:id="0"/>${r("Penutup")}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
    `<w:p><w:pPr><w:rPr><w:ins w:id="25" ${EX}/></w:rPr></w:pPr><w:ins w:id="26" ${EX}>${r("Pasal 5 usulan eksternal")}</w:ins></w:p>`;
  const extComment = `<w:comment w:id="0" w:author="Sari (Pihak Eksternal)" w:date="2026-11-10T09:00:00Z" w:initials="S"><w:p><w:r><w:t>Mohon klarifikasi penutup</w:t></w:r></w:p></w:comment>`;
  const returned = await docx(returnedBody, extComment);

  const res = await mergeExternalMarkup({ master, baseline: copy, returned });
  ok(res.ok, "penggabungan berhasil" + (res.ok ? "" : ` (${res.reason})`));
  if (!res.ok) { process.exit(1); }
  const mdoc = await bodyOf(res.buf);
  const mc = await bodyOf(res.buf, "word/comments.xml");
  ok(/Budi \(Internal\)/.test(mdoc) && /pihak pertama/.test(mdoc) && /lama/.test(mdoc), "markup internal (sisipan & hapusan) tetap ada di master");
  ok(/Pasal dihapus/.test(mdoc), "paragraf yang dihapus internal tetap tercatat");
  ok(/RAHASIA internal/.test(mc), "komentar internal tetap ada di master");
  ok(/Mohon klarifikasi penutup/.test(mc), "komentar eksternal ditambahkan ke master");
  const ids = [...mc.matchAll(/<w:comment\b[^>]*\bw:id="(\d+)"/g)].map((m) => m[1]);
  ok(new Set(ids).size === ids.length && ids.length === 2, "id komentar unik (internal 0, eksternal dinomori ulang)");
  const mh = await readChangeHistory(res.buf);
  const authors = new Set([...mh.changes.map((c) => c.author), ...mh.comments.map((c) => c.author)]);
  ok(authors.has("Budi (Internal)") && authors.has("Sari (Pihak Eksternal)"), "riwayat memuat markup internal DAN eksternal dalam satu berkas");
  ok(mh.changes.some((c) => c.author.includes("Eksternal") && c.kind === "delete" && c.text.includes("seratus")), "coretan eksternal 'seratus' tercatat");
  ok(mh.changes.some((c) => c.author.includes("Eksternal") && c.kind === "insert" && c.text.includes("Pasal 5")), "paragraf baru eksternal tercatat");
  const finalText = text(await bodyOf(await acceptAllChanges(res.buf)));
  ok(finalText === "Pasal 1 pihak pertama membayar tunaiPasal 2 harga dua ratus jutaPenutupPasal 5 usulan eksternal", "terima semua = teks akhir yang benar: " + finalText);
  ok((await originalTextOf(res.buf)) === (await originalTextOf(master)), "tampilan Original master tidak berubah");

  // 3. Eksternal menyisip di tengah sisipan internal (harus memecah w:ins internal, bukan menumpuk).
  const mid = await docx(`<w:p>${r("A ")}<w:ins w:id="1" ${IN}>${r("sisip internal panjang ")}</w:ins>${r("Z")}</w:p>`);
  const midCopy = await buildExternalCopy(mid);
  const midRet = await docx(`<w:p>${r("A sisip ")}<w:ins w:id="9" ${EX}>${r("EKS ")}</w:ins>${r("internal panjang Z")}</w:p>`);
  const midRes = await mergeExternalMarkup({ master: mid, baseline: midCopy, returned: midRet });
  ok(midRes.ok, "sisip di tengah sisipan internal berhasil" + (midRes.ok ? "" : ` (${midRes.reason})`));
  if (midRes.ok) {
    const x = await bodyOf(midRes.buf);
    ok(!/<w:ins\b[^>]*>(?:(?!<\/w:ins>).)*<w:ins\b/s.test(x), "tidak ada w:ins bersarang");
    ok(text(await bodyOf(await acceptAllChanges(midRes.buf))) === "A sisip EKS internal panjang Z", "teks akhir benar setelah sisip di tengah");
  }

  // 4. Hapus teks yang sebelumnya disisipkan internal -> w:ins > w:del.
  const delRet = await docx(`<w:p>${r("A ")}<w:del w:id="9" ${EX}>${dr("sisip ")}</w:del>${r("internal panjang Z")}</w:p>`);
  const delRes = await mergeExternalMarkup({ master: mid, baseline: midCopy, returned: delRet });
  ok(delRes.ok, "hapus teks sisipan internal berhasil" + (delRes.ok ? "" : ` (${delRes.reason})`));
  if (delRes.ok) ok(text(await bodyOf(await acceptAllChanges(delRes.buf))) === "A internal panjang Z", "teks akhir benar setelah hapus sebagian sisipan internal");

  // 5. Fail-safe: perubahan format belum didukung -> ditolak, master tidak diubah.
  const fmtRet = await docx(`<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="5" ${EX}><w:rPr/></w:rPrChange></w:rPr><w:t>A sisip internal panjang Z</w:t></w:r></w:p>`);
  const fmtRes = await mergeExternalMarkup({ master: mid, baseline: midCopy, returned: fmtRet });
  ok(!fmtRes.ok, "perubahan format ditolak aman (fallback)");

  // 6. Fail-safe: paragraf dipecah.
  const splitRet = await docx(`<w:p><w:pPr><w:rPr><w:ins w:id="7" ${EX}/></w:rPr></w:pPr>${r("A sisip ")}</w:p><w:p>${r("internal panjang Z")}</w:p>`);
  const splitRes = await mergeExternalMarkup({ master: mid, baseline: midCopy, returned: splitRet });
  ok(!splitRes.ok, "paragraf dipecah ditolak aman (fallback)");

  // 7. Fail-safe: berkas eksternal yang teksnya diubah tanpa Track Changes.
  const sneakyRet = await docx(`<w:p>${r("A sisip internal panjang TANPA TRACK")}</w:p>`);
  const sneakyRes = await mergeExternalMarkup({ master: mid, baseline: midCopy, returned: sneakyRet });
  ok(!sneakyRes.ok, "perubahan teks tanpa Track Changes ditolak");

  console.log(failed ? `\n${failed} uji GAGAL` : "\nSemua uji lulus");
  process.exit(failed ? 1 : 0);
})();
