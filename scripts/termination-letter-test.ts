import JSZip from "jszip";
import assert from "node:assert/strict";
import { buildTerminationLetter, tanggalPanjang, letterParagraphs } from "../termination-letter.js";
import { readFileSync } from "node:fs";

const base = { contractTitle: "PKS <Test> & Co", contractNumber: "001/PKS/XI/2026", party1Name: "PT A", party2Name: "PT B", startDate: "2026-01-01", endDate: "2027-01-01", reasonType: "Wanprestasi / pelanggaran", reason: "Terlambat bayar", effectiveDate: "2026-12-31", today: "2026-11-12" };
(async () => {
  assert.equal(tanggalPanjang("2026-12-31"), "31 Desember 2026");
  for (const kind of ["Terminated", "Archived"] as const) {
    const buf = await buildTerminationLetter({ ...base, kind });
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file("word/document.xml")!.async("string");
    assert.ok(xml.includes("&lt;Test&gt; &amp; Co"), "XML di-escape");
    assert.ok(xml.includes("31 Desember 2026") && xml.includes("Terlambat bayar") && xml.includes("PT B"));
    assert.ok(xml.includes(kind === "Terminated" ? "SURAT PEMUTUSAN PERJANJIAN" : "SURAT PEMBATALAN PERJANJIAN"));
    assert.ok(zip.file("[Content_Types].xml") && zip.file("_rels/.rels"));
    assert.ok(letterParagraphs({ ...base, kind }).length > 10);
  }
  const srv = readFileSync(new URL("../server.ts", import.meta.url), "utf8");
  assert.ok(srv.includes('"/api/contracts/:id/termination/decision"') && srv.includes("applyDueTerminations(db)"));
  assert.ok(/Aktif: \[\],\n\s+TidakAktif: \[\],/.test(srv), "PUT tidak lagi boleh menutup kontrak");
  console.log("ok - surat pengakhiran + guard server");
})().catch((e) => { console.error(e); process.exit(1); });
