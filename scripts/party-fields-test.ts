// Uji "Data Pihak & Kontrak" (party-fields.ts) + integrasi variabel Word (template-docx.ts).
// Jalankan: npx tsx scripts/party-fields-test.ts
//
// Yang dijaga:
//  1. PARITAS dgn modal "Edit Data Pihak & Kontrak" di src/App.tsx — field, label, tipe input, pesan validasi.
//     Bila seseorang menambah/menghapus field di modal tanpa menyamakan panel, uji ini GAGAL.
//  2. Validasi & penguncian Nomor Dokumen.
//  3. Rekonsiliasi lama TIDAK berubah perilakunya (dibandingkan dgn salinan implementasi lama).
//  4. Putaran dua arah (panel -> DB -> dokumen, dokumen -> DB, aplikasi utama -> dokumen) pada .docx sungguhan.
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import type { Contract } from "../src/types";
import {
  PARTY_FIELDS, PARTY_TEXT_KEYS, PARTY_LOGO_KEYS, PARTY_WORD_KEYS, partyFormOf, partyRevOf, partyEditability,
  buildPartyView, normalizePartyPatch, applyPartyPatch, partyWordValues, reconcileCore, englishInvalidationBetween,
  isPlaceholderValue, parseIdDate, ISO_DATE, formatShortDateId, partyFieldLabel, type PartyCompany,
} from "../party-fields";
import { listVariables, setVariables, refreshEnglishVars, valueToEnglish } from "../template-docx";

const here = path.dirname(fileURLToPath(import.meta.url));
let passed = 0;
const ok = (name: string) => { passed++; console.log(`  ok  ${name}`); };
const cloneC = <T,>(x: T): T => JSON.parse(JSON.stringify(x));

const company: PartyCompany = { name: "PT Semesta Digital Terpadu", address: "Jakarta Selatan", repTitle: "Direktur" };
const baseContract = (): Contract => ({
  id: "c1", tenantId: "t-01", contractNumber: "174/PKS-KUY-MEDEX/GKM/XI/2026", title: "CEK HAK AKSES ANTAR ENTITAS",
  docType: "Perjanjian Kerjasama", startDate: "2026-11-10", endDate: "2028-11-22", status: "Draft",
  party1Name: "PT Semesta Digital Terpadu", party2Name: "Xaxa", party2Type: "Vendor", variables: {}, parties: [],
  createdById: "u-1",
} as unknown as Contract);

// ---------------------------------------------------------------------------
// 1. Paritas dgn modal aplikasi utama
// ---------------------------------------------------------------------------
{
  const app = fs.readFileSync(path.join(here, "..", "src", "App.tsx"), "utf8");
  const from = app.indexOf("{editPartyModalContract && (");
  const to = app.indexOf("{renewSource && (", from);
  assert.ok(from > 0 && to > from, "blok modal Edit Data Pihak & Kontrak tidak ditemukan di App.tsx");
  const modal = app.slice(from, to);

  // (a) himpunan field modal == himpunan field skema
  const inModal = new Set(Array.from(modal.matchAll(/editPartyForm\.(\w+)/g)).map((m) => m[1]));
  // Logo = tiga nilai satu kesatuan: modal membaca `party2LogoUrl`, sedangkan key & mimeType diisi handler
  // unggah / tombol hapus (properti objek, bukan `editPartyForm.X`) — dicek eksplisit di bawah.
  const inSchema = new Set<string>([...PARTY_TEXT_KEYS, "party2LogoUrl"]);
  assert.ok(app.includes("party2LogoUrl: data.url, party2LogoKey: data.key, party2LogoMimeType: data.mimeType"), "handler unggah logo berubah");
  assert.ok(modal.includes('party2LogoUrl: "", party2LogoKey: "", party2LogoMimeType: ""'), "tombol hapus logo berubah");
  assert.deepEqual([...PARTY_LOGO_KEYS], ["party2LogoUrl", "party2LogoKey", "party2LogoMimeType"]);
  const missingInPanel = [...inModal].filter((k) => !inSchema.has(k));
  const missingInModal = [...inSchema].filter((k) => !inModal.has(k));
  assert.deepEqual(missingInPanel, [], `Field ada di modal tapi TIDAK di Panel CLM: ${missingInPanel.join(", ")}`);
  assert.deepEqual(missingInModal, [], `Field ada di Panel CLM tapi tidak di modal: ${missingInModal.join(", ")}`);
  ok(`himpunan field modal == skema panel (${inSchema.size} field)`);

  // (b) tipe input tiap field sama dgn modal
  const typeInModal = (key: string): string => {
    const i = modal.indexOf(`value={editPartyForm.${key}}`);
    assert.ok(i > 0, `value untuk ${key} tidak ditemukan`);
    const head = modal.slice(Math.max(0, i - 260), i);
    const lastTag = Math.max(head.lastIndexOf("<input"), head.lastIndexOf("<select"));
    const tag = head.slice(lastTag);
    if (tag.startsWith("<select")) return "select";
    return /type="date"/.test(tag) ? "date" : "text";
  };
  for (const f of PARTY_FIELDS) assert.equal(f.type, typeInModal(f.key), `tipe input ${f.key}: panel=${f.type}, modal=${typeInModal(f.key)}`);
  ok("tipe input (text/date/select) tiap field sama dgn modal");

  // (c) label & placeholder yang dipakai modal ada di skema
  for (const lbl of ["Judul Dokumen", "Jenis Dokumen", 'Tanggal ("pada hari ini, tanggal...")', "Tanggal Berakhir Kontrak", "Nomor Dokumen", "Logo Pihak Kedua"]) {
    assert.ok(modal.includes(lbl), `label modal berubah: ${lbl}`);
  }
  assert.equal(PARTY_FIELDS.find((f) => f.key === "title")!.label, "Judul Dokumen");
  assert.equal(PARTY_FIELDS.find((f) => f.key === "docType")!.label, "Jenis Dokumen");
  assert.equal(PARTY_FIELDS.find((f) => f.key === "startDate")!.label, 'Tanggal ("pada hari ini, tanggal...")');
  assert.equal(PARTY_FIELDS.find((f) => f.key === "endDate")!.label, "Tanggal Berakhir Kontrak");
  for (const ph of ["mis. Surat Perjanjian Kerjasama", "Label ID (mis. NPWP)", "Label ID (mis. NIK)", "Nomor identitas", 'placeholder="Alamat"', 'placeholder="Jabatan"', 'placeholder="Nama"']) {
    assert.ok(modal.includes(ph), `placeholder modal berubah: ${ph}`);
  }
  const view = buildPartyView(baseContract(), company, ["Vendor", "Karyawan"]);
  const ph = (k: string) => view.groups.flatMap((g) => g.fields).find((f) => f.key === k)!.placeholder;
  assert.equal(ph("docType"), "mis. Surat Perjanjian Kerjasama");
  assert.equal(ph("party1IdLabel"), "Label ID (mis. NPWP)");
  assert.equal(ph("party2IdLabel"), "Label ID (mis. NIK)");
  assert.equal(ph("party1IdNumber"), "Nomor identitas");
  assert.equal(ph("party2Name"), "Nama");
  assert.equal(ph("party2Address"), "Alamat");
  assert.equal(ph("party2Position"), "Jabatan");
  assert.equal(ph("party1Address"), 'Alamat (kosong = "Jakarta Selatan")');
  assert.equal(ph("party1Position"), 'Jabatan Penandatanganan (kosong = "Direktur")');
  assert.ok(modal.includes("Alamat (kosong = "), "placeholder dinamis Alamat Pihak Pertama berubah di modal");
  assert.ok(modal.includes("Jabatan Penandatanganan (kosong = "), "placeholder dinamis Jabatan berubah di modal");
  ok("label & placeholder (termasuk yang dinamis) sama dgn modal");

  // (d) pesan validasi sama dgn modal
  assert.ok(modal.length > 0 && app.includes("Judul dokumen tidak boleh kosong") && app.includes("Nama Pihak Kedua tidak boleh kosong"));
  const bad1 = normalizePartyPatch(baseContract(), { title: "  " }, { partyTypes: ["Vendor"] });
  const bad2 = normalizePartyPatch(baseContract(), { party2Name: "" }, { partyTypes: ["Vendor"] });
  assert.ok(bad1.ok === false && bad1.error === "Judul dokumen tidak boleh kosong");
  assert.ok(bad2.ok === false && bad2.error === "Nama Pihak Kedua tidak boleh kosong");
  ok("pesan validasi wajib-isi sama persis dgn modal");

  // (e) pengelompokan & urutan: Informasi Kontrak, Pihak Pertama, Pihak Kedua (urutan modal)
  assert.deepEqual(view.groups.map((g) => g.id), ["kontrak", "pihak1", "pihak2"]);
  assert.equal(view.groups[0].title, "Informasi Kontrak");
  assert.equal(view.groups[1].title, "Pihak Pertama (PT Semesta Digital Terpadu)");
  assert.equal(view.groups[2].title, "Pihak Kedua");
  assert.deepEqual(view.groups[0].fields.map((f) => f.key), ["title", "docType", "startDate", "endDate"]);
  assert.deepEqual(view.groups[1].fields.map((f) => f.key), ["party1Address", "party1Position", "party1IdLabel", "party1IdNumber"]);
  assert.deepEqual(view.groups[2].fields.map((f) => f.key), ["party2Name", "party2Type", "party2Address", "party2Position", "party2IdLabel", "party2IdNumber"]);
  assert.equal(view.locked.value, "174/PKS-KUY-MEDEX/GKM/XI/2026");
  ok("pengelompokan Informasi Kontrak / Pihak Pertama / Pihak Kedua + Nomor Dokumen terkunci");

  // (f) dropdown tipe: nilai saat ini tetap tampil walau tidak ada di daftar master (sama dgn modal)
  const c = baseContract(); c.party2Type = "Rekanan Lama";
  const sel = buildPartyView(c, company, ["Vendor", "Karyawan"]).groups[2].fields.find((f) => f.key === "party2Type")!;
  assert.deepEqual(sel.options, ["Rekanan Lama", "Vendor", "Karyawan"]);
  assert.deepEqual(buildPartyView(baseContract(), company, ["Vendor", "Karyawan"]).groups[2].fields.find((f) => f.key === "party2Type")!.options, ["Vendor", "Karyawan"]);
  ok("dropdown Tipe Pihak Kedua = daftar master (+ nilai lama bila di luar daftar)");
}

// ---------------------------------------------------------------------------
// 2. Validasi, penguncian, normalisasi
// ---------------------------------------------------------------------------
{
  const t = ["Vendor", "Karyawan"];
  // Nomor Dokumen terkunci
  const lock = normalizePartyPatch(baseContract(), { contractNumber: "999/XX" }, { partyTypes: t });
  assert.ok(lock.ok === false && lock.field === "contractNumber" && /terkunci/.test(lock.error));
  const same = normalizePartyPatch(baseContract(), { contractNumber: "174/PKS-KUY-MEDEX/GKM/XI/2026", title: "Baru" }, { partyTypes: t });
  assert.ok(same.ok === true && same.changed.join() === "title" && !("contractNumber" in same.next));
  ok("Nomor Dokumen: perubahan ditolak, nilai sama diabaikan & tidak pernah masuk `next`");

  // tanggal
  for (const bad of ["", "10/11/2026", "2026-02-31", "2026-13-01"]) {
    const r = normalizePartyPatch(baseContract(), { endDate: bad }, { partyTypes: t });
    assert.ok(r.ok === false && r.field === "endDate", `tanggal "${bad}" harus ditolak`);
  }
  assert.ok(normalizePartyPatch(baseContract(), { endDate: "2029-02-28" }, { partyTypes: t }).ok);
  ok("tanggal: kosong / format salah / tanggal fiktif ditolak");

  // tipe pihak kedua
  assert.ok(normalizePartyPatch(baseContract(), { party2Type: "Hacker" }, { partyTypes: t }).ok === false);
  assert.ok(normalizePartyPatch(baseContract(), { party2Type: "Karyawan" }, { partyTypes: t }).ok === true);
  const keep = normalizePartyPatch(baseContract(), { party2Type: "  " }, { partyTypes: t });
  assert.ok(keep.ok === true && !("party2Type" in keep.next), "kosong = tetap nilai lama (sama dgn modal)");
  const legacy = baseContract(); legacy.party2Type = "Rekanan Lama";
  assert.ok(normalizePartyPatch(legacy, { party2Type: "Rekanan Lama" }, { partyTypes: t }).ok === true);
  ok("Tipe Pihak Kedua: harus dari daftar (atau nilai lama); kosong = tidak berubah");

  // trim + kosong = undefined untuk bidang opsional
  const c = baseContract(); c.party2Position = "Manajer";
  const r = normalizePartyPatch(c, { party2Position: "   ", party2Address: "  Jl. Mawar 1  ", docType: "" }, { partyTypes: t });
  assert.ok(r.ok === true);
  if (r.ok) {
    assert.equal(r.next.party2Position, undefined);
    assert.equal(r.next.party2Address, "Jl. Mawar 1");
    assert.equal(r.next.docType, undefined);
    applyPartyPatch(c, r);
    assert.equal(c.party2Position, undefined);
    assert.equal(c.party2Address, "Jl. Mawar 1");
  }
  ok("trim; bidang opsional kosong disimpan undefined (sama dgn modal)");

  // field tak dikenal / tipe salah / terlalu panjang
  assert.ok(normalizePartyPatch(baseContract(), { status: "Aktif" }, { partyTypes: t }).ok === false);
  assert.ok(normalizePartyPatch(baseContract(), { title: 123 as any }, { partyTypes: t }).ok === false);
  assert.ok(normalizePartyPatch(baseContract(), { title: "x".repeat(1001) }, { partyTypes: t }).ok === false);
  assert.ok(normalizePartyPatch(baseContract(), { party2LogoMimeType: "image/gif" }, { partyTypes: t }).ok === false);
  ok("field tak dikenal (mis. status), tipe salah, terlalu panjang, logo bukan PNG/JPEG ditolak");

  // konflik: nilai di server sudah berubah dari yang dilihat pengguna
  const cur = baseContract(); cur.party2Address = "Alamat dari aplikasi utama";
  const conf = normalizePartyPatch(cur, { party2Address: "Alamat dari panel" }, { partyTypes: t, base: { party2Address: "" } });
  assert.ok(conf.ok === false && conf.status === 409 && conf.field === "party2Address");
  const noConf = normalizePartyPatch(cur, { party2Address: "Alamat dari panel" }, { partyTypes: t, base: { party2Address: "Alamat dari aplikasi utama" } });
  assert.ok(noConf.ok === true);
  ok("konflik penyimpanan (perubahan di aplikasi utama saat panel terbuka) -> 409, bukan menimpa diam-diam");

  // terjemahan Inggris basi
  const base = baseContract();
  const t1 = normalizePartyPatch(base, { title: "Judul Baru" }, { partyTypes: t });
  assert.ok(t1.ok && t1.invalidate.title && !t1.invalidate.docType && !t1.invalidate.preamble);
  const t2 = normalizePartyPatch(base, { docType: "MoU" }, { partyTypes: t });
  assert.ok(t2.ok && !t2.invalidate.title && t2.invalidate.docType && !t2.invalidate.preamble);
  const t3 = normalizePartyPatch(base, { party1IdNumber: "01.234" }, { partyTypes: t });
  assert.ok(t3.ok && !t3.invalidate.title && !t3.invalidate.docType && t3.invalidate.preamble);
  const t4 = normalizePartyPatch(base, { endDate: "2030-01-01", party2Type: "Karyawan", party2LogoUrl: "/u/x.png", party2LogoKey: "k", party2LogoMimeType: "image/png" }, { partyTypes: t });
  assert.ok(t4.ok && !t4.invalidate.title && !t4.invalidate.docType && !t4.invalidate.preamble, "tanggal berakhir/tipe/logo tidak menyentuh narasi (sama dgn modal)");
  const e = baseContract(); e.titleEn = "Old"; e.docTypeEn = "Old"; e.preambleEn = "Old"; e.preambleEnOriginal = "Old"; e.translationWarnings = ["w"];
  const t5 = normalizePartyPatch(e, { title: "B", party2Name: "Yyy" }, { partyTypes: t });
  assert.ok(t5.ok === true);
  if (t5.ok) { applyPartyPatch(e, t5); assert.equal(e.titleEn, undefined); assert.equal(e.docTypeEn, "Old"); assert.equal(e.preambleEn, undefined); assert.equal(e.translationWarnings, undefined); }
  ok("terjemahan Inggris yang basi dibuang HANYA untuk bagian yang berubah (sama dgn modal)");

  // tidak ada perubahan
  const noop = normalizePartyPatch(baseContract(), partyFormOf(baseContract()) as any, { partyTypes: t });
  assert.ok(noop.ok === true && noop.changed.length === 0);
  ok("mengirim data yang sama = tidak ada perubahan");

  // label untuk audit
  assert.equal(partyFieldLabel("party2Address"), "Pihak Kedua — Alamat");
  assert.equal(partyFieldLabel("title"), "Judul Dokumen");
  ok("label audit mudah dibaca (tanpa nilai pribadi)");
}

// ---------------------------------------------------------------------------
// 3. Boleh diubah?
// ---------------------------------------------------------------------------
{
  const c = baseContract();
  assert.equal(partyEditability(c, { id: "u-1", role: "legal" }, true).editable, true);
  assert.equal(partyEditability({ ...c, status: "OnReview" } as Contract, { id: "u-1", role: "legal" }, true).editable, false);
  assert.equal(partyEditability(c, { id: "u-1", role: "viewer" }, false).editable, false);
  const rev = { ...c, revisionStage: "internal" } as Contract;
  assert.equal(partyEditability(rev, { id: "u-1", role: "legal" }, true).editable, true, "pembuat boleh saat revisi");
  assert.equal(partyEditability(rev, { id: "u-9", role: "legal" }, true).editable, false, "bukan pembuat tidak boleh saat revisi");
  assert.equal(partyEditability(rev, { id: "u-9", role: "super_admin" }, true).editable, true, "Super Admin boleh saat revisi");
  assert.equal(partyEditability({ ...rev, createdById: undefined } as Contract, { id: "u-9", role: "legal" }, true).editable, true, "tanpa pemilik tercatat: ikut izin edit");
  ok("aturan boleh-diubah = isContractEditable (Draft + izin + pemilik saat revisi)");
}

// ---------------------------------------------------------------------------
// 4. Rekonsiliasi: perilaku lama TIDAK berubah (bandingkan dgn salinan implementasi lama)
// ---------------------------------------------------------------------------
function legacyReconcile(c: Contract, inDoc: Record<string, string>, want: Record<string, string>, snap: Record<string, string>) {
  const push: Record<string, string> = {}; const pushed: string[] = [], pulled: string[] = [];
  for (const [key, docVal] of Object.entries(inDoc)) {
    const sys = want[key];
    if (sys === undefined) continue;
    if (snap[key] === undefined) { snap[key] = docVal; continue; }
    if (sys !== snap[key]) {
      if (sys !== docVal) { push[key] = sys; pushed.push(key); }
      snap[key] = sys;
    } else if (docVal !== snap[key]) {
      const identity: Record<string, () => void> = {
        Party2Name: () => { if (docVal.trim()) c.party2Name = docVal.trim(); },
        Party2Address: () => { if (docVal.trim() && !docVal.startsWith("[")) c.party2Address = docVal.trim(); },
      };
      const iso = parseIdDate(docVal);
      if (key === "StartDate" && iso) { c.startDate = iso; pulled.push(key); snap[key] = docVal; }
      else if (key === "EndDate" && iso) { c.endDate = iso; if (typeof c.variables?.EndDate === "string") (c.variables as any).EndDate = iso; pulled.push(key); snap[key] = docVal; }
      else if (identity[key]) { identity[key](); pulled.push(key); snap[key] = docVal; }
      else if (typeof c.variables?.[key] === "string" && ISO_DATE.test(c.variables[key] as string)) {
        if (iso) { (c.variables as any)[key] = iso; pulled.push(key); }
        snap[key] = docVal;
      } else if (typeof c.variables?.[key] === "string" && !(key in { StartDate: 1, EndDate: 1, Party1Name: 1, Party1Address: 1, Party1Representative: 1, Party1RepTitle: 1 })) {
        (c.variables as any)[key] = docVal; pulled.push(key); snap[key] = docVal;
      } else snap[key] = docVal;
    }
  }
  return { push, pushed, pulled };
}
{
  // Pembangkit tetap -> skenario acak tapi selalu sama.
  let seed = 20261110;
  const rnd = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);
  const pick = <T,>(a: T[]) => a[Math.floor(rnd() * a.length)];
  // Kunci lama saja (Party1Address sengaja dikecualikan: kini memang BOLEH ditarik dari dokumen).
  const keys = ["StartDate", "EndDate", "Party1Name", "Party1Representative", "Party1RepTitle", "Party2Name", "Party2Address", "ContractNo", "Deadline", "Note"];
  let scenarios = 0;
  for (let n = 0; n < 4000; n++) {
    const mk = () => {
      const c = baseContract();
      c.variables = { ContractNo: "A-1", Deadline: "2026-12-01", Note: "x" } as any;
      return c;
    };
    const want: Record<string, string> = {
      StartDate: pick(["Selasa, 10 November 2026", "Rabu, 11 November 2026"]), EndDate: pick(["22 November 2028", "23 November 2028"]),
      Party1Name: pick(["PT A", "PT B"]), Party1Representative: pick(["Budi", "Siti"]), Party1RepTitle: pick(["Direktur", "Komisaris"]),
      Party2Name: pick(["Xaxa", "Yyy"]), Party2Address: pick(["Jl. A", "[Alamat belum diisi]", "Jl. B"]),
      ContractNo: pick(["A-1", "A-2"]), Deadline: pick(["1 Desember 2026", "2 Desember 2026"]), Note: pick(["x", "y"]),
    };
    const inDoc: Record<string, string> = {};
    for (const k of keys) if (rnd() < 0.85) inDoc[k] = pick([want[k], "Selasa, 10 November 2026", "5 Januari 2027", "Jl. Lain", "", "[Alamat belum diisi]", "Teks bebas", "22 November 2028"]);
    const snap0: Record<string, string> = {};
    for (const k of keys) if (rnd() < 0.7) snap0[k] = pick([want[k], inDoc[k] ?? "", "Teks bebas", "Selasa, 10 November 2026"]);
    const a = mk(), b = mk(); const sa = { ...snap0 }, sb = { ...snap0 };
    const old = legacyReconcile(a, inDoc, want, sa);
    const neu = reconcileCore({ contract: b, inDoc, want, snap: sb, company });
    assert.deepEqual(neu, old, `skenario ${n}: push/pulled berbeda`);
    assert.deepEqual(sb, sa, `skenario ${n}: potret berbeda`);
    assert.deepEqual(b, a, `skenario ${n}: kontrak berbeda`);
    scenarios++;
  }
  ok(`reconcileCore == implementasi lama pada ${scenarios} skenario acak (kunci lama)`);
}

// ---------------------------------------------------------------------------
// 5. Pemetaan variabel Word untuk field baru
// ---------------------------------------------------------------------------
{
  const c = baseContract();
  const w = partyWordValues(c, company);
  assert.equal(w.DocumentTitle, "CEK HAK AKSES ANTAR ENTITAS");
  assert.equal(w.Party1Position, "Direktur", "kosong = jabatan perwakilan perusahaan (sama dgn hint modal)");
  assert.equal(w.Party1IdLabel, "No. Identitas");
  assert.ok(isPlaceholderValue(w.Party2Position) && isPlaceholderValue(w.Party1IdNumber) && isPlaceholderValue(w.Party2IdNumber));
  c.party2Position = "Manajer"; c.party2IdLabel = "NIK"; c.party2IdNumber = "3171"; c.party1Position = "Presiden Direktur";
  const w2 = partyWordValues(c, company);
  assert.equal(w2.Party2Position, "Manajer"); assert.equal(w2.Party2IdLabel, "NIK"); assert.equal(w2.Party2IdNumber, "3171");
  assert.equal(w2.Party1Position, "Presiden Direktur");
  for (const f of PARTY_FIELDS) assert.ok(f.word && PARTY_WORD_KEYS.includes(f.word));
  assert.equal(isPlaceholderValue("[DRAFT] Kontrak"), false, "judul berkurung siku bukan placeholder");
  assert.equal(isPlaceholderValue("[Alamat belum diisi]"), true);
  ok("pemetaan variabel Word field baru + penanda 'belum diisi'");
}

// ---------------------------------------------------------------------------
// 6. Dokumen .docx sungguhan: placeholder (var:*) diperbarui, termasuk varian @en
// ---------------------------------------------------------------------------
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const sdtXml = (tag: string, text: string) =>
  `<w:sdt><w:sdtPr><w:alias w:val="${tag}"/><w:tag w:val="var:${tag}"/><w:id w:val="${Math.floor(Math.random() * 1e8)}"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:sdtContent></w:sdt>`;
async function makeDocx(vars: Record<string, string>, extraEn: Record<string, string> = {}): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  const paras = Object.entries(vars).map(([k, v]) => `<w:p><w:r><w:t xml:space="preserve">${k}: </w:t></w:r>${sdtXml(k, v)}</w:p>`).join("");
  const en = Object.entries(extraEn).map(([k, v]) => `<w:p><w:r><w:t xml:space="preserve">EN ${k}: </w:t></w:r>${sdtXml(k + "@en", v)}</w:p>`).join("");
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${paras}${en}<w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: "nodebuffer" });
}
const rawXml = async (b: Buffer) => (await JSZip.loadAsync(b)).file("word/document.xml")!.async("string");

(async () => {
  {
    // Skenario penuh: dokumen berisi placeholder untuk SEMUA field data pihak.
    const c = baseContract();
    const settings = { companyName: company.name, companyAddress: company.address, companyRepresentativeTitle: company.repTitle };
    // desiredWordValues (server) = kunci lama + partyWordValues; di sini disusun sama utk uji.
    const want = (x: Contract): Record<string, string> => ({
      StartDate: new Date(x.startDate).toLocaleDateString("id-ID", { weekday: "long", day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }),
      EndDate: formatShortDateId(x.endDate),
      Party1Name: settings.companyName, Party1Address: x.party1Address || settings.companyAddress,
      Party2Name: x.party2Name, Party2Address: x.party2Address || "[Alamat belum diisi]",
      ...partyWordValues(x, company),
    });
    const initial = want(c);
    let docx = await makeDocx(initial, { StartDate: valueToEnglish(initial.StartDate), Party2Name: initial.Party2Name });
    let snap: Record<string, string> = {};
    // sinkron pertama: hanya mencatat potret
    let r = reconcileCore({ contract: cloneC(c), inDoc: await listVariables(docx), want: want(c), snap, company });
    assert.deepEqual([r.pushed, r.pulled], [[], []]);

    // --- A. Panel CLM menyimpan (panel -> DB): ubah beberapa field sekaligus ---
    const patch = normalizePartyPatch(c, {
      title: "KONTRAK BARU", party2Name: "PT Maju Jaya", party2Position: "Direktur Utama", party2IdLabel: "NPWP", party2IdNumber: "01.234.567.8-901.000",
      party1Position: "Presiden Direktur", party1IdLabel: "NPWP", party1IdNumber: "09.876", party1Address: "Jl. Sudirman 1", startDate: "2026-12-01",
      endDate: "2029-12-31", party2Address: "Jl. Mawar 5", docType: "MoU", party2Type: "Karyawan",
    }, { partyTypes: ["Vendor", "Karyawan"] });
    if (patch.ok === false) throw new Error("patch ditolak: " + patch.error);
    applyPartyPatch(c, patch);
    assert.equal(c.contractNumber, "174/PKS-KUY-MEDEX/GKM/XI/2026", "Nomor Dokumen tidak berubah");
    // plugin menerapkan push ke dokumen (diwakili setVariables milik server untuk variabel yang sama)
    const pushKeys = patch.changed.map((k) => PARTY_FIELDS.find((f) => f.key === k)?.word).filter(Boolean) as string[];
    const w1 = want(c);
    const toPush: Record<string, string> = {}; for (const k of pushKeys) toPush[k] = w1[k];
    docx = (await setVariables(docx, toPush)).docx;
    const after = await listVariables(docx);
    for (const k of pushKeys) assert.equal(after[k], w1[k], `variabel ${k} di dokumen harus mengikuti data baru`);
    assert.equal(after.Party2Position, "Direktur Utama");
    assert.equal(after.Party2IdNumber, "01.234.567.8-901.000");
    assert.equal(after.DocumentTitle, "KONTRAK BARU");
    // format run (tebal) dipertahankan
    assert.ok((await rawXml(docx)).includes("<w:b/>"), "format huruf tetap");
    ok("panel -> DB -> dokumen: placeholder Word untuk semua field yang diubah ikut berubah");

    // tick sinkron berikutnya tidak 'mengembalikan' data lama ke sistem (potret baru dicatat)
    const cc = cloneC(c);
    r = reconcileCore({ contract: cc, inDoc: await listVariables(docx), want: want(c), snap, company, live: { pullable: () => true } });
    assert.deepEqual(r.pulled, [], "tidak boleh menarik balik apa pun");
    assert.deepEqual(r.pushed, []);
    assert.equal(snap.Party2Name, "PT Maju Jaya");
    assert.deepEqual(partyFormOf(cc), partyFormOf(c), "data sistem tidak berubah");
    ok("sinkron setelah simpan panel: tidak ada pembalikan data");

    // --- B. @en: varian Inggris mengikuti (tanggal diterjemahkan) ---
    const en = await refreshEnglishVars(docx, want(c));
    assert.ok(en.changed >= 1);
    const enAfter = (await rawXml(en.docx));
    assert.ok(enAfter.includes(valueToEnglish(w1.StartDate)), "tanggal paragraf Inggris diterjemahkan");
    assert.ok(enAfter.includes("PT Maju Jaya"));
    ok("varian @en (paragraf Inggris) ikut diperbarui dgn tanggal diterjemahkan");

    // --- C. Dokumen -> DB: pengguna mengetik langsung di placeholder ---
    const typed = await setVariables(docx, { Party2Position: "Komisaris", Party2IdNumber: "99.999", DocumentTitle: "JUDUL DARI WORD", Party1Address: "Jl. Baru 9" });
    const dv = await listVariables(typed.docx);
    // belum stabil -> belum ditarik
    const c2 = cloneC(c);
    r = reconcileCore({ contract: c2, inDoc: dv, want: want(c), snap, company, live: { pullable: () => false } });
    assert.deepEqual(r.pulled, []);
    assert.equal(c2.party2Position, "Direktur Utama");
    // stabil -> ditarik ke data sistem
    r = reconcileCore({ contract: c2, inDoc: dv, want: want(c), snap, company, live: { pullable: () => true } });
    assert.deepEqual(r.pulled.sort(), ["DocumentTitle", "Party1Address", "Party2IdNumber", "Party2Position"]);
    assert.equal(c2.party2Position, "Komisaris"); assert.equal(c2.party2IdNumber, "99.999");
    assert.equal(c2.title, "JUDUL DARI WORD"); assert.equal(c2.party1Address, "Jl. Baru 9");
    const inv = englishInvalidationBetween(c, c2);
    assert.ok(inv.title && inv.preamble && !inv.docType);
    ok("dokumen -> DB: isian yang diketik langsung di Word ditarik ke data sistem (hanya setelah stabil)");

    // --- D. Aplikasi utama -> dokumen: data sistem berubah, dokumen mengikuti ---
    const c3 = cloneC(c2); c3.party2Address = "Alamat dari modal aplikasi utama"; c3.endDate = "2031-01-01";
    r = reconcileCore({ contract: cloneC(c3), inDoc: dv, want: want(c3), snap, company, live: { pullable: () => true } });
    assert.deepEqual(r.pushed.sort(), ["EndDate", "Party2Address"]);
    assert.equal(r.push.Party2Address, "Alamat dari modal aplikasi utama");
    // mode langsung: potret BELUM dicatat sampai dokumen terbukti memuat nilainya
    assert.notEqual(snap.Party2Address, "Alamat dari modal aplikasi utama");
    const applied = await setVariables(typed.docx, r.push);
    const r2 = reconcileCore({ contract: cloneC(c3), inDoc: await listVariables(applied.docx), want: want(c3), snap, company, live: { pullable: () => true } });
    assert.deepEqual([r2.pushed, r2.pulled], [[], []]);
    assert.equal(snap.Party2Address, "Alamat dari modal aplikasi utama");
    ok("aplikasi utama -> dokumen: data sistem baru didorong; potret dicatat setelah dokumen memuatnya");

    // --- E. Kegagalan menerapkan push TIDAK membuat data sistem 'dikembalikan' oleh dokumen lama ---
    const c4 = cloneC(c3); c4.party2Address = "Alamat E";
    const snapE = { ...snap };
    const eDoc = await listVariables(applied.docx); // dokumen lama (push gagal diterapkan)
    for (let i = 0; i < 3; i++) {
      const cx = cloneC(c4);
      const rr = reconcileCore({ contract: cx, inDoc: eDoc, want: want(c4), snap: snapE, company, live: { pullable: () => true } });
      assert.deepEqual(rr.pulled, [], `putaran ${i}: tidak boleh menarik nilai lama dari dokumen`);
      assert.equal(cx.party2Address, "Alamat E");
      assert.ok(rr.pushed.includes("Party2Address"), "terus mencoba mendorong sampai berhasil");
    }
    ok("push gagal diterapkan: sistem tetap menang & terus mencoba (tidak ada pembalikan data)");
  }

  {
    // Dokumen tanpa placeholder untuk field tertentu: tidak error, tidak ada perubahan
    const docx = await makeDocx({ Party2Name: "Xaxa" });
    const res = await setVariables(docx, { Party2Position: "Manajer", Party2Name: "Baru" });
    assert.equal(res.changed, 1);
    assert.deepEqual(await listVariables(res.docx), { Party2Name: "Baru" });
    ok("variabel yang tidak ada di dokumen dilewati tanpa merusak dokumen");
  }

  {
    // partyRevOf: berubah bila nilai berubah, stabil bila tidak
    const a = baseContract(), b = baseContract();
    assert.equal(partyRevOf(a), partyRevOf(b));
    b.party2IdNumber = "1";
    assert.notEqual(partyRevOf(a), partyRevOf(b));
    const c = baseContract(); (c as any).updatedAt = "x"; c.wordSync = { a: "1" };
    assert.equal(partyRevOf(a), partyRevOf(c), "updatedAt/wordSync tidak memengaruhi penanda (autosave editor mengubah updatedAt terus)");
    ok("penanda perubahan data pihak stabil & tidak terpengaruh autosave editor");
  }

  console.log(`ALL PARTY-FIELDS CHECKS PASSED (${passed})`);
})().catch((e) => { console.error(e); process.exit(1); });
