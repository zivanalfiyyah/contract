// Uji aturan sesi review (Setujui vs Ajukan Revisi). Jalankan: npx tsx scripts/review-session-test.ts
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import {
  summarizeReviewActivity, addedByAuthor, savedChangesDetails, EXTERNAL_AUTHOR_SUFFIX,
  COMMENT_EDIT_AUDIT_ACTION, COMMENT_DELETE_AUDIT_ACTION, REVIEW_EDITOR_AUDIT_ACTION, REVIEW_SAVED_AUDIT_ACTION,
  type SessionReviewer,
} from "../review-session";

const __dirname = path.dirname(new URL(import.meta.url).pathname);
let n = 0;
const t = (name: string, fn: () => void) => { fn(); n++; console.log("ok -", name); };

const START = "2026-11-10T10:00:00.000Z";
const BEFORE = "2026-11-09T09:00:00.000Z";
const AFTER = "2026-11-10T11:30:00.000Z";

const cm = (o: any = {}) => ({ id: "c" + Math.random(), tenantId: "t", contractId: "k1", clauseId: "x", clauseTitle: "X", userId: "u-ani", userName: "Ani", userRole: "legal", text: "teks", resolved: false, mentions: [], createdAt: AFTER, ...o });
const au = (o: any = {}) => ({ id: "a" + Math.random(), contractId: "k1", userId: "u-ani", userName: "Ani", userRole: "legal", action: "Add Comment", details: "", timestamp: AFTER, ipAddress: "", ...o });
const saved = (author: string, o: any = {}) => au({ action: REVIEW_SAVED_AUDIT_ACTION, details: savedChangesDetails(author, 1), ...o });
const tc = (o: any = {}) => ({ id: "1", kind: "insert", author: "Ani", date: "x", text: "baru", where: "", context: "ctx", ...o });
const dc = (o: any = {}) => ({ id: "1", author: "Ani", date: "x", text: "komentar", where: "pasal 1", ...o });
const ani: SessionReviewer = { kind: "internal", userId: "u-ani", name: "Ani" };
const ext: SessionReviewer = { kind: "external" };
const run = (reviewer: SessionReviewer, o: any = {}) => summarizeReviewActivity({ reviewer, sessionStart: START, comments: [], audits: [], ...o });

t("tanpa aktivitas → Setujui & Revisi sama-sama boleh (hasActivity false)", () => {
  assert.equal(run(ani).hasActivity, false);
  assert.equal(run(ext).hasActivity, false);
});

t("komentar baru reviewer di sesi ini → Setujui dikunci", () => {
  assert.equal(run(ani, { comments: [cm()] }).hasActivity, true);
});
t("balasan komentar (parentId) juga dihitung", () => {
  assert.equal(run(ani, { comments: [cm({ parentId: "c0" })] }).counts.comments, 1);
});
t("usulan coret / ganti dihitung", () => {
  assert.equal(run(ani, { comments: [cm({ kind: "strike" }), cm({ kind: "replace", replacement: "z" })] }).counts.comments, 2);
});
t("komentar sesi/ronde SEBELUMNYA tidak dihitung", () => {
  assert.equal(run(ani, { comments: [cm({ createdAt: BEFORE })] }).hasActivity, false);
});
t("komentar reviewer lain tidak mengunci reviewer ini", () => {
  assert.equal(run(ani, { comments: [cm({ userId: "u-budi", userName: "Budi" })] }).hasActivity, false);
});
t("komentar tamu eksternal tidak mengunci reviewer internal (dan sebaliknya)", () => {
  assert.equal(run(ani, { comments: [cm({ external: true, userId: "external" })] }).hasActivity, false);
  assert.equal(run(ext, { comments: [cm()] }).hasActivity, false);
  assert.equal(run(ext, { comments: [cm({ external: true, userId: "external" })] }).hasActivity, true);
});
t("komentar internal (balasan staf) pada thread eksternal tidak dihitung sebagai aktivitas tamu", () => {
  assert.equal(run(ext, { comments: [cm({ visibility: "eksternal" })] }).hasActivity, false);
});

t("komentar diedit / dihapus pada sesi ini dihitung (hapus tidak membuka kunci lagi)", () => {
  assert.equal(run(ani, { audits: [au({ action: COMMENT_DELETE_AUDIT_ACTION })] }).counts.commentDeletions, 1);
  assert.equal(run(ani, { audits: [au({ action: COMMENT_EDIT_AUDIT_ACTION })] }).counts.commentEdits, 1);
  assert.equal(run(ani, { comments: [], audits: [au({ action: COMMENT_DELETE_AUDIT_ACTION })] }).hasActivity, true);
});
t("audit lama / milik orang lain / aksi lain tidak dihitung", () => {
  assert.equal(run(ani, { audits: [au({ action: COMMENT_DELETE_AUDIT_ACTION, timestamp: BEFORE })] }).hasActivity, false);
  assert.equal(run(ani, { audits: [au({ action: COMMENT_DELETE_AUDIT_ACTION, userId: "u-budi" })] }).hasActivity, false);
  assert.equal(run(ani, { audits: [au({ action: "Setujui Kontrak (Approval Matrix)" })] }).hasActivity, false);
});

t("editor Word melaporkan perubahan (berkas belum tersimpan) → langsung dikunci di server", () => {
  const r = run(ani, { audits: [au({ action: REVIEW_EDITOR_AUDIT_ACTION })] });
  assert.equal(r.hasActivity, true);
  assert.equal(r.counts.editorChanges, 1);
});
t("penanda editor tamu eksternal (userId 'external') hanya mengunci sisi eksternal", () => {
  const a = au({ action: REVIEW_EDITOR_AUDIT_ACTION, userId: "external" });
  assert.equal(run(ext, { audits: [a] }).hasActivity, true);
  assert.equal(run(ani, { audits: [a] }).hasActivity, false);
  assert.equal(run(ext, { audits: [au({ action: REVIEW_EDITOR_AUDIT_ACTION })] }).hasActivity, false);
  assert.equal(run(ext, { audits: [au({ action: REVIEW_EDITOR_AUDIT_ACTION, userId: "external", timestamp: BEFORE })] }).hasActivity, false);
});

t("berkas tinjauan tersimpan berisi perubahan baru penulis ini (jam server) → dikunci", () => {
  assert.equal(run(ani, { audits: [saved("Ani")] }).counts.savedChanges, 1);
  assert.equal(run(ext, { audits: [saved(`Pak Joni ${EXTERNAL_AUTHOR_SUFFIX}`, { userId: "external" })] }).hasActivity, true);
});
t("penulis dicocokkan lewat nama (tanpa peduli huruf besar/spasi); userId pencatat tidak menentukan", () => {
  assert.equal(run(ani, { audits: [saved("  ANI ", { userId: "u-lain" })] }).hasActivity, true);
});
t("perubahan penulis lain / sisi lain tidak mengunci", () => {
  assert.equal(run(ani, { audits: [saved("Budi")] }).hasActivity, false);
  assert.equal(run(ani, { audits: [saved(`Pak Joni ${EXTERNAL_AUTHOR_SUFFIX}`)] }).hasActivity, false);
  assert.equal(run(ext, { audits: [saved("Ani")] }).hasActivity, false);
});
t("simpanan di sesi lalu tidak dihitung", () => {
  assert.equal(run(ani, { audits: [saved("Ani", { timestamp: BEFORE })] }).hasActivity, false);
});
t("REGRESI: berkas tersimpan TANPA perubahan baru (audit 'Review Track Changes' biasa) bukan aktivitas", () => {
  // Document Server menyimpan setelah editor ditutup; perubahan ronde lama bisa tercatat tersimpan di ronde baru.
  assert.equal(run(ext, { audits: [au({ action: "Review Track Changes", userId: "external" })] }).hasActivity, false);
  assert.equal(run(ani, { audits: [au({ action: "Review Track Changes" })] }).hasActivity, false);
});

t("addedByAuthor: hanya yang BARU dibanding berkas sebelumnya, per penulis, tanpa melihat tanggal/id", () => {
  const prev = { changes: [tc({ author: `sasa ${EXTERNAL_AUTHOR_SUFFIX}`, text: "lama", date: "2026-11-12T04:50:00Z", id: "1" })], comments: [dc({ author: `sasa ${EXTERNAL_AUTHOR_SUFFIX}`, text: "k lama" })] };
  // simpan ulang: id dinomori ulang & tanggal bisa berbeda format → bukan perubahan baru
  const sameAgain = { changes: [tc({ author: `sasa ${EXTERNAL_AUTHOR_SUFFIX}`, text: "lama", date: "2026-11-11T21:50:00Z", id: "7" })], comments: [dc({ author: `sasa ${EXTERNAL_AUTHOR_SUFFIX}`, text: "k lama", id: "9" })] };
  assert.equal(addedByAuthor(prev as any, sameAgain as any).size, 0);
  const next = { changes: [...sameAgain.changes, tc({ author: "Ani", text: "tambah" })], comments: [...sameAgain.comments, dc({ author: "Ani", text: "komentar baru" }), dc({ author: "Ani", text: "komentar lain" })] };
  const added = addedByAuthor(prev as any, next as any);
  assert.equal(added.get("Ani"), 3);
  assert.equal(added.has(`sasa ${EXTERNAL_AUTHOR_SUFFIX}`), false);
});
t("addedByAuthor: perubahan identik yang ditambah lagi (multiset) tetap terhitung", () => {
  const one = { changes: [tc()], comments: [] };
  const two = { changes: [tc(), tc({ id: "2" })], comments: [] };
  assert.equal(addedByAuthor(one as any, two as any).get("Ani"), 1);
});

t("pengajuan ulang = sesi baru: aktivitas lama berhenti mengunci", () => {
  const comments = [cm({ createdAt: AFTER })];
  assert.equal(summarizeReviewActivity({ reviewer: ani, sessionStart: START, comments, audits: [] }).hasActivity, true);
  assert.equal(summarizeReviewActivity({ reviewer: ani, sessionStart: "2026-11-12T08:00:00.000Z", comments, audits: [saved("Ani")] }).hasActivity, false);
});
t("sessionStart tidak diketahui (data lama) → seluruh riwayat dihitung (aman)", () => {
  assert.equal(summarizeReviewActivity({ reviewer: ani, sessionStart: null, comments: [cm({ createdAt: BEFORE })], audits: [] }).hasActivity, true);
});

t("modul murni: hanya membaca, tidak mengubah masukan (riwayat tidak dihapus)", () => {
  const comments = [cm(), cm({ createdAt: BEFORE })];
  const audits = [saved("Ani"), au({ action: COMMENT_DELETE_AUDIT_ACTION })];
  const snap = JSON.stringify({ comments, audits });
  run(ani, { comments, audits });
  assert.equal(JSON.stringify({ comments, audits }), snap);
});

t("server: Setujui divalidasi di backend (internal & eksternal); jam berkas tidak dipakai; callback mencatat isi baru", () => {
  const src = fs.readFileSync(path.resolve(__dirname, "../server.ts"), "utf8");
  const internal = src.slice(src.indexOf('"/api/contracts/:id/approval-decision"'), src.indexOf('app.delete("/api/contracts/:id", requireAuth'));
  assert.ok(internal.includes('decision === "approved"') && internal.includes("REVIEW_ACTIVITY_DETECTED") && internal.includes("status(409)"));
  assert.ok(internal.includes("revisionActivity?.hasActivity"));
  const ext = src.slice(src.indexOf('"/api/external-review/:token/approve"'), src.indexOf('"/api/external-review/:token/request-revision"'));
  assert.ok(ext.includes("loadReviewActivity") && ext.includes("REVIEW_ACTIVITY_DETECTED") && ext.includes("status(409)"));
  assert.ok(src.includes('"/api/contracts/:id/review-session/activity"') && src.includes('"/api/external-review/:token/review-session/activity"'));
  // callback: perubahan baru dicatat (REVIEW_SAVED_AUDIT_ACTION), dengan penjaga editor lama (iat)
  const cb = src.slice(src.indexOf('"/api/onlyoffice/review-callback/:id"'), src.indexOf("async function changeHistoryOf"));
  assert.ok(cb.includes("REVIEW_SAVED_AUDIT_ACTION") && cb.includes("iat"));
  // riwayat tidak boleh dihapus oleh logika sesi, dan tanggal berkas (w:date) tidak dipakai di modul
  const mod = fs.readFileSync(path.resolve(__dirname, "../review-session.ts"), "utf8");
  const code = mod.replace(/\/\/.*$/gm, "");
  assert.ok(!/splice|delete |\.pop\(|writeFile/.test(code));
  assert.ok(!/parseDocxDate|\.date\b/.test(code));
});

console.log(`\n${n} kelompok uji lolos`);
