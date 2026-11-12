// Deteksi AKTIVITAS REVIEWER pada SESI REVIEW yang sedang berjalan.
//
// Aturan bisnis: reviewer (internal maupun eksternal) hanya boleh "Setujui" bila pada sesi review ini ia
// TIDAK menambah komentar, mengedit, menghapus/mencoret, atau membuat markup/Track Changes apa pun. Bila ada,
// ia wajib memakai "Ajukan Revisi".
//
//  - SESI = satu putaran review sejak pembuat dokumen mengajukan (internal) / membuka link (eksternal)
//    sampai ada keputusan. Putaran baru dimulai saat dokumen diajukan ulang → `sessionStart` baru.
//  - Hanya aktivitas yang tercatat SETELAH `sessionStart` yang dihitung. Markup lama (sesi/ronde sebelumnya,
//    reviewer lain) tidak ikut — dan TIDAK PERNAH dihapus: modul ini hanya MEMBACA.
//  - Internal: aktivitas milik reviewer itu sendiri (user id / nama pengubah). Eksternal: aktivitas tamu.
//
// PENTING — JAM: semua bukti di sini memakai JAM SERVER (waktu komentar/audit dicatat server). Tanggal DI DALAM
// berkas Word (w:date) sengaja TIDAK dipakai: ONLYOFFICE menulisnya dalam jam lokal pengguna yang diberi label
// UTC ("Z"), jadi membandingkannya dengan jam server meleset sebesar selisih zona waktu (mis. +7 jam di WIB →
// perubahan ronde lama tampak "baru"; zona barat UTC → perubahan baru tampak "lama").
//
// Modul ini murni (tanpa akses DB/berkas) supaya bisa diuji: lihat scripts/review-session-test.ts.
import type { ClauseComment, AuditTrail } from "./src/types";
import type { ChangeHistory } from "./review-changes.js";

/** Akhiran nama pengubah Track Changes untuk tamu eksternal (lihat server.ts: review-doc/config). */
export const EXTERNAL_AUTHOR_SUFFIX = "(Pihak Eksternal)";

/** Aksi audit yang menandai komentar diubah/dihapus (komentar yang dihapus hilang dari daftar, jejaknya di audit). */
export const COMMENT_EDIT_AUDIT_ACTION = "Edit Comment";
export const COMMENT_DELETE_AUDIT_ACTION = "Delete Comment";
/**
 * Editor Word (Track Changes) melaporkan reviewer baru mengubah dokumen. Server baru menerima berkasnya setelah
 * editor ditutup, jadi penanda ini dicatat SEGERA saat perubahan pertama (satu per sesi per reviewer).
 */
export const REVIEW_EDITOR_AUDIT_ACTION = "Review Editor Activity";
/**
 * Berkas tinjauan tersimpan dan MENGANDUNG perubahan/komentar baru dari seorang penulis dibanding berkas
 * sebelumnya (dicatat callback ONLYOFFICE di server.ts, hanya bila editor dibuka pada sesi review ini).
 * Penulis dibawa di `details` ("Penulis: <nama> | …"), bukan di userId, karena callback bisa datang dari sesi
 * editor bersama.
 */
export const REVIEW_SAVED_AUDIT_ACTION = "Review Changes Saved";
// CATATAN: audit "Review Track Changes" (berkas tersimpan, apa pun isinya) sengaja TIDAK dipakai: Document
// Server baru menyimpan setelah editor ditutup (bisa tertunda), jadi perubahan ronde lama bisa tercatat tersimpan
// di ronde baru.

export const savedChangesDetails = (author: string, added: number) => `Penulis: ${author} | perubahan/komentar baru di berkas tinjauan: ${added}`;
const savedAuthorOf = (details: string): string | null => /^Penulis: (.+?) \| /.exec(details || "")?.[1] ?? null;

export type SessionReviewer =
  | { kind: "internal"; userId: string; name: string }
  | { kind: "external" };

export interface ReviewSessionActivity {
  hasActivity: boolean;
  /** Awal sesi (ISO). null = tidak diketahui (data lama) → seluruh riwayat dianggap bagian sesi. */
  sessionStart: string | null;
  counts: {
    comments: number;        // komentar / balasan / usulan coret / usulan ganti baru
    commentEdits: number;    // komentar diubah
    commentDeletions: number; // komentar dihapus
    editorChanges: number;   // editor Word melaporkan reviewer mengubah dokumen (belum tentu sudah tersimpan)
    savedChanges: number;    // berkas tinjauan tersimpan berisi Track Changes/komentar baru dari reviewer
  };
}

export const EMPTY_ACTIVITY = (sessionStart: string | null): ReviewSessionActivity => ({
  hasActivity: false, sessionStart,
  counts: { comments: 0, commentEdits: 0, commentDeletions: 0, editorChanges: 0, savedChanges: 0 },
});

const norm = (s: string) => s.replace(/\s+/g, " ").trim().toLowerCase();

/** Apakah nama pengubah di berkas Word adalah reviewer ini. */
export function authorIsReviewer(author: string, reviewer: SessionReviewer): boolean {
  return reviewer.kind === "external"
    ? norm(author).endsWith(norm(EXTERNAL_AUTHOR_SUFFIX))
    : norm(author) === norm(reviewer.name);
}

/**
 * Perubahan/komentar yang BARU di `next` dibanding `prev`, per penulis. Murni selisih ISI (tanpa tanggal dan tanpa
 * id yang dinomori ulang tiap simpan): kunci = jenis + penulis + teks + konteks, dibandingkan sebagai multiset.
 */
export function addedByAuthor(prev: ChangeHistory, next: ChangeHistory): Map<string, number> {
  const keys = (h: ChangeHistory) => {
    const m = new Map<string, { author: string; n: number }>();
    const add = (author: string, key: string) => { const e = m.get(key); if (e) e.n++; else m.set(key, { author, n: 1 }); };
    for (const c of h.changes) add(c.author, `c|${c.kind}|${c.author}|${c.text}|${c.context}`);
    for (const c of h.comments) add(c.author, `m|${c.author}|${c.text}|${c.where}`);
    return m;
  };
  const before = keys(prev);
  const out = new Map<string, number>();
  for (const [k, e] of keys(next)) {
    const grew = e.n - (before.get(k)?.n ?? 0);
    if (grew > 0) out.set(e.author, (out.get(e.author) ?? 0) + grew);
  }
  return out;
}

export function summarizeReviewActivity(input: {
  reviewer: SessionReviewer;
  sessionStart: string | null;
  /** Komentar kontrak ini (semua; disaring di sini). */
  comments: ClauseComment[];
  /** Jejak audit kontrak ini. */
  audits: AuditTrail[];
}): ReviewSessionActivity {
  const { reviewer, sessionStart } = input;
  const startMs = sessionStart ? Date.parse(sessionStart) : Number.NEGATIVE_INFINITY;
  const out = EMPTY_ACTIVITY(sessionStart);

  const isMine = (c: ClauseComment) => reviewer.kind === "external" ? c.external === true : !c.external && c.userId === reviewer.userId;
  out.counts.comments = input.comments.filter((c) => isMine(c) && Date.parse(c.createdAt) >= startMs).length;

  const inSession = input.audits.filter((a) => Date.parse(a.timestamp) >= startMs);
  // Jejak audit milik reviewer ini (tamu eksternal dicatat dengan userId "external").
  const mine = inSession.filter((a) => a.userId === (reviewer.kind === "external" ? "external" : reviewer.userId));
  if (reviewer.kind === "internal") {
    out.counts.commentEdits = mine.filter((a) => a.action === COMMENT_EDIT_AUDIT_ACTION).length;
    out.counts.commentDeletions = mine.filter((a) => a.action === COMMENT_DELETE_AUDIT_ACTION).length;
  }
  out.counts.editorChanges = mine.filter((a) => a.action === REVIEW_EDITOR_AUDIT_ACTION).length;
  out.counts.savedChanges = inSession.filter((a) => {
    if (a.action !== REVIEW_SAVED_AUDIT_ACTION) return false;
    const author = savedAuthorOf(a.details);
    return author !== null && authorIsReviewer(author, reviewer);
  }).length;

  const k = out.counts;
  out.hasActivity = k.comments + k.commentEdits + k.commentDeletions + k.editorChanges + k.savedChanges > 0;
  return out;
}
