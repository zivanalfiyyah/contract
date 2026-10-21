// Alur status pekerjaan legal & kontrak — SATU sumber kebenaran, dipakai
// server (server.ts, reminders.ts) DAN frontend (App.tsx) supaya label,
// urutan, dan hitungan status selalu identik di Dashboard Legal, Pekerjaan
// Legal, Kontrak Eksternal, Monitoring, filter, badge, dan pencarian.
//
// Alur (12 tahap):
//   Pekerjaan Masuk → Belum Diproses → Draft → Review Internal →
//   Review Pihak Kedua → Revisi/Negosiasi → Finalisasi → Proses TTD →
//   Dokumen Kembali → Arsip → Distribusi/Implementasi → Aktif
//
//  - Pekerjaan Masuk  : permintaan baru (formulir/divisi), belum disetujui Legal.
//  - Belum Diproses   : sudah disetujui Legal, kontrak belum dibuat.
//  - Draft … Dokumen Kembali : mengikuti status kontrak yang tertaut.
//  - Arsip            : scan TTD basah sudah diunggah (kontrak sudah mengikat;
//                       masa berlaku & pengingat sudah berjalan sejak tahap ini).
//  - Distribusi/Implementasi : Legal sudah mengumumkan rilis ke orang internal.
//  - Aktif            : SEMUA penerima pengumuman sudah membuka notifikasinya.
//
// Status akhir di luar alur: Ditolak (pekerjaan), Tidak Aktif (lewat masa
// berlaku), Tidak Berlaku (ditutup manual — dulu "Diarsipkan"), Diakhiri.

export const LEGAL_STAGES = [
  "PekerjaanMasuk",
  "BelumDiproses",
  "Draft",
  "ReviewInternal",
  "ReviewPihakKedua",
  "RevisiNegosiasi",
  "Finalisasi",
  "ProsesTTD",
  "DokumenKembali",
  "Arsip",
  "Distribusi",
  "Aktif",
] as const;
export type LegalStage = (typeof LEGAL_STAGES)[number];
export type LegalEndStatus = "Ditolak" | "TidakAktif" | "Archived" | "Terminated";
export type UnifiedStatus = LegalStage | LegalEndStatus;

export const UNIFIED_STATUS_LABEL: Record<string, string> = {
  PekerjaanMasuk: "Pekerjaan Masuk",
  BelumDiproses: "Belum Diproses",
  Draft: "Draft",
  ReviewInternal: "Review Internal",
  ReviewPihakKedua: "Review Pihak Kedua",
  RevisiNegosiasi: "Revisi/Negosiasi",
  Finalisasi: "Finalisasi",
  ProsesTTD: "Proses TTD",
  DokumenKembali: "Dokumen Kembali",
  Arsip: "Arsip",
  Distribusi: "Distribusi/Implementasi",
  Aktif: "Aktif",
  Ditolak: "Ditolak",
  TidakAktif: "Tidak Aktif",
  Archived: "Tidak Berlaku",
  Terminated: "Diakhiri",
};

export const unifiedStatusLabel = (s?: string) => (s && UNIFIED_STATUS_LABEL[s]) || s || "-";

/**
 * Kontrak yang SUDAH MENGIKAT (TTD basah sudah diunggah): Arsip, Distribusi,
 * Aktif. Dipakai untuk monitoring & pengingat masa berlaku, perpanjangan,
 * addendum, pengumuman rilis, watermark CONTROLLED, dsb. — semua yang dulu
 * memeriksa `status === "Aktif"` dengan arti "kontrak sedang berlaku".
 */
export const IN_FORCE_STATUSES = ["Arsip", "Distribusi", "Aktif"] as const;
export const isInForce = (status?: string) => !!status && (IN_FORCE_STATUSES as readonly string[]).includes(status);

/** Status yang pernah berlaku (berlaku sekarang atau sudah berakhir). */
export const EVER_IN_FORCE_STATUSES = ["Arsip", "Distribusi", "Aktif", "TidakAktif", "Archived", "Terminated"] as const;
export const wasEverInForce = (status?: string) => !!status && (EVER_IN_FORCE_STATUSES as readonly string[]).includes(status);

interface ContractLike {
  status: string;
  revisionStage?: string;
  externalRoundCount?: number;
  externalApprovals?: unknown[];
  externalReviewToken?: string;
  signingReceipt?: unknown;
}

/** Tahap alur untuk sebuah kontrak (berdasarkan status mentah + penanda). */
export function contractStage(c: ContractLike): UnifiedStatus {
  switch (c.status) {
    case "Draft":
      if (c.revisionStage === "internal" || c.revisionStage === "eksternal") return "RevisiNegosiasi";
      // Kontrak lama: link review dibuka saat Draft (sebelum alur Okt 2026).
      if (!c.externalRoundCount) {
        if ((c.externalApprovals?.length || 0) > 0) return "RevisiNegosiasi";
        if (c.externalReviewToken) return "ReviewPihakKedua";
      }
      return "Draft";
    case "OnReview": return "ReviewInternal";
    case "ApprovedInternal": // disetujui internal, siap dikirim ke mitra
    case "ReviewEksternal": return "ReviewPihakKedua";
    case "FinalReview": return "Finalisasi";
    case "FullyApproved": return c.signingReceipt ? "DokumenKembali" : "ProsesTTD";
    case "Arsip": return "Arsip";
    case "Distribusi": return "Distribusi";
    case "Aktif": return "Aktif";
    case "TidakAktif": return "TidakAktif";
    case "Archived": return "Archived";
    case "Terminated": return "Terminated";
    default: return "Draft";
  }
}

/** Tahap alur untuk sebuah Pekerjaan Legal (mengikuti kontrak tertaut bila ada). */
export function jobStage(job: { status: string }, linked?: ContractLike | null): UnifiedStatus {
  if (job.status === "menunggu_persetujuan") return "PekerjaanMasuk";
  if (job.status === "ditolak") return "Ditolak";
  if (linked) return contractStage(linked);
  return "BelumDiproses";
}
