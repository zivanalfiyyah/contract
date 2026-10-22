// Model peran (RBAC) bersama — dipakai server (server.ts, auth.ts) DAN
// frontend (App.tsx) supaya aturan "siapa boleh apa" tidak pernah berbeda
// antara tombol yang tampil dan endpoint yang menerima.
//
// Ringkasan aturan (hasil diskusi):
//  - Super Admin   : semua fitur, lintas perusahaan.
//  - IT/System     : semua PENGATURAN (user, master data, sistem), tanpa hak
//                    bisnis (tidak buat/ajukan/approve kontrak) & tidak melihat
//                    kontrak karyawan.
//  - Admin         : admin per perusahaan — user & konfigurasi PT-nya + semua
//                    dokumen di PT-nya.
//  - Legal         : buat, ajukan & tindak lanjuti revisi, approve, kirim ke
//                    eksternal/TTD; lihat semua dokumen.
//  - Manajemen     : lihat semua (baca) + approve bila dipilih di matriks.
//  - Divisi bisnis : Marketing / Business Partnership / Account Executive /
//                    Operasional — lihat dokumen DIVISINYA saja, mengajukan
//                    Permintaan Kontrak ke Legal, boleh dipilih jadi reviewer.
//  - HR/GA         : buat & ajukan sendiri kontrak karyawan (+ DCS); tidak
//                    melihat kontrak eksternal.
//  - staff/viewer  : peran lama — hanya lihat sampai di-set ulang di Kelola
//                    Pengguna. "manager" lama dimigrasikan ke "manajemen".

export const DIVISION_ROLES = ["marketing", "business_partnership", "account_executive", "operasional"] as const;
export type DivisionRole = (typeof DIVISION_ROLES)[number];

export const ROLE_DEFS: { value: string; label: string; desc: string; assignable: boolean }[] = [
  { value: "super_admin", label: "Super Admin", desc: "Semua fitur, lintas perusahaan.", assignable: false },
  { value: "it_system", label: "IT / System", desc: "Kelola pengguna, master data & pengaturan sistem. Tidak membuat/menyetujui kontrak dan tidak melihat kontrak karyawan.", assignable: true },
  { value: "admin", label: "Admin Perusahaan", desc: "Kelola pengguna & konfigurasi perusahaannya, serta seluruh dokumen di perusahaan itu.", assignable: true },
  { value: "legal", label: "Legal", desc: "Susun kontrak, ajukan review, tindak lanjuti revisi, kirim ke pihak eksternal & TTD. Melihat semua dokumen.", assignable: true },
  { value: "manajemen", label: "Manajemen", desc: "Melihat semua dokumen dan menyetujui bila dipilih di matriks approval.", assignable: true },
  { value: "marketing", label: "Marketing", desc: "Melihat dokumen divisi Marketing & mengajukan Permintaan Kontrak ke Legal.", assignable: true },
  { value: "business_partnership", label: "Business Partnership", desc: "Melihat dokumen divisi Business Partnership & mengajukan Permintaan Kontrak ke Legal.", assignable: true },
  { value: "account_executive", label: "Account Executive", desc: "Melihat dokumen divisi Account Executive & mengajukan Permintaan Kontrak ke Legal.", assignable: true },
  { value: "operasional", label: "Operasional", desc: "Melihat dokumen divisi Operasional & mengajukan Permintaan Kontrak ke Legal.", assignable: true },
  { value: "hr_ga", label: "HR / GA", desc: "Membuat & mengajukan sendiri kontrak karyawan, serta dokumen DCS.", assignable: true },
  // Peran lama — tetap dikenali supaya data lama tidak rusak.
  { value: "staff", label: "Staff (lama — hanya lihat)", desc: "Peran lama. Hanya bisa melihat sampai perannya diatur ulang.", assignable: false },
  { value: "viewer", label: "Viewer (hanya lihat)", desc: "Hanya melihat dokumen — tanpa mengubah apa pun.", assignable: true },
  { value: "manager", label: "Manajemen", desc: "Peran lama (otomatis jadi Manajemen).", assignable: false },
];

export const roleLabelOf = (role?: string) => ROLE_DEFS.find((r) => r.value === role)?.label || role || "-";
export const isDivisionRole = (role?: string): role is DivisionRole => !!role && (DIVISION_ROLES as readonly string[]).includes(role);

// ----- Kelompok peran per kemampuan (dipakai requireRole di server) -----
// super_admin selalu lolos requireRole (lihat auth.ts), jadi tidak perlu ditulis.
export const R = {
  // Pengaturan: user, registrasi, settings, master data, logo.
  settings: ["admin", "it_system"],
  // Template & klausul.
  templates: ["admin", "legal"],
  // Buat/ubah/ajukan dokumen kontrak (HR/GA dibatasi ke kontrak karyawan
  // oleh canViewContract — kontrak lain "tidak ditemukan" baginya).
  contractEditors: ["admin", "legal", "hr_ga"],
  // Kirim ke pihak eksternal, legal hold, pemusnahan, pengumuman.
  legalOnly: ["admin", "legal"],
  // Anggaran.
  budgets: ["admin", "legal", "manajemen"],
  // Data karyawan (modul).
  employeeData: ["admin", "hr_ga"],
  // Data vendor (modul).
  vendorData: ["admin", "legal"],
  // Dokumen DCS (buat/ubah).
  dcsEditors: ["admin", "legal", "hr_ga"],
  // Permintaan Kontrak ke Legal (divisi bisnis).
  requesters: [...DIVISION_ROLES],
} as const;

// Peran yang boleh dipilih sebagai approver / di-mention untuk review.
export const canBeApprover = (role?: string) =>
  !!role && !["viewer", "staff", "it_system"].includes(role);

export const isAdminLike = (role?: string) => role === "admin" || role === "super_admin";

// Kontrak karyawan (internal) vs eksternal — sama dengan isEmployeeContract di App.
export const isEmployeeContractLike = (c: { party2Type?: string; employmentType?: string }) =>
  c.party2Type === "Karyawan" || c.party2Type === "Employee";

// Siapa boleh MELIHAT sebuah kontrak. Orang yang terlibat langsung (pembuat,
// pengaju, approver di ronde mana pun) selalu boleh melihat, apa pun perannya
// — kalau tidak, approver dari divisi lain tidak bisa membuka dokumen yang
// harus ia setujui.
export function canViewContract(
  user: { id: string; role: string; actingForIds?: string[] },
  c: {
    party2Type?: string; employmentType?: string; requestingDivision?: string;
    createdById?: string; approvalSubmittedById?: string;
    approvalSteps?: { approverId?: string; decidedByDelegateId?: string }[];
    approvalHistory?: { steps?: { approverId?: string }[] }[];
  },
): boolean {
  const role = user.role;
  if (role === "super_admin" || role === "admin" || role === "legal" || role === "manajemen" || role === "manager") return true;
  if (c.createdById === user.id || c.approvalSubmittedById === user.id) return true;
  const me = [user.id, ...(user.actingForIds || [])];
  if ((c.approvalSteps || []).some((s) => me.includes(s.approverId || "")) ) return true;
  if ((c.approvalSteps || []).some((s) => s.approverId === user.id || s.decidedByDelegateId === user.id)) return true;
  if ((c.approvalHistory || []).some((r) => (r.steps || []).some((s) => s.approverId === user.id))) return true;
  const employee = isEmployeeContractLike(c);
  if (role === "hr_ga") return employee;
  if (role === "it_system" || role === "staff" || role === "viewer") return !employee;
  if (isDivisionRole(role)) return !employee && c.requestingDivision === role;
  return false;
}

// ===================== MATRIKS HAK AKSES AKSI =====================
// SATU-SATUNYA tempat daftar izin aksi didefinisikan. Setiap izin punya
// default (= aturan sebelum matriks ada, jadi perilaku tidak berubah sampai
// admin mengubahnya) dan bisa ditimpa per role lewat Konfigurasi > Hak Akses
// (disimpan di pengaturan tenant: actionPermissions). Server (requirePerm)
// DAN frontend (tombol) membaca sumber yang sama lewat isPermAllowed().
// Super Admin SELALU lolos semua izin (tidak bisa mengunci diri sendiri).
// Aturan yang bukan per-role tetap di kode: Super-Admin-only (backup, dll)
// dan penugasan per dokumen (approver / pembuat dokumen di Mode Tinjau).

export type PermDef = { key: string; label: string; group: string; defaults: readonly string[]; hint?: string };

export const PERM_GROUPS = [
  "Dokumen (Eksternal & Karyawan)",
  "Review & Persetujuan",
  "TTD & Arsip",
  "Library & Template",
  "Dokumen Terkontrol (DCS)",
  "Pekerjaan Legal",
  "Data & Anggaran",
  "Pengaturan Sistem",
] as const;

export const PERMISSIONS: readonly PermDef[] = [
  // --- Dokumen ---
  { key: "contract.edit", group: "Dokumen (Eksternal & Karyawan)", label: "Buat, ubah & ajukan dokumen", defaults: ["admin", "legal", "hr_ga"], hint: "Buat dokumen, edit draft, ajukan review internal, perpanjang, addendum, versi berkas, unggah logo, kelola sub folder. HR/GA dibatasi ke dokumen karyawan oleh aturan lihat-dokumen." },
  { key: "contract.delete", group: "Dokumen (Eksternal & Karyawan)", label: "Hapus dokumen", defaults: ["legal"] },
  { key: "contract.move", group: "Dokumen (Eksternal & Karyawan)", label: "Pindahkan dokumen ke folder/kategori arsip", defaults: ["legal"] },
  { key: "contract.classify", group: "Dokumen (Eksternal & Karyawan)", label: "Ubah klasifikasi & nomor dokumen", defaults: ["admin", "legal", "hr_ga"] },
  { key: "contract.bulk", group: "Dokumen (Eksternal & Karyawan)", label: "Impor massal & kirim pengingat massal", defaults: ["admin", "legal", "hr_ga"] },
  // --- Review ---
  { key: "review.access", group: "Review & Persetujuan", label: "Mode Tinjau: komentar / coret / ganti", defaults: ["admin", "legal"], hint: "Approver yang ditunjuk Legal & pembuat dokumen tetap bisa meninjau dokumen mereka walau role-nya tidak dicentang." },
  { key: "review.markup_decide", group: "Review & Persetujuan", label: "Terima / tolak usulan coret & ganti", defaults: ["admin", "legal", "hr_ga"], hint: "Tetap hanya untuk pembuat/pengaju dokumen (atau admin)." },
  { key: "review.send_external", group: "Review & Persetujuan", label: "Kirim/buka akses review pihak kedua & lewati review", defaults: ["admin", "legal"] },
  // --- TTD & Arsip ---
  { key: "signing.manage", group: "TTD & Arsip", label: "Catat pengiriman/penerimaan TTD & unggah TTD basah", defaults: ["admin", "legal", "hr_ga"] },
  { key: "release.announce", group: "TTD & Arsip", label: "Pengumuman rilis (internal & eksternal)", defaults: ["admin", "legal"] },
  { key: "retention.manage", group: "TTD & Arsip", label: "Legal hold & pemusnahan dokumen", defaults: ["admin", "legal"] },
  // --- Library ---
  { key: "templates.manage", group: "Library & Template", label: "Kelola klausul, kategori klausul & template", defaults: ["admin", "legal"] },
  // --- DCS ---
  { key: "dcs.config", group: "Dokumen Terkontrol (DCS)", label: "Atur jenis dokumen, aturan nomor & klausul DCS", defaults: ["admin", "legal"] },
  { key: "dcs.edit", group: "Dokumen Terkontrol (DCS)", label: "Buat & ubah dokumen DCS (versi, berkas, terjemahan)", defaults: ["admin", "legal", "hr_ga"] },
  { key: "dcs.release", group: "Dokumen Terkontrol (DCS)", label: "Rilis, distribusi & tarik dokumen DCS", defaults: ["admin", "legal", "manajemen"] },
  // --- Pekerjaan Legal ---
  { key: "legaljob.request", group: "Pekerjaan Legal", label: "Ajukan Pekerjaan Legal", defaults: ["admin", "legal", "marketing", "business_partnership", "account_executive", "operasional"] },
  { key: "legaljob.process", group: "Pekerjaan Legal", label: "Setujui/tolak Pekerjaan Legal & kelola link formulir", defaults: ["admin", "legal"] },
  // --- Data ---
  { key: "budgets.manage", group: "Data & Anggaran", label: "Kelola anggaran", defaults: ["admin", "legal", "manajemen"] },
  { key: "employees.manage", group: "Data & Anggaran", label: "Kelola data karyawan", defaults: ["admin", "hr_ga"] },
  { key: "vendors.manage", group: "Data & Anggaran", label: "Kelola data vendor", defaults: ["admin", "legal"] },
  // --- Pengaturan ---
  { key: "settings.manage", group: "Pengaturan Sistem", label: "Kelola pengguna, pengaturan & master data", defaults: ["admin", "it_system"] },
  { key: "settings.employment_types", group: "Pengaturan Sistem", label: "Ubah Kategori Dokumen Karyawan saja", defaults: ["hr_ga"], hint: "Khusus bagian kategori karyawan; pengaturan lain tetap butuh izin di atas." },
];

export const PERM_KEYS = PERMISSIONS.map((p) => p.key);

/** Role yang bisa dicentang di matriks (Super Admin selalu penuh, tidak ditampilkan sebagai kolom bebas). */
export const MATRIX_ROLES = ROLE_DEFS.filter((r) => r.assignable).map((r) => r.value);

export type ActionPermissions = Record<string, string[]>;

/** Role legacy dinormalkan: "manager" = "manajemen". */
const normRole = (role?: string) => (role === "manager" ? "manajemen" : role || "");

/**
 * Apakah `role` boleh melakukan `key`? Urutan: super_admin → selalu; override
 * tenant (actionPermissions[key], bila array) → isinya; kalau tidak ada →
 * default izin tsb. Key tak dikenal → ditolak (aman).
 */
export function isPermAllowed(role: string | undefined, key: string, overrides?: ActionPermissions | null): boolean {
  if (role === "super_admin") return true;
  const def = PERMISSIONS.find((p) => p.key === key);
  if (!def) return false;
  const r = normRole(role);
  const ov = overrides?.[key];
  const list: readonly string[] = Array.isArray(ov) ? ov : def.defaults;
  return list.includes(r);
}

/** Bersihkan kiriman actionPermissions dari klien: hanya key & role yang dikenal. */
export function sanitizeActionPermissions(input: unknown): ActionPermissions {
  const out: ActionPermissions = {};
  if (!input || typeof input !== "object") return out;
  const roles = new Set<string>(ROLE_DEFS.map((r) => r.value));
  for (const [k, v] of Object.entries(input as Record<string, unknown>)) {
    if (!PERM_KEYS.includes(k) || !Array.isArray(v)) continue;
    out[k] = Array.from(new Set(v.map(String).filter((x) => roles.has(x) && x !== "super_admin")));
  }
  return out;
}

// Siapa boleh MENGAKSES Mode Tinjau & memberi komentar / coret / ganti pada
// sebuah dokumen. Role yang dicentang di matriks (review.access) selalu boleh.
// Role lain TIDAK punya akses bebas — hanya bila ada tindakan/suruhan dari
// Legal, yaitu: ditunjuk sebagai approver pada tahap review yang sedang
// berjalan (Review Internal / Review Final Internal, termasuk yang menerima
// delegasi), atau pembuat/pengaju dokumen itu sendiri. Dipakai server (403)
// DAN frontend (tombol disembunyikan) supaya keduanya selalu sepakat.
export function canReviewContract(
  user: { id: string; role: string; actingForIds?: string[] },
  c: {
    status?: string; createdById?: string; approvalSubmittedById?: string;
    approvalSteps?: { approverId?: string; decidedByDelegateId?: string }[];
  },
  overrides?: ActionPermissions | null,
): boolean {
  if (isPermAllowed(user.role, "review.access", overrides)) return true;
  if (c.createdById === user.id || c.approvalSubmittedById === user.id) return true;
  const inReviewStage = c.status === "OnReview" || c.status === "FinalReview";
  if (!inReviewStage) return false;
  const me = [user.id, ...(user.actingForIds || [])];
  return (c.approvalSteps || []).some((s) => me.includes(s.approverId || "") || s.decidedByDelegateId === user.id);
}

// Default menu sidebar per peran (dipakai bila admin belum mengatur matriks
// Hak Akses untuk peran tsb). id harus sama dengan MENU_ITEMS di App.tsx.
export const DEFAULT_MENUS_BY_ROLE: Record<string, string[]> = {
  it_system: ["dashboard", "monitoring", "internal-docs", "arsip", "all-docs", "konfigurasi", "audits", "audit-kontrak-dcs"],
  marketing: ["dashboard", "pekerjaan-legal", "monitoring", "arsip", "all-docs", "kalender-legal"],
  business_partnership: ["dashboard", "pekerjaan-legal", "monitoring", "arsip", "all-docs", "kalender-legal"],
  account_executive: ["dashboard", "pekerjaan-legal", "monitoring", "arsip", "all-docs", "kalender-legal"],
  operasional: ["dashboard", "pekerjaan-legal", "monitoring", "internal-docs", "arsip", "all-docs", "kalender-legal"],
  hr_ga: ["dashboard", "kontrak-karyawan", "clauses", "internal-docs", "arsip", "all-docs", "kalender-legal", "konfigurasi"],
  manajemen: ["dashboard", "dashboard-legal", "pekerjaan-legal", "monitoring", "kontrak-karyawan", "internal-docs", "arsip", "all-docs", "audits", "audit-kontrak-dcs", "kalender-legal", "retensi-arsip", "performa-vendor", "ai-tools"],
  staff: ["dashboard", "monitoring", "arsip", "all-docs"],
  viewer: ["dashboard", "monitoring", "arsip", "all-docs"],
};
