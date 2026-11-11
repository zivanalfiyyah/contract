// "Data Pihak & Kontrak" — satu sumber kebenaran untuk modal "Edit Data Pihak &
// Kontrak" (aplikasi utama) dan plugin ONLYOFFICE "Panel CLM".
//
// Modul ini MURNI (tanpa database / Express) supaya bisa diuji langsung
// (scripts/party-fields-test.ts) dan dipakai server di dua jalur:
//   1. Panel CLM menyimpan data pihak  -> normalizePartyPatch + applyPartyPatch
//   2. Sinkron dua arah dgn variabel Word (var:<Nama>) -> reconcileCore
//
// Aturan yang SENGAJA dipertahankan dari modal aplikasi utama
// (handleSaveEditPartyModal di src/App.tsx):
//  - Nomor Dokumen terkunci (jejak audit) — tidak pernah ikut diubah.
//  - Judul & Nama Pihak Kedua wajib diisi.
//  - Bidang opsional yang dikosongkan disimpan sebagai `undefined`.
//  - Mengubah judul / jenis dokumen / data yang tertanam di narasi membuang
//    padanan terjemahan Inggris yang sudah basi (titleEn / docTypeEn / preambleEn).
//  - Hanya dokumen Draft (dan pemilik dokumen saat revisi) yang boleh diubah.
import type { Contract } from "./src/types";

// ---------------------------------------------------------------------------
// Skema field
// ---------------------------------------------------------------------------
export type PartyGroupId = "kontrak" | "pihak1" | "pihak2";

export const PARTY_TEXT_KEYS = [
  "title", "docType", "startDate", "endDate",
  "party1Address", "party1Position", "party1IdLabel", "party1IdNumber",
  "party2Name", "party2Type", "party2Address", "party2Position", "party2IdLabel", "party2IdNumber",
] as const;
export type PartyTextKey = (typeof PARTY_TEXT_KEYS)[number];

export const PARTY_LOGO_KEYS = ["party2LogoUrl", "party2LogoKey", "party2LogoMimeType"] as const;
export type PartyLogoKey = (typeof PARTY_LOGO_KEYS)[number];

export type PartyKey = PartyTextKey | PartyLogoKey;
export type PartyForm = Record<PartyKey, string>;

/** Field yang kosongnya BOLEH berarti "hapus" (disimpan undefined). */
const OPTIONAL_KEYS: PartyKey[] = [
  "docType",
  "party1Address", "party1Position", "party1IdLabel", "party1IdNumber",
  "party2Address", "party2Position", "party2IdLabel", "party2IdNumber",
  "party2LogoUrl", "party2LogoKey", "party2LogoMimeType",
];

export interface PartyFieldDef {
  key: PartyTextKey;
  group: PartyGroupId;
  label: string;
  type: "text" | "date" | "select";
  required?: boolean;
  /** Placeholder statis (sama dgn modal). Yang dinamis dihitung di buildPartyView. */
  placeholder?: string;
  /** Nama variabel Word ("var:<word>") yang mewakili field ini, bila ada. */
  word?: string;
  /** Field berdampingan dalam satu baris (mis. Label ID + Nomor identitas). */
  row?: string;
  /** Lebar relatif dalam baris: "narrow" = sempit, selain itu mengisi sisa. */
  size?: "narrow";
}

// Urutan = urutan di modal aplikasi utama.
export const PARTY_FIELDS: PartyFieldDef[] = [
  { key: "title", group: "kontrak", label: "Judul Dokumen", type: "text", required: true, word: "DocumentTitle" },
  { key: "docType", group: "kontrak", label: "Jenis Dokumen", type: "text", placeholder: "mis. Surat Perjanjian Kerjasama", word: "DocumentType" },
  { key: "startDate", group: "kontrak", label: 'Tanggal ("pada hari ini, tanggal...")', type: "date", word: "StartDate" },
  { key: "endDate", group: "kontrak", label: "Tanggal Berakhir Kontrak", type: "date", word: "EndDate" },

  { key: "party1Address", group: "pihak1", label: "Alamat", type: "text", word: "Party1Address" },
  { key: "party1Position", group: "pihak1", label: "Jabatan Penandatanganan", type: "text", word: "Party1Position" },
  { key: "party1IdLabel", group: "pihak1", label: "Label ID", type: "text", placeholder: "Label ID (mis. NPWP)", word: "Party1IdLabel", row: "p1-id", size: "narrow" },
  { key: "party1IdNumber", group: "pihak1", label: "Nomor Identitas", type: "text", placeholder: "Nomor identitas", word: "Party1IdNumber", row: "p1-id" },

  { key: "party2Name", group: "pihak2", label: "Nama", type: "text", required: true, placeholder: "Nama", word: "Party2Name", row: "p2-nama" },
  { key: "party2Type", group: "pihak2", label: "Tipe", type: "select", word: "Party2Type", row: "p2-nama", size: "narrow" },
  { key: "party2Address", group: "pihak2", label: "Alamat", type: "text", placeholder: "Alamat", word: "Party2Address" },
  { key: "party2Position", group: "pihak2", label: "Jabatan", type: "text", placeholder: "Jabatan", word: "Party2Position" },
  { key: "party2IdLabel", group: "pihak2", label: "Label ID", type: "text", placeholder: "Label ID (mis. NIK)", word: "Party2IdLabel", row: "p2-id", size: "narrow" },
  { key: "party2IdNumber", group: "pihak2", label: "Nomor Identitas", type: "text", placeholder: "Nomor identitas", word: "Party2IdNumber", row: "p2-id" },
];

/** Nama variabel Word yang dikelola form (tidak ditampilkan lagi di "variabel lain"). */
export const PARTY_WORD_KEYS: string[] = PARTY_FIELDS.map((f) => f.word!).filter(Boolean);

export const PARTY_LOCKED_NOTE = "Nomor dokumen terkunci demi jejak audit — tidak bisa diubah dari sini.";

// ---------------------------------------------------------------------------
// Tanggal (dipakai server.ts juga — dipindah ke sini agar satu definisi)
// ---------------------------------------------------------------------------
export const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

export const formatShortDateId = (iso: string): string => {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? iso : d.toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" });
};

const ID_MONTHS: Record<string, string> = { januari: "01", februari: "02", maret: "03", april: "04", mei: "05", juni: "06", juli: "07", agustus: "08", september: "09", oktober: "10", november: "11", desember: "12" };

/** "Rabu, 1 Juli 2026" / "31 Desember 2026" -> "2026-07-01"; null bila tidak terbaca. */
export function parseIdDate(text: string): string | null {
  const m = /(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/.exec(text || "");
  if (!m) return null;
  const mm = ID_MONTHS[m[2].toLowerCase()];
  const d = Number(m[1]);
  if (!mm || d < 1 || d > 31) return null;
  return `${m[3]}-${mm}-${String(d).padStart(2, "0")}`;
}

/** YYYY-MM-DD yang benar-benar ada di kalender (2026-02-31 ditolak). */
export function isRealIsoDate(s: string): boolean {
  if (!ISO_DATE.test(s)) return false;
  const d = new Date(`${s}T00:00:00Z`);
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}

// ---------------------------------------------------------------------------
// Bentuk form (sama dgn handleOpenEditPartyModal di App.tsx)
// ---------------------------------------------------------------------------
export function partyFormOf(c: Partial<Contract>): PartyForm {
  return {
    title: c.title || "",
    docType: c.docType || "",
    startDate: c.startDate || "",
    endDate: c.endDate || "",
    party1Address: c.party1Address || "",
    party1Position: c.party1Position || "",
    party1IdLabel: c.party1IdLabel || "",
    party1IdNumber: c.party1IdNumber || "",
    party2Name: c.party2Name || "",
    party2Type: c.party2Type || "",
    party2Address: c.party2Address || "",
    party2Position: c.party2Position || "",
    party2IdLabel: c.party2IdLabel || "",
    party2IdNumber: c.party2IdNumber || "",
    party2LogoUrl: c.party2LogoUrl || "",
    party2LogoKey: c.party2LogoKey || "",
    party2LogoMimeType: c.party2LogoMimeType || "",
  };
}

/** Penanda perubahan murah: berubah bila SALAH SATU nilai form berubah. */
export function partyRevOf(c: Partial<Contract>): string {
  const f = partyFormOf(c);
  const s = PARTY_FIELDS.map((d) => f[d.key]).concat(PARTY_LOGO_KEYS.map((k) => f[k])).join("\u0001");
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0;
  return h.toString(36) + ":" + s.length;
}

// ---------------------------------------------------------------------------
// Boleh diubah? (aturan = isContractEditable di App.tsx)
// ---------------------------------------------------------------------------
export function partyEditability(
  c: Pick<Contract, "status" | "revisionStage" | "approvalSubmittedById" | "createdById">,
  user: { id: string; role: string },
  canEditPerm: boolean,
): { editable: boolean; reason?: string } {
  if (c.status !== "Draft") return { editable: false, reason: "Data pihak & kontrak hanya bisa diubah selagi dokumen berstatus Draft." };
  if (!canEditPerm) return { editable: false, reason: "Anda tidak punya izin mengubah dokumen ini." };
  if (c.revisionStage) {
    const adminLike = user.role === "super_admin";
    const owner = c.approvalSubmittedById || c.createdById;
    const ownerSide = adminLike || (owner ? owner === user.id : canEditPerm);
    if (!ownerSide) return { editable: false, reason: "Saat revisi, hanya pembuat/pengaju dokumen (atau admin) yang dapat mengubah data." };
  }
  return { editable: true };
}

// ---------------------------------------------------------------------------
// Tampilan form untuk plugin (label, placeholder dinamis, opsi dropdown)
// ---------------------------------------------------------------------------
export interface PartyCompany { name?: string; address?: string; repTitle?: string }

export interface PartyViewField {
  key: PartyTextKey;
  label: string;
  type: "text" | "date" | "select";
  required: boolean;
  placeholder: string;
  value: string;
  word?: string;
  row?: string;
  size?: "narrow";
  options?: string[];
}
export interface PartyViewGroup { id: PartyGroupId; title: string; fields: PartyViewField[] }
export interface PartyView {
  rev: string;
  groups: PartyViewGroup[];
  locked: { key: "contractNumber"; label: string; value: string; note: string };
  logo: { label: string; hint: string; url: string; key: string; mimeType: string };
  wordKeys: string[];
}

export function buildPartyView(c: Partial<Contract>, company: PartyCompany, partyTypes: string[]): PartyView {
  const form = partyFormOf(c);
  const dyn: Partial<Record<PartyTextKey, string>> = {
    party1Address: `Alamat (kosong = "${company.address || "alamat perusahaan"}")`,
    party1Position: `Jabatan Penandatanganan (kosong = "${company.repTitle || "jabatan perwakilan"}")`,
  };
  const titles: Record<PartyGroupId, string> = {
    kontrak: "Informasi Kontrak",
    pihak1: `Pihak Pertama (${company.name || "Perusahaan"})`,
    pihak2: "Pihak Kedua",
  };
  const groups: PartyViewGroup[] = (["kontrak", "pihak1", "pihak2"] as PartyGroupId[]).map((g) => ({
    id: g,
    title: titles[g],
    fields: PARTY_FIELDS.filter((f) => f.group === g).map((f) => {
      const out: PartyViewField = {
        key: f.key, label: f.label, type: f.type, required: !!f.required,
        placeholder: dyn[f.key] ?? f.placeholder ?? "",
        value: form[f.key], word: f.word, row: f.row, size: f.size,
      };
      if (f.type === "select") {
        // Sama dgn modal: nilai saat ini tetap tampil walau tidak ada di daftar master.
        const cur = form[f.key];
        out.options = (cur && !partyTypes.includes(cur) ? [cur] : []).concat(partyTypes);
      }
      return out;
    }),
  }));
  return {
    rev: partyRevOf(c),
    groups,
    locked: { key: "contractNumber", label: "Nomor Dokumen", value: c.contractNumber || "", note: PARTY_LOCKED_NOTE },
    logo: {
      label: "Logo Pihak Kedua",
      hint: "Opsional. Kosong = kop surat cuma tampilkan logo Pihak Pertama saja — normal, bukan error. Dipakai utk kop surat dua-logo pada kontrak kerjasama B2B antar dua perusahaan independen.",
      url: form.party2LogoUrl, key: form.party2LogoKey, mimeType: form.party2LogoMimeType,
    },
    wordKeys: PARTY_WORD_KEYS,
  };
}

// ---------------------------------------------------------------------------
// Validasi + normalisasi perubahan (padanan handleSaveEditPartyModal)
// ---------------------------------------------------------------------------
export interface EnglishInvalidation { title: boolean; docType: boolean; preamble: boolean }

export type PartyPatchResult =
  | { ok: true; next: Partial<Record<PartyKey, string | undefined>>; changed: PartyKey[]; invalidate: EnglishInvalidation }
  | { ok: false; error: string; field?: PartyKey | "contractNumber"; status: 400 | 409 };

const MAX_LEN = 1000;
const LABEL_OF: Record<string, string> = Object.fromEntries([
  ...PARTY_FIELDS.map((f) => [f.key, f.label]),
  ["party2LogoUrl", "Logo Pihak Kedua"], ["party2LogoKey", "Logo Pihak Kedua"], ["party2LogoMimeType", "Logo Pihak Kedua"],
]);
const GROUP_LABEL: Record<string, string> = { kontrak: "", pihak1: "Pihak Pertama — ", pihak2: "Pihak Kedua — " };
/** Nama field yang enak dibaca untuk audit/pesan: "Pihak Kedua — Alamat". */
export function partyFieldLabel(key: string): string {
  const def = PARTY_FIELDS.find((f) => f.key === key);
  return (def ? GROUP_LABEL[def.group] : "") + (LABEL_OF[key] || key);
}

const PREAMBLE_KEYS: PartyKey[] = [
  "startDate", "party1Address", "party1Position", "party1IdLabel", "party1IdNumber",
  "party2Name", "party2Address", "party2Position", "party2IdLabel", "party2IdNumber",
];

/**
 * @param current  keadaan kontrak sekarang
 * @param patch    nilai baru (hanya field yang dikirim yang diproses)
 * @param base     (opsional) nilai yang dilihat pengguna saat membuka form; bila diberikan,
 *                 field yang nilainya di server sudah berbeda dari `base` ditolak (409) —
 *                 mencegah panel menimpa perubahan yang dibuat orang lain di aplikasi utama.
 */
export function normalizePartyPatch(
  current: Partial<Contract>,
  patch: Record<string, unknown>,
  ctx: { partyTypes: string[]; base?: Record<string, unknown> },
): PartyPatchResult {
  if (patch.contractNumber !== undefined && String(patch.contractNumber) !== (current.contractNumber || "")) {
    return { ok: false, field: "contractNumber", status: 400, error: PARTY_LOCKED_NOTE };
  }
  const cur = partyFormOf(current);
  const allowed = new Set<string>([...PARTY_TEXT_KEYS, ...PARTY_LOGO_KEYS]);
  const next: Partial<Record<PartyKey, string | undefined>> = {};
  const changed: PartyKey[] = [];

  for (const [k, raw] of Object.entries(patch)) {
    if (k === "contractNumber") continue;
    if (!allowed.has(k)) return { ok: false, status: 400, error: `Field tidak dikenal: ${k}` };
    if (raw !== undefined && raw !== null && typeof raw !== "string") return { ok: false, status: 400, field: k as PartyKey, error: `${partyFieldLabel(k)} harus berupa teks.` };
  }

  for (const key of [...PARTY_TEXT_KEYS, ...PARTY_LOGO_KEYS] as PartyKey[]) {
    if (!(key in patch)) continue;
    const value = String(patch[key] ?? "").trim();
    if (value.length > MAX_LEN) return { ok: false, status: 400, field: key, error: `${partyFieldLabel(key)} terlalu panjang (maks ${MAX_LEN} karakter).` };

    if (ctx.base && key in ctx.base && String(ctx.base[key] ?? "").trim() !== cur[key].trim()) {
      return { ok: false, status: 409, field: key, error: `${partyFieldLabel(key)} baru saja diubah di tempat lain. Data dimuat ulang — ulangi perubahan Anda bila masih perlu.` };
    }

    if (key === "title" && !value) return { ok: false, status: 400, field: key, error: "Judul dokumen tidak boleh kosong" };
    if (key === "party2Name" && !value) return { ok: false, status: 400, field: key, error: "Nama Pihak Kedua tidak boleh kosong" };
    if (key === "startDate" || key === "endDate") {
      if (!value) return { ok: false, status: 400, field: key, error: `${partyFieldLabel(key)} tidak boleh kosong` };
      if (!isRealIsoDate(value)) return { ok: false, status: 400, field: key, error: `${partyFieldLabel(key)} bukan tanggal yang valid (format YYYY-MM-DD).` };
    }
    if (key === "party2Type") {
      // Modal: party2Type.trim() || nilai lama. Nilai baru harus dari daftar master (atau nilai lama).
      if (!value) continue;
      if (!ctx.partyTypes.includes(value) && value !== cur.party2Type) return { ok: false, status: 400, field: key, error: `Tipe Pihak Kedua "${value}" tidak ada di daftar.` };
    }
    if (key === "party2LogoMimeType" && value && !/^image\/(png|jpe?g)$/.test(value)) {
      return { ok: false, status: 400, field: key, error: "Logo harus berformat PNG atau JPEG" };
    }

    next[key] = OPTIONAL_KEYS.includes(key) ? (value || undefined) : value;
    if (value !== cur[key].trim()) changed.push(key);
  }

  const touched = (ks: PartyKey[]) => changed.some((k) => ks.includes(k));
  return {
    ok: true, next, changed,
    invalidate: { title: changed.includes("title"), docType: changed.includes("docType"), preamble: touched(PREAMBLE_KEYS) },
  };
}

/** Terapkan hasil normalizePartyPatch ke kontrak (mutasi) + buang terjemahan Inggris yang basi. */
export function applyPartyPatch(c: Contract, r: Extract<PartyPatchResult, { ok: true }>): void {
  for (const [k, v] of Object.entries(r.next)) (c as any)[k] = v;
  clearStaleEnglish(c, r.invalidate);
}

export function clearStaleEnglish(c: Partial<Contract>, inv: EnglishInvalidation): void {
  if (inv.title) c.titleEn = undefined;
  if (inv.docType) c.docTypeEn = undefined;
  if (inv.preamble) { c.preambleEn = undefined; c.preambleEnOriginal = undefined; c.translationWarnings = undefined; }
}

/** Bandingkan dua keadaan kontrak -> terjemahan Inggris mana yang basi (untuk perubahan dari Word). */
export function englishInvalidationBetween(before: Partial<Contract>, after: Partial<Contract>): EnglishInvalidation {
  const a = partyFormOf(before), b = partyFormOf(after);
  const diff = (k: PartyKey) => a[k] !== b[k];
  return { title: diff("title"), docType: diff("docType"), preamble: PREAMBLE_KEYS.some(diff) };
}

/** "[Alamat belum diisi]" dst. — penanda kosong yang ditulis sistem ke dokumen, bukan isian pengguna. */
export const isPlaceholderValue = (v: string): boolean => /^\[[^\]]*belum diisi\]$/.test(v.trim());

// ---------------------------------------------------------------------------
// Pemetaan field -> variabel Word (untuk field BARU; kunci lama tetap dihitung
// di desiredWordValues server.ts supaya perilakunya tidak berubah)
// ---------------------------------------------------------------------------
export function partyWordValues(c: Partial<Contract>, company: PartyCompany): Record<string, string> {
  const ph = (label: string) => `[${label} belum diisi]`;
  return {
    DocumentTitle: c.title || ph("Judul dokumen"),
    DocumentType: c.docType || ph("Jenis dokumen"),
    Party1Position: c.party1Position || company.repTitle || ph("Jabatan Pihak Pertama"),
    Party1IdLabel: c.party1IdLabel || "No. Identitas",
    Party1IdNumber: c.party1IdNumber || ph("Nomor identitas Pihak Pertama"),
    Party2Type: c.party2Type || ph("Tipe Pihak Kedua"),
    Party2Position: c.party2Position || ph("Jabatan Pihak Kedua"),
    Party2IdLabel: c.party2IdLabel || "No. Identitas",
    Party2IdNumber: c.party2IdNumber || ph("Nomor identitas Pihak Kedua"),
  };
}

// ---------------------------------------------------------------------------
// Rekonsiliasi tiga arah: data sistem <-> isi variabel di dokumen Word,
// dengan potret terakhir (contract.wordSync) sebagai pembanding.
//  - data sistem berubah  -> dorong ke dokumen
//  - isi dokumen berubah  -> bila variabel bisa dibaca dgn aman, salin ke data sistem
// Dipakai dua jalur: saat editor dibuka (membaca berkas .docx) dan saat Panel
// CLM sinkron langsung dari editor yang sedang terbuka (mode `live`).
// ---------------------------------------------------------------------------
export interface ReconcileOptions {
  /** Kontrak yang boleh dimutasi (salinan kerja). */
  contract: Contract;
  /** Nilai variabel yang ada di dokumen sekarang (kunci tanpa "@en"). */
  inDoc: Record<string, string>;
  /** Nilai yang SEHARUSNYA tampil menurut data sistem. */
  want: Record<string, string>;
  /** Potret terakhir; dimutasi menjadi potret baru. */
  snap: Record<string, string>;
  /** Alamat/jabatan perusahaan — untuk menghindari mengunci nilai bawaan perusahaan jadi "ubahan kontrak". */
  company?: PartyCompany;
  /**
   * Mode langsung (Panel CLM): `pullable` membatasi variabel yang boleh DITARIK dari dokumen (hanya yang
   * isinya stabil, bukan sedang diketik). Tanpa potret baru untuk variabel yang didorong — potret baru
   * dicatat pada sinkron berikutnya setelah dokumen terbukti memuat nilainya (kalau penerapan di editor
   * gagal, data sistem tidak "dikembalikan" oleh isi dokumen yang lama).
   */
  live?: { pullable: (key: string) => boolean };
}
export interface ReconcileOutcome { push: Record<string, string>; pushed: string[]; pulled: string[] }

const NOT_PULLABLE_FROM_DOC = new Set(["Party1Name", "Party1Representative", "Party1RepTitle"]);

export function reconcileCore(o: ReconcileOptions): ReconcileOutcome {
  const { contract: c, inDoc, want, snap, company } = o;
  const push: Record<string, string> = {};
  const pushed: string[] = [], pulled: string[] = [];
  const real = (v: string) => !!v.trim() && !isPlaceholderValue(v); // "[... belum diisi]" bukan data

  // Variabel yang boleh ditulis balik ke kolom kontrak saat isi dokumen diubah orang.
  const identity: Record<string, (v: string) => void> = {
    Party2Name: (v) => { if (v.trim()) c.party2Name = v.trim(); },
    Party2Address: (v) => { if (real(v)) c.party2Address = v.trim(); },
    Party1Address: (v) => {
      if (!real(v)) return;
      // Sama dgn alamat perusahaan = ikuti perusahaan (tidak dibekukan sbg ubahan kontrak).
      c.party1Address = company?.address && v.trim() === company.address.trim() ? undefined : v.trim();
    },
    DocumentTitle: (v) => { if (real(v)) c.title = v.trim(); },
    DocumentType: (v) => { if (real(v)) c.docType = v.trim(); },
    Party1Position: (v) => { if (real(v)) c.party1Position = company?.repTitle && v.trim() === company.repTitle.trim() ? undefined : v.trim(); },
    Party1IdLabel: (v) => { if (real(v)) c.party1IdLabel = v.trim() === "No. Identitas" ? undefined : v.trim(); },
    Party1IdNumber: (v) => { if (real(v)) c.party1IdNumber = v.trim(); },
    Party2Position: (v) => { if (real(v)) c.party2Position = v.trim(); },
    Party2IdLabel: (v) => { if (real(v)) c.party2IdLabel = v.trim() === "No. Identitas" ? undefined : v.trim(); },
    Party2IdNumber: (v) => { if (real(v)) c.party2IdNumber = v.trim(); },
  };

  for (const [key, docVal] of Object.entries(inDoc)) {
    const sys = want[key];
    if (sys === undefined) continue;
    if (snap[key] === undefined) { snap[key] = docVal; continue; } // pertama kali: catat saja
    if (sys !== snap[key]) {
      // Data sistem berubah sejak sinkron terakhir -> dokumen mengikuti.
      if (sys !== docVal) { push[key] = sys; pushed.push(key); if (!o.live) snap[key] = sys; }
      else snap[key] = sys; // dokumen sudah memuat nilainya
    } else if (docVal !== snap[key]) {
      if (o.live && !o.live.pullable(key)) continue; // masih diketik / belum stabil: tunggu
      // Isi dokumen berubah, data sistem tidak: salin balik bila bisa dibaca dengan aman.
      const iso = parseIdDate(docVal);
      if (key === "StartDate" && iso) { c.startDate = iso; pulled.push(key); snap[key] = docVal; }
      else if (key === "EndDate" && iso) { c.endDate = iso; if (typeof c.variables?.EndDate === "string") (c.variables as any).EndDate = iso; pulled.push(key); snap[key] = docVal; }
      else if (identity[key]) { identity[key](docVal); pulled.push(key); snap[key] = docVal; }
      else if (typeof c.variables?.[key] === "string" && ISO_DATE.test(c.variables[key] as string)) {
        if (iso) { (c.variables as any)[key] = iso; pulled.push(key); }
        snap[key] = docVal;
      } else if (typeof c.variables?.[key] === "string" && !(key in { StartDate: 1, EndDate: 1 }) && !NOT_PULLABLE_FROM_DOC.has(key)) {
        (c.variables as any)[key] = docVal; pulled.push(key); snap[key] = docVal;
      } else snap[key] = docVal; // tidak bisa dibaca aman (mis. identitas Pihak Pertama dari pengaturan): dibiarkan
    }
  }
  return { push, pushed, pulled };
}
