// Master Mitra/Operator bawaan untuk Arsip Eksternal (PT › Jenis Dokumen ›
// Platform › Mitra/Operator). Hanya SEED: begitu admin menyimpan
// `operatorGroups` di Konfigurasi, daftar tersimpan itulah yang dipakai.
//
// Aturan berlaku grup (dicocokkan tanpa peduli huruf besar/kecil):
//  - `companies`: PT (nama/kode/id perusahaan) pemilik grup. Kosong = semua PT.
//  - `platforms`: platform yang dicakup. Kosong = SEMUA platform milik PT tsb.
// Bila ada grup yang menyebut platform terpilih secara EKSPLISIT, hanya grup
// eksplisit itu yang dipakai (mis. platform Bank punya mitra sendiri, jadi
// daftar GKM tidak ikut). Kalau tidak ada, dipakai grup yang platformnya kosong.
export type OperatorGroup = { id: string; name: string; companies: string[]; platforms: string[]; operators: string[] };

export const DEFAULT_OPERATOR_GROUPS: OperatorGroup[] = [
  {
    id: "og-tbk",
    name: "Mitra/Operator TBK",
    companies: [],
    platforms: ["Tiketux"],
    operators: [
      "Abbata", "ADIBUZZ", "ANS", "Antar Lintas Sumatrra", "Aragon Transport", "Arnes Shuttle", "ATS", "Banyumili",
      "Baraya Travel", "Besttrans", "Bhinneka Sangkuriang", "Bina Sarana", "BTM Travel", "Cahaya Shuttle", "CGTrans",
      "Cititrans", "City Trans Utama", "aytrans", "E-Shuttle", "Go Shuttle", "Gracias Shuttle", "Harum BSI",
      "Jackal Holidays", "Joglosemar", "Kalisari", "Karunia Bakti", "Kencana", "KPM Trans", "Krakaline", "Kruzz",
      "Kupu-Kupu Ayu", "Lintas Shuttle", "Minanga Express", "MRTrans", "MRZ Trans", "MSTrans", "Nam Trans", "Namaste",
      "Nonumy", "NPM", "Nussaraya Trans", "Putra Remaja", "Rama Sakti", "Raputri Shuttle", "Rimba Raya Premium Travel",
      "Riyan Transport", "Royal Kencana", "Rukun Trans", "Sabila Shuttle", "Sadia Trans", "Sari Harum", "Sehati",
      "Selamat Trans", "Semeru Trans", "Sumber Alam", "Sumber Jaya Executive Shuttle", "Sunjaya Trans", "Transkita",
      "Tri Kusuma", "Unicorn Indorent", "Urban Travel", "VTrans", "WB Shuttle", "Wisata Komodo", "Yanti Group Shuttle",
      "ZTrans", "Plitamas", "AO Shuttle",
    ],
  },
  {
    id: "og-bank",
    name: "Mitra/Operator Bank",
    companies: [],
    platforms: ["Bank"],
    operators: [
      "PT BANK NEO COMMERCE TBK", "PT ESPAY DEBIT INDONESIA KOE/DANA", "PT Nusa SATU INTI ARTA/DOKU",
      "PT FLIPTECH LENTERA INSPIRASI PERTIWI/FLIP", "PT MEDIA INDONUSA/PASPEY", "PT MIDTRANS", "PT MITRA SENTRAL TERPADU",
      "PT BANK BPTPN TBK/JENUSPAY", "PT MMC TEKNOLOGI NUSANTARA", "PT IYONPAY NETWORK/NICEPAY",
      "PT WAHANA PEMBAYARAN DIGITAL/PAYLABS", "PT PRISMALINK INTERNASIONAL/PRISMALINK", "FLIP",
      "PT TELEKOMUNIKASI INDONESIA, TBK", "PT BIMASAKTI MULTISINARGI/WINPAY", "KREDIVO", "SHOPEEPAY", "INDODANA",
      "ISAKU INDOMARET", "BCA", "BANK PERMATA", "PT SURGANYA MOTOR INDONESIA/PLANET BAN",
    ],
  },
  {
    id: "og-gkm",
    name: "Mitra/Operator GKM",
    // Semua platform milik PT GKM, kecuali yang punya grup sendiri (Bank).
    companies: ["GKM"],
    platforms: [],
    operators: [
      "PRIMER KOPERASI ANGKATAN UDARA (PRIMKOPAU)/RS SALAMUN", "RS ADVEN BANDUNG", "RS HARMINA PASTEUR BANDUNG",
      "RS HARMINA ANCAMANI BANDUNG", "RSUD BANDUNG KIWARI", "RSUD KOTA BANDUNG", "RS PINDAD BANDUNG",
      "RS KEBON JATI BANDUNG", "RSUD CIBABAT CIMAHI", "RS MITRA KASIH CIMAHI", "RS KASIH BUNDA CIMAHI",
      "RS MURNI TEGUH BANDUNG", "RS OETOMO/UTOMO HOSPITAL BANDUNG", "MAYAPADA HOSPITAL BANDUNG",
      "RS CAHAYA KAWALUYAAN", "RS AL ISLAM BANDUNG", "RS SANTOBOROMEUS",
    ],
  },
];

// Rapikan kiriman dari klien: buang kosong/duplikat, batasi panjang.
export function sanitizeOperatorGroups(input: unknown): OperatorGroup[] {
  if (!Array.isArray(input)) return [];
  const clean = (arr: unknown, max: number) => {
    const out: string[] = [];
    (Array.isArray(arr) ? arr : []).forEach((x) => {
      const v = String(x ?? "").trim().slice(0, 120);
      if (v && !out.some((y) => y.toLowerCase() === v.toLowerCase())) out.push(v);
    });
    return out.slice(0, max);
  };
  return input.slice(0, 50).map((g: any, i) => ({
    id: String(g?.id || `og-${Date.now()}-${i}`).slice(0, 60),
    name: String(g?.name || "").trim().slice(0, 80) || `Grup ${i + 1}`,
    companies: clean(g?.companies, 50),
    platforms: clean(g?.platforms, 50),
    operators: clean(g?.operators, 1000),
  }));
}
