// ---------------------------------------------------------------------------
// Pengolah gambar TTD Digital (Konfigurasi > Kelola Perusahaan).
//
// Dua jalan masuk, satu hasil: PNG TRANSPARAN yang sudah dipotong rapat
// (tanpa ruang kosong) dan dikecilkan ke ukuran wajar — itulah satu-satunya
// yang dikirim ke server, sehingga penempatannya di kolom TTD dokumen bisa
// rapi untuk semua format TTD.
//
//  1. Foto tanda tangan di kertas → latar kertas dihapus (transparan), tinta
//     dijadikan warna seragam, noda/tepi meja dibuang, lalu dipotong rapat.
//  2. Gambar langsung di kanvas → cukup dipotong rapat (kanvas sudah bening).
//
// Bagian "inti" di atas file ini MURNI (hanya hitungan piksel RGBA, tanpa DOM)
// supaya bisa diuji lewat scripts/signature-image-test.ts. Fungsi yang
// memakai kanvas/DOM ada di bagian bawah dan hanya dipanggil dari browser.
// ---------------------------------------------------------------------------

export interface RgbaImage {
  data: Uint8ClampedArray;
  width: number;
  height: number;
}

export interface SignatureResult {
  blob: Blob;
  dataUrl: string;
  width: number;
  height: number;
}

/** Sisi terpanjang foto saat dianalisis (px). Foto ponsel 4000px+ diperkecil dulu. */
export const SIGNATURE_ANALYSIS_MAX_SIDE = 1600;
/** Batas ukuran PNG hasil akhir (px). Cukup tajam utk cetak ±4 cm lebar. */
export const SIGNATURE_OUTPUT_MAX_W = 900;
export const SIGNATURE_OUTPUT_MAX_H = 450;
/** Batas byte PNG hasil akhir — sedikit di bawah batas server (400 KB). */
export const SIGNATURE_MAX_BYTES = 380 * 1024;
/** Warna tinta goresan di kanvas gambar langsung. */
export const SIGNATURE_DRAW_INK = "#10172a";

// Rentang kontras "tinta" (0 = sama terang dgn kertas, 1 = hitam pekat).
// Di bawah LO → bening (noise kertas), di atas HI → tinta penuh.
const INK_LO = 0.12;
const INK_HI = 0.45;
const SOLID_ALPHA = 128;
const BOUNDS_ALPHA = 64;
const MIN_COMPONENT_PX = 12;
const MIN_SIDE_PX = 24;

const clamp = (v: number, lo: number, hi: number) => (v < lo ? lo : v > hi ? hi : v);
const smoothstep = (lo: number, hi: number, v: number) => {
  const t = clamp((v - lo) / (hi - lo), 0, 1);
  return t * t * (3 - 2 * t);
};

/** Ukuran gambar `w`×`h` yang muat di `maxW`×`maxH` tanpa mengubah rasio (boleh diperbesar). */
export function fitWithin(w: number, h: number, maxW: number, maxH: number): { width: number; height: number } {
  if (!(w > 0) || !(h > 0) || !(maxW > 0) || !(maxH > 0)) return { width: 0, height: 0 };
  const s = Math.min(maxW / w, maxH / h);
  return { width: Math.max(1, Math.round(w * s)), height: Math.max(1, Math.round(h * s)) };
}

/** Sudah ada bagian bening yang berarti (mis. PNG tanda tangan siap pakai)? */
export function hasTransparency(img: RgbaImage): boolean {
  const { data } = img;
  let transparent = 0;
  for (let i = 3; i < data.length; i += 4) if (data[i] < 250) transparent++;
  return transparent > img.width * img.height * 0.02;
}

function percentile(hist: Uint32Array, count: number, q: number): number {
  const target = count * q;
  let acc = 0;
  for (let v = 0; v < 256; v++) {
    acc += hist[v];
    if (acc >= target) return v;
  }
  return 255;
}

/**
 * Perkiraan tingkat terang KERTAS di tiap piksel (peta latar). Foto ponsel
 * jarang rata — ada bayangan/gradasi — jadi latar dihitung per blok (persentil
 * ke-90 kecerahan blok = warna kertas) lalu diinterpolasi halus. Goresan
 * tanda tangan tipis, jadi tidak menggeser angka ini.
 */
function paperLevels(lum: Uint8Array, w: number, h: number): Float32Array {
  const block = Math.max(24, Math.round(Math.min(w, h) / 10));
  const gx = Math.ceil(w / block);
  const gy = Math.ceil(h / block);

  const globalHist = new Uint32Array(256);
  for (let i = 0; i < lum.length; i++) globalHist[lum[i]]++;
  const globalPaper = percentile(globalHist, lum.length, 0.9);

  const grid = new Float32Array(gx * gy);
  const hist = new Uint32Array(256);
  for (let by = 0; by < gy; by++) {
    for (let bx = 0; bx < gx; bx++) {
      hist.fill(0);
      const x0 = bx * block, x1 = Math.min(w, x0 + block);
      const y0 = by * block, y1 = Math.min(h, y0 + block);
      let count = 0;
      for (let y = y0; y < y1; y++) {
        for (let x = x0; x < x1; x++) { hist[lum[y * w + x]]++; count++; }
      }
      // Lantai = separuh kertas global: blok yang nyaris tertutup tinta tidak
      // dianggap "kertas gelap" (yang membuat goresannya hilang).
      grid[by * gx + bx] = Math.max(percentile(hist, count, 0.9), globalPaper * 0.5);
    }
  }

  const x0s = new Int32Array(w), x1s = new Int32Array(w);
  const txs = new Float32Array(w);
  for (let x = 0; x < w; x++) {
    const fx = (x + 0.5) / block - 0.5;
    const a = clamp(Math.floor(fx), 0, gx - 1);
    x0s[x] = a;
    x1s[x] = Math.min(a + 1, gx - 1);
    txs[x] = clamp(fx - a, 0, 1);
  }
  const out = new Float32Array(w * h);
  for (let y = 0; y < h; y++) {
    const fy = (y + 0.5) / block - 0.5;
    const ya = clamp(Math.floor(fy), 0, gy - 1);
    const yb = Math.min(ya + 1, gy - 1);
    const ty = clamp(fy - ya, 0, 1);
    for (let x = 0; x < w; x++) {
      const top = grid[ya * gx + x0s[x]] * (1 - txs[x]) + grid[ya * gx + x1s[x]] * txs[x];
      const bot = grid[yb * gx + x0s[x]] * (1 - txs[x]) + grid[yb * gx + x1s[x]] * txs[x];
      out[y * w + x] = top * (1 - ty) + bot * ty;
    }
  }
  return out;
}

interface InkPiece { members: number[]; minX: number; maxX: number; minY: number; maxY: number }

/**
 * Buang "bukan tanda tangan":
 *  - bidang gelap padat, mis. tepi meja yang ikut terfoto di pinggir kertas.
 *    Goresan itu tipis (isi < 50% dari kotak pembatasnya), sedangkan tepi
 *    meja/bayangan pekat itu padat — itulah pembedanya;
 *  - noda debu: gumpalan kecil yang JAUH dari badan tanda tangan utama
 *    (titik di dekat badan, mis. titik huruf "i", dipertahankan).
 */
function dropArtifacts(alpha: Uint8Array, w: number, h: number): void {
  const total = w * h;
  const seen = new Uint8Array(total);
  const stack = new Int32Array(total);
  const cleared = new Uint8Array(total);

  // 1. Pecah menjadi gumpalan-gumpalan tinta yang saling terhubung.
  const pieces: InkPiece[] = [];
  for (let start = 0; start < total; start++) {
    if (alpha[start] < SOLID_ALPHA || seen[start]) continue;
    const piece: InkPiece = { members: [], minX: w, maxX: 0, minY: h, maxY: 0 };
    let sp = 0;
    stack[sp++] = start;
    seen[start] = 1;
    while (sp > 0) {
      const p = stack[--sp];
      piece.members.push(p);
      const x = p % w, y = (p - x) / w;
      if (x < piece.minX) piece.minX = x;
      if (x > piece.maxX) piece.maxX = x;
      if (y < piece.minY) piece.minY = y;
      if (y > piece.maxY) piece.maxY = y;
      for (let dy = -1; dy <= 1; dy++) {
        const ny = y + dy;
        if (ny < 0 || ny >= h) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nx = x + dx;
          if (nx < 0 || nx >= w || (dx === 0 && dy === 0)) continue;
          const q = ny * w + nx;
          if (!seen[q] && alpha[q] >= SOLID_ALPHA) { seen[q] = 1; stack[sp++] = q; }
        }
      }
    }
    pieces.push(piece);
  }

  // 2. Putuskan mana yang dibuang.
  const drop = new Set<InkPiece>();
  const boxAreaOf = (p: InkPiece) => (p.maxX - p.minX + 1) * (p.maxY - p.minY + 1);
  for (const p of pieces) {
    if (p.members.length > total * 0.01 && p.members.length / boxAreaOf(p) > 0.5) drop.add(p);
  }
  let main: InkPiece | null = null;
  for (const p of pieces) {
    if (!drop.has(p) && (!main || p.members.length > main.members.length)) main = p;
  }
  if (main) {
    const reach = 0.12 * Math.max(main.maxX - main.minX + 1, main.maxY - main.minY + 1) + 8;
    for (const p of pieces) {
      if (drop.has(p) || p === main) continue;
      const near = p.maxX >= main.minX - reach && p.minX <= main.maxX + reach
        && p.maxY >= main.minY - reach && p.minY <= main.maxY + reach;
      if (p.members.length < MIN_COMPONENT_PX || (!near && p.members.length < main.members.length * 0.08)) drop.add(p);
    }
  }

  if (drop.size === 0) return;
  for (const p of drop) {
    for (const m of p.members) { alpha[m] = 0; cleared[m] = 1; }
  }
  // Bersihkan juga tepi lembut (antialias) di sekeliling bagian yang dibuang.
  for (let round = 0; round < 3; round++) {
    const grow: number[] = [];
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const p = y * w + x;
        if (cleared[p] || alpha[p] === 0 || alpha[p] >= SOLID_ALPHA) continue;
        let near = false;
        for (let dy = -1; dy <= 1 && !near; dy++) {
          const ny = y + dy;
          if (ny < 0 || ny >= h) continue;
          for (let dx = -1; dx <= 1; dx++) {
            const nx = x + dx;
            if (nx >= 0 && nx < w && cleared[ny * w + nx]) { near = true; break; }
          }
        }
        if (near) grow.push(p);
      }
    }
    for (const p of grow) { alpha[p] = 0; cleared[p] = 1; }
  }
}

/** Kotak pembatas piksel yang cukup pekat; null bila kosong. */
function inkBounds(img: RgbaImage, alphaOf: (i: number) => number): { x: number; y: number; w: number; h: number } | null {
  const { width: w, height: h } = img;
  let minX = w, maxX = -1, minY = h, maxY = -1;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      if (alphaOf(y * w + x) >= BOUNDS_ALPHA) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) return null;
  return { x: minX, y: minY, w: maxX - minX + 1, h: maxY - minY + 1 };
}

/** Potong rapat di sekitar kotak goresan (+ bantalan tipis), sisanya dibuang. */
function cropWithPadding(
  img: RgbaImage,
  box: { x: number; y: number; w: number; h: number },
  pixelAt: (i: number, out: Uint8ClampedArray, o: number) => void,
): RgbaImage {
  const pad = Math.max(2, Math.round(Math.max(box.w, box.h) * 0.03));
  const x0 = Math.max(0, box.x - pad), y0 = Math.max(0, box.y - pad);
  const x1 = Math.min(img.width, box.x + box.w + pad), y1 = Math.min(img.height, box.y + box.h + pad);
  const w = x1 - x0, h = y1 - y0;
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) pixelAt((y0 + y) * img.width + (x0 + x), data, (y * w + x) * 4);
  }
  return { data, width: w, height: h };
}

/**
 * Gambar yang SUDAH bening (hasil kanvas / PNG tanda tangan): cukup dipotong
 * rapat di sekitar goresan. null bila tak ada goresan berarti.
 */
export function cropToInk(img: RgbaImage): RgbaImage | null {
  const box = inkBounds(img, (i) => img.data[i * 4 + 3]);
  if (!box || Math.max(box.w, box.h) < 8) return null; // noda tunggal, bukan tanda tangan
  return cropWithPadding(img, box, (i, out, o) => {
    out[o] = img.data[i * 4]; out[o + 1] = img.data[i * 4 + 1];
    out[o + 2] = img.data[i * 4 + 2]; out[o + 3] = img.data[i * 4 + 3];
  });
}

/**
 * Foto tanda tangan di kertas → PNG transparan terpotong rapat berwarna tinta
 * seragam. null = tidak terdeteksi tanda tangan (foto kosong / terlalu pucat).
 */
export function extractSignature(img: RgbaImage): RgbaImage | null {
  if (hasTransparency(img)) return cropToInk(img); // sudah bening, jangan diutak-atik

  const { width: w, height: h, data } = img;
  const total = w * h;
  const lum = new Uint8Array(total);
  for (let i = 0; i < total; i++) {
    lum[i] = Math.round(0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2]);
  }
  const paper = paperLevels(lum, w, h);

  const alpha = new Uint8Array(total);
  for (let i = 0; i < total; i++) {
    const contrast = (paper[i] - lum[i]) / Math.max(paper[i], 40);
    alpha[i] = Math.round(smoothstep(INK_LO, INK_HI, contrast) * 255);
  }
  dropArtifacts(alpha, w, h);

  // Warna tinta = rata-rata piksel yang jelas tinta, digelapkan bila kepucatan
  // (supaya hasil cetak tegas); tinta biru tetap biru, hitam tetap hitam.
  let r = 0, g = 0, b = 0, n = 0;
  for (let i = 0; i < total; i++) {
    if (alpha[i] >= 230) { r += data[i * 4]; g += data[i * 4 + 1]; b += data[i * 4 + 2]; n++; }
  }
  if (n < Math.max(MIN_COMPONENT_PX, total * 0.0003)) return null; // nyaris tak ada tinta
  r /= n; g /= n; b /= n;
  const inkLum = 0.299 * r + 0.587 * g + 0.114 * b;
  if (inkLum > 80) { const f = 80 / inkLum; r *= f; g *= f; b *= f; }
  const ink = [Math.round(r), Math.round(g), Math.round(b)];

  const box = inkBounds(img, (i) => alpha[i]);
  if (!box || Math.max(box.w, box.h) < MIN_SIDE_PX) return null;
  return cropWithPadding(img, box, (i, out, o) => {
    out[o] = ink[0]; out[o + 1] = ink[1]; out[o + 2] = ink[2]; out[o + 3] = alpha[i];
  });
}

// ===========================================================================
// Bagian browser (kanvas / DOM) — jangan dipanggil dari Node.
// ===========================================================================

function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result));
    fr.onerror = () => reject(new Error("Gagal membaca gambar tanda tangan."));
    fr.readAsDataURL(blob);
  });
}

/** Kecilkan ke ukuran akhir & ubah ke PNG (≤ SIGNATURE_MAX_BYTES bila memungkinkan). */
async function encodeSignature(img: RgbaImage): Promise<SignatureResult> {
  const src = document.createElement("canvas");
  src.width = img.width;
  src.height = img.height;
  src.getContext("2d")!.putImageData(new ImageData(new Uint8ClampedArray(img.data), img.width, img.height), 0, 0);

  let scale = Math.min(1, SIGNATURE_OUTPUT_MAX_W / img.width, SIGNATURE_OUTPUT_MAX_H / img.height);
  for (let attempt = 0; ; attempt++) {
    const width = Math.max(1, Math.round(img.width * scale));
    const height = Math.max(1, Math.round(img.height * scale));
    const dst = document.createElement("canvas");
    dst.width = width;
    dst.height = height;
    const ctx = dst.getContext("2d")!;
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(src, 0, 0, width, height);
    const blob = await new Promise<Blob | null>((res) => dst.toBlob(res, "image/png"));
    if (!blob) throw new Error("Gagal membuat gambar tanda tangan.");
    if (blob.size <= SIGNATURE_MAX_BYTES || attempt >= 3) {
      return { blob, dataUrl: await blobToDataUrl(blob), width, height };
    }
    scale *= 0.75;
  }
}

async function openBitmap(file: Blob): Promise<{ source: CanvasImageSource; width: number; height: number; close: () => void }> {
  if (typeof createImageBitmap === "function") {
    try {
      const bmp = await createImageBitmap(file); // orientasi EXIF foto ponsel ikut dihormati
      return { source: bmp, width: bmp.width, height: bmp.height, close: () => bmp.close() };
    } catch { /* jatuh ke <img> */ }
  }
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = () => reject(new Error("Foto tidak bisa dibuka. Gunakan berkas JPG atau PNG."));
      el.src = url;
    });
    return { source: img, width: img.naturalWidth, height: img.naturalHeight, close: () => URL.revokeObjectURL(url) };
  } catch (e) {
    URL.revokeObjectURL(url);
    throw e;
  }
}

/** Foto/berkas gambar tanda tangan → PNG transparan siap simpan. Melempar Error berpesan Indonesia. */
export async function processSignaturePhoto(file: Blob): Promise<SignatureResult> {
  const bmp = await openBitmap(file);
  try {
    const scale = Math.min(1, SIGNATURE_ANALYSIS_MAX_SIDE / Math.max(bmp.width, bmp.height));
    const w = Math.max(1, Math.round(bmp.width * scale));
    const h = Math.max(1, Math.round(bmp.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d", { willReadFrequently: true })!;
    ctx.drawImage(bmp.source, 0, 0, w, h);
    const out = extractSignature(ctx.getImageData(0, 0, w, h));
    if (!out) {
      throw new Error(
        "Tanda tangan tidak terdeteksi. Foto di atas kertas putih polos dengan tinta yang jelas, " +
        "pastikan cahaya merata dan tanda tangan memenuhi sebagian besar foto.",
      );
    }
    return await encodeSignature(out);
  } finally {
    bmp.close();
  }
}

/** Isi kanvas gambar langsung → PNG transparan terpotong rapat. */
export async function exportDrawnSignature(canvas: HTMLCanvasElement): Promise<SignatureResult> {
  const ctx = canvas.getContext("2d", { willReadFrequently: true })!;
  const out = cropToInk(ctx.getImageData(0, 0, canvas.width, canvas.height));
  if (!out) throw new Error("Belum ada goresan tanda tangan. Gambar dulu di kotak.");
  return await encodeSignature(out);
}
