// Uji pengolah gambar TTD Digital (src/signature-image.ts) dengan gambar buatan.
// Jalankan: npx tsx scripts/signature-image-test.ts
import assert from "node:assert/strict";
import { cropToInk, extractSignature, fitWithin, hasTransparency, type RgbaImage } from "../src/signature-image";

// Pembangkit angka acak tetap supaya hasil uji selalu sama.
let seed = 1234567;
const rnd = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);

function blank(w: number, h: number, fill: (x: number, y: number) => [number, number, number, number]): RgbaImage {
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const [r, g, b, a] = fill(x, y);
      const o = (y * w + x) * 4;
      data[o] = r; data[o + 1] = g; data[o + 2] = b; data[o + 3] = a;
    }
  }
  return { data, width: w, height: h };
}

/** Titik-titik goresan "tanda tangan" (kurva bergelombang) pada gambar w×h. */
function strokePoints(w: number, h: number): { x: number; y: number }[] {
  const pts: { x: number; y: number }[] = [];
  for (let i = 0; i <= 400; i++) {
    const t = i / 400;
    pts.push({ x: w * (0.15 + 0.7 * t), y: h * (0.5 + 0.18 * Math.sin(t * 14) * (1 - 0.4 * t)) });
  }
  return pts;
}

/** Cat goresan berujung bulat (anti-alias) di atas gambar: warna tinta dicampur ke piksel yang ada. */
function paint(img: RgbaImage, pts: { x: number; y: number }[], radius: number, ink: [number, number, number]) {
  for (const p of pts) {
    for (let y = Math.floor(p.y - radius - 1); y <= Math.ceil(p.y + radius + 1); y++) {
      for (let x = Math.floor(p.x - radius - 1); x <= Math.ceil(p.x + radius + 1); x++) {
        if (x < 0 || y < 0 || x >= img.width || y >= img.height) continue;
        const d = Math.hypot(x - p.x, y - p.y);
        const cov = Math.max(0, Math.min(1, radius + 0.5 - d));
        if (cov <= 0) continue;
        const o = (y * img.width + x) * 4;
        const cur = Math.min(1, img.data[o + 3] === 255 ? cov : cov);
        for (let c = 0; c < 3; c++) img.data[o + c] = img.data[o + c] * (1 - cur) + ink[c] * cur;
      }
    }
  }
}

let passed = 0;
const ok = (name: string) => { passed++; console.log("  ok  -", name); };

// ---------------------------------------------------------------------------
// 1. Foto kertas bergradasi (bayangan) + noda debu + tepi meja gelap di kiri
// ---------------------------------------------------------------------------
{
  const W = 1200, H = 800;
  const photo = blank(W, H, (x) => {
    const base = 238 - 45 * (x / W) + (rnd() - 0.5) * 8; // terang di kiri → redup di kanan, plus noise
    return [base, base, base - 4, 255];
  });
  const pts = strokePoints(W, H);
  paint(photo, pts, 4.5, [25, 45, 150]); // tinta biru
  // noda debu kecil di pojok
  for (const [dx, dy] of [[1100, 60], [1150, 740], [700, 30]]) paint(photo, [{ x: dx, y: dy }], 2, [90, 90, 90]);
  // tepi meja gelap padat di sisi kiri (mis. foto tidak dipotong rapi)
  for (let y = 0; y < H; y++) for (let x = 0; x < 70; x++) {
    const o = (y * W + x) * 4;
    photo.data[o] = 70; photo.data[o + 1] = 48; photo.data[o + 2] = 34;
  }

  const t0 = Date.now();
  const out = extractSignature(photo);
  const ms = Date.now() - t0;
  assert.ok(out, "tanda tangan harus terdeteksi");
  console.log(`  (waktu proses ${W}x${H}: ${ms} ms, hasil ${out!.width}x${out!.height})`);

  // Lebar hasil ≈ lebar goresan (0.7*W) — membuktikan tepi meja & debu tidak ikut.
  const strokeW = 0.7 * W;
  assert.ok(out!.width >= strokeW * 0.97 && out!.width <= strokeW * 1.12, `lebar ${out!.width} harus ≈ ${strokeW}`);
  ok("dipotong rapat sebatas goresan (tepi meja & debu tidak ikut)");

  const { data, width, height } = out!;
  const corners = [0, (width - 1), (height - 1) * width, height * width - 1].map((i) => data[i * 4 + 3]);
  assert.deepEqual(corners, [0, 0, 0, 0]);
  ok("latar kertas jadi transparan (4 pojok alpha 0)");

  let solid = 0, blueish = 0, partial = 0;
  for (let i = 0; i < width * height; i++) {
    const a = data[i * 4 + 3];
    if (a >= 230) { solid++; if (data[i * 4 + 2] > data[i * 4] + 20) blueish++; }
    else if (a > 20) partial++;
  }
  assert.ok(solid > 800, "harus ada piksel tinta penuh");
  assert.ok(blueish / solid > 0.95, "warna tinta biru terjaga");
  assert.ok(partial > 0 && partial < solid * 1.5, "tepi goresan halus (anti-alias), bukan bergerigi");
  ok(`tinta penuh ${solid} px, berwarna biru, tepi halus ${partial} px`);

  // Hasil tidak boleh punya bayangan/gradasi kertas: bagian dalam kotak tapi di luar goresan harus bening.
  let paperLeak = 0;
  for (let i = 0; i < width * height; i++) if (data[i * 4 + 3] > 20 && data[i * 4 + 3] < 230) paperLeak++;
  assert.ok(paperLeak < solid * 1.5);
  ok("tidak ada sisa bayangan kertas");
}

// Titik dekat badan tanda tangan (mis. titik huruf "i") dipertahankan, debu jauh dibuang.
{
  const W = 1000, H = 600;
  const mk = (extra: { x: number; y: number }[]) => {
    const p = blank(W, H, () => [236, 236, 236, 255]);
    paint(p, strokePoints(W, H), 4, [20, 20, 20]);
    for (const e of extra) paint(p, [e], 4, [20, 20, 20]);
    return p;
  };
  const base = extractSignature(mk([]))!;
  const withNearDot = extractSignature(mk([{ x: W * 0.5, y: H * 0.5 - 0.18 * H - 40 }]))!; // 40px di atas kurva
  const withFarDust = extractSignature(mk([{ x: W * 0.5, y: 12 }]))!;
  assert.ok(withNearDot.height > base.height + 20, "titik dekat harus ikut");
  assert.equal(withFarDust.height, base.height);
  ok("titik dekat badan dipertahankan, debu jauh dibuang");
}

// ---------------------------------------------------------------------------
// 2. Kertas kosong / terlalu pucat → tidak terdeteksi
// ---------------------------------------------------------------------------
{
  const empty = blank(800, 600, () => { const v = 235 + (rnd() - 0.5) * 6; return [v, v, v, 255]; });
  assert.equal(extractSignature(empty), null);
  ok("kertas kosong → null (tidak berpura-pura ada tanda tangan)");

  const faint = blank(800, 600, () => [238, 238, 238, 255]);
  paint(faint, strokePoints(800, 600), 3, [222, 222, 222]); // nyaris sama terang dgn kertas
  assert.equal(extractSignature(faint), null);
  ok("goresan terlalu pucat → null");
}

// ---------------------------------------------------------------------------
// 3. PNG yang sudah bening dipakai apa adanya (hanya dipotong rapat)
// ---------------------------------------------------------------------------
{
  // Cakupan goresan digambar di atas putih, lalu dijadikan alpha (hitam, latar bening).
  const mask = blank(600, 300, () => [255, 255, 255, 255]);
  paint(mask, strokePoints(600, 300), 3, [0, 0, 0]);
  const png = blank(600, 300, (x, y) => [0, 0, 0, 255 - mask.data[(y * 600 + x) * 4]]);
  assert.equal(hasTransparency(png), true);
  const out = extractSignature(png);
  assert.ok(out);
  assert.ok(out!.width < 600 && out!.height < 300);
  ok("PNG bening → hanya dipotong rapat");
}

// ---------------------------------------------------------------------------
// 4. Kanvas gambar langsung
// ---------------------------------------------------------------------------
{
  const canvas = blank(1200, 480, () => [0, 0, 0, 0]);
  const pts = strokePoints(1200, 480).filter((p) => p.x > 300 && p.x < 800);
  for (const p of pts) {
    for (let y = Math.floor(p.y - 4); y <= Math.ceil(p.y + 4); y++) for (let x = Math.floor(p.x - 4); x <= Math.ceil(p.x + 4); x++) {
      if (Math.hypot(x - p.x, y - p.y) <= 4) { const o = (y * 1200 + x) * 4; canvas.data[o] = 16; canvas.data[o + 1] = 23; canvas.data[o + 2] = 42; canvas.data[o + 3] = 255; }
    }
  }
  const out = cropToInk(canvas);
  assert.ok(out);
  assert.ok(out!.width >= 490 && out!.width <= 560, `lebar ${out!.width}`);
  ok(`kanvas → dipotong rapat (${out!.width}x${out!.height})`);

  assert.equal(cropToInk(blank(400, 200, () => [0, 0, 0, 0])), null);
  ok("kanvas kosong → null");
}

// ---------------------------------------------------------------------------
// 5. fitWithin
// ---------------------------------------------------------------------------
{
  assert.deepEqual(fitWithin(400, 200, 100, 100), { width: 100, height: 50 });
  assert.deepEqual(fitWithin(200, 400, 100, 100), { width: 50, height: 100 });
  assert.deepEqual(fitWithin(0, 10, 100, 100), { width: 0, height: 0 });
  ok("fitWithin menjaga rasio");
}

console.log(`ALL SIGNATURE-IMAGE CHECKS PASSED (${passed})`);
