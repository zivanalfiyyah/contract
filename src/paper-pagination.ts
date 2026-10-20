// Tampilan "kertas ala Word" untuk preview kontrak template.
//
// Dokumen tetap dirender sebagai SATU alur (flow) HTML biasa — supaya Mode
// Edit (contentEditable), komentar, dll. tidak perlu dipecah per halaman.
// Fungsi di sini mengukur blok-blok di dalam alur itu lalu MENDORONG blok
// yang tidak muat ke awal halaman berikutnya (menambah margin-top-nya),
// sehingga isi tepat jatuh di area isi tiap lembar A4 yang digambar di
// belakangnya. Export PDF memotong PERSIS di lembar-lembar yang sama, jadi
// yang tampil di layar = yang tercetak.
//
// Aturan pemotongan (sama untuk layar & PDF):
//  - Potongan hanya di antara blok (paragraf, butir daftar, baris judul, tabel,
//    dua-kolom bahasa, blok TTD) — tidak pernah memotong satu baris teks.
//  - Judul pasal (data-pg-keep) tidak pernah ditinggal sendirian di dasar
//    halaman; ikut pindah bersama blok sesudahnya.
//  - Blok bertanda data-pdf-force-break selalu mulai di halaman baru
//    (dipakai rangkap ke-2 dst pada format "2 Rangkap + Meterai").
//  - Satu blok yang sendirian lebih tinggi dari satu halaman penuh tidak
//    dipotong: lembarnya dibuat lebih panjang (saat export diperkecil agar
//    muat di A4) — supaya tidak ada teks yang hilang di sela halaman.

export interface PaperPage {
  top: number;    // px (CSS, tanpa skala) dari atas tumpukan lembar
  height: number; // px; normalnya = tinggi A4, lebih bila ada blok sangat tinggi
}

export interface PaperGeometry {
  sheetHeight: number; // px
  topMargin: number;   // px
  bottomMargin: number;// px
  gap: number;         // px jarak antar lembar
}

export const MM_TO_PX = 96 / 25.4;
export const A4_WIDTH_PX = 210 * MM_TO_PX;
export const A4_HEIGHT_PX = 297 * MM_TO_PX;

const PUSH_ATTR = "data-pg-push";
const FLOW_DISPLAYS = new Set(["block", "list-item", "flow-root"]);
const BLOCKISH = new Set(["block", "list-item", "flow-root", "table", "flex", "grid"]);

interface Atom { el: HTMLElement; forceBreak: boolean; keep: boolean; }

function isSkipped(cs: CSSStyleDeclaration): boolean {
  return cs.display === "none" || cs.position === "absolute" || cs.position === "fixed";
}

// Elemen inline (mis. <span> pembungkus HTML pasal) yang di dalamnya ada blok
// (<p>, <ol>, ...) diperlakukan "transparan": blok di dalamnya tetap bisa
// dipisah per paragraf.
function containsBlock(el: HTMLElement): boolean {
  for (const child of Array.from(el.children) as HTMLElement[]) {
    const ccs = getComputedStyle(child);
    if (isSkipped(ccs)) continue;
    if (BLOCKISH.has(ccs.display)) return true;
    if ((ccs.display === "inline" || ccs.display === "contents") && containsBlock(child)) return true;
  }
  return false;
}

function isAtomic(el: HTMLElement, cs: CSSStyleDeclaration): boolean {
  if (el.hasAttribute("data-pg-atom")) return true;
  if (cs.display === "inline" || cs.display === "contents") return !containsBlock(el);
  if (!FLOW_DISPLAYS.has(cs.display)) return true; // flex/grid/table/inline-block = satu kesatuan
  return !containsBlock(el); // hanya berisi teks/inline
}

// Kumpulkan blok terkecil (urutan dokumen). `force` = penanda paksa-halaman-
// baru yang diwariskan ke blok PERTAMA di dalam elemen bertanda.
function collect(el: HTMLElement, out: Atom[], force: boolean): boolean {
  for (const child of Array.from(el.children) as HTMLElement[]) {
    const cs = getComputedStyle(child);
    if (isSkipped(cs)) continue;
    const f = force || child.hasAttribute("data-pdf-force-break");
    if (isAtomic(child, cs)) {
      out.push({ el: child, forceBreak: f, keep: child.hasAttribute("data-pg-keep") });
      force = false;
    } else if (collect(child, out, f)) {
      force = false;
    }
  }
  return !force;
}

/** Hapus semua dorongan halaman (kembalikan margin-top asli). */
export function clearPagination(root: HTMLElement): void {
  root.querySelectorAll<HTMLElement>(`[${PUSH_ATTR}]`).forEach((el) => {
    el.style.marginTop = el.getAttribute(PUSH_ATTR) || "";
    el.removeAttribute(PUSH_ATTR);
    if (!el.getAttribute("style")) el.removeAttribute("style");
  });
}

/** Bersihkan penanda paginasi dari HTML hasil contentEditable sebelum disimpan. */
export function stripPaginationMarks(html: string): string {
  if (!html || !html.includes(PUSH_ATTR) || typeof document === "undefined") return html;
  const tpl = document.createElement("template");
  tpl.innerHTML = html;
  tpl.content.querySelectorAll<HTMLElement>(`[${PUSH_ATTR}]`).forEach((el) => {
    el.style.marginTop = el.getAttribute(PUSH_ATTR) || "";
    el.removeAttribute(PUSH_ATTR);
    if (!el.getAttribute("style")) el.removeAttribute("style");
  });
  return tpl.innerHTML;
}

/**
 * Pecah alur `root` menjadi lembar-lembar. `root` harus position:relative dan
 * padding-top = margin atas (isi halaman 1 mulai tepat di margin atas).
 * Mengukur dgn getBoundingClientRect dibagi skala, jadi aman walau tumpukan
 * lembar sedang di-scale (transform) agar muat di panel.
 */
export function paginateFlow(root: HTMLElement, g: PaperGeometry): PaperPage[] {
  clearPagination(root);
  const rootRect0 = root.getBoundingClientRect();
  const k = root.offsetWidth > 0 ? rootRect0.width / root.offsetWidth : 1;
  const relTop = (el: HTMLElement) => (el.getBoundingClientRect().top - root.getBoundingClientRect().top) / k;
  const heightOf = (el: HTMLElement) => el.getBoundingClientRect().height / k;

  const atoms: Atom[] = [];
  collect(root, atoms, false);

  const contentH = Math.max(50, g.sheetHeight - g.topMargin - g.bottomMargin);
  const pages: PaperPage[] = [];
  let sheetTop = 0;
  let start = g.topMargin;
  let end = start + contentH;
  let lastBottom = start;

  const pushTo = (el: HTMLElement, target: number) => {
    el.setAttribute(PUSH_ATTR, el.style.marginTop || "");
    let m = (parseFloat(getComputedStyle(el).marginTop) || 0) + (target - relTop(el));
    // Margin bisa "collapse" dgn margin saudara/induk — ukur ulang & koreksi.
    for (let t = 0; t < 4; t++) {
      el.style.marginTop = `${m}px`;
      const d = target - relTop(el);
      if (Math.abs(d) <= 0.5) break;
      m += d;
    }
  };

  let i = 0;
  let justPushed = -1; // pengaman: blok yg baru dipindah dianggap awal halaman
  while (i < atoms.length) {
    const a = atoms[i];
    const h = heightOf(a.el);
    const top = relTop(a.el);
    const bottom = top + h;
    if (h <= 0.5) { i++; continue; }
    const atPageStart = top <= start + 1.5 || justPushed === i;
    justPushed = -1;
    const overflow = bottom > end + 1.5;
    if (atPageStart || (!overflow && !a.forceBreak)) {
      if (bottom > end) end = bottom; // blok sangat tinggi: lembar dipanjangkan
      lastBottom = Math.max(lastBottom, bottom);
      i++;
      continue;
    }
    // Pindahkan ke halaman berikutnya — bawa judul pasal yg menggantung.
    let b = i;
    if (!a.forceBreak && i > 0 && atoms[i - 1].keep) {
      const pt = relTop(atoms[i - 1].el);
      if (pt > start + 1.5) b = i - 1;
    }
    const pageHeight = Math.max(g.sheetHeight, end - sheetTop + g.bottomMargin);
    pages.push({ top: sheetTop, height: pageHeight });
    sheetTop += pageHeight + g.gap;
    start = sheetTop + g.topMargin;
    end = start + contentH;
    pushTo(atoms[b].el, start);
    lastBottom = start;
    i = b; // ukur ulang mulai dari blok yg baru dipindah
    justPushed = b;
  }
  const lastHeight = Math.max(g.sheetHeight, Math.max(end, lastBottom) - sheetTop + g.bottomMargin);
  pages.push({ top: sheetTop, height: lastHeight });
  return pages;
}

/** Halaman (index) tempat sebuah elemen berada — dipakai format 2 rangkap. */
export function pageIndexOf(root: HTMLElement, el: HTMLElement, pages: PaperPage[]): number {
  const k = root.offsetWidth > 0 ? root.getBoundingClientRect().width / root.offsetWidth : 1;
  const top = (el.getBoundingClientRect().top - root.getBoundingClientRect().top) / k;
  for (let p = pages.length - 1; p >= 0; p--) if (top >= pages[p].top - 1) return p;
  return 0;
}

export function samePages(a: PaperPage[], b: PaperPage[]): boolean {
  if (a.length !== b.length) return false;
  return a.every((p, i) => Math.abs(p.top - b[i].top) < 0.5 && Math.abs(p.height - b[i].height) < 0.5);
}
