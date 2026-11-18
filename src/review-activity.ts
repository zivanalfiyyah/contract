import { useCallback, useEffect, useRef, useState } from "react";

// Aktivitas SESI REVIEW reviewer (komentar, hapus, tambah, edit). Dipakai tampilan untuk menonaktifkan tombol
// Setujui selama ada aktivitas. Server tetap memvalidasi ulang (409 REVIEW_ACTIVITY_DETECTED), jadi ini hanya
// kemudahan: gagal mengambil data TIDAK mengunci tombol (tampilan mengikuti status terakhir yang diketahui).

export interface ReviewActivityCounts {
  comments: number; commentEdits: number; commentDeletions: number; editorChanges: number; savedChanges: number;
}
export interface ReviewActivity { hasActivity: boolean; counts: ReviewActivityCounts | null }

/** Ringkasan singkat yang ditampilkan di dekat tombol, mis. "2 komentar/usulan, perubahan di editor Word". */
export function describeReviewActivity(c: ReviewActivityCounts | null): string {
  if (!c) return "";
  const parts: string[] = [];
  if (c.comments) parts.push(`${c.comments} komentar/usulan`);
  if (c.commentEdits) parts.push(`${c.commentEdits} komentar diubah`);
  if (c.commentDeletions) parts.push(`${c.commentDeletions} komentar dihapus`);
  if (c.editorChanges || c.savedChanges) parts.push("perubahan di editor Word");
  return parts.join(", ");
}

/**
 * Ambil aktivitas sesi dari `url` (null = tidak aktif). Disegarkan berkala dan setiap `refreshKey` berubah
 * (mis. setelah komentar terkirim). `refresh()` bisa dipanggil manual (mis. saat editor Word melaporkan perubahan).
 */
export function useReviewActivity(url: string | null, refreshKey?: unknown): ReviewActivity & { refresh: () => Promise<void> } {
  const [state, setState] = useState<ReviewActivity>({ hasActivity: false, counts: null });
  const alive = useRef(true);
  useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);
  const refresh = useCallback(async () => {
    if (!url) { setState((s) => (s.hasActivity || s.counts ? { hasActivity: false, counts: null } : s)); return; }
    try {
      const r = await fetch(url);
      if (!r.ok) return;
      const j = await r.json();
      if (alive.current) setState({ hasActivity: !!j.hasActivity, counts: j.counts || null });
    } catch { /* gagal ambil data: pertahankan status terakhir */ }
  }, [url]);
  useEffect(() => {
    refresh();
    if (!url) return;
    const t = setInterval(refresh, 10000);
    return () => clearInterval(t);
  }, [refresh, url, refreshKey]);
  return { ...state, refresh };
}
