// Tombol nyala/mati kop surat, logo, tanda tangan, meterai & jumlah rangkap untuk dokumen Word.
// Bisa dipakai kapan saja selagi Draft; hanya blok kop/tanda tangan yang berubah, ketikan bebas tetap utuh.
import React, { useCallback, useEffect, useState } from "react";
import { Building2, Image as ImageIcon, PenTool, Stamp, Copy, Loader2, RefreshCw, Columns2, Rows2, Languages } from "lucide-react";
import type { Contract } from "./types";

interface Layout { available: boolean; lang?: "id" | "bilingual" | "en"; layout?: "side" | "stacked"; hasNew?: boolean; translated?: boolean; geminiReady?: boolean; kop?: boolean; logo?: boolean; ttd?: boolean; digital?: boolean; digitalAvailable?: boolean; meterai?: boolean; copies?: 1 | 2; editable?: boolean }

export default function WordLayoutBar({ cat, contract, onChanged, showToast, askConfirm }: {
  cat: "lang" | "view" | "sign";
  contract: Contract;
  onChanged: (c: Contract) => void;
  showToast: (msg: string, type?: "success" | "warning" | "info") => void;
  askConfirm?: (msg: string, opts?: { title?: string; confirmLabel?: string; danger?: boolean }) => Promise<boolean>;
}) {
  const [lay, setLay] = useState<Layout>({ available: false });
  const [busy, setBusy] = useState(false);

  const load = useCallback(async () => {
    try {
      const r = await fetch(`/api/contracts/${contract.id}/word-layout`);
      setLay(r.ok ? await r.json() : { available: false });
    } catch { setLay({ available: false }); }
  }, [contract.id]);
  useEffect(() => { load(); }, [load, contract.currentDocumentVersion]);

  if (!lay.available || !lay.editable) return null;

  const apply = async (body: Record<string, unknown>) => {
    if (busy) return;
    setBusy(true);
    try {
      const r = await fetch(`/api/contracts/${contract.id}/word-layout`, {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
      });
      const data = await r.json().catch(() => ({}));
      if (!r.ok || !data.success) { showToast(data.error || "Gagal mengubah tata letak", "warning"); return; }
      onChanged(data.contract as Contract);
      await load();
    } catch (e: any) {
      showToast(e?.message || "Gagal mengubah tata letak", "warning");
    } finally { setBusy(false); }
  };

  // Bahasa dokumen: terjemahan AI disisipkan di bawah teks Indonesia; bahasa lain hanya disembunyikan (tidak dihapus).
  const setLang = async (mode: "id" | "bilingual" | "en", retranslate = false, layout?: "side" | "stacked") => {
    if (busy) return;
    setBusy(true);
    if (mode !== "id" && (!lay.translated || retranslate)) showToast("Menerjemahkan dokumen dengan AI… bisa memakan waktu beberapa menit.", "info");
    try {
      const r = await fetch(`/api/contracts/${contract.id}/word-language`, {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ mode, retranslate, layout: layout || lay.layout || "side" }),
      });
      const data = await r.json().catch(() => ({}));
      if (!r.ok || !data.success) { showToast(data.error || "Gagal mengubah bahasa dokumen", "warning"); return; }
      onChanged(data.contract as Contract);
      await load();
    } catch (e: any) {
      showToast(e?.message || "Gagal mengubah bahasa dokumen", "warning");
    } finally { setBusy(false); }
  };
  const curLang = lay.lang || "id";
  const curLayout = lay.layout || "side";

  const cls = (on: boolean) => `px-3 py-2 text-xs font-semibold rounded-xl border transition flex items-center gap-1 ${
    on ? "bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700" : "bg-slate-900/50 hover:bg-slate-800 text-slate-500 border-slate-800"}`;
  const hint = "Tunggu perubahan terakhir di editor tersimpan dulu. Editor akan dimuat ulang.";

  return (
    <>
      {cat === "lang" && (<>
      <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-xl p-0.5" title="Bahasa dokumen. Terjemahan dibuat AI saat pertama kali dipilih; teks Anda tidak dihapus, hanya disembunyikan.">
        {busy && <Loader2 className="w-3.5 h-3.5 animate-spin text-slate-400 ml-1.5" />}
        {([
          { k: "id" as const, label: "ID", tip: "Bahasa Indonesia saja" },
          { k: "en" as const, label: "EN", tip: "Bahasa Inggris saja (terjemahan AI)" },
          { k: "bilingual" as const, label: "ID+EN", tip: "Dwibahasa — Indonesia & Inggris" },
        ]).map((o) => (
          <button key={o.k} type="button" disabled={busy} title={o.tip} onClick={() => curLang !== o.k && setLang(o.k)}
            className={`px-2.5 py-1.5 rounded-lg text-[11px] font-bold transition cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed ${curLang === o.k ? "bg-indigo-600 text-white" : "text-slate-400 hover:text-slate-200"}`}>
            {o.label}
          </button>
        ))}
      </div>
      <div className="flex items-center gap-1 bg-slate-900 border border-slate-800 rounded-xl p-0.5" title="Susunan dua bahasa. Memilih salah satunya otomatis memakai mode ID+EN.">
        {([
          { k: "side" as const, label: "Kiri-Kanan", tip: "Berdampingan — Indonesia di kiri, Inggris di kanan", Icon: Columns2 },
          { k: "stacked" as const, label: "Atas-Bawah", tip: "Berselang per pasal — judul+isi Indonesia dulu, lalu judul+isi Inggris di bawahnya", Icon: Rows2 },
        ]).map(({ k, label, tip, Icon }) => {
          const active = curLang === "bilingual" && curLayout === k;
          return (
            <button key={k} type="button" disabled={busy} title={tip} onClick={() => !active && setLang("bilingual", false, k)}
              className={`px-2.5 py-1.5 rounded-lg text-[11px] font-bold transition cursor-pointer flex items-center gap-1 disabled:opacity-50 disabled:cursor-not-allowed ${active ? "bg-indigo-600 text-white" : "text-slate-400 hover:text-slate-200"}`}>
              <Icon className="w-3.5 h-3.5" /> {label}
            </button>
          );
        })}
      </div>
      {lay.hasNew && curLang === "bilingual" && (
        <button type="button" disabled={busy} title="Ada teks Indonesia yang baru Anda ketik dan belum punya terjemahan. Hanya bagian baru yang diterjemahkan; terjemahan lain (termasuk yang Anda edit) tidak disentuh." className={cls(true)}
          onClick={() => setLang("bilingual")}>
          <Languages className="w-3.5 h-3.5" /> Terjemahkan Teks Baru
        </button>
      )}
      {lay.translated && (
        <button type="button" disabled={busy} title="Buang semua terjemahan Inggris yang ada lalu terjemahkan ulang seluruh dokumen dari teks Indonesia saat ini. Pakai kalau teks Indonesia banyak berubah. Teks Indonesia tidak berubah; perubahan manual di bagian Inggris hilang." className={cls(false)}
          onClick={async () => {
            const ok = askConfirm ? await askConfirm("Seluruh terjemahan Inggris akan dibuat ulang dari teks Indonesia yang sekarang. Perubahan manual pada bagian Inggris akan hilang. Teks Indonesia tidak berubah.", { title: "Terjemahkan Ulang Seluruhnya", confirmLabel: "Ya, Terjemahkan Ulang", danger: false }) : true;
            if (ok) setLang(curLang === "id" ? "bilingual" : curLang, true);
          }}>
          <RefreshCw className="w-3.5 h-3.5" /> Terjemahkan Ulang
        </button>
      )}
      </>)}
      {cat === "view" && (<>
      <button type="button" disabled={busy} title={hint} onClick={() => apply({ kop: !lay.kop })} className={cls(!!lay.kop)}>
        {busy ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Building2 className="w-3.5 h-3.5" />} Kop Surat: {lay.kop ? "Nyala" : "Mati"}
      </button>
      {lay.kop && (
        <button type="button" disabled={busy} title={hint} onClick={() => apply({ kop: true, logo: !lay.logo })} className={cls(!!lay.logo)}>
          <ImageIcon className="w-3.5 h-3.5" /> Logo: {lay.logo ? "Nyala" : "Mati"}
        </button>
      )}
      </>)}
      {cat === "sign" && (<>
      <button type="button" disabled={busy} title={hint} onClick={() => apply({ ttd: !lay.ttd })} className={cls(!!lay.ttd)}>
        <PenTool className="w-3.5 h-3.5" /> Tanda Tangan: {lay.ttd ? "Nyala" : "Mati"}
      </button>
      {lay.ttd && (
        <>
          <button type="button" disabled={busy} title={hint} onClick={() => apply({ meterai: !lay.meterai })} className={cls(!!lay.meterai)}>
            <Stamp className="w-3.5 h-3.5" /> Meterai: {lay.meterai ? "Ya" : "Tidak"}
          </button>
          <button type="button" disabled={busy} title={hint} onClick={() => apply({ copies: lay.copies === 2 ? 1 : 2 })} className={cls(true)}>
            <Copy className="w-3.5 h-3.5" /> {lay.copies === 2 ? "2 Rangkap" : "1 Rangkap"}
          </button>
          {lay.copies === 2 && (
            <button type="button" disabled={busy} title="Hitung ulang tulisan yang sehalaman dengan tanda tangan lalu salin ke rangkap 2. Pakai setelah Anda mengubah isi dokumen." onClick={() => apply({ refreshTail: true })} className={cls(false)}>
              <RefreshCw className="w-3.5 h-3.5" /> Segarkan Rangkap 2
            </button>
          )}
          {lay.digitalAvailable && (
            <button type="button" disabled={busy} title={hint} onClick={() => apply({ digital: !lay.digital })} className={cls(!!lay.digital)}>
              <PenTool className="w-3.5 h-3.5" /> TTD Digital: {lay.digital ? "Nyala" : "Mati"}
            </button>
          )}
        </>
      )}
      </>)}
    </>
  );
}
