// Tombol nyala/mati kop surat, logo, tanda tangan, meterai & jumlah rangkap untuk dokumen Word.
// Bisa dipakai kapan saja selagi Draft; hanya blok kop/tanda tangan yang berubah, ketikan bebas tetap utuh.
import React, { useCallback, useEffect, useState } from "react";
import { Building2, Image as ImageIcon, PenTool, Stamp, Copy, Loader2 } from "lucide-react";
import type { Contract } from "./types";

interface Layout { available: boolean; kop?: boolean; logo?: boolean; ttd?: boolean; meterai?: boolean; copies?: 1 | 2; editable?: boolean }

export default function WordLayoutBar({ contract, onChanged, showToast }: {
  contract: Contract;
  onChanged: (c: Contract) => void;
  showToast: (msg: string, type?: "success" | "warning" | "info") => void;
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

  const cls = (on: boolean) => `px-3 py-2 text-xs font-semibold rounded-xl border transition flex items-center gap-1 ${
    on ? "bg-slate-800 hover:bg-slate-700 text-slate-200 border-slate-700" : "bg-slate-900/50 hover:bg-slate-800 text-slate-500 border-slate-800"}`;
  const hint = "Tunggu perubahan terakhir di editor tersimpan dulu. Editor akan dimuat ulang.";

  return (
    <>
      <button type="button" disabled={busy} title={hint} onClick={() => apply({ kop: !lay.kop })} className={cls(!!lay.kop)}>
        {busy ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Building2 className="w-3.5 h-3.5" />} Kop Surat: {lay.kop ? "Nyala" : "Mati"}
      </button>
      {lay.kop && (
        <button type="button" disabled={busy} title={hint} onClick={() => apply({ kop: true, logo: !lay.logo })} className={cls(!!lay.logo)}>
          <ImageIcon className="w-3.5 h-3.5" /> Logo: {lay.logo ? "Nyala" : "Mati"}
        </button>
      )}
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
        </>
      )}
    </>
  );
}
