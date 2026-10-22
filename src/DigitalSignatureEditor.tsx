// ---------------------------------------------------------------------------
// Jendela pengaturan TTD Digital Pihak Pertama — Konfigurasi > Kelola
// Perusahaan, HANYA Super Admin (server menegakkan: /api/settings/digital-signature).
//
// Dua cara membuat tanda tangan:
//  1. Unggah Foto  : foto tanda tangan di kertas → latar dihapus, dipotong rapat.
//  2. Gambar Langsung : kotak putih seperti kertas, digores dengan mouse / jari /
//                      stylus. Tombol Hapus untuk mengulang, Simpan untuk menyimpan.
//
// Keduanya menghasilkan PNG transparan terpotong rapat (src/signature-image.ts),
// sehingga di dokumen otomatis tampil rapi di kolom TTD Pihak Pertama, apa pun
// format TTD-nya. Tanda tangan ini baru dipakai pada dokumen yang tombol "TTD
// Digital"-nya dinyalakan.
// ---------------------------------------------------------------------------
import React, { useCallback, useEffect, useRef, useState } from "react";
import { Eraser, ImageUp, PenLine, Save, Trash2, X } from "lucide-react";
import {
  SIGNATURE_DRAW_INK,
  exportDrawnSignature,
  processSignaturePhoto,
  type SignatureResult,
} from "./signature-image";

interface Props {
  tenantId: string;
  tenantName: string;
  onClose: () => void;
  /** Dipanggil setelah tanda tangan tersimpan / dihapus (mis. untuk menyegarkan info di layar). */
  onChanged: () => void;
  notify: (message: string, type?: "success" | "info" | "warning") => void;
}

interface CurrentSignature {
  configured: boolean;
  dataUrl?: string;
  source?: "upload" | "draw";
  updatedAt?: string;
  updatedByName?: string;
}

// Ukuran kanvas gambar langsung. Isi piksel 2× supaya goresan tajam di layar
// retina; tampilannya mengikuti lebar jendela (aspek tetap).
const PAD_W = 1120;
const PAD_H = 440;
const PAD_BASE_WIDTH = 6;

type Tab = "foto" | "gambar";

const DigitalSignatureEditor: React.FC<Props> = ({ tenantId, tenantName, onClose, onChanged, notify }) => {
  const [tab, setTab] = useState<Tab>("foto");
  const [current, setCurrent] = useState<CurrentSignature | null>(null);
  const [saving, setSaving] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);

  // --- Unggah Foto ---
  const [photo, setPhoto] = useState<SignatureResult | null>(null);
  const [photoBusy, setPhotoBusy] = useState(false);
  const [photoError, setPhotoError] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- Gambar Langsung ---
  const padRef = useRef<HTMLCanvasElement>(null);
  const [hasInk, setHasInk] = useState(false);
  const drawing = useRef<{ id: number; last: { x: number; y: number }; mid: { x: number; y: number } } | null>(null);

  const qs = `?tenantId=${encodeURIComponent(tenantId)}`;

  const loadCurrent = useCallback(async () => {
    try {
      const res = await fetch(`/api/settings/digital-signature${qs}`);
      setCurrent(res.ok ? await res.json() : { configured: false });
    } catch {
      setCurrent({ configured: false });
    }
  }, [qs]);
  useEffect(() => { loadCurrent(); }, [loadCurrent]);

  // Tutup dengan Esc.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape" && !saving) onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose, saving]);

  // ----- kanvas -----
  const clearPad = () => {
    const c = padRef.current;
    if (!c) return;
    c.getContext("2d")!.clearRect(0, 0, c.width, c.height);
    drawing.current = null;
    setHasInk(false);
  };

  const pointOf = (e: { clientX: number; clientY: number }) => {
    const c = padRef.current!;
    const r = c.getBoundingClientRect();
    return { x: ((e.clientX - r.left) * c.width) / r.width, y: ((e.clientY - r.top) * c.height) / r.height };
  };
  const widthOf = (e: { pointerType?: string; pressure?: number }) =>
    e.pointerType === "pen" && e.pressure && e.pressure > 0 ? PAD_BASE_WIDTH * (0.45 + e.pressure * 1.1) : PAD_BASE_WIDTH;

  const onPadDown = (e: React.PointerEvent<HTMLCanvasElement>) => {
    if (e.button !== 0 && e.pointerType === "mouse") return;
    const c = padRef.current!;
    c.setPointerCapture(e.pointerId);
    const p = pointOf(e);
    const ctx = c.getContext("2d")!;
    ctx.fillStyle = SIGNATURE_DRAW_INK;
    ctx.beginPath();
    ctx.arc(p.x, p.y, widthOf(e) / 2, 0, Math.PI * 2);
    ctx.fill();
    drawing.current = { id: e.pointerId, last: p, mid: p };
    setHasInk(true);
  };
  const onPadMove = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const st = drawing.current;
    if (!st || st.id !== e.pointerId) return;
    const ctx = padRef.current!.getContext("2d")!;
    ctx.strokeStyle = SIGNATURE_DRAW_INK;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    const native = e.nativeEvent as PointerEvent;
    const events = typeof native.getCoalescedEvents === "function" && native.getCoalescedEvents().length
      ? native.getCoalescedEvents()
      : [native];
    for (const ev of events) {
      const p = pointOf(ev);
      const mid = { x: (st.last.x + p.x) / 2, y: (st.last.y + p.y) / 2 };
      ctx.lineWidth = widthOf(ev);
      ctx.beginPath();
      ctx.moveTo(st.mid.x, st.mid.y);
      ctx.quadraticCurveTo(st.last.x, st.last.y, mid.x, mid.y); // garis dihaluskan lewat titik tengah
      ctx.stroke();
      st.mid = mid;
      st.last = p;
    }
  };
  const onPadUp = (e: React.PointerEvent<HTMLCanvasElement>) => {
    const st = drawing.current;
    if (!st || st.id !== e.pointerId) return;
    const ctx = padRef.current!.getContext("2d")!;
    ctx.strokeStyle = SIGNATURE_DRAW_INK;
    ctx.lineCap = "round";
    ctx.lineWidth = widthOf(e);
    ctx.beginPath();
    ctx.moveTo(st.mid.x, st.mid.y);
    ctx.lineTo(st.last.x, st.last.y);
    ctx.stroke();
    drawing.current = null;
  };

  // ----- foto -----
  const onPickPhoto = async (file: File | undefined) => {
    if (!file) return;
    setPhotoError("");
    setPhoto(null);
    if (!/^image\/(png|jpe?g|webp)$/i.test(file.type)) {
      setPhotoError("Berkas harus berupa gambar JPG atau PNG.");
      return;
    }
    setPhotoBusy(true);
    try {
      setPhoto(await processSignaturePhoto(file));
    } catch (err: any) {
      setPhotoError(err?.message || "Foto tidak bisa diproses.");
    } finally {
      setPhotoBusy(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  // ----- simpan / hapus -----
  const upload = async (result: SignatureResult, source: "upload" | "draw") => {
    const fd = new FormData();
    fd.append("source", source);
    fd.append("file", result.blob, "ttd-digital.png");
    const res = await fetch(`/api/settings/digital-signature${qs}`, { method: "PUT", body: fd });
    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.success) throw new Error(data.error || `Gagal menyimpan tanda tangan (HTTP ${res.status}).`);
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      if (tab === "foto") {
        if (!photo) throw new Error("Pilih foto tanda tangan dulu.");
        await upload(photo, "upload");
      } else {
        if (!padRef.current) return;
        await upload(await exportDrawnSignature(padRef.current), "draw");
      }
      notify(`Tanda tangan digital ${tenantName} tersimpan.`, "success");
      onChanged();
      onClose();
    } catch (err: any) {
      notify(err?.message || "Gagal menyimpan tanda tangan.", "warning");
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async () => {
    setSaving(true);
    try {
      const res = await fetch(`/api/settings/digital-signature${qs}`, { method: "DELETE" });
      if (!res.ok) throw new Error("Gagal menghapus tanda tangan.");
      notify("Tanda tangan digital dihapus.", "success");
      setConfirmDelete(false);
      setPhoto(null);
      await loadCurrent();
      onChanged();
    } catch (err: any) {
      notify(err?.message || "Gagal menghapus tanda tangan.", "warning");
    } finally {
      setSaving(false);
    }
  };

  const canSave = !saving && (tab === "foto" ? !!photo : hasInk);
  const tabBtn = (k: Tab, label: string, icon: React.ReactNode) => (
    <button
      type="button"
      role="tab"
      aria-selected={tab === k}
      onClick={() => setTab(k)}
      className={`flex-1 px-3 py-2 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 transition cursor-pointer ${
        tab === k ? "bg-indigo-600 text-white" : "text-slate-400 hover:text-slate-200"
      }`}
    >
      {icon}
      {label}
    </button>
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-4 animate-fadeIn" onMouseDown={(e) => { if (e.target === e.currentTarget && !saving) onClose(); }}>
      <div role="dialog" aria-modal="true" aria-label="TTD Digital" className="bg-slate-950 border border-slate-800 rounded-2xl w-full max-w-xl shadow-2xl overflow-hidden max-h-[92vh] flex flex-col">
        <div className="p-5 border-b border-slate-800 bg-slate-900/50 flex justify-between items-start gap-3">
          <div className="min-w-0">
            <h3 className="font-bold text-slate-100 text-sm">TTD Digital Pihak Pertama</h3>
            <p className="text-[12px] text-slate-500 truncate">{tenantName}</p>
          </div>
          <button onClick={onClose} disabled={saving} aria-label="Tutup" className="text-slate-400 hover:text-slate-100 cursor-pointer shrink-0"><X className="w-4 h-4" /></button>
        </div>

        <div className="p-5 space-y-4 overflow-y-auto">
          {/* Tanda tangan yang sedang tersimpan */}
          <div className="flex items-center gap-3 p-3 rounded-xl bg-slate-900/60 border border-slate-800">
            <div className="w-28 h-14 rounded-lg bg-white border border-slate-300 flex items-center justify-center shrink-0 overflow-hidden">
              {current?.configured && current.dataUrl
                ? <img src={current.dataUrl} alt="Tanda tangan tersimpan" className="max-w-full max-h-full object-contain" />
                : <span className="text-[10px] text-slate-400 px-1 text-center leading-tight">{current === null ? "Memuat…" : "Belum ada"}</span>}
            </div>
            <div className="min-w-0 flex-1 text-[12px]">
              {current?.configured ? (
                <>
                  <p className="font-semibold text-emerald-400">Sudah diatur</p>
                  <p className="text-slate-500 truncate">
                    {current.source === "draw" ? "Digambar langsung" : "Dari foto"}
                    {current.updatedAt ? ` · ${new Date(current.updatedAt).toLocaleDateString("id-ID", { day: "numeric", month: "short", year: "numeric" })}` : ""}
                    {current.updatedByName ? ` · ${current.updatedByName}` : ""}
                  </p>
                  <p className="text-slate-500">Menyimpan yang baru akan menggantikan yang ini.</p>
                </>
              ) : (
                <p className="text-slate-400">{current === null ? "Memeriksa…" : "Belum ada tanda tangan. Buat satu di bawah."}</p>
              )}
            </div>
            {current?.configured && (confirmDelete ? (
              <div className="flex items-center gap-1.5 shrink-0">
                <button onClick={handleDelete} disabled={saving} className="px-2.5 py-1.5 bg-rose-600 hover:bg-rose-500 disabled:opacity-50 text-white text-[11px] font-bold rounded-lg cursor-pointer">Ya, hapus</button>
                <button onClick={() => setConfirmDelete(false)} disabled={saving} className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-[11px] font-semibold rounded-lg cursor-pointer">Batal</button>
              </div>
            ) : (
              <button onClick={() => setConfirmDelete(true)} disabled={saving} title="Hapus tanda tangan tersimpan" aria-label="Hapus tanda tangan tersimpan" className="p-2 text-slate-400 hover:text-rose-400 hover:bg-slate-800 rounded-lg cursor-pointer shrink-0"><Trash2 className="w-4 h-4" /></button>
            ))}
          </div>

          <div role="tablist" className="flex gap-1 bg-slate-900 border border-slate-800 rounded-xl p-1">
            {tabBtn("foto", "Unggah Foto", <ImageUp className="w-3.5 h-3.5" />)}
            {tabBtn("gambar", "Gambar Langsung", <PenLine className="w-3.5 h-3.5" />)}
          </div>

          {tab === "foto" && (
            <div className="space-y-3">
              <p className="text-[12px] text-slate-400 leading-relaxed">
                Tanda tangan di kertas putih polos dengan pulpen gelap, foto, lalu unggah. Latar kertas dihapus otomatis dan hasilnya dipotong rapat.
                Cahaya sebaiknya merata dan tanda tangan memenuhi sebagian besar foto.
              </p>
              <input ref={fileInputRef} type="file" accept="image/png,image/jpeg,image/webp" className="hidden" onChange={(e) => onPickPhoto(e.target.files?.[0])} />
              <div className="rounded-xl border border-dashed border-slate-700 bg-slate-900/40 p-3">
                <div className="h-36 rounded-lg bg-white border border-slate-300 flex items-center justify-center overflow-hidden">
                  {photoBusy ? <span className="text-xs text-slate-500">Memproses foto…</span>
                    : photo ? <img src={photo.dataUrl} alt="Hasil tanda tangan dari foto" className="max-w-full max-h-full object-contain p-2" />
                    : <span className="text-xs text-slate-400 px-4 text-center">Hasil akan tampil di sini</span>}
                </div>
                <div className="mt-3 flex items-center justify-between gap-3">
                  <button type="button" onClick={() => fileInputRef.current?.click()} disabled={photoBusy || saving} className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 disabled:opacity-50 text-slate-100 text-xs font-semibold rounded-lg flex items-center gap-1.5 cursor-pointer">
                    <ImageUp className="w-3.5 h-3.5" /> {photo ? "Ganti foto" : "Pilih foto"}
                  </button>
                  {photo && <span className="text-[11px] text-slate-500">Periksa hasilnya sebelum disimpan.</span>}
                </div>
              </div>
              {photoError && <p role="alert" className="text-[12px] text-rose-400">{photoError}</p>}
            </div>
          )}

          {tab === "gambar" && (
            <div className="space-y-3">
              <p className="text-[12px] text-slate-400 leading-relaxed">
                Goreskan tanda tangan di kotak putih dengan mouse, jari, atau stylus. Tekan <b className="text-slate-200">Hapus</b> untuk mengulang, <b className="text-slate-200">Simpan</b> bila sudah sesuai.
              </p>
              <div className="relative rounded-xl bg-white border border-slate-300 shadow-inner overflow-hidden" style={{ aspectRatio: `${PAD_W} / ${PAD_H}` }}>
                <canvas
                  ref={padRef}
                  width={PAD_W}
                  height={PAD_H}
                  aria-label="Kotak untuk menggambar tanda tangan"
                  className="absolute inset-0 w-full h-full cursor-crosshair"
                  style={{ touchAction: "none" }}
                  onPointerDown={onPadDown}
                  onPointerMove={onPadMove}
                  onPointerUp={onPadUp}
                  onPointerCancel={onPadUp}
                  onContextMenu={(e) => e.preventDefault()}
                />
                {/* Garis bantu & petunjuk — elemen biasa, BUKAN bagian kanvas, jadi tidak ikut tersimpan. */}
                <div className="pointer-events-none absolute left-[6%] right-[6%] bottom-[24%] border-b border-dashed border-slate-300" />
                {!hasInk && <span className="pointer-events-none absolute inset-0 flex items-center justify-center text-sm text-slate-300 select-none">Tanda tangan di sini</span>}
              </div>
              <div className="flex">
                <button type="button" onClick={clearPad} disabled={!hasInk || saving} className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 disabled:opacity-40 text-slate-100 text-xs font-semibold rounded-lg flex items-center gap-1.5 cursor-pointer">
                  <Eraser className="w-3.5 h-3.5" /> Hapus
                </button>
              </div>
            </div>
          )}

          <p className="text-[11px] text-slate-500 leading-relaxed">
            Tanda tangan ini otomatis menempati kolom TTD <b className="text-slate-400">Pihak Pertama</b> pada dokumen yang tombol <b className="text-slate-400">TTD Digital</b>-nya dinyalakan.
            Ini gambar tanda tangan, bukan tanda tangan elektronik tersertifikasi — berkas resmi yang mengikat tetap salinan bertanda tangan yang diunggah lewat "Unggah Bukti TTD".
          </p>
        </div>

        <div className="p-5 border-t border-slate-800 bg-slate-900/50 flex justify-end gap-3">
          <button onClick={onClose} disabled={saving} className="px-4 py-2 bg-slate-850 hover:bg-slate-800 text-slate-400 text-xs font-semibold rounded-xl cursor-pointer">Batal</button>
          <button onClick={handleSave} disabled={!canSave} className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-40 disabled:cursor-not-allowed text-white text-xs font-bold rounded-xl flex items-center gap-1.5 cursor-pointer">
            <Save className="w-3.5 h-3.5" /> {saving ? "Menyimpan…" : "Simpan"}
          </button>
        </div>
      </div>
    </div>
  );
};

export default DigitalSignatureEditor;
