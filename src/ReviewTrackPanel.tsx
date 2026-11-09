// Tinjauan dengan Track Changes (internal & eksternal): dokumen dibuka di editor Word dalam mode tinjau —
// setiap tambah/ganti/coret tercatat dengan nama + waktu, komentar juga. Riwayat siapa mengubah apa dilihat
// LANGSUNG DI DALAM editor (panel Review / Komentar & balon perubahan bawaan Word), tidak ada panel terpisah.
import React from "react";
import { Download } from "lucide-react";
import OnlyOfficeEditor from "./OnlyOfficeEditor";

interface Props {
  contractId: string;
  mode: "internal" | "external";
  /** Token link review (mode external). */
  token?: string;
  /** Nama tamu (mode external) — dicatat sebagai pengubah. */
  reviewerName?: string;
  canApply?: boolean;
  onApplied?: () => void;
  onUnavailable?: (reason: string) => void;
  showToast?: (msg: string, type?: "success" | "warning" | "info") => void;
  askConfirm?: (msg: string, opts?: { title?: string; confirmLabel?: string; danger?: boolean }) => Promise<boolean>;
  height?: string;
}

export default function ReviewTrackPanel({ contractId, mode, token, reviewerName, onUnavailable, height }: Props) {
  const configUrl = mode === "external"
    ? `/api/external-review/${token}/review-doc/config?name=${encodeURIComponent(reviewerName || "")}`
    : `/api/contracts/${contractId}/review-doc/config`;
  return (
    <div className="flex-1 min-h-0 flex flex-col bg-slate-950" data-testid="review-track-panel">
      <div className="shrink-0 px-4 py-2 flex items-center gap-3 flex-wrap border-b border-slate-800 bg-slate-900/60 text-[11px] text-slate-400">
        <span className="flex-1 min-w-48">
          Ubah teks langsung di dokumen — setiap perubahan &amp; komentar tercatat dengan <b className="text-slate-200">nama dan waktu</b>.
          Lihat riwayatnya lewat tab <b className="text-slate-200">Review</b> / ikon komentar di editor.
          {mode === "external" ? ` Nama Anda tercatat sebagai "${reviewerName}".` : ""}
        </span>
        {mode === "internal" && (
          <>
            <a href={`/api/contracts/${contractId}/review-doc/file?kind=marked`} className="px-2.5 py-1.5 text-[10px] font-bold rounded-lg border bg-slate-950 border-slate-800 text-slate-200 hover:border-indigo-500/50 flex items-center gap-1"><Download className="w-3 h-3" /> Unduh versi coretan</a>
            <a href={`/api/contracts/${contractId}/review-doc/file?kind=clean`} className="px-2.5 py-1.5 text-[10px] font-bold rounded-lg border bg-slate-950 border-slate-800 text-slate-200 hover:border-indigo-500/50 flex items-center gap-1"><Download className="w-3 h-3" /> Unduh versi bersih</a>
          </>
        )}
      </div>
      <OnlyOfficeEditor
        key={`rv-${contractId}-${mode}-${reviewerName || ""}`}
        contractId={contractId} version="current" canEdit={false} configUrl={configUrl} height={height || "80vh"}
        onUnavailable={onUnavailable}
      />
    </div>
  );
}
