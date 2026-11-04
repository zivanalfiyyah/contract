// Editor Word sungguhan (ONLYOFFICE Document Server) untuk dokumen .docx.
// Tampilan, tabel, kotak teks, kolom, header/footer, daftar, dan copy-paste memakai
// mesin editor ONLYOFFICE — bukan lagi render HTML buatan sendiri. Server menyimpan
// hasil edit sebagai versi baru lewat callback (lihat onlyoffice.ts / server.ts).
import React, { useEffect, useRef, useState } from "react";
import { AlertTriangle, Loader2 } from "lucide-react";

declare global {
  interface Window { DocsAPI?: any }
}

const scriptLoads = new Map<string, Promise<void>>();
function loadDocsApi(serverUrl: string): Promise<void> {
  const src = `${serverUrl}/web-apps/apps/api/documents/api.js`;
  const known = scriptLoads.get(src);
  if (known) return known;
  const p = new Promise<void>((resolve, reject) => {
    if (window.DocsAPI) return resolve();
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => (window.DocsAPI ? resolve() : reject(new Error("api.js termuat tetapi DocsAPI tidak ada")));
    s.onerror = () => reject(new Error("Document Server tidak bisa dijangkau"));
    document.head.appendChild(s);
  });
  p.catch(() => scriptLoads.delete(src)); // boleh dicoba lagi setelah gagal
  scriptLoads.set(src, p);
  return p;
}

interface Props {
  key?: React.Key;
  contractId: string;
  /** "current" atau nomor versi. */
  version: number | "current";
  /** false => mode lihat saja. Server tetap memutuskan hak edit sebenarnya. */
  canEdit: boolean;
  /** Dipanggil (sekali per perubahan) setelah editor melaporkan dokumen sudah tersimpan. */
  onSaved?: () => void;
  /** Dipanggil bila editor tidak bisa dimuat — pemanggil beralih ke editor cadangan. */
  onUnavailable?: (reason: string) => void;
}

export default function OnlyOfficeEditor({ contractId, version, canEdit, onSaved, onUnavailable }: Props) {
  const holderId = useRef(`oo-${Math.random().toString(36).slice(2, 10)}`).current;
  const editorRef = useRef<any>(null);
  const [state, setState] = useState<"loading" | "ready" | "error">("loading");
  const [error, setError] = useState("");
  const onSavedRef = useRef(onSaved);
  const onUnavailableRef = useRef(onUnavailable);
  onSavedRef.current = onSaved;
  onUnavailableRef.current = onUnavailable;

  useEffect(() => {
    let cancelled = false;
    let modified = false;
    setState("loading");
    (async () => {
      try {
        const r = await fetch(`/api/contracts/${contractId}/onlyoffice/config?version=${version}${canEdit ? "" : "&mode=view"}`);
        const data = await r.json().catch(() => ({}));
        if (!r.ok) throw new Error(data.error || `Gagal menyiapkan editor (HTTP ${r.status})`);
        if (!data.enabled) throw new Error("Editor Word belum diaktifkan di server");
        await loadDocsApi(data.documentServerUrl);
        if (cancelled) return;
        const config = {
          ...data.config,
          events: {
            onAppReady: () => { if (!cancelled) setState("ready"); },
            onDocumentStateChange: (e: { data: boolean }) => {
              // true = ada perubahan belum tersimpan, false = sudah tersimpan di server.
              if (e.data) { modified = true; return; }
              if (modified) { modified = false; setTimeout(() => { if (!cancelled) onSavedRef.current?.(); }, 1500); }
            },
            onError: (e: any) => {
              if (cancelled) return;
              setError(`Editor melaporkan kesalahan: ${e?.data?.errorDescription || e?.data || "tidak diketahui"}`);
              setState("error");
            },
          },
        };
        editorRef.current = new window.DocsAPI.DocEditor(holderId, config);
      } catch (e: any) {
        if (cancelled) return;
        setError(e?.message || "Editor Word tidak bisa dimuat");
        setState("error");
        onUnavailableRef.current?.(e?.message || "");
      }
    })();
    return () => {
      cancelled = true;
      try { editorRef.current?.destroyEditor?.(); } catch { /* noop */ }
      editorRef.current = null;
    };
  }, [contractId, version, canEdit, holderId]);

  return (
    <div className="relative bg-white" style={{ height: "78vh", minHeight: 480 }} data-testid="onlyoffice-editor">
      {state === "loading" && (
        <div className="absolute inset-0 z-10 flex items-center justify-center text-slate-500 text-xs gap-2 bg-white/80 pointer-events-none">
          <Loader2 className="w-4 h-4 animate-spin" /> Memuat editor Word…
        </div>
      )}
      {state === "error" && (
        <div className="absolute inset-0 z-10 p-6 text-xs text-rose-600 flex items-start gap-2 bg-white">
          <AlertTriangle className="w-4 h-4 shrink-0" /> {error}
        </div>
      )}
      <div id={holderId} style={{ width: "100%", height: "100%" }} />
    </div>
  );
}
