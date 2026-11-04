// Opsi paste ala Microsoft Word, dipakai SEMUA editor dokumen (template maupun upload):
//
//   Ikuti Format Dokumen  (Word: Merge Formatting)       — bawaan Ctrl+V
//   Format Asal           (Word: Keep Source Formatting)
//   Teks Saja             (Word: Keep Text Only)         — Ctrl+Shift+V
//
// Alurnya mirip Word: paste langsung memakai pilihan bawaan, lalu popup kecil
// muncul di dekat kursor untuk mengganti pilihan. Mengganti pilihan = membatalkan
// paste tadi (undo) lalu menempel ulang dgn pilihan baru. Popup hilang begitu
// pengguna mengetik, mengklik di luar, menggulir, atau menekan Esc.
import React, { useEffect, useState } from "react";
import { cleanPastedHtml, type PasteMode } from "./paste-clean";

export type PasteChoice = "merge" | "source" | "text";

export interface PasteOptions {
  /** "rich" = editor template (HTML), "flat" = editor .docx upload. */
  mode: PasteMode;
  /** Pembersih tambahan utk HTML hasil paste (mis. DOMPurify di editor template). */
  sanitize?: (html: string) => string;
  /** Dipanggil setiap kali isi editor berubah karena paste. */
  onChange?: () => void;
}

// ---- Shortcut Ctrl/Cmd+Shift+V = paste teks saja ----------------------------
// Browser tidak memberi tahu event "paste" apakah ia berasal dari Ctrl+Shift+V,
// jadi tombolnya dicatat sendiri.
let lastPlainShortcutAt = 0;
if (typeof window !== "undefined" && !(window as any).__pasteShortcutInstalled) {
  (window as any).__pasteShortcutInstalled = true;
  window.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && (e.key === "v" || e.key === "V")) lastPlainShortcutAt = Date.now();
  }, true);
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function insertHtml(html: string): boolean {
  try { document.execCommand("styleWithCSS", false, "true"); } catch { /* noop */ }
  return document.execCommand("insertHTML", false, html);
}

/** Sisipkan isi clipboard sesuai pilihan. true kalau ada yang disisipkan. */
export function performPaste(clip: { html: string; text: string }, choice: PasteChoice, o: PasteOptions): boolean {
  if (choice !== "text" && clip.html) {
    let cleaned = cleanPastedHtml(clip.html, o.mode, choice === "source" ? "source" : "merge");
    if (cleaned && o.sanitize) cleaned = o.sanitize(cleaned);
    if (cleaned) return insertHtml(cleaned);
  }
  // Teks polos (pilihan "Teks Saja", atau clipboard tanpa HTML berguna).
  const t = clip.text.replace(/\r\n?/g, "\n").replace(/\n$/, "");
  if (!t) return false;
  if (o.mode === "flat") {
    const lines = t.split("\n");
    if (lines.length === 1) return document.execCommand("insertText", false, t);
    // Satu insertHTML (bukan insertParagraph+insertText berulang) supaya paste
    // ini hanya satu langkah undo — dibutuhkan utk mengganti pilihan lewat popup.
    return insertHtml(lines.map((l) => `<p>${l ? esc(l) : "<br>"}</p>`).join(""));
  }
  return insertHtml(esc(t).replace(/\n/g, "<br>"));
}

function caretRect(): DOMRect | null {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  const r = sel.getRangeAt(0).cloneRange();
  r.collapse(false);
  let rect = r.getBoundingClientRect();
  if (!rect || (rect.width === 0 && rect.height === 0)) {
    const n = r.endContainer;
    const el = n.nodeType === Node.ELEMENT_NODE ? (n as Element) : n.parentElement;
    if (!el) return null;
    rect = el.getBoundingClientRect();
  }
  return rect;
}

/**
 * Gulirkan layar ke akhir teks yang baru ditempel (seperti Word). Paste bawaan browser
 * melakukannya sendiri, tapi kita mengambil alih paste (preventDefault + insertHTML)
 * sehingga untuk tempelan panjang kursor bisa tertinggal jauh di luar layar.
 */
function revealCaret() {
  try {
    const sel = window.getSelection();
    const n = sel?.focusNode;
    const el = n ? (n.nodeType === Node.ELEMENT_NODE ? (n as Element) : n.parentElement) : null;
    el?.scrollIntoView({ block: "nearest", inline: "nearest" });
  } catch { /* noop */ }
}

export function hidePasteOptions() {
  window.dispatchEvent(new CustomEvent("docpaste:hide"));
}

/**
 * Handler event paste untuk contentEditable manapun. Selalu mengambil alih
 * paste (preventDefault), menempel dgn pilihan bawaan, lalu menampilkan popup.
 */
export function handlePaste(ev: ClipboardEvent, o: PasteOptions) {
  ev.preventDefault();
  const dt = ev.clipboardData;
  const clip = { html: dt?.getData("text/html") || "", text: dt?.getData("text/plain") || "" };
  const plain = Date.now() - lastPlainShortcutAt < 1000;
  const first: PasteChoice = plain ? "text" : "merge";
  const apply = (c: PasteChoice) => {
    const ok = performPaste(clip, c, o);
    o.onChange?.();
    if (ok) revealCaret();
    return ok;
  };
  hidePasteOptions();
  const ok = apply(first);
  // Tanpa HTML tidak ada yang bisa dipilih; Ctrl+Shift+V sudah eksplisit "teks saja".
  if (!ok || plain || !clip.html) return;
  requestAnimationFrame(() => {
    const rect = caretRect();
    if (!rect) return;
    window.dispatchEvent(new CustomEvent("docpaste:show", {
      detail: {
        rect,
        current: first,
        redo: (c: PasteChoice) => {
          try { document.execCommand("undo"); } catch { /* noop */ }
          apply(c);
          return caretRect();
        },
      },
    }));
  });
}

// ---- Popup -------------------------------------------------------------------
const CHOICES: { k: PasteChoice; label: string; tip: string }[] = [
  { k: "merge", label: "Ikuti Format Dokumen", tip: "Merge Formatting — mengikuti format dokumen ini (tebal/miring/list tetap)" },
  { k: "source", label: "Format Asal", tip: "Keep Source Formatting — format sama seperti dokumen asal (font, ukuran, warna)" },
  { k: "text", label: "Teks Saja", tip: "Keep Text Only — hanya teks, semua format dibuang" },
];

interface PopupState { left: number; bottom: number; current: PasteChoice; redo: (c: PasteChoice) => DOMRect | null }

/** Dipasang SEKALI di root App (fixed position), mendengarkan event dari handlePaste. */
export function PasteOptionsOverlay() {
  const [st, setSt] = useState<PopupState | null>(null);

  useEffect(() => {
    let shownAt = 0;
    const show = (e: Event) => {
      const d = (e as CustomEvent).detail;
      shownAt = Date.now();
      setSt({ left: d.rect.left, bottom: d.rect.bottom, current: d.current, redo: d.redo });
    };
    const hide = () => setSt(null);
    // Paste bisa menggulir editor sendiri (kursor dibawa ke tampilan) — abaikan scroll sesaat setelah popup muncul.
    const onScroll = () => { if (Date.now() - shownAt > 500) setSt(null); };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Control" || e.key === "Shift" || e.key === "Alt" || e.key === "Meta") return;
      setSt(null);
    };
    const onDown = (e: MouseEvent) => {
      if (!(e.target as HTMLElement | null)?.closest?.("[data-paste-popup]")) setSt(null);
    };
    window.addEventListener("docpaste:show", show);
    window.addEventListener("docpaste:hide", hide);
    window.addEventListener("keydown", onKey, true);
    window.addEventListener("mousedown", onDown, true);
    window.addEventListener("scroll", onScroll, true);
    window.addEventListener("resize", hide);
    return () => {
      window.removeEventListener("docpaste:show", show);
      window.removeEventListener("docpaste:hide", hide);
      window.removeEventListener("keydown", onKey, true);
      window.removeEventListener("mousedown", onDown, true);
      window.removeEventListener("scroll", onScroll, true);
      window.removeEventListener("resize", hide);
    };
  }, []);

  if (!st) return null;
  const width = 380;
  const left = Math.max(8, Math.min(st.left, window.innerWidth - width - 8));
  // Selalu di dalam layar: kalau kursor ada di bawah/atas area terlihat, popup menempel di tepi layar.
  let top = st.bottom + 6;
  if (top + 44 > window.innerHeight) top = st.bottom - 52;
  top = Math.max(8, Math.min(top, window.innerHeight - 52));

  return (
    <div
      data-paste-popup
      role="toolbar"
      aria-label="Opsi tempel"
      // mousedown di popup TIDAK boleh memindahkan fokus/seleksi dari editor (undo+tempel ulang bergantung padanya).
      onMouseDown={(e) => e.preventDefault()}
      style={{ position: "fixed", left, top, zIndex: 90 }}
      className="flex items-center gap-1 rounded-xl border border-slate-700 bg-slate-900 p-1 shadow-xl"
    >
      <span className="px-2 text-[10px] font-bold uppercase tracking-wide text-slate-500">Tempel</span>
      {CHOICES.map((c) => {
        const active = st.current === c.k;
        return (
          <button
            key={c.k}
            type="button"
            title={c.tip}
            aria-pressed={active}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              if (active) return;
              const r = st.redo(c.k);
              setSt((s) => (s ? { ...s, current: c.k, left: r ? r.left : s.left, bottom: r ? r.bottom : s.bottom } : s));
            }}
            className={`px-2.5 py-1.5 rounded-lg text-[11px] font-semibold transition cursor-pointer ${
              active ? "bg-indigo-600 text-white" : "text-slate-300 hover:bg-slate-800"
            }`}
          >
            {c.label}
          </button>
        );
      })}
    </div>
  );
}
