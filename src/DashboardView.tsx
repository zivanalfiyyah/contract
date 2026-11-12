// Dashboard Manajemen Kontrak — tata letak "selayar" (satu layar, tanpa halaman memanjang).
// Kartu yang isinya melebihi tinggi kartu discroll di dalam kartu dengan efek blur di tepi bawah/atas.
// Isi (sesuai urutan): 1 Kontrak Aktif · 2 Draft Kontrak · 3 Segera Berakhir · 4 Portofolio Value (GA)
// · 5 Perlu Persetujuan Saya · 6 Produktivitas Pengguna & Log Audit · 7 Sharing Fee Kontrak Aktif
// · 8 Peringatan Kadaluwarsa Kontrak & Mesin Perpanjangan.
import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  FileCheck2, FilePen, AlertTriangle, TrendingUp, Stamp, History, RefreshCw, CheckCircle2, AlertCircle, FilePlus, CreditCard,
} from "lucide-react";
import type { Contract, AuditTrail } from "./types";

const ACCENT = "#635BFF";

// ---------------------------------------------------------------------------
// Scroll di dalam kartu dengan blur/memudar di tepi yang masih ada isinya.
// ---------------------------------------------------------------------------
function BlurScroll({ children, className = "" }: { children: React.ReactNode; className?: string }) {
  const ref = useRef<HTMLDivElement>(null);
  const [edge, setEdge] = useState({ top: false, bottom: false });
  const check = () => {
    const el = ref.current;
    if (!el) return;
    const next = { top: el.scrollTop > 4, bottom: el.scrollHeight - el.scrollTop - el.clientHeight > 4 };
    setEdge((prev) => (prev.top === next.top && prev.bottom === next.bottom ? prev : next));
  };
  useEffect(() => {
    check();
    const el = ref.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(check);
    ro.observe(el);
    Array.from(el.children).forEach((c) => ro.observe(c as Element));
    return () => ro.disconnect();
  });
  const top = edge.top ? "transparent 0, #000 28px" : "#000 0";
  const bottom = edge.bottom ? "#000 calc(100% - 36px), transparent 100%" : "#000 100%";
  const mask = `linear-gradient(to bottom, ${top}, ${bottom})`;
  return (
    <div className="relative flex-1 min-h-0">
      <div
        ref={ref}
        onScroll={check}
        className={`absolute inset-0 overflow-y-auto ${className}`}
        style={{ WebkitMaskImage: mask, maskImage: mask }}
      >
        {children}
      </div>
    </div>
  );
}

// Kelas disalin dari HTML referensi Prodex (nama "slate" -> "rs", "indigo" -> "ri": palet Tailwind asli, lihat index.css).
const CARD = "bg-white dark:bg-rs-900 rounded-2xl border border-rs-200/80 dark:border-rs-800";
const PILL = "text-xs font-bold text-rs-500 hover:text-rs-800 dark:hover:text-rs-200 bg-rs-100 dark:bg-rs-800 px-3 py-1 rounded-xl";

function Panel({ icon: Icon, title, big, right, children, className = "" }: {
  icon?: any; title: string; big?: boolean; right?: React.ReactNode; children: React.ReactNode; className?: string;
}) {
  return (
    <section className={`${CARD} p-5 flex flex-col min-w-0 ${className}`}>
      <div className="flex items-center justify-between gap-3 mb-4 shrink-0">
        <div className="flex items-center gap-2 min-w-0">
          {Icon && <Icon className="w-4 h-4 text-rs-400 shrink-0" />}
          <h3 className={`${big ? "text-sm" : "text-xs"} font-bold text-rs-900 dark:text-white truncate`}>{title}</h3>
        </div>
        {right}
      </div>
      {children}
    </section>
  );
}

function StatCard({ icon: Icon, label, value, title, tone }: { icon: any; label: string; value: string; title?: string; tone: string }) {
  return (
    <div className={`${CARD} p-4 flex items-center gap-4 min-w-0`}>
      <div className={`w-12 h-12 rounded-2xl flex items-center justify-center shrink-0 ${tone}`}>
        <Icon className="w-5 h-5" />
      </div>
      <div className="min-w-0">
        <span className="text-[11px] font-semibold text-rs-400 dark:text-rs-500 block truncate">{label}</span>
        <span className={`${value.length > 9 ? "text-lg" : "text-xl"} font-extrabold text-rs-900 dark:text-white tracking-tight block truncate`} title={title}>{value}</span>
      </div>
    </div>
  );
}

// Kategori kegiatan -> ikon, warna & lencana (seperti 4 contoh kegiatan di referensi).
function activityKind(action: string) {
  const a = (action || "").toLowerCase();
  if (/approv|setuju|putus|reject|tolak/.test(a)) return { Icon: CheckCircle2, label: "Approval", box: "bg-ri-50 dark:bg-ri-950/50 text-[#635BFF]", badge: "text-ri-600 bg-ri-50 dark:bg-ri-950/60 dark:text-ri-300" };
  if (/expire|kadaluwarsa|alert|reminder|renew|perpanjang/.test(a)) return { Icon: AlertCircle, label: "System Alert", box: "bg-amber-50 dark:bg-amber-950/50 text-amber-500", badge: "text-rose-500 bg-rose-50 dark:bg-rose-950/60" };
  if (/draft|create|upload|buat|template|clause/.test(a)) return { Icon: FilePlus, label: "Draft", box: "bg-purple-50 dark:bg-purple-950/50 text-purple-600", badge: "text-purple-600 bg-purple-50 dark:bg-purple-950/60 dark:text-purple-300" };
  if (/fee|payment|bayar|invoice/.test(a)) return { Icon: CreditCard, label: "Payment", box: "bg-emerald-50 dark:bg-emerald-950/50 text-emerald-600", badge: "text-emerald-600 bg-emerald-50 dark:bg-emerald-950/60" };
  return { Icon: History, label: "Aktivitas", box: "bg-rs-100 dark:bg-rs-800 text-rs-500", badge: "text-rs-500 bg-rs-100 dark:bg-rs-800" };
}

const rpShort = (v: number) =>
  v >= 1e12 ? `${(v / 1e12).toLocaleString("id-ID", { maximumFractionDigits: 1 })} T`
  : v >= 1e9 ? `${(v / 1e9).toLocaleString("id-ID", { maximumFractionDigits: 1 })} M`
  : v >= 1e6 ? `${(v / 1e6).toLocaleString("id-ID", { maximumFractionDigits: 1 })} jt`
  : v >= 1e3 ? `${Math.round(v / 1e3).toLocaleString("id-ID")} rb`
  : `${Math.round(v)}`;

// ---------------------------------------------------------------------------
type Period = "bulan" | "triwulan" | "tahun";

function buildBuckets(period: Period, now: Date) {
  const y = now.getFullYear(), m = now.getMonth();
  if (period === "bulan") {
    return Array.from({ length: 6 }, (_, i) => {
      const d = new Date(y, m - (5 - i), 1);
      return { label: d.toLocaleDateString("id-ID", { month: "short" }), match: (t: Date) => t.getFullYear() === d.getFullYear() && t.getMonth() === d.getMonth() };
    });
  }
  if (period === "triwulan") {
    const q = Math.floor(m / 3);
    return Array.from({ length: 4 }, (_, i) => {
      const idx = y * 4 + q - (3 - i);
      const yy = Math.floor(idx / 4), qq = ((idx % 4) + 4) % 4;
      return { label: `Q${qq + 1} ${String(yy).slice(2)}`, match: (t: Date) => t.getFullYear() === yy && Math.floor(t.getMonth() / 3) === qq };
    });
  }
  return Array.from({ length: 3 }, (_, i) => {
    const yy = y - (2 - i);
    return { label: String(yy), match: (t: Date) => t.getFullYear() === yy };
  });
}

export interface DashboardApprovalItem {
  contract: Contract;
  step: { order: number; approverName?: string };
  isMine: boolean;
  asDelegate: boolean;
  totalSteps: number;
  waitingSince: string;
}

export interface DashboardViewProps {
  contracts: Contract[];
  activeContracts: Contract[];
  draftCount: number;
  expiring: Contract[];
  valueText: string;
  valueFull: string;
  sharingStats: { hasAny: boolean; nominalSum: number; nominalContracts: number; pctContracts: number };
  approvals: DashboardApprovalItem[];
  slaDays: number;
  audits: AuditTrail[];
  roleLabel: (r: string) => string;
  daysRemaining: (d: string) => number;
  onOpenContract: (c: Contract) => void;
  onRenew: (c: Contract) => void;
  onOpenAudits: () => void;
}

export default function DashboardView(p: DashboardViewProps) {
  const [period, setPeriod] = useState<Period>("bulan");
  const [hover, setHover] = useState<number | null>(null);

  // ----- 7. Sharing fee: nominal per periode (persentase tidak dijumlah ke rupiah) -----
  const bars = useMemo(() => {
    const buckets = buildBuckets(period, new Date());
    return buckets.map((b) => {
      const sum = p.activeContracts.reduce((s, c) => {
        const t = new Date(c.startDate);
        if (isNaN(t.getTime()) || !b.match(t)) return s;
        const items = (c as any).sharingFeeItems || [];
        return s + items.filter((it: any) => it.feeType === "nominal").reduce((a: number, it: any) => a + (Number(it.value) || 0), 0);
      }, 0);
      return { label: b.label, value: sum };
    });
  }, [period, p.activeContracts]);
  const maxBar = Math.max(1, ...bars.map((b) => b.value));
  const pctShare = p.activeContracts.length ? Math.round((p.sharingStats.pctContracts / p.activeContracts.length) * 100) : 0;

  // ----- 8. Kadaluwarsa & mesin perpanjangan -----
  const expiryRows = useMemo(() => {
    return [...p.expiring]
      .map((c) => {
        const hasRenewalDraft = p.contracts.some((x) => (x as any).renewedFromId === c.id);
        const state: "draft" | "auto" | "action" = hasRenewalDraft ? "draft" : (c as any).isAutoRenew ? "auto" : "action";
        return { c, days: p.daysRemaining(c.endDate), state };
      })
      .sort((a, b) => a.days - b.days);
  }, [p.expiring, p.contracts]);
  const segs = [
    { key: "action", label: "Perlu tindakan", color: "#EF4444", n: expiryRows.filter((r) => r.state === "action").length },
    { key: "draft", label: "Draf perpanjangan dibuat", color: ACCENT, n: expiryRows.filter((r) => r.state === "draft").length },
    { key: "auto", label: "Perpanjangan otomatis", color: "#F59E0B", n: expiryRows.filter((r) => r.state === "auto").length },
  ];
  const totalExp = expiryRows.length;
  const R = 40, CIRC = 2 * Math.PI * R;
  let acc = 0;

  const hari = (since: string) => {
    const d = Math.max(0, Math.floor((Date.now() - new Date(since).getTime()) / 86400000));
    return Number.isFinite(d) ? d : 0;
  };

  const periodTabs = ([["bulan", "Bulanan"], ["triwulan", "Triwulan"], ["tahun", "Tahunan"]] as [Period, string][]);
  const DR = 43.75, DC = 2 * Math.PI * DR; // donat tebal 12,5% (cutout 75% seperti Chart.js referensi)
  let dacc = 0;

  return (
    <div className="space-y-6">
      {/* 1-4 */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <StatCard icon={FileCheck2} label="Kontrak Aktif" value={String(p.activeContracts.length)} tone="bg-ri-50 dark:bg-ri-950/50 text-[#635BFF]" />
        <StatCard icon={FilePen} label="Draft Kontrak" value={String(p.draftCount)} tone="bg-purple-50 dark:bg-purple-950/50 text-purple-600" />
        <StatCard icon={AlertTriangle} label="Segera Berakhir" value={String(p.expiring.length)} tone="bg-amber-50 dark:bg-amber-950/50 text-amber-500" />
        <StatCard icon={TrendingUp} label="Portofolio Value (GA)" value={p.valueText} title={p.valueFull} tone="bg-emerald-50 dark:bg-emerald-950/50 text-emerald-600" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* 7. Sharing fee */}
        <section className={`lg:col-span-7 ${CARD} p-5 flex flex-col justify-between min-w-0`}>
          <div>
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
              <div>
                <h3 className="text-sm font-bold text-rs-900 dark:text-white">Sharing Fee Kontrak Aktif</h3>
                <div className="flex items-center gap-4 mt-1 text-[11px] text-rs-500 dark:text-rs-400 font-medium">
                  <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-[#635BFF]" />Fee Tetap</span>
                  <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-ri-200 dark:bg-ri-400" />Fee Persentase: {p.sharingStats.pctContracts} kontrak · {pctShare}%</span>
                </div>
              </div>
              <div className="flex items-center bg-rs-100 dark:bg-rs-800 p-1 rounded-xl self-start text-[11px] font-bold text-rs-500">
                {periodTabs.map(([k, l]) => (
                  <button key={k} type="button" onClick={() => { setPeriod(k); setHover(null); }}
                    className={`px-3 py-1 rounded-lg cursor-pointer ${period === k ? "bg-white dark:bg-rs-700 text-rs-900 dark:text-white" : "hover:text-rs-900 dark:hover:text-white"}`}>{l}</button>
                ))}
              </div>
            </div>
            {!p.sharingStats.hasAny ? (
              <div className="h-64 flex flex-col items-center justify-center text-center text-xs text-rs-400 font-medium">
                <CheckCircle2 className="w-6 h-6 mb-1.5 text-[#635BFF]" />
                Belum ada kontrak aktif dengan sharing fee.
              </div>
            ) : (
              <div className="relative h-64 w-full flex gap-2">
                <div className="flex flex-col justify-between text-[10px] text-rs-400 tabular-nums pb-6 text-right w-10 shrink-0">
                  {[1, 0.66, 0.33, 0].map((f) => <span key={f}>{f === 0 ? "0" : rpShort(maxBar * f)}</span>)}
                </div>
                <div className="relative flex-1 min-w-0">
                  <div className="absolute inset-x-0 top-1.5 bottom-6 flex flex-col justify-between pointer-events-none">
                    {[0, 1, 2, 3].map((i) => <div key={i} className="border-t border-rs-100 dark:border-rs-800" />)}
                  </div>
                  <div className="absolute inset-x-0 top-0 bottom-0 flex items-stretch justify-around">
                    {bars.map((b, i) => (
                      <div key={i} className="flex-1 flex flex-col items-center min-w-0" onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(null)}>
                        <div className="relative flex-1 w-full flex items-end justify-center pt-1.5">
                          <span
                            className="w-7 max-w-[60%] rounded-md transition-all"
                            style={{ height: `${b.value === 0 ? 2 : Math.max(5, (b.value / maxBar) * 100)}%`, background: hover === i ? "#4f46e5" : ACCENT, opacity: b.value === 0 ? 0.35 : 1 }}
                          />
                          {hover === i && (
                            <div className="absolute z-10 top-0 left-1/2 -translate-x-1/2 bg-white dark:bg-rs-800 border border-rs-200 dark:border-rs-700 shadow-xl rounded-xl p-2.5 text-[11px] whitespace-nowrap pointer-events-none">
                              <div className="flex items-center gap-1.5 font-bold text-rs-700 dark:text-rs-300"><span className="w-1.5 h-1.5 rounded-full bg-[#635BFF]" />Fee Tetap · {b.label}</div>
                              <div className="font-extrabold text-rs-900 dark:text-white ml-3">Rp {b.value.toLocaleString("id-ID")}</div>
                            </div>
                          )}
                        </div>
                        <span className="h-6 text-[10px] leading-6 text-rs-400">{b.label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        </section>

        {/* 8. Kadaluwarsa & mesin perpanjangan */}
        <section className={`lg:col-span-5 ${CARD} p-5 flex flex-col justify-between min-w-0`}>
          <div>
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-xs font-bold text-rs-900 dark:text-white">Peringatan Kadaluwarsa & Mesin Perpanjangan</h3>
              <span className={`${PILL} whitespace-nowrap tabular-nums`}>{totalExp} kontrak</span>
            </div>
            <div className="relative h-48 w-full flex items-center justify-center my-2">
              <svg viewBox="0 0 100 100" className="h-48 w-48 -rotate-90">
                <circle cx="50" cy="50" r={DR} fill="none" strokeWidth="12.5" className="stroke-rs-100 dark:stroke-rs-800" />
                {totalExp > 0 && segs.map((sg) => {
                  if (!sg.n) return null;
                  const len = (sg.n / totalExp) * DC;
                  const el = <circle key={sg.key} cx="50" cy="50" r={DR} fill="none" stroke={sg.color} strokeWidth="12.5" strokeDasharray={`${len} ${DC}`} strokeDashoffset={-dacc} />;
                  dacc += len;
                  return el;
                })}
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center text-center pointer-events-none">
                <span className="text-[10px] font-semibold text-rs-400">Total Kadaluwarsa</span>
                <span className="text-base font-extrabold text-rs-900 dark:text-white">{totalExp} Kontrak</span>
              </div>
            </div>
            <div className="space-y-2 mt-4 text-xs font-semibold">
              {segs.map((sg) => (
                <div key={sg.key} className="flex items-center justify-between">
                  <div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full" style={{ background: sg.color }} /><span className="text-rs-600 dark:text-rs-300">{sg.label}</span></div>
                  <div className="text-rs-900 dark:text-white font-bold">{sg.n} <span className="text-rs-400 font-medium ml-1">{totalExp ? Math.round((sg.n / totalExp) * 100) : 0}%</span></div>
                </div>
              ))}
            </div>
            {expiryRows.length > 0 && (
              <div className="mt-4 border-t border-rs-100 dark:border-rs-800 max-h-56 overflow-y-auto">
                {expiryRows.map(({ c, days, state }) => (
                  <div key={c.id} className="flex items-center gap-2.5 py-2.5 border-b border-rs-50 dark:border-rs-800/60 last:border-0">
                    <div className="min-w-0 flex-1">
                      <p className="text-xs font-bold text-rs-900 dark:text-white truncate">{c.title}</p>
                      <p className="text-[11px] text-rs-400 truncate">
                        <span className={days <= 7 ? "text-rose-500 font-semibold" : "text-amber-500 font-semibold"}>{days} hari lagi</span>
                        {" · "}{state === "draft" ? "draf perpanjangan ada" : state === "auto" ? "perpanjangan otomatis" : "perlu tindakan"}
                      </p>
                    </div>
                    <button onClick={() => p.onOpenContract(c)} className="bg-rs-100 dark:bg-rs-800 text-rs-600 dark:text-rs-300 text-[10px] font-bold px-2 py-1 rounded-lg cursor-pointer shrink-0">Tinjau</button>
                    <button onClick={() => p.onRenew(c)} className="bg-[#635BFF] hover:bg-ri-700 text-white text-[10px] font-bold px-2.5 py-1 rounded-lg cursor-pointer shrink-0 flex items-center gap-1">
                      <RefreshCw className="w-3 h-3" /> {state === "draft" ? "Buka" : "Perpanjang"}
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>
        </section>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* 6. Produktivitas & log audit */}
        <Panel icon={History} title="Produktivitas Pengguna & Log Audit Aktivitas" className="lg:col-span-5 justify-between"
          right={<button onClick={p.onOpenAudits} className={`${PILL} shrink-0`}>Semua</button>}>
          <div className="space-y-4 max-h-[420px] overflow-y-auto">
            {p.audits.length === 0 && <p className="text-xs text-rs-400 text-center py-6">Belum ada aktivitas.</p>}
            {p.audits.slice(0, 12).map((a) => {
              const k = activityKind(a.action);
              return (
                <div key={a.id} className="flex items-center justify-between gap-3" title={a.details}>
                  <div className="flex items-center gap-3 min-w-0">
                    <div className={`w-8 h-8 rounded-full flex items-center justify-center shrink-0 ${k.box}`}><k.Icon className="w-4 h-4" /></div>
                    <div className="min-w-0">
                      <h4 className="text-xs font-bold text-rs-900 dark:text-white truncate">{a.action}</h4>
                      <p className="text-[11px] text-rs-400 font-medium truncate">{a.userName} • {new Date(a.timestamp).toLocaleDateString("id-ID", { day: "numeric", month: "short", year: "2-digit" })}</p>
                    </div>
                  </div>
                  <span className={`text-[10px] font-bold px-2.5 py-1 rounded-lg shrink-0 ${k.badge}`}>{k.label}</span>
                </div>
              );
            })}
          </div>
        </Panel>

        {/* 5. Perlu persetujuan saya */}
        <Panel icon={Stamp} title="Perlu Persetujuan Saya" className="lg:col-span-7 justify-between"
          right={p.approvals.length > 0 ? <span className="text-[11px] font-bold px-2.5 py-1 rounded-xl bg-rose-50 dark:bg-rose-950/60 text-rose-500 tabular-nums">{p.approvals.length}</span> : null}>
          {p.approvals.length === 0 ? (
            <div className="flex-1 min-h-[180px] flex flex-col items-center justify-center text-center text-xs text-rs-400 font-medium">
              <CheckCircle2 className="w-6 h-6 mb-1.5 text-[#635BFF]" />
              Tidak ada dokumen yang menunggu persetujuan Anda.
            </div>
          ) : (
            <div className="overflow-x-auto max-h-[420px] overflow-y-auto">
              <table className="plain-table w-full text-left border-collapse">
                <thead>
                  <tr className="border-b border-rs-100 dark:border-rs-800 text-[10px] font-bold text-rs-400 uppercase tracking-wider">
                    <th className="py-2.5 px-2 font-bold">Nama Kontrak</th>
                    <th className="py-2.5 px-2 text-center font-bold">Langkah</th>
                    <th className="py-2.5 px-2 text-center font-bold">Menunggu</th>
                    <th className="py-2.5 px-2 text-center font-bold">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-rs-50 dark:divide-rs-800/60 text-xs font-semibold text-rs-700 dark:text-rs-300">
                  {p.approvals.map(({ contract: c, step, isMine, asDelegate, totalSteps, waitingSince }) => {
                    const d = hari(waitingSince);
                    const late = p.slaDays > 0 && d >= p.slaDays;
                    return (
                      <tr key={c.id} className="hover:bg-rs-50/50 dark:hover:bg-rs-800/40">
                        <td className="py-3 px-2">
                          <span className="font-bold text-rs-900 dark:text-white block">{c.title}</span>
                          <span className="text-[10px] text-rs-400 font-normal">{c.contractNumber}{asDelegate && !isMine ? ` · pengganti ${step.approverName || ""}` : ""}</span>
                        </td>
                        <td className="py-3 px-2 text-center font-medium text-rs-500">{step.order}/{totalSteps}</td>
                        <td className={`py-3 px-2 text-center font-medium ${late ? "text-rose-500" : "text-rs-500"}`}>{d === 0 ? "hari ini" : `${d} hari`}</td>
                        <td className="py-3 px-2 text-center">
                          <button onClick={() => p.onOpenContract(c)} className={isMine ? "bg-[#635BFF] hover:bg-ri-700 text-white text-[10px] font-bold px-2.5 py-1 rounded-lg cursor-pointer" : "bg-rs-100 dark:bg-rs-800 text-rs-600 dark:text-rs-300 text-[10px] font-bold px-2 py-1 rounded-lg cursor-pointer"}>
                            {isMine ? "Putuskan" : "Lihat"}
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </Panel>
      </div>
    </div>
  );
}
