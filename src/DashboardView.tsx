// Dashboard Manajemen Kontrak — tata letak "selayar" (satu layar, tanpa halaman memanjang).
// Kartu yang isinya melebihi tinggi kartu discroll di dalam kartu dengan efek blur di tepi bawah/atas.
// Isi (sesuai urutan): 1 Kontrak Aktif · 2 Draft Kontrak · 3 Segera Berakhir · 4 Portofolio Value (GA)
// · 5 Perlu Persetujuan Saya · 6 Produktivitas Pengguna & Log Audit · 7 Sharing Fee Kontrak Aktif
// · 8 Peringatan Kadaluwarsa Kontrak & Mesin Perpanjangan.
import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  FileCheck2, FilePen, AlertTriangle, TrendingUp, Stamp, History, BadgePercent, CalendarClock, RefreshCw, CheckCircle2, ChevronRight,
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

function Panel({ icon: Icon, title, right, children, className = "" }: {
  icon: any; title: string; right?: React.ReactNode; children: React.ReactNode; className?: string;
}) {
  return (
    <section className={`rounded-2xl bg-slate-900 border border-slate-800 card-soft p-5 flex flex-col gap-3 min-h-0 overflow-hidden ${className}`}>
      <div className="flex items-center justify-between gap-3 shrink-0">
        <div className="flex items-center gap-2 min-w-0">
          <Icon className="w-4 h-4 text-slate-500 shrink-0" />
          <h3 className="text-xs font-bold text-slate-100 truncate">{title}</h3>
        </div>
        {right}
      </div>
      {children}
    </section>
  );
}

function StatCard({ icon: Icon, label, value, title, tone }: { icon: any; label: string; value: string; title?: string; tone: string }) {
  return (
    <div className="rounded-2xl bg-slate-900 border border-slate-800 card-soft p-4 flex items-center gap-4 min-w-0">
      <span className={`w-12 h-12 rounded-2xl flex items-center justify-center shrink-0 ${tone}`}>
        <Icon className="w-5 h-5" />
      </span>
      <div className="min-w-0">
        <p className="text-[11px] font-semibold text-slate-500 truncate">{label}</p>
        <p className={`${value.length > 9 ? "text-[17px]" : "text-[22px]"} leading-tight font-extrabold text-slate-100 tabular truncate`} title={title}>{value}</p>
      </div>
    </div>
  );
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

  return (
    <div className="flex flex-col gap-5 lg:h-[calc(100vh-132px)] lg:min-h-[620px]">
      {/* 1-4 */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-5 shrink-0">
        <StatCard icon={FileCheck2} label="Kontrak Aktif" value={String(p.activeContracts.length)} tone="bg-indigo-500/10 text-[#635BFF]" />
        <StatCard icon={FilePen} label="Draft Kontrak" value={String(p.draftCount)} tone="bg-purple-500/10 text-purple-500" />
        <StatCard icon={AlertTriangle} label="Segera Berakhir" value={String(p.expiring.length)} tone="bg-amber-500/10 text-amber-500" />
        <StatCard icon={TrendingUp} label="Portofolio Value (GA)" value={p.valueText} title={p.valueFull} tone="bg-emerald-500/10 text-emerald-500" />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 lg:grid-rows-2 gap-5 flex-1 min-h-0">
        {/* 7. Sharing fee */}
        <Panel
          icon={BadgePercent}
          title="Sharing Fee Kontrak Aktif"
          className="lg:col-span-7 h-[340px] lg:h-auto"
          right={
            <div className="flex items-center bg-slate-950 p-0.5 rounded-xl text-[10.5px] font-bold text-slate-500">
              {([["bulan", "Bulanan"], ["triwulan", "Triwulan"], ["tahun", "Tahunan"]] as [Period, string][]).map(([k, l]) => (
                <button key={k} type="button" onClick={() => { setPeriod(k); setHover(null); }}
                  className={`px-2.5 py-1 rounded-lg cursor-pointer ${period === k ? "bg-slate-900 text-slate-100 shadow-sm" : "hover:text-slate-100"}`}>{l}</button>
              ))}
            </div>
          }
        >
          {!p.sharingStats.hasAny ? (
            <div className="flex-1 flex flex-col items-center justify-center text-center text-[12px] text-slate-500">
              <BadgePercent className="w-7 h-7 mb-2 opacity-50" strokeWidth={1.5} />
              Belum ada kontrak aktif dengan sharing fee.
            </div>
          ) : (
            <>
              <div className="flex items-end justify-between gap-3 shrink-0">
                <div>
                  <p className="text-[11px] text-slate-500">Fee Tetap (nominal)</p>
                  <p className="text-[20px] leading-tight font-extrabold text-slate-100 tabular">Rp {p.sharingStats.nominalSum.toLocaleString("id-ID")}</p>
                  <p className="text-[10.5px] text-slate-500">{p.sharingStats.nominalContracts > 0 ? `dari ${p.sharingStats.nominalContracts} kontrak` : "belum ada fee nominal"}</p>
                </div>
                <div className="text-right min-w-0">
                  <p className="text-[11px] text-slate-500">Fee Persentase (dipisah)</p>
                  <p className="text-[15px] font-bold text-slate-100 tabular">{p.sharingStats.pctContracts} kontrak · {pctShare}%</p>
                  <div className="mt-1 h-1.5 w-40 max-w-full ml-auto rounded-full bg-slate-950 overflow-hidden">
                    <div className="h-full rounded-full" style={{ width: `${pctShare}%`, background: ACCENT }} />
                  </div>
                </div>
              </div>
              <div className="relative flex-1 min-h-0 flex gap-2">
                <div className="flex flex-col justify-between text-[9px] text-slate-500 tabular pb-5 text-right w-9 shrink-0">
                  {[1, 0.66, 0.33, 0].map((f) => <span key={f}>{f === 0 ? "0" : rpShort(maxBar * f)}</span>)}
                </div>
                <div className="relative flex-1 min-w-0">
                  <div className="absolute inset-x-0 top-0 bottom-5 flex flex-col justify-between pointer-events-none">
                    {[0, 1, 2, 3].map((i) => <div key={i} className="border-t border-slate-800" />)}
                  </div>
                  <div className="absolute inset-x-0 top-0 bottom-0 flex items-stretch justify-around">
                    {bars.map((b, i) => (
                      <div key={i} className="flex-1 flex flex-col items-center min-w-0" onMouseEnter={() => setHover(i)} onMouseLeave={() => setHover(null)}>
                        <div className="relative flex-1 w-full flex items-end justify-center">
                          <span
                            className="w-7 max-w-[60%] rounded-t-lg transition-all"
                            style={{ height: `${b.value === 0 ? 3 : Math.max(6, (b.value / maxBar) * 100)}%`, background: hover === i ? ACCENT : "#C7C9FF", opacity: b.value === 0 ? 0.5 : 1 }}
                          />
                          {hover === i && (
                            <div className="absolute z-10 -top-1 left-1/2 -translate-x-1/2 -translate-y-full bg-slate-900 border border-slate-800 shadow-xl rounded-xl px-2.5 py-1.5 text-[10.5px] whitespace-nowrap pointer-events-none">
                              <span className="text-slate-500">Fee Tetap · {b.label}</span>
                              <span className="block font-extrabold text-slate-100">Rp {b.value.toLocaleString("id-ID")}</span>
                            </div>
                          )}
                        </div>
                        <span className="h-5 text-[10px] leading-5 text-slate-500">{b.label}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </>
          )}
        </Panel>

        {/* 8. Kadaluwarsa & mesin perpanjangan */}
        <Panel
          icon={CalendarClock}
          title="Peringatan Kadaluwarsa & Mesin Perpanjangan"
          className="lg:col-span-5 h-[380px] lg:h-auto"
          right={<span className="text-[10.5px] font-bold px-2.5 py-1 rounded-lg bg-slate-950 text-slate-500 tabular-nums whitespace-nowrap">{totalExp} kontrak</span>}
        >
          <div className="flex items-center gap-4 shrink-0">
            <div className="relative w-[84px] h-[84px] shrink-0">
              <svg viewBox="0 0 100 100" className="w-full h-full -rotate-90">
                <circle cx="50" cy="50" r={R} fill="none" stroke="currentColor" strokeWidth="12" className="text-slate-950" />
                {totalExp > 0 && segs.map((s) => {
                  if (!s.n) return null;
                  const len = (s.n / totalExp) * CIRC;
                  const el = <circle key={s.key} cx="50" cy="50" r={R} fill="none" stroke={s.color} strokeWidth="12" strokeDasharray={`${Math.max(0, len - 1.5)} ${CIRC}`} strokeDashoffset={-acc} />;
                  acc += len;
                  return el;
                })}
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center text-center pointer-events-none">
                <span className="text-[16px] leading-none font-extrabold text-slate-100 tabular">{totalExp}</span>
                <span className="text-[8.5px] text-slate-500 mt-0.5 leading-none">kontrak</span>
              </div>
            </div>
            <div className="flex-1 min-w-0 space-y-1.5 text-[11.5px] font-semibold">
              {segs.map((s) => (
                <div key={s.key} className="flex items-center justify-between gap-2">
                  <span className="flex items-center gap-1.5 text-slate-300 min-w-0"><i className="w-2 h-2 rounded-full shrink-0" style={{ background: s.color }} /><span className="truncate">{s.label}</span></span>
                  <span className="text-slate-100 tabular shrink-0">{s.n} <span className="text-slate-500 font-medium">{totalExp ? Math.round((s.n / totalExp) * 100) : 0}%</span></span>
                </div>
              ))}
            </div>
          </div>
          {expiryRows.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center text-center border-t border-slate-800 text-[12px] text-slate-500">
              <CheckCircle2 className="w-6 h-6 mb-1.5" style={{ color: ACCENT }} />
              Tidak ada kontrak yang mendekati masa berakhir.
            </div>
          ) : (
            <BlurScroll className="border-t border-slate-800">
              {expiryRows.map(({ c, days, state }) => (
                <div key={c.id} className="flex items-center gap-2.5 py-2.5 border-b border-slate-800 last:border-0">
                  <div className="min-w-0 flex-1">
                    <p className="text-[12px] font-bold text-slate-100 truncate">{c.title}</p>
                    <p className="text-[10.5px] text-slate-500 truncate">
                      <span className={days <= 7 ? "text-rose-500 font-semibold" : "text-amber-500 font-semibold"}>{days} hari lagi</span>
                      {" · "}{state === "draft" ? "draf perpanjangan ada" : state === "auto" ? "perpanjangan otomatis" : "perlu tindakan"}
                    </p>
                  </div>
                  <button onClick={() => p.onOpenContract(c)} className="px-2.5 py-1 rounded-lg bg-slate-950 hover:bg-slate-800 text-[10.5px] font-bold text-slate-300 cursor-pointer shrink-0">Tinjau</button>
                  <button onClick={() => p.onRenew(c)} className="px-2.5 py-1 rounded-lg text-white text-[10.5px] font-bold cursor-pointer shrink-0 flex items-center gap-1 hover:opacity-90" style={{ background: ACCENT }}>
                    <RefreshCw className="w-3 h-3" /> {state === "draft" ? "Buka" : "Perpanjang"}
                  </button>
                </div>
              ))}
            </BlurScroll>
          )}
        </Panel>

        {/* 6. Produktivitas & log audit */}
        <Panel
          icon={History}
          title="Produktivitas Pengguna & Log Audit Aktivitas"
          className="lg:col-span-5 h-[340px] lg:h-auto"
          right={<button onClick={p.onOpenAudits} className="text-[10.5px] font-bold text-slate-500 hover:text-slate-100 bg-slate-950 px-2.5 py-1 rounded-lg flex items-center gap-0.5 cursor-pointer">Semua <ChevronRight className="w-3 h-3" /></button>}
        >
          <BlurScroll>
            {p.audits.length === 0 && <p className="text-[12px] text-slate-500 text-center py-6">Belum ada aktivitas.</p>}
            {p.audits.slice(0, 20).map((a) => (
              <div key={a.id} className="flex items-start gap-2.5 py-2.5 border-b border-slate-800 last:border-0">
                <span className="w-8 h-8 rounded-full bg-indigo-500/10 text-[#635BFF] flex items-center justify-center text-[11px] font-bold shrink-0">
                  {(a.userName || "?").trim().charAt(0).toUpperCase()}
                </span>
                <div className="min-w-0 flex-1">
                  <div className="flex items-baseline justify-between gap-2">
                    <p className="text-[12px] font-bold text-slate-100 truncate">{a.userName} <span className="text-[10px] font-medium text-slate-500">{p.roleLabel(a.userRole)}</span></p>
                    <span className="text-[10px] text-slate-500 shrink-0">{new Date(a.timestamp).toLocaleDateString("id-ID", { day: "numeric", month: "short" })}</span>
                  </div>
                  <p className="text-[11px] text-slate-500 leading-snug line-clamp-2">
                    {a.action} • {a.details}
                    {a.contractNumber ? <span className="font-mono text-[10px]"> · {a.contractNumber}</span> : null}
                  </p>
                </div>
              </div>
            ))}
          </BlurScroll>
        </Panel>

        {/* 5. Perlu persetujuan saya */}
        <Panel
          icon={Stamp}
          title="Perlu Persetujuan Saya"
          className="lg:col-span-7 h-[340px] lg:h-auto"
          right={p.approvals.length > 0 ? <span className="text-[10.5px] font-bold px-2.5 py-1 rounded-lg bg-rose-500/10 text-rose-500 tabular-nums">{p.approvals.length}</span> : null}
        >
          {p.approvals.length === 0 ? (
            <div className="flex-1 flex flex-col items-center justify-center text-center text-[12px] text-slate-500">
              <CheckCircle2 className="w-6 h-6 mb-1.5" style={{ color: ACCENT }} />
              Tidak ada dokumen yang menunggu persetujuan Anda.
            </div>
          ) : (
            <>
              <div className="grid grid-cols-[1fr_88px_88px_76px] gap-2 px-1 text-[9.5px] font-bold text-slate-500 uppercase tracking-wider shrink-0">
                <span>Nama kontrak</span><span className="text-center">Langkah</span><span className="text-center">Menunggu</span><span className="text-center">Aksi</span>
              </div>
              <BlurScroll className="border-t border-slate-800">
                {p.approvals.map(({ contract: c, step, isMine, asDelegate, totalSteps, waitingSince }) => {
                  const d = hari(waitingSince);
                  const late = p.slaDays > 0 && d >= p.slaDays;
                  return (
                    <div key={c.id} className="grid grid-cols-[1fr_88px_88px_76px] gap-2 items-center py-2.5 px-1 border-b border-slate-800 last:border-0">
                      <div className="min-w-0">
                        <p className="text-[12px] font-bold text-slate-100 truncate">{c.title}</p>
                        <p className="text-[10.5px] text-slate-500 truncate"><span className="font-mono">{c.contractNumber}</span>
                          {asDelegate && !isMine ? ` · pengganti ${step.approverName || ""}` : ""}</p>
                      </div>
                      <span className="text-center text-[11px] text-slate-500 tabular">{step.order}/{totalSteps}</span>
                      <span className={`text-center text-[11px] font-semibold tabular ${late ? "text-rose-500" : "text-slate-500"}`}>{d === 0 ? "hari ini" : `${d} hari`}</span>
                      <span className="text-center">
                        <button onClick={() => p.onOpenContract(c)} className={`px-2.5 py-1 rounded-lg text-[10.5px] font-bold cursor-pointer ${isMine ? "text-white hover:opacity-90" : "bg-slate-950 text-slate-300 hover:bg-slate-800"}`} style={isMine ? { background: ACCENT } : undefined}>
                          {isMine ? "Putuskan" : "Lihat"}
                        </button>
                      </span>
                    </div>
                  );
                })}
              </BlurScroll>
            </>
          )}
        </Panel>
      </div>
    </div>
  );
}
