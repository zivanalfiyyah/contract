// Uji RBAC (hapus Admin + lintas PT). Jalankan: npx tsx scripts/rbac-test.ts
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { ROLE_DEFS, PERMISSIONS, canAccessContract, canViewContract, isDirectlyAssigned, isCrossTenantRole, isFullAccessRole, isAdminLike, isPermAllowed, DEFAULT_MENUS_BY_ROLE, canBeApprover } from "../src/roles";

const __dirname = path.dirname(new URL(import.meta.url).pathname);
let n = 0;
const t = (name: string, fn: () => void) => { fn(); n++; console.log("ok -", name); };

const ext = (tenantId: string, extra: any = {}) => ({ tenantId, party2Type: "Perusahaan", ...extra });
const emp = (tenantId: string, extra: any = {}) => ({ tenantId, party2Type: "Karyawan", ...extra });
const u = (role: string, tenantId: string, id = role + "-" + tenantId) => ({ id, role, tenantId });
const ROLES = ROLE_DEFS.map((r) => r.value);

t("role admin hilang total", () => {
  assert.ok(!ROLES.includes("admin" as any));
  for (const p of PERMISSIONS) assert.ok(!p.defaults.includes("admin"), p.key);
  for (const k of Object.keys(DEFAULT_MENUS_BY_ROLE)) assert.notEqual(k, "admin");
  assert.equal(isAdminLike("admin"), false);
  assert.equal(isFullAccessRole("admin"), false);
  assert.equal(isPermAllowed("admin", "settings.manage", {} as any), false);
});

t("tidak ada referensi role admin tersisa di kode", () => {
  const root = path.resolve(__dirname, "..");
  const files = ["src/roles.ts", "src/App.tsx", "server.ts", "auth.ts", "db.ts", "reminders.ts", "template-docx.ts", "party-fields.ts", "dcs/routes.ts", "dcs/repo.ts"];
  for (const f of files) {
    const src = fs.readFileSync(path.join(root, f), "utf8");
    const hits = src.split("\n").filter((l) => /(role|Role)\s*[=!]==?\s*"admin"|\["admin"|"admin",|, "admin"\]|role: "admin"/.test(l) && !/previousRole|migrasi|Migrasi|=== "admin"\) \{ \/\/ migrasi/.test(l));
    // satu-satunya pengecualian sah: migrasi akun lama di server.ts
    const bad = hits.filter((l) => !(f === "server.ts" && /u\.role === "admin"|role === "admin"/.test(l)));
    assert.deepEqual(bad, [], f);
  }
});

t("Super Admin: semua PT, internal & eksternal", () => {
  for (const tid of ["TBK", "GKM", "X"]) {
    assert.ok(canAccessContract(u("super_admin", "TBK"), ext(tid)));
    assert.ok(canAccessContract(u("super_admin", "TBK"), emp(tid)));
  }
});

t("Legal: semua PT tapi HANYA eksternal", () => {
  const l = u("legal", "TBK");
  assert.ok(canAccessContract(l, ext("TBK")));
  assert.ok(canAccessContract(l, ext("GKM")));
  assert.ok(!canAccessContract(l, emp("TBK")));
  assert.ok(!canAccessContract(l, emp("GKM")));
});

t("HR/GA: hanya internal di PT sendiri", () => {
  const h = u("hr_ga", "TBK");
  assert.ok(canAccessContract(h, emp("TBK")));
  assert.ok(!canAccessContract(h, ext("TBK")));
  assert.ok(!canAccessContract(h, emp("GKM")));
  assert.ok(!canAccessContract(h, ext("GKM")));
});

t("peran lain: hanya yang melibatkan dirinya, tidak lintas PT otomatis", () => {
  for (const role of ROLES.filter((r) => !["super_admin", "legal", "hr_ga"].includes(r))) {
    const x = u(role, "TBK");
    assert.ok(!canAccessContract(x, ext("TBK")), role + " tak boleh lihat semua dokumen PT-nya");
    assert.ok(!canAccessContract(x, ext("GKM")), role + " lintas PT otomatis");
    assert.ok(canAccessContract(x, ext("TBK", { createdById: x.id })), role + " pembuat");
  }
});

t("approver lintas PT: hanya dokumen yang ditugaskan langsung", () => {
  const dir = u("direktur", "TBK", "dir1");
  const assigned = ext("GKM", { approvalSteps: [{ approverId: "dir1" }] });
  const other = ext("GKM", { approvalSteps: [{ approverId: "lain" }] });
  assert.ok(canAccessContract(dir, assigned));
  assert.ok(!canAccessContract(dir, other));
  assert.ok(!canAccessContract(dir, ext("GKM")));
  // riwayat ronde lama tetap boleh
  assert.ok(canAccessContract(dir, ext("GKM", { approvalHistory: [{ steps: [{ approverId: "dir1" }] }] })));
  // delegasi
  assert.ok(canAccessContract({ ...u("finance", "TBK", "del"), actingForIds: ["dir1"] } as any, assigned));
  // HR yang ditugaskan di dokumen internal PT lain hanya itu saja
  const hr = u("hr_ga", "TBK", "hr1");
  assert.ok(canAccessContract(hr, emp("GKM", { approvalSteps: [{ approverId: "hr1" }] })));
  assert.ok(!canAccessContract(hr, emp("GKM")));
});

t("isCrossTenantRole & isFullAccessRole", () => {
  assert.ok(isCrossTenantRole("super_admin") && isCrossTenantRole("legal"));
  for (const r of ROLES.filter((r) => !["super_admin", "legal"].includes(r))) assert.ok(!isCrossTenantRole(r), r);
  assert.ok(isFullAccessRole("legal") && isFullAccessRole("super_admin") && !isFullAccessRole("hr_ga"));
});

t("viewer/staff tetap tidak bisa jadi approver", () => {
  assert.ok(!canBeApprover("viewer") && !canBeApprover("staff") && canBeApprover("direktur"));
});

t("server: validasi 'satu PT' dihapus & ada konteks tenant per dokumen", () => {
  const src = fs.readFileSync(path.resolve(__dirname, "../server.ts"), "utf8");
  assert.ok(src.includes("app.param(\"id\""));
  assert.ok(src.includes("/api/users/approver-candidates"));
  assert.ok(!/id === id && u\.tenantId === tid && u\.active/.test(src));
  const auth = fs.readFileSync(path.resolve(__dirname, "../auth.ts"), "utf8");
  assert.ok(auth.includes("tenantCtx"));
});

console.log(`\n${n} kelompok uji lolos`);
