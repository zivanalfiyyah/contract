import jwt from "jsonwebtoken";

// ---------------------------------------------------------------------------
// Integrasi ONLYOFFICE Document Server (editor .docx persis seperti Word).
//
//   Browser  ──(api.js, iframe)──▶  Document Server (DS)
//   DS       ──GET  /api/onlyoffice/file/:id?t=…──▶  server ini   (mengambil berkas)
//   DS       ──POST /api/onlyoffice/callback/:id?t=…▶ server ini  (mengirim hasil edit)
//
// DS tidak punya cookie login pengguna, jadi dua URL di atas dilindungi token
// berumur pendek (JWT) yang kita tanda tangani sendiri. Permintaan DS ke
// callback juga wajib bertanda tangan JWT DS (ONLYOFFICE_JWT_SECRET).
//
// Variabel lingkungan:
//   ONLYOFFICE_URL         alamat DS yang bisa dibuka BROWSER, mis. https://office.example.com
//   ONLYOFFICE_JWT_SECRET  secret yang sama dengan JWT_SECRET di DS (WAJIB)
//   ONLYOFFICE_APP_URL     alamat aplikasi ini yang bisa dijangkau DS (bawaan: APP_URL)
//   ONLYOFFICE_INTERNAL_URL (opsional) alamat DS dari sisi server ini, bila beda dgn ONLYOFFICE_URL
// ---------------------------------------------------------------------------

export interface OnlyOfficeSettings {
  enabled: boolean;
  publicUrl: string;
  internalUrl: string;
  appUrl: string;
  secret: string;
}

const trimSlash = (s: string) => s.replace(/\/+$/, "");

export function onlyOfficeSettings(env: NodeJS.ProcessEnv = process.env): OnlyOfficeSettings {
  const publicUrl = trimSlash(env.ONLYOFFICE_URL || "");
  const secret = env.ONLYOFFICE_JWT_SECRET || "";
  const appRaw = env.ONLYOFFICE_APP_URL || (env.APP_URL && env.APP_URL !== "MY_APP_URL" ? env.APP_URL : "");
  return {
    enabled: !!publicUrl && !!secret && !!appRaw,
    publicUrl,
    internalUrl: trimSlash(env.ONLYOFFICE_INTERNAL_URL || publicUrl),
    appUrl: trimSlash(appRaw),
    secret,
  };
}

// ---- Token tautan (file & callback) ---------------------------------------
export interface LinkClaims {
  kind: "file" | "callback" | "plugin";
  contractId: string;
  /** Nomor versi berkas yang dibuka. */
  version: number;
  userId: string;
  tenantId: string;
  /** Id sesi edit (satu sesi editor = satu versi baru, simpanan berikutnya menimpa versi itu). */
  session: string;
}

export function signLink(s: OnlyOfficeSettings, claims: LinkClaims, ttlSeconds: number): string {
  return jwt.sign(claims, s.secret, { expiresIn: ttlSeconds, audience: "onlyoffice-link" });
}

export function verifyLink(s: OnlyOfficeSettings, token: unknown, kind: LinkClaims["kind"], contractId: string): LinkClaims | null {
  if (typeof token !== "string" || !token) return null;
  try {
    const c = jwt.verify(token, s.secret, { audience: "onlyoffice-link", algorithms: ["HS256"] }) as LinkClaims;
    return c.kind === kind && c.contractId === contractId ? c : null;
  } catch {
    return null;
  }
}

// ---- JWT DS (config editor & callback) ------------------------------------
export function signDsPayload(s: OnlyOfficeSettings, payload: object): string {
  return jwt.sign(payload, s.secret, { algorithm: "HS256" });
}

/** Ambil & verifikasi token dari DS: header Authorization: Bearer … atau body.token. */
export function verifyDsRequest(s: OnlyOfficeSettings, headers: Record<string, any>, body: any): any | null {
  const auth = String(headers?.authorization || "");
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : (typeof body?.token === "string" ? body.token : "");
  if (!token) return null;
  try {
    const decoded = jwt.verify(token, s.secret, { algorithms: ["HS256"] }) as any;
    // DS membungkus isi di "payload" bila token dikirim lewat header.
    return decoded?.payload && typeof decoded.payload === "object" ? decoded.payload : decoded;
  } catch {
    return null;
  }
}

/** Hanya boleh mengunduh hasil edit dari host DS sendiri (cegah SSRF lewat body callback). */
export function isAllowedDsDownloadUrl(s: OnlyOfficeSettings, raw: unknown): boolean {
  if (typeof raw !== "string") return false;
  try {
    const u = new URL(raw);
    if (u.protocol !== "http:" && u.protocol !== "https:") return false;
    const allowed = [s.publicUrl, s.internalUrl].filter(Boolean).map((x) => new URL(x).host);
    return allowed.includes(u.host);
  } catch {
    return false;
  }
}

/** Alamat unduh dari DS bisa memakai host publik; ganti ke alamat internal bila diset. */
export function rewriteDsUrlToInternal(s: OnlyOfficeSettings, raw: string): string {
  if (!s.internalUrl || s.internalUrl === s.publicUrl) return raw;
  const u = new URL(raw);
  const pub = new URL(s.publicUrl);
  if (u.host !== pub.host) return raw;
  const internal = new URL(s.internalUrl);
  u.protocol = internal.protocol;
  u.host = internal.host;
  return u.toString();
}

// ---- Konfigurasi editor ----------------------------------------------------
export interface EditorConfigInput {
  contractId: string;
  fileName: string;
  version: number;
  sha256: string;
  canEdit: boolean;
  /** Dokumen yang dikirim ke editor berbeda per kelompok pengguna (mis. kunci pasal) -> kunci DS dibedakan agar cache tidak tertukar. */
  keyVariant?: string;
  user: { id: string; name: string };
  fileUrl: string;
  callbackUrl: string;
  lang?: string;
  /** Plugin "Sisipkan Pasal": URL config.json (harus terjangkau BROWSER) + opsi untuk plugin. */
  plugin?: { configUrl: string; guid: string; options: Record<string, unknown> };
}

/** Kunci dokumen DS: berubah setiap isi berkas berubah (DS menyimpan cache per key). */
export function documentKey(contractId: string, version: number, sha256: string, variant = ""): string {
  return `${contractId}-v${version}-${sha256.slice(0, 12)}${variant ? "-" + variant : ""}`.replace(/[^0-9a-zA-Z._=-]/g, "_").slice(0, 128);
}

export function buildEditorConfig(s: OnlyOfficeSettings, i: EditorConfigInput) {
  const config: any = {
    documentType: "word",
    document: {
      fileType: "docx",
      key: documentKey(i.contractId, i.version, i.sha256, i.keyVariant),
      title: i.fileName,
      url: i.fileUrl,
      permissions: {
        edit: i.canEdit,
        download: true,
        print: true,
        review: i.canEdit,
        comment: i.canEdit,
        copy: true,
        chat: false,
      },
    },
    editorConfig: {
      mode: i.canEdit ? "edit" : "view",
      lang: i.lang || "id",
      callbackUrl: i.callbackUrl,
      user: { id: i.user.id, name: i.user.name },
      customization: {
        autosave: true,
        forcesave: true,
        compactHeader: true,
        hideRightMenu: false,
        feedback: false,
        help: false,
        goback: false,
      },
    },
    width: "100%",
    height: "100%",
  };
  if (i.plugin && i.canEdit) {
    config.editorConfig.plugins = { pluginsData: [i.plugin.configUrl], options: { [i.plugin.guid]: i.plugin.options } };
  }
  config.token = signDsPayload(s, config);
  return config;
}
