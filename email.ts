import nodemailer, { Transporter } from "nodemailer";
import { logger } from "./logger.js";

// Transactional email via SMTP — just needs credentials in .env, so this is
// fully functional, not scaffolding. Until those env vars are set,
// isConfigured() is false and callers fall back to manual mailto:/wa.me
// links — the app must never claim an email was sent when it wasn't.

export interface EmailConfig {
  host: string;
  port: number;
  secure: boolean;
  user: string;
  pass: string;
  fromEmail: string;
  fromName: string;
}

function readConfig(): Partial<EmailConfig> {
  return {
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT) || 587,
    secure: process.env.SMTP_SECURE === "true",
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
    fromEmail: process.env.SMTP_FROM_EMAIL || process.env.SMTP_USER,
    fromName: process.env.SMTP_FROM_NAME || "Smart CLM Enterprise",
  } as Partial<EmailConfig>;
}

export function isEmailConfigured(): boolean {
  const cfg = readConfig();
  return !!(cfg.host && cfg.user && cfg.pass && cfg.fromEmail);
}

let cachedTransporter: Transporter | null = null;
let cachedConfigKey = "";

function getTransporter(): Transporter {
  const cfg = readConfig();
  const key = `${cfg.host}:${cfg.port}:${cfg.user}`;
  // Re-create the transporter if env vars changed since last build (e.g. an
  // admin just updated SMTP settings and hit "Test Connection" again).
  if (!cachedTransporter || cachedConfigKey !== key) {
    cachedTransporter = nodemailer.createTransport({
      host: cfg.host,
      port: cfg.port,
      secure: cfg.secure, // true for port 465, false for 587/25 (STARTTLS)
      auth: { user: cfg.user, pass: cfg.pass },
    });
    cachedConfigKey = key;
  }
  return cachedTransporter;
}

export interface SendEmailParams {
  to: string;
  subject: string;
  html: string;
  text?: string;
  // Lampiran opsional (mis. PDF kontrak) — diteruskan apa adanya ke
  // nodemailer, yang mendukung Buffer content secara native.
  attachments?: { filename: string; content: Buffer; contentType?: string }[];
}

export interface SendEmailResult {
  sent: boolean;
  error?: string;
}

export async function sendEmail(params: SendEmailParams): Promise<SendEmailResult> {
  if (!isEmailConfigured()) {
    return { sent: false, error: "SMTP belum dikonfigurasi." };
  }
  const cfg = readConfig();
  try {
    await getTransporter().sendMail({
      from: `"${cfg.fromName}" <${cfg.fromEmail}>`,
      to: params.to,
      subject: params.subject,
      html: params.html,
      text: params.text,
      attachments: params.attachments,
    });
    return { sent: true };
  } catch (err: any) {
    logger.error({ err, to: params.to }, "Failed to send email via SMTP");
    return { sent: false, error: err?.message || "Gagal mengirim email." };
  }
}

export interface TestConnectionResult {
  connected: boolean;
  message: string;
}

export async function testEmailConnection(): Promise<TestConnectionResult> {
  if (!isEmailConfigured()) {
    return { connected: false, message: "SMTP_HOST, SMTP_USER, SMTP_PASS, dan SMTP_FROM_EMAIL belum lengkap di .env." };
  }
  try {
    await getTransporter().verify();
    return { connected: true, message: "Berhasil terhubung ke server SMTP." };
  } catch (err: any) {
    return { connected: false, message: `Gagal terhubung ke SMTP: ${err?.message || "unknown error"}` };
  }
}

// Teks dari pengguna (nama, keterangan) WAJIB di-escape sebelum disisipkan ke
// bodyHtml — tanpa ini nama pendaftar seperti "<a href=...>" ikut dirender
// sebagai tautan di email admin.
export function escapeEmailHtml(s: string): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

// Shared HTML wrapper so every transactional email looks consistent.
// Palet mengikuti tema aplikasi (putih + biru + aksen peach, lihat
// src/index.css) supaya email tidak terasa datang dari sistem lain.
export function emailTemplate(opts: { title: string; bodyHtml: string; ctaLabel?: string; ctaUrl?: string }): string {
  return `<!doctype html>
<html><body style="margin:0;padding:0;background:#eef5fd;font-family:'Plus Jakarta Sans',system-ui,-apple-system,Segoe UI,Roboto,sans-serif;">
<div style="max-width:560px;margin:0 auto;padding:32px 20px;">
  <div style="font-size:13px;font-weight:700;color:#2a6293;margin:0 0 14px 4px;">Smart CLM Enterprise</div>
  <div style="background:#ffffff;border-radius:22px;padding:30px;box-shadow:0 12px 34px -10px rgba(110,145,200,0.28);">
    <div style="height:4px;width:56px;border-radius:4px;background:linear-gradient(90deg,#72b6e6,#f9cfa8);margin:0 0 18px;"></div>
    <h1 style="color:#1c2433;font-size:18px;margin:0 0 14px;">${opts.title}</h1>
    <div style="color:#475469;font-size:14px;line-height:1.65;">${opts.bodyHtml}</div>
    ${opts.ctaUrl ? `<div style="margin-top:24px;"><a href="${opts.ctaUrl}" style="display:inline-block;background:#3b7fb6;background-image:linear-gradient(135deg,#3b7fb6,#4f98d0);color:#ffffff;text-decoration:none;padding:12px 26px;border-radius:999px;font-weight:700;font-size:14px;">${opts.ctaLabel || "Buka Tautan"}</a></div>
    <p style="color:#86929f;font-size:11px;line-height:1.5;margin:18px 0 0;">Tombol tidak berfungsi? Salin tautan ini ke browser:<br><span style="color:#3a86c2;word-break:break-all;">${opts.ctaUrl}</span></p>` : ""}
  </div>
  <p style="color:#86929f;font-size:11px;text-align:center;margin-top:20px;">Email otomatis dari Smart Contract Lifecycle Management System.</p>
</div>
</body></html>`;
}
