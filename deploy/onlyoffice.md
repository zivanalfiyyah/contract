# Editor Word (ONLYOFFICE Document Server)

Dokumen **Upload (.docx)** dibuka dengan ONLYOFFICE Document Server: tampilan, tabel, kotak teks,
kolom, header/footer, daftar, dan copy-paste memakai mesin editor Word-compatible, bukan render buatan sendiri.
Bila `ONLYOFFICE_URL` kosong atau Document Server tidak terjangkau, aplikasi otomatis memakai editor lama.

## 1. Jalankan Document Server

```bash
docker run -d --name onlyoffice --restart unless-stopped -p 8080:80 \
  -e JWT_ENABLED=true \
  -e JWT_SECRET='GANTI_DENGAN_SECRET_ACAK_PANJANG' \
  onlyoffice/documentserver
```

Minimal RAM 4 GB. Pasang di belakang HTTPS (nginx/Caddy) bila aplikasi diakses lewat HTTPS,
karena browser memblokir iframe http di halaman https.

## 2. Isi `.env` aplikasi

```
ONLYOFFICE_URL=https://office.example.com          # alamat DS yang dibuka browser
ONLYOFFICE_JWT_SECRET=GANTI_DENGAN_SECRET_ACAK_PANJANG   # sama persis dgn JWT_SECRET di DS
ONLYOFFICE_APP_URL=https://app.example.com         # alamat aplikasi ini yang bisa dijangkau DS
# ONLYOFFICE_INTERNAL_URL=http://onlyoffice:80     # opsional: alamat DS dari sisi server aplikasi
```

Jika aplikasi dan DS berada di jaringan privat/localhost, DS menolak memanggil IP privat secara bawaan.
Tambahkan di container DS: `-e ALLOW_PRIVATE_IP_ADDRESS=true` (hanya untuk jaringan tepercaya).

## 3. Cara kerja & keamanan

- Browser membuka editor lewat `api.js` dari Document Server.
- DS mengambil berkas dari `/api/onlyoffice/file/:id?t=…` dan mengirim hasil edit ke
  `/api/onlyoffice/callback/:id?t=…`. Keduanya memakai token JWT berumur 12 jam yang terikat ke
  satu kontrak, satu versi, dan satu pengguna. Callback juga wajib bertanda tangan JWT DS.
- Hasil edit hanya diunduh dari host Document Server sendiri (anti-SSRF).
- Satu sesi editor = satu versi baru (Versi N+1); simpanan berikutnya di sesi yang sama memperbarui
  versi itu, bukan menumpuk versi. Versi original tidak pernah ditimpa. Hanya kontrak berstatus Draft yang bisa diedit.

## Plugin "Sisipkan Pasal"

Folder `onlyoffice-plugin/` (di akar proyek, sejajar `server.ts`) disajikan server di `/onlyoffice-plugins/...`
dan dimuat BROWSER dari halaman editor. **Folder ini wajib ikut ter-deploy.**

Checklist produksi:
1. `ONLYOFFICE_PLUGIN_BASE_URL="https://domain-aplikasi-anda"` di `.env`. Tanpa ini, alamat plugin ditebak dari
   host permintaan; di belakang proxy tebakan itu bisa `http` sehingga diblokir browser (mixed content).
2. Aplikasi dan Document Server sama-sama `https`.
3. Skrip editor (`plugins.js`) diambil dari Document Server sendiri (`<ONLYOFFICE_URL>/sdkjs-plugins/v1/plugins.js`);
   cadangannya `onlyoffice.github.io`. Pastikan jalur `/sdkjs-plugins/` di Document Server tidak diblokir proxy.
4. Hanya jalur `/onlyoffice-plugins` yang boleh dibingkai & diakses lintas asal (diatur di `server.ts`);
   daftar pasal dilindungi token tautan per kontrak (berlaku 12 jam), bukan cookie.
5. Jangan lupa `JWT_SECRET` Document Server yang kuat dan sama dengan `ONLYOFFICE_JWT_SECRET`.

Di editor: tab **Plugin** -> **Sisipkan Pasal** -> cari pasal -> **Sisipkan**.
