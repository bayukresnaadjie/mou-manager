# Sistem Manajemen MoU & Perjanjian Kerja Sama

Aplikasi web berbasis **Google Apps Script** dan **Google Sheets** untuk mengelola
MoU dan perjanjian kerja sama, dengan pengingat otomatis sebelum kontrak berakhir.

## Fitur

- **Admin:** melihat, menambah, mengubah, dan menghapus dokumen; mengelola akun mitra; melihat riwayat notifikasi.
- **Mitra:** hanya melihat dokumen milik perusahaannya sendiri (tanpa kontak dan catatan internal).
- **Pengingat otomatis** lewat email dan WhatsApp ke kedua pihak pada H-90, H-30, H-14, H-7, H-1, dan hari berakhir.
- Login dengan email dan password (hash SHA-256 + salt), pembatasan percobaan login, sesi 6 jam.

## Struktur

| File | Fungsi |
|---|---|
| `Code.gs` | Backend: autentikasi, CRUD, notifikasi, trigger |
| `Index.html` | Tampilan web app |
| `appsscript.json` | Manifest (zona waktu, akses web app) |

## Cara pasang

1. Buat Google Sheet baru, lalu buka **Extensions → Apps Script**.
2. Salin `Code.gs`, buat file HTML bernama `Index` lalu salin `Index.html`, dan salin `appsscript.json`
   (aktifkan "Show appsscript.json manifest file" di Project Settings).
3. Di `Code.gs`, ubah `ADMIN_EMAIL` dan `NAMA_ORGANISASI`. Biarkan `ADMIN_PASSWORD` kosong.
4. Jalankan fungsi `setup` dan setujui izin akses. Password awal admin tampil di **Execution log**.
5. **Deploy → New deployment → Web app**, dengan *Execute as: Me* dan *Who has access: Anyone*.
6. Buka URL `/exec`, login sebagai admin, lalu segera ganti password lewat tombol **Ganti password**.

## Mengaktifkan WhatsApp (opsional)

Pengiriman WhatsApp memakai [Fonnte](https://fonnte.com).

1. Daftar di Fonnte, hubungkan satu nomor pengirim (scan QR), lalu salin token device.
2. Di Apps Script: **Project Settings → Script Properties → tambah `FONNTE_TOKEN`** berisi token tersebut.

Token disimpan di Script Properties, **bukan di kode**. Tanpa token, pengiriman WhatsApp dilewati
(email tetap terkirim). Untuk penyedia lain, ganti isi fungsi `sendWhatsApp_`.

## Memperbarui setelah mengubah kode

Gunakan **Deploy → Manage deployments → ikon pensil → New version → Deploy** agar URL `/exec` tetap sama.

## Catatan keamanan

- Jangan menaruh password, token, atau data pribadi di kode yang diunggah ke repository.
- Nama mitra pada akun harus sama persis dengan kolom **Pihak kedua** pada dokumen.
- Aplikasi dijalankan lewat URL `/exec` milik Apps Script. GitHub hanya menyimpan kode sumber;
  `Index.html` tidak berfungsi jika dibuka langsung atau lewat GitHub Pages karena memakai `google.script.run`.
