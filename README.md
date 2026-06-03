# Aerizen Asset Management v3.8 - Recovery Config Fix

Perbaikan penting:
- Memperbaiki error `ConfigurationView is not defined` yang membuat aplikasi masuk Recovery Mode.
- Mengembalikan semua halaman enterprise yang sempat hilang pada update tampilan Asset ID Card: Kontrak, Operasional, Maintenance, Finance, Sparepart, Vendor, Ticket, Dokumen, Approval, Laporan, Akun, dan Konfigurasi.
- Tampilan foto kendaraan model Asset ID Card/BPKB tetap dipertahankan.
- Fitur installer update tanpa uninstall tetap aman.
- Data lokal dan Supabase tidak perlu dihapus untuk fix ini.

# Aerizen Asset Management

Versi ini adalah **v3.6 Installer + Foto Kendaraan HD**. Build GitHub Actions default sekarang menghasilkan file **Setup installer**, bukan portable, agar update tidak perlu uninstall dan tidak membingungkan. Foto aset kendaraan juga ditampilkan dengan mode HD/contain agar tidak terpotong atau terlihat jelek.
 v3.1 Enterprise Workflow

Aerizen Asset Management adalah aplikasi Electron + React + Vite untuk manajemen aset kendaraan rental, aset operasional, dan perangkat IT. Aplikasi ini dibuat **offline-first** dan dapat memakai **Supabase realtime** untuk sinkronisasi antar akun/perangkat.

## Fitur utama v3.1

- Dashboard enterprise command center
- Data Aset dengan tambah/edit/hapus, upload foto, import Excel kendaraan sewa, QR/Barcode, dan Aset 360
- Reminder otomatis STNK merah jika expired atau tersisa maksimal 30 hari
- Reminder KIR, pajak kendaraan, asuransi, service berkala, kontrak, invoice, ticket, dan approval
- Kontrak sewa: customer, aset, nilai bulanan, deposit, periode, dokumen, dan renewal
- Check-in / Check-out aset: KM, BBM, kondisi, aksesoris, penerima, dan tanda tangan digital
- Digital Inspection Checklist kendaraan dan perangkat IT
- Maintenance Work Order dengan Nominal Biaya Aktual, vendor, sparepart, foto before/after, history per nopol, export Excel, dan tombol Selesai cepat
- Finance & Invoice: billing bulanan, unpaid/partial/paid, profit per aset, export Excel
- Sparepart & Inventory: stok minimum, nilai persediaan, lokasi, vendor
- Vendor Management: bengkel/supplier, rating, rata-rata biaya, lead time, transaksi
- Ticket & SLA: keluhan customer/internal, prioritas, SLA, status, PIC, close ticket
- Digital Document Vault: STNK, BPKB, KIR, asuransi, garansi IT, kontrak, invoice service, BAST
- Lonceng notifikasi kecil di pojok atas untuk semua reminder otomatis, terpisah dari menu utama
- Role & Permission User, Approval Workflow, dan Audit Trail
- Akun: login/logout, lupa password, lihat password, reset password via Supabase Auth
- Konfigurasi Supabase: URL, anon key, workspace, realtime listener, auto sync, push/pull queue
- Backup/restore JSON dan export Excel per modul

## Cara menjalankan lokal

```bash
npm install
npm run electron:dev
```

## Cara build EXE lokal

```bash
npm run dist
```

Hasil build ada di folder `release`.

## Cara build dari GitHub Actions

1. Upload semua isi project ke repository GitHub.
2. Buka tab **Actions**.
3. Pilih **Build Aerizen Windows EXE**.
4. Klik **Run workflow**.
5. Download artifact `aerizen-windows-exe`.

## Supabase setup

Jalankan SQL di `supabase/schema.sql` pada SQL Editor Supabase.

Aplikasi memakai satu tabel fleksibel:

```sql
create table public.aerizen_records (
  id text primary key,
  collection text not null,
  payload jsonb not null,
  updated_at timestamptz default now(),
  deleted_at timestamptz
);
```

Setelah itu isi **Supabase URL**, **anon key**, dan **workspace** di menu Konfigurasi.

## Catatan

- Aplikasi tetap bisa dipakai tanpa Supabase karena mode offline selalu aktif.
- Supabase diperlukan untuk realtime multi akun, login akun asli, dan lupa password.
- Untuk reset password, aktifkan email recovery di Supabase Authentication.

## Update aplikasi tanpa uninstall

Mulai versi 3.3, jalankan installer `.exe` versi terbaru untuk menimpa versi lama. Tidak perlu uninstall.

Gunakan file installer NSIS dari folder `release`, bukan file `portable`.

```bash
npm run dist:installer
```

Pastikan `appId` dan `productName` tidak diganti agar data lokal/offline dan instalasi lama tetap dikenali sebagai aplikasi yang sama.


### Catatan v3.4 blank screen fix
Jika EXE sebelumnya hanya menampilkan layar gelap/kosong, upload versi v3.4 ini ke GitHub, jalankan Actions lagi, lalu install installer NSIS terbaru di atas versi lama. Penyebab utamanya adalah asset Vite yang memakai path absolut saat dibuka lewat `file://` di Electron packaged.
