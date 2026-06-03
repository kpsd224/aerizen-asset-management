# Changelog

## v3.6.0 - Installer resmi + Foto Kendaraan HD
- GitHub Actions default dibuat installer-only.
- Nama output installer menjadi `Aerizen-Asset-Management-Setup-3.6.0-x64.exe`.
- Menghindari salah download/menjalankan portable EXE.
- Foto aset/kendaraan sekarang memakai `object-fit: contain`, bukan `cover`.
- Tinggi preview foto Aset 360 diperbesar.
- Foto asli tetap disimpan tanpa kompresi.


## v3.5.0 - GitHub Actions Clean Warning

- Update workflow GitHub Actions agar memakai action yang kompatibel dengan runtime Node 24.
- Ganti `actions/checkout@v4` menjadi `actions/checkout@v5`.
- Ganti `actions/setup-node@v4` menjadi `actions/setup-node@v6`.
- Ganti `actions/upload-artifact@v4` menjadi `actions/upload-artifact@v6`.
- Pin runner ke `windows-2022` agar tidak muncul notice migrasi `windows-latest`.
- Build project memakai Node.js 22.
- Fitur aplikasi v3.4 tetap dipertahankan.


## v3.3.0 - Update Installer Tanpa Uninstall

- Installer Windows dibuat update-friendly agar versi baru bisa langsung dipasang di atas versi lama.
- `appId` dan `productName` dikunci tetap sama supaya Windows mengenali aplikasi sebagai aplikasi yang sama.
- `allowToChangeInstallationDirectory` diubah menjadi `false` agar update tidak membuat instalasi ganda di folder berbeda.
- `deleteAppDataOnUninstall` dibuat `false` supaya data lokal/offline tidak terhapus saat uninstall manual.
- Menambahkan script `npm run dist:installer` untuk build installer update tanpa portable.
- Menambahkan panduan `UPGRADE-INSTALL-GUIDE.md`.
- Fitur v3.2 tetap aman: notification bell, edit aset, Aset 360, maintenance, finance, kontrak, approval, Supabase sync, lupa password, dan lihat password.

# Aerizen Asset Management Changelog

## v3.2.0 - GitHub Actions Build Fix
- Memperbaiki error `Dependencies lock file is not found` di step Setup Node.js.
- Menghapus opsi `cache: npm` dari `actions/setup-node` supaya workflow tidak membutuhkan `package-lock.json` sebelum `npm install`.
- Mengubah instalasi dependency menjadi `npm install --no-audit --no-fund`.
- Fitur v3.1 tetap dipertahankan: notification bell, enterprise workflow, offline + Supabase, maintenance history, export Excel, akun, lupa password, dan lihat password.


## v3.1.0

- Menu **Notifikasi** dihapus dari daftar navigasi utama agar tidak memenuhi menu.
- Ditambahkan ikon **lonceng notifikasi** kecil di pojok atas kanan header.
- Lonceng menampilkan badge jumlah notifikasi merah/kritis.
- Klik lonceng membuka panel daftar reminder STNK, KIR, pajak, asuransi, service, kontrak, invoice, ticket, dan approval.
- Klik item notifikasi langsung mengarah ke modul terkait.
- Dashboard tetap bersih dan hanya menampilkan ringkasan alert utama.


## v3.0.0 - Enterprise Workflow

- Menambahkan modul Kontrak Sewa.
- Menambahkan modul Check-in / Check-out aset.
- Menambahkan Digital Inspection Checklist kendaraan dan perangkat IT.
- Menambahkan Finance & Invoice.
- Menambahkan profit per aset.
- Menambahkan Sparepart & Inventory.
- Menambahkan Vendor / Bengkel / Supplier Management.
- Menambahkan Ticket & SLA.
- Menambahkan Digital Document Vault.
- Menambahkan Notification Center.
- Menambahkan Role & Permission User.
- Menambahkan Approval Workflow terpisah.
- Menambahkan reminder KIR, pajak, asuransi, service, kontrak, invoice, ticket, dan approval.
- Menambahkan QR/Barcode field per aset.
- Maintenance tetap memakai Nominal Biaya Aktual, bukan estimasi.
- History maintenance tetap bisa difilter per nopol dan diekspor ke Excel.
- Tombol Selesai maintenance cepat tetap tersedia tanpa membuka form Work Order.
- Fitur akun tetap punya login/logout, lupa password, dan lihat password.
- Fitur offline-first dan Supabase realtime tetap dipertahankan.
- Dashboard dibuat sebagai command center dan tidak lagi berisi shortcut/work order double.

## v2.8.0

- Tombol Selesai maintenance langsung dari tabel.

## v2.7.0

- Dashboard dibersihkan dari pintasan kerja double.


## v3.4.0 - Blank Screen Fix
- Memperbaiki layar blank pada EXE packaged dengan `base: './'` di Vite.
- Menambahkan loading fallback di `index.html`.
- Menambahkan recovery/error boundary agar error render tidak menjadi layar kosong.
- Menambahkan migrasi data lokal agar update dari v2/v3 lama tetap aman.
