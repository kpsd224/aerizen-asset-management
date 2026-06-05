# Aerizen v5.5 UI Polish

- Merapikan layout Ringkasan Aset pada Aset 360 agar status, nopol, GPS, dan KM tidak terpotong.
- Menambah responsive wrapping dan ukuran teks adaptif pada kartu ringkasan.
- Mempertahankan seluruh fitur v3.8-v5.4.

## v5.5.0 — Complete Invoice & Aerizen Branding
- Melengkapi form Invoice dan menghubungkannya dengan Kontrak, Customer, Aset, serta Nopol/Serial.
- Menambahkan rincian biaya, diskon, PPN, PPh, pembayaran, sisa piutang, faktur pajak, dan lampiran invoice.
- Menambahkan perhitungan otomatis subtotal, pajak, total invoice, status pembayaran, dan sisa piutang.
- Menambahkan edit invoice serta informasi invoice yang lebih lengkap pada tabel Finance.
- Menggunakan logo Aerizen pada UI, favicon, splash, dan build Windows.
- Menampilkan versi v5.5 pada identitas dan judul aplikasi.
- Mempertahankan seluruh fitur v3.8-v5.3.

## v5.3.0
- Menambahkan pencarian dan filter pada daftar aset di form Dokumen.
- Menghapus banner teknis dan teks restore yang tidak diperlukan dari UI.
- Mempertahankan seluruh fitur v5.2 dan versi sebelumnya.

# Changelog

## v5.2.0 — Contract & Document Link Fix

- Form kontrak rental dilengkapi secara menyeluruh.
- ID kontrak berurutan dengan prefix `CTR`.
- ID dokumen dibuat manual.
- Pencarian/pemilihan nopol, serial, dan ID aset ditambahkan pada form dokumen.
- Dokumen otomatis terhubung ke Aset 360 melalui aset dan nopol.
- Tabel Kontrak dan Dokumen menampilkan informasi relasi yang lebih lengkap.
- Seluruh modul dan fitur v3.8-v5.1 dipertahankan.

## v5.1.0 — Form, Code & Notification Fix
- Menyamakan ID checklist dengan format `CHK-IN`, `CHK-OUT`, dan `CHK-INS`.
- Menyamakan ID vendor menjadi urutan `VDR-001`, `VDR-002`, dan seterusnya.
- Membuat ID sparepart otomatis berdasarkan kode nama item, seperti `SP-OLI-002`.
- Melengkapi form Work Order, Sparepart, dan Vendor.
- Mengembalikan lonceng notifikasi ke area navigasi agar selalu terlihat.
- Mempertahankan filter dan Export Excel Work Order per nopol.
- Tidak mengurangi modul atau fitur v3.8.

## v4.9.0 — Context Label Fix
- Menghapus istilah tambah-data yang tidak sesuai konteks dari antarmuka.
- Tombol dan judul form sekarang mengikuti modul: Kontrak, Checklist, Invoice, Sparepart, Vendor, Ticket, dan Approval.
- ID otomatis tetap memakai prefix modul: CTR, CHK, INV, SP, VDR, TCK, dan APR.
- Export Excel tetap tersedia per modul tanpa mengurangi fitur sebelumnya.

# Aerizen Asset Management v4.4 - Professional Full Restore Lock

## Fokus perbaikan
Versi ini dibuat dari sistem lengkap sebelumnya dan dikunci agar fitur v3.8 tidak hilang saat ada improvement UI/UX.

## Fitur yang wajib dipertahankan
- Dashboard enterprise
- Data Aset lengkap
- Import Excel kendaraan sewa + template XLSX/CSV
- Kontrak Sewa dengan nominal nilai bulanan dan deposit
- Operasional Check-in / Check-out
- Work Order (WO), bukan diganti Maintenance
- Nominal Biaya Aktual WO
- Finance & Invoice dengan Amount, Paid, Sisa, Due Date, Status, dan tombol Tandai Lunas
- Profitability per Aset
- Sparepart / Inventory dengan nilai stok
- Vendor Management dengan avg cost, lead time, rating, transaksi
- Ticket & Complaint
- Dokumen Aset
- Approval & Permission
- Laporan + export per collection
- Akun
- Konfigurasi Supabase/offline sync

## Improvement v4.4
- Tampilan Aset dibuat seperti kartu BPKB/contract card.
- ID Aset otomatis dan read-only.
- ID Dokumen otomatis dan read-only.
- Upload Dokumen Aset ditambahkan di form Aset dan modul Dokumen.
- GPS punya checkbox sendiri.
- KM/Odometer punya field sendiri.
- Import Excel Aset mendukung GPS, ID GPS, Provider GPS, dan KM/Odometer.
- Tombol Tambah data sesuai modul ditampilkan jelas pada Kontrak, Operasional, Finance, Inventory, Vendor, Ticket, dan Approval.
- Export Excel tetap tersedia per modul.
- Smoke-test diperketat untuk mencegah Akun, Konfigurasi, WO, Export Excel, Import Excel, dan Nominal Biaya Aktual hilang.

## Catatan penting
Jangan pakai versi v4.0-v4.3 yang sebelumnya untuk produksi. Pakai v4.4 ini sebagai basis lanjut.


## v4.7 Approval Workflow Vertical Fix
- Panel Approval Workflow dipindahkan benar-benar ke bawah Role & Permission User, bukan berada di samping.
- Preview HTML juga dikunci agar tidak memakai layout split/dua kolom pada halaman Approval.
- Semua fitur v3.8/v4.6 tetap dipertahankan: Import Excel Aset, Export Excel per modul, Work Order WO, Akun, Konfigurasi, Laporan, Dokumen, GPS/KM, dan ID otomatis.

## v4.8.0 - Contract ID CTR Fix
- ID otomatis kontrak menggunakan prefix `CTR-YYYY-XXXX`, dengan prefix `CTR-`.
- Tombol dan judul form kontrak diperjelas menjadi **Tambah Kontrak**.
- Preview dan source app memakai aturan prefix yang sama.
- Semua fitur v3.8 serta perbaikan v4.7 tetap dipertahankan.
