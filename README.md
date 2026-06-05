# Aerizen Asset Management v5.5

## Pembaruan v5.5

- Form Invoice dibuat lengkap dan terhubung dengan Kontrak serta Aset/Nopol.
- Data customer, PIC, email, telepon, alamat tagihan, NPWP, PO, periode, termin, dan jatuh tempo tersedia di form.
- Rincian tagihan mencakup nilai sewa, over KM, BBM, denda, service, biaya lain, diskon, PPN, PPh, total invoice, pembayaran, dan sisa piutang.
- Tersedia data metode pembayaran, rekening tujuan, referensi pembayaran, nomor faktur pajak, upload file invoice, upload faktur pajak, serta catatan.
- Total, PPN, potongan PPh, status, dan sisa piutang dihitung otomatis.
- Tabel Finance menampilkan tanggal, kontrak, aset/nopol, total, pembayaran, sisa, status, tombol edit, dan file invoice.
- Logo Aerizen yang diberikan digunakan pada header aplikasi, splash/loading, favicon, dan icon build Windows.
- Identitas versi ditampilkan sebagai **Aerizen Asset Management System · v5.5** dan judul aplikasi **Aerizen Asset Management v5.5**.
- Semua fitur v3.8 sampai v5.3 tetap dipertahankan.

## Perbaikan v5.3

- Form Kontrak dibuat lengkap: data customer dan PIC, aset/nopol, PO, periode, billing, deposit, biaya, diskon, PPN, total kontrak, batas KM, SLA/coverage, asuransi, replacement unit, approval, file kontrak, BAST, renewal, dan catatan.
- Nomor kontrak otomatis berurutan mengikuti format `CTR-YYYY-0001`.
- ID Dokumen pada menu Dokumen sekarang **manual**, sehingga dapat mengikuti nopol atau standar nomor dokumen perusahaan.
- Dokumen dapat dicari/dipilih berdasarkan nopol, serial, atau ID aset. Data aset, nopol, dan nama unit tertarik otomatis.
- Dokumen yang disimpan terhubung ke Digital Document Vault, kartu aset, dan Aset 360 berdasarkan `asetId` serta nopol/serial.
- Semua fitur v3.8 sampai v5.1 tetap dipertahankan.

# Riwayat v5.1

Versi ini memakai basis **Full Restore Lock v3.8**. Semua modul lama tetap dipertahankan: Dashboard, Aset, Kontrak, Operasional, Work Order, Finance, Sparepart, Vendor, Ticket, Dokumen, Approval, Laporan, Akun, dan Konfigurasi.

## Perbaikan v5.1

- ID Checklist Operasional mengikuti format data lama:
  - `CHK-OUT-003` untuk Check-out.
  - `CHK-IN-003` untuk Check-in.
  - `CHK-INS-003` untuk Inspection.
  - Saat tipe checklist diganti, prefix ID ikut berubah otomatis.
- ID Vendor mengikuti urutan lama `VDR-001`, `VDR-002`, `VDR-003`, lalu `VDR-004`.
- ID Sparepart mengikuti nama item, contohnya `SP-OLI-002`, `SP-BAN-002`, dan `SP-CHG-002`.
- Form Work Order dilengkapi dengan identitas WO, aset/nopol/KM, PIC, teknisi, vendor, jadwal, rincian biaya, PO, invoice, approval, warranty/claim, diagnosa, tindakan, QC, foto before/after, dan lampiran.
- Form Sparepart dilengkapi dengan part number, merek, spesifikasi, aset kompatibel, stok minimum/maksimum, rak, harga beli, nilai keluar, vendor, tanggal pembelian, kedaluwarsa, kondisi, status, dan catatan.
- Form Vendor dilengkapi dengan nama legal, PIC, telepon, email, alamat, NPWP, rekening, termin pembayaran, layanan, wilayah, rating, biaya rata-rata, lead time, SLA, kontrak, dokumen, status, dan catatan evaluasi.
- Lonceng notifikasi dipindahkan ke area navigasi agar selalu terlihat.
- Filter dan Export Excel Work Order per nopol tetap dipertahankan.

## Fitur yang tetap dipertahankan

- Import Excel kendaraan sewa beserta template XLSX/CSV.
- Kartu aset BPKB, upload foto, dan upload dokumen aset.
- GPS checkbox, ID GPS, provider GPS, dan KM/Odometer.
- Finance dengan Amount, Paid, Sisa, Due Date, Status, serta Profitability per Aset.
- Export Excel per modul.
- Approval Workflow berada di bawah Role & Permission User.
- Akun, Konfigurasi Supabase/offline sync, Laporan, backup, dan restore.

## Menjalankan aplikasi

```bash
npm install
npm run dev
npm run test
npm run dist
```

## Preview cepat

Buka `preview-review.html` untuk review langsung tanpa install npm.


## Pembaruan v5.3
- Pencarian aset langsung pada pilihan aset di form Dokumen.
- Daftar aset otomatis terfilter berdasarkan nopol, serial, ID, atau nama aset.
- Banner teknis Full Restore Lock dihapus dari antarmuka.
- Seluruh fitur versi sebelumnya tetap dipertahankan.
