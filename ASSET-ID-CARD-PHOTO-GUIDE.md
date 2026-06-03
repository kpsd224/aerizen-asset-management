# Aerizen v3.7 - Asset ID Card Photo

Perubahan utama:

- Tampilan foto kendaraan di kartu aset dan Aset 360 dibuat kembali seperti kartu BPKB / asset identity card.
- Foto kendaraan memakai `object-fit: contain`, sehingga tidak terpotong dan tidak gepeng.
- Kartu menampilkan nomor BPKB/asset card, contract number, customer, STNK, payment status, custody info, timeline, dan attachment BPKB.
- Fitur sebelumnya tetap dipertahankan: edit aset, Aset 360, STNK alert, maintenance, export Excel, akun, lupa password, lihat password, lonceng notifikasi, Supabase sync, dan installer resmi.

Build installer tetap memakai GitHub Actions:

1. Upload semua file project ke GitHub.
2. Buka Actions.
3. Jalankan Build Aerizen Windows Installer.
4. Download artifact `aerizen-windows-installer`.
5. Jalankan file `Aerizen-Asset-Management-Setup-3.7.0-x64.exe`.

Tidak perlu uninstall versi lama karena appId dan productName tetap sama.
