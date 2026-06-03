# Aerizen v3.6 - Installer dan Foto Kendaraan HD

## Kenapa versi sebelumnya terasa tidak ter-install?
Build sebelumnya membuat installer NSIS dan portable EXE sekaligus dengan nama file yang mirip. Akibatnya pengguna bisa keliru menjalankan file portable, bukan installer. Versi ini membuat workflow default hanya menghasilkan installer resmi.

File yang harus dipakai untuk update/install:

```text
Aerizen-Asset-Management-Setup-3.6.0-x64.exe
```

Jalankan file Setup tersebut untuk menimpa versi lama. Tidak perlu uninstall versi lama selama appId tetap `com.aerizen.assetmanagement`.

## Perbaikan foto kendaraan
Tampilan foto aset sebelumnya memakai `object-fit: cover`, sehingga foto kendaraan bisa terlihat terpotong, gepeng, atau blur saat dimasukkan ke kartu yang kecil. Versi ini memakai `object-fit: contain`, tinggi preview lebih besar, dan foto asli tetap disimpan tanpa kompresi.

## Catatan
Mode portable masih tersedia lewat perintah manual:

```bash
npm run dist:portable
```

Namun GitHub Actions default dibuat installer-only agar proses update tidak membingungkan.
