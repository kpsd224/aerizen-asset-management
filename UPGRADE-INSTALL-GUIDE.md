# Panduan Update Tanpa Uninstall - Aerizen Asset Management

Mulai v3.3, installer Windows dibuat supaya pembaruan bisa langsung ditimpa ke versi lama.

## Cara update aplikasi

1. Build dari GitHub Actions.
2. Download artifact `aerizen-windows-exe`.
3. Jalankan file installer NSIS:
   - `Aerizen-Asset-Management-3.3.0-x64.exe`
4. Jangan pilih file `portable` untuk update aplikasi terpasang.
5. Installer akan memasang versi baru di lokasi aplikasi yang sama.
6. Data lokal/offline tetap aman selama `appId` dan `productName` tidak diganti.

## Penting

- Jangan ganti `appId`: `com.aerizen.assetmanagement`.
- Jangan ganti `productName`: `Aerizen Asset Management`.
- Jangan uninstall aplikasi lama kalau hanya mau update.
- File portable tidak melakukan upgrade instalasi. Portable hanya aplikasi lepas yang bisa dijalankan tanpa instalasi.
- Kalau pernah install versi lama di folder berbeda, jalankan installer v3.3 lalu gunakan lokasi default agar versi berikutnya konsisten.

## Pengaturan installer yang dipakai

```json
"nsis": {
  "oneClick": false,
  "allowToChangeInstallationDirectory": false,
  "deleteAppDataOnUninstall": false,
  "runAfterFinish": true
}
```

Dengan pengaturan ini, user cukup menjalankan installer versi terbaru. Tidak perlu uninstall versi lama.


## Jika aplikasi blank setelah update
Gunakan v3.4 atau lebih baru. Versi ini memakai path asset relatif dan recovery mode. Jalankan installer baru tanpa uninstall. Data lokal tidak dihapus.
