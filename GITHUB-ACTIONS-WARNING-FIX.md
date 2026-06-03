# GitHub Actions Warning Fix

Versi ini memperbaiki annotation/warning yang muncul di GitHub Actions:

1. `Node.js 20 actions are deprecated`
   - Workflow sekarang memakai action yang sudah kompatibel dengan runtime Node 24:
     - `actions/checkout@v5`
     - `actions/setup-node@v6`
     - `actions/upload-artifact@v6`

2. `windows-latest requests are being redirected`
   - Workflow dipin ke `windows-2022` supaya runner stabil dan tidak mengikuti migrasi otomatis `windows-latest`.

3. Node untuk build project
   - Project di-build dengan Node.js `22` melalui `setup-node`.
   - Ini berbeda dari runtime internal GitHub Actions.

Jika build berhasil tetapi masih ada annotation kecil dari GitHub, itu biasanya notice platform, bukan error aplikasi.
