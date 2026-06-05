# GitHub Build Fix v5.6

Perbaikan ini khusus mengatasi kegagalan pada langkah **Install dependencies** di GitHub Actions.

- Seluruh URL dependency di `package-lock.json` memakai registry publik `https://registry.npmjs.org/`.
- `.npmrc` mengunci registry publik npm.
- Workflow memakai Node.js `22.16.0`, npm `10.9.2`, dan `npm ci`.
- Workflow menjalankan smoke test, build Vite, build installer Windows, lalu mengunggah file EXE.
- Semua fitur aplikasi v5.5 tetap dipertahankan.

Upload seluruh isi folder ini ke root repository GitHub, lalu buka **Actions → Build Aerizen Windows Installer → Run workflow**.
