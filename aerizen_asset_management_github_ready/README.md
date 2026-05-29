# Aerizen Asset Management

Aplikasi desktop Asset Management untuk kendaraan sewa, aset internal, perangkat IT, maintenance, work order, dan import Excel kendaraan sewa.

## Cara upload ke GitHub

1. Buat repository baru di GitHub.
2. Download dan ekstrak ZIP project ini.
3. Upload semua isi folder ini ke repository GitHub, bukan file ZIP-nya.
4. Pastikan struktur file di GitHub terlihat seperti ini:

```text
.github/workflows/build-windows-exe.yml
src/App.jsx
src/main.jsx
electron/main.cjs
public/templates/template_import_kendaraan_sewa.xlsx
package.json
index.html
```

## Cara build EXE otomatis dari GitHub

1. Buka repository GitHub.
2. Masuk ke tab **Actions**.
3. Pilih workflow **Build Windows EXE**.
4. Klik **Run workflow**.
5. Tunggu proses selesai.
6. Buka hasil workflow.
7. Download artifact bernama **Aerizen-Asset-Management-Windows-Installer**.
8. Di dalam artifact tersebut ada file:

```text
Aerizen-Asset-Management-Setup-1.0.0.exe
```

## Cara build di laptop sendiri

```bash
npm install
npm run dist
```

File `.exe` akan muncul di folder:

```text
release/
```

## Cara preview development

```bash
npm install
npm run electron:dev
```

## Format Import Excel Kendaraan Sewa

Template tersedia di:

```text
public/templates/template_import_kendaraan_sewa.xlsx
```

Kolom yang didukung:

- no
- nama kendaraan
- warna
- tahun kendaraan
- nomor rangka
- nomor mesin
- perusahaan user
- alamat user
- wilayah
- tanggal STNK
- memiliki BPKB
