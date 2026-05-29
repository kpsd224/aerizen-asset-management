# Aerizen Asset Management - Build EXE

## Cara menjalankan preview desktop

1. Install Node.js LTS.
2. Buka folder project ini di VS Code atau Terminal.
3. Jalankan:

```bash
npm install
npm run electron:dev
```

## Cara membuat installer .exe

```bash
npm install
npm run dist
```

Hasil installer akan muncul di folder:

```text
release/Aerizen-Asset-Management-Setup-1.0.0.exe
```

## Cara membuat portable .exe

```bash
npm run dist:portable
```

## Fitur utama

- Dashboard manajemen aset enterprise.
- Aset kendaraan, perangkat IT, peralatan kantor, dan lainnya.
- Import Excel kendaraan sewa.
- Template Excel tersedia di `public/templates/template_import_kendaraan_sewa.xlsx`.
- Form edit, tambah, hapus aset.
- Work Order maintenance.
- Detail Aset 360.

## Format Excel Kendaraan Sewa

Header yang didukung:

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
