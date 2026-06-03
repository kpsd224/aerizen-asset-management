const fs = require('fs');
const path = require('path');

const required = [
  'package.json',
  'index.html',
  'src/App.jsx',
  'src/styles.css',
  'electron/main.cjs',
  'electron/preload.cjs',
  'supabase/schema.sql',
  'public/template_import_kendaraan_sewa.csv',
  '.github/workflows/build-windows.yml',
  'GITHUB-ACTIONS-WARNING-FIX.md',
  'UPGRADE-INSTALL-GUIDE.md',
];

for (const file of required) {
  const full = path.join(process.cwd(), file);
  if (!fs.existsSync(full)) {
    console.error(`Missing file: ${file}`);
    process.exit(1);
  }
}

const app = fs.readFileSync(path.join(process.cwd(), 'src/App.jsx'), 'utf8');
const checks = [
  'Aerizen',
  'createClient',
  'aerizen_records',
  'Import Excel Kendaraan Sewa',
  'Offline + Supabase',
  'STNK Perlu Perpanjang',
  'Nominal Biaya Aktual',
  'Export Excel',
  'sendPasswordReset',
  'resetPasswordForEmail',
  'PasswordField',
  'password-toggle',
  'completeWorkOrder',
  'Tanggal Selesai',
  'QR / Barcode',
  'NotificationBell',
  'notification-bell',
  'Lonceng Notifikasi',
  'AssetPhotoCard',
  'asset-id-card',
  'View Scanned BPKB',
];

for (const check of checks) {
  if (!app.includes(check)) {
    console.error(`Missing expected feature marker: ${check}`);
    process.exit(1);
  }
}

console.log('Aerizen v3.7 smoke-test passed.');
