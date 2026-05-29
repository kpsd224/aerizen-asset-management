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
  'Lupa Password',
  'resetPasswordForEmail',
  'PasswordField',
  'password-toggle',
  'completeWorkOrder',
  'Aksi Cepat',
  'Tanggal Selesai',
  'Role & Permission User',
  'QR / Barcode',
  'Kontrak Sewa',
  'Check-in / Check-out',
  'Digital Inspection Checklist',
  'Billing & Invoice',
  'Sparepart & Inventory',
  'Vendor / Bengkel / Supplier Management',
  'SLA & Ticketing',
  'Digital Document Vault',
  'Notification Center',
  'NotificationBell',
  'notification-bell',
  'Lonceng Notifikasi',
  'Approval Workflow',
  'Audit Trail',
];

for (const check of checks) {
  if (!app.includes(check)) {
    console.error(`Missing expected feature marker: ${check}`);
    process.exit(1);
  }
}

console.log('Aerizen v3.1 smoke-test passed.');
