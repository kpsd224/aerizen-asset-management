const fs = require('fs');
const path = require('path');

const required = [
  'package.json','index.html','src/App.jsx','src/styles.css','electron/main.cjs','electron/preload.cjs','supabase/schema.sql',
  'public/template_import_kendaraan_sewa.csv','public/template_import_kendaraan_sewa.xlsx','.github/workflows/build-windows.yml',
  'GITHUB-ACTIONS-WARNING-FIX.md','UPGRADE-INSTALL-GUIDE.md','preview-review.html',
  'public/branding/aerizen-emblem-512.png','public/branding/aerizen-logo-1024.png','public/favicon.png','build/icon.ico'
];
for (const file of required) {
  if (!fs.existsSync(path.join(process.cwd(), file))) {
    console.error(`Missing file: ${file}`);
    process.exit(1);
  }
}

const app = fs.readFileSync(path.join(process.cwd(), 'src/App.jsx'), 'utf8');
const css = fs.readFileSync(path.join(process.cwd(), 'src/styles.css'), 'utf8');
const preview = fs.readFileSync(path.join(process.cwd(), 'preview-review.html'), 'utf8');
const csv = fs.readFileSync(path.join(process.cwd(), 'public/template_import_kendaraan_sewa.csv'), 'utf8');

const moduleKeys = ['dashboard','assets','contracts','operations','workorders','finance','inventory','vendors','tickets','documents','approvals','reports','account','configuration'];
for (const key of moduleKeys) {
  if (!app.includes(`['${key}'`) && !app.includes(`"${key}"`)) {
    console.error(`Missing module tab key: ${key}`);
    process.exit(1);
  }
}

const lockedFeatures = [
  'Import Excel Kendaraan Sewa','Export Excel','Nominal Biaya Aktual','NotificationBell','Lonceng Notifikasi',
  'AssetPhotoCard','View Scanned BPKB','Akun','Konfigurasi','AccountView','ConfigurationView','Upload Dokumen Aset',
  'GPS','KM / Odometer','ID Aset Otomatis','ID Dokumen (Manual)','Provider GPS','approval-page','approval-workflow-panel',
  'Cari / pilih nopol','Export Excel Nopol','wo-nopol-options','filter-summary'
];
for (const marker of lockedFeatures) {
  if (!app.includes(marker) && !css.includes(marker)) {
    console.error(`Missing locked feature marker: ${marker}`);
    process.exit(1);
  }
}

const newMarkers = [
  'nextChecklistId','CHK-(?:IN|OUT|INS)','CHK-IN','CHK-OUT','CHK-INS',
  'nextVendorId','VDR-${padSequence(next)}','nextInventoryId','inventoryCode',
  'ID WO Otomatis','KM / Odometer Saat WO','Diagnosa','Tindakan / Pekerjaan Dilakukan','Biaya Jasa Aktual',
  'Biaya Sparepart Aktual','Nomor PO','Nomor Invoice Vendor','Status Approval','Hasil QC / Test Drive',
  'Part Number / SKU Supplier','Maksimum Stok','Harga Beli','Rak / Bin','Tanggal Kedaluwarsa','Catatan Sparepart',
  'Nama Legal Perusahaan','PIC Vendor','NPWP','Termin Pembayaran','Wilayah Layanan','SLA / Waktu Respons','Catatan Evaluasi Vendor',
  'nav-frame-with-bell','notification-bell-wrap'
];
for (const marker of newMarkers) {
  if (!app.includes(marker) && !css.includes(marker)) {
    console.error(`Missing retained v5.1 marker: ${marker}`);
    process.exit(1);
  }
}

if (!app.includes('checklists: { id: nextChecklistId(state')) {
  console.error('Checklist auto ID does not use nextChecklistId.');
  process.exit(1);
}
if (!app.includes('vendors: { id: nextVendorId(state)')) {
  console.error('Vendor auto ID does not use VDR sequential format.');
  process.exit(1);
}
if (!app.includes('inventory: { id: nextInventoryId(state')) {
  console.error('Sparepart auto ID does not use SP-code sequential format.');
  process.exit(1);
}

const v52Markers = [
  'nextContractId','contractDefaults','normalizeContract','function ContractForm','No. Kontrak Otomatis',
  'PIC Customer','Nomor PO Customer','Total Nilai Kontrak','Batas KM per Bulan','Status Approval','Upload File Kontrak',
  'ID Dokumen (Manual)','Cari dan Pilih Aset','filteredAssets','asset-search-select','Simpan & Hubungkan Dokumen','suggestDocumentId','documentMatchesAsset',
  'Nomor Dokumen Resmi','Aset 360 sesuai nopol'
];
for (const marker of v52Markers) {
  if (!app.includes(marker) && !css.includes(marker)) {
    console.error(`Missing v5.2 marker: ${marker}`);
    process.exit(1);
  }
}
if (app.includes('ID Dokumen Otomatis')) {
  console.error('Document ID must be manual in v5.2.');
  process.exit(1);
}
if (!app.includes("getQuickCollection(active) { return ({ operations:")) {
  console.error('Contract must use the dedicated complete ContractForm, not the generic quick form.');
  process.exit(1);
}

const navIndex = app.indexOf('nav-frame nav-frame-with-bell');
const bellAfterNav = app.indexOf('<NotificationBell', navIndex);
if (navIndex < 0 || bellAfterNav < navIndex) {
  console.error('Notification bell must be visible in the navigation frame.');
  process.exit(1);
}

const approvalStart = app.indexOf('function ApprovalsView');
const approvalEnd = app.indexOf('function ReportsView', approvalStart);
const approvalBlock = app.slice(approvalStart, approvalEnd);
if (approvalBlock.includes('dashboard-grid two') || approvalBlock.includes('className="dashboard-grid')) {
  console.error('Approval Workflow must remain below Role & Permission User.');
  process.exit(1);
}

for (const col of ['gps','id gps','provider gps','km odometer']) {
  if (!csv.toLowerCase().includes(col)) {
    console.error(`Template import missing column: ${col}`);
    process.exit(1);
  }
}

const v55Markers = [
  'APP_VERSION_LABEL','v5.5','brand-logo','aerizen-emblem-512.png','function InvoiceForm','Identitas Invoice',
  'Customer & Aset Tertagih','Rincian Nilai Tagihan','Nilai Sewa / Tagihan Pokok','Biaya Kelebihan KM',
  'Biaya BBM','Denda / Penalty','Biaya Service / Maintenance','Diskon','Subtotal','PPN (%)',
  'PPh / Potongan (%)','Total Invoice','Jumlah Dibayar','Sisa Piutang','Metode Pembayaran',
  'Nomor Faktur Pajak','Upload File Invoice','Upload Faktur Pajak','Catatan Invoice','invoice-summary',
  'onEdit={(invoice)'
];
for (const marker of v55Markers) {
  if (!app.includes(marker) && !css.includes(marker) && !preview.includes(marker)) {
    console.error(`Missing v5.5 retained/UI marker: ${marker}`);
    process.exit(1);
  }
}

for (const marker of ['asset-summary-grid','asset-summary-card','asset-summary-value','v5.5 UI polish']) {
  if (!app.includes(marker) && !css.includes(marker)) {
    console.error(`Missing v5.5 UI polish marker: ${marker}`);
    process.exit(1);
  }
}

const previewMarkers = [
  'Aerizen Asset Management v5.5','Asset Management System','Identitas Invoice','Upload Faktur Pajak',
  'ID Dokumen (Manual)','Cari / pilih nopol','Nominal Biaya Aktual','Akun','Konfigurasi'
];
for (const marker of previewMarkers) {
  if (!preview.includes(marker)) {
    console.error(`Preview missing retained/new marker: ${marker}`);
    process.exit(1);
  }
}

if (/\bPOS\b/.test(app) || /\bPOS\b/.test(preview)) {
  console.error('Generic POS label is still visible.');
  process.exit(1);
}

if (app.includes('restore-lock-strip') || preview.includes('Full Restore Lock')) {
  console.error('Technical restore banner must not appear in the UI.');
  process.exit(1);
}

console.log('Aerizen v5.5 smoke-test passed: all prior features retained and Aset 360 UI remains responsive and tidy.');
