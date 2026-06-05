import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

const STORAGE_KEY = 'aerizen.v3.state';
const SETTINGS_KEY = 'aerizen.v3.settings';
const QUEUE_KEY = 'aerizen.v3.syncQueue';
// v5.5 UI Polish: seluruh fitur v3.8-v5.4 dipertahankan. Layout Aset 360 dan komponen responsif dirapikan agar tidak terpotong atau keluar dari kartu.
const APP_VERSION = '5.5.0';
const APP_VERSION_LABEL = 'v5.5';

const today = new Date().toISOString().slice(0, 10);
const currency = new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 });

const statusOptions = ['Tersedia', 'Disewakan', 'Dipakai Internal', 'Maintenance', 'Perbaikan', 'Overdue', 'Nonaktif'];


const categoryOptions = ['Kendaraan', 'Perangkat IT', 'Peralatan Kantor', 'Mesin', 'Properti', 'Lainnya'];
const workOrderStatus = ['Draft', 'Menunggu Approval', 'Dijadwalkan', 'Dikerjakan', 'Menunggu Sparepart', 'Menunggu Vendor', 'Menunggu QC', 'Selesai', 'Ditutup'];
const workOrderType = ['Preventive Maintenance', 'Corrective Maintenance', 'Emergency Repair', 'Vendor Repair', 'Warranty Claim', 'Inspection', 'Reconditioning'];
const priorityOptions = ['Rendah', 'Sedang', 'Tinggi', 'Kritis'];

const tabs = [
  ['dashboard', 'Dashboard'],
  ['assets', 'Aset'],
  ['contracts', 'Kontrak'],
  ['operations', 'Operasional'],
  ['workorders', 'Work Order (WO)'],
  ['finance', 'Finance'],
  ['inventory', 'Sparepart'],
  ['vendors', 'Vendor'],
  ['tickets', 'Ticket'],
  ['documents', 'Dokumen'],
  ['approvals', 'Approval'],
  ['reports', 'Laporan'],
  ['account', 'Akun'],
  ['configuration', 'Konfigurasi'],
];

const REQUIRED_V38_MODULES = ['dashboard','assets','contracts','operations','workorders','finance','inventory','vendors','tickets','documents','approvals','reports','account','configuration'];
const missingV38Modules = REQUIRED_V38_MODULES.filter((key) => !tabs.some(([tabKey]) => tabKey === key));
if (missingV38Modules.length) console.warn('Aerizen module guard: modul wajib v3.8 belum masuk:', missingV38Modules.join(', '));

const initialAssets = [
  {
    id: 'AST-KND-2026-0184',
    nama: 'Toyota Innova Zenix',
    kategori: 'Kendaraan', tipe: 'Aset Rental', status: 'Disewakan', kondisi: 'Sangat Baik', cabang: 'Jakarta', lokasi: 'Pool Jakarta A',
    pemegang: 'PT Nusantara Retail', perusahaanUser: 'PT Nusantara Retail', alamatUser: 'Jl. Sudirman No. 10, Jakarta', wilayah: 'Jakarta',
    warna: 'Hitam', tahunKendaraan: '2024', nomorRangka: 'MHFABCD1234567890', nomorMesin: '2GDE123456', nomorPolisi: 'B 184 AZN',
    qrCode: 'QR-AERIZEN-AST-KND-2026-0184', tanggalSTNK: '2026-12-31', tanggalKIR: '2026-08-12', tanggalPajak: '2026-12-31', tanggalAsuransi: '2026-09-30', tanggalServiceBerikutnya: '2026-07-05', memilikiBPKB: 'Ya',
    gambar: '', risiko: 18, utilisasi: 91, nilaiBuku: 287500000, pendapatanBulanan: 18400000, biayaBulanan: 4200000,
    aksiBerikutnya: 'Perpanjangan kontrak dalam 6 hari', dokumen: 'STNK berlaku sampai 2026-12-31 · BPKB: Ya', telemetri: 'GPS Online · 42.180 km',
    alur: 'Disewakan → Monitoring → Perpanjangan / Pengembalian', riwayat: ['Terdaftar', 'Dipasang QR', 'Disewakan', 'GPS Sinkron', 'Invoice Terbit'], updatedAt: new Date().toISOString(),
  },
  {
    id: 'AST-IT-2026-0441',
    nama: 'HP EliteBook 840 G11', kategori: 'Perangkat IT', tipe: 'Aset Rental IT', status: 'Dipakai Internal', kondisi: 'Baik', cabang: 'Bandung', lokasi: 'Cabang Bandung',
    pemegang: 'Ayu Pramesti · Tim Produk', perusahaanUser: 'Internal', alamatUser: 'Cabang Bandung', wilayah: 'Bandung', warna: 'Silver', tahunKendaraan: '-',
    nomorRangka: 'SN-HP840-2026-0441', nomorMesin: 'IMEI-MDM-0441', nomorPolisi: '-', qrCode: 'QR-AERIZEN-AST-IT-2026-0441', tanggalSTNK: '', tanggalKIR: '', tanggalPajak: '', tanggalAsuransi: '', tanggalServiceBerikutnya: '2026-06-18', memilikiBPKB: '-',
    gambar: '', risiko: 26, utilisasi: 76, nilaiBuku: 31900000, pendapatanBulanan: 0, biayaBulanan: 850000,
    aksiBerikutnya: 'Cek kesehatan baterai dan garansi', dokumen: 'Garansi berlaku 218 hari lagi', telemetri: 'MDM Online · Enkripsi Aktif',
    alur: 'Ditugaskan → Pemeriksaan Berkala → Kembali / Mutasi', riwayat: ['Terdaftar', 'Setup Keamanan', 'Ditugaskan', 'Cek MDM'], updatedAt: new Date().toISOString(),
  },
  {
    id: 'AST-KND-2025-0098',
    nama: 'Mitsubishi L300 Box', kategori: 'Kendaraan', tipe: 'Aset Operasional', status: 'Maintenance', kondisi: 'Perlu Servis', cabang: 'Surabaya', lokasi: 'Yard Surabaya',
    pemegang: 'Logistik Internal', perusahaanUser: 'Internal Logistik', alamatUser: 'Yard Surabaya', wilayah: 'Surabaya', warna: 'Putih', tahunKendaraan: '2022',
    nomorRangka: 'MK2L3009876543210', nomorMesin: '4D56ABC987', nomorPolisi: 'L 98 AZN', qrCode: 'QR-AERIZEN-AST-KND-2025-0098', tanggalSTNK: '2026-06-20', tanggalKIR: '2026-06-17', tanggalPajak: '2026-06-20', tanggalAsuransi: '2026-07-02', tanggalServiceBerikutnya: '2026-06-08', memilikiBPKB: 'Ya',
    gambar: '', risiko: 72, utilisasi: 58, nilaiBuku: 132000000, pendapatanBulanan: 9700000, biayaBulanan: 7100000,
    aksiBerikutnya: 'Inspeksi rem menunggu QC', dokumen: 'KIR habis dalam 19 hari', telemetri: 'GPS Online · 83.900 km',
    alur: 'Work Order → Vendor Repair → QC → Tersedia', riwayat: ['Terdaftar', 'Dipakai Logistik', 'Insiden', 'Work Order', 'Menunggu QC'], updatedAt: new Date().toISOString(),
  },
];

const initialWorkOrders = [
  { id: 'WO-2026-0912', asetId: 'AST-KND-2025-0098', namaAset: 'Mitsubishi L300 Box', nopol: 'L 98 AZN', jenis: 'Corrective Maintenance', prioritas: 'Tinggi', status: 'Menunggu QC', pic: 'Vendor Surabaya', vendorId: 'VDR-002', tanggalMulai: '2026-05-28', jatuhTempo: '2026-05-29', tanggalSelesai: '', biaya: 3500000, sparepart: 'Brake pad depan', fotoBefore: 'Belum diunggah', fotoAfter: 'Menunggu QC', keluhan: 'Rem kurang pakem dan muncul suara saat pengereman.', checklist: 'Cek rem depan, cek rem belakang, test drive, foto evidence, QC akhir', updatedAt: new Date().toISOString() },
  { id: 'WO-2026-0825', asetId: 'AST-KND-2026-0184', namaAset: 'Toyota Innova Zenix', nopol: 'B 184 AZN', jenis: 'Preventive Maintenance', prioritas: 'Sedang', status: 'Selesai', pic: 'Bengkel Rekanan Jakarta', vendorId: 'VDR-001', tanggalMulai: '2026-05-10', jatuhTempo: '2026-05-10', tanggalSelesai: '2026-05-10', biaya: 1250000, sparepart: 'Oli mesin + filter', fotoBefore: 'Ada', fotoAfter: 'Ada', keluhan: 'Servis berkala 10.000 km dan ganti oli.', checklist: 'Ganti oli, cek filter, cek rem, QC akhir', updatedAt: new Date().toISOString() },
];

const initialContracts = [
  { id: 'CTR-2026-0007', customer: 'PT Nusantara Retail', asetId: 'AST-KND-2026-0184', aset: 'Toyota Innova Zenix', nilaiBulanan: 18400000, deposit: 30000000, mulai: '2026-01-01', selesai: '2026-06-30', status: 'Aktif', billingCycle: 'Bulanan', dokumen: 'Kontrak_Nusantara_0007.pdf', renewal: 'Perlu follow up H-30' },
  { id: 'CTR-2026-0012', customer: 'PT Sinar Logistik', asetId: 'AST-KND-2025-0098', aset: 'Mitsubishi L300 Box', nilaiBulanan: 9700000, deposit: 15000000, mulai: '2026-03-15', selesai: '2026-10-15', status: 'Menunggu Dokumen', billingCycle: 'Bulanan', dokumen: 'Draft_Sinar_0012.pdf', renewal: 'Menunggu PO' },
];

const initialInvoices = [
  { id: 'INV-2026-0501', contractId: 'CTR-2026-0007', customer: 'PT Nusantara Retail', periode: 'Mei 2026', amount: 18400000, paid: 18400000, dueDate: '2026-05-25', status: 'Paid' },
  { id: 'INV-2026-0601', contractId: 'CTR-2026-0007', customer: 'PT Nusantara Retail', periode: 'Juni 2026', amount: 18400000, paid: 0, dueDate: '2026-06-25', status: 'Unpaid' },
  { id: 'INV-2026-0602', contractId: 'CTR-2026-0012', customer: 'PT Sinar Logistik', periode: 'Juni 2026', amount: 9700000, paid: 2500000, dueDate: '2026-06-20', status: 'Partial' },
];

const initialChecklists = [
  { id: 'CHK-OUT-001', tipe: 'Check-out', asetId: 'AST-KND-2026-0184', aset: 'Toyota Innova Zenix', tanggal: '2026-05-24', pic: 'Driver Jakarta', penerima: 'Rizky · PT Nusantara Retail', km: '42.180', bbm: '80%', kondisi: 'Body baik, interior bersih', aksesoris: 'STNK, toolkit, dongkrak, kunci cadangan', tandaTangan: 'Sudah', status: 'Selesai' },
  { id: 'CHK-IN-002', tipe: 'Check-in', asetId: 'AST-IT-2026-0441', aset: 'HP EliteBook 840 G11', tanggal: '2026-05-28', pic: 'IT Support Bandung', penerima: 'Ayu Pramesti', km: '-', bbm: '-', kondisi: 'Keyboard dan charger lengkap', aksesoris: 'Charger, sleeve, stylus', tandaTangan: 'Menunggu', status: 'Draft' },
];

const initialInventory = [
  { id: 'SP-OLI-001', nama: 'Oli Mesin 5W-30', partNumber: 'OLI-5W30-4L', merek: 'OEM', kategori: 'Kendaraan', spesifikasi: 'Full synthetic 5W-30', asetKompatibel: 'Toyota Innova Zenix', stok: 24, min: 10, stokMax: 40, satuan: 'Liter', lokasi: 'Gudang Jakarta', rak: 'A-01', hargaBeli: 155000, nilai: 185000, supplierId: 'VDR-001', vendor: 'Bengkel Rekanan Jakarta', tanggalBeli: '2026-05-10', tanggalExpired: '2028-05-10', kondisi: 'Baru', status: 'Aktif', catatan: '' },
  { id: 'SP-BAN-001', nama: 'Ban 205/65 R16', partNumber: 'BAN-20565R16', merek: 'Bridgestone', kategori: 'Kendaraan', spesifikasi: '205/65 R16', asetKompatibel: 'MPV / Toyota Innova', stok: 8, min: 8, stokMax: 20, satuan: 'Pcs', lokasi: 'Pool Surabaya', rak: 'B-04', hargaBeli: 850000, nilai: 980000, supplierId: '', vendor: 'PT Ban Prima', tanggalBeli: '2026-04-20', tanggalExpired: '', kondisi: 'Baru', status: 'Minimum Stock', catatan: 'Segera restock' },
  { id: 'SP-CHG-001', nama: 'Charger USB-C 65W', partNumber: 'CHG-USBC-65W', merek: 'HP', kategori: 'Perangkat IT', spesifikasi: 'USB-C PD 65W', asetKompatibel: 'HP EliteBook 840 G11', stok: 17, min: 12, stokMax: 30, satuan: 'Unit', lokasi: 'Gudang IT Bandung', rak: 'IT-C-02', hargaBeli: 275000, nilai: 325000, supplierId: 'VDR-003', vendor: 'IT Supplier Bandung', tanggalBeli: '2026-05-01', tanggalExpired: '', kondisi: 'Baru', status: 'Aktif', catatan: '' },
];

const initialVendors = [
  { id: 'VDR-001', nama: 'Bengkel Rekanan Jakarta', namaLegal: 'PT Bengkel Rekanan Jakarta', jenis: 'Bengkel Kendaraan', pic: 'Budi Santoso', kontak: '0812-0000-1001', email: 'service@bengkelrekanan.co.id', alamat: 'Jakarta Selatan', kota: 'Jakarta', provinsi: 'DKI Jakarta', npwp: '01.234.567.8-001.000', bank: 'BCA', rekening: '1234567890', paymentTerm: '30 hari', layanan: 'Preventive, corrective, emergency repair', wilayahLayanan: 'Jabodetabek', rating: 4.8, avgCost: 1450000, avgLeadTime: '1.2 hari', sla: 'Respons 2 jam', transaksi: 38, kontrakMulai: '2026-01-01', kontrakSelesai: '2026-12-31', dokumen: 'PKS-VDR-001.pdf', status: 'Aktif', catatan: '' },
  { id: 'VDR-002', nama: 'Vendor Surabaya', namaLegal: 'CV Vendor Surabaya', jenis: 'Repair Kendaraan', pic: 'Andi Wijaya', kontak: '0812-0000-2002', email: 'repair@vendorsurabaya.co.id', alamat: 'Surabaya Barat', kota: 'Surabaya', provinsi: 'Jawa Timur', npwp: '', bank: 'Mandiri', rekening: '', paymentTerm: '14 hari', layanan: 'Repair kendaraan dan body repair', wilayahLayanan: 'Surabaya dan sekitarnya', rating: 4.2, avgCost: 3200000, avgLeadTime: '2.4 hari', sla: 'Respons 4 jam', transaksi: 21, kontrakMulai: '2026-01-01', kontrakSelesai: '2026-12-31', dokumen: 'PKS-VDR-002.pdf', status: 'Aktif', catatan: '' },
  { id: 'VDR-003', nama: 'IT Supplier Bandung', namaLegal: 'PT IT Supplier Bandung', jenis: 'Perangkat IT', pic: 'Rina Putri', kontak: '0812-0000-3003', email: 'sales@itsupplierbandung.co.id', alamat: 'Bandung', kota: 'Bandung', provinsi: 'Jawa Barat', npwp: '', bank: 'BNI', rekening: '', paymentTerm: '30 hari', layanan: 'Sparepart dan dukungan perangkat IT', wilayahLayanan: 'Jawa Barat', rating: 4.6, avgCost: 825000, avgLeadTime: '1.7 hari', sla: 'Respons 4 jam', transaksi: 16, kontrakMulai: '2026-01-01', kontrakSelesai: '2026-12-31', dokumen: 'PKS-VDR-003.pdf', status: 'Aktif', catatan: '' },
];

const initialTickets = [
  { id: 'TCK-2026-0044', customer: 'PT Nusantara Retail', asetId: 'AST-KND-2026-0184', aset: 'Toyota Innova Zenix', kategori: 'Keluhan customer', prioritas: 'Sedang', sla: '2026-05-30', status: 'Open', pic: 'Customer Success', deskripsi: 'Customer meminta jadwal penggantian unit saat maintenance.' },
  { id: 'TCK-2026-0045', customer: 'Internal Logistik', asetId: 'AST-KND-2025-0098', aset: 'Mitsubishi L300 Box', kategori: 'Mobilisasi', prioritas: 'Tinggi', sla: '2026-05-29', status: 'In Progress', pic: 'Ops Surabaya', deskripsi: 'Unit belum bisa keluar karena menunggu QC rem.' },
];

const initialDocuments = [
  { id: 'DOC-001', asetId: 'AST-KND-2026-0184', nama: 'STNK Toyota Innova Zenix', tipe: 'STNK', expired: '2026-12-31', pemilik: 'Legal Asset', status: 'Valid', file: 'stnk-innova.pdf' },
  { id: 'DOC-002', asetId: 'AST-KND-2025-0098', nama: 'KIR Mitsubishi L300 Box', tipe: 'KIR', expired: '2026-06-17', pemilik: 'Operasional Surabaya', status: 'Hampir Habis', file: 'kir-l300.pdf' },
  { id: 'DOC-003', asetId: 'AST-IT-2026-0441', nama: 'Garansi HP EliteBook', tipe: 'Garansi IT', expired: '2027-01-05', pemilik: 'IT Asset', status: 'Valid', file: 'garansi-elitebook.pdf' },
];

const initialApprovals = [
  { id: 'APR-REPAIR-001', proses: 'Repair > Rp5 juta', approval: 'Supervisor + Finance', limit: 5000000, status: 'Aktif', pending: '0' },
  { id: 'APR-DISPOSAL-001', proses: 'Disposal aset', approval: 'Asset Manager + Direktur', limit: 0, status: 'Aktif', pending: '1' },
  { id: 'APR-RENTAL-001', proses: 'Harga rental khusus', approval: 'Sales Manager', limit: 0, status: 'Aktif', pending: '2' },
  { id: 'APR-MUTASI-001', proses: 'Mutasi antar cabang', approval: 'Dua Branch Manager', limit: 0, status: 'Aktif', pending: '0' },
];

const initialRoles = [
  { id: 'ROLE-001', role: 'Super Admin', akses: 'Semua modul, konfigurasi, hapus data, restore backup', user: 'Owner / Direktur' },
  { id: 'ROLE-002', role: 'Admin Asset', akses: 'Tambah/edit aset, dokumen, import Excel, QR aset', user: 'Asset team' },
  { id: 'ROLE-003', role: 'Operasional', akses: 'Check-in/out, kontrak, delivery, pickup, ticket SLA', user: 'Field ops' },
  { id: 'ROLE-004', role: 'Maintenance', akses: 'Work Order, biaya aktual, sparepart, vendor, foto before/after', user: 'Maintenance team' },
  { id: 'ROLE-005', role: 'Finance', akses: 'Invoice, billing, revenue, profit, laporan', user: 'Finance team' },
  { id: 'ROLE-006', role: 'Client Viewer', akses: 'Hanya lihat aset, kontrak, invoice, ticket milik customer', user: 'Customer portal' },
];

const initialState = {
  assets: initialAssets.map((asset) => normalizeAsset(asset)),
  workOrders: initialWorkOrders,
  contracts: initialContracts,
  invoices: initialInvoices.map((invoice) => normalizeInvoice(invoice, { contracts: initialContracts, assets: initialAssets })),
  checklists: initialChecklists,
  inventory: initialInventory,
  vendors: initialVendors,
  tickets: initialTickets,
  documents: initialDocuments.map((doc) => normalizeDocument(doc)),
  approvals: initialApprovals,
  roles: initialRoles,
  auditLogs: [
    { id: 'LOG-001', waktu: new Date().toISOString(), aksi: 'Seed data', detail: 'Data awal Aerizen v3.0 Enterprise dibuat offline', user: 'System' },
  ],
};

function safeParse(value, fallback) {
  try { return value ? JSON.parse(value) : fallback; } catch { return fallback; }
}
function uid(prefix) { return `${prefix}-${new Date().getFullYear()}-${Math.floor(1000 + Math.random() * 8999)}`; }
function padSequence(value, digits = 3) { return String(value).padStart(digits, '0'); }
function nextSuffix(rows = [], matcher, fallback = 1) {
  const numbers = rows.map((row) => String(row?.id || '').match(matcher)).filter(Boolean).map((match) => Number(match[1])).filter(Number.isFinite);
  return numbers.length ? Math.max(...numbers) + 1 : fallback;
}
function checklistPrefix(type = 'Check-out') { return type === 'Check-in' ? 'CHK-IN' : type === 'Inspection' ? 'CHK-INS' : 'CHK-OUT'; }
function nextChecklistId(state = {}, type = 'Check-out') {
  const prefix = checklistPrefix(type);
  const next = nextSuffix(state.checklists || [], /^CHK-(?:IN|OUT|INS)-(\d+)$/i);
  return `${prefix}-${padSequence(next)}`;
}
function nextVendorId(state = {}) {
  const next = nextSuffix(state.vendors || [], /^VDR-(\d+)$/i);
  return `VDR-${padSequence(next)}`;
}
function nextContractId(state = {}) {
  const year = new Date().getFullYear();
  const next = nextSuffix(state.contracts || [], new RegExp(`^CTR-${year}-(\\d+)$`, 'i'));
  return `CTR-${year}-${padSequence(next, 4)}`;
}
function compactReference(value = '') { return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function suggestDocumentId(asset = {}, type = 'DOC', fileName = '') {
  const typeCode = compactReference(type).slice(0, 10) || 'DOC';
  const assetCode = compactReference(asset.nomorPolisi || asset.nomorRangka || asset.id || 'ASET').slice(0, 18) || 'ASET';
  const fileCode = compactReference(String(fileName || '').replace(/\.[^.]+$/, '')).slice(0, 10);
  return `${typeCode}-${assetCode}${fileCode ? `-${fileCode}` : ''}`;
}
function documentMatchesAsset(doc = {}, asset = {}) {
  if (doc.asetId && asset.id && doc.asetId === asset.id) return true;
  const docRef = compactReference(doc.nopol || doc.nomorPolisi || doc.nomorRangka);
  const assetRef = compactReference(asset.nomorPolisi || asset.nomorRangka);
  return !!docRef && !!assetRef && docRef === assetRef;
}
function monthsBetween(startDate, endDate) {
  const start = parseDateOnly(startDate); const end = parseDateOnly(endDate);
  if (!start || !end || end < start) return 0;
  const raw = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  return Math.max(1, raw + (end.getDate() >= start.getDate() ? 1 : 0));
}
function contractDefaults(state = {}) {
  const asset = (state.assets || [])[0] || {};
  return {
    id: nextContractId(state), customer: '', picCustomer: '', teleponCustomer: '', emailCustomer: '', alamatCustomer: '', npwpCustomer: '',
    asetId: asset.id || '', aset: asset.nama || '', nopol: asset.nomorPolisi || asset.nomorRangka || '', nomorRangka: asset.nomorRangka || '', cabang: asset.cabang || '', lokasiSerahTerima: asset.lokasi || '',
    nomorPO: '', tanggalPO: '', mulai: today, selesai: today, durasiBulan: 1, billingCycle: 'Bulanan', terminPembayaran: '30 hari', tanggalJatuhTempo: '25',
    nilaiBulanan: 0, deposit: 0, biayaAdmin: 0, diskon: 0, ppnPersen: 11, totalKontrak: 0, batasKmBulanan: 0, tarifOverKm: 0, kmAwal: asset.km || '',
    cakupanService: 'Termasuk maintenance berkala', unitPengganti: 'Sesuai SLA', asuransi: 'Termasuk', pengemudi: 'Tidak termasuk',
    approvalStatus: 'Belum Diajukan', status: 'Menunggu Dokumen', renewal: '', dokumen: '', dokumenFileData: '', dokumenFileType: '', bast: '', catatan: ''
  };
}
function normalizeContract(contract = {}, state = {}) {
  const asset = (state.assets || []).find((item) => item.id === contract.asetId) || {};
  const duration = numeric(contract.durasiBulan) || monthsBetween(contract.mulai, contract.selesai) || 1;
  const subtotal = numeric(contract.nilaiBulanan) * duration + numeric(contract.biayaAdmin) - numeric(contract.diskon);
  const ppn = Math.round(subtotal * (Number(contract.ppnPersen || 0) / 100));
  return {
    ...contract,
    id: String(contract.id || nextContractId(state)).trim(),
    asetId: asset.id || contract.asetId || '', aset: asset.nama || contract.aset || '',
    nopol: asset.nomorPolisi || asset.nomorRangka || contract.nopol || '', nomorRangka: asset.nomorRangka || contract.nomorRangka || '', cabang: asset.cabang || contract.cabang || '',
    nilaiBulanan: numeric(contract.nilaiBulanan), deposit: numeric(contract.deposit), biayaAdmin: numeric(contract.biayaAdmin), diskon: numeric(contract.diskon),
    ppnPersen: Number(contract.ppnPersen || 0), durasiBulan: duration, totalKontrak: numeric(contract.totalKontrak) || subtotal + ppn,
    batasKmBulanan: numeric(contract.batasKmBulanan), tarifOverKm: numeric(contract.tarifOverKm), kmAwal: contract.kmAwal || asset.km || '',
    status: contract.status || 'Menunggu Dokumen', approvalStatus: contract.approvalStatus || 'Belum Diajukan', updatedAt: new Date().toISOString()
  };
}
function addDays(dateValue, days = 30) {
  const base = parseDateOnly(dateValue) || new Date();
  base.setDate(base.getDate() + Number(days || 0));
  return base.toISOString().slice(0, 10);
}
function paymentTermDays(term = '') {
  const match = String(term || '').match(/(\d+)/);
  return match ? Number(match[1]) : 30;
}
function invoiceDefaults(state = {}) {
  const contract = firstContract(state);
  const asset = (state.assets || []).find((item) => item.id === contract.asetId) || firstAsset(state);
  const baseAmount = numeric(contract.nilaiBulanan);
  const subtotal = baseAmount;
  const taxRate = 11;
  const taxAmount = Math.round(subtotal * taxRate / 100);
  const amount = subtotal + taxAmount;
  return {
    id: uid('INV'), contractId: contract.id || '', customer: contract.customer || '', customerPic: contract.picCustomer || '', customerEmail: contract.emailCustomer || '', customerPhone: contract.teleponCustomer || '', billingAddress: contract.alamatCustomer || '', npwpCustomer: contract.npwpCustomer || '',
    assetId: asset.id || '', asset: asset.nama || '', nopol: asset.nomorPolisi || asset.nomorRangka || '', poCustomer: contract.nomorPO || '', invoiceDate: today, periode: new Date().toLocaleDateString('id-ID', { month: 'long', year: 'numeric' }), periodStart: today, periodEnd: today,
    paymentTerm: contract.terminPembayaran || '30 hari', dueDate: addDays(today, paymentTermDays(contract.terminPembayaran || '30 hari')), baseAmount, overKmCharge: 0, fuelCharge: 0, penaltyCharge: 0, serviceCharge: 0, otherCharge: 0, discount: 0,
    subtotal, taxRate, taxAmount, withholdingRate: 0, withholdingAmount: 0, amount, paid: 0, remaining: amount, paymentDate: '', paymentMethod: 'Transfer Bank', paymentReference: '', bankAccount: '', taxInvoiceNumber: '', status: 'Unpaid', notes: '', invoiceFileName: '', invoiceFileData: '', taxInvoiceFileName: '', taxInvoiceFileData: '', updatedAt: new Date().toISOString()
  };
}
function normalizeInvoice(invoice = {}, state = {}) {
  const contracts = state.contracts || [];
  const assets = state.assets || [];
  const contract = contracts.find((item) => item.id === invoice.contractId) || {};
  const asset = assets.find((item) => item.id === (invoice.assetId || contract.asetId)) || {};
  const isLegacy = invoice.baseAmount === undefined && invoice.taxRate === undefined && invoice.subtotal === undefined;
  const baseAmount = numeric(invoice.baseAmount ?? (isLegacy ? invoice.amount : contract.nilaiBulanan));
  const overKmCharge = numeric(invoice.overKmCharge);
  const fuelCharge = numeric(invoice.fuelCharge);
  const penaltyCharge = numeric(invoice.penaltyCharge);
  const serviceCharge = numeric(invoice.serviceCharge);
  const otherCharge = numeric(invoice.otherCharge);
  const discount = numeric(invoice.discount);
  const subtotal = Math.max(0, numeric(invoice.subtotal) || (baseAmount + overKmCharge + fuelCharge + penaltyCharge + serviceCharge + otherCharge - discount));
  const taxRate = invoice.taxRate === undefined ? (isLegacy ? 0 : 11) : Number(invoice.taxRate || 0);
  const taxAmount = invoice.taxAmount === undefined ? Math.round(subtotal * taxRate / 100) : numeric(invoice.taxAmount);
  const withholdingRate = Number(invoice.withholdingRate || 0);
  const withholdingAmount = invoice.withholdingAmount === undefined ? Math.round(subtotal * withholdingRate / 100) : numeric(invoice.withholdingAmount);
  const amount = isLegacy ? numeric(invoice.amount) : (numeric(invoice.amount) || subtotal + taxAmount);
  const paid = numeric(invoice.paid);
  const remaining = Math.max(0, amount - paid - withholdingAmount);
  const status = remaining <= 0 ? 'Paid' : paid > 0 || withholdingAmount > 0 ? 'Partial' : (invoice.status || 'Unpaid');
  return {
    ...invoice, id: invoice.id || uid('INV'), contractId: invoice.contractId || contract.id || '', customer: invoice.customer || contract.customer || '', customerPic: invoice.customerPic || contract.picCustomer || '', customerEmail: invoice.customerEmail || contract.emailCustomer || '', customerPhone: invoice.customerPhone || contract.teleponCustomer || '', billingAddress: invoice.billingAddress || contract.alamatCustomer || '', npwpCustomer: invoice.npwpCustomer || contract.npwpCustomer || '',
    assetId: asset.id || invoice.assetId || contract.asetId || '', asset: asset.nama || invoice.asset || contract.aset || '', nopol: asset.nomorPolisi || asset.nomorRangka || invoice.nopol || contract.nopol || '', poCustomer: invoice.poCustomer || contract.nomorPO || '', invoiceDate: invoice.invoiceDate || (invoice.dueDate ? addDays(invoice.dueDate, -30) : today), periode: invoice.periode || '', periodStart: invoice.periodStart || '', periodEnd: invoice.periodEnd || '', paymentTerm: invoice.paymentTerm || contract.terminPembayaran || '30 hari', dueDate: invoice.dueDate || addDays(invoice.invoiceDate || today, paymentTermDays(invoice.paymentTerm || contract.terminPembayaran || '30 hari')),
    baseAmount, overKmCharge, fuelCharge, penaltyCharge, serviceCharge, otherCharge, discount, subtotal, taxRate, taxAmount, withholdingRate, withholdingAmount, amount, paid, remaining, paymentDate: invoice.paymentDate || '', paymentMethod: invoice.paymentMethod || 'Transfer Bank', paymentReference: invoice.paymentReference || '', bankAccount: invoice.bankAccount || '', taxInvoiceNumber: invoice.taxInvoiceNumber || '', status, notes: invoice.notes || '', invoiceFileName: invoice.invoiceFileName || '', invoiceFileData: invoice.invoiceFileData || '', taxInvoiceFileName: invoice.taxInvoiceFileName || '', taxInvoiceFileData: invoice.taxInvoiceFileData || '', updatedAt: new Date().toISOString()
  };
}

function inventoryCode(name = '') {
  const normalized = String(name).trim().toLowerCase();
  const aliases = { oli: 'OLI', ban: 'BAN', charger: 'CHG', aki: 'AKI', filter: 'FLT', rem: 'REM', brake: 'REM', lampu: 'LMP', adapter: 'ADP', baterai: 'BAT', battery: 'BAT' };
  const firstWord = normalized.split(/\s+/).filter(Boolean)[0] || 'ITEM';
  return aliases[firstWord] || firstWord.replace(/[^a-z0-9]/g, '').slice(0, 3).toUpperCase() || 'ITEM';
}
function nextInventoryId(state = {}, name = '') {
  const code = inventoryCode(name);
  const next = nextSuffix(state.inventory || [], new RegExp(`^SP-${code}-(\\d+)$`, 'i'));
  return `SP-${code}-${padSequence(next)}`;
}
function generateAssetId(kategori = 'Kendaraan') { const kode = kategori === 'Kendaraan' ? 'KND' : kategori === 'Perangkat IT' ? 'IT' : 'AST'; const unique = `${Date.now().toString().slice(-5)}${Math.floor(100 + Math.random() * 899)}`.slice(-7); return `AST-${kode}-${new Date().getFullYear()}-${unique}`; }
function booleanValue(value) { if (typeof value === 'boolean') return value; const text = String(value || '').toLowerCase().trim(); return ['ya', 'y', 'yes', 'true', 'ada', 'aktif', 'active', 'online', '1', 'terpasang'].some((item) => text.includes(item)); }
function extractKm(asset = {}) { const direct = asset.km ?? asset.odometerKm ?? asset.odometer ?? asset.kilometer; if (direct !== undefined && direct !== null && String(direct).trim() !== '') return String(direct).replace(/\s*km$/i, '').trim(); const match = String(asset.telemetri || '').match(/([0-9][0-9.,]*)\s*km/i); return match ? match[1] : ''; }
function assetTelemetry(asset = {}) { const gps = booleanValue(asset.gpsAktif) ? `GPS aktif${asset.gpsDeviceId ? ` · ${asset.gpsDeviceId}` : ''}` : 'GPS belum aktif'; const km = extractKm(asset) ? `${extractKm(asset)} km` : 'KM belum diisi'; return `${gps} · ${km}`; }
function normalizeDocument(doc = {}) { return { id: String(doc.id || '').trim(), asetId: doc.asetId || '', asetNama: doc.asetNama || '', nopol: doc.nopol || '', nama: doc.nama || doc.fileName || 'Dokumen Aset', tipe: doc.tipe || 'Dokumen Aset', nomorDokumen: doc.nomorDokumen || '', tanggalTerbit: doc.tanggalTerbit || '', expired: doc.expired || '', pemilik: doc.pemilik || 'Asset Team', status: doc.status || 'Valid', keterangan: doc.keterangan || '', file: doc.file || doc.fileName || '', fileName: doc.fileName || doc.file || '', fileType: doc.fileType || '', fileData: doc.fileData || '', uploadedAt: doc.uploadedAt || new Date().toISOString() }; }
function numeric(value) { if (typeof value === 'number') return value; const cleaned = String(value || '').replace(/[^0-9-]/g, ''); return Number(cleaned || 0); }
function parseDateOnly(value) { if (!value) return null; const parsed = new Date(`${value}T00:00:00`); if (!Number.isNaN(parsed.getTime())) return parsed; const fallback = new Date(value); return Number.isNaN(fallback.getTime()) ? null : fallback; }
function daysUntil(value) { const target = parseDateOnly(value); if (!target) return null; const start = new Date(); start.setHours(0,0,0,0); target.setHours(0,0,0,0); return Math.ceil((target.getTime() - start.getTime()) / 86400000); }
function overdueLevel(date, redDays = 30, amberDays = 60) { const days = daysUntil(date); if (days === null) return { level: 'slate', label: 'Tidak ada tanggal', days }; if (days < 0) return { level: 'red', label: `Lewat ${Math.abs(days)} hari`, days }; if (days <= redDays) return { level: 'red', label: `${days} hari lagi`, days }; if (days <= amberDays) return { level: 'amber', label: `${days} hari lagi`, days }; return { level: 'green', label: `${days} hari lagi`, days }; }
function stnkInfo(asset = {}) { if (asset.kategori && asset.kategori !== 'Kendaraan') return { due: false, level: 'slate', label: 'Tidak wajib', message: 'STNK hanya untuk kendaraan', days: null }; const info = overdueLevel(asset.tanggalSTNK, 30, 60); if (info.days === null) return { due: true, level: 'amber', label: 'STNK belum diisi', message: 'Lengkapi tanggal masa berlaku STNK', days: null }; if (info.days < 0) return { due: true, level: 'red', label: 'STNK expired', message: `STNK lewat ${Math.abs(info.days)} hari. Segera perpanjang.`, days: info.days }; if (info.days <= 30) return { due: true, level: 'red', label: 'Perpanjang STNK', message: `STNK habis dalam ${info.days} hari. Perpanjang maksimal bulan ini.`, days: info.days }; if (info.days <= 60) return { due: false, level: 'amber', label: 'Pantau STNK', message: `STNK habis dalam ${info.days} hari. Siapkan dokumen.`, days: info.days }; return { due: false, level: 'green', label: 'STNK aman', message: `STNK masih ${info.days} hari.`, days: info.days }; }
function riskClass(risk) { if (Number(risk) >= 70) return 'red'; if (Number(risk) >= 35) return 'amber'; return 'green'; }
function normalizeHeader(header) { return String(header || '').toLowerCase().trim().replace(/[._-]+/g, ' ').replace(/\s+/g, ' '); }
function findValue(row, names) { const normalized = Object.entries(row || {}).reduce((acc, [key, value]) => { acc[normalizeHeader(key)] = value; return acc; }, {}); for (const name of names) { const value = normalized[normalizeHeader(name)]; if (value !== undefined && value !== null && String(value).trim() !== '') return value; } return ''; }
function formatDateValue(value) { if (!value) return ''; if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10); if (typeof value === 'number') { const parsed = XLSX.SSF.parse_date_code(value); if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`; } const text = String(value).trim(); const parsed = new Date(text); if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10); return text; }
function formatBpkb(value) { const text = String(value || '').toLowerCase().trim(); if (['ya', 'y', 'yes', 'true', 'ada', '1', 'memiliki'].includes(text)) return 'Ya'; if (['tidak', 'no', 'false', '0', 'belum', 'tidak ada'].includes(text)) return 'Tidak'; return value ? String(value) : 'Belum diisi'; }
function normalizeAsset(asset = {}) { const kategori = asset.kategori || 'Kendaraan'; const id = String(asset.id || '').trim() || generateAssetId(kategori); const tanggalSTNKValue = asset.tanggalSTNK || asset.tanggalStnk || ''; const stnk = stnkInfo({ kategori, tanggalSTNK: tanggalSTNKValue }); const risikoAwal = Number(asset.risiko ?? 10); const risikoFinal = kategori === 'Kendaraan' && stnk.due ? Math.max(risikoAwal, stnk.level === 'red' ? 82 : 45) : risikoAwal; const gpsAktif = asset.gpsAktif !== undefined ? booleanValue(asset.gpsAktif) : /gps\s*(online|aktif|active|terpasang)/i.test(asset.telemetri || ''); const km = extractKm(asset); const baseAsset = { id, nama: asset.nama || 'Aset Baru', kategori, tipe: asset.tipe || 'Aset Rental', status: asset.status || 'Tersedia', kondisi: asset.kondisi || 'Baik', cabang: asset.cabang || asset.wilayah || '-', lokasi: asset.lokasi || asset.wilayah || '-', pemegang: asset.pemegang || asset.perusahaanUser || 'Belum ditugaskan', perusahaanUser: asset.perusahaanUser || asset.pemegang || 'Belum diisi', alamatUser: asset.alamatUser || '-', wilayah: asset.wilayah || asset.cabang || '-', warna: asset.warna || '-', tahunKendaraan: asset.tahunKendaraan || asset.tahun || '-', nomorRangka: asset.nomorRangka || '-', nomorMesin: asset.nomorMesin || '-', nomorPolisi: asset.nomorPolisi || '-', qrCode: asset.qrCode || `QR-AERIZEN-${id}`, tanggalSTNK: tanggalSTNKValue, tanggalKIR: asset.tanggalKIR || '', tanggalPajak: asset.tanggalPajak || '', tanggalAsuransi: asset.tanggalAsuransi || '', tanggalServiceBerikutnya: asset.tanggalServiceBerikutnya || '', memilikiBPKB: asset.memilikiBPKB || asset.memilikiBpkb || 'Belum diisi', gpsAktif, gpsDeviceId: asset.gpsDeviceId || '', gpsProvider: asset.gpsProvider || '', gpsLastUpdate: asset.gpsLastUpdate || '', km, odometerKm: km, gambar: asset.gambar || '', risiko: risikoFinal, utilisasi: Number(asset.utilisasi ?? 0), nilaiBuku: numeric(asset.nilaiBuku), pendapatanBulanan: numeric(asset.pendapatanBulanan || asset.pendapatan), biayaBulanan: numeric(asset.biayaBulanan || asset.biaya), aksiBerikutnya: asset.aksiBerikutnya || (kategori === 'Kendaraan' && stnk.due ? stnk.message : 'Belum ada aksi berikutnya'), dokumen: asset.dokumen || `STNK: ${tanggalSTNKValue || 'Belum diisi'} · BPKB: ${asset.memilikiBPKB || asset.memilikiBpkb || 'Belum diisi'}`, alur: asset.alur || 'Tersedia → Dipesan → Digunakan → Kembali', riwayat: Array.isArray(asset.riwayat) && asset.riwayat.length ? asset.riwayat : ['Terdaftar'], updatedAt: new Date().toISOString() }; return { ...baseAsset, telemetri: asset.telemetri || assetTelemetry(baseAsset) }; }
function normalizeWorkOrder(wo = {}, asset) {
  const biayaJasa = numeric(wo.biayaJasa);
  const biayaSparepart = numeric(wo.biayaSparepart);
  const biayaLain = numeric(wo.biayaLain);
  const pajak = numeric(wo.pajak);
  const componentTotal = biayaJasa + biayaSparepart + biayaLain + pajak;
  const biaya = numeric(wo.biaya ?? wo.estimasiBiaya) || componentTotal;
  return {
    id: wo.id || uid('WO'), asetId: wo.asetId || asset?.id || '', namaAset: wo.namaAset || asset?.nama || '', nopol: wo.nopol || asset?.nomorPolisi || '-', kilometer: wo.kilometer || asset?.km || '',
    jenis: wo.jenis || 'Preventive Maintenance', prioritas: wo.prioritas || 'Sedang', status: wo.status || 'Draft', pic: wo.pic || 'Belum ditugaskan', teknisi: wo.teknisi || '', vendorId: wo.vendorId || '', lokasiPekerjaan: wo.lokasiPekerjaan || asset?.lokasi || '',
    tanggalMulai: wo.tanggalMulai || today, jatuhTempo: wo.jatuhTempo || today, tanggalSelesai: wo.tanggalSelesai || (['Selesai', 'Ditutup'].includes(wo.status) ? today : ''), downtimeJam: numeric(wo.downtimeJam),
    estimasiBiaya: numeric(wo.estimasiBiaya), biayaJasa, biayaSparepart, biayaLain, pajak, biaya,
    sparepart: wo.sparepart || '', noPO: wo.noPO || '', noInvoice: wo.noInvoice || '', warrantyStatus: wo.warrantyStatus || 'Tidak', approvalStatus: wo.approvalStatus || 'Belum Diajukan',
    fotoBefore: wo.fotoBefore || wo.fotoBeforeName || 'Belum diunggah', fotoBeforeName: wo.fotoBeforeName || '', fotoBeforeData: wo.fotoBeforeData || '',
    fotoAfter: wo.fotoAfter || wo.fotoAfterName || 'Belum diunggah', fotoAfterName: wo.fotoAfterName || '', fotoAfterData: wo.fotoAfterData || '',
    lampiranName: wo.lampiranName || '', lampiranData: wo.lampiranData || '',
    keluhan: wo.keluhan || '', diagnosa: wo.diagnosa || '', tindakan: wo.tindakan || '', checklist: wo.checklist || 'Inspeksi awal, foto evidence, update progress, QC akhir', hasilQC: wo.hasilQC || '', catatan: wo.catatan || '', updatedAt: new Date().toISOString(),
  };
}
function detectHeaderRows(worksheet) { const grid = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false }); const headerIndex = grid.findIndex((row) => row.some((cell) => normalizeHeader(cell).includes('nama kendaraan'))); if (headerIndex >= 0) { const headers = grid[headerIndex].map((cell) => String(cell || '').trim()); return grid.slice(headerIndex + 1).filter((row) => row.some((cell) => String(cell || '').trim() !== '')).map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? '']))); } return XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false }); }
function assetFromImportRow(row, index) { const no = findValue(row, ['no', 'nomor', 'no.']); const nama = String(findValue(row, ['nama kendaraan', 'kendaraan', 'nama aset', 'nama']) || `Kendaraan Sewa ${index + 1}`); const nomorRangka = String(findValue(row, ['nomor rangka', 'no rangka', 'vin']) || '-'); const nomorPolisi = String(findValue(row, ['nomor polisi', 'no polisi', 'nopol', 'plat nomor']) || '-'); const perusahaanUser = String(findValue(row, ['perusahaan user', 'perusahaan', 'customer']) || 'Belum diisi'); const wilayah = String(findValue(row, ['wilayah', 'area', 'region', 'cabang']) || '-'); const tanggalSTNK = formatDateValue(findValue(row, ['tanggal STNK', 'tanggal stnk', 'stnk', 'masa berlaku stnk', 'expired stnk'])); return normalizeAsset({ nama, kategori: 'Kendaraan', tipe: 'Aset Rental', status: 'Disewakan', kondisi: 'Baik', cabang: wilayah, lokasi: wilayah, pemegang: perusahaanUser, perusahaanUser, alamatUser: String(findValue(row, ['alamat user', 'alamat']) || '-'), wilayah, warna: String(findValue(row, ['warna']) || '-'), tahunKendaraan: String(findValue(row, ['tahun kendaraan', 'tahun']) || '-'), nomorRangka, nomorMesin: String(findValue(row, ['nomor mesin', 'no mesin']) || '-'), nomorPolisi, tanggalSTNK, memilikiBPKB: formatBpkb(findValue(row, ['memiliki BPKB', 'bpkb'])), tanggalKIR: formatDateValue(findValue(row, ['tanggal KIR', 'kir'])), tanggalPajak: formatDateValue(findValue(row, ['tanggal pajak', 'pajak'])), tanggalAsuransi: formatDateValue(findValue(row, ['tanggal asuransi', 'asuransi'])), gpsAktif: booleanValue(findValue(row, ['gps', 'gps aktif', 'memiliki gps', 'terpasang gps'])), gpsDeviceId: String(findValue(row, ['id gps', 'gps id', 'device gps', 'imei gps']) || ''), gpsProvider: String(findValue(row, ['provider gps', 'vendor gps', 'gps provider']) || ''), km: String(findValue(row, ['km', 'odometer', 'kilometer', 'kilometer kendaraan']) || ''), utilisasi: 80, risiko: tanggalSTNK ? 18 : 45, riwayat: ['Import Excel', `Baris ${no || index + 1}`, 'Validasi Kendaraan Sewa'] }); }

function asArray(value, fallback = []) { return Array.isArray(value) ? value : fallback; }
const defaultSettings = { supabaseUrl: import.meta.env.VITE_SUPABASE_URL || '', supabaseKey: import.meta.env.VITE_SUPABASE_ANON_KEY || '', workspace: 'aerizen-main', accountEmail: '', accountPassword: '', loggedInUser: '', role: 'Super Admin', realtime: true, autoSync: true };
function loadState() {
  const stored = safeParse(localStorage.getItem(STORAGE_KEY), null);
  if (!stored || typeof stored !== 'object') return initialState;
  const migratedAssets = asArray(stored.assets, initialState.assets).map((asset) => normalizeAsset(asset));
  return {
    ...initialState,
    ...stored,
    assets: migratedAssets,
    workOrders: asArray(stored.workOrders, initialState.workOrders).map((wo) => normalizeWorkOrder(wo, migratedAssets.find((asset) => asset.id === wo.asetId))),
    contracts: asArray(stored.contracts, initialState.contracts),
    invoices: asArray(stored.invoices, initialState.invoices).map((invoice) => normalizeInvoice(invoice, { contracts: asArray(stored.contracts, initialState.contracts), assets: migratedAssets })),
    checklists: asArray(stored.checklists, initialState.checklists),
    inventory: asArray(stored.inventory, initialState.inventory),
    vendors: asArray(stored.vendors, initialState.vendors),
    tickets: asArray(stored.tickets, initialState.tickets),
    documents: asArray(stored.documents, initialState.documents).map((doc) => normalizeDocument(doc)),
    approvals: asArray(stored.approvals, initialState.approvals),
    roles: asArray(stored.roles, initialState.roles),
    auditLogs: asArray(stored.auditLogs, initialState.auditLogs),
  };
}
function loadSettings() {
  const stored = safeParse(localStorage.getItem(SETTINGS_KEY), {});
  return { ...defaultSettings, ...(stored && typeof stored === 'object' ? stored : {}) };
}

function useLocalState() {
  const [state, setState] = useState(loadState);
  const [settings, setSettings] = useState(loadSettings);
  const [queue, setQueue] = useState(() => asArray(safeParse(localStorage.getItem(QUEUE_KEY), []), []));
  useEffect(() => localStorage.setItem(STORAGE_KEY, JSON.stringify(state)), [state]);
  useEffect(() => localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings)), [settings]);
  useEffect(() => localStorage.setItem(QUEUE_KEY, JSON.stringify(queue)), [queue]);
  return { state, setState, settings, setSettings, queue, setQueue };
}

function App() {
  const { state, setState, settings, setSettings, queue, setQueue } = useLocalState();
  const [active, setActive] = useState(() => { const rawHash = window.location.hash.replace('#', ''); const hash = rawHash === 'maintenance' ? 'workorders' : rawHash; return tabs.some(([key]) => key === hash) ? hash : 'dashboard'; });
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('Semua');
  const [selectedAsset, setSelectedAsset] = useState(null);
  const [assetForm, setAssetForm] = useState(null);
  const [workOrderForm, setWorkOrderForm] = useState(null);
  const [contractForm, setContractForm] = useState(null);
  const [documentForm, setDocumentForm] = useState(null);
  const [quickAddForm, setQuickAddForm] = useState(null);
  const [notice, setNotice] = useState('');
  const [syncStatus, setSyncStatus] = useState('Offline + Supabase siap. Data aman di storage lokal.');
  const [desktop, setDesktop] = useState(null);
  const [notificationOpen, setNotificationOpen] = useState(false);
  const channelRef = useRef(null);

  const pageMeta = {
    dashboard: ['Command center', 'Dashboard Enterprise', 'Ringkasan risiko, STNK, profit, SLA, kontrak, ticket, dan notifikasi utama.'],
    assets: ['Master data', 'Data Aset', 'Kelola aset, foto, dokumen, GPS, KM/odometer, import Excel, QR aset, nilai, profit, dan Aset 360.'],
    contracts: ['Rental lifecycle', 'Kontrak Sewa', 'Kontrak lengkap: customer, PIC, aset/nopol, periode, biaya, PO, SLA, coverage, dokumen, approval, dan renewal.'],
    operations: ['Field operation', 'Check-in / Check-out', 'Serah-terima, inspeksi kendaraan/perangkat, foto before-after, aksesoris, dan tanda tangan digital.'],
    workorders: ['Work order', 'Work Order (WO)', 'WO tetap memakai kode otomatis WO-xxxx, Nominal Biaya Aktual, history per nopol, filter nopol, Export Excel, foto before/after, sparepart, vendor, dan klik Selesai.'],
    finance: ['Billing', 'Finance & Invoice', 'Invoice bulanan, nominal amount/paid/sisa, status pembayaran, profit aset, revenue customer, tombol lunas, Tambah Invoice, dan Export Excel.'],
    inventory: ['Sparepart', 'Inventory Sparepart', 'Stok oli, ban, aki, charger, adapter, minimum stock, vendor, dan nilai persediaan.'],
    vendors: ['Supplier performance', 'Vendor Management', 'Database bengkel, supplier IT, rating, biaya rata-rata, lead time, dan performa vendor.'],
    tickets: ['SLA support', 'Ticket & Complaint', 'Keluhan customer/internal, prioritas, SLA, PIC, status, dan tindak lanjut.'],
    documents: ['Digital vault', 'Dokumen Aset', 'ID dokumen manual dan keterkaitan aset melalui nopol/serial agar file otomatis muncul pada Aset 360.'],
    approvals: ['Control tower', 'Approval & Permission', 'Role, permission, approval matrix, limit biaya, pending approval, dan kontrol akses user.'],
    reports: ['Audit & export', 'Laporan', 'Audit trail, backup/restore, export Excel, dan laporan management.'],
    account: ['Autentikasi', 'Akun', 'Login/logout, lupa password, lihat password, role aktif, dan reset password Supabase.'],
    configuration: ['Integrasi', 'Konfigurasi', 'Offline-first, Supabase realtime, workspace, queue sync, SQL, dan reset data demo.'],
  };

  const supabase = useMemo(() => { if (!settings.supabaseUrl || !settings.supabaseKey) return null; try { return createClient(settings.supabaseUrl, settings.supabaseKey); } catch { return null; } }, [settings.supabaseUrl, settings.supabaseKey]);

  useEffect(() => { if (window.aerizenDesktop?.platform) window.aerizenDesktop.platform().then(setDesktop).catch(() => setDesktop(null)); }, []);
  useEffect(() => { const onHashChange = () => { const rawHash = window.location.hash.replace('#', ''); const hash = rawHash === 'maintenance' ? 'workorders' : rawHash; if (tabs.some(([key]) => key === hash)) setActive(hash); }; window.addEventListener('hashchange', onHashChange); return () => window.removeEventListener('hashchange', onHashChange); }, []);
  useEffect(() => { if (window.location.hash !== `#${active}`) window.location.hash = active; }, [active]);

  const addLog = (aksi, detail) => setState((prev) => ({ ...prev, auditLogs: [{ id: uid('LOG'), waktu: new Date().toISOString(), aksi, detail, user: settings.loggedInUser || 'Admin Aerizen' }, ...(prev.auditLogs || [])].slice(0, 200) }));
  const queueSync = (collection, item, action = 'upsert') => setQueue((prev) => [{ id: uid('SYNC'), collection, action, payload: item, createdAt: new Date().toISOString() }, ...prev].slice(0, 500));
  const upsertLocal = (collection, item, doSync = true) => { setState((prev) => { const list = prev[collection] || []; const exists = list.some((row) => row.id === item.id); return { ...prev, [collection]: exists ? list.map((row) => row.id === item.id ? item : row) : [item, ...list] }; }); if (doSync) queueSync(collection, item, 'upsert'); };
  const deleteLocal = (collection, id) => { setState((prev) => ({ ...prev, [collection]: (prev[collection] || []).filter((row) => row.id !== id) })); queueSync(collection, { id }, 'delete'); addLog(`Delete ${collection}`, id); };

  const notifications = useMemo(() => buildNotifications(state), [state]);
  const metrics = useMemo(() => buildMetrics(state, notifications), [state, notifications]);
  const filteredAssets = useMemo(() => { const text = query.toLowerCase(); return (state.assets || []).filter((asset) => { const matchFilter = filter === 'Semua' || asset.kategori === filter || asset.status === filter || asset.cabang === filter; const matchQuery = [asset.nama, asset.id, asset.pemegang, asset.cabang, asset.nomorPolisi, asset.nomorRangka, asset.nomorMesin, asset.wilayah, asset.qrCode].join(' ').toLowerCase().includes(text); return matchFilter && matchQuery; }); }, [state.assets, query, filter]);
  const filterOptions = useMemo(() => { const dynamic = new Set(['Semua', ...categoryOptions, ...statusOptions]); (state.assets || []).forEach((asset) => dynamic.add(asset.cabang)); return Array.from(dynamic).filter(Boolean); }, [state.assets]);

  async function processSyncQueue() {
    if (!supabase) { setSyncStatus('Supabase belum dikonfigurasi. Data tetap aman di offline storage.'); return; }
    if (!navigator.onLine) { setSyncStatus('Perangkat offline. Perubahan masuk antrean sync.'); return; }
    const items = [...queue].reverse();
    if (!items.length) { setSyncStatus('Tidak ada antrean yang perlu dikirim.'); return; }
    setSyncStatus(`Mengirim ${items.length} perubahan ke Supabase...`);
    for (const item of items) {
      const row = { id: `${item.collection}:${item.payload.id}`, collection: item.collection, payload: item.payload, updated_at: new Date().toISOString(), deleted_at: item.action === 'delete' ? new Date().toISOString() : null };
      const { error } = await supabase.from('aerizen_records').upsert(row);
      if (error) { setSyncStatus(`Gagal sinkron ${item.payload.id}: ${error.message}`); return; }
    }
    setQueue([]); setSyncStatus(`${items.length} perubahan berhasil dikirim ke Supabase.`);
  }
  async function pullSupabase() { if (!supabase) { setSyncStatus('Isi Supabase URL dan anon key dulu.'); return; } const { data, error } = await supabase.from('aerizen_records').select('*').is('deleted_at', null).order('updated_at', { ascending: false }); if (error) { setSyncStatus(`Gagal pull: ${error.message}`); return; } const next = { ...state }; for (const row of data || []) { const list = next[row.collection] || []; const exists = list.some((item) => item.id === row.payload.id); next[row.collection] = exists ? list.map((item) => item.id === row.payload.id ? row.payload : item) : [row.payload, ...list]; } setState(next); setSyncStatus(`${data?.length || 0} record dibaca dari Supabase.`); }
  async function loginAccount() { if (!settings.accountEmail || !settings.accountPassword) { setNotice('Isi email akun dan password terlebih dahulu.'); return; } if (!supabase) { setSettings((prev) => ({ ...prev, loggedInUser: prev.accountEmail })); setNotice('Supabase belum diatur. Login disimpan lokal untuk demo.'); return; } const { data, error } = await supabase.auth.signInWithPassword({ email: settings.accountEmail, password: settings.accountPassword }); if (error) { setNotice(`Login gagal: ${error.message}`); return; } setSettings((prev) => ({ ...prev, loggedInUser: data.user?.email || prev.accountEmail, accountPassword: '' })); setNotice(`Login berhasil sebagai ${data.user?.email || settings.accountEmail}.`); }
  async function logoutAccount() { if (supabase) await supabase.auth.signOut().catch(() => null); setSettings((prev) => ({ ...prev, loggedInUser: '', accountPassword: '' })); setNotice('Akun berhasil logout.'); }
  async function sendPasswordReset(email) { const targetEmail = String(email || settings.accountEmail || '').trim(); if (!targetEmail) { setNotice('Isi email akun terlebih dahulu untuk mengirim link lupa password.'); return; } if (!supabase) { setNotice('Fitur lupa password membutuhkan Supabase URL dan anon key di menu Konfigurasi.'); return; } const redirectTo = window.location.protocol.startsWith('http') ? `${window.location.origin}${window.location.pathname}#account` : undefined; const { error } = await supabase.auth.resetPasswordForEmail(targetEmail, redirectTo ? { redirectTo } : undefined); if (error) { setNotice(`Gagal mengirim link reset password: ${error.message}`); return; } setNotice(`Link reset password berhasil dikirim ke ${targetEmail}. Cek inbox/spam email tersebut.`); }
  async function updateAccountPassword(newPassword, confirmPassword) { if (!newPassword || !confirmPassword) { setNotice('Isi password baru dan konfirmasi password.'); return; } if (newPassword.length < 6) { setNotice('Password baru minimal 6 karakter.'); return; } if (newPassword !== confirmPassword) { setNotice('Konfirmasi password tidak sama.'); return; } if (!supabase) { setNotice('Update password membutuhkan Supabase. Isi konfigurasi Supabase terlebih dahulu.'); return; } const { data, error } = await supabase.auth.updateUser({ password: newPassword }); if (error) { setNotice(`Gagal update password: ${error.message}. Buka link reset dari email terlebih dahulu, lalu coba lagi.`); return; } setSettings((prev) => ({ ...prev, loggedInUser: data.user?.email || prev.loggedInUser, accountPassword: '' })); setNotice('Password akun berhasil diperbarui. Silakan login ulang jika sesi berakhir.'); }

  useEffect(() => { if (!supabase || !settings.realtime) return undefined; if (channelRef.current) supabase.removeChannel(channelRef.current); const channel = supabase.channel('aerizen-records-realtime').on('postgres_changes', { event: '*', schema: 'public', table: 'aerizen_records' }, (payload) => { const nextRecord = payload.new; if (!nextRecord?.payload || nextRecord.deleted_at) return; upsertLocal(nextRecord.collection, nextRecord.payload, false); setSyncStatus(`Realtime update: ${nextRecord.collection} / ${nextRecord.payload.id}`); }).subscribe(); channelRef.current = channel; return () => supabase.removeChannel(channel); }, [supabase, settings.realtime]);
  useEffect(() => { if (!supabase) return undefined; const { data } = supabase.auth.onAuthStateChange((event, session) => { if (event === 'PASSWORD_RECOVERY') { setActive('account'); setSettings((prev) => ({ ...prev, loggedInUser: session?.user?.email || prev.loggedInUser })); setNotice('Mode reset password aktif. Masukkan password baru di halaman Akun, lalu klik Simpan Password Baru.'); } if (event === 'SIGNED_IN' && session?.user?.email) setSettings((prev) => ({ ...prev, loggedInUser: session.user.email })); }); return () => data.subscription.unsubscribe(); }, [supabase]);
  useEffect(() => { if (!settings.autoSync) return undefined; const timer = window.setInterval(() => { if (queue.length && navigator.onLine && supabase) processSyncQueue(); }, 10000); return () => window.clearInterval(timer); }, [queue, settings.autoSync, supabase]);

  async function importExcel(file) { if (!file) return; try { const buffer = await file.arrayBuffer(); const workbook = XLSX.read(buffer, { type: 'array', cellDates: true }); const worksheet = workbook.Sheets[workbook.SheetNames[0]]; const rows = detectHeaderRows(worksheet); const imported = rows.filter((row) => Object.values(row).some((v) => String(v || '').trim() !== '')).map(assetFromImportRow); if (!imported.length) { setNotice('File tidak memiliki data yang bisa diimpor. Pastikan header sesuai template.'); return; } setState((prev) => { const ids = new Set((prev.assets || []).map((item) => item.id)); const unique = imported.map((asset, index) => ids.has(asset.id) ? { ...asset, id: `${asset.id}-${index + 1}` } : asset); unique.forEach((asset) => queueSync('assets', asset, 'upsert')); return { ...prev, assets: [...unique, ...(prev.assets || [])] }; }); addLog('Import Excel Kendaraan Sewa', `${imported.length} kendaraan sewa diimpor dari ${file.name}`); setNotice(`${imported.length} kendaraan sewa berhasil diimpor dari ${file.name}.`); } catch (error) { setNotice(`Import gagal: ${error.message || 'format file tidak terbaca'}`); } }
  function saveAsset(form) { const existing = (state.assets || []).some((item) => item.id === form.id); const { assetDocuments = [], ...assetPayload } = form || {}; const asset = normalizeAsset(assetPayload); upsertLocal('assets', asset); (assetDocuments || []).forEach((doc) => { const quickId = doc.id || suggestDocumentId(asset, doc.tipe || 'Dokumen', doc.fileName || doc.file); upsertLocal('documents', normalizeDocument({ ...doc, id: quickId, asetId: asset.id, asetNama: asset.nama, nopol: asset.nomorPolisi || asset.nomorRangka || '', pemilik: settings.loggedInUser || 'Asset Team', status: doc.status || 'Valid' })); }); addLog(existing ? 'Update aset' : 'Tambah aset', `${asset.id} - ${asset.nama}${assetDocuments?.length ? ` · ${assetDocuments.length} dokumen diupload` : ''}`); setAssetForm(null); setSelectedAsset(asset); setNotice(`Aset ${asset.nama} berhasil disimpan${assetDocuments?.length ? ` dengan ${assetDocuments.length} dokumen aset` : ''}.`); }
  function saveContract(form) { const contract = normalizeContract(form, state); if (!contract.id || !contract.customer || !contract.asetId) { setNotice('Nomor kontrak, customer, dan aset wajib diisi.'); return; } upsertLocal('contracts', contract); addLog('Simpan kontrak', `${contract.id} · ${contract.customer} · ${contract.nopol}`); setContractForm(null); setNotice(`Kontrak ${contract.id} berhasil disimpan untuk ${contract.aset}.`); }
  function saveDocument(form) { const asset = (state.assets || []).find((item) => item.id === form.asetId); if (!String(form.id || '').trim()) { setNotice('ID Dokumen wajib diisi manual.'); return; } if (!asset) { setNotice('Pilih nopol/aset agar dokumen terhubung ke Aset 360.'); return; } const doc = normalizeDocument({ ...form, asetId: asset.id, asetNama: asset.nama, nopol: asset.nomorPolisi || asset.nomorRangka || '' }); upsertLocal('documents', doc); addLog('Upload dokumen aset', `${doc.id} - ${doc.nama} - ${doc.nopol}`); setDocumentForm(null); setNotice(`Dokumen ${doc.id} terhubung ke ${asset.nama} (${doc.nopol}).`); }
  function saveQuickAdd(collection, values) { const row = normalizeQuickRow(collection, values, state); upsertLocal(collection, row); addLog(`Tambah data ${collection}`, row.id || row.nama || row.customer || row.proses); setQuickAddForm(null); setNotice(`Data ${row.id || row.nama || 'baru'} berhasil ditambahkan.`); }
  function saveWorkOrder(form) { const asset = (state.assets || []).find((item) => item.id === form.asetId); const completeStatus = ['Selesai', 'Ditutup'].includes(form.status); const wo = normalizeWorkOrder({ ...form, nopol: form.nopol || asset?.nomorPolisi, tanggalSelesai: completeStatus ? (form.tanggalSelesai || today) : '' }, asset); upsertLocal('workOrders', wo); if (asset && !['Draft', 'Selesai', 'Ditutup'].includes(wo.status)) upsertLocal('assets', normalizeAsset({ ...asset, status: 'Maintenance', aksiBerikutnya: `${wo.jenis} - ${wo.status}`, riwayat: [...(asset.riwayat || []), `Work Order ${wo.id}`] })); if (asset && ['Selesai', 'Ditutup'].includes(wo.status) && ['Maintenance', 'Perbaikan', 'Overdue'].includes(asset.status)) upsertLocal('assets', normalizeAsset({ ...asset, status: 'Tersedia', kondisi: asset.kondisi === 'Perlu Servis' ? 'Baik' : asset.kondisi, aksiBerikutnya: `Work Order selesai ${wo.tanggalSelesai || today}`, riwayat: [...(asset.riwayat || []), `Work Order ${wo.id} selesai`] })); addLog('Work Order', `${wo.id} untuk ${wo.namaAset}, Nominal Biaya Aktual ${currency.format(numeric(wo.biaya))}`); setWorkOrderForm(null); setNotice(`Work Order ${wo.id} berhasil disimpan.`); }
  function completeWorkOrder(woId) { const existing = (state.workOrders || []).find((item) => item.id === woId); if (!existing) { setNotice('Work Order tidak ditemukan.'); return; } if (['Selesai', 'Ditutup'].includes(existing.status)) { setNotice(`Work Order ${woId} sudah selesai.`); return; } const asset = (state.assets || []).find((item) => item.id === existing.asetId); const completed = normalizeWorkOrder({ ...existing, status: 'Selesai', tanggalSelesai: today, fotoAfter: existing.fotoAfter || 'Ada' }, asset); upsertLocal('workOrders', completed); if (asset && ['Maintenance', 'Perbaikan', 'Overdue'].includes(asset.status)) upsertLocal('assets', normalizeAsset({ ...asset, status: 'Tersedia', kondisi: asset.kondisi === 'Perlu Servis' ? 'Baik' : asset.kondisi, aksiBerikutnya: `Work Order selesai ${today}. Siap digunakan.`, riwayat: [...(asset.riwayat || []), `Work Order ${completed.id} selesai`] })); addLog('Selesai Work Order', `${completed.id} selesai untuk ${completed.namaAset} (${completed.nopol})`); setNotice(`Work Order ${completed.id} ditandai selesai.`); }
  function markInvoicePaid(invoiceId) { const inv = (state.invoices || []).find((item) => item.id === invoiceId); if (!inv) return; const normalized = normalizeInvoice(inv, state); const paid = { ...normalized, paid: Math.max(0, numeric(normalized.amount) - numeric(normalized.withholdingAmount)), remaining: 0, paymentDate: normalized.paymentDate || today, status: 'Paid' }; upsertLocal('invoices', paid); addLog('Invoice paid', `${paid.id} dilunasi ${currency.format(paid.amount)}`); setNotice(`${paid.id} ditandai lunas.`); }
  function closeTicket(ticketId) { const ticket = (state.tickets || []).find((item) => item.id === ticketId); if (!ticket) return; upsertLocal('tickets', { ...ticket, status: 'Closed' }); addLog('Ticket closed', ticketId); setNotice(`${ticketId} ditutup.`); }
  function approveRule(approvalId) { const item = (state.approvals || []).find((row) => row.id === approvalId); if (!item) return; const pending = Math.max(0, numeric(item.pending) - 1); upsertLocal('approvals', { ...item, pending: String(pending) }); addLog('Approval', `${approvalId} disetujui`); setNotice(`${approvalId} disetujui.`); }
  function exportBackup() { const blob = new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' }); const url = URL.createObjectURL(blob); const anchor = document.createElement('a'); anchor.href = url; anchor.download = `aerizen-v5.5-backup-${today}.json`; anchor.click(); URL.revokeObjectURL(url); }
  async function restoreBackup(file) { if (!file) return; const text = await file.text(); const restored = safeParse(text, null); if (!restored?.assets) { setNotice('File backup tidak valid.'); return; } setState(restored); setNotice('Backup berhasil direstore ke offline storage.'); addLog('Restore backup', file.name); }
  function resetDemo() { setState(initialState); setQueue([]); setNotice('Data demo Aerizen v5.5 dikembalikan ke awal.'); }
  function exportCollection(collection, filename) { const rows = state[collection] || []; const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, collection); XLSX.writeFile(wb, `${filename}-${today}.xlsx`); addLog('Export Excel', filename); }
  function saveConfiguration() { setNotice('Konfigurasi berhasil disimpan.'); }

  const [eyebrow, title, description] = pageMeta[active] || pageMeta.dashboard;

  return (
    <div className="workspace-shell enterprise-v3">
      <header className="workspace-header">
        <div className="workspace-top">
          <div className="brand"><img className="brand-logo" src="./branding/aerizen-emblem-512.png" alt="Logo Aerizen" /><div><strong>Aerizen</strong><span>Asset Management System · {APP_VERSION_LABEL}</span></div></div>
          <div className="header-right"><div className="header-status"><span className={`pill ${navigator.onLine ? 'green' : 'amber'}`}>{navigator.onLine ? 'Online' : 'Offline'}</span><span className="pill blue">Queue: {queue.length}</span><span className="pill slate">Workspace: {settings.workspace || 'aerizen-main'}</span><span className="pill blue">Versi: {APP_VERSION_LABEL}</span><span className="pill purple">Role: {settings.role || 'Super Admin'}</span><span className="muted">{desktop?.packaged ? 'EXE mode' : 'Development mode'}</span></div></div>
        </div>
        <div className="nav-frame nav-frame-with-bell"><nav className="top-nav enterprise-nav">{tabs.map(([key, label]) => <button key={key} className={active === key ? 'active' : ''} onClick={() => setActive(key)}>{label}</button>)}</nav><NotificationBell notifications={notifications} open={notificationOpen} setOpen={setNotificationOpen} setActive={setActive} /></div>
      </header>
      <main className="workspace-main">
        <div className="page-header"><div><p className="eyebrow">{eyebrow}</p><h1>{title}</h1><p className="page-description">{description}</p></div><div className="top-actions">{active === 'assets' && <button className="btn primary" onClick={() => setAssetForm({})}>Tambah Aset</button>}{active === 'workorders' && <button className="btn primary" onClick={() => setWorkOrderForm({ asetId: state.assets?.[0]?.id || '', namaAset: state.assets?.[0]?.nama || '' })}>Buat WO</button>}{active === 'documents' && <button className="btn primary" onClick={() => setDocumentForm({ asetId: state.assets?.[0]?.id || '' })}>Upload Dokumen Aset</button>}{active === 'contracts' && <button className="btn primary" onClick={() => setContractForm(contractDefaults(state))}>Tambah Kontrak</button>}{getQuickCollection(active) && active !== 'workorders' && active !== 'documents' && active !== 'contracts' && <button className="btn primary" onClick={() => setQuickAddForm({ collection: getQuickCollection(active), values: quickDefaults(getQuickCollection(active), state) })}>{quickModuleConfig(getQuickCollection(active)).button}</button>}{active === 'reports' && <button className="btn ghost" onClick={exportBackup}>Export Backup</button>}{active === 'reports' && <label className="btn ghost">Restore JSON<input type="file" accept=".json" onChange={(e) => restoreBackup(e.target.files?.[0])} hidden /></label>}{active === 'configuration' && <button className="btn primary" onClick={processSyncQueue}>Sinkronkan Sekarang</button>}</div></div>
        {notice && <div className="notice"><span>{notice}</span><button onClick={() => setNotice('')}>×</button></div>}
        {active === 'dashboard' && <Dashboard metrics={metrics} state={state} notifications={notifications} onOpenNotifications={() => setNotificationOpen(true)} />}
        {active === 'assets' && <AssetsView assets={filteredAssets} contracts={state.contracts} invoices={state.invoices} documents={state.documents} query={query} setQuery={setQuery} filter={filter} setFilter={setFilter} filterOptions={filterOptions} importExcel={importExcel} onDetail={setSelectedAsset} onEdit={setAssetForm} onWO={(a) => setWorkOrderForm({ asetId: a.id, namaAset: a.nama })} onDelete={(id) => deleteLocal('assets', id)} />}
        {active === 'contracts' && <ContractsView contracts={state.contracts || []} assets={state.assets || []} onAdd={() => setContractForm(contractDefaults(state))} onEdit={setContractForm} exportExcel={() => exportCollection('contracts', 'kontrak-sewa')} />}
        {active === 'operations' && <OperationsView checklists={state.checklists || []} assets={state.assets || []} onAdd={() => setQuickAddForm({ collection: 'checklists', values: quickDefaults('checklists', state) })} onExport={() => exportCollection('checklists', 'check-in-out-inspection')} />}
        {active === 'workorders' && <WorkOrderView workOrders={state.workOrders || []} assets={state.assets || []} vendors={state.vendors || []} onCreate={(asset) => setWorkOrderForm({ asetId: asset.id, namaAset: asset.nama })} onEdit={setWorkOrderForm} onComplete={completeWorkOrder} />}
        {active === 'finance' && <FinanceView invoices={state.invoices || []} assets={state.assets || []} contracts={state.contracts || []} onAdd={() => setQuickAddForm({ collection: 'invoices', values: invoiceDefaults(state), isEditing: false })} onEdit={(invoice) => setQuickAddForm({ collection: 'invoices', values: normalizeInvoice(invoice, state), isEditing: true })} onPaid={markInvoicePaid} onExport={() => exportCollection('invoices', 'billing-invoice')} />}
        {active === 'inventory' && <InventoryView inventory={state.inventory || []} onAdd={() => setQuickAddForm({ collection: 'inventory', values: quickDefaults('inventory', state) })} onExport={() => exportCollection('inventory', 'inventory-sparepart')} />}
        {active === 'vendors' && <VendorView vendors={state.vendors || []} workOrders={state.workOrders || []} onAdd={() => setQuickAddForm({ collection: 'vendors', values: quickDefaults('vendors', state) })} onExport={() => exportCollection('vendors', 'vendor-management')} />}
        {active === 'tickets' && <TicketView tickets={state.tickets || []} onAdd={() => setQuickAddForm({ collection: 'tickets', values: quickDefaults('tickets', state) })} onCloseTicket={closeTicket} onExport={() => exportCollection('tickets', 'ticket-sla')} />}
        {active === 'documents' && <DocumentsView documents={state.documents || []} assets={state.assets || []} onAdd={() => setDocumentForm({ asetId: state.assets?.[0]?.id || '' })} onExport={() => exportCollection('documents', 'dokumen-aset')} />}
        {active === 'approvals' && <ApprovalsView approvals={state.approvals || []} roles={state.roles || []} onAdd={() => setQuickAddForm({ collection: 'approvals', values: quickDefaults('approvals', state) })} onApprove={approveRule} onExport={() => exportCollection('approvals', 'approval-workflow')} />}
        {active === 'reports' && <ReportsView state={state} queue={queue} notifications={notifications} onExport={exportBackup} onExportExcel={(collection) => exportCollection(collection, `laporan-${collection}`)} />}
        {active === 'account' && <AccountView settings={settings} setSettings={setSettings} roles={state.roles || []} onLogin={loginAccount} onLogout={logoutAccount} onForgotPassword={sendPasswordReset} onUpdatePassword={updateAccountPassword} />}
        {active === 'configuration' && <ConfigurationView settings={settings} setSettings={setSettings} queue={queue} syncStatus={syncStatus} onSave={saveConfiguration} onPush={processSyncQueue} onPull={pullSupabase} onReset={resetDemo} />}
      </main>
      {selectedAsset && <AssetDrawer asset={selectedAsset} documents={state.documents || []} workOrders={state.workOrders || []} contracts={state.contracts || []} invoices={state.invoices || []} onClose={() => setSelectedAsset(null)} onEdit={setAssetForm} onWO={(asset) => setWorkOrderForm({ asetId: asset.id, namaAset: asset.nama })} />}
      {assetForm && <AssetForm asset={assetForm} onClose={() => setAssetForm(null)} onSave={saveAsset} />}
      {workOrderForm && <WorkOrderForm form={workOrderForm} assets={state.assets || []} vendors={state.vendors || []} onClose={() => setWorkOrderForm(null)} onSave={saveWorkOrder} />}
      {contractForm && <ContractForm form={contractForm} assets={state.assets || []} state={state} onClose={() => setContractForm(null)} onSave={saveContract} />}
      {documentForm && <DocumentForm form={documentForm} assets={state.assets || []} onClose={() => setDocumentForm(null)} onSave={saveDocument} />}
      {quickAddForm && (quickAddForm.collection === 'invoices' ? <InvoiceForm values={quickAddForm.values} state={state} isEditing={quickAddForm.isEditing} onClose={() => setQuickAddForm(null)} onSave={(values) => saveQuickAdd('invoices', values)} /> : <QuickAddForm collection={quickAddForm.collection} values={quickAddForm.values} state={state} onClose={() => setQuickAddForm(null)} onSave={saveQuickAdd} />)}
    </div>
  );
}


function getQuickCollection(active) { return ({ operations: 'checklists', finance: 'invoices', inventory: 'inventory', vendors: 'vendors', tickets: 'tickets', approvals: 'approvals' })[active] || ''; }
const QUICK_MODULE_CONFIG = {
  contracts: { button: 'Tambah Kontrak', title: 'Tambah Kontrak', save: 'Simpan Kontrak' },
  checklists: { button: 'Tambah Checklist', title: 'Tambah Checklist Operasional', save: 'Simpan Checklist' },
  invoices: { button: 'Tambah Invoice', title: 'Tambah Invoice', save: 'Simpan Invoice' },
  inventory: { button: 'Tambah Sparepart', title: 'Tambah Sparepart / Inventory', save: 'Simpan Sparepart' },
  vendors: { button: 'Tambah Vendor', title: 'Tambah Vendor', save: 'Simpan Vendor' },
  tickets: { button: 'Tambah Ticket', title: 'Tambah Ticket', save: 'Simpan Ticket' },
  approvals: { button: 'Tambah Approval', title: 'Tambah Rule Approval', save: 'Simpan Approval' },
};
function quickModuleConfig(collection) { return QUICK_MODULE_CONFIG[collection] || { button: 'Tambah Data', title: 'Tambah Data', save: 'Simpan Data' }; }
function firstAsset(state = {}) { return (state.assets || [])[0] || {}; }
function firstContract(state = {}) { return (state.contracts || [])[0] || {}; }
function quickDefaults(collection, state = {}) { const asset = firstAsset(state); const contract = firstContract(state); const defaults = {
  contracts: contractDefaults(state),
  checklists: { id: nextChecklistId(state, 'Check-out'), tipe: 'Check-out', asetId: asset.id || '', aset: asset.nama || '', tanggal: today, waktu: new Date().toTimeString().slice(0, 5), pic: '', penerima: '', km: asset.km || '', bbm: '', kondisi: '', aksesoris: '', kerusakan: '', fotoBefore: '', fotoAfter: '', latitude: '', longitude: '', tandaTangan: 'Belum', status: 'Draft', catatan: '' },
  invoices: invoiceDefaults(state),
  inventory: { id: nextInventoryId(state, ''), nama: '', partNumber: '', merek: '', kategori: 'Kendaraan', spesifikasi: '', asetKompatibel: '', stok: 0, min: 0, stokMax: 0, satuan: 'Pcs', lokasi: '', rak: '', hargaBeli: 0, nilai: 0, supplierId: '', vendor: '', tanggalBeli: today, tanggalExpired: '', kondisi: 'Baru', status: 'Aktif', catatan: '' },
  vendors: { id: nextVendorId(state), nama: '', namaLegal: '', jenis: 'Bengkel Kendaraan', pic: '', kontak: '', email: '', alamat: '', kota: '', provinsi: '', npwp: '', bank: '', rekening: '', paymentTerm: '30 hari', layanan: '', wilayahLayanan: '', rating: 0, avgCost: 0, avgLeadTime: '', sla: '', transaksi: 0, kontrakMulai: '', kontrakSelesai: '', dokumen: '', status: 'Aktif', catatan: '' },
  tickets: { id: uid('TCK'), customer: '', asetId: asset.id || '', aset: asset.nama || '', kategori: 'Keluhan customer', prioritas: 'Sedang', sla: today, status: 'Open', pic: '', deskripsi: '' },
  approvals: { id: uid('APR'), proses: '', approval: '', limit: 0, status: 'Aktif', pending: '0' },
}; return defaults[collection] || { id: uid('DATA') }; }
function quickFields(collection, state = {}) { const assets = state.assets || []; const contracts = state.contracts || []; const vendors = state.vendors || []; const assetOptions = assets.map((asset) => asset.id); const contractOptions = contracts.map((contract) => contract.id); const vendorOptions = vendors.map((vendor) => vendor.id); return ({
  contracts: [ ['customer','Customer'], ['asetId','Aset', 'select', assetOptions], ['nilaiBulanan','Nilai Bulanan','number'], ['deposit','Deposit','number'], ['mulai','Mulai','date'], ['selesai','Selesai','date'], ['status','Status','select',['Aktif','Menunggu Dokumen','Selesai','Dibatalkan']], ['billingCycle','Billing Cycle','select',['Bulanan','Tahunan','Sekali Bayar']], ['dokumen','Nama Dokumen Kontrak'], ['renewal','Renewal / Catatan','textarea'] ],
  checklists: [ ['tipe','Tipe Checklist','select',['Check-out','Check-in','Inspection']], ['asetId','Aset','select',assetOptions], ['tanggal','Tanggal','date'], ['waktu','Waktu','time'], ['pic','PIC / Petugas'], ['penerima','Penerima / Penyerah'], ['km','KM / Odometer','number'], ['bbm','BBM / Battery Level'], ['kondisi','Kondisi Umum','textarea'], ['kerusakan','Kerusakan / Temuan','textarea'], ['aksesoris','Aksesoris & Dokumen','textarea'], ['fotoBefore','Foto Before / Nama File'], ['fotoAfter','Foto After / Nama File'], ['latitude','Latitude'], ['longitude','Longitude'], ['tandaTangan','Tanda Tangan','select',['Belum','Sudah','Menunggu']], ['status','Status','select',['Draft','Selesai','Menunggu']], ['catatan','Catatan Operasional','textarea'] ],
  invoices: [ ['contractId','Kontrak','select',contractOptions], ['customer','Customer'], ['periode','Periode'], ['amount','Amount','number'], ['paid','Paid','number'], ['dueDate','Due Date','date'], ['status','Status','select',['Unpaid','Partial','Paid']] ],
  inventory: [ ['nama','Nama Item'], ['partNumber','Part Number / SKU Supplier'], ['merek','Merek'], ['kategori','Kategori','select',categoryOptions], ['spesifikasi','Spesifikasi','textarea'], ['asetKompatibel','Aset / Kendaraan Kompatibel'], ['stok','Stok Saat Ini','number'], ['min','Minimum Stok','number'], ['stokMax','Maksimum Stok','number'], ['satuan','Satuan'], ['lokasi','Gudang / Lokasi'], ['rak','Rak / Bin'], ['hargaBeli','Harga Beli','number'], ['nilai','Nilai / Harga Keluar','number'], ['supplierId','Vendor Utama','select',['',...vendorOptions]], ['vendor','Nama Vendor Alternatif'], ['tanggalBeli','Tanggal Pembelian Terakhir','date'], ['tanggalExpired','Tanggal Kedaluwarsa','date'], ['kondisi','Kondisi','select',['Baru','Baik','Bekas Layak','Rusak']], ['status','Status','select',['Aktif','Minimum Stock','Habis','Nonaktif']], ['catatan','Catatan Sparepart','textarea'] ],
  vendors: [ ['nama','Nama Vendor / Bengkel'], ['namaLegal','Nama Legal Perusahaan'], ['jenis','Jenis Vendor','select',['Bengkel Kendaraan','Repair Kendaraan','Supplier Sparepart','Perangkat IT','Body Repair','Towing','Asuransi','Lainnya']], ['pic','PIC Vendor'], ['kontak','Nomor Telepon / WhatsApp','tel'], ['email','Email','email'], ['alamat','Alamat Lengkap','textarea'], ['kota','Kota'], ['provinsi','Provinsi'], ['npwp','NPWP'], ['bank','Bank'], ['rekening','Nomor Rekening'], ['paymentTerm','Termin Pembayaran'], ['layanan','Jenis Layanan','textarea'], ['wilayahLayanan','Wilayah Layanan'], ['rating','Rating','number'], ['avgCost','Biaya Rata-rata','number'], ['avgLeadTime','Lead Time Rata-rata'], ['sla','SLA / Waktu Respons'], ['transaksi','Jumlah Transaksi','number'], ['kontrakMulai','Kontrak Mulai','date'], ['kontrakSelesai','Kontrak Selesai','date'], ['dokumen','Dokumen / Perjanjian'], ['status','Status','select',['Aktif','Nonaktif','Evaluasi','Blacklist']], ['catatan','Catatan Evaluasi Vendor','textarea'] ],
  tickets: [ ['customer','Customer'], ['asetId','Aset','select',assetOptions], ['kategori','Kategori'], ['prioritas','Prioritas','select',priorityOptions], ['sla','SLA','date'], ['status','Status','select',['Open','In Progress','Closed']], ['pic','PIC'], ['deskripsi','Deskripsi','textarea'] ],
  approvals: [ ['proses','Proses'], ['approval','Approval Oleh'], ['limit','Limit Biaya','number'], ['status','Status','select',['Aktif','Nonaktif']], ['pending','Pending','number'] ],
})[collection] || []; }
function normalizeQuickRow(collection, values = {}, state = {}) { const asset = (state.assets || []).find((item) => item.id === values.asetId) || {}; const contract = (state.contracts || []).find((item) => item.id === values.contractId) || {}; const selectedVendor = (state.vendors || []).find((item) => item.id === values.supplierId) || {}; const map = {
  contracts: () => ({ ...values, id: values.id || uid('CTR'), aset: asset.nama || values.aset || '', nilaiBulanan: numeric(values.nilaiBulanan), deposit: numeric(values.deposit) }),
  checklists: () => ({ ...values, id: values.id || nextChecklistId(state, values.tipe), aset: asset.nama || values.aset || '', km: values.km || asset.km || '', status: values.status || 'Draft' }),
  invoices: () => normalizeInvoice({ ...values, id: values.id || uid('INV') }, state),
  inventory: () => ({ ...values, id: values.id || nextInventoryId(state, values.nama), stok: numeric(values.stok), min: numeric(values.min), stokMax: numeric(values.stokMax), hargaBeli: numeric(values.hargaBeli), nilai: numeric(values.nilai), vendor: selectedVendor.nama || values.vendor || '', status: values.status || 'Aktif' }),
  vendors: () => ({ ...values, id: values.id || nextVendorId(state), rating: Number(values.rating || 0), avgCost: numeric(values.avgCost), transaksi: numeric(values.transaksi) }),
  tickets: () => ({ ...values, id: values.id || uid('TCK'), aset: asset.nama || values.aset || '', status: values.status || 'Open' }),
  approvals: () => ({ ...values, id: values.id || uid('APR'), limit: numeric(values.limit), pending: String(values.pending ?? '0') }),
}; return map[collection] ? map[collection]() : { ...values, id: values.id || uid('DATA') }; }
function buildMetrics(state, notifications) { const assets = state.assets || []; const total = assets.length; const rented = assets.filter((a) => a.status === 'Disewakan').length; const maintenance = assets.filter((a) => ['Maintenance', 'Perbaikan'].includes(a.status)).length; const highRisk = assets.filter((a) => Number(a.risiko || 0) >= 60).length; const stnkDue = assets.filter((a) => a.kategori === 'Kendaraan' && stnkInfo(a).due).length; const utilization = total ? Math.round(assets.reduce((sum, a) => sum + Number(a.utilisasi || 0), 0) / total) : 0; const revenue = assets.reduce((sum, a) => sum + numeric(a.pendapatanBulanan), 0); const cost = assets.reduce((sum, a) => sum + numeric(a.biayaBulanan), 0); const unpaid = (state.invoices || []).filter((i) => i.status !== 'Paid').reduce((s, i) => s + Math.max(0, numeric(i.amount) - numeric(i.paid)), 0); return { total, rented, maintenance, highRisk, stnkDue, utilization, revenue, cost, profit: revenue - cost, notifications: notifications.filter((n) => n.level === 'red').length, ticketsOpen: (state.tickets || []).filter((t) => t.status !== 'Closed').length, unpaid }; }
function buildNotifications(state) { const list = []; (state.assets || []).forEach((asset) => { const stnk = stnkInfo(asset); if (stnk.due) list.push({ id: `NTF-STNK-${asset.id}`, type: 'STNK', level: 'red', title: `${asset.nama} perlu perpanjang STNK`, detail: `${asset.nomorPolisi || '-'} · ${stnk.message}`, due: asset.tanggalSTNK, action: 'Perpanjang STNK' }); ['tanggalKIR', 'tanggalPajak', 'tanggalAsuransi', 'tanggalServiceBerikutnya'].forEach((key) => { const label = { tanggalKIR: 'KIR', tanggalPajak: 'Pajak kendaraan', tanggalAsuransi: 'Asuransi', tanggalServiceBerikutnya: 'Service berkala' }[key]; const info = overdueLevel(asset[key], 30, 60); if (info.days !== null && info.days <= 30) list.push({ id: `NTF-${key}-${asset.id}`, type: label, level: info.level, title: `${label} ${asset.nama}`, detail: `${asset.nomorPolisi || asset.nomorRangka || '-'} · ${info.label}`, due: asset[key], action: `Tindak lanjut ${label}` }); }); }); (state.contracts || []).forEach((c) => { const info = overdueLevel(c.selesai, 30, 60); if (info.days !== null && info.days <= 30) list.push({ id: `NTF-CTR-${c.id}`, type: 'Kontrak', level: info.level, title: `Kontrak ${c.customer}`, detail: `${c.aset} · selesai ${info.label}`, due: c.selesai, action: 'Follow up renewal' }); }); (state.invoices || []).filter((i) => i.status !== 'Paid').forEach((i) => { const info = overdueLevel(i.dueDate, 7, 14); if (info.days !== null && info.days <= 14) list.push({ id: `NTF-INV-${i.id}`, type: 'Invoice', level: info.level, title: `Tagihan ${i.customer}`, detail: `${i.id} · sisa ${currency.format(Math.max(0, numeric(i.amount) - numeric(i.paid)))} · ${info.label}`, due: i.dueDate, action: 'Follow up pembayaran' }); }); (state.tickets || []).filter((t) => t.status !== 'Closed').forEach((t) => { const info = overdueLevel(t.sla, 1, 3); if (info.days !== null && info.days <= 3) list.push({ id: `NTF-TCK-${t.id}`, type: 'Ticket SLA', level: t.prioritas === 'Tinggi' || info.level === 'red' ? 'red' : 'amber', title: `${t.id} SLA ${t.customer}`, detail: `${t.aset} · ${t.status} · ${info.label}`, due: t.sla, action: 'Selesaikan ticket' }); }); (state.approvals || []).filter((a) => numeric(a.pending) > 0).forEach((a) => list.push({ id: `NTF-APR-${a.id}`, type: 'Approval', level: 'amber', title: `${a.proses} menunggu approval`, detail: `${a.pending} item pending · ${a.approval}`, due: today, action: 'Review approval' })); return list.sort((a, b) => (a.level === 'red' ? -1 : 1) - (b.level === 'red' ? -1 : 1)); }

function notificationTarget(type) {
  if (['STNK', 'KIR', 'Pajak kendaraan', 'Asuransi', 'Service berkala'].includes(type)) return 'assets';
  if (type === 'Invoice') return 'finance';
  if (type === 'Kontrak') return 'contracts';
  if (type === 'Ticket SLA') return 'tickets';
  if (type === 'Approval') return 'approvals';
  return 'dashboard';
}

function NotificationBell({ notifications, open, setOpen, setActive }) {
  const redCount = notifications.filter((n) => n.level === 'red').length;
  const topItems = notifications.slice(0, 10);
  const goTo = (type) => { setActive(notificationTarget(type)); setOpen(false); };
  return <div className="notification-bell-wrap"><button className="notification-bell" title="Notifikasi" onClick={() => setOpen(!open)}><span aria-hidden="true">🔔</span>{notifications.length > 0 && <b className={redCount ? 'danger' : ''}>{redCount || notifications.length}</b>}</button>{open && <div className="notification-popover"><div className="notification-popover-head"><div><p className="eyebrow">Notification Center</p><h3>Lonceng Notifikasi</h3><span>{notifications.length} reminder aktif · {redCount} kritis</span></div><button onClick={() => setOpen(false)}>×</button></div><div className="notification-popover-list">{topItems.length ? topItems.map((item) => <button className={`notification-item ${item.level === 'red' ? 'critical' : ''}`} key={item.id} onClick={() => goTo(item.type)}><div><b>{item.title}</b><span>{item.detail} · due: {item.due || '-'}</span></div><i className={`pill ${item.level}`}>{item.type}</i></button>) : <div className="notification-empty">Tidak ada notifikasi aktif.</div>}</div><div className="notification-popover-foot"><span>Notifikasi tidak lagi berada di menu utama.</span><button className="btn ghost" onClick={() => { setActive('reports'); setOpen(false); }}>Audit & Laporan</button></div></div>}</div>;
}


function Dashboard({ metrics, state, notifications, onOpenNotifications }) { const topAlerts = notifications.slice(0, 7); return <section className="stack"><div className="metrics enterprise-metrics"><Metric label="Total Aset" value={metrics.total} note="Rental + internal" /><Metric label="STNK Perlu Perpanjang" value={metrics.stnkDue} note="Notif merah ≤ 30 hari / expired" tone="danger" /><Metric label="Profit Bulanan" value={currency.format(metrics.profit)} note={`${currency.format(metrics.revenue)} revenue`} tone="success" /><Metric label="Tagihan Belum Lunas" value={currency.format(metrics.unpaid)} note="Invoice unpaid/partial" tone="danger" /><Metric label="Utilisasi" value={`${metrics.utilization}%`} note="Rata-rata seluruh aset" /><Metric label="Ticket Open" value={metrics.ticketsOpen} note="SLA berjalan" /><Metric label="Aset Maintenance" value={metrics.maintenance} note="Unit sedang diperbaiki" /><Metric label="Notif Kritis" value={metrics.notifications} note="Butuh tindakan cepat" tone="danger" /></div><div className="dashboard-grid two"><div className="panel"><div className="panel-head compact"><div><h2>Command Center</h2><p>Dashboard hanya menampilkan tindak lanjut utama. Detail lengkap tetap ada di halaman masing-masing supaya tidak double.</p></div></div><div className="quick-summary"><div><b>{(state.contracts || []).filter((c) => c.status === 'Aktif').length}</b><span>Kontrak aktif</span></div><div><b>{(state.inventory || []).filter((i) => numeric(i.stok) <= numeric(i.min)).length}</b><span>Sparepart minimum</span></div><div><b>{(state.approvals || []).reduce((s, a) => s + numeric(a.pending), 0)}</b><span>Approval pending</span></div></div></div><div className="panel"><div className="panel-head compact"><div><h2>Alert Utama</h2><p>Ringkasan alert tetap di Dashboard. Daftar lengkapnya sekarang lewat ikon lonceng kecil di pojok atas, bukan menu tab.</p></div><button className="btn ghost" onClick={onOpenNotifications}>Buka Lonceng</button></div><div className="alert-list">{topAlerts.map((item) => <div className={`alert-row ${item.level === 'red' ? 'alert-critical' : ''}`} key={item.id}><div><b>{item.title}</b><span>{item.detail}</span></div><span className={`pill ${item.level}`}>{item.type}</span></div>)}</div></div></div></section>; }
function Metric({ label, value, note, tone = '' }) { return <div className={`metric ${tone}`}><span>{label}</span><b>{value}</b><p>{note}</p></div>; }
function AssetsView({ assets, contracts = [], invoices = [], documents = [], query, setQuery, filter, setFilter, filterOptions, importExcel, onDetail, onEdit, onWO, onDelete }) { return <section className="stack"><ImportPanel onImport={importExcel} /><div className="panel"><div className="panel-head"><div><h2>Daftar Aset - Tampilan Kartu BPKB</h2><p>Aset sekarang tampil seperti kartu identitas BPKB/kontrak: foto unit, customer, STNK, masa berlaku, payment status, custody, timeline, dan attachment tetap terhubung ke Aset 360.</p></div><div className="filters"><input placeholder="Cari aset, QR, customer, nopol, rangka, wilayah..." value={query} onChange={(e) => setQuery(e.target.value)} /><select value={filter} onChange={(e) => setFilter(e.target.value)}>{filterOptions.map((item) => <option key={item}>{item}</option>)}</select></div></div><div className="asset-grid asset-card-grid">{assets.map((asset) => { const contract = getContractForAsset(asset, contracts); const relatedInvoices = invoices.filter((invoice) => invoice.contractId === contract?.id); return <AssetCard key={asset.id} asset={asset} contract={contract} invoices={relatedInvoices} documents={documents} onDetail={onDetail} onEdit={onEdit} onWO={onWO} onDelete={onDelete} />; })}</div></div></section>; }
function ImportPanel({ onImport }) { return <div className="panel import-panel"><div><p className="eyebrow">Import Excel Kendaraan Sewa</p><h2>Upload data kendaraan sewa dari Excel/CSV</h2><p>Kolom didukung: no, nama kendaraan, warna, tahun kendaraan, nomor rangka, nomor mesin, nomor polisi/nopol, perusahaan user, alamat user, wilayah, tanggal STNK, BPKB, KIR, pajak, asuransi, GPS, ID GPS, dan KM/odometer.</p></div><div className="import-actions"><a className="btn ghost" href="/template_import_kendaraan_sewa.xlsx" download>Template XLSX</a><a className="btn ghost" href="/template_import_kendaraan_sewa.csv" download>Template CSV</a><label className="btn primary">Import Excel<input hidden type="file" accept=".xlsx,.xls,.csv" onChange={(e) => onImport(e.target.files?.[0])} /></label></div></div>; }
function getAssetBpkbNumber(asset) {
  const raw = String(asset.nomorBPKB || asset.bpkbNumber || asset.id || '').replace(/\D/g, '').slice(-5);
  return `BPKB - No. ${raw ? raw.padStart(5, '0') : '00123'}`;
}
function getContractForAsset(asset, contracts = []) {
  return contracts.find((contract) => contract.asetId === asset.id || contract.aset === asset.nama) || null;
}
function getPaymentSummary(contract, invoices = []) {
  const related = invoices.filter((invoice) => invoice.contractId === contract?.id);
  if (!related.length) return { text: 'Belum ada invoice', status: 'Menunggu', tone: 'amber' };
  const paid = related.filter((invoice) => invoice.status === 'Paid').length;
  const text = `${paid}/${related.length} invoice lunas`;
  const allPaid = paid === related.length;
  return { text, status: allPaid ? 'On-Time' : 'Follow Up', tone: allPaid ? 'green' : 'red' };
}
function formatPeriod(contract) {
  if (!contract) return '-';
  return `${contract.mulai || '-'} s/d ${contract.selesai || '-'}`;
}
function formatCardDate(value) {
  if (!value) return '-';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return new Intl.DateTimeFormat('en-US', { month: 'long', year: 'numeric' }).format(parsed);
}
function formatCardMonth(value) {
  if (!value) return '-';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return new Intl.DateTimeFormat('en-US', { month: 'short', year: 'numeric' }).format(parsed);
}
function formatLoanDuration(contract) {
  if (!contract?.mulai || !contract?.selesai) return '-';
  const start = new Date(contract.mulai);
  const end = new Date(contract.selesai);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return formatPeriod(contract);
  const months = Math.max(1, ((end.getFullYear() - start.getFullYear()) * 12) + (end.getMonth() - start.getMonth()) + 1);
  return `${months} months`;
}
function formatLoanPeriod(contract) {
  if (!contract) return '';
  return `(${formatCardMonth(contract.mulai)} - ${formatCardMonth(contract.selesai)})`;
}
function getAssetDocForCard(asset, documents = []) {
  const docs = (documents || []).filter((doc) => documentMatchesAsset(doc, asset));
  return docs.find((doc) => /bpkb/i.test(`${doc.tipe} ${doc.nama}`)) || docs.find((doc) => /stnk/i.test(`${doc.tipe} ${doc.nama}`)) || docs[0] || null;
}
function buildAssetCardTimeline(asset, contract) {
  return [
    { label: 'Received', value: formatCardMonth(contract?.mulai || asset.updatedAt || today), tone: 'green' },
    { label: 'Stored', value: asset.cabang || asset.lokasi || '-', tone: 'green' },
    { label: asset.status === 'Disewakan' ? 'Released' : 'Updated', value: formatCardMonth(contract?.selesai || asset.updatedAt || today), tone: 'red' },
  ];
}
function AssetPhotoCard({ asset, contract, invoices = [], documents = [], compact = false }) {
  const payment = getPaymentSummary(contract, invoices);
  const stnk = stnkInfo(asset);
  const attachment = getAssetDocForCard(asset, documents);
  const timeline = buildAssetCardTimeline(asset, contract);
  const title = asset.kategori === 'Kendaraan' ? getAssetBpkbNumber(asset) : `ASSET CARD - No. ${String(asset.id || '').split('-').pop() || '00001'}`;
  const identityLine = [asset.nomorPolisi || asset.nomorRangka || '-', asset.km ? `${asset.km} KM` : '', asset.warna || ''].filter(Boolean).join(' · ');
  return <div className={`asset-id-card ${compact ? 'compact' : ''}`}><div className="asset-id-head"><b>{title}</b><span>Contract #{contract?.id || asset.id}</span></div><div className="asset-id-photo">{asset.gambar ? <img alt={asset.nama} src={asset.gambar} /> : <div className="asset-photo-placeholder"><span>{asset.kategori === 'Kendaraan' ? '🚗' : '💻'}</span><small>{asset.kategori === 'Kendaraan' ? 'Foto Kendaraan' : 'Foto Perangkat'}</small></div>}</div><div className="asset-id-name"><b>{asset.nama}</b><span>{identityLine}</span></div>{!compact && <><div className="asset-id-info"><div><span>Customer</span><b>{contract?.customer || asset.pemegang || '-'}</b></div><div><span>STNK</span><b>{asset.nomorSTNK || asset.tanggalSTNK || '-'}</b></div><div><span>Expiration Date</span><b>{formatCardDate(asset.tanggalSTNK || asset.tanggalAsuransi)}</b></div><div><span>Loan Duration</span><b>{formatLoanDuration(contract)}<small>{formatLoanPeriod(contract)}</small></b></div><div><span>Payments</span><b>{payment.text}</b></div><div><span>Payment Status</span><b className={`status-${payment.tone}`}>{payment.status}</b></div><div><span>Custody Info</span><b>{asset.cabang || asset.lokasi || '-'}</b></div></div><div className="asset-id-timeline"><span>Timeline</span><div>{timeline.map((item) => <p key={`${item.label}-${item.value}`}><i className={item.tone === 'red' ? 'red' : 'green'} /><b>{item.label}</b> - {item.value}</p>)}</div></div><div className="asset-id-attach"><span>Attachments</span><div className="asset-id-attachment-box"><div className="doc-thumb"><i /><i /><i /></div>{attachment?.fileData ? <a className="scan-button" href={attachment.fileData} target="_blank" rel="noreferrer">View Scanned BPKB</a> : <button type="button">{attachment ? 'View Scanned BPKB' : 'No scanned file'}</button>}</div></div></>}{stnk.due && <div className="asset-id-warning">STNK merah: {stnk.message}</div>}</div>;
}

function AssetCard({ asset, contract, invoices = [], documents = [], onDetail, onEdit, onWO, onDelete }) { const profit = numeric(asset.pendapatanBulanan) - numeric(asset.biayaBulanan); const stnk = stnkInfo(asset); const docCount = (documents || []).filter((doc) => documentMatchesAsset(doc, asset)).length; return <article className={`asset-card asset-card-pass ${stnk.due ? 'asset-card-alert' : ''}`}><div className="asset-image asset-pass-stage"><AssetPhotoCard asset={asset} contract={contract} invoices={invoices} documents={documents} /></div><div className="asset-body asset-pass-footer"><div className="asset-title"><div><b>{asset.id}</b><span>{asset.kategori} · {asset.status}</span></div><div className="title-pills"><span className={`pill ${asset.gpsAktif ? 'green' : 'slate'}`}>{asset.gpsAktif ? 'GPS aktif' : 'GPS off'}</span><span className={`pill ${riskClass(asset.risiko)}`}>Risiko {asset.risiko}%</span></div></div>{stnk.due && <div className="stnk-alert"><div><b>Notifikasi merah STNK</b><span>{stnk.message}</span></div></div>}<div className="asset-kpis"><span><b>{asset.km ? `${asset.km} km` : '-'}</b>KM</span><span><b>{docCount}</b>Dokumen</span><span><b>{currency.format(profit)}</b>Profit/bln</span></div><p>{asset.aksiBerikutnya}</p></div><div className="card-actions"><button onClick={() => onDetail(asset)}>Aset 360</button><button onClick={() => onEdit(asset)}>Edit</button><button onClick={() => onWO(asset)}>Buat WO</button><button className="danger" onClick={() => onDelete(asset.id)}>Hapus</button></div></article>; }

// v3.8: restored enterprise page views that were accidentally removed in v3.7 asset card update.
function ContractsView({ contracts, assets, onAdd, onEdit, exportExcel }) {
  const active = contracts.filter((c) => c.status === 'Aktif').length;
  const revenue = contracts.reduce((sum, contract) => sum + numeric(contract.nilaiBulanan), 0);
  const contractValue = contracts.reduce((sum, contract) => sum + numeric(contract.totalKontrak || contract.nilaiBulanan), 0);
  const rows = contracts.map((contract) => ({ ...contract, nopol: contract.nopol || assets.find((asset) => asset.id === contract.asetId)?.nomorPolisi || '-', durasi: contract.durasiBulan ? `${contract.durasiBulan} bulan` : `${monthsBetween(contract.mulai, contract.selesai)} bulan` }));
  return <section className="stack"><div className="metrics"><Metric label="Kontrak Aktif" value={active} note="Rental berjalan" /><Metric label="Revenue Kontrak" value={currency.format(revenue)} note="Bulanan" tone="success" /><Metric label="Nilai Kontrak" value={currency.format(contractValue)} note="Termasuk biaya & pajak" /><Metric label="Aset Disewakan" value={assets.filter((a) => a.status === 'Disewakan').length} note="Unit aktif" /></div><div className="panel"><div className="panel-head"><div><h2>Modul Kontrak Sewa</h2><p>Customer, PIC, unit/nopol, nilai, deposit, PO, periode, SLA, dokumen, approval, renewal, dan status kontrak.</p></div><div className="module-actions"><button className="btn primary" onClick={onAdd}>Tambah Kontrak</button><button className="btn ghost" onClick={exportExcel}>Export Excel</button></div></div><DataTable rows={rows} columns={[[ 'id', 'No Kontrak' ], [ 'customer', 'Customer' ], [ 'aset', 'Aset' ], [ 'nopol', 'Nopol / Serial' ], [ 'nomorPO', 'No. PO' ], [ 'nilaiBulanan', 'Nilai/Bulan', (v) => currency.format(numeric(v)) ], [ 'deposit', 'Deposit', (v) => currency.format(numeric(v)) ], [ 'durasi', 'Durasi' ], [ 'mulai', 'Mulai' ], [ 'selesai', 'Selesai' ], [ 'approvalStatus', 'Approval' ], [ 'status', 'Status' ], [ 'dokumen', 'Dokumen', (v, row) => row.dokumenFileData ? <a className="inline-link" href={row.dokumenFileData} target="_blank" rel="noreferrer">Review kontrak</a> : (v || '-') ], [ 'aksi', 'Aksi', (_v, row) => <button className="btn small" onClick={() => onEdit(row)}>Edit</button> ]]} /></div></section>;
}
function OperationsView({ checklists, assets, onAdd, onExport }) { return <section className="stack"><div className="metrics"><Metric label="Checklist" value={checklists.length} note="Check-in / check-out" /><Metric label="Tanda Tangan" value={checklists.filter((c) => c.tandaTangan === 'Sudah').length} note="Digital sign" tone="success" /><Metric label="Draft" value={checklists.filter((c) => c.status === 'Draft').length} note="Perlu dilengkapi" tone="danger" /><Metric label="Unit Siap" value={assets.filter((a) => a.status === 'Tersedia').length} note="Ready operation" /></div><div className="panel"><div className="panel-head"><div><h2>Check-in / Check-out & Digital Inspection Checklist</h2><p>Serah-terima aset, KM, BBM, kondisi, aksesoris, foto before/after, penerima, dan tanda tangan digital.</p></div><div className="module-actions"><button className="btn primary" onClick={onAdd}>Tambah Checklist</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><DataTable rows={checklists} columns={[[ 'id', 'ID' ], [ 'tipe', 'Tipe' ], [ 'aset', 'Aset' ], [ 'tanggal', 'Tanggal' ], [ 'pic', 'PIC' ], [ 'penerima', 'Penerima' ], [ 'km', 'KM' ], [ 'bbm', 'BBM' ], [ 'kondisi', 'Kondisi' ], [ 'aksesoris', 'Aksesoris' ], [ 'tandaTangan', 'Tanda Tangan' ], [ 'status', 'Status' ]]} /></div><div className="panel"><h2>Checklist Kendaraan & Perangkat IT</h2><div className="check-grid"><ChecklistCard title="Kendaraan" items={['Body', 'Ban', 'Mesin', 'Interior', 'AC', 'Lampu', 'Rem', 'Oli', 'STNK', 'Toolkit', 'Dongkrak', 'Segitiga pengaman']} /><ChecklistCard title="Perangkat IT" items={['Layar', 'Keyboard', 'Touchpad', 'Charger', 'Battery health', 'Serial number', 'OS', 'Office/license', 'Tas/sleeve', 'Stylus', 'Kondisi fisik']} /></div></div></section>; }
function ChecklistCard({ title, items }) { return <div className="check-card"><h3>{title}</h3>{items.map((item) => <label key={item}><input type="checkbox" readOnly checked /> {item}</label>)}</div>; }
function WorkOrderView({ workOrders, assets, vendors, onCreate, onEdit, onComplete }) {
  const [nopolSearch, setNopolSearch] = useState('');
  const assetMap = useMemo(() => Object.fromEntries((assets || []).map((asset) => [asset.id, asset])), [assets]);
  const historyRows = useMemo(() => (workOrders || []).map((wo) => {
    const asset = assetMap[wo.asetId] || {};
    return {
      id: wo.id,
      asetId: wo.asetId,
      namaAset: wo.namaAset || asset.nama || '-',
      nopol: wo.nopol || asset.nomorPolisi || '-',
      jenis: wo.jenis,
      prioritas: wo.prioritas,
      status: wo.status,
      vendor: vendors.find((v) => v.id === wo.vendorId)?.nama || wo.pic,
      pic: wo.pic,
      tanggalMulai: wo.tanggalMulai,
      jatuhTempo: wo.jatuhTempo,
      tanggalSelesai: wo.tanggalSelesai || '',
      biaya: numeric(wo.biaya ?? wo.estimasiBiaya),
      sparepart: wo.sparepart || '-',
      fotoBefore: wo.fotoBefore || '-',
      fotoAfter: wo.fotoAfter || '-',
      keluhan: wo.keluhan,
      checklist: wo.checklist,
    };
  }), [workOrders, assetMap, vendors]);

  const nopolOptions = useMemo(
    () => Array.from(new Set(historyRows.map((row) => row.nopol).filter((nopol) => nopol && nopol !== '-'))).sort((a, b) => a.localeCompare(b, 'id')),
    [historyRows],
  );
  const normalizeNopol = (value) => String(value || '').replace(/\s+/g, '').toLowerCase();
  const normalizedSearch = normalizeNopol(nopolSearch);
  const filteredRows = useMemo(
    () => normalizedSearch
      ? historyRows.filter((row) => normalizeNopol(row.nopol).includes(normalizedSearch))
      : historyRows,
    [historyRows, normalizedSearch],
  );
  const totalBiaya = filteredRows.reduce((sum, row) => sum + numeric(row.biaya), 0);
  const activeFilterLabel = nopolSearch.trim() || 'Semua nopol';

  function exportMaintenanceExcel() {
    const rows = filteredRows.map((row) => ({
      'Nomor Polisi': row.nopol,
      'Nama Kendaraan': row.namaAset,
      'ID Work Order': row.id,
      'Jenis WO': row.jenis,
      Prioritas: row.prioritas,
      Status: row.status,
      PIC: row.pic,
      Vendor: row.vendor,
      'Tanggal Mulai': row.tanggalMulai,
      'Jatuh Tempo': row.jatuhTempo,
      'Tanggal Selesai': row.tanggalSelesai || '-',
      'Nominal Biaya Aktual': row.biaya,
      Sparepart: row.sparepart,
      'Foto Before': row.fotoBefore,
      'Foto After': row.fotoAfter,
      Keluhan: row.keluhan,
      Checklist: row.checklist,
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows.length ? rows : [{
      'Nomor Polisi / Pencarian': activeFilterLabel,
      Keterangan: 'Belum ada histori WO yang sesuai filter',
    }]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'History WO');
    const safeNopol = nopolSearch.trim()
      ? nopolSearch.replace(/[^a-z0-9]+/gi, '-').replace(/^-+|-+$/g, '').toLowerCase()
      : 'semua-nopol';
    XLSX.writeFile(workbook, `history-wo-${safeNopol || 'hasil-pencarian'}-${today}.xlsx`);
  }

  return <section className="stack">
    <div className="metrics">
      <Metric label="Total Work Order" value={filteredRows.length} note={activeFilterLabel} />
      <Metric label="Total Biaya" value={currency.format(totalBiaya)} note="Nominal Biaya Aktual sesuai filter" tone="danger" />
      <Metric label="WO Selesai" value={filteredRows.filter((row) => ['Selesai', 'Ditutup'].includes(row.status)).length} note="Histori tertutup sesuai filter" tone="success" />
      <Metric label="Nopol Tercatat" value={nopolOptions.length} note="Kendaraan punya histori" />
    </div>
    <div className="panel">
      <div className="panel-head">
        <div>
          <h2>Work Order (WO) - History per Nopol</h2>
          <p>Pilih atau cari nopol untuk menampilkan hanya histori kendaraan tersebut. Export Excel mengikuti nopol atau hasil pencarian yang aktif.</p>
        </div>
        <div className="maintenance-actions">
          <label className="nopol-search-control">
            <span>Cari / pilih nopol</span>
            <input
              type="search"
              list="wo-nopol-options"
              value={nopolSearch}
              onChange={(event) => setNopolSearch(event.target.value)}
              placeholder="Contoh: B 184 AZN"
              aria-label="Cari atau pilih nomor polisi"
            />
            <datalist id="wo-nopol-options">
              {nopolOptions.map((nopol) => <option key={nopol} value={nopol} />)}
            </datalist>
          </label>
          {nopolSearch && <button className="btn ghost" onClick={() => setNopolSearch('')}>Semua Nopol</button>}
          <button className="btn ghost" onClick={exportMaintenanceExcel}>Export Excel Nopol</button>
          <button className="btn primary" onClick={() => onCreate(assets[0] || {})}>Buat WO</button>
        </div>
      </div>
      <div className="filter-summary">
        <b>{filteredRows.length} WO ditampilkan</b>
        <span>Filter aktif: {activeFilterLabel}. File export hanya memuat data yang sedang tampil.</span>
      </div>
      <div className="table-wrap"><table><thead><tr><th>Nopol</th><th>Kendaraan</th><th>ID WO</th><th>Jenis</th><th>Prioritas</th><th>Status</th><th>Mulai</th><th>Jatuh Tempo</th><th>Tanggal Selesai</th><th>Vendor</th><th>Sparepart</th><th>Nominal Biaya Aktual</th><th>Aksi Cepat</th></tr></thead><tbody>{filteredRows.length ? filteredRows.map((wo) => {
        const done = ['Selesai', 'Ditutup'].includes(wo.status);
        return <tr key={wo.id}><td><b>{wo.nopol}</b></td><td>{wo.namaAset}</td><td>{wo.id}</td><td>{wo.jenis}</td><td><span className={`pill ${wo.prioritas === 'Kritis' ? 'red' : wo.prioritas === 'Tinggi' ? 'amber' : 'blue'}`}>{wo.prioritas}</span></td><td><span className={`pill ${done ? 'green' : 'amber'}`}>{wo.status}</span></td><td>{wo.tanggalMulai}</td><td>{wo.jatuhTempo}</td><td>{wo.tanggalSelesai || '-'}</td><td>{wo.vendor}</td><td>{wo.sparepart}</td><td><b>{currency.format(numeric(wo.biaya))}</b></td><td><div className="row-actions"><button className="btn small" onClick={() => onEdit(wo)}>Edit</button>{!done && <button className="btn small success" onClick={() => onComplete(wo.id)}>Selesai</button>}</div></td></tr>;
      }) : <tr><td colSpan="13"><div className="empty-filter-state"><b>Nopol tidak ditemukan</b><span>Coba kata kunci lain atau klik “Semua Nopol”.</span></div></td></tr>}</tbody></table></div>
    </div>
  </section>;
}
function FinanceView({ invoices, assets, contracts, onAdd, onEdit, onPaid, onExport }) { const normalizedInvoices = invoices.map((invoice) => normalizeInvoice(invoice, { contracts, assets })); const revenue = normalizedInvoices.reduce((s, i) => s + numeric(i.amount), 0); const paid = normalizedInvoices.reduce((s, i) => s + numeric(i.paid) + numeric(i.withholdingAmount), 0); const unpaid = normalizedInvoices.reduce((s, i) => s + numeric(i.remaining), 0); const profitRows = assets.map((a) => ({ id: a.id, aset: a.nama, nopol: a.nomorPolisi, pendapatan: numeric(a.pendapatanBulanan), biaya: numeric(a.biayaBulanan), profit: numeric(a.pendapatanBulanan) - numeric(a.biayaBulanan), utilisasi: `${a.utilisasi}%`, status: a.status })); return <section className="stack"><div className="metrics"><Metric label="Invoice Terbit" value={currency.format(revenue)} note="Total billing" /><Metric label="Terbayar" value={currency.format(paid)} note="Pembayaran + potongan pajak" tone="success" /><Metric label="Belum Lunas" value={currency.format(unpaid)} note="Sisa piutang aktif" tone="danger" /><Metric label="Kontrak" value={contracts.length} note="Sumber tagihan" /></div><div className="panel"><div className="panel-head"><div><h2>Billing & Invoice</h2><p>Invoice terhubung ke kontrak dan aset, rincian biaya, pajak, pembayaran, lampiran, serta status piutang.</p></div><div className="module-actions"><button className="btn primary" onClick={onAdd}>Tambah Invoice</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><div className="table-wrap"><table><thead><tr><th>Invoice</th><th>Tanggal</th><th>Kontrak</th><th>Customer</th><th>Aset / Nopol</th><th>Periode</th><th>Total</th><th>Paid</th><th>Sisa</th><th>Due Date</th><th>Status</th><th>Aksi</th></tr></thead><tbody>{normalizedInvoices.map((i) => <tr key={i.id}><td><b>{i.id}</b></td><td>{i.invoiceDate}</td><td>{i.contractId || '-'}</td><td>{i.customer}</td><td>{i.asset || '-'}<br/><small>{i.nopol || '-'}</small></td><td>{i.periode}</td><td>{currency.format(numeric(i.amount))}</td><td>{currency.format(numeric(i.paid))}</td><td>{currency.format(numeric(i.remaining))}</td><td>{i.dueDate}</td><td><span className={`pill ${i.status === 'Paid' ? 'green' : i.status === 'Partial' ? 'amber' : 'red'}`}>{i.status}</span></td><td><div className="row-actions"><button className="btn small" onClick={() => onEdit(i)}>Edit</button>{i.status !== 'Paid' && <button className="btn small success" onClick={() => onPaid(i.id)}>Tandai Lunas</button>}{i.invoiceFileData && <a className="btn small ghost" href={i.invoiceFileData} target="_blank" rel="noreferrer">File</a>}</div></td></tr>)}</tbody></table></div></div><div className="panel"><div className="panel-head compact"><h2>Profitability per Aset</h2></div><DataTable rows={profitRows} columns={[[ 'aset', 'Aset' ], [ 'nopol', 'Nopol/Serial' ], [ 'pendapatan', 'Pendapatan', (v) => currency.format(v) ], [ 'biaya', 'Biaya', (v) => currency.format(v) ], [ 'profit', 'Profit', (v) => currency.format(v) ], [ 'utilisasi', 'Utilisasi' ], [ 'status', 'Status' ]]} /></div></section>; }
function InventoryView({ inventory, onAdd, onExport }) { const low = inventory.filter((i) => numeric(i.stok) <= numeric(i.min)); return <section className="stack"><div className="metrics"><Metric label="Item Sparepart" value={inventory.length} note="SKU aktif" /><Metric label="Minimum Stock" value={low.length} note="Perlu restock" tone="danger" /><Metric label="Nilai Stok" value={currency.format(inventory.reduce((s, i) => s + numeric(i.stok) * numeric(i.nilai), 0))} note="Estimasi persediaan" /><Metric label="Lokasi" value={new Set(inventory.map((i) => i.lokasi)).size} note="Gudang/cabang" /></div><div className="panel"><div className="panel-head"><div><h2>Sparepart & Inventory</h2><p>Stok oli, ban, aki, filter, charger, adapter, unit pengganti, minimum stock, dan vendor.</p></div><div className="module-actions"><button className="btn primary" onClick={onAdd}>Tambah Sparepart</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><DataTable rows={inventory} columns={[[ 'id', 'SKU' ], [ 'nama', 'Item' ], [ 'partNumber', 'Part Number' ], [ 'merek', 'Merek' ], [ 'kategori', 'Kategori' ], [ 'stok', 'Stok' ], [ 'min', 'Min' ], [ 'stokMax', 'Max' ], [ 'satuan', 'Satuan' ], [ 'lokasi', 'Lokasi' ], [ 'rak', 'Rak' ], [ 'hargaBeli', 'Harga Beli', (v) => currency.format(numeric(v)) ], [ 'nilai', 'Nilai/Unit', (v) => currency.format(numeric(v)) ], [ 'vendor', 'Vendor' ], [ 'status', 'Status' ]]} /></div></section>; }
function VendorView({ vendors, workOrders, onAdd, onExport }) { return <section className="stack"><div className="metrics"><Metric label="Vendor Aktif" value={vendors.filter((v) => v.status === 'Aktif').length} note="Bengkel + IT" /><Metric label="Transaksi" value={vendors.reduce((s, v) => s + numeric(v.transaksi), 0)} note="Total histori" /><Metric label="WO Vendor" value={workOrders.filter((wo) => wo.vendorId).length} note="Terhubung WO" /><Metric label="Rating Avg" value={(vendors.reduce((s, v) => s + Number(v.rating), 0) / Math.max(1, vendors.length)).toFixed(1)} note="Performa" tone="success" /></div><div className="panel"><div className="panel-head"><div><h2>Vendor / Bengkel / Supplier Management</h2><p>Rating vendor, biaya rata-rata, lead time, transaksi, jenis layanan, dan performa vendor.</p></div><div className="module-actions"><button className="btn primary" onClick={onAdd}>Tambah Vendor</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><DataTable rows={vendors} columns={[[ 'id', 'ID' ], [ 'nama', 'Vendor' ], [ 'namaLegal', 'Nama Legal' ], [ 'jenis', 'Jenis' ], [ 'pic', 'PIC' ], [ 'kontak', 'Kontak' ], [ 'email', 'Email' ], [ 'kota', 'Kota' ], [ 'wilayahLayanan', 'Wilayah' ], [ 'rating', 'Rating' ], [ 'avgCost', 'Avg Cost', (v) => currency.format(numeric(v)) ], [ 'avgLeadTime', 'Lead Time' ], [ 'sla', 'SLA' ], [ 'transaksi', 'Transaksi' ], [ 'status', 'Status' ]]} /></div></section>; }
function TicketView({ tickets, onAdd, onCloseTicket, onExport }) { return <section className="stack"><div className="metrics"><Metric label="Ticket Open" value={tickets.filter((t) => t.status !== 'Closed').length} note="Aktif" tone="danger" /><Metric label="High Priority" value={tickets.filter((t) => t.prioritas === 'Tinggi').length} note="Butuh cepat" tone="danger" /><Metric label="Closed" value={tickets.filter((t) => t.status === 'Closed').length} note="Selesai" tone="success" /><Metric label="Total Ticket" value={tickets.length} note="Customer + internal" /></div><div className="panel"><div className="panel-head"><div><h2>SLA & Ticketing</h2><p>Keluhan customer, unit mogok, laptop error, delivery/pickup, status, PIC, deadline SLA, dan closure.</p></div><div className="module-actions"><button className="btn primary" onClick={onAdd}>Tambah Ticket</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><div className="table-wrap"><table><thead><tr><th>Ticket</th><th>Customer</th><th>Aset</th><th>Kategori</th><th>Prioritas</th><th>SLA</th><th>Status</th><th>PIC</th><th>Deskripsi</th><th>Aksi</th></tr></thead><tbody>{tickets.map((t) => <tr key={t.id}><td><b>{t.id}</b></td><td>{t.customer}</td><td>{t.aset}</td><td>{t.kategori}</td><td><span className={`pill ${t.prioritas === 'Tinggi' ? 'red' : 'amber'}`}>{t.prioritas}</span></td><td>{t.sla}</td><td>{t.status}</td><td>{t.pic}</td><td>{t.deskripsi}</td><td>{t.status !== 'Closed' && <button className="btn small success" onClick={() => onCloseTicket(t.id)}>Close</button>}</td></tr>)}</tbody></table></div></div></section>; }
function DocumentsView({ documents, assets, onAdd, onExport }) {
  const rows = documents.map((document) => { const asset = assets.find((item) => item.id === document.asetId); return { ...document, aset: asset?.nama || document.asetNama || document.asetId || '-', nopol: document.nopol || asset?.nomorPolisi || asset?.nomorRangka || '-' }; });
  return <section className="stack"><div className="metrics"><Metric label="Dokumen" value={documents.length} note="Vault aktif" /><Metric label="Hampir Habis" value={documents.filter((d) => d.status === 'Hampir Habis').length} note="Perlu update" tone="danger" /><Metric label="Valid" value={documents.filter((d) => d.status === 'Valid').length} note="Aman" tone="success" /><Metric label="Aset Terkait" value={new Set(documents.map((d) => d.asetId).filter(Boolean)).size} note="Linked by nopol/aset" /></div><div className="panel"><div className="panel-head"><div><h2>Digital Document Vault</h2><p>ID dokumen diisi manual. Pilih/cari nopol agar STNK, BPKB, KIR, asuransi, kontrak, BAST, dan file lain otomatis tersambung ke Aset 360.</p></div><div className="top-actions"><button className="btn primary" onClick={onAdd}>Upload Dokumen Aset</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><DataTable rows={rows} columns={[[ 'id', 'ID Dokumen' ], [ 'nopol', 'Nopol / Serial' ], [ 'aset', 'Aset' ], [ 'nama', 'Dokumen' ], [ 'tipe', 'Tipe' ], [ 'nomorDokumen', 'Nomor Dokumen' ], [ 'expired', 'Expired' ], [ 'pemilik', 'PIC' ], [ 'status', 'Status' ], [ 'file', 'File / Preview', (v, row) => row.fileData ? <a className="inline-link" href={row.fileData} target="_blank" rel="noreferrer">Review file</a> : (v || '-') ]]} /></div></section>;
}
function ApprovalsView({ approvals, roles, onAdd, onApprove, onExport }) { return <section className="approval-page stack"><div className="metrics"><Metric label="Role" value={roles.length} note="Permission user" /><Metric label="Rule Approval" value={approvals.length} note="Matriks aktif" /><Metric label="Pending" value={approvals.reduce((s, a) => s + numeric(a.pending), 0)} note="Menunggu approval" tone="danger" /><Metric label="Auto Escalation" value="Aktif" note="Rule based" /></div><div className="panel approval-role-panel"><div className="panel-head compact"><h2>Role & Permission User</h2></div><DataTable rows={roles} columns={[[ 'role', 'Role' ], [ 'akses', 'Akses' ], [ 'user', 'User' ]]} /></div><div className="panel approval-workflow-panel"><div className="panel-head compact vertical-actions"><div><h2>Approval Workflow</h2><p>Limit approval, pending approval, kontrol akses, dan tambah rule baru. Panel ini sengaja berada di bawah Role & Permission User, bukan disamping.</p></div><div className="module-actions below-title"><button className="btn primary" onClick={onAdd}>Tambah Approval</button><button className="btn ghost" onClick={onExport}>Export Excel</button></div></div><div className="table-wrap"><table><thead><tr><th>ID</th><th>Proses</th><th>Approval</th><th>Limit</th><th>Status</th><th>Pending</th><th>Aksi</th></tr></thead><tbody>{approvals.map((a) => <tr key={a.id}><td><b>{a.id}</b></td><td>{a.proses}</td><td>{a.approval}</td><td>{currency.format(numeric(a.limit))}</td><td>{a.status}</td><td><span className={`pill ${numeric(a.pending) ? 'amber' : 'green'}`}>{a.pending}</span></td><td>{numeric(a.pending) > 0 && <button className="btn small success" onClick={() => onApprove(a.id)}>Approve</button>}</td></tr>)}</tbody></table></div></div></section>; }
function ReportsView({ state, queue, notifications, onExport, onExportExcel }) { const collections = ['assets', 'contracts', 'workOrders', 'invoices', 'inventory', 'vendors', 'tickets', 'documents', 'checklists']; return <section className="stack"><div className="metrics"><Metric label="Total Log" value={state.auditLogs?.length || 0} note="Aktivitas terekam" /><Metric label="Queue Sync" value={queue.length} note="Menunggu Supabase" /><Metric label="Notif" value={notifications.length} note="Reminder aktif" /><Metric label="Collection" value={collections.length} note="Laporan tersedia" /></div><div className="panel"><div className="panel-head"><div><h2>Export Laporan Management</h2><p>Export per modul ke Excel atau backup JSON penuh untuk restore offline.</p></div><button className="btn ghost" onClick={onExport}>Export Backup JSON</button></div><div className="export-grid">{collections.map((c) => <button className="btn ghost" key={c} onClick={() => onExportExcel(c)}>Export {c}</button>)}</div></div><div className="panel"><div className="panel-head compact"><h2>Audit Trail</h2></div><div className="log-list">{(state.auditLogs || []).slice(0, 30).map((log) => <div className="log-row" key={log.id}><div><b>{log.aksi}</b><span>{new Date(log.waktu).toLocaleString('id-ID')} · {log.detail}</span></div><span className="pill slate">{log.user}</span></div>)}</div></div></section>; }
function AccountView({ settings, setSettings, roles, onLogin, onLogout, onForgotPassword, onUpdatePassword }) { const set = (key, value) => setSettings((prev) => ({ ...prev, [key]: value })); const [resetEmail, setResetEmail] = useState(settings.accountEmail || ''); const [newPassword, setNewPassword] = useState(''); const [confirmPassword, setConfirmPassword] = useState(''); async function handleUpdatePassword() { await onUpdatePassword(newPassword, confirmPassword); setNewPassword(''); setConfirmPassword(''); } return <section className="stack"><div className="account-grid"><div className="panel"><div className="panel-head compact"><h2>Status Akun</h2></div><div className="account-summary"><div className="summary-card"><span>Workspace</span><b>{settings.workspace || 'aerizen-main'}</b></div><div className="summary-card"><span>User aktif</span><b>{settings.loggedInUser || 'Belum login'}</b></div><div className="summary-card"><span>Role aktif</span><b>{settings.role || 'Super Admin'}</b></div><div className="summary-card"><span>Keamanan akun</span><b>Lupa Password + Lihat Password aktif</b></div></div></div><div className="panel"><div className="panel-head"><div><h2>Login Akun</h2><p>Menu akun tetap terpisah dari konfigurasi.</p></div></div><div className="form-grid"><Field label="Workspace" value={settings.workspace || ''} onChange={(v) => set('workspace', v)} /><Field label="Email akun" value={settings.accountEmail || ''} onChange={(v) => { set('accountEmail', v); if (!resetEmail) setResetEmail(v); }} placeholder="email@perusahaan.com" /><PasswordField label="Password" value={settings.accountPassword || ''} onChange={(v) => set('accountPassword', v)} placeholder="••••••••" /><Select label="Role" value={settings.role || 'Super Admin'} options={roles.map((r) => r.role)} onChange={(v) => set('role', v)} /></div><div className="modal-actions split-actions"><button className="btn primary" onClick={onLogin}>Login</button><button className="btn ghost" onClick={onLogout}>Logout</button><button className="btn ghost" onClick={() => onForgotPassword(settings.accountEmail || resetEmail)}>Lupa Password</button></div></div></div><div className="panel password-panel"><div className="panel-head"><div><h2>Lupa Password</h2><p>Kirim link reset ke email akun. Setelah membuka link dari email, kembali ke halaman Akun lalu isi password baru.</p></div><span className="pill blue">Supabase Auth</span></div><div className="form-grid"><Field label="Email reset password" value={resetEmail} onChange={setResetEmail} placeholder="email@perusahaan.com" /><div className="security-note"><b>Catatan</b><span>Fitur ini memakai Supabase Auth. Pastikan Supabase URL dan anon key sudah diisi di menu Konfigurasi.</span></div></div><div className="modal-actions split-actions"><button className="btn primary" onClick={() => onForgotPassword(resetEmail || settings.accountEmail)}>Kirim Link Reset Password</button></div><div className="password-divider" /><div className="form-grid"><PasswordField label="Password baru" value={newPassword} onChange={setNewPassword} placeholder="Minimal 6 karakter" /><PasswordField label="Konfirmasi password baru" value={confirmPassword} onChange={setConfirmPassword} placeholder="Ulangi password baru" /></div><div className="modal-actions split-actions"><button className="btn success" onClick={handleUpdatePassword}>Simpan Password Baru</button></div></div></section>; }
function ConfigurationView({ settings, setSettings, queue, syncStatus, onSave, onPush, onPull, onReset }) { const set = (key, value) => setSettings((prev) => ({ ...prev, [key]: value })); return <section className="stack"><div className="config-badges"><span className={`pill ${settings.realtime ? 'green' : 'amber'}`}>{settings.realtime ? 'Online realtime aktif' : 'Realtime nonaktif'}</span><span className="pill blue">Queue: {queue.length}</span><span className="pill slate">Workspace: {settings.workspace || 'aerizen-main'}</span></div><div className="panel config-surface"><div className="config-grid"><div><div className="panel-head compact"><div><h2>Realtime Sync + Offline Mode</h2><p>Offline + Supabase: data disimpan lokal dulu agar aplikasi tetap bisa dipakai offline. Saat internet tersedia, perubahan dikirim ke Supabase dan muncul realtime di akun lain.</p></div></div><div className="sync-card"><b>Status</b><p>{syncStatus}</p><span>Sinkron terakhir mengikuti aksi user dan status koneksi perangkat.</span></div></div><div className="config-form"><Field label="Supabase URL" value={settings.supabaseUrl || ''} onChange={(v) => set('supabaseUrl', v)} placeholder="https://xxxx.supabase.co" /><Field label="Supabase anon key" value={settings.supabaseKey || ''} onChange={(v) => set('supabaseKey', v)} placeholder="sb_publishable_xxx" /><Field label="Workspace" value={settings.workspace || ''} onChange={(v) => set('workspace', v)} placeholder="aerizen-main" /><div className="modal-actions split-actions"><button className="btn primary" onClick={onSave}>Simpan Konfigurasi</button><button className="btn success" onClick={onPush}>Sinkronkan Sekarang</button></div><div className="toggle-list"><label className="switch"><input type="checkbox" checked={settings.realtime} onChange={(e) => set('realtime', e.target.checked)} /><span>Realtime listener aktif</span></label><label className="switch"><input type="checkbox" checked={settings.autoSync} onChange={(e) => set('autoSync', e.target.checked)} /><span>Auto sync tiap 10 detik saat online</span></label></div><div className="modal-actions split-actions"><button className="btn ghost" onClick={onPull}>Pull Data</button><button className="btn danger" onClick={onReset}>Reset Demo</button></div></div></div></div><div className="panel"><h2>SQL Supabase yang dipakai</h2><pre>{`create table public.aerizen_records (\n  id text primary key,\n  collection text not null,\n  payload jsonb not null,\n  updated_at timestamptz default now(),\n  deleted_at timestamptz\n);`}</pre></div></section>; }

function AssetDrawer({ asset, documents, workOrders, contracts = [], invoices = [], onClose, onEdit, onWO }) { const stnk = stnkInfo(asset); const finance = numeric(asset.pendapatanBulanan) - numeric(asset.biayaBulanan); const docs = documents.filter((d) => documentMatchesAsset(d, asset)); const wos = workOrders.filter((wo) => wo.asetId === asset.id); const contract = getContractForAsset(asset, contracts); const relatedInvoices = invoices.filter((invoice) => invoice.contractId === contract?.id); return <div className="drawer-backdrop" onClick={onClose}><aside className="drawer" onClick={(e) => e.stopPropagation()}><div className="drawer-head"><div><p className="eyebrow">Aset 360</p><h2>{asset.nama}</h2><p className="muted">{asset.id} · {asset.qrCode}</p></div><button onClick={onClose}>×</button></div><div className="asset-360-hero"><AssetPhotoCard asset={asset} contract={contract} invoices={relatedInvoices} documents={docs} /><div className="asset-360-summary"><h3>Ringkasan Aset</h3><p>Detail aset lengkap: foto, dokumen, GPS yang bisa dicentang, KM/odometer khusus, kontrak, Work Order (WO), dan histori aktivitas.</p><div className="quick-summary two-cols asset-summary-grid"><div className="asset-summary-card"><b className="asset-summary-value" title={asset.status}>{asset.status}</b><span className="asset-summary-label">Status unit</span></div><div className="asset-summary-card"><b className="asset-summary-value asset-summary-nopol" title={asset.nomorPolisi || '-'}>{asset.nomorPolisi || '-'}</b><span className="asset-summary-label">Nopol / serial</span></div><div className="asset-summary-card"><b className="asset-summary-value" title={asset.gpsAktif ? 'GPS aktif' : 'GPS belum aktif'}>{asset.gpsAktif ? 'GPS aktif' : 'GPS belum aktif'}</b><span className="asset-summary-label">Status GPS</span></div><div className="asset-summary-card"><b className="asset-summary-value asset-summary-km" title={asset.km ? `${asset.km} km` : '-'}>{asset.km ? `${Number(asset.km).toLocaleString('id-ID')} km` : '-'}</b><span className="asset-summary-label">KM / odometer</span></div></div></div></div>{stnk.due && <div className="stnk-alert"><div><b>Notifikasi merah STNK</b><span>{stnk.message}</span></div><span className="pill red">Wajib diperpanjang</span></div>}<div className="detail-grid"><Info label="Status" value={asset.status} /><Info label="Kategori" value={asset.kategori} /><Info label="Cabang" value={asset.cabang} /><Info label="Pemegang" value={asset.pemegang} /><Info label="Nopol / Serial" value={asset.nomorPolisi || asset.nomorRangka} /><Info label="QR / Barcode" value={asset.qrCode} /><Info label="Nomor Rangka" value={asset.nomorRangka} /><Info label="Nomor Mesin / IMEI" value={asset.nomorMesin} /><Info label="GPS" value={asset.gpsAktif ? 'Aktif / dicentang' : 'Tidak aktif'} /><Info label="ID Perangkat GPS" value={asset.gpsDeviceId || '-'} /><Info label="Provider GPS" value={asset.gpsProvider || '-'} /><Info label="KM / Odometer" value={asset.km ? `${asset.km} km` : '-'} /><Info label="STNK" value={asset.tanggalSTNK || '-'} /><Info label="Status STNK" value={asset.kategori === 'Kendaraan' ? stnk.message : '-'} /><Info label="KIR" value={asset.tanggalKIR || '-'} /><Info label="Pajak" value={asset.tanggalPajak || '-'} /><Info label="Asuransi" value={asset.tanggalAsuransi || '-'} /><Info label="Service Berikutnya" value={asset.tanggalServiceBerikutnya || '-'} /><Info label="BPKB" value={asset.memilikiBPKB} /><Info label="Nilai Buku" value={currency.format(numeric(asset.nilaiBuku))} /><Info label="Profit Bulanan" value={currency.format(finance)} /><Info label="Telemetri" value={asset.telemetri || assetTelemetry(asset)} /></div><div className="timeline"><h3>Timeline Aktivitas</h3>{(asset.riwayat || []).map((item, index) => <div className="timeline-row" key={`${item}-${index}`}><b>{index + 1}</b><span>{item}</span></div>)}</div><div className="drawer-section"><h3>Dokumen Aset</h3>{docs.length ? docs.map((d) => <div className="mini-row" key={d.id}><b>{d.tipe}</b><span>{d.nama} · {d.status} {d.fileData && <a className="inline-link" href={d.fileData} target="_blank" rel="noreferrer">Review</a>}</span></div>) : <p className="muted">Belum ada dokumen terkait.</p>}</div><div className="drawer-section"><h3>History Work Order (WO)</h3>{wos.length ? wos.map((wo) => <div className="mini-row" key={wo.id}><b>{wo.id}</b><span>{wo.status} · {currency.format(numeric(wo.biaya))}</span></div>) : <p className="muted">Belum ada WO terkait.</p>}</div><div className="drawer-actions"><button className="btn primary" onClick={() => onEdit(asset)}>Edit Aset</button><button className="btn ghost" onClick={() => onWO(asset)}>Buat Work Order</button><button className="btn ghost" onClick={() => navigator.clipboard?.writeText(asset.qrCode)}>Copy QR</button></div></aside></div>; }
function Info({ label, value }) { return <div className="info"><span>{label}</span><b>{value || '-'}</b></div>; }
function AssetForm({ asset, onClose, onSave }) { const isNew = !asset?.id; const [form, setForm] = useState(normalizeAsset(asset || {})); const set = (key, value) => setForm((prev) => ({ ...prev, [key]: value })); function readAsDataUrl(file, callback) { if (!file) return; const reader = new FileReader(); reader.onload = () => callback(reader.result); reader.readAsDataURL(file); } function imageToDataUrl(file) { readAsDataUrl(file, (result) => set('gambar', result)); } function docsToDataUrl(files) { const list = Array.from(files || []); if (!list.length) return; list.forEach((file) => readAsDataUrl(file, (result) => setForm((prev) => ({ ...prev, assetDocuments: [...(prev.assetDocuments || []), { nama: file.name, tipe: 'Dokumen Aset', file: file.name, fileName: file.name, fileType: file.type, fileData: result, expired: '', status: 'Valid' }] })))); } return <Modal title={isNew ? 'Tambah Aset - ID Otomatis' : 'Edit Aset'} onClose={onClose}><div className="form-grid"><Field label="ID Aset Otomatis" value={form.id} onChange={() => null} readOnly /><Field label="Nama Aset" value={form.nama} onChange={(v) => set('nama', v)} /><Select label="Kategori" value={form.kategori} options={categoryOptions} onChange={(v) => setForm((prev) => normalizeAsset({ ...prev, kategori: v, id: prev.id || generateAssetId(v) }))} /><Select label="Status" value={form.status} options={statusOptions} onChange={(v) => set('status', v)} /><Field label="Tipe" value={form.tipe} onChange={(v) => set('tipe', v)} /><Field label="Kondisi" value={form.kondisi} onChange={(v) => set('kondisi', v)} /><Field label="Cabang / Wilayah" value={form.cabang} onChange={(v) => { set('cabang', v); set('wilayah', v); }} /><Field label="Lokasi" value={form.lokasi} onChange={(v) => set('lokasi', v)} /><Field label="Pemegang / Customer" value={form.pemegang} onChange={(v) => { set('pemegang', v); set('perusahaanUser', v); }} /><Field label="Alamat User" value={form.alamatUser} onChange={(v) => set('alamatUser', v)} /><Field label="Warna" value={form.warna} onChange={(v) => set('warna', v)} /><Field label="Tahun Kendaraan" value={form.tahunKendaraan} onChange={(v) => set('tahunKendaraan', v)} /><Field label="Nomor Polisi" value={form.nomorPolisi} onChange={(v) => set('nomorPolisi', v)} /><Field label="QR / Barcode Otomatis" value={form.qrCode} onChange={() => null} readOnly /><Field label="Nomor Rangka / Serial" value={form.nomorRangka} onChange={(v) => set('nomorRangka', v)} /><Field label="Nomor Mesin / IMEI" value={form.nomorMesin} onChange={(v) => set('nomorMesin', v)} /><label className="field switch-field"><span>GPS</span><div className="check-row"><input type="checkbox" checked={!!form.gpsAktif} onChange={(e) => set('gpsAktif', e.target.checked)} /><b>Centang jika GPS terpasang / aktif</b></div></label><Field label="KM / Odometer" type="number" value={form.km} onChange={(v) => { set('km', v); set('odometerKm', v); }} /><Field label="ID Perangkat GPS" value={form.gpsDeviceId} onChange={(v) => set('gpsDeviceId', v)} /><Field label="Provider GPS" value={form.gpsProvider} onChange={(v) => set('gpsProvider', v)} /><Field label="Tanggal STNK" type="date" value={form.tanggalSTNK} onChange={(v) => set('tanggalSTNK', v)} /><Field label="Tanggal KIR" type="date" value={form.tanggalKIR} onChange={(v) => set('tanggalKIR', v)} /><Field label="Tanggal Pajak" type="date" value={form.tanggalPajak} onChange={(v) => set('tanggalPajak', v)} /><Field label="Tanggal Asuransi" type="date" value={form.tanggalAsuransi} onChange={(v) => set('tanggalAsuransi', v)} /><Field label="Service Berikutnya" type="date" value={form.tanggalServiceBerikutnya} onChange={(v) => set('tanggalServiceBerikutnya', v)} /><Select label="Memiliki BPKB" value={form.memilikiBPKB} options={['Ya', 'Tidak', 'Belum diisi', '-']} onChange={(v) => set('memilikiBPKB', v)} /><Field label="Utilisasi %" type="number" value={form.utilisasi} onChange={(v) => set('utilisasi', v)} /><Field label="Risiko %" type="number" value={form.risiko} onChange={(v) => set('risiko', v)} /><Field label="Nilai Buku" type="number" value={form.nilaiBuku} onChange={(v) => set('nilaiBuku', v)} /><Field label="Pendapatan Bulanan" type="number" value={form.pendapatanBulanan} onChange={(v) => set('pendapatanBulanan', v)} /><Field label="Biaya Bulanan" type="number" value={form.biayaBulanan} onChange={(v) => set('biayaBulanan', v)} /><label className="field full"><span>Upload Foto Aset</span><input type="file" accept="image/*" onChange={(e) => imageToDataUrl(e.target.files?.[0])} /><small className="field-help">Foto aset tetap disimpan dan tampilan kartu tidak dibuat gepeng.</small></label><label className="field full"><span>Upload Dokumen Aset</span><input type="file" multiple accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={(e) => docsToDataUrl(e.target.files)} /><small className="field-help">Tempat upload dokumen aset: STNK, BPKB, KIR, polis asuransi, kontrak, invoice service, garansi, BAST, dan dokumen lain.</small>{!!form.assetDocuments?.length && <div className="upload-list">{form.assetDocuments.map((doc, index) => <span key={`${doc.fileName}-${index}`}>{doc.fileName}</span>)}</div>}</label><TextArea label="Aksi Berikutnya" value={form.aksiBerikutnya} onChange={(v) => set('aksiBerikutnya', v)} /><TextArea label="Telemetri / Catatan" value={form.telemetri || assetTelemetry(form)} onChange={(v) => set('telemetri', v)} /></div><div className="modal-actions"><button className="btn ghost" onClick={onClose}>Batal</button><button className="btn primary" onClick={() => onSave({ ...form, telemetri: form.telemetri || assetTelemetry(form) })}>Simpan</button></div></Modal>; }

function ContractForm({ form, assets, state, onClose, onSave }) {
  const [draft, setDraft] = useState(normalizeContract(form || contractDefaults(state), state));
  const set = (key, value) => setDraft((prev) => ({ ...prev, [key]: value }));
  const selectedAsset = assets.find((asset) => asset.id === draft.asetId) || {};
  const duration = numeric(draft.durasiBulan) || monthsBetween(draft.mulai, draft.selesai) || 1;
  const subtotal = numeric(draft.nilaiBulanan) * duration + numeric(draft.biayaAdmin) - numeric(draft.diskon);
  const ppn = Math.round(subtotal * (Number(draft.ppnPersen || 0) / 100));
  const calculatedTotal = Math.max(0, subtotal + ppn);
  function chooseAsset(asetId) {
    const asset = assets.find((item) => item.id === asetId) || {};
    setDraft((prev) => ({ ...prev, asetId, aset: asset.nama || '', nopol: asset.nomorPolisi || asset.nomorRangka || '', nomorRangka: asset.nomorRangka || '', cabang: asset.cabang || '', lokasiSerahTerima: asset.lokasi || prev.lokasiSerahTerima, kmAwal: asset.km || prev.kmAwal }));
  }
  function updatePeriod(key, value) {
    setDraft((prev) => { const next = { ...prev, [key]: value }; next.durasiBulan = monthsBetween(next.mulai, next.selesai) || prev.durasiBulan || 1; return next; });
  }
  function contractFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setDraft((prev) => ({ ...prev, dokumen: file.name, dokumenFileData: reader.result, dokumenFileType: file.type }));
    reader.readAsDataURL(file);
  }
  function submit() {
    if (!String(draft.id || '').trim()) { window.alert('Nomor kontrak wajib diisi.'); return; }
    if (!String(draft.customer || '').trim()) { window.alert('Nama customer wajib diisi.'); return; }
    if (!draft.asetId) { window.alert('Pilih aset/nopol untuk kontrak.'); return; }
    onSave({ ...draft, durasiBulan: duration, totalKontrak: calculatedTotal });
  }
  return <Modal title="Tambah / Edit Kontrak Sewa" onClose={onClose}>
    <div className="form-section-title"><b>Identitas Kontrak & Customer</b><span>Nomor kontrak, PIC, kontak, alamat, dan data legal customer.</span></div>
    <div className="form-grid">
      <Field label="No. Kontrak Otomatis" value={draft.id} onChange={() => null} readOnly />
      <Select label="Status Kontrak" value={draft.status} options={['Draft','Menunggu Dokumen','Menunggu Approval','Aktif','Selesai','Dibatalkan']} onChange={(v) => set('status', v)} />
      <Field label="Nama Customer / Perusahaan" value={draft.customer} onChange={(v) => set('customer', v)} />
      <Field label="PIC Customer" value={draft.picCustomer} onChange={(v) => set('picCustomer', v)} />
      <Field label="Telepon / WhatsApp Customer" type="tel" value={draft.teleponCustomer} onChange={(v) => set('teleponCustomer', v)} />
      <Field label="Email Customer" type="email" value={draft.emailCustomer} onChange={(v) => set('emailCustomer', v)} />
      <Field label="NPWP Customer" value={draft.npwpCustomer} onChange={(v) => set('npwpCustomer', v)} />
      <TextArea label="Alamat Customer" value={draft.alamatCustomer} onChange={(v) => set('alamatCustomer', v)} />
    </div>
    <div className="form-section-title"><b>Aset, Nopol & Serah Terima</b><span>Pilih unit agar nama aset, nopol/serial, rangka, cabang, lokasi, dan KM awal tertarik otomatis.</span></div>
    <div className="form-grid">
      <label className="field"><span>Aset / Unit</span><select value={draft.asetId || ''} onChange={(e) => chooseAsset(e.target.value)}><option value="">Pilih aset / nopol</option>{assets.map((asset) => <option key={asset.id} value={asset.id}>{asset.nomorPolisi || asset.nomorRangka || '-'} · {asset.nama} · {asset.id}</option>)}</select></label>
      <Field label="Nama Aset" value={selectedAsset.nama || draft.aset} onChange={() => null} readOnly />
      <Field label="Nopol / Serial" value={selectedAsset.nomorPolisi || selectedAsset.nomorRangka || draft.nopol} onChange={() => null} readOnly />
      <Field label="Nomor Rangka / Serial" value={selectedAsset.nomorRangka || draft.nomorRangka} onChange={() => null} readOnly />
      <Field label="Cabang / Wilayah" value={selectedAsset.cabang || draft.cabang} onChange={() => null} readOnly />
      <Field label="Lokasi Serah Terima" value={draft.lokasiSerahTerima} onChange={(v) => set('lokasiSerahTerima', v)} />
      <Field label="KM / Odometer Awal" type="number" value={draft.kmAwal} onChange={(v) => set('kmAwal', v)} />
      <Field label="Batas KM per Bulan" type="number" value={draft.batasKmBulanan} onChange={(v) => set('batasKmBulanan', v)} />
      <Field label="Tarif Kelebihan per KM" type="number" value={draft.tarifOverKm} onChange={(v) => set('tarifOverKm', v)} />
    </div>
    <div className="form-section-title"><b>Periode, Billing & Nilai Kontrak</b><span>Periode, PO, termin, deposit, biaya, diskon, PPN, dan total kontrak.</span></div>
    <div className="form-grid">
      <Field label="Tanggal Mulai" type="date" value={draft.mulai} onChange={(v) => updatePeriod('mulai', v)} />
      <Field label="Tanggal Selesai" type="date" value={draft.selesai} onChange={(v) => updatePeriod('selesai', v)} />
      <Field label="Durasi (Bulan)" type="number" value={duration} onChange={(v) => set('durasiBulan', v)} />
      <Select label="Billing Cycle" value={draft.billingCycle} options={['Bulanan','Tahunan','Sekali Bayar','Custom']} onChange={(v) => set('billingCycle', v)} />
      <Field label="Nomor PO Customer" value={draft.nomorPO} onChange={(v) => set('nomorPO', v)} />
      <Field label="Tanggal PO" type="date" value={draft.tanggalPO} onChange={(v) => set('tanggalPO', v)} />
      <Field label="Termin Pembayaran" value={draft.terminPembayaran} onChange={(v) => set('terminPembayaran', v)} />
      <Field label="Tanggal Jatuh Tempo Tiap Bulan" type="number" value={draft.tanggalJatuhTempo} onChange={(v) => set('tanggalJatuhTempo', v)} />
      <Field label="Nilai Sewa per Bulan" type="number" value={draft.nilaiBulanan} onChange={(v) => set('nilaiBulanan', v)} />
      <Field label="Deposit" type="number" value={draft.deposit} onChange={(v) => set('deposit', v)} />
      <Field label="Biaya Administrasi" type="number" value={draft.biayaAdmin} onChange={(v) => set('biayaAdmin', v)} />
      <Field label="Diskon" type="number" value={draft.diskon} onChange={(v) => set('diskon', v)} />
      <Field label="PPN (%)" type="number" value={draft.ppnPersen} onChange={(v) => set('ppnPersen', v)} />
      <Field label="Total Nilai Kontrak" value={currency.format(calculatedTotal)} onChange={() => null} readOnly />
    </div>
    <div className="form-section-title"><b>Layanan, Dokumen & Approval</b><span>Coverage layanan, replacement unit, asuransi, pengemudi, approval, file kontrak, BAST, dan renewal.</span></div>
    <div className="form-grid">
      <TextArea label="Cakupan Service / Maintenance" value={draft.cakupanService} onChange={(v) => set('cakupanService', v)} />
      <Field label="Ketentuan Unit Pengganti" value={draft.unitPengganti} onChange={(v) => set('unitPengganti', v)} />
      <Field label="Asuransi" value={draft.asuransi} onChange={(v) => set('asuransi', v)} />
      <Field label="Pengemudi" value={draft.pengemudi} onChange={(v) => set('pengemudi', v)} />
      <Select label="Status Approval" value={draft.approvalStatus} options={['Belum Diajukan','Menunggu Approval','Disetujui','Ditolak','Tidak Perlu Approval']} onChange={(v) => set('approvalStatus', v)} />
      <Field label="Nomor / Nama BAST" value={draft.bast} onChange={(v) => set('bast', v)} />
      <label className="field full"><span>Upload File Kontrak</span><input type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx" onChange={(e) => contractFile(e.target.files?.[0])} /><small className="field-help">{draft.dokumen || 'Belum ada file kontrak.'}</small></label>
      <TextArea label="Renewal / Tindak Lanjut" value={draft.renewal} onChange={(v) => set('renewal', v)} />
      <TextArea label="Catatan Kontrak" value={draft.catatan} onChange={(v) => set('catatan', v)} />
    </div>
    <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Batal</button><button className="btn primary" onClick={submit}>Simpan Kontrak</button></div>
  </Modal>;
}

function DocumentForm({ form, assets, onClose, onSave }) {
  const initialAsset = assets.find((asset) => asset.id === form?.asetId) || {};
  const [draft, setDraft] = useState(normalizeDocument({ ...form, id: form?.id || '', asetNama: initialAsset.nama || form?.asetNama || '', nopol: initialAsset.nomorPolisi || initialAsset.nomorRangka || form?.nopol || '' }));
  const [assetSearch, setAssetSearch] = useState(initialAsset.id ? `${initialAsset.nomorPolisi || initialAsset.nomorRangka || '-'} · ${initialAsset.nama}` : '');
  const set = (key, value) => setDraft((prev) => ({ ...prev, [key]: value }));
  const selectedAsset = assets.find((asset) => asset.id === draft.asetId) || {};
  const normalizedAssetSearch = compactReference(assetSearch);
  const filteredAssets = assets.filter((asset) => {
    if (!normalizedAssetSearch) return true;
    return [asset.nomorPolisi, asset.nomorRangka, asset.id, asset.nama, asset.nomorMesin]
      .map(compactReference)
      .some((reference) => reference && reference.includes(normalizedAssetSearch));
  });
  const idSuggestion = suggestDocumentId(selectedAsset, draft.tipe, draft.fileName || draft.nama);
  function linkAsset(asset) {
    if (!asset) return;
    setAssetSearch(`${asset.nomorPolisi || asset.nomorRangka || '-'} · ${asset.nama}`);
    setDraft((prev) => ({ ...prev, asetId: asset.id, asetNama: asset.nama, nopol: asset.nomorPolisi || asset.nomorRangka || '' }));
  }
  function chooseAssetBySearch(value) {
    setAssetSearch(value);
    const compact = compactReference(value);
    if (!compact) return;
    const exactMatch = assets.find((asset) => [asset.nomorPolisi, asset.nomorRangka, asset.id]
      .map(compactReference)
      .some((reference) => reference && reference === compact));
    if (exactMatch) linkAsset(exactMatch);
  }
  function fileToDataUrl(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setDraft((prev) => ({ ...prev, nama: !prev.nama || prev.nama === 'Dokumen Aset' ? file.name : prev.nama, file: file.name, fileName: file.name, fileType: file.type, fileData: reader.result }));
    reader.readAsDataURL(file);
  }
  function submit() {
    if (!String(draft.id || '').trim()) { window.alert('ID Dokumen wajib diisi manual, misalnya BPKB-B184AZN.'); return; }
    if (!draft.asetId) { window.alert('Pilih atau cari nopol agar dokumen terhubung ke aset.'); return; }
    onSave(draft);
  }
  return <Modal title="Upload Dokumen Aset" onClose={onClose}>
    <div className="form-section-title"><b>Identitas Dokumen dan Aset</b></div>
    <div className="form-grid">
      <Field label="ID Dokumen (Manual)" value={draft.id} placeholder="Contoh: BPKB-B184AZN" onChange={(v) => set('id', v.toUpperCase())} />
      <div className="field asset-search-select">
        <span>Cari dan Pilih Aset</span>
        <input value={assetSearch} placeholder="Cari nopol, serial, ID, atau nama aset" onChange={(e) => chooseAssetBySearch(e.target.value)} />
        <select value={draft.asetId || ''} onChange={(e) => linkAsset(assets.find((asset) => asset.id === e.target.value))}>
          <option value="">Pilih aset</option>
          {filteredAssets.map((asset) => <option key={asset.id} value={asset.id}>{asset.nomorPolisi || asset.nomorRangka || '-'} · {asset.nama}</option>)}
        </select>
        {!filteredAssets.length && <small className="field-error">Aset tidak ditemukan.</small>}
      </div>
      <Field label="ID Aset Tertaut" value={draft.asetId} onChange={() => null} readOnly />
      <Field label="Nama Aset" value={selectedAsset.nama || draft.asetNama} onChange={() => null} readOnly />
      <Field label="Nopol / Serial Tertaut" value={selectedAsset.nomorPolisi || selectedAsset.nomorRangka || draft.nopol} onChange={() => null} readOnly />
      <div className="field full document-id-suggestion"><span>Saran Format ID</span><div><code>{idSuggestion}</code><button type="button" className="btn small ghost" onClick={() => set('id', idSuggestion)}>Pakai Saran</button></div></div>
    </div>
    <div className="form-section-title"><b>Informasi Dokumen</b></div>
    <div className="form-grid">
      <Field label="Nama Dokumen" value={draft.nama} onChange={(v) => set('nama', v)} />
      <Select label="Tipe Dokumen" value={draft.tipe} options={['Dokumen Aset','STNK','BPKB','KIR','Pajak','Asuransi','Garansi IT','Kontrak','Invoice Service','BAST','Surat Jalan','Lainnya']} onChange={(v) => set('tipe', v)} />
      <Field label="Nomor Dokumen Resmi" value={draft.nomorDokumen} onChange={(v) => set('nomorDokumen', v)} />
      <Field label="Tanggal Terbit" type="date" value={draft.tanggalTerbit} onChange={(v) => set('tanggalTerbit', v)} />
      <Field label="Expired / Masa Berlaku" type="date" value={draft.expired} onChange={(v) => set('expired', v)} />
      <Field label="PIC / Pemilik Dokumen" value={draft.pemilik} onChange={(v) => set('pemilik', v)} />
      <Select label="Status" value={draft.status} options={['Valid','Hampir Habis','Expired','Draft','Menunggu Dokumen','Perlu Review']} onChange={(v) => set('status', v)} />
      <label className="field full"><span>Upload File Dokumen</span><input type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={(e) => fileToDataUrl(e.target.files?.[0])} /><small className="field-help">Setelah disimpan, file tampil di Digital Document Vault dan Aset 360 sesuai nopol yang dipilih.</small>{draft.fileName && <div className="upload-list"><span>{draft.fileName}</span></div>}</label>
      <TextArea label="Keterangan Dokumen" value={draft.keterangan} onChange={(v) => set('keterangan', v)} />
    </div>
    <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Batal</button><button className="btn primary" onClick={submit}>Simpan & Hubungkan Dokumen</button></div>
  </Modal>;
}
function InvoiceForm({ values, state, isEditing = false, onClose, onSave }) {
  const [draft, setDraft] = useState(normalizeInvoice(values || invoiceDefaults(state), state));
  const contracts = state.contracts || [];
  const assets = state.assets || [];
  const set = (key, value) => setDraft((prev) => ({ ...prev, [key]: value }));
  const selectedContract = contracts.find((item) => item.id === draft.contractId) || {};
  const selectedAsset = assets.find((item) => item.id === draft.assetId) || assets.find((item) => item.id === selectedContract.asetId) || {};
  const subtotal = Math.max(0, numeric(draft.baseAmount) + numeric(draft.overKmCharge) + numeric(draft.fuelCharge) + numeric(draft.penaltyCharge) + numeric(draft.serviceCharge) + numeric(draft.otherCharge) - numeric(draft.discount));
  const taxAmount = Math.round(subtotal * Number(draft.taxRate || 0) / 100);
  const withholdingAmount = Math.round(subtotal * Number(draft.withholdingRate || 0) / 100);
  const amount = subtotal + taxAmount;
  const remaining = Math.max(0, amount - numeric(draft.paid) - withholdingAmount);
  const calculatedStatus = remaining <= 0 ? 'Paid' : numeric(draft.paid) > 0 || withholdingAmount > 0 ? 'Partial' : 'Unpaid';
  function selectContract(contractId) {
    const contract = contracts.find((item) => item.id === contractId) || {};
    const asset = assets.find((item) => item.id === contract.asetId) || {};
    const baseAmount = numeric(contract.nilaiBulanan);
    setDraft((prev) => ({ ...prev, contractId, customer: contract.customer || prev.customer, customerPic: contract.picCustomer || '', customerEmail: contract.emailCustomer || '', customerPhone: contract.teleponCustomer || '', billingAddress: contract.alamatCustomer || '', npwpCustomer: contract.npwpCustomer || '', assetId: asset.id || contract.asetId || '', asset: asset.nama || contract.aset || '', nopol: asset.nomorPolisi || asset.nomorRangka || contract.nopol || '', poCustomer: contract.nomorPO || '', paymentTerm: contract.terminPembayaran || prev.paymentTerm, dueDate: addDays(prev.invoiceDate || today, paymentTermDays(contract.terminPembayaran || prev.paymentTerm)), baseAmount }));
  }
  function changeInvoiceDate(value) { setDraft((prev) => ({ ...prev, invoiceDate: value, dueDate: addDays(value, paymentTermDays(prev.paymentTerm)) })); }
  function changePaymentTerm(value) { setDraft((prev) => ({ ...prev, paymentTerm: value, dueDate: addDays(prev.invoiceDate || today, paymentTermDays(value)) })); }
  function fileToDataUrl(file, nameKey, dataKey) { if (!file) return; const reader = new FileReader(); reader.onload = () => setDraft((prev) => ({ ...prev, [nameKey]: file.name, [dataKey]: reader.result })); reader.readAsDataURL(file); }
  const savePayload = { ...draft, assetId: selectedAsset.id || draft.assetId, asset: selectedAsset.nama || draft.asset, nopol: selectedAsset.nomorPolisi || selectedAsset.nomorRangka || draft.nopol, subtotal, taxAmount, withholdingAmount, amount, remaining, status: calculatedStatus };
  return <Modal title={isEditing ? `Edit Invoice ${values.id}` : 'Tambah Invoice'} onClose={onClose}>
    <div className="form-section-title"><b>Identitas Invoice</b><span>Nomor invoice, kontrak, customer, aset, periode, dan jatuh tempo.</span></div>
    <div className="form-grid">
      <Field label="ID Invoice Otomatis" value={draft.id} onChange={() => null} readOnly />
      <Select label="Nomor Kontrak" value={draft.contractId} options={['', ...contracts.map((item) => item.id)]} onChange={selectContract} />
      <Field label="Tanggal Invoice" type="date" value={draft.invoiceDate} onChange={changeInvoiceDate} />
      <Field label="Periode Tagihan" value={draft.periode} onChange={(v) => set('periode', v)} placeholder="Contoh: Juni 2026" />
      <Field label="Periode Mulai" type="date" value={draft.periodStart} onChange={(v) => set('periodStart', v)} />
      <Field label="Periode Selesai" type="date" value={draft.periodEnd} onChange={(v) => set('periodEnd', v)} />
      <Field label="Termin Pembayaran" value={draft.paymentTerm} onChange={changePaymentTerm} placeholder="Contoh: 30 hari" />
      <Field label="Jatuh Tempo" type="date" value={draft.dueDate} onChange={(v) => set('dueDate', v)} />
      <Field label="Nomor PO Customer" value={draft.poCustomer} onChange={(v) => set('poCustomer', v)} />
      <Field label="Status Pembayaran Otomatis" value={calculatedStatus} onChange={() => null} readOnly />
    </div>
    <div className="form-section-title"><b>Customer & Aset Tertagih</b><span>Data otomatis ditarik dari kontrak dan tetap dapat dilengkapi.</span></div>
    <div className="form-grid">
      <Field label="Customer" value={draft.customer} onChange={(v) => set('customer', v)} />
      <Field label="PIC Customer" value={draft.customerPic} onChange={(v) => set('customerPic', v)} />
      <Field label="Email Customer" type="email" value={draft.customerEmail} onChange={(v) => set('customerEmail', v)} />
      <Field label="Telepon Customer" type="tel" value={draft.customerPhone} onChange={(v) => set('customerPhone', v)} />
      <Field label="NPWP Customer" value={draft.npwpCustomer} onChange={(v) => set('npwpCustomer', v)} />
      <Field label="ID Aset" value={selectedAsset.id || draft.assetId} onChange={() => null} readOnly />
      <Field label="Nama Aset" value={selectedAsset.nama || draft.asset} onChange={() => null} readOnly />
      <Field label="Nopol / Serial" value={selectedAsset.nomorPolisi || selectedAsset.nomorRangka || draft.nopol} onChange={() => null} readOnly />
      <TextArea label="Alamat Penagihan" value={draft.billingAddress} onChange={(v) => set('billingAddress', v)} />
    </div>
    <div className="form-section-title"><b>Rincian Nilai Tagihan</b><span>Total dihitung otomatis dari sewa, biaya tambahan, diskon, PPN, dan potongan pajak.</span></div>
    <div className="form-grid">
      <Field label="Nilai Sewa / Tagihan Pokok" type="number" value={draft.baseAmount} onChange={(v) => set('baseAmount', v)} />
      <Field label="Biaya Kelebihan KM" type="number" value={draft.overKmCharge} onChange={(v) => set('overKmCharge', v)} />
      <Field label="Biaya BBM" type="number" value={draft.fuelCharge} onChange={(v) => set('fuelCharge', v)} />
      <Field label="Denda / Penalty" type="number" value={draft.penaltyCharge} onChange={(v) => set('penaltyCharge', v)} />
      <Field label="Biaya Service / Maintenance" type="number" value={draft.serviceCharge} onChange={(v) => set('serviceCharge', v)} />
      <Field label="Biaya Lain-lain" type="number" value={draft.otherCharge} onChange={(v) => set('otherCharge', v)} />
      <Field label="Diskon" type="number" value={draft.discount} onChange={(v) => set('discount', v)} />
      <Field label="Subtotal" type="number" value={subtotal} onChange={() => null} readOnly />
      <Field label="PPN (%)" type="number" value={draft.taxRate} onChange={(v) => set('taxRate', v)} />
      <Field label="Nominal PPN" type="number" value={taxAmount} onChange={() => null} readOnly />
      <Field label="PPh / Potongan (%)" type="number" value={draft.withholdingRate} onChange={(v) => set('withholdingRate', v)} />
      <Field label="Nominal Potongan PPh" type="number" value={withholdingAmount} onChange={() => null} readOnly />
      <Field label="Total Invoice" type="number" value={amount} onChange={() => null} readOnly />
      <Field label="Jumlah Dibayar" type="number" value={draft.paid} onChange={(v) => set('paid', v)} />
      <Field label="Sisa Piutang" type="number" value={remaining} onChange={() => null} readOnly />
    </div>
    <div className="form-section-title"><b>Pembayaran, Pajak & Lampiran</b><span>Rekonsiliasi pembayaran, faktur pajak, file invoice, dan catatan.</span></div>
    <div className="form-grid">
      <Field label="Tanggal Pembayaran" type="date" value={draft.paymentDate} onChange={(v) => set('paymentDate', v)} />
      <Select label="Metode Pembayaran" value={draft.paymentMethod} options={['Transfer Bank','Virtual Account','Giro','Cash','Kartu','Lainnya']} onChange={(v) => set('paymentMethod', v)} />
      <Field label="Rekening Tujuan" value={draft.bankAccount} onChange={(v) => set('bankAccount', v)} />
      <Field label="Referensi Pembayaran" value={draft.paymentReference} onChange={(v) => set('paymentReference', v)} />
      <Field label="Nomor Faktur Pajak" value={draft.taxInvoiceNumber} onChange={(v) => set('taxInvoiceNumber', v)} />
      <label className="field full"><span>Upload File Invoice</span><input type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={(e) => fileToDataUrl(e.target.files?.[0], 'invoiceFileName', 'invoiceFileData')} /><small className="field-help">{draft.invoiceFileName || 'Belum ada file invoice'}</small></label>
      <label className="field full"><span>Upload Faktur Pajak</span><input type="file" accept=".pdf,.jpg,.jpeg,.png" onChange={(e) => fileToDataUrl(e.target.files?.[0], 'taxInvoiceFileName', 'taxInvoiceFileData')} /><small className="field-help">{draft.taxInvoiceFileName || 'Belum ada faktur pajak'}</small></label>
      <TextArea label="Catatan Invoice" value={draft.notes} onChange={(v) => set('notes', v)} />
    </div>
    <div className="invoice-summary"><div><span>Subtotal</span><b>{currency.format(subtotal)}</b></div><div><span>PPN</span><b>{currency.format(taxAmount)}</b></div><div><span>Total Invoice</span><b>{currency.format(amount)}</b></div><div><span>Sisa Piutang</span><b>{currency.format(remaining)}</b></div></div>
    <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Batal</button><button className="btn primary" onClick={() => onSave(savePayload)}>Simpan Invoice</button></div>
  </Modal>;
}

function QuickAddForm({ collection, values, state, onClose, onSave }) {
  const [draft, setDraft] = useState(values || quickDefaults(collection, state));
  const set = (key, value) => setDraft((prev) => {
    const next = { ...prev, [key]: value };
    if (collection === 'checklists' && key === 'tipe') next.id = nextChecklistId(state, value);
    if (collection === 'inventory' && key === 'nama') next.id = nextInventoryId(state, value);
    if (collection === 'inventory' && key === 'hargaBeli' && !numeric(prev.nilai)) next.nilai = value;
    if (collection === 'checklists' && key === 'asetId') {
      const asset = (state.assets || []).find((item) => item.id === value);
      next.aset = asset?.nama || '';
      if (!next.km) next.km = asset?.km || '';
    }
    return next;
  });
  const fields = quickFields(collection, state);
  const config = quickModuleConfig(collection);
  return <Modal title={config.title} onClose={onClose}><div className="form-grid"><Field label="ID Otomatis" value={draft.id} onChange={() => null} readOnly />{fields.map(([key, labelText, type = 'text', options = []]) => type === 'select' ? <Select key={key} label={labelText} value={draft[key]} options={options.length ? options : ['']} onChange={(v) => set(key, v)} /> : type === 'textarea' ? <TextArea key={key} label={labelText} value={draft[key]} onChange={(v) => set(key, v)} /> : <Field key={key} label={labelText} type={type} value={draft[key]} onChange={(v) => set(key, v)} />)}</div><div className="modal-actions"><button className="btn ghost" onClick={onClose}>Batal</button><button className="btn primary" onClick={() => onSave(collection, draft)}>{config.save}</button></div></Modal>;
}
function WorkOrderForm({ form, assets, vendors, onClose, onSave }) {
  const [draft, setDraft] = useState(normalizeWorkOrder(form, assets.find((a) => a.id === form.asetId)));
  const set = (key, value) => setDraft((prev) => ({ ...prev, [key]: value }));
  const selectedAsset = assets.find((asset) => asset.id === draft.asetId);
  const actualTotal = numeric(draft.biayaJasa) + numeric(draft.biayaSparepart) + numeric(draft.biayaLain) + numeric(draft.pajak);
  function fileToDataUrl(file, nameKey, dataKey, statusKey) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setDraft((prev) => ({ ...prev, [nameKey]: file.name, [dataKey]: reader.result, ...(statusKey ? { [statusKey]: file.name } : {}) }));
    reader.readAsDataURL(file);
  }
  return <Modal title={form?.id ? `Edit Work Order ${form.id}` : 'Buat Work Order'} onClose={onClose}>
    <div className="form-section-title"><b>Identitas Work Order</b><span>ID, aset, lokasi, dan penanggung jawab.</span></div>
    <div className="form-grid">
      <Field label="ID WO Otomatis" value={draft.id} onChange={() => null} readOnly />
      <Select label="Aset" value={draft.asetId} options={assets.map((asset) => asset.id)} onChange={(v) => { const a = assets.find((asset) => asset.id === v); setDraft((prev) => ({ ...prev, asetId: v, namaAset: a?.nama || '', nopol: a?.nomorPolisi || '-', kilometer: a?.km || prev.kilometer, lokasiPekerjaan: a?.lokasi || prev.lokasiPekerjaan })); }} />
      <Field label="Nama Aset" value={selectedAsset?.nama || draft.namaAset} onChange={(v) => set('namaAset', v)} readOnly={!!selectedAsset} />
      <Field label="Nomor Polisi / Serial" value={selectedAsset?.nomorPolisi || draft.nopol || '-'} onChange={(v) => set('nopol', v)} readOnly={!!selectedAsset} />
      <Field label="KM / Odometer Saat WO" type="number" value={draft.kilometer} onChange={(v) => set('kilometer', v)} />
      <Field label="Lokasi Pekerjaan" value={draft.lokasiPekerjaan} onChange={(v) => set('lokasiPekerjaan', v)} />
      <Select label="Jenis WO" value={draft.jenis} options={workOrderType} onChange={(v) => set('jenis', v)} />
      <Select label="Prioritas" value={draft.prioritas} options={priorityOptions} onChange={(v) => set('prioritas', v)} />
      <Select label="Status" value={draft.status} options={workOrderStatus} onChange={(v) => set('status', v)} />
      <Field label="PIC Internal" value={draft.pic} onChange={(v) => set('pic', v)} />
      <Field label="Teknisi / Mekanik" value={draft.teknisi} onChange={(v) => set('teknisi', v)} />
      <Select label="Vendor / Bengkel" value={draft.vendorId || ''} options={['', ...vendors.map((v) => v.id)]} onChange={(v) => set('vendorId', v)} />
    </div>
    <div className="form-section-title"><b>Jadwal, Biaya & Administrasi</b><span>Estimasi, realisasi biaya, approval, PO, invoice, dan warranty.</span></div>
    <div className="form-grid">
      <Field label="Tanggal Mulai" type="date" value={draft.tanggalMulai} onChange={(v) => set('tanggalMulai', v)} />
      <Field label="Jatuh Tempo" type="date" value={draft.jatuhTempo} onChange={(v) => set('jatuhTempo', v)} />
      <Field label="Tanggal Selesai" type="date" value={draft.tanggalSelesai} onChange={(v) => set('tanggalSelesai', v)} />
      <Field label="Downtime (Jam)" type="number" value={draft.downtimeJam} onChange={(v) => set('downtimeJam', v)} />
      <Field label="Estimasi Biaya" type="number" value={draft.estimasiBiaya} onChange={(v) => set('estimasiBiaya', v)} />
      <Field label="Biaya Jasa Aktual" type="number" value={draft.biayaJasa} onChange={(v) => set('biayaJasa', v)} />
      <Field label="Biaya Sparepart Aktual" type="number" value={draft.biayaSparepart} onChange={(v) => set('biayaSparepart', v)} />
      <Field label="Biaya Lain-lain" type="number" value={draft.biayaLain} onChange={(v) => set('biayaLain', v)} />
      <Field label="Pajak / PPN" type="number" value={draft.pajak} onChange={(v) => set('pajak', v)} />
      <Field label="Nominal Biaya Aktual Total" type="number" value={actualTotal || draft.biaya} onChange={() => null} readOnly />
      <Field label="Nomor PO" value={draft.noPO} onChange={(v) => set('noPO', v)} />
      <Field label="Nomor Invoice Vendor" value={draft.noInvoice} onChange={(v) => set('noInvoice', v)} />
      <Select label="Warranty / Claim" value={draft.warrantyStatus} options={['Tidak','Ya - Warranty','Ya - Asuransi','Dalam Proses Claim','Claim Disetujui','Claim Ditolak']} onChange={(v) => set('warrantyStatus', v)} />
      <Select label="Status Approval" value={draft.approvalStatus} options={['Belum Diajukan','Menunggu Approval','Disetujui','Ditolak','Tidak Perlu Approval']} onChange={(v) => set('approvalStatus', v)} />
    </div>
    <div className="form-section-title"><b>Detail Pekerjaan, Evidence & QC</b><span>Keluhan, diagnosa, tindakan, sparepart, foto, lampiran, dan hasil QC.</span></div>
    <div className="form-grid">
      <TextArea label="Keluhan / Trigger" value={draft.keluhan} onChange={(v) => set('keluhan', v)} />
      <TextArea label="Diagnosa" value={draft.diagnosa} onChange={(v) => set('diagnosa', v)} />
      <TextArea label="Tindakan / Pekerjaan Dilakukan" value={draft.tindakan} onChange={(v) => set('tindakan', v)} />
      <TextArea label="Sparepart Dipakai" value={draft.sparepart} onChange={(v) => set('sparepart', v)} />
      <TextArea label="Checklist Pekerjaan" value={draft.checklist} onChange={(v) => set('checklist', v)} />
      <TextArea label="Hasil QC / Test Drive" value={draft.hasilQC} onChange={(v) => set('hasilQC', v)} />
      <TextArea label="Catatan Tambahan" value={draft.catatan} onChange={(v) => set('catatan', v)} />
      <label className="field full"><span>Foto Before</span><input type="file" accept="image/*" onChange={(e) => fileToDataUrl(e.target.files?.[0], 'fotoBeforeName', 'fotoBeforeData', 'fotoBefore')} /><small className="field-help">{draft.fotoBeforeName || draft.fotoBefore || 'Belum ada file'}</small></label>
      <label className="field full"><span>Foto After</span><input type="file" accept="image/*" onChange={(e) => fileToDataUrl(e.target.files?.[0], 'fotoAfterName', 'fotoAfterData', 'fotoAfter')} /><small className="field-help">{draft.fotoAfterName || draft.fotoAfter || 'Belum ada file'}</small></label>
      <label className="field full"><span>Lampiran WO / Invoice / Laporan Vendor</span><input type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={(e) => fileToDataUrl(e.target.files?.[0], 'lampiranName', 'lampiranData')} /><small className="field-help">{draft.lampiranName || 'Belum ada lampiran'}</small></label>
    </div>
    <div className="modal-actions"><button className="btn ghost" onClick={onClose}>Batal</button><button className="btn primary" onClick={() => onSave({ ...draft, biaya: actualTotal || numeric(draft.biaya) })}>Simpan Work Order</button></div>
  </Modal>;
}
function DataTable({ rows, columns }) { return <div className="table-wrap"><table><thead><tr>{columns.map(([, label]) => <th key={label}>{label}</th>)}</tr></thead><tbody>{rows.map((row) => <tr key={row.id}>{columns.map(([key, , format]) => <td key={key}>{format ? format(row[key], row) : row[key]}</td>)}</tr>)}</tbody></table></div>; }
function Modal({ title, children, onClose }) { return <div className="modal-backdrop" onClick={onClose}><div className="modal" onClick={(e) => e.stopPropagation()}><div className="modal-head"><h2>{title}</h2><button onClick={onClose}>×</button></div>{children}</div></div>; }
function Field({ label, value, onChange, type = 'text', placeholder = '', readOnly = false }) { return <label className="field"><span>{label}</span><input type={type} value={value ?? ''} placeholder={placeholder} readOnly={readOnly} className={readOnly ? 'readonly-input' : ''} onChange={(e) => !readOnly && onChange(e.target.value)} /></label>; }
function PasswordField({ label, value, onChange, placeholder = '' }) { const [visible, setVisible] = useState(false); return <label className="field password-field"><span>{label}</span><div className="password-control"><input type={visible ? 'text' : 'password'} value={value ?? ''} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} /><button type="button" className="password-toggle" onClick={() => setVisible((current) => !current)}>{visible ? 'Sembunyikan' : 'Lihat'}</button></div></label>; }
function Select({ label, value, options, onChange }) { return <label className="field"><span>{label}</span><select value={value ?? ''} onChange={(e) => onChange(e.target.value)}>{options.map((item) => <option key={item} value={item}>{item || '-'}</option>)}</select></label>; }
function TextArea({ label, value, onChange }) { return <label className="field full"><span>{label}</span><textarea value={value ?? ''} onChange={(e) => onChange(e.target.value)} rows={4} /></label>; }

export default App;
