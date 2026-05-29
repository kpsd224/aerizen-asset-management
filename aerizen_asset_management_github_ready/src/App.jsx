import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { motion, AnimatePresence } from "framer-motion";
import {
  Activity,
  AlertTriangle,
  Archive,
  BarChart3,
  Bell,
  Building2,
  CalendarClock,
  Car,
  CheckCircle2,
  ChevronRight,
  ClipboardCheck,
  Database,
  Edit3,
  FileText,
  Filter,
  Gauge,
  GitBranch,
  Laptop,
  Layers3,
  MapPin,
  Plus,
  QrCode,
  RefreshCw,
  Save,
  Search,
  ShieldCheck,
  Sparkles,
  Trash2,
  Truck,
  Upload,
  Users,
  Workflow,
  Wrench,
  X,
} from "lucide-react";

const dataAsetAwal = [
  {
    id: "AST-KND-2026-0184",
    nama: "Toyota Innova Zenix",
    kategori: "Kendaraan",
    tipe: "Aset Rental",
    gambar: "https://images.unsplash.com/photo-1542362567-b07e54358753?q=80&w=1200&auto=format&fit=crop",
    lokasi: "Pool Jakarta A",
    cabang: "Jakarta",
    pemegang: "PT Nusantara Retail",
    status: "Disewakan",
    kondisi: "Sangat Baik",
    risiko: 18,
    utilisasi: 91,
    roi: 32,
    pendapatan: "Rp 18.400.000",
    biaya: "Rp 4.200.000",
    nilaiBuku: "Rp 287.500.000",
    aksiBerikutnya: "Perpanjangan kontrak dalam 6 hari",
    dokumen: "STNK berlaku 84 hari lagi",
    telemetri: "GPS Online · 42.180 km",
    alur: "Disewakan → Monitoring → Perpanjangan / Pengembalian",
    riwayat: ["Terdaftar", "Dipasang QR", "Disewakan", "GPS Sinkron", "Invoice Terbit"],
  },
  {
    id: "AST-IT-2026-0441",
    nama: "MacBook Pro M3 14”",
    kategori: "Perangkat IT",
    tipe: "Penugasan Karyawan",
    gambar: "https://images.unsplash.com/photo-1517336714731-489689fd1ca8?q=80&w=1200&auto=format&fit=crop",
    lokasi: "Cabang Bandung",
    cabang: "Bandung",
    pemegang: "Ayu Pramesti · Tim Produk",
    status: "Dipakai Internal",
    kondisi: "Baik",
    risiko: 26,
    utilisasi: 76,
    roi: 19,
    pendapatan: "Biaya internal",
    biaya: "Rp 850.000",
    nilaiBuku: "Rp 31.900.000",
    aksiBerikutnya: "Cek kesehatan baterai",
    dokumen: "Garansi berlaku 218 hari lagi",
    telemetri: "MDM Online · Enkripsi Aktif",
    alur: "Ditugaskan → Pemeriksaan Berkala → Kembali / Mutasi",
    riwayat: ["Terdaftar", "Setup Keamanan", "Ditugaskan", "Cek MDM", "Review Garansi"],
  },
  {
    id: "AST-KND-2025-0098",
    nama: "Mitsubishi L300 Box",
    kategori: "Kendaraan",
    tipe: "Aset Operasional",
    gambar: "https://images.unsplash.com/photo-1601584115197-04ecc0da31d7?q=80&w=1200&auto=format&fit=crop",
    lokasi: "Yard Surabaya",
    cabang: "Surabaya",
    pemegang: "Logistik Internal",
    status: "Maintenance",
    kondisi: "Perlu Servis",
    risiko: 72,
    utilisasi: 58,
    roi: 11,
    pendapatan: "Rp 9.700.000",
    biaya: "Rp 7.100.000",
    nilaiBuku: "Rp 132.000.000",
    aksiBerikutnya: "Inspeksi rem menunggu QC",
    dokumen: "KIR habis dalam 19 hari",
    telemetri: "GPS Online · 83.900 km",
    alur: "Work Order → Vendor Repair → QC → Tersedia",
    riwayat: ["Terdaftar", "Dipakai Logistik", "Insiden", "Work Order", "Menunggu QC"],
  },
  {
    id: "AST-OFC-2026-0067",
    nama: "Canon ImageRunner DX",
    kategori: "Peralatan Kantor",
    tipe: "Aset Cabang",
    gambar: "https://images.unsplash.com/photo-1612815154858-60aa4c59eaa6?q=80&w=1200&auto=format&fit=crop",
    lokasi: "Cabang Medan",
    cabang: "Medan",
    pemegang: "Departemen Finance",
    status: "Tersedia",
    kondisi: "Baik",
    risiko: 12,
    utilisasi: 44,
    roi: 8,
    pendapatan: "Biaya internal",
    biaya: "Rp 380.000",
    nilaiBuku: "Rp 18.400.000",
    aksiBerikutnya: "Siap untuk booking internal",
    dokumen: "Kontrak servis aktif",
    telemetri: "Meter 128.440 halaman",
    alur: "Tersedia → Dipesan → Ditugaskan → Audit",
    riwayat: ["Terdaftar", "Kontrak Servis", "Audit Lulus", "Tersedia"],
  },
];

const dataWorkOrderAwal = [
  {
    id: "WO-2026-0912",
    asetId: "AST-KND-2025-0098",
    namaAset: "Mitsubishi L300 Box",
    jenis: "Corrective Maintenance",
    prioritas: "Tinggi",
    status: "Menunggu QC",
    pic: "Vendor Surabaya",
    tanggalMulai: "2026-05-28",
    jatuhTempo: "2026-05-29",
    estimasiBiaya: "Rp 3.500.000",
    keluhan: "Rem terasa kurang pakem dan muncul suara saat pengereman.",
    diagnosis: "Perlu pemeriksaan kampas rem, minyak rem, dan kaliper.",
    checklist: "Cek rem depan, cek rem belakang, test drive, foto evidence, QC akhir",
    catatan: "Prioritas tinggi karena kendaraan dipakai untuk operasional logistik.",
  },
];

const opsiStatus = ["Tersedia", "Disewakan", "Dipakai Internal", "Maintenance", "Perbaikan", "Overdue", "Nonaktif"];
const opsiKategori = ["Kendaraan", "Perangkat IT", "Peralatan Kantor", "Mesin", "Properti", "Lainnya"];
const opsiJenisWorkOrder = ["Preventive Maintenance", "Corrective Maintenance", "Emergency Repair", "Vendor Repair", "Warranty Claim", "Inspection", "Reconditioning"];
const opsiPrioritasWorkOrder = ["Rendah", "Sedang", "Tinggi", "Kritis"];
const opsiStatusWorkOrder = ["Draft", "Menunggu Approval", "Dijadwalkan", "Dikerjakan", "Menunggu Sparepart", "Menunggu Vendor", "Menunggu QC", "Selesai", "Ditutup"];
const alurKerja = ["Permintaan", "Cek Risiko", "Approval", "Reservasi", "Inspeksi", "Serah Terima", "Monitoring", "Kembali / Perpanjang"];

const modul = [
  { icon: Archive, judul: "Aset 360", deskripsi: "Identitas, dokumen, foto, riwayat, biaya, risiko, dan timeline aset." },
  { icon: FileText, judul: "Kontrak Rental", deskripsi: "Quotation, kontrak, deposit, billing, renewal, denda, dan klaim." },
  { icon: Wrench, judul: "Maintenance", deskripsi: "Preventive, corrective, vendor repair, work order, QC, dan sparepart." },
  { icon: ClipboardCheck, judul: "Audit & Stock Opname", deskripsi: "Scan QR/RFID, selisih fisik, bukti foto, dan approval adjustment." },
  { icon: ShieldCheck, judul: "Kepatuhan", deskripsi: "STNK, KIR, garansi, asuransi, kontrak servis, dan reminder expired." },
  { icon: Sparkles, judul: "Rekomendasi AI", deskripsi: "Prediksi kerusakan, rekomendasi aset, dan deteksi anomali biaya." },
];

const tugasOperasional = [
  { id: "WO-2026-0912", jenis: "Maintenance", judul: "Inspeksi rem", aset: "AST-KND-2025-0098", pic: "Vendor Surabaya", sla: "Sisa 4 jam", level: "Tinggi" },
  { id: "DLV-2026-0043", jenis: "Delivery", judul: "Serah terima customer", aset: "AST-KND-2026-0184", pic: "Driver Jakarta", sla: "Hari ini", level: "Sedang" },
  { id: "APR-2026-0188", jenis: "Approval", judul: "Approval biaya repair", aset: "AST-KND-2025-0098", pic: "Finance Manager", sla: "Sisa 2 jam", level: "Kritis" },
];

const apiEndpoints = [
  "GET /api/aset?status=&cabang=&kategori=",
  "POST /api/aset/daftar",
  "GET /api/aset/:id/360",
  "POST /api/kontrak-rental",
  "POST /api/work-order",
  "POST /api/stock-opname/scan",
  "POST /api/approval/:id/setujui",
  "GET /api/laporan/utilisasi",
];

const tabelDatabase = ["perusahaan", "cabang", "pengguna", "role", "aset", "dokumen_aset", "foto_aset", "riwayat_aset", "kendaraan", "perangkat_it", "kontrak_rental", "pengiriman", "pengembalian", "work_order", "ticket_repair", "approval", "keuangan_aset", "stock_opname", "audit_log"];

function buatIdAset(kategori) {
  const kode = kategori === "Kendaraan" ? "KND" : kategori === "Perangkat IT" ? "IT" : kategori === "Peralatan Kantor" ? "OFC" : "AST";
  return `AST-${kode}-2026-${Math.floor(1000 + Math.random() * 8999)}`;
}

function buatIdWorkOrder() {
  return `WO-2026-${Math.floor(1000 + Math.random() * 8999)}`;
}

function gayaStatus(status) {
  const styles = {
    Disewakan: "bg-blue-50 text-blue-700 border-blue-100",
    "Dipakai Internal": "bg-violet-50 text-violet-700 border-violet-100",
    Maintenance: "bg-amber-50 text-amber-700 border-amber-100",
    Perbaikan: "bg-orange-50 text-orange-700 border-orange-100",
    Overdue: "bg-red-50 text-red-700 border-red-100",
    Tersedia: "bg-emerald-50 text-emerald-700 border-emerald-100",
    Nonaktif: "bg-slate-100 text-slate-600 border-slate-200",
  };
  return styles[status] || "bg-slate-50 text-slate-700 border-slate-100";
}

function gayaRisiko(risiko) {
  if (Number(risiko) >= 70) return "text-red-600 bg-red-50 border-red-100";
  if (Number(risiko) >= 35) return "text-amber-600 bg-amber-50 border-amber-100";
  return "text-emerald-600 bg-emerald-50 border-emerald-100";
}

function hitungRataUtilisasi(daftarAset) {
  if (!Array.isArray(daftarAset) || daftarAset.length === 0) return 0;
  return Math.round(daftarAset.reduce((sum, aset) => sum + Number(aset.utilisasi || 0), 0) / daftarAset.length);
}

function hitungAsetBerisiko(daftarAset, ambang = 60) {
  if (!Array.isArray(daftarAset)) return 0;
  return daftarAset.filter((aset) => Number(aset.risiko || 0) >= ambang).length;
}

function normalisasiAset(form) {
  return {
    ...form,
    id: form.id || buatIdAset(form.kategori),
    risiko: Number(form.risiko || 0),
    utilisasi: Number(form.utilisasi || 0),
    roi: Number(form.roi || 0),
    riwayat: Array.isArray(form.riwayat) && form.riwayat.length > 0 ? form.riwayat : ["Terdaftar"],
  };
}

function normalisasiWorkOrder(form) {
  return {
    ...form,
    id: form.id || buatIdWorkOrder(),
    status: form.status || "Draft",
    prioritas: form.prioritas || "Sedang",
    estimasiBiaya: form.estimasiBiaya || "Rp 0",
    checklist: form.checklist || "Inspeksi awal, foto evidence, update progress, QC akhir",
  };
}

function normalisasiHeaderImport(header) {
  return String(header || "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[._-]+/g, " ");
}

function ambilNilaiKolom(row, daftarNamaKolom) {
  const pasangan = Object.entries(row || {}).reduce((hasil, [key, value]) => {
    hasil[normalisasiHeaderImport(key)] = value;
    return hasil;
  }, {});
  for (const nama of daftarNamaKolom) {
    const nilai = pasangan[normalisasiHeaderImport(nama)];
    if (nilai !== undefined && nilai !== null && String(nilai).trim() !== "") return nilai;
  }
  return "";
}

function formatTanggalExcel(nilai) {
  if (!nilai) return "";
  if (nilai instanceof Date && !Number.isNaN(nilai.getTime())) return nilai.toISOString().slice(0, 10);
  if (typeof nilai === "number") {
    const tanggal = XLSX.SSF.parse_date_code(nilai);
    if (tanggal) {
      const bulan = String(tanggal.m).padStart(2, "0");
      const hari = String(tanggal.d).padStart(2, "0");
      return `${tanggal.y}-${bulan}-${hari}`;
    }
  }
  const teks = String(nilai).trim();
  const parsed = new Date(teks);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  return teks;
}

function formatBpkb(nilai) {
  const teks = String(nilai || "").trim().toLowerCase();
  if (["ya", "y", "yes", "true", "ada", "1", "memiliki"].includes(teks)) return "Ya";
  if (["tidak", "no", "false", "0", "belum", "tidak ada"].includes(teks)) return "Tidak";
  return nilai ? String(nilai) : "Belum diisi";
}

function buatAsetDariImportKendaraan(row, index) {
  const no = ambilNilaiKolom(row, ["no", "nomor", "no."]);
  const namaKendaraan = String(ambilNilaiKolom(row, ["nama kendaraan", "kendaraan", "nama aset", "nama"]) || `Kendaraan Sewa ${index + 1}`);
  const warna = String(ambilNilaiKolom(row, ["warna", "color"]) || "-");
  const tahunKendaraan = String(ambilNilaiKolom(row, ["tahun kendaraan", "tahun", "year"]) || "-");
  const nomorRangka = String(ambilNilaiKolom(row, ["nomor rangka", "no rangka", "norangka", "vin", "chassis number"]) || "-");
  const nomorMesin = String(ambilNilaiKolom(row, ["nomor mesin", "no mesin", "nomesin", "engine number"]) || "-");
  const perusahaanUser = String(ambilNilaiKolom(row, ["perusahaan user", "perusahaan", "customer", "nama perusahaan"]) || "Belum diisi");
  const alamatUser = String(ambilNilaiKolom(row, ["alamat user", "alamat", "alamat perusahaan"]) || "-");
  const wilayah = String(ambilNilaiKolom(row, ["wilayah", "area", "region", "cabang"]) || "-");
  const tanggalSTNK = formatTanggalExcel(ambilNilaiKolom(row, ["tanggal STNK", "tanggal stnk", "stnk", "masa berlaku stnk", "expired stnk"]));
  const memilikiBPKB = formatBpkb(ambilNilaiKolom(row, ["memiliki BPKB", "bpkb", "ada bpkb"]));
  const kodeUnik = nomorRangka !== "-" ? nomorRangka.slice(-6).toUpperCase() : String(no || index + 1).padStart(4, "0");

  return normalisasiAset({
    id: `AST-KND-SEWA-${kodeUnik}`,
    nama: namaKendaraan,
    kategori: "Kendaraan",
    tipe: "Aset Rental",
    gambar: "https://images.unsplash.com/photo-1542362567-b07e54358753?q=80&w=1200&auto=format&fit=crop",
    lokasi: wilayah,
    cabang: wilayah,
    pemegang: perusahaanUser,
    status: "Disewakan",
    kondisi: "Baik",
    risiko: tanggalSTNK ? 18 : 45,
    utilisasi: 80,
    roi: 20,
    pendapatan: "Rp 0",
    biaya: "Rp 0",
    nilaiBuku: "Rp 0",
    aksiBerikutnya: tanggalSTNK ? `Monitoring STNK ${tanggalSTNK}` : "Lengkapi tanggal STNK",
    dokumen: `STNK: ${tanggalSTNK || "Belum diisi"} · BPKB: ${memilikiBPKB}`,
    telemetri: "GPS belum tersambung",
    alur: "Import Excel → Validasi Dokumen → Disewakan → Monitoring → Pengembalian",
    riwayat: ["Import Excel", "Validasi Kendaraan Sewa"],
    noImport: no,
    warna,
    tahunKendaraan,
    nomorRangka,
    nomorMesin,
    perusahaanUser,
    alamatUser,
    wilayah,
    tanggalSTNK,
    memilikiBPKB,
  });
}

async function bacaImportKendaraanSewa(file) {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "", raw: false });
  return rows
    .filter((row) => Object.values(row).some((value) => String(value || "").trim() !== ""))
    .map((row, index) => buatAsetDariImportKendaraan(row, index));
}

function jalankanTesMandiri() {
  console.assert(gayaStatus("Tersedia").includes("emerald"), "Status Tersedia harus memakai gaya hijau/emerald.");
  console.assert(gayaStatus("Status Tidak Ada").includes("slate"), "Status tidak dikenal harus memakai gaya default slate.");
  console.assert(gayaRisiko(72).includes("red"), "Risiko tinggi harus memakai gaya merah.");
  console.assert(gayaRisiko(40).includes("amber"), "Risiko sedang harus memakai gaya amber.");
  console.assert(gayaRisiko(10).includes("emerald"), "Risiko rendah harus memakai gaya emerald.");
  console.assert(hitungRataUtilisasi([{ utilisasi: 10 }, { utilisasi: 30 }]) === 20, "Rata-rata utilisasi harus benar.");
  console.assert(hitungRataUtilisasi([]) === 0, "Rata-rata utilisasi daftar kosong harus 0.");
  console.assert(hitungAsetBerisiko([{ risiko: 59 }, { risiko: 60 }, { risiko: 90 }]) === 2, "Hitung aset berisiko harus memakai ambang >= 60.");
  console.assert(normalisasiAset({ kategori: "Kendaraan", risiko: "7", utilisasi: "8", roi: "9" }).id.startsWith("AST-KND-2026-"), "ID otomatis kendaraan harus memakai prefix AST-KND-2026.");
  console.assert(buatIdWorkOrder().startsWith("WO-2026-"), "ID Work Order otomatis harus memakai prefix WO-2026.");
  console.assert(normalisasiWorkOrder({ asetId: "AST-TEST" }).status === "Draft", "Work Order baru tanpa status harus menjadi Draft.");
  console.assert(normalisasiWorkOrder({ asetId: "AST-TEST", prioritas: "Kritis" }).prioritas === "Kritis", "Prioritas Work Order yang diisi harus dipertahankan.");
}

jalankanTesMandiri();

function KartuMetrik({ icon: Icon, label, nilai, detail }) {
  return (
    <motion.div initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }} className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex items-center justify-between">
        <div className="rounded-2xl bg-slate-100 p-3"><Icon className="h-5 w-5 text-slate-700" /></div>
        <span className="text-xs font-medium text-emerald-600">Live</span>
      </div>
      <div className="mt-5 text-3xl font-semibold tracking-tight text-slate-950">{nilai}</div>
      <div className="mt-1 text-sm font-medium text-slate-700">{label}</div>
      <div className="mt-2 text-xs text-slate-500">{detail}</div>
    </motion.div>
  );
}

function Input({ label, value, onChange, type = "text", placeholder = "" }) {
  return (
    <label className="block">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      <input type={type} value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} className="mt-1 w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none focus:border-slate-400" />
    </label>
  );
}

function Select({ label, value, options, onChange }) {
  return (
    <label className="block">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      <select value={value} onChange={(e) => onChange(e.target.value)} className="mt-1 w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none focus:border-slate-400">
        {options.map((item) => <option key={item}>{item}</option>)}
      </select>
    </label>
  );
}

function TextArea({ label, value, onChange, placeholder = "" }) {
  return (
    <label className="block">
      <span className="text-xs font-semibold text-slate-500">{label}</span>
      <textarea value={value} placeholder={placeholder} onChange={(e) => onChange(e.target.value)} rows={4} className="mt-1 w-full resize-none rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm outline-none focus:border-slate-400" />
    </label>
  );
}

function FormAset({ aset, onClose, onSave }) {
  const [form, setForm] = useState(aset || {
    id: "",
    nama: "",
    kategori: "Kendaraan",
    tipe: "Aset Rental",
    gambar: "https://images.unsplash.com/photo-1542362567-b07e54358753?q=80&w=1200&auto=format&fit=crop",
    lokasi: "",
    cabang: "",
    pemegang: "",
    status: "Tersedia",
    kondisi: "Baik",
    risiko: 10,
    utilisasi: 0,
    roi: 0,
    pendapatan: "Rp 0",
    biaya: "Rp 0",
    nilaiBuku: "Rp 0",
    aksiBerikutnya: "Belum ada aksi",
    dokumen: "Dokumen belum lengkap",
    telemetri: "Belum tersambung",
    alur: "Tersedia → Dipesan → Digunakan → Kembali",
    riwayat: ["Terdaftar"],
  });

  const ubah = (key, value) => setForm((prev) => ({ ...prev, [key]: value }));
  const simpan = () => onSave(normalisasiAset(form));

  return (
    <AnimatePresence>
      <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-sm">
        <motion.div initial={{ x: 560 }} animate={{ x: 0 }} exit={{ x: 560 }} transition={{ type: "spring", damping: 30, stiffness: 260 }} className="absolute right-0 top-0 h-full w-full max-w-2xl overflow-y-auto bg-white p-6 shadow-2xl">
          <div className="flex items-start justify-between gap-4">
            <div>
              <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Form Aset</div>
              <h2 className="mt-1 text-2xl font-bold tracking-tight text-slate-950">{aset ? "Edit Data Aset" : "Tambah Aset Baru"}</h2>
              <p className="mt-1 text-sm text-slate-500">Semua data di daftar bisa diedit, ditambah, dan dihapus langsung dari tampilan.</p>
            </div>
            <button onClick={onClose} className="rounded-2xl border border-slate-200 p-3 hover:bg-slate-50"><X className="h-5 w-5" /></button>
          </div>
          <div className="mt-6 grid gap-4 sm:grid-cols-2">
            <Input label="Kode Aset" value={form.id} onChange={(v) => ubah("id", v)} placeholder="Otomatis jika kosong" />
            <Input label="Nama Aset" value={form.nama} onChange={(v) => ubah("nama", v)} />
            <Select label="Kategori" value={form.kategori} options={opsiKategori} onChange={(v) => ubah("kategori", v)} />
            <Select label="Status" value={form.status} options={opsiStatus} onChange={(v) => ubah("status", v)} />
            <Input label="Tipe" value={form.tipe} onChange={(v) => ubah("tipe", v)} />
            <Input label="Kondisi" value={form.kondisi} onChange={(v) => ubah("kondisi", v)} />
            <Input label="Cabang" value={form.cabang} onChange={(v) => ubah("cabang", v)} />
            <Input label="Lokasi" value={form.lokasi} onChange={(v) => ubah("lokasi", v)} />
            <Input label="Pemegang / Customer" value={form.pemegang} onChange={(v) => ubah("pemegang", v)} />
            <Input label="URL Gambar" value={form.gambar} onChange={(v) => ubah("gambar", v)} />
            <Input label="Risiko (%)" type="number" value={form.risiko} onChange={(v) => ubah("risiko", v)} />
            <Input label="Utilisasi (%)" type="number" value={form.utilisasi} onChange={(v) => ubah("utilisasi", v)} />
            <Input label="ROI (%)" type="number" value={form.roi} onChange={(v) => ubah("roi", v)} />
            <Input label="Nilai Buku" value={form.nilaiBuku} onChange={(v) => ubah("nilaiBuku", v)} />
            <Input label="Pendapatan Bulanan" value={form.pendapatan} onChange={(v) => ubah("pendapatan", v)} />
            <Input label="Biaya Bulanan" value={form.biaya} onChange={(v) => ubah("biaya", v)} />
            <Input label="Dokumen" value={form.dokumen} onChange={(v) => ubah("dokumen", v)} />
            <Input label="Telemetri" value={form.telemetri} onChange={(v) => ubah("telemetri", v)} />
            <div className="sm:col-span-2"><Input label="Aksi Berikutnya" value={form.aksiBerikutnya} onChange={(v) => ubah("aksiBerikutnya", v)} /></div>
            <div className="sm:col-span-2"><Input label="Alur Aset" value={form.alur} onChange={(v) => ubah("alur", v)} /></div>
          </div>
          <div className="mt-6 flex gap-3">
            <button onClick={simpan} className="flex flex-1 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white hover:bg-slate-800"><Save className="h-4 w-4" /> Simpan Data</button>
            <button onClick={onClose} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold hover:bg-slate-50">Batal</button>
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
}

function FormWorkOrder({ aset, onClose, onSave }) {
  const today = new Date().toISOString().slice(0, 10);
  const [form, setForm] = useState({
    id: "",
    asetId: aset?.id || "",
    namaAset: aset?.nama || "",
    jenis: aset?.kategori === "Kendaraan" ? "Corrective Maintenance" : "Inspection",
    prioritas: Number(aset?.risiko || 0) >= 70 ? "Tinggi" : "Sedang",
    status: "Draft",
    pic: aset?.kategori === "Kendaraan" ? "Teknisi Fleet" : "Teknisi IT / GA",
    tanggalMulai: today,
    jatuhTempo: today,
    estimasiBiaya: "Rp 0",
    keluhan: aset?.aksiBerikutnya || "",
    diagnosis: "",
    checklist: aset?.kategori === "Kendaraan" ? "Cek oli, cek rem, cek ban, cek lampu, foto evidence, QC akhir" : "Cek kondisi fisik, cek fungsi, foto evidence, update status, QC akhir",
    catatan: "",
  });

  const ubah = (key, value) => setForm((prev) => ({ ...prev, [key]: value }));
  const simpan = () => onSave(normalisasiWorkOrder(form));

  return (
    <AnimatePresence>
      <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-sm">
        <motion.div initial={{ x: 560 }} animate={{ x: 0 }} exit={{ x: 560 }} transition={{ type: "spring", damping: 30, stiffness: 260 }} className="absolute right-0 top-0 h-full w-full max-w-2xl overflow-y-auto bg-white p-6 shadow-2xl">
          <div className="flex items-start justify-between gap-4">
            <div>
              <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Form Work Order</div>
              <h2 className="mt-1 text-2xl font-bold tracking-tight text-slate-950">Buat Work Order</h2>
              <p className="mt-1 text-sm text-slate-500">Isi pekerjaan maintenance/perbaikan untuk aset: <b>{aset?.nama}</b>.</p>
            </div>
            <button onClick={onClose} className="rounded-2xl border border-slate-200 p-3 hover:bg-slate-50"><X className="h-5 w-5" /></button>
          </div>
          <div className="mt-6 grid gap-4 sm:grid-cols-2">
            <Input label="Nomor Work Order" value={form.id} onChange={(v) => ubah("id", v)} placeholder="Otomatis jika kosong" />
            <Input label="Kode Aset" value={form.asetId} onChange={(v) => ubah("asetId", v)} />
            <Input label="Nama Aset" value={form.namaAset} onChange={(v) => ubah("namaAset", v)} />
            <Select label="Jenis Pekerjaan" value={form.jenis} options={opsiJenisWorkOrder} onChange={(v) => ubah("jenis", v)} />
            <Select label="Prioritas" value={form.prioritas} options={opsiPrioritasWorkOrder} onChange={(v) => ubah("prioritas", v)} />
            <Select label="Status" value={form.status} options={opsiStatusWorkOrder} onChange={(v) => ubah("status", v)} />
            <Input label="PIC / Vendor" value={form.pic} onChange={(v) => ubah("pic", v)} />
            <Input label="Estimasi Biaya" value={form.estimasiBiaya} onChange={(v) => ubah("estimasiBiaya", v)} />
            <Input label="Tanggal Mulai" type="date" value={form.tanggalMulai} onChange={(v) => ubah("tanggalMulai", v)} />
            <Input label="Jatuh Tempo" type="date" value={form.jatuhTempo} onChange={(v) => ubah("jatuhTempo", v)} />
            <div className="sm:col-span-2"><TextArea label="Keluhan / Permintaan" value={form.keluhan} onChange={(v) => ubah("keluhan", v)} placeholder="Contoh: Rem kurang pakem, laptop sering panas, printer error..." /></div>
            <div className="sm:col-span-2"><TextArea label="Diagnosis Awal" value={form.diagnosis} onChange={(v) => ubah("diagnosis", v)} placeholder="Isi hasil pengecekan awal atau dugaan penyebab masalah." /></div>
            <div className="sm:col-span-2"><TextArea label="Checklist Pekerjaan" value={form.checklist} onChange={(v) => ubah("checklist", v)} placeholder="Pisahkan checklist dengan koma." /></div>
            <div className="sm:col-span-2"><TextArea label="Catatan Tambahan" value={form.catatan} onChange={(v) => ubah("catatan", v)} /></div>
          </div>
          <div className="mt-6 flex gap-3">
            <button onClick={simpan} className="flex flex-1 items-center justify-center gap-2 rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white hover:bg-slate-800"><Save className="h-4 w-4" /> Simpan Work Order</button>
            <button onClick={onClose} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold hover:bg-slate-50">Batal</button>
          </div>
        </motion.div>
      </motion.div>
    </AnimatePresence>
  );
}

function KartuAset({ aset, onOpen, onEdit, onDelete }) {
  return (
    <motion.article layout initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} className="group overflow-hidden rounded-[2rem] border border-slate-200 bg-white shadow-sm transition hover:-translate-y-1 hover:shadow-xl">
      <div className="relative h-48 overflow-hidden bg-slate-100">
        <img src={aset.gambar} alt={aset.nama} className="h-full w-full object-cover transition duration-500 group-hover:scale-105" />
        <div className="absolute inset-0 bg-gradient-to-t from-slate-950/70 via-slate-950/10 to-transparent" />
        <div className="absolute left-4 top-4 flex gap-2">
          <span className={`rounded-full border px-3 py-1 text-xs font-semibold ${gayaStatus(aset.status)}`}>{aset.status}</span>
          <span className={`rounded-full border px-3 py-1 text-xs font-semibold ${gayaRisiko(aset.risiko)}`}>Risiko {aset.risiko}%</span>
        </div>
        <div className="absolute bottom-4 left-4 right-4 text-white">
          <div className="text-lg font-semibold">{aset.nama}</div>
          <div className="mt-1 flex items-center gap-2 text-xs text-white/85"><QrCode className="h-3.5 w-3.5" /> {aset.id}</div>
        </div>
      </div>
      <div className="p-5">
        <div className="grid grid-cols-3 gap-3">
          <div className="rounded-2xl bg-slate-50 p-3"><div className="text-[11px] text-slate-500">Utilisasi</div><div className="mt-1 font-semibold text-slate-950">{aset.utilisasi}%</div></div>
          <div className="rounded-2xl bg-slate-50 p-3"><div className="text-[11px] text-slate-500">ROI</div><div className="mt-1 font-semibold text-slate-950">{aset.roi}%</div></div>
          <div className="rounded-2xl bg-slate-50 p-3"><div className="text-[11px] text-slate-500">Kondisi</div><div className="mt-1 truncate font-semibold text-slate-950">{aset.kondisi}</div></div>
        </div>
        <div className="mt-5 space-y-3 text-sm">
          <div className="flex gap-3 text-slate-700"><MapPin className="mt-0.5 h-4 w-4 shrink-0 text-slate-400" />{aset.lokasi || "Lokasi belum diisi"}</div>
          <div className="flex gap-3 text-slate-700"><Users className="mt-0.5 h-4 w-4 shrink-0 text-slate-400" />{aset.pemegang || "Pemegang belum diisi"}</div>
          <div className="flex gap-3 text-slate-700"><Gauge className="mt-0.5 h-4 w-4 shrink-0 text-slate-400" />{aset.telemetri}</div>
          <div className="flex gap-3 text-slate-700"><ShieldCheck className="mt-0.5 h-4 w-4 shrink-0 text-slate-400" />{aset.dokumen}</div>
        </div>
        <div className="mt-5 rounded-2xl border border-slate-200 bg-slate-50 p-4">
          <div className="flex items-start gap-3">
            <div className="rounded-xl bg-white p-2 shadow-sm"><Bell className="h-4 w-4 text-slate-700" /></div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Aksi Cerdas Berikutnya</div>
              <div className="mt-1 text-sm font-medium text-slate-900">{aset.aksiBerikutnya}</div>
              <div className="mt-1 text-xs text-slate-500">{aset.alur}</div>
            </div>
          </div>
        </div>
        <div className="mt-5 grid grid-cols-3 gap-2">
          <button onClick={() => onOpen(aset)} className="rounded-2xl bg-slate-950 px-3 py-3 text-xs font-semibold text-white hover:bg-slate-800">Detail</button>
          <button onClick={() => onEdit(aset)} className="flex items-center justify-center gap-1 rounded-2xl border border-slate-200 px-3 py-3 text-xs font-semibold hover:bg-slate-50"><Edit3 className="h-3.5 w-3.5" /> Edit</button>
          <button onClick={() => onDelete(aset.id)} className="flex items-center justify-center gap-1 rounded-2xl border border-red-100 bg-red-50 px-3 py-3 text-xs font-semibold text-red-700 hover:bg-red-100"><Trash2 className="h-3.5 w-3.5" /> Hapus</button>
        </div>
      </div>
    </motion.article>
  );
}

function DetailAset({ aset, onClose, onEdit, onCreateWorkOrder }) {
  return (
    <AnimatePresence>
      {aset && (
        <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-sm">
          <motion.div initial={{ x: 520 }} animate={{ x: 0 }} exit={{ x: 520 }} transition={{ type: "spring", damping: 30, stiffness: 260 }} className="absolute right-0 top-0 h-full w-full max-w-2xl overflow-y-auto bg-white p-6 shadow-2xl">
            <div className="flex items-start justify-between gap-4">
              <div>
                <div className="text-xs font-semibold uppercase tracking-wide text-slate-500">Detail Aset 360</div>
                <h2 className="mt-1 text-2xl font-bold tracking-tight text-slate-950">{aset.nama}</h2>
                <div className="mt-2 flex flex-wrap gap-2">
                  <span className={`rounded-full border px-3 py-1 text-xs font-semibold ${gayaStatus(aset.status)}`}>{aset.status}</span>
                  <span className={`rounded-full border px-3 py-1 text-xs font-semibold ${gayaRisiko(aset.risiko)}`}>Risiko {aset.risiko}%</span>
                  <span className="rounded-full border border-slate-200 bg-slate-50 px-3 py-1 text-xs font-semibold text-slate-700">{aset.id}</span>
                </div>
              </div>
              <button onClick={onClose} className="rounded-2xl border border-slate-200 p-3 hover:bg-slate-50"><X className="h-5 w-5" /></button>
            </div>
            <img src={aset.gambar} alt={aset.nama} className="mt-6 h-64 w-full rounded-[2rem] object-cover" />
            <div className="mt-6 grid gap-3 sm:grid-cols-3">
              <div className="rounded-3xl bg-slate-50 p-4"><div className="text-xs text-slate-500">Nilai Buku</div><div className="mt-1 font-bold">{aset.nilaiBuku}</div></div>
              <div className="rounded-3xl bg-slate-50 p-4"><div className="text-xs text-slate-500">Pendapatan</div><div className="mt-1 font-bold">{aset.pendapatan}</div></div>
              <div className="rounded-3xl bg-slate-50 p-4"><div className="text-xs text-slate-500">Biaya</div><div className="mt-1 font-bold">{aset.biaya}</div></div>
            </div>
            <div className="mt-6 grid gap-5 lg:grid-cols-2">
              <section className="rounded-[2rem] border border-slate-200 p-5">
                <h3 className="font-bold">Identitas & Lokasi</h3>
                <div className="mt-4 space-y-3 text-sm text-slate-700">
                  <div className="flex justify-between gap-4"><span>Kategori</span><b>{aset.kategori}</b></div>
                  <div className="flex justify-between gap-4"><span>Tipe</span><b>{aset.tipe}</b></div>
                  <div className="flex justify-between gap-4"><span>Cabang</span><b>{aset.cabang || "-"}</b></div>
                  <div className="flex justify-between gap-4"><span>Lokasi</span><b>{aset.lokasi || "-"}</b></div>
                  <div className="flex justify-between gap-4"><span>Pemegang</span><b className="text-right">{aset.pemegang || "-"}</b></div>
                </div>
              </section>
              <section className="rounded-[2rem] border border-slate-200 p-5">
                <h3 className="font-bold">Kepatuhan & Telemetri</h3>
                <div className="mt-4 space-y-3 text-sm text-slate-700">
                  <div className="flex gap-3"><ShieldCheck className="h-4 w-4 text-emerald-600" />{aset.dokumen}</div>
                  <div className="flex gap-3"><Gauge className="h-4 w-4 text-blue-600" />{aset.telemetri}</div>
                  <div className="flex gap-3"><Bell className="h-4 w-4 text-amber-600" />{aset.aksiBerikutnya}</div>
                </div>
              </section>
            </div>
            {aset.kategori === "Kendaraan" && (
              <section className="mt-5 rounded-[2rem] border border-slate-200 p-5">
                <h3 className="font-bold">Data Kendaraan Sewa</h3>
                <div className="mt-4 grid gap-3 text-sm text-slate-700 sm:grid-cols-2">
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>Warna</span><b>{aset.warna || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>Tahun</span><b>{aset.tahunKendaraan || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>No. Rangka</span><b className="text-right">{aset.nomorRangka || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>No. Mesin</span><b className="text-right">{aset.nomorMesin || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>Perusahaan User</span><b className="text-right">{aset.perusahaanUser || aset.pemegang || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>Wilayah</span><b>{aset.wilayah || aset.cabang || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>Tanggal STNK</span><b>{aset.tanggalSTNK || "-"}</b></div>
                  <div className="flex justify-between gap-4 rounded-2xl bg-slate-50 p-3"><span>BPKB</span><b>{aset.memilikiBPKB || "-"}</b></div>
                  <div className="rounded-2xl bg-slate-50 p-3 sm:col-span-2"><span>Alamat User</span><div className="mt-1 font-semibold text-slate-950">{aset.alamatUser || "-"}</div></div>
                </div>
              </section>
            )}
            <section className="mt-5 rounded-[2rem] border border-slate-200 p-5">
              <h3 className="font-bold">Timeline Siklus Hidup</h3>
              <div className="mt-4 space-y-3">
                {(aset.riwayat || []).map((item, index) => (
                  <div key={`${item}-${index}`} className="flex items-center gap-3 rounded-2xl bg-slate-50 p-3 text-sm">
                    <div className="flex h-7 w-7 items-center justify-center rounded-full bg-white text-xs font-bold shadow-sm">{index + 1}</div>
                    <span>{item}</span>
                  </div>
                ))}
              </div>
            </section>
            <div className="mt-6 grid gap-3 sm:grid-cols-3">
              <button onClick={() => onEdit(aset)} className="rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white">Edit Aset</button>
              <button onClick={() => onCreateWorkOrder(aset)} className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold hover:bg-slate-50">Buat Work Order</button>
              <button className="rounded-2xl border border-slate-200 px-4 py-3 text-sm font-semibold hover:bg-slate-50">Generate QR</button>
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
}

export default function AdvancedAssetManagementApp() {
  const [daftarAset, setDaftarAset] = useState(dataAsetAwal);
  const [pencarian, setPencarian] = useState("");
  const [filter, setFilter] = useState("Semua");
  const [detailAset, setDetailAset] = useState(null);
  const [formAset, setFormAset] = useState(null);
  const [modeForm, setModeForm] = useState(null);
  const [daftarWorkOrder, setDaftarWorkOrder] = useState(dataWorkOrderAwal);
  const [asetWorkOrder, setAsetWorkOrder] = useState(null);
  const [statusImport, setStatusImport] = useState("");

  const asetTampil = useMemo(() => daftarAset.filter((aset) => {
    const cocokCari = `${aset.nama} ${aset.id} ${aset.kategori} ${aset.status} ${aset.cabang}`.toLowerCase().includes(pencarian.toLowerCase());
    const cocokFilter = filter === "Semua" || aset.status === filter || aset.kategori === filter;
    return cocokCari && cocokFilter;
  }), [daftarAset, pencarian, filter]);

  const totalAset = daftarAset.length;
  const asetBerisiko = hitungAsetBerisiko(daftarAset);
  const rataUtilisasi = hitungRataUtilisasi(daftarAset);
  const slaHariIni = tugasOperasional.length + daftarWorkOrder.filter((wo) => !["Selesai", "Ditutup"].includes(wo.status)).length;

  const bukaTambah = () => {
    setFormAset(null);
    setModeForm("tambah");
  };

  const bukaEdit = (aset) => {
    setFormAset(aset);
    setModeForm("edit");
    setDetailAset(null);
  };

  const tutupForm = () => {
    setFormAset(null);
    setModeForm(null);
  };

  const simpanAset = (aset) => {
    setDaftarAset((prev) => prev.some((item) => item.id === aset.id) ? prev.map((item) => item.id === aset.id ? aset : item) : [aset, ...prev]);
    tutupForm();
  };

  const hapusAset = (id) => {
    setDaftarAset((prev) => prev.filter((aset) => aset.id !== id));
    setDaftarWorkOrder((prev) => prev.filter((wo) => wo.asetId !== id));
    if (detailAset?.id === id) setDetailAset(null);
  };

  const bukaWorkOrder = (aset) => {
    setDetailAset(null);
    setAsetWorkOrder(aset);
  };

  const tutupWorkOrder = () => setAsetWorkOrder(null);

  const simpanWorkOrder = (workOrder) => {
    setDaftarWorkOrder((prev) => [workOrder, ...prev]);
    setDaftarAset((prev) => prev.map((aset) => {
      if (aset.id !== workOrder.asetId) return aset;
      return {
        ...aset,
        status: workOrder.status === "Draft" ? aset.status : "Maintenance",
        aksiBerikutnya: `${workOrder.jenis} - ${workOrder.status}`,
        alur: "Work Order → Dikerjakan → QC → Selesai",
        riwayat: [...(aset.riwayat || []), `Work Order ${workOrder.id}`],
      };
    }));
    tutupWorkOrder();
  };

  const importKendaraanSewa = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setStatusImport("Membaca file Excel...");
      const asetBaru = await bacaImportKendaraanSewa(file);
      if (asetBaru.length === 0) {
        setStatusImport("File Excel tidak memiliki data kendaraan yang bisa diimpor.");
        return;
      }
      setDaftarAset((prev) => {
        const idLama = new Set(prev.map((aset) => aset.id));
        const asetUnik = asetBaru.map((aset, index) => idLama.has(aset.id) ? { ...aset, id: `${aset.id}-${index + 1}` } : aset);
        return [...asetUnik, ...prev];
      });
      setStatusImport(`${asetBaru.length} kendaraan sewa berhasil diimpor dari ${file.name}.`);
    } catch (error) {
      setStatusImport(`Import gagal: ${error?.message || "format file tidak terbaca"}. Pastikan memakai format .xlsx/.xls/.csv sesuai template.`);
    } finally {
      event.target.value = "";
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-950">
      <aside className="fixed left-0 top-0 hidden h-screen w-72 border-r border-slate-200 bg-white p-6 lg:block">
        <div className="flex items-center gap-3">
          <div className="rounded-2xl bg-slate-950 p-3 text-white"><Building2 className="h-5 w-5" /></div>
          <div><div className="text-lg font-bold tracking-tight">AssetNexus</div><div className="text-xs text-slate-500">Sistem Operasi Aset</div></div>
        </div>
        <nav className="mt-10 space-y-2 text-sm font-medium">
          {[[BarChart3, "Pusat Komando"], [Archive, "Aset 360"], [Car, "Manajemen Kendaraan"], [Laptop, "Perangkat IT"], [Truck, "Operasi Rental"], [Wrench, "Maintenance"], [ClipboardCheck, "Audit & Stock Opname"], [ShieldCheck, "Kepatuhan"], [Database, "Model Data & API"]].map(([Icon, label], index) => (
            <button key={label} className={`flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left transition ${index === 0 ? "bg-slate-950 text-white" : "text-slate-600 hover:bg-slate-100"}`}><Icon className="h-4 w-4" /> {label}</button>
          ))}
        </nav>
        <div className="absolute bottom-6 left-6 right-6 rounded-3xl bg-slate-950 p-5 text-white">
          <div className="flex items-center gap-2 text-sm font-semibold"><Sparkles className="h-4 w-4" /> Lapisan AI</div>
          <p className="mt-2 text-xs leading-5 text-slate-300">Prediksi maintenance, rekomendasi aset, dan deteksi anomali aktif.</p>
        </div>
      </aside>

      <main className="lg:pl-72">
        <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/85 px-5 py-4 backdrop-blur md:px-8">
          <div className="flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div><div className="text-sm font-medium text-slate-500">Pusat Komando</div><h1 className="text-2xl font-bold tracking-tight md:text-3xl">Sistem Manajemen Aset Enterprise</h1></div>
            <div className="flex items-center gap-3">
              <button className="rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold shadow-sm hover:bg-slate-50"><RefreshCw className="mr-2 inline h-4 w-4" /> Sinkron</button>
              <label className="cursor-pointer rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm font-semibold shadow-sm hover:bg-slate-50">
                <Upload className="mr-2 inline h-4 w-4" /> Import Kendaraan Sewa
                <input type="file" accept=".xlsx,.xls,.csv" onChange={importKendaraanSewa} className="hidden" />
              </label>
              <button onClick={bukaTambah} className="rounded-2xl bg-slate-950 px-4 py-3 text-sm font-semibold text-white shadow-sm hover:bg-slate-800"><Plus className="mr-2 inline h-4 w-4" /> Tambah Aset</button>
            </div>
          </div>
        </header>

        <section className="p-5 md:p-8">
          <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
            <KartuMetrik icon={Archive} label="Total Aset Terkelola" nilai={totalAset.toLocaleString("id-ID")} detail="Data dapat ditambah, diedit, dan dihapus" />
            <KartuMetrik icon={Activity} label="Rata-rata Utilisasi" nilai={`${rataUtilisasi}%`} detail="Gabungan rental dan penggunaan internal" />
            <KartuMetrik icon={AlertTriangle} label="Aset Berisiko" nilai={asetBerisiko} detail="Dokumen expired, service overdue, atau downtime tinggi" />
            <KartuMetrik icon={CalendarClock} label="SLA Hari Ini" nilai={slaHariIni} detail="Delivery, return, maintenance, dan approval" />
          </div>

          <div className="mt-8 rounded-[2rem] border border-slate-200 bg-white p-5 shadow-sm">
            <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
              <div><h2 className="text-xl font-bold tracking-tight">Alur Kerja Aset Cerdas</h2><p className="mt-1 text-sm text-slate-500">Alur modular, siap audit, dan berbasis rule engine.</p></div>
              <div className="flex flex-wrap gap-2">
                {alurKerja.map((item, index) => <div key={item} className="flex items-center gap-2"><div className="rounded-full border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-700">{item}</div>{index < alurKerja.length - 1 && <ChevronRight className="hidden h-4 w-4 text-slate-300 sm:block" />}</div>)}
              </div>
            </div>
          </div>

          <div className="mt-8 grid gap-4 md:grid-cols-2 xl:grid-cols-3">
            {modul.map(({ icon: Icon, judul, deskripsi }) => <div key={judul} className="rounded-3xl border border-slate-200 bg-white p-5 shadow-sm"><div className="flex items-start gap-4"><div className="rounded-2xl bg-slate-100 p-3"><Icon className="h-5 w-5 text-slate-700" /></div><div><div className="font-semibold text-slate-950">{judul}</div><p className="mt-1 text-sm leading-6 text-slate-500">{deskripsi}</p></div></div></div>)}
          </div>

          <div className="mt-10 grid gap-5 xl:grid-cols-3">
            <section className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm xl:col-span-2">
              <div className="flex items-center gap-3"><Workflow className="h-5 w-5 text-slate-600" /><h2 className="text-xl font-bold tracking-tight">Papan Tugas Operasional</h2></div>
              <div className="mt-5 grid gap-3 lg:grid-cols-3">
                {tugasOperasional.map((tugas) => <div key={tugas.id} className="rounded-3xl border border-slate-200 bg-slate-50 p-4"><div className="flex items-center justify-between"><span className="text-xs font-bold text-slate-500">{tugas.id}</span><span className={`rounded-full px-2.5 py-1 text-[11px] font-bold ${tugas.level === "Kritis" ? "bg-red-100 text-red-700" : tugas.level === "Tinggi" ? "bg-amber-100 text-amber-700" : "bg-blue-100 text-blue-700"}`}>{tugas.level}</span></div><div className="mt-4 font-semibold">{tugas.judul}</div><div className="mt-2 text-xs text-slate-500">{tugas.jenis} · {tugas.aset}</div><div className="mt-4 flex items-center justify-between text-xs"><span>{tugas.pic}</span><b>{tugas.sla}</b></div></div>)}
              </div>
            </section>
            <section className="rounded-[2rem] border border-slate-200 bg-slate-950 p-6 text-white shadow-sm">
              <div className="flex items-center gap-3"><GitBranch className="h-5 w-5" /><h2 className="text-xl font-bold tracking-tight">Matriks Approval</h2></div>
              <div className="mt-5 space-y-3 text-sm text-slate-200"><div className="rounded-2xl bg-white/10 p-4">Repair &gt; Rp5 juta → Supervisor + Finance</div><div className="rounded-2xl bg-white/10 p-4">Disposal aset → Asset Manager + Direktur</div><div className="rounded-2xl bg-white/10 p-4">Harga rental khusus → Sales Manager</div><div className="rounded-2xl bg-white/10 p-4">Mutasi antar cabang → Dua Branch Manager</div></div>
            </section>
          </div>

          <div className="mt-10 rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm">
            <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
              <div>
                <h2 className="text-xl font-bold tracking-tight">Import Excel Kendaraan Sewa</h2>
                <p className="mt-1 text-sm text-slate-500">Kolom yang didukung: no, nama kendaraan, warna, tahun kendaraan, nomor rangka, nomor mesin, perusahaan user, alamat user, wilayah, tanggal STNK, memiliki BPKB.</p>
              </div>
              <label className="cursor-pointer rounded-2xl bg-slate-950 px-4 py-3 text-center text-sm font-semibold text-white shadow-sm hover:bg-slate-800">
                <Upload className="mr-2 inline h-4 w-4" /> Upload Excel
                <input type="file" accept=".xlsx,.xls,.csv" onChange={importKendaraanSewa} className="hidden" />
              </label>
            </div>
            {statusImport && <div className="mt-4 rounded-2xl bg-slate-50 px-4 py-3 text-sm font-semibold text-slate-700">{statusImport}</div>}
          </div>

          <div className="mt-10 flex flex-col gap-4 md:flex-row md:items-center md:justify-between">
            <div><h2 className="text-2xl font-bold tracking-tight">Daftar Aset yang Bisa Diedit</h2><p className="mt-1 text-sm text-slate-500">Tambah, edit, hapus, cari, filter, dan buka detail Aset 360.</p></div>
            <div className="flex flex-col gap-3 sm:flex-row"><div className="relative"><Search className="absolute left-3 top-3.5 h-4 w-4 text-slate-400" /><input value={pencarian} onChange={(e) => setPencarian(e.target.value)} placeholder="Cari aset..." className="w-full rounded-2xl border border-slate-200 bg-white py-3 pl-10 pr-4 text-sm shadow-sm outline-none focus:border-slate-400 sm:w-72" /></div><div className="relative"><Filter className="absolute left-3 top-3.5 h-4 w-4 text-slate-400" /><select value={filter} onChange={(e) => setFilter(e.target.value)} className="w-full appearance-none rounded-2xl border border-slate-200 bg-white py-3 pl-10 pr-10 text-sm font-medium shadow-sm outline-none focus:border-slate-400 sm:w-56">{["Semua", ...opsiKategori, ...opsiStatus].map((item) => <option key={item}>{item}</option>)}</select></div></div>
          </div>

          <div className="mt-5 grid gap-5 md:grid-cols-2 xl:grid-cols-4">
            {asetTampil.map((aset) => <KartuAset key={aset.id} aset={aset} onOpen={setDetailAset} onEdit={bukaEdit} onDelete={hapusAset} />)}
          </div>

          <div className="mt-10 rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm">
            <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
              <div><h2 className="text-xl font-bold tracking-tight">Daftar Work Order</h2><p className="mt-1 text-sm text-slate-500">Work Order dibuat dari tombol “Buat Work Order” di halaman Detail Aset 360.</p></div>
              <span className="rounded-full bg-slate-100 px-4 py-2 text-xs font-semibold text-slate-700">{daftarWorkOrder.length} WO</span>
            </div>
            <div className="mt-5 grid gap-3 xl:grid-cols-3">
              {daftarWorkOrder.map((wo) => (
                <article key={wo.id} className="rounded-3xl border border-slate-200 bg-slate-50 p-4">
                  <div className="flex items-center justify-between gap-3"><span className="text-xs font-bold text-slate-500">{wo.id}</span><span className={`rounded-full px-2.5 py-1 text-[11px] font-bold ${wo.prioritas === "Kritis" ? "bg-red-100 text-red-700" : wo.prioritas === "Tinggi" ? "bg-amber-100 text-amber-700" : "bg-blue-100 text-blue-700"}`}>{wo.prioritas}</span></div>
                  <h3 className="mt-4 font-semibold text-slate-950">{wo.jenis}</h3>
                  <p className="mt-1 text-xs text-slate-500">{wo.namaAset} · {wo.asetId}</p>
                  <div className="mt-4 grid grid-cols-2 gap-2 text-xs">
                    <div className="rounded-2xl bg-white p-3"><span className="text-slate-500">Status</span><div className="mt-1 font-bold">{wo.status}</div></div>
                    <div className="rounded-2xl bg-white p-3"><span className="text-slate-500">PIC</span><div className="mt-1 font-bold">{wo.pic}</div></div>
                    <div className="rounded-2xl bg-white p-3"><span className="text-slate-500">Due</span><div className="mt-1 font-bold">{wo.jatuhTempo}</div></div>
                    <div className="rounded-2xl bg-white p-3"><span className="text-slate-500">Biaya</span><div className="mt-1 font-bold">{wo.estimasiBiaya}</div></div>
                  </div>
                  <p className="mt-4 line-clamp-3 text-sm text-slate-600">{wo.keluhan}</p>
                </article>
              ))}
            </div>
          </div>

          <div className="mt-10 grid gap-5 xl:grid-cols-2">
            <section className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm"><div className="flex items-center gap-3"><Database className="h-5 w-5 text-slate-600" /><h3 className="text-lg font-bold tracking-tight">Blueprint Database</h3></div><div className="mt-5 flex flex-wrap gap-2">{tabelDatabase.map((table) => <span key={table} className="rounded-full border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-semibold text-slate-700">{table}</span>)}</div></section>
            <section className="rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm"><div className="flex items-center gap-3"><Layers3 className="h-5 w-5 text-slate-600" /><h3 className="text-lg font-bold tracking-tight">Kontrak API Inti</h3></div><div className="mt-5 space-y-2">{apiEndpoints.map((endpoint) => <code key={endpoint} className="block rounded-2xl bg-slate-950 px-4 py-3 text-xs font-semibold text-slate-100">{endpoint}</code>)}</div></section>
          </div>

          <div className="mt-10 rounded-[2rem] border border-slate-200 bg-white p-6 shadow-sm"><h3 className="text-lg font-bold tracking-tight">Rule Engine</h3><div className="mt-5 space-y-3">{["Jika STNK/KIR expired < 14 hari, aset masuk peringatan kepatuhan dan tidak bisa delivery.", "Jika return inspection belum QC Passed, status tidak bisa menjadi Tersedia.", "Jika biaya repair melebihi limit cabang, approval otomatis naik ke Finance Manager.", "Jika aset idle lebih dari 30 hari, sistem merekomendasikan promo rental atau review disposal."].map((rule) => <div key={rule} className="flex items-start gap-3 rounded-2xl bg-slate-50 p-4 text-sm text-slate-700"><CheckCircle2 className="mt-0.5 h-4 w-4 shrink-0 text-emerald-600" /> {rule}</div>)}</div></div>
        </section>
      </main>

      <DetailAset aset={detailAset} onClose={() => setDetailAset(null)} onEdit={bukaEdit} onCreateWorkOrder={bukaWorkOrder} />
      {asetWorkOrder && <FormWorkOrder aset={asetWorkOrder} onClose={tutupWorkOrder} onSave={simpanWorkOrder} />}
      {modeForm && <FormAset aset={formAset} onClose={tutupForm} onSave={simpanAset} />}
    </div>
  );
}
