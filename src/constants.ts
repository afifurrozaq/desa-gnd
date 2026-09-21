/**
 * Application Static Text & Configuration Constants
 * File ini memuat seluruh data teks statis (List Dapukan, List Kelompok, Kategori Jamaah, Dropdown Options, Jadwal Sambung, dll)
 * untuk mempermudah penambahan atau pengubahan data statis aplikasi di satu tempat.
 */

import { MosqueLocation } from './types';

export interface DropdownOption<T = string> {
  value: T;
  label: string;
}

// --- APP IDENTITY & CREDITS ---
export const APP_TITLE = 'Desa GND Management';
export const APP_SUBTITLE = 'Sistem Manajemen Desa GND';
export const APP_CREDIT_STUDIO = 'ARSH Studio';
export const APP_CREDIT_URL = 'https://github.com/afifurrozaq';

// --- LIST KELOMPOK / LOKASI MASJID ---
export const MOSQUE_LOCATION_OPTIONS: DropdownOption<MosqueLocation | ''>[] = [
  { value: '', label: 'Pilih Kelompok' },
  { value: 'Kramat Batu', label: 'Kramat Batu' },
  { value: 'Karya Utama', label: 'Karya Utama' },
  { value: 'Radio Dalam', label: 'Radio Dalam' },
  { value: 'Cipete', label: 'Cipete' },
  { value: 'Antena', label: 'Antena' }
];

export const MOSQUE_LOCATIONS: MosqueLocation[] = MOSQUE_LOCATION_OPTIONS
  .filter((opt): opt is DropdownOption<MosqueLocation> => opt.value !== '')
  .map(opt => opt.value);

export const DEFAULT_LOCATION: MosqueLocation = 'Pilih Kelompok' as MosqueLocation;

// Kode Prefix Anggota per Kelompok
export const LOCATION_PREFIXES: Record<string, string> = {
  'Kramat Batu': 'KB',
  'Karya Utama': 'KU',
  'Radio Dalam': 'RD',
  'Cipete': 'CP',
  'Antena': 'AN'
};

// --- LIST DAPUKAN ---
export const DAPUKAN_OPTIONS_OBJ: DropdownOption[] = [
  { value: "Ides", label: "Ides" },
  { value: "Wides", label: "Wides" },
  { value: "Koor. Lupg", label: "Koor. Lupg" },
  { value: "Bosdes", label: "Bosdes" },
  { value: "Mubdes", label: "Mubdes" },
  { value: "Tim Aghniya'", label: "Tim Aghniya'" },
  { value: "Ku Des", label: "Ku Des" },
  { value: "Tim Bk", label: "Tim Bk" },
  { value: "Tim Bacaan", label: "Tim Bacaan" },
  { value: "Tim Basyiron Wa Nadziron", label: "Tim Basyiron Wa Nadziron" },
  { value: "Tim Benda Sb", label: "Tim Benda Sb" },
  { value: "Tim DhuaFa'", label: "Tim DhuaFa'" },
  { value: "Tim Faraoid", label: "Tim Faraoid" },
  { value: "Tim Gambuh", label: "Tim Gambuh" },
  { value: "Tim Haji", label: "Tim Haji" },
  { value: "Tim Keluarga Bahagia", label: "Tim Keluarga Bahagia" },
  { value: "Tim Kematian", label: "Tim Kematian" },
  { value: "Tim Manula", label: "Tim Manula" },
  { value: "Tim Mondar Mandir", label: "Tim Mondar Mandir" },
  { value: "Tim Muballigh", label: "Tim Muballigh" },
  { value: "Tim Organisasi", label: "Tim Organisasi" },
  { value: "Tim Pembangunan", label: "Tim Pembangunan" },
  { value: "Tim Pkw", label: "Tim Pkw" },
  { value: "Tim Penyelesaian", label: "Tim Penyelesaian" },
  { value: "Tim Pramuka", label: "Tim Pramuka" },
  { value: "Tim Sarjana", label: "Tim Sarjana" },
  { value: "Tim Ub", label: "Tim Ub" },
  { value: "Tim Zakat", label: "Tim Zakat" },
  { value: "Tim Cai & Remaja (Karemdes)", label: "Tim Cai & Remaja (Karemdes)" },
  { value: "Keputrian Des", label: "Keputrian Des" },
  { value: "Ikel", label: "Ikel" },
  { value: "Wikel", label: "Wikel" },
  { value: "Pjkbm", label: "Pjkbm" },
  { value: "Boskel", label: "Boskel" },
  { value: "Mubkel", label: "Mubkel" },
  { value: "Ku Kel", label: "Ku Kel" },
  { value: "Pakar Pendidik", label: "Pakar Pendidik" },
  { value: "Ptk", label: "Ptk" },
  { value: "Keputkel", label: "Keputkel" },
  { value: "Rokyah", label: "Rokyah" }
];

export const DAPUKAN_OPTIONS = DAPUKAN_OPTIONS_OBJ.map(opt => opt.value);

// --- KATEGORI JAMAAH ---
export const JAMAAH_CATEGORY_OPTIONS = [
  { value: 'UMUM', label: 'UMUM' },
  { value: 'GPN_A', label: 'GPN A' },
  { value: 'GPN_B', label: 'GPN B' },
  { value: 'GPN_B_PLUS', label: 'GPN B+' },
  { value: 'AR', label: 'AR' },
  { value: 'APR', label: 'APR' },
  { value: 'ACR', label: 'ACR' },
  { value: 'AUD', label: 'AUD' },
  { value: 'DUDA', label: 'DUDA' },
  { value: 'JANDA', label: 'JANDA' }
] as const;

export const DEFAULT_JAMAAH_CATEGORY = JAMAAH_CATEGORY_OPTIONS[0].value;

// --- STATUS PERNIKAHAN ---
export const MARITAL_STATUS_OPTIONS: DropdownOption[] = [
  { value: 'Belum Menikah', label: 'Belum Menikah' },
  { value: 'Nikah', label: 'Nikah' },
  { value: 'Janda', label: 'Janda' },
  { value: 'Duda', label: 'Duda' }
];

// --- SESI SAMBUNG / PENGAJIAN ---
export const SESSION_TYPE_OPTIONS: DropdownOption<'Kelompok' | 'Desa' | 'Acara'>[] = [
  { value: 'Kelompok', label: 'Kelompok' },
  { value: 'Desa', label: 'Desa' },
  { value: 'Acara', label: 'Acara' }
];

// --- GOLONGAN DARAH ---
export const BLOOD_TYPE_OPTIONS: DropdownOption[] = [
  { value: 'A', label: 'A' },
  { value: 'B', label: 'B' },
  { value: 'AB', label: 'AB' },
  { value: 'O', label: 'O' },
  { value: '-', label: '-' }
];

// --- JENIS KELAMIN ---
export const GENDER_OPTIONS: DropdownOption[] = [
  { value: 'Laki-laki', label: 'Laki-laki' },
  { value: 'Perempuan', label: 'Perempuan' }
];

// --- PENDIDIKAN TERAKHIR ---
export const EDUCATION_OPTIONS: DropdownOption[] = [
  { value: 'SD', label: 'SD / Sederajat' },
  { value: 'SMP', label: 'SMP / Sederajat' },
  { value: 'SMA', label: 'SMA / SMK / Sederajat' },
  { value: 'D3', label: 'D3 / Diploma' },
  { value: 'S1', label: 'S1 / Sarjana' },
  { value: 'S2', label: 'S2 / Magister' },
  { value: 'S3', label: 'S3 / Doktor' },
  { value: 'Lainnya', label: 'Lainnya' }
];

// --- OPSI YA / TIDAK ---
export const YES_NO_OPTIONS: DropdownOption[] = [
  { value: 'Ya', label: 'Ya' },
  { value: 'Tidak', label: 'Tidak' }
];

// --- STATUS PRESENSI ---
export const ATTENDANCE_STATUS_OPTIONS: DropdownOption<'hadir' | 'izin'>[] = [
  { value: 'hadir', label: 'Hadir' },
  { value: 'izin', label: 'Izin' }
];

// --- JADWAL SAMBUNG / PRESENSI RUTIN ABSENSI ---
export interface JadwalSambung {
  day: number; // 0: Minggu, 1: Senin, ..., 6: Sabtu
  dayName: string;
  sessions: string[];
}

export const JADWAL_SAMBUNG_RUTIN: JadwalSambung[] = [
  { day: 1, dayName: 'Senin', sessions: ['UMUM (Kelompok/Desa)', 'GPN'] },
  { day: 2, dayName: 'Selasa', sessions: ['GPN'] },
  { day: 3, dayName: 'Rabu', sessions: ['UMUM (Kelompok/Desa)', 'GPN'] },
  { day: 4, dayName: 'Kamis', sessions: ['UMUM (Kelompok/Desa)', 'GPN'] },
  { day: 5, dayName: 'Jumat', sessions: ['Pengajian Rutin Jumat'] },
  { day: 6, dayName: 'Sabtu', sessions: ['Acara Kelompok/Desa'] },
  { day: 0, dayName: 'Minggu', sessions: ['Acara Kelompok/Desa'] }
];

/**
 * Helper untuk mendapatkan teks status / jadwal sambung hari ini di halaman absensi
 */
export const getTodayJadwalSambung = (): string => {
  const today = new Date().getDay(); // 0-6 (Sun-Sat)
  const current = JADWAL_SAMBUNG_RUTIN.find(s => s.day === today);
  return current && current.sessions.length > 0 ? current.sessions.join(', ') : 'Tidak ada jadwal rutin';
};

// --- KATEGORI KEGIATAN & WARNA TEMA DAN UKURAN TAMPILAN ---
export interface ActivityCategoryConfig {
  id: string;
  label: string;
  bg: string;
  text: string;
  border: string;
  dot: string;
  badge: string;
  width?: string;       // Pengaturan lebar box (e.g. 'col-span-1', 'md:col-span-2', 'w-full')
  height?: string;      // Pengaturan tinggi box (e.g. 'min-h-[180px]', 'min-h-[220px]', 'h-64')
  cardWidth?: string;   // Pengaturan opsi lebar kartu
  cardHeight?: string;  // Pengaturan opsi tinggi kartu
}

export const ACTIVITY_CATEGORIES: ActivityCategoryConfig[] = [
  { 
    id: 'Sambung', 
    label: 'Sambung / Pengajian Sambung', 
    bg: 'bg-purple-50', 
    text: 'text-purple-700', 
    border: 'border-purple-200', 
    dot: 'bg-purple-500', 
    badge: 'bg-purple-600 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'row-span-2 md:row-span-2',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Pengajian Rutin', 
    label: 'Pengajian Rutin', 
    bg: 'bg-emerald-50', 
    text: 'text-emerald-700', 
    border: 'border-emerald-200', 
    dot: 'bg-emerald-500', 
    badge: 'bg-emerald-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Musyawarah / Rapat', 
    label: 'Musyawarah / Rapat', 
    bg: 'bg-indigo-50', 
    text: 'text-indigo-700', 
    border: 'border-indigo-200', 
    dot: 'bg-indigo-500', 
    badge: 'bg-indigo-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Kerja Bakti', 
    label: 'Kerja Bakti', 
    bg: 'bg-amber-50', 
    text: 'text-amber-700', 
    border: 'border-amber-200', 
    dot: 'bg-amber-500', 
    badge: 'bg-amber-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Pembinaan Generus', 
    label: 'Pembinaan Generus', 
    bg: 'bg-sky-50', 
    text: 'text-sky-700', 
    border: 'border-sky-200', 
    dot: 'bg-sky-500', 
    badge: 'bg-sky-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Keputrian', 
    label: 'Keputrian', 
    bg: 'bg-rose-50', 
    text: 'text-rose-700', 
    border: 'border-rose-200', 
    dot: 'bg-rose-500', 
    badge: 'bg-rose-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Khotmil Qur\'an', 
    label: 'Khotmil Qur\'an', 
    bg: 'bg-teal-50', 
    text: 'text-teal-700', 
    border: 'border-teal-200', 
    dot: 'bg-teal-500', 
    badge: 'bg-teal-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Acara Khusus / Sambung Desa', 
    label: 'Acara Khusus / Sambung Desa', 
    bg: 'bg-purple-50', 
    text: 'text-purple-700', 
    border: 'border-purple-200', 
    dot: 'bg-purple-500', 
    badge: 'bg-purple-500 text-white',
    width: 'col-span-1 md:col-span-2',
    height: 'min-h-[220px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Olahraga & Keakraban', 
    label: 'Olahraga & Keakraban', 
    bg: 'bg-orange-50', 
    text: 'text-orange-700', 
    border: 'border-orange-200', 
    dot: 'bg-orange-500', 
    badge: 'bg-orange-500 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  },
  { 
    id: 'Lainnya', 
    label: 'Lainnya', 
    bg: 'bg-slate-50', 
    text: 'text-slate-700', 
    border: 'border-slate-200', 
    dot: 'bg-slate-500', 
    badge: 'bg-slate-600 text-white',
    width: 'col-span-1 md:col-span-1',
    height: 'min-h-[180px]',
    cardWidth: 'w-full',
    cardHeight: 'h-auto'
  }
];

// --- DATES & MONTHS ---
export const MONTH_NAMES = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

export const DAY_NAMES = ['Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab', 'Aha'];
export const DAY_FULL_NAMES = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];

// --- KONDISI & TIPE ASET ---
export const ASSET_STATUS_OPTIONS: DropdownOption[] = [
  { value: 'baik', label: 'Baik' },
  { value: 'rusak', label: 'Rusak' },
  { value: 'perlu perbaikan', label: 'Perlu Perbaikan' },
  { value: 'wakaf', label: 'Wakaf' },
  { value: 'sertifikasi', label: 'Sertifikasi' },
  { value: 'terjual', label: 'Terjual' }
];

export const ASSET_TYPE_OPTIONS: DropdownOption<'barang' | 'tanah'>[] = [
  { value: 'barang', label: 'Barang' },
  { value: 'tanah', label: 'Tanah' }
];

// --- HAK AKSES USER ---
export const USER_ROLE_OPTIONS: DropdownOption[] = [
  { value: 'super_admin', label: 'Super Admin' },
  { value: 'admin_desa', label: 'Admin Desa' },
  { value: 'pengurus', label: 'Pengurus Kelompok' },
  { value: 'jamaah', label: 'Jamaah / Anggota' }
];

