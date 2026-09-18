/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type UserRole = 'admin' | 'pengurus';

export type MosqueLocation = 
  | 'Kramat Batu' 
  | 'Karya Utama' 
  | 'Radio Dalam' 
  | 'Cipete' 
  | 'Antena';

export interface UserProfile {
  uid: string;
  email: string;
  displayName: string;
  role: UserRole;
  location?: MosqueLocation; // Only for 'pengurus'
  isVerified: boolean;
  createdAt: number;
  spreadsheetId?: string;
}

export type JamaahCategory = 'UMUM' | 'ACR' | 'APR' | 'GPN' | 'DUDA' | 'JANDA' | string;

export interface Jamaah {
  id: string;
  memberId: string; // KB001, KB001-01, etc.
  name: string;
  nickname?: string;
  gender?: string;
  originAddress?: string;
  currentAddress?: string;
  address?: string; // Legacy fallback
  phone: string;
  location: MosqueLocation;
  photoUrl?: string;
  category: JamaahCategory;
  isKK: boolean;
  kkId?: string; // Reference to the KK's memberId
  familyOrder: number; // 0 for KK, 1, 2, etc. for members
  dapukan?: string[];
  positions?: string[]; // Legacy fallback
  registeredAt: number;
  placeOfBirth?: string;
  dateOfBirth?: string;
  fatherName?: string;
  motherName?: string;
  parentPhone?: string;
  lastEducation?: string;
  majorOrClass?: string;
  schoolOrUniversity?: string;
  currentJob?: string;
  workplaceAddress?: string;
  maritalStatus?: string;
  marriageYear?: string;
  spouseName?: string;
  hasJurusKeras?: string;
  hasJurusHalus?: string;
  bloodType?: string;
  hasUbShares?: string;
  previousDapukan?: string;
  isMubaligh?: string;
  hasHajj?: string;
  hajjPortionNumber?: string;
  plannedHajjYear?: string;
  hajjName?: string;
  hajjYear?: string;
  medicalHistory?: string;
}

export interface Asset {
  id: string;
  name: string;
  category: string;
  status: 'baik' | 'rusak' | 'perlu perbaikan' | 'wakaf' | 'sertifikasi' | 'terjual';
  quantity: number;
  location: MosqueLocation | 'Utama';
  lastChecked: number;
  description?: string;
  photoUrl?: string;
  areaSize?: number; // for land in m2
  assetType: 'barang' | 'tanah';
}

export interface Activity {
  id: string;
  title: string;
  description: string;
  date: number;
  type: 'harian' | 'mingguan';
  imageUrls: string[];
  createdBy: string;
  location: MosqueLocation;
}

export interface FacilityStat {
  id: string;
  name: string;
  currentUsage: number;
  capacity: number;
  updatedAt: number;
}

export interface Attendance {
  id: string;
  jamaahId: string;
  jamaahName: string;
  location: MosqueLocation;
  category: JamaahCategory;
  date: number;
  sessionType: 'Kelompok' | 'Desa' | 'Acara';
  day: string;
  status: 'hadir' | 'izin';
  reason?: string;
  isConfirmed?: boolean;
}

export interface UBShopping {
  id: string;
  jamaahId: string;
  jamaahName: string;
  location: MosqueLocation;
  amount: number;
  date: number; // timestamp
  month: number; // 0-11
  year: number;
  note?: string;
  createdAt: number;
}

export interface RegistrationLink {
  id: string;
  token: string;
  location?: MosqueLocation | 'Seluruh Lokasi';
  expiresAt: number;
  createdAt: number;
  createdBy: string;
  note?: string;
  isActive: boolean;
}

