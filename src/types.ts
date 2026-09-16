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
}

export type JamaahCategory = 'UMUM' | 'ACR' | 'APR' | 'GPN';

export interface Jamaah {
  id: string;
  memberId: string; // KB001, KB001-01, etc.
  name: string;
  address: string;
  phone: string;
  location: MosqueLocation;
  photoUrl?: string;
  category: JamaahCategory;
  isKK: boolean;
  kkId?: string; // Reference to the KK's memberId
  familyOrder: number; // 0 for KK, 1, 2, etc. for members
  registeredAt: number;
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
