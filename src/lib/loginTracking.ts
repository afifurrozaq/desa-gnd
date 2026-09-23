import { saveData } from './dataService';
import { UserProfile, LoginHistoryEntry } from '../types';

export interface ClientLoginInfo {
  ip: string;
  location: string;
  city: string;
  region: string;
  country: string;
  device: string;
  dateTimeStr: string;
  timestamp: number;
}

/**
 * Mendeteksi jenis perangkat dan browser dari user agent
 */
export function getDeviceInfo(): string {
  if (typeof navigator === 'undefined') return 'Unknown Device';
  const ua = navigator.userAgent;

  // Deteksi OS
  let os = 'Unknown OS';
  if (/windows/i.test(ua)) os = 'Windows';
  else if (/macintosh|mac os x/i.test(ua)) os = 'macOS';
  else if (/android/i.test(ua)) os = 'Android';
  else if (/iphone|ipad|ipod/i.test(ua)) os = 'iOS';
  else if (/linux/i.test(ua)) os = 'Linux';

  // Deteksi Browser
  let browser = 'Browser';
  if (/edg/i.test(ua)) browser = 'Edge';
  else if (/chrome|crios/i.test(ua)) browser = 'Chrome';
  else if (/firefox|fxios/i.test(ua)) browser = 'Firefox';
  else if (/safari/i.test(ua) && !/chrome/i.test(ua)) browser = 'Safari';
  else if (/opr\//i.test(ua)) browser = 'Opera';

  // Deteksi Tipe (Mobile / Tablet / Desktop)
  const isMobile = /mobile|android|iphone|ipad|phone/i.test(ua);
  const type = isMobile ? 'Mobile' : 'Desktop';

  return `${type} (${browser} di ${os})`;
}

/**
 * Format tanggal dan waktu dalam format Indonesia yang rapi (WIB/Lokal)
 */
export function formatLoginDateTime(date: Date = new Date()): string {
  try {
    return new Intl.DateTimeFormat('id-ID', {
      day: 'numeric',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: false
    }).format(date) + ' WIB';
  } catch (e) {
    return date.toLocaleString('id-ID');
  }
}

/**
 * Mendapatkan IP Address dan Lokasi Geografis Pengguna
 */
export async function getClientLoginInfo(): Promise<ClientLoginInfo> {
  const timestamp = Date.now();
  const dateTimeStr = formatLoginDateTime(new Date(timestamp));
  const device = getDeviceInfo();

  let ip = '';
  let city = '';
  let region = '';
  let country = '';
  let location = '';

  // 1. Coba lewat ipwho.is (CORS free, cepat, tanpa API Key)
  try {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 3500);
    const res = await fetch('https://ipwho.is/', { signal: controller.signal });
    clearTimeout(timeoutId);
    if (res.ok) {
      const data = await res.json();
      if (data && data.success !== false) {
        ip = data.ip || '';
        city = data.city || '';
        region = data.region || '';
        country = data.country || '';
        const parts = [city, region, country].filter(Boolean);
        location = parts.length > 0 ? parts.join(', ') : 'Lokasi Terdeteksi';
      }
    }
  } catch (e) {
    // Failover
  }

  // 2. Jika ipwho.is gagal, coba ipapi.co
  if (!ip) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 3000);
      const res = await fetch('https://ipapi.co/json/', { signal: controller.signal });
      clearTimeout(timeoutId);
      if (res.ok) {
        const data = await res.json();
        ip = data.ip || '';
        city = data.city || '';
        region = data.region || '';
        country = data.country_name || '';
        const parts = [city, region, country].filter(Boolean);
        location = parts.length > 0 ? parts.join(', ') : '';
      }
    } catch (e) {
      // Failover
    }
  }

  // 3. Jika masih belum ada, coba endpoint lokal /api/client-info atau api.ipify.org
  if (!ip) {
    try {
      const res = await fetch('/api/client-info');
      if (res.ok) {
        const data = await res.json();
        ip = data.ip || '';
      }
    } catch (e) {}

    if (!ip) {
      try {
        const res = await fetch('https://api.ipify.org?format=json');
        if (res.ok) {
          const data = await res.json();
          ip = data.ip || '';
        }
      } catch (e) {}
    }
  }

  if (!ip) {
    ip = '127.0.0.1';
  }
  if (!location) {
    location = 'Indonesia';
  }

  return {
    ip,
    location,
    city,
    region,
    country,
    device,
    dateTimeStr,
    timestamp
  };
}

/**
 * Parse login history dari UserProfile
 */
export function parseLoginHistory(user?: UserProfile | null): LoginHistoryEntry[] {
  if (!user || !user.loginHistory) return [];
  if (Array.isArray(user.loginHistory)) {
    return user.loginHistory;
  }
  if (typeof user.loginHistory === 'string') {
    try {
      const parsed = JSON.parse(user.loginHistory);
      return Array.isArray(parsed) ? parsed : [];
    } catch (e) {
      return [];
    }
  }
  return [];
}

/**
 * Mencatat login pengguna ke sheet users dan login_history
 */
export async function recordUserLogin(
  userId: string,
  userEmail: string,
  userName?: string,
  spreadsheetId?: string | null,
  accessToken?: string | null
): Promise<LoginHistoryEntry | null> {
  if (!userId) return null;

  // Hindari pencatatan berulang dalam durasi singkat (misal reload halaman < 30 detik)
  const lastRecordedKey = `last_login_recorded_${userId}`;
  const lastRecordedTime = Number(sessionStorage.getItem(lastRecordedKey) || '0');
  const now = Date.now();
  if (now - lastRecordedTime < 30000) {
    return null; // Sudah dicatat baru saja dalam sesi ini
  }
  sessionStorage.setItem(lastRecordedKey, now.toString());

  try {
    const info = await getClientLoginInfo();

    const newEntry: LoginHistoryEntry = {
      id: `lh_${now}_${Math.random().toString(36).substring(2, 7)}`,
      timestamp: info.timestamp,
      dateTimeStr: info.dateTimeStr,
      ip: info.ip,
      location: info.location,
      device: info.device,
      city: info.city,
      region: info.region,
      country: info.country
    };

    // 1. Ambil data user dari cache local dan profile tersimpan
    let existingUser: any = {};
    let existingHistory: LoginHistoryEntry[] = [];
    
    // Cek dari stored profile terlebih dahulu
    const profileKey = `user_profile_${userId}`;
    let storedProfileParsed: any = null;
    try {
      const storedProfile = localStorage.getItem(profileKey);
      if (storedProfile) {
        storedProfileParsed = JSON.parse(storedProfile);
      }
    } catch (e) {}

    // Cek dari cache_users
    try {
      const cached = localStorage.getItem('cache_users');
      if (cached) {
        const users = JSON.parse(cached);
        const found = users.find((u: any) => 
          String(u.id || u.uid) === String(userId) || 
          (u.email && u.email.toLowerCase() === userEmail.toLowerCase())
        );
        if (found) {
          existingUser = found;
          existingHistory = parseLoginHistory(found);
        }
      }
    } catch (e) {}

    if (existingHistory.length === 0 && storedProfileParsed) {
      existingHistory = parseLoginHistory(storedProfileParsed);
    }

    // Batasi riwayat maksimal 20 data terakhir agar hemat ruang penyimpanan Google Sheets
    const updatedHistory = [newEntry, ...existingHistory.filter(h => h.id !== newEntry.id)].slice(0, 20);

    const effectiveSpreadsheetId = spreadsheetId || localStorage.getItem('app_spreadsheet_id') || (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
    const activeToken = accessToken || localStorage.getItem('app_access_token');

    // 2. Validasi status verifikasi dan kepemilikan akun agar TIDAK PERNAH terhapus / hilang saat login
    const isOwnerEmail = Boolean(userEmail && (
      userEmail.toLowerCase() === 'travelio11111@gmail.com' ||
      userEmail.toLowerCase().startsWith('admin')
    ));

    let isVerified = Boolean(isOwnerEmail);
    if (!isVerified) {
      const checkCandidates = [
        existingUser?.isVerified,
        storedProfileParsed?.isVerified
      ];
      for (const cand of checkCandidates) {
        if (cand === true || cand === 'true' || cand === 'TRUE') {
          isVerified = true;
          break;
        }
      }
    }

    const role = isOwnerEmail ? 'admin' : (existingUser?.role || storedProfileParsed?.role || 'pengurus');
    const location = existingUser?.location || storedProfileParsed?.location || '';
    const finalDisplayName = userName || existingUser?.displayName || storedProfileParsed?.displayName || userEmail.split('@')[0];
    const createdAt = existingUser?.createdAt || storedProfileParsed?.createdAt || Date.now();
    const finalId = existingUser?.id || existingUser?.uid || storedProfileParsed?.id || storedProfileParsed?.uid || userId;

    const updatePayload: UserProfile = {
      ...existingUser,
      uid: userId,
      id: finalId,
      email: userEmail,
      displayName: finalDisplayName,
      role,
      location,
      isVerified,
      createdAt,
      spreadsheetId: existingUser?.spreadsheetId || storedProfileParsed?.spreadsheetId || effectiveSpreadsheetId,
      lastLoginAt: info.timestamp,
      lastLoginFormatted: info.dateTimeStr,
      lastLoginIp: info.ip,
      lastLoginLocation: info.location,
      lastLoginDevice: info.device,
      loginHistory: updatedHistory
    };

    // Update profile di localStorage SEBELUM memicu sync apapun
    const mergedProfile = { ...(storedProfileParsed || {}), ...updatePayload, isVerified };
    localStorage.setItem(profileKey, JSON.stringify(mergedProfile));

    // Simpan ke users
    await saveData(null, activeToken, effectiveSpreadsheetId, 'users', updatePayload, finalId);

    // Simpan juga audit trail terpisah ke collection/sheet 'login_history'
    try {
      await saveData(null, activeToken, effectiveSpreadsheetId, 'login_history', {
        id: newEntry.id,
        userId,
        email: userEmail,
        displayName: finalDisplayName,
        timestamp: info.timestamp,
        dateTimeStr: info.dateTimeStr,
        ip: info.ip,
        location: info.location,
        device: info.device
      }, newEntry.id);
    } catch (auditErr) {
      console.warn('Simpan login_history audit notice:', auditErr);
    }

    console.info(`[LoginTracking] Login berhasil dicatat untuk ${userEmail}: IP ${info.ip}, Lokasi: ${info.location}, Waktu: ${info.dateTimeStr}`);
    return newEntry;
  } catch (err) {
    console.warn('[LoginTracking] Gagal mencatat history login:', err);
    return null;
  }
}
