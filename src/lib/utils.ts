import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

/**
 * Normalizes any date format (DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD, ISO string,
 * Indonesian month format, Excel serial numbers, timestamps) into standard HTML5 "YYYY-MM-DD".
 */
export function normalizeDateToInputFormat(val: any): string {
  if (val === null || val === undefined || val === '') return '';

  // 1. Numeric handling (Excel serial dates or Unix timestamps)
  if (typeof val === 'number' && !isNaN(val)) {
    if (val > 1000 && val < 100000) {
      // Excel serial date (e.g., 34865 = 1995-06-15)
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      if (!isNaN(date.getTime())) {
        const y = date.getUTCFullYear();
        const m = String(date.getUTCMonth() + 1).padStart(2, '0');
        const d = String(date.getUTCDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
    }
    const ms = val > 10000000000 ? val : val * 1000;
    const date = new Date(ms);
    if (!isNaN(date.getTime())) {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    return '';
  }

  const str = String(val).trim();
  if (!str || str === '-' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined') return '';

  // 2. YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD (optionally followed by T or space and time)
  const ymdMatch = str.match(/^(\d{4})[-/. ](\d{1,2})[-/. ](\d{1,2})/);
  if (ymdMatch) {
    const year = ymdMatch[1];
    const month = ymdMatch[2].padStart(2, '0');
    const day = ymdMatch[3].padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  // 3. DD-MM-YYYY or DD/MM/YYYY or DD.MM.YYYY
  const dmyMatch = str.match(/^(\d{1,2})[-/. ](\d{1,2})[-/. ](\d{2,4})/);
  if (dmyMatch) {
    const day = dmyMatch[1].padStart(2, '0');
    const month = dmyMatch[2].padStart(2, '0');
    let year = dmyMatch[3];
    if (year.length === 2) {
      const yrNum = parseInt(year, 10);
      year = String(yrNum < 40 ? 2000 + yrNum : 1900 + yrNum);
    }
    return `${year}-${month}-${day}`;
  }

  // 4. Indonesian month names (e.g., "17 Agustus 1990", "5 Jan 2001", "25-Des-1995")
  const idMonths: Record<string, string> = {
    jan: '01', januari: '01',
    feb: '02', februari: '02',
    mar: '03', maret: '03',
    apr: '04', april: '04',
    mei: '05', may: '05',
    jun: '06', juni: '06',
    jul: '07', juli: '07',
    agu: '08', agt: '08', agustus: '08', aug: '08', august: '08',
    sep: '09', september: '09',
    okt: '10', oktober: '10', oct: '10', october: '10',
    nov: '11', november: '11',
    des: '12', desember: '12', dec: '12', december: '12',
  };

  const textMonthMatch = str.match(/^(\d{1,2})[\s\-_/]+([a-zA-Z]+)[\s\-_/]+(\d{2,4})/);
  if (textMonthMatch) {
    const day = textMonthMatch[1].padStart(2, '0');
    const mName = textMonthMatch[2].toLowerCase();
    let year = textMonthMatch[3];
    if (year.length === 2) {
      const yrNum = parseInt(year, 10);
      year = String(yrNum < 40 ? 2000 + yrNum : 1900 + yrNum);
    }
    const month = idMonths[mName];
    if (month) {
      return `${year}-${month}-${day}`;
    }
  }

  // 5. Native JS Date parsing fallback
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  return '';
}

/**
 * Extracts and normalizes birth date from any common property aliases on a Jamaah object.
 */
export function extractBirthDate(j?: any): string {
  if (!j) return '';
  const raw =
    j.dateOfBirth ??
    j.birthDate ??
    j.tanggalLahir ??
    j.tglLahir ??
    j.tgl_lahir ??
    j.tanggal_lahir ??
    j.dob ??
    j['Tanggal Lahir'] ??
    j['Date of Birth'] ??
    '';
  return normalizeDateToInputFormat(raw);
}
