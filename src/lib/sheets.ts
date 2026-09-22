/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

const GOOGLE_SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

export const normalizeAppsScriptUrl = (url: string | null | undefined): string | null => {
  if (!url || typeof url !== 'string') return null;
  let trimmed = url.trim();
  if (!trimmed) return null;
  if (!/^https?:\/\//i.test(trimmed)) {
    trimmed = `https://${trimmed}`;
  }
  if (trimmed.includes('googleusercontent.com') || trimmed.includes('/echo?')) {
    return null;
  }
  if (!trimmed.includes('script.google.com')) {
    return null;
  }
  // Convert /edit to /exec if user accidentally pasted script editor URL
  if (trimmed.includes('/edit')) {
    trimmed = trimmed.replace(/\/edit.*$/i, '/exec');
  }
  return trimmed;
};

export const getAppsScriptUrl = (): string | null => {
  const raw = (
    (import.meta as any).env?.VITE_APPS_SCRIPT_URL ||
    localStorage.getItem('app_script_url') ||
    null
  );

  const normalized = normalizeAppsScriptUrl(raw);
  if (!normalized && raw) {
    console.warn('[Sheets] URL Apps Script tidak valid atau bukan URL Web App yang benar (/exec):', raw);
  }
  return normalized;
};

export interface SheetData {
  range: string;
  values: any[][];
}

const handleResponseError = async (response: Response, actionName: string) => {
  const errorObj = await response.json().catch(() => ({}));
  const isAuthError = response.status === 401 || 
                      response.status === 403 || 
                      errorObj?.error?.code === 401 || 
                      errorObj?.error?.code === 403 || 
                      errorObj?.error?.status === 'UNAUTHENTICATED' || 
                      errorObj?.error?.status === 'PERMISSION_DENIED' || 
                      (typeof errorObj?.error?.message === 'string' && (
                        errorObj.error.message.includes('UNAUTHENTICATED') || 
                        errorObj.error.message.includes('PERMISSION_DENIED')
                      ));
  
  if (isAuthError) {
    console.warn(`[Sheets] Auth error (${response.status}) on ${actionName}`);
    localStorage.removeItem('app_access_token');
    window.dispatchEvent(new Event('storage'));
    throw new Error('Akses Google Sheets memerlukan autentikasi Google (403/401). Silakan masuk menggunakan Google.');
  }

  const msg = errorObj?.error?.message || JSON.stringify(errorObj);
  throw new Error(`Gagal pada ${actionName}: ${msg}`);
};

export const createSpreadsheet = async (accessToken: string, title: string) => {
  const response = await fetch(GOOGLE_SHEETS_API_BASE, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      properties: { title },
      sheets: [
        { properties: { title: 'jamaah' } },
        { properties: { title: 'assets' } },
        { properties: { title: 'activities' } },
        { properties: { title: 'ub_shopping' } },
        { properties: { title: 'attendance' } },
        { properties: { title: 'facility_stats' } },
        { properties: { title: 'users' } },
      ],
    }),
  });

  if (!response.ok) {
    await handleResponseError(response, 'membuat spreadsheet');
  }

  return await response.json();
};

export const updateSheetValues = async (accessToken: string, spreadsheetId: string, range: string, values: any[][]) => {
  clearSheetMemoryCache(range.split('!')[0]);
  const response = await fetch(`${GOOGLE_SHEETS_API_BASE}/${spreadsheetId}/values/${range}?valueInputOption=RAW`, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ 
      range,
      majorDimension: 'ROWS',
      values 
    }),
  });

  if (!response.ok) {
    await handleResponseError(response, 'memperbarui sheet');
  }

  return await response.json();
};

// In-memory cache to prevent duplicate fetch requests and provide instant 0ms tab switching
const memoryCache = new Map<string, { data: SheetData; expiresAt: number }>();
const CACHE_TTL_MS = 3000; // 3 seconds cache for fast fresh updates

export const clearSheetMemoryCache = (collectionName?: string) => {
  if (!collectionName) {
    memoryCache.clear();
    return;
  }
  for (const key of memoryCache.keys()) {
    if (key.includes(collectionName)) {
      memoryCache.delete(key);
    }
  }
};

export function parseSheetRowsToObjects<T = any>(rows: any[][]): T[] {
  if (!rows || rows.length === 0) return [];
  const headers = rows[0] || [];
  return rows.slice(1).map(row => {
    const item: any = {};
    headers.forEach((header, index) => {
      let val = row[index];
      try {
        if (typeof val === 'string') {
          if (val.startsWith('{') || val.startsWith('[')) {
            val = JSON.parse(val);
          } else if (val.toLowerCase() === 'true') {
            val = true;
          } else if (val.toLowerCase() === 'false') {
            val = false;
          } else if (!isNaN(Number(val)) && val !== '' && !val.startsWith('0')) {
            val = Number(val);
          }
        }
      } catch (e) {}
      if (header === 'imageUrls' || header === 'images') {
        if (!Array.isArray(val)) {
          if (typeof val === 'string' && val.trim()) {
            if (val.trim().startsWith('[') && val.trim().endsWith(']')) {
              try { val = JSON.parse(val.trim()); } catch (e) {}
            } else if (val.includes(',')) {
              val = val.split(',').map((s: string) => s.trim()).filter(Boolean);
            } else {
              val = [val.trim()];
            }
          } else {
            val = [];
          }
        }
      }
      item[header] = val;
    });
    return item as T;
  });
}

// In-flight request deduplication map to prevent multiple simultaneous requests for the same sheet
const inFlightRequests = new Map<string, Promise<SheetData>>();

// Concurrency queue to avoid overwhelming Google Apps Script's concurrent execution limit
let activeAppsScriptRequests = 0;
const MAX_CONCURRENT_REQUESTS = 2;
const appsScriptQueue: Array<() => void> = [];

const queueAppsScriptRequest = <T>(task: () => Promise<T>): Promise<T> => {
  return new Promise<T>((resolve, reject) => {
    const execute = () => {
      activeAppsScriptRequests++;
      task()
        .then(resolve)
        .catch(reject)
        .finally(() => {
          activeAppsScriptRequests--;
          if (appsScriptQueue.length > 0) {
            const next = appsScriptQueue.shift();
            next?.();
          }
        });
    };

    if (activeAppsScriptRequests < MAX_CONCURRENT_REQUESTS) {
      execute();
    } else {
      appsScriptQueue.push(execute);
    }
  });
};

// Batch All-in-One Fetch State
let batchFetchPromise: Promise<boolean> | null = null;
let lastBatchFetchTime = 0;
const BATCH_CACHE_COOLDOWN_MS = 6000; // 6 seconds cooldown between batch fetches

/**
 * Fetch ALL sheets simultaneously in a single HTTP request using action=getAll.
 * Dramatically cuts load time by triggering Google Apps Script cold-start only once.
 */
export const fetchAllSheetsBatch = async (
  spreadsheetId?: string | null,
  force: boolean = false
): Promise<boolean> => {
  const appsScriptUrl = getAppsScriptUrl();
  if (!appsScriptUrl) return false;

  const now = Date.now();
  if (!force && now - lastBatchFetchTime < BATCH_CACHE_COOLDOWN_MS) {
    return true;
  }

  if (batchFetchPromise) {
    return batchFetchPromise;
  }

  batchFetchPromise = (async () => {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => {
        try { controller.abort('Apps Script timeout'); } catch {}
      }, 45000);

      const sep = appsScriptUrl.includes('?') ? '&' : '?';
      const fetchUrl = `${appsScriptUrl}${sep}action=getAll&_t=${Date.now()}`;

      const res = await fetch(fetchUrl, {
        method: 'GET',
        redirect: 'follow',
        cache: 'no-store',
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      if (!res.ok) return false;

      const contentType = (res.headers.get('content-type') || '').toLowerCase();
      if (!contentType.includes('application/json') && !contentType.includes('+json')) {
        return false;
      }

      const json = await res.json();
      if (json && json.status === 'success' && json.data && typeof json.data === 'object') {
        const dataMap = json.data as Record<string, any[][]>;
        const curTime = Date.now();
        lastBatchFetchTime = curTime;

        Object.keys(dataMap).forEach((collection) => {
          const rawValues = dataMap[collection] || [];
          const cacheKey = `${spreadsheetId || 'default'}_${collection}`;
          memoryCache.set(cacheKey, {
            data: { range: `${collection}!A:ZZ`, values: rawValues },
            expiresAt: curTime + CACHE_TTL_MS
          });

          // Also populate localStorage cache so it's instantly usable
          try {
            const items = parseSheetRowsToObjects(rawValues);
            if (items.length > 0) {
              localStorage.setItem(`cache_${collection}`, JSON.stringify(items));
              window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName: collection } }));
            }
          } catch (e) {}
        });

        console.info('[Sheets] Batch loading all-in-one berhasil! Seluruh data tabel dimuat dalam 1 request.');
        return true;
      }
      return false;
    } catch (err) {
      return false;
    } finally {
      batchFetchPromise = null;
    }
  })();

  return batchFetchPromise;
};

/**
 * Fetch values EXCLUSIVELY via Google Apps Script Web App.
 * Uses Batch All-in-One when possible, with individual fallback.
 */
export const getSheetValues = async (
  _accessToken: string | null | undefined,
  spreadsheetId: string | null | undefined,
  range: string
): Promise<SheetData> => {
  const sheetName = range.split('!')[0] || 'Sheet1';
  const cacheKey = `${spreadsheetId || 'default'}_${sheetName}`;

  // 1. Instant return from in-memory cache (0ms)
  const cached = memoryCache.get(cacheKey);
  const now = Date.now();
  if (cached && cached.expiresAt > now) {
    return cached.data;
  }

  // 2. Return active in-flight promise if the same sheet is already being fetched
  const inFlight = inFlightRequests.get(cacheKey);
  if (inFlight) {
    return inFlight;
  }

  const appsScriptUrl = getAppsScriptUrl();
  if (!appsScriptUrl) {
    console.warn('[Sheets] URL Apps Script belum dikonfigurasi. Silakan atur URL Web App Apps Script di menu pengaturan.');
    return { range, values: [] };
  }

  // 3. Attempt Batch All-in-One Fetch first to load all sheets in a single round-trip
  if (now - lastBatchFetchTime > BATCH_CACHE_COOLDOWN_MS) {
    await fetchAllSheetsBatch(spreadsheetId);
    const batchCached = memoryCache.get(cacheKey);
    if (batchCached && batchCached.expiresAt > Date.now()) {
      return batchCached.data;
    }
  }

  const fetchPromise = (async (): Promise<SheetData> => {
    const executeFetch = async (attempt: number = 1): Promise<SheetData> => {
      return queueAppsScriptRequest(async () => {
        const controller = new AbortController();
        // 45 seconds timeout for cold start tolerance
        const timeoutId = setTimeout(() => {
          try {
            controller.abort('Apps Script timeout');
          } catch {
            controller.abort();
          }
        }, 45000);

        try {
          const sep = appsScriptUrl.includes('?') ? '&' : '?';
          // Send collection, sheet, and action for maximum compatibility with all Apps Script doGet variations
          const fetchUrl = `${appsScriptUrl}${sep}action=get&collection=${encodeURIComponent(sheetName)}&sheet=${encodeURIComponent(sheetName)}&_t=${Date.now()}`;
          
          const res = await fetch(fetchUrl, {
            method: 'GET',
            redirect: 'follow',
            cache: 'no-store',
            signal: controller.signal
          });

          if (!res.ok) {
            throw new Error(`HTTP ${res.status}: ${res.statusText}`);
          }

          const contentType = (res.headers.get('content-type') || '').toLowerCase();
          if (!contentType.includes('application/json') && !contentType.includes('+json')) {
            const text = await res.text();
            if (text.includes('<html') || text.includes('<!DOCTYPE')) {
              console.warn(
                '[Sheets] Apps Script mengembalikan halaman HTML. Pastikan Web App di-deploy dengan akses: "Anyone" (Siapa saja) dan URL berakhiran /exec.'
              );
            }
            throw new Error('Apps Script tidak mengembalikan format JSON');
          }

          const json = await res.json();
          if (json && Array.isArray(json.values)) {
            const result: SheetData = { range, values: json.values };
            memoryCache.set(cacheKey, { data: result, expiresAt: Date.now() + CACHE_TTL_MS });
            return result;
          }

          throw new Error('Respons Apps Script tidak memiliki array values');
        } catch (err: any) {
          const isTimeout = err?.name === 'AbortError' || String(err?.message || '').toLowerCase().includes('timeout') || String(err?.message || '').toLowerCase().includes('aborted');
          
          // Retry once on timeout/transient error with a brief cooldown
          if (attempt === 1) {
            await new Promise((r) => setTimeout(r, 1200));
            return executeFetch(2);
          }

          console.warn(`[Sheets] Gagal mengambil sheet '${sheetName}' via Apps Script (${isTimeout ? 'Waktu habis / Timeout' : err?.message || err}). Menggunakan data lokal.`);
          return { range, values: [] };
        } finally {
          clearTimeout(timeoutId);
        }
      });
    };

    try {
      return await executeFetch(1);
    } finally {
      inFlightRequests.delete(cacheKey);
    }
  })();

  inFlightRequests.set(cacheKey, fetchPromise);
  return fetchPromise;
};

export const appendSheetValues = async (accessToken: string, spreadsheetId: string, range: string, values: any[][]) => {
  clearSheetMemoryCache(range.split('!')[0]);
  const response = await fetch(`${GOOGLE_SHEETS_API_BASE}/${spreadsheetId}/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ 
      range,
      majorDimension: 'ROWS',
      values 
    }),
  });

  if (!response.ok) {
    await handleResponseError(response, 'menambah baris sheet');
  }

  return await response.json();
};

export const clearSheetValues = async (accessToken: string, spreadsheetId: string, range: string) => {
  clearSheetMemoryCache(range.split('!')[0]);
  const response = await fetch(`${GOOGLE_SHEETS_API_BASE}/${spreadsheetId}/values/${range}:clear`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  if (!response.ok) {
    await handleResponseError(response, 'menghapus isi sheet');
  }

  return await response.json();
};