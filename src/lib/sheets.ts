/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

const GOOGLE_SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
const FALLBACK_API_KEY = 'AIzaSyCBRAeVjFpslb7v4ck_VEehjjWjBt49LUI';

export interface SheetData {
  range: string;
  values: any[][];
}

const handleResponseError = async (response: Response, actionName: string) => {
  const errorObj = await response.json().catch(() => ({}));
  const is401 = response.status === 401 || 
                errorObj?.error?.code === 401 || 
                errorObj?.error?.status === 'UNAUTHENTICATED' || 
                (typeof errorObj?.error?.message === 'string' && errorObj.error.message.includes('UNAUTHENTICATED'));
  
  if (is401) {
    console.warn(`[Sheets] 401 Unauthenticated on ${actionName}`);
    throw new Error('Sesi login Google Sheets Anda telah berakhir atau belum terhubung (401).');
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
const CACHE_TTL_MS = 20000; // 20 seconds cache

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

/**
 * Fast GViz fetch for public spreadsheets (sub-150ms directly from Google CDN)
 */
const fetchViaGViz = async (spreadsheetId: string, sheetName: string): Promise<SheetData | null> => {
  try {
    const gvizUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}`;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 3500);

    const gvizResponse = await fetch(gvizUrl, { signal: controller.signal });
    clearTimeout(timeoutId);

    if (gvizResponse.ok) {
      const text = await gvizResponse.text();
      const jsonStart = text.indexOf('{');
      const jsonEnd = text.lastIndexOf('}');
      if (jsonStart !== -1 && jsonEnd !== -1) {
        const parsed = JSON.parse(text.substring(jsonStart, jsonEnd + 1));
        if (parsed.status === 'error') return null;

        const cols = parsed.table?.cols || [];
        const rows = parsed.table?.rows || [];
        const headers = cols.map((c: any) => c.label || c.id || '');
        const values = [
          headers,
          ...rows.map((r: any) => (r.c || []).map((cell: any) => cell?.v ?? ''))
        ];
        return { range: `${sheetName}!A:ZZ`, values };
      }
    }
  } catch (err) {
    // Non-fatal, will fall back
  }
  return null;
};

/**
 * Fetch values with intelligent fast strategy:
 * 1. Return from memory cache if fresh (0ms)
 * 2. If spreadsheet is open public, use GViz directly (sub-150ms, no token needed)
 * 3. Fallback to OAuth token, Apps Script Web App, or API Key if private
 */
export const getSheetValues = async (accessToken: string | null | undefined, spreadsheetId: string, range: string): Promise<SheetData> => {
  if (!spreadsheetId) {
    return { range, values: [] };
  }

  const sheetName = range.split('!')[0] || 'Sheet1';
  const cacheKey = `${spreadsheetId}_${sheetName}`;

  // 1. Instant return from in-memory cache
  const cached = memoryCache.get(cacheKey);
  const now = Date.now();
  if (cached && cached.expiresAt > now) {
    return cached.data;
  }

  // 2. Try with OAuth Access Token first if provided (fast direct authorized access to Google Sheets API)
  if (accessToken) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 4000);

      const response = await fetch(`${GOOGLE_SHEETS_API_BASE}/${spreadsheetId}/values/${range}`, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      if (response.ok) {
        const json = await response.json();
        memoryCache.set(cacheKey, { data: json, expiresAt: now + CACHE_TTL_MS });
        return json;
      }

      if (response.status === 401 || response.status === 403) {
        // Token is invalid/expired - clear it so subsequent calls don't waste time
        localStorage.removeItem('app_access_token');
      }
    } catch (err) {
      console.warn('[Sheets] Token request network error:', err);
    }
  }

  // 3. Fallback: Ultra-fast check via GViz (works instantly for public sheets)
  const gvizResult = await fetchViaGViz(spreadsheetId, sheetName);
  if (gvizResult && gvizResult.values && gvizResult.values.length > 0) {
    memoryCache.set(cacheKey, { data: gvizResult, expiresAt: now + CACHE_TTL_MS });
    return gvizResult;
  }

  // 4. Try with Apps Script Web App URL (Token-Free read fallback)
  const appsScriptUrl = (import.meta as any).env?.VITE_APPS_SCRIPT_URL || localStorage.getItem('app_script_url');
  if (appsScriptUrl) {
    try {
      const fetchUrl = `${appsScriptUrl}${appsScriptUrl.includes('?') ? '&' : '?'}collection=${encodeURIComponent(sheetName)}`;
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 4000);

      const res = await fetch(fetchUrl, { signal: controller.signal });
      clearTimeout(timeoutId);

      if (res.ok) {
        const json = await res.json();
        if (json && Array.isArray(json.values)) {
          const result = { range, values: json.values };
          memoryCache.set(cacheKey, { data: result, expiresAt: now + CACHE_TTL_MS });
          return result;
        }
      }
    } catch (err) {
      console.warn('[Sheets] Apps Script GET fetch error:', err);
    }
  }

  // 5. Try with API Key (works if sheet is shared as "Anyone with link can view")
  const apiKey = (import.meta as any).env?.VITE_GOOGLE_API_KEY || FALLBACK_API_KEY;
  if (apiKey) {
    try {
      const keyResponse = await fetch(`${GOOGLE_SHEETS_API_BASE}/${spreadsheetId}/values/${range}?key=${apiKey}`);
      if (keyResponse.ok) {
        const json = await keyResponse.json();
        memoryCache.set(cacheKey, { data: json, expiresAt: now + CACHE_TTL_MS });
        return json;
      }
    } catch (err) {
      console.warn('[Sheets] API key fetch error:', err);
    }
  }

  // If all fail, return empty structure rather than throwing an unhandled rejection
  return { range, values: [] };
};

export const appendSheetValues = async (accessToken: string, spreadsheetId: string, range: string, values: any[][]) => {
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
