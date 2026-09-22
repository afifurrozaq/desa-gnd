import { getSheetValues, updateSheetValues, appendSheetValues, clearSheetValues, createSpreadsheet, clearSheetMemoryCache, getAppsScriptUrl } from './sheets';
export { clearSheetMemoryCache, getAppsScriptUrl };

export async function getOrRefreshAccessToken(providedToken?: string | null): Promise<string | null> {
  const token = providedToken || localStorage.getItem('app_access_token');
  return token || null;
}

export async function saveData<T>(
  db: any, // Unused - Firestore disabled
  accessToken: string | null,
  spreadsheetId: string | null | undefined,
  collectionName: string,
  data: any,
  id?: string
) {
  const envSpreadsheetId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
  let effectiveSpreadsheetId = spreadsheetId || envSpreadsheetId || localStorage.getItem('app_spreadsheet_id') || '';

  const finalId = id || data.id || data.uid || Math.random().toString(36).substring(2, 11);

  // 1. Merge with existing cached item to safeguard against data erasure on partial updates (e.g. isConfirmed)
  let existingCachedItem: any = {};
  try {
    const cached = localStorage.getItem(`cache_${collectionName}`);
    const items = cached ? JSON.parse(cached) : [];
    const found = items.find((it: any) => String(it.id || it.uid) === String(finalId));
    if (found) {
      existingCachedItem = found;
    }
  } catch (err) {
    console.warn('[Cache] Could not read local cache for merge:', err);
  }

  const itemToSave = {
    ...existingCachedItem,
    ...data,
    id: finalId,
    _localAddedAt: existingCachedItem._localAddedAt || Date.now()
  };

  // 2. Immediately update local persistent cache and clear memory cache to prevent ANY data loss
  clearSheetMemoryCache(collectionName);
  try {
    const cached = localStorage.getItem(`cache_${collectionName}`);
    const items = cached ? JSON.parse(cached) : [];
    const existingIndex = items.findIndex((it: any) => String(it.id || it.uid) === String(finalId));
    if (existingIndex !== -1) {
      items[existingIndex] = { ...items[existingIndex], ...itemToSave };
    } else {
      items.push(itemToSave);
    }
    localStorage.setItem(`cache_${collectionName}`, JSON.stringify(items));
    window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
  } catch (err) {
    console.warn('[Cache] Could not write to local cache:', err);
  }

  // Sanitize item properties for Google Sheets
  const cleanItem = { ...itemToSave };
  Object.keys(cleanItem).forEach(key => {
    if (cleanItem[key] && typeof cleanItem[key] === 'object' && typeof cleanItem[key].toDate === 'function') {
      cleanItem[key] = cleanItem[key].toDate().toISOString();
    } else if (cleanItem[key] instanceof Date) {
      cleanItem[key] = cleanItem[key].toISOString();
    } else if (typeof cleanItem[key] === 'object' && cleanItem[key] !== null) {
      cleanItem[key] = JSON.stringify(cleanItem[key]);
    }
  });

  // 3. If Apps Script Web App URL is configured: Save directly without OAuth token!
  const appsScriptUrl = getAppsScriptUrl();
  if (appsScriptUrl) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 6000);
      await fetch(appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'save',
          collection: collectionName,
          item: cleanItem,
          id: finalId,
          spreadsheetId: effectiveSpreadsheetId
        }),
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      console.info(`[Sheets] Data '${collectionName}' berhasil disimpan via Apps Script (Bebas Token)!`);
      window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
      return finalId;
    } catch (scriptErr) {
      console.warn('[Sheets] Gagal kirim ke Apps Script URL / Timeout, fallback ke cache lokal:', scriptErr);
      return finalId;
    }
  }

  // 4. Try to obtain active Google Access Token
  let effectiveAccessToken = await getOrRefreshAccessToken(accessToken);

  if (!effectiveAccessToken) {
    console.info(`[Sheets] Token tidak tersedia. Data '${collectionName}' telah disimpan ke cache lokal.`);
    return finalId;
  }

  const executeSave = async (token: string) => {
    let currentSpreadsheetId = effectiveSpreadsheetId;
    if (!currentSpreadsheetId) {
      const created = await createSpreadsheet(token, 'Database Mosque Management');
      currentSpreadsheetId = created.spreadsheetId;
      if (currentSpreadsheetId) {
        effectiveSpreadsheetId = currentSpreadsheetId;
        localStorage.setItem('app_spreadsheet_id', currentSpreadsheetId);
        window.dispatchEvent(new Event('spreadsheet_id_updated'));
      }
    }

    let result: any = { values: [] };
    try {
      result = await getSheetValues(token, currentSpreadsheetId!, `${collectionName}!A:ZZ`);
    } catch (e: any) {
      result = { values: [] };
    }
    const rows = (result.values as any[][]) || [];
    let headers = rows[0] || [];

    if (rows.length === 0) {
      const newHeaders = Object.keys(cleanItem);
      const firstRow = newHeaders.map(h => cleanItem[h] ?? '');
      await updateSheetValues(token, currentSpreadsheetId!, `${collectionName}!A1`, [newHeaders, firstRow]);
      window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
      return finalId;
    }

    const itemKeys = Object.keys(cleanItem);
    const missingHeaders = itemKeys.filter(k => !headers.includes(k));
    if (missingHeaders.length > 0) {
      headers = [...headers, ...missingHeaders];
      await updateSheetValues(token, currentSpreadsheetId!, `${collectionName}!A1`, [headers]);
    }

    if (finalId && rows.length > 0) {
      let idIndex = headers.indexOf('id');
      if (idIndex === -1) idIndex = headers.indexOf('uid');
      if (idIndex !== -1) {
        const rowIndex = rows.findIndex((r, idx) => idx > 0 && String(r[idIndex]) === String(finalId));
        if (rowIndex !== -1) {
          const existingRowValues = rows[rowIndex] || [];
          const rowValues = headers.map((h, colIdx) => {
            const val = cleanItem[h];
            if (val !== undefined && val !== null && val !== '') {
              return val;
            }
            return existingRowValues[colIdx] !== undefined ? existingRowValues[colIdx] : '';
          });
          await updateSheetValues(token, currentSpreadsheetId!, `${collectionName}!A${rowIndex + 1}`, [rowValues]);
          window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
          return finalId;
        }
      }
    }

    const rowValues = headers.map(h => cleanItem[h] ?? '');
    await appendSheetValues(token, currentSpreadsheetId!, collectionName, [rowValues]);

    window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
    return finalId;
  };

  try {
    return await executeSave(effectiveAccessToken);
  } catch (error: any) {
    const errorMsg = error.message || String(error);
    const is401 = errorMsg.includes('401') || errorMsg.includes('UNAUTHENTICATED') || errorMsg.includes('invalid authentication credentials');

    if (is401) {
      window.dispatchEvent(new CustomEvent('google_token_expired'));
    }
    console.warn(`[Sheets] Gagal sync ke spreadsheet (${errorMsg}), data aman tersimpan di cache lokal.`);
    return finalId;
  }
}

export async function deleteData(
  db: any, // Unused - Firestore disabled
  accessToken: string | null,
  spreadsheetId: string | null | undefined,
  collectionName: string,
  id: string
) {
  const envSpreadsheetId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
  const effectiveSpreadsheetId = spreadsheetId || envSpreadsheetId || localStorage.getItem('app_spreadsheet_id');

  // 1. Immediately delete from local cache and clear memory cache
  clearSheetMemoryCache(collectionName);
  try {
    let recentDeletes: Array<{ id: string; time: number }> = [];
    const raw = localStorage.getItem(`recent_deletes_${collectionName}`);
    if (raw) recentDeletes = JSON.parse(raw);
    recentDeletes.push({ id: String(id), time: Date.now() });
    localStorage.setItem(`recent_deletes_${collectionName}`, JSON.stringify(recentDeletes));

    const cached = localStorage.getItem(`cache_${collectionName}`);
    if (cached) {
      const items = JSON.parse(cached);
      const remaining = items.filter((it: any) => String(it.id || it.uid) !== String(id));
      localStorage.setItem(`cache_${collectionName}`, JSON.stringify(remaining));
      window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
    }
  } catch (err) {
    console.warn('[Cache] Could not delete from local cache:', err);
  }

  // 2. If Apps Script Web App URL is configured: Delete directly without OAuth token!
  const appsScriptUrl = getAppsScriptUrl();
  if (appsScriptUrl) {
    try {
      await fetch(appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'delete',
          collection: collectionName,
          id: id,
          spreadsheetId: effectiveSpreadsheetId
        }),
      });
      console.info(`[Sheets] Data '${id}' berhasil dihapus via Apps Script (Bebas Token)!`);
      window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
      return;
    } catch (scriptErr) {
      console.warn('[Sheets] Gagal delete via Apps Script URL:', scriptErr);
    }
  }

  let effectiveAccessToken = await getOrRefreshAccessToken(accessToken);

  if (!effectiveSpreadsheetId || !effectiveAccessToken) {
    console.info(`[Sheets] Token/Spreadsheet tidak tersedia saat delete. Data '${id}' telah dihapus dari cache lokal.`);
    return;
  }

  const executeDelete = async (token: string) => {
    const result = await getSheetValues(token, effectiveSpreadsheetId, `${collectionName}!A:ZZ`);
    const rows = (result.values as any[][]) || [];
    const headers = rows[0] || [];
    let idIndex = headers.indexOf('id');
    if (idIndex === -1) idIndex = headers.indexOf('uid');

    if (idIndex !== -1) {
      const rowIndex = rows.findIndex((r, idx) => idx > 0 && String(r[idIndex]) === String(id));
      if (rowIndex !== -1) {
        const newRows = rows.filter((r, i) => i !== rowIndex);
        // Clear old sheet contents so trailing deleted row doesn't persist
        await clearSheetValues(token, effectiveSpreadsheetId, `${collectionName}!A:ZZ`);
        if (newRows.length > 0) {
          await updateSheetValues(token, effectiveSpreadsheetId, `${collectionName}!A1`, newRows);
        }
        window.dispatchEvent(new CustomEvent('data_updated', { detail: { collectionName } }));
        return;
      }
    }
  };

  try {
    return await executeDelete(effectiveAccessToken);
  } catch (error: any) {
    const errorMsg = error.message || String(error);
    const is401 = errorMsg.includes('401') || errorMsg.includes('UNAUTHENTICATED') || errorMsg.includes('invalid authentication credentials');

    if (is401) {
      window.dispatchEvent(new CustomEvent('google_token_expired'));
    }
    console.warn(`[Sheets] Gagal delete di spreadsheet (${errorMsg}), data telah dihapus di lokal.`);
  }
}