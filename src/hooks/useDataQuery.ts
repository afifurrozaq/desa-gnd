import { useState, useEffect, useMemo, useCallback } from 'react';
import { getSheetValues, clearSheetMemoryCache, getAppsScriptUrl } from '../lib/sheets';
import { useFirebase } from '../components/FirebaseProvider';

export interface QueryWhereConstraint {
  field: string;
  op: '==' | '>=' | '<=' | '<' | '>' | '!=' | 'in';
  value: any;
}

export function where(field: string, op: '==' | '>=' | '<=' | '<' | '>' | '!=' | 'in', value: any): QueryWhereConstraint {
  return { field, op, value };
}

function filterItemsByConstraints<T>(items: T[], constraints: QueryWhereConstraint[]): T[] {
  if (!constraints || constraints.length === 0) return items;

  return items.filter(item => {
    return constraints.every(c => {
      if (!c || !c.field) return true;
      // If constraint value is undefined/empty or 'Seluruh Lokasi', don't filter out all items
      if (c.field === 'location' && (!c.value || c.value === 'Seluruh Lokasi')) return true;
      
      const val = (item as any)[c.field];
      if (c.op === '==') {
        if (typeof val === 'boolean' || typeof c.value === 'boolean') {
          return Boolean(val) === Boolean(c.value);
        }
        return String(val ?? '').trim().toLowerCase() === String(c.value ?? '').trim().toLowerCase();
      }
      if (c.op === '!=') {
        return String(val ?? '').trim().toLowerCase() !== String(c.value ?? '').trim().toLowerCase();
      }
      if (c.op === '>=') return val >= c.value;
      if (c.op === '<') return val < c.value;
      if (c.op === '<=') return val <= c.value;
      if (c.op === '>') return val > c.value;
      if (c.op === 'in') return Array.isArray(c.value) && c.value.includes(val);
      return true;
    });
  });
}

export interface UseDataQueryOptions {
  autoRefresh?: boolean;
  pollIntervalMs?: number;
}

export function useDataQuery<T>(
  collectionName: string,
  constraints: QueryWhereConstraint[] = [],
  overriddenSpreadsheetId?: string,
  options: UseDataQueryOptions = { autoRefresh: true, pollIntervalMs: 0 }
) {
  const { accessToken, spreadsheetId: globalSpreadsheetId } = useFirebase();
  const envSpreadsheetId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
  const effectiveSpreadsheetId = overriddenSpreadsheetId || globalSpreadsheetId || envSpreadsheetId || localStorage.getItem('app_spreadsheet_id') || '';

  const constraintsStr = useMemo(() => JSON.stringify(constraints || []), [constraints]);

  // Read current local cache
  const getCachedItems = useCallback((): T[] => {
    try {
      const cached = localStorage.getItem(`cache_${collectionName}`);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed)) {
          const sanitized = parsed.map((item: any) => {
            if (item && typeof item === 'object') {
              if (item.imageUrls && !Array.isArray(item.imageUrls)) {
                if (typeof item.imageUrls === 'string' && item.imageUrls.trim()) {
                  item.imageUrls = [item.imageUrls.trim()];
                } else {
                  item.imageUrls = [];
                }
              }
            }
            return item;
          });
          return filterItemsByConstraints<T>(sanitized, constraints);
        }
      }
    } catch (e) {}
    return [];
  }, [collectionName, constraintsStr]);

  const [sheetData, setSheetData] = useState<T[]>(() => getCachedItems());
  const [sheetLoading, setSheetLoading] = useState<boolean>(false);
  const [sheetError, setSheetError] = useState<Error | null>(null);

  const fetchFromSheets = useCallback(async (bypassCache = false) => {
    if (bypassCache) {
      clearSheetMemoryCache(collectionName);
    }

    const cachedItems = getCachedItems();
    if (cachedItems.length > 0) {
      setSheetData(cachedItems);
    }

    const activeSpreadsheetId = overriddenSpreadsheetId || globalSpreadsheetId || (import.meta as any).env?.VITE_SPREADSHEET_ID || localStorage.getItem('app_spreadsheet_id') || '';
    const activeAccessToken = accessToken || localStorage.getItem('app_access_token') || null;
    const appsScriptUrl = getAppsScriptUrl();

    if (!activeSpreadsheetId && !appsScriptUrl) {
      setSheetLoading(false);
      return;
    }
    
    if (cachedItems.length === 0) {
      setSheetLoading(true);
    }

    try {
      const result = await getSheetValues(activeAccessToken, activeSpreadsheetId, `${collectionName}!A:ZZ`);
      const rows = result.values as any[][];
      
      let items: T[] = [];
      if (rows && rows.length > 0) {
        const headers = rows[0] || [];
        items = rows.slice(1).map(row => {
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

      // Smart merge: protect recent local additions/deletions against stale remote sheets responses
      const now = Date.now();
      let recentDeletes: Array<{ id: string; time: number }> = [];
      try {
        const rawDeletes = localStorage.getItem(`recent_deletes_${collectionName}`);
        if (rawDeletes) recentDeletes = JSON.parse(rawDeletes);
      } catch (e) {}
      recentDeletes = recentDeletes.filter(d => now - d.time < 180000); // 3 min memory
      const deletedIds = new Set(recentDeletes.map(d => String(d.id)));

      // Filter out deleted items
      items = items.filter((it: any) => !deletedIds.has(String(it.id || it.uid)));

      // Merge local items that exist in current cache but haven't appeared in remote items yet
      const remoteMap = new Map<string, any>();
      items.forEach((it: any) => {
        const key = String(it.id || it.uid);
        if (key) remoteMap.set(key, it);
      });

      const currentLocalCache: any[] = JSON.parse(localStorage.getItem(`cache_${collectionName}`) || '[]');
      currentLocalCache.forEach(localItem => {
        const key = String(localItem.id || localItem.uid);
        const isRecentlyAddedLocally = localItem._localAddedAt && (now - Number(localItem._localAddedAt) < 120000);
        if (key && !deletedIds.has(key) && !remoteMap.has(key) && isRecentlyAddedLocally) {
          // Keep local item only if added locally very recently and pending remote write
          items.push(localItem);
        }
      });

      // Update local persistent cache with merged items
      try {
        localStorage.setItem(`cache_${collectionName}`, JSON.stringify(items));
      } catch (cacheErr) {}

      const filteredItems = filterItemsByConstraints<T>(items, constraints);
      setSheetData(filteredItems);
      setSheetError(null);
    } catch (err: any) {
      setSheetError(err);
      const fallback = getCachedItems();
      if (fallback.length > 0) {
        setSheetData(fallback);
      }
    } finally {
      setSheetLoading(false);
    }
  }, [collectionName, constraintsStr, effectiveSpreadsheetId, accessToken, overriddenSpreadsheetId, globalSpreadsheetId, getCachedItems]);

  const autoRefresh = options.autoRefresh !== false;
  const pollIntervalMs = typeof options.pollIntervalMs === 'number' ? options.pollIntervalMs : 0;

  useEffect(() => {
    let isMounted = true;

    const safeFetch = (bypass = false) => {
      if (isMounted) {
        fetchFromSheets(bypass);
      }
    };

    safeFetch();

    const handleDataUpdated = (e: Event) => {
      const customEvt = e as CustomEvent;
      if (!customEvt.detail || !customEvt.detail.collectionName || customEvt.detail.collectionName === collectionName) {
        safeFetch(true);
      }
    };

    const handleStorage = () => {
      safeFetch(true);
    };

    const handleFocus = () => {
      safeFetch(true);
    };

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        safeFetch(true);
      }
    };

    const shouldUseTimerPolling = autoRefresh && Number.isFinite(pollIntervalMs) && pollIntervalMs > 0;

    const pollInterval = shouldUseTimerPolling ? setInterval(() => {
      if (document.visibilityState === 'visible') {
        safeFetch(false);
      }
    }, pollIntervalMs) : null;

    if (autoRefresh) {
      window.addEventListener('data_updated', handleDataUpdated);
      window.addEventListener('token_refreshed', handleStorage);
      window.addEventListener('spreadsheet_id_updated', handleStorage);
      window.addEventListener('focus', handleFocus);
      document.addEventListener('visibilitychange', handleVisibilityChange);
    }

    return () => {
      isMounted = false;
      if (pollInterval) clearInterval(pollInterval);
      if (autoRefresh) {
        window.removeEventListener('data_updated', handleDataUpdated);
        window.removeEventListener('token_refreshed', handleStorage);
        window.removeEventListener('spreadsheet_id_updated', handleStorage);
        window.removeEventListener('focus', handleFocus);
        document.removeEventListener('visibilitychange', handleVisibilityChange);
      }
    };
  }, [fetchFromSheets, collectionName, autoRefresh, pollIntervalMs]);

  return useMemo(() => {
    return { 
      data: sheetData, 
      loading: sheetLoading, 
      error: sheetError,
      refetch: () => fetchFromSheets(true),
      refresh: () => fetchFromSheets(true)
    };
  }, [sheetData, sheetLoading, sheetError, fetchFromSheets]);
}
