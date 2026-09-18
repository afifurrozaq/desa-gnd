import { useState, useEffect, useMemo, useCallback } from 'react';
import { getSheetValues } from '../lib/sheets';
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

export function useDataQuery<T>(
  collectionName: string,
  constraints: QueryWhereConstraint[] = [],
  overriddenSpreadsheetId?: string
) {
  const { accessToken, spreadsheetId: globalSpreadsheetId } = useFirebase();
  const envSpreadsheetId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
  const effectiveSpreadsheetId = overriddenSpreadsheetId || globalSpreadsheetId || envSpreadsheetId || localStorage.getItem('app_spreadsheet_id') || '';

  const constraintsStr = useMemo(() => JSON.stringify(constraints || []), [constraints]);

  // 1. Immediately read from local cache so data never flickers or disappears
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

  useEffect(() => {
    let isMounted = true;

    const fetchFromSheets = async () => {
      // Re-populate from local cache first in case cache changed
      const cachedItems = getCachedItems();
      if (cachedItems.length > 0 && isMounted) {
        setSheetData(cachedItems);
      }

      const activeSpreadsheetId = overriddenSpreadsheetId || globalSpreadsheetId || (import.meta as any).env?.VITE_SPREADSHEET_ID || localStorage.getItem('app_spreadsheet_id') || '';
      const activeAccessToken = accessToken || localStorage.getItem('app_access_token') || null;

      if (!activeSpreadsheetId) {
        if (isMounted) {
          setSheetLoading(false);
        }
        return;
      }
      
      // Only show full loading spinner if cache is completely empty to maintain 0ms instant UI feel
      if (isMounted && cachedItems.length === 0) {
        setSheetLoading(true);
      }

      try {
        const result = await getSheetValues(activeAccessToken, activeSpreadsheetId, `${collectionName}!A:ZZ`);
        const rows = result.values as any[][];
        
        if (!rows || rows.length === 0) {
          // If sheet is empty or not yet seeded, keep cached items if any
          if (isMounted) {
            if (cachedItems.length === 0) {
              setSheetData([]);
            }
            setSheetLoading(false);
          }
          return;
        }

        const headers = rows[0] || [];
        const items = rows.slice(1).map(row => {
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

        // Update local persistent cache with fresh sheet items
        try {
          localStorage.setItem(`cache_${collectionName}`, JSON.stringify(items));
        } catch (cacheErr) {}

        const filteredItems = filterItemsByConstraints<T>(items, constraints);

        if (isMounted) {
          setSheetData(filteredItems);
          setSheetError(null);
        }
      } catch (err: any) {
        if (isMounted) {
          setSheetError(err);
          // On network or auth error, fall back to local cache so data stays visible
          const fallback = getCachedItems();
          if (fallback.length > 0) {
            setSheetData(fallback);
          }
        }
      } finally {
        if (isMounted) setSheetLoading(false);
      }
    };

    fetchFromSheets();

    const handleDataUpdated = (e: Event) => {
      const customEvt = e as CustomEvent;
      if (!customEvt.detail || !customEvt.detail.collectionName || customEvt.detail.collectionName === collectionName) {
        fetchFromSheets();
      }
    };

    const handleStorage = () => {
      fetchFromSheets();
    };

    window.addEventListener('data_updated', handleDataUpdated);
    window.addEventListener('token_refreshed', handleStorage);
    window.addEventListener('spreadsheet_id_updated', handleStorage);

    return () => {
      isMounted = false;
      window.removeEventListener('data_updated', handleDataUpdated);
      window.removeEventListener('token_refreshed', handleStorage);
      window.removeEventListener('spreadsheet_id_updated', handleStorage);
    };
  }, [effectiveSpreadsheetId, accessToken, collectionName, constraintsStr]);

  return useMemo(() => {
    return { data: sheetData, loading: sheetLoading, error: sheetError };
  }, [sheetData, sheetLoading, sheetError]);
}
