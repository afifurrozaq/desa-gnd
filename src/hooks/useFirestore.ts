// Legacy Firestore Hook - Firestore has been completely disabled in favor of Google Sheets.

export function useFirestoreQuery<T = any>(
  db: any,
  collectionName: string,
  constraints: any[] = [],
  enabled: boolean = false
) {
  return { data: [], loading: false, error: null };
}
