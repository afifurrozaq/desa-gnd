import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, User, Auth, GoogleAuthProvider, signInWithPopup } from 'firebase/auth';
import { saveData } from '../lib/dataService';
import { getSheetValues } from '../lib/sheets';

interface FirebaseContextType {
  user: User | null;
  loading: boolean;
  db: null;
  auth: Auth | null;
  accessToken: string | null;
  googleSignIn: () => Promise<void>;
  refreshAccessToken: () => Promise<string | null>;
  profile: any | null;
  spreadsheetId: string | null;
  setSpreadsheetId: (id: string | null) => void;
  updateProfile: (data: any) => void;
  syncProfileFromSheet: (uid?: string) => Promise<void>;
}

const FirebaseContext = createContext<FirebaseContextType>({ 
  user: null, 
  loading: true, 
  db: null,
  auth: null,
  accessToken: null,
  googleSignIn: async () => {},
  refreshAccessToken: async () => null,
  profile: null,
  spreadsheetId: null,
  setSpreadsheetId: () => {},
  updateProfile: () => {},
  syncProfileFromSheet: async () => {}
});

export const refreshGoogleAccessToken = async (authInstance?: Auth | null): Promise<string | null> => {
  let activeAuth = authInstance;
  if (!activeAuth) {
    try {
      activeAuth = getAuth();
    } catch (e) {}
  }
  if (!activeAuth) return null;

  const provider = new GoogleAuthProvider();
  provider.addScope('https://www.googleapis.com/auth/spreadsheets');
  provider.addScope('https://www.googleapis.com/auth/drive.file');

  try {
    const result = await signInWithPopup(activeAuth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (credential?.accessToken) {
      const token = credential.accessToken;
      localStorage.setItem('app_access_token', token);
      localStorage.setItem('app_access_token_time', Date.now().toString());
      window.dispatchEvent(new Event('storage'));
      window.dispatchEvent(new CustomEvent('token_refreshed', { detail: { token } }));
      return token;
    }
  } catch (error: any) {
    const code = error?.code || '';
    const msg = error?.message || String(error);
    if (code === 'auth/popup-closed-by-user' || code === 'auth/cancelled-popup-request' || msg.includes('popup-closed-by-user') || msg.includes('cancelled-popup-request')) {
      console.info('Google popup ditutup oleh pengguna.');
      return null;
    }
    console.warn('Google auth notice:', msg);
  }
  return null;
};

export const FirebaseProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [auth, setAuth] = useState<Auth | null>(null);

  const envSpreadsheetId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';

  const [accessToken, setAccessToken] = useState<string | null>(() => {
    return localStorage.getItem('app_access_token');
  });

  const [profile, setProfile] = useState<any | null>(null);

  const [spreadsheetId, setSpreadsheetIdState] = useState<string | null>(() => {
    return envSpreadsheetId || localStorage.getItem('app_spreadsheet_id') || null;
  });

  const setSpreadsheetId = (id: string | null) => {
    const nextId = id || envSpreadsheetId || null;
    setSpreadsheetIdState(nextId);
    if (id) {
      localStorage.setItem('app_spreadsheet_id', id);
    } else {
      localStorage.removeItem('app_spreadsheet_id');
    }
    window.dispatchEvent(new Event('spreadsheet_id_updated'));
  };

  const updateProfile = (data: any) => {
    setProfile((prev: any) => {
      const updated = { ...prev, ...data };
      if (user?.uid) {
        localStorage.setItem(`user_profile_${user.uid}`, JSON.stringify(updated));
        const activeSpreadsheetId = spreadsheetId || (import.meta as any).env?.VITE_SPREADSHEET_ID || localStorage.getItem('app_spreadsheet_id');
        const activeToken = accessToken || localStorage.getItem('app_access_token');
        if (activeSpreadsheetId) {
          saveData(null, activeToken, activeSpreadsheetId, 'users', updated, user.uid).catch(console.warn);
        }
      }
      return updated;
    });
  };

  const googleSignIn = async () => {
    if (!auth) return;
    const provider = new GoogleAuthProvider();
    provider.addScope('https://www.googleapis.com/auth/spreadsheets');
    provider.addScope('https://www.googleapis.com/auth/drive.file');
    try {
      const result = await signInWithPopup(auth, provider);
      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (credential?.accessToken) {
        setAccessToken(credential.accessToken);
        localStorage.setItem('app_access_token', credential.accessToken);
        localStorage.setItem('app_access_token_time', Date.now().toString());
        window.dispatchEvent(new Event('storage'));
        window.dispatchEvent(new CustomEvent('token_refreshed', { detail: { token: credential.accessToken } }));
      }
    } catch (error: any) {
      const code = error?.code || '';
      const msg = error?.message || String(error);
      if (code === 'auth/popup-closed-by-user' || code === 'auth/cancelled-popup-request' || msg.includes('popup-closed-by-user') || msg.includes('cancelled-popup-request')) {
        console.info('Google Sign-In popup ditutup oleh pengguna.');
        return;
      }
      console.warn('Google Sign-In notice:', msg);
    }
  };

  const refreshAccessToken = async (): Promise<string | null> => {
    const token = await refreshGoogleAccessToken(auth);
    if (token) {
      setAccessToken(token);
    }
    return token;
  };

  // Check /api/config for environment spreadsheetId
  useEffect(() => {
    fetch('/api/config')
      .then(res => res.json())
      .then(cfg => {
        if (cfg?.spreadsheetId && !spreadsheetId) {
          setSpreadsheetIdState(cfg.spreadsheetId);
          localStorage.setItem('app_spreadsheet_id', cfg.spreadsheetId);
          window.dispatchEvent(new Event('spreadsheet_id_updated'));
        }
      })
      .catch(() => {});
  }, []);

  useEffect(() => {
    const handleStorage = () => {
      const currentEnvId = (import.meta as any).env?.VITE_SPREADSHEET_ID || '';
      setSpreadsheetIdState(currentEnvId || localStorage.getItem('app_spreadsheet_id') || null);
      const storedToken = localStorage.getItem('app_access_token');
      if (storedToken) {
        setAccessToken(storedToken);
      }
    };
    window.addEventListener('storage', handleStorage);
    window.addEventListener('spreadsheet_id_updated', handleStorage);
    window.addEventListener('token_refreshed', handleStorage);
    return () => {
      window.removeEventListener('storage', handleStorage);
      window.removeEventListener('spreadsheet_id_updated', handleStorage);
      window.removeEventListener('token_refreshed', handleStorage);
    };
  }, []);

  // Check user profile from Google Sheets or local cache
  const syncProfileFromSheet = useCallback(async (uid?: string) => {
    const targetUid = uid || auth?.currentUser?.uid || user?.uid;
    if (!targetUid) return;

    // 1. Immediately check local cache_users
    const cachedUsersStr = localStorage.getItem('cache_users');
    if (cachedUsersStr) {
      try {
        const cachedUsers = JSON.parse(cachedUsersStr);
        const currentUserEmail = auth?.currentUser?.email || user?.email;
        const found = cachedUsers.find((u: any) => 
          String(u.id || u.uid) === targetUid ||
          (currentUserEmail && u.email && u.email.toLowerCase() === currentUserEmail.toLowerCase())
        );
        if (found) {
          setProfile((prev: any) => {
            if (prev) {
              const hasDiff = Object.keys(found).some(k => prev[k] !== found[k]);
              if (!hasDiff) return prev;
            }
            const updated = { ...(prev || {}), ...found };
            localStorage.setItem(`user_profile_${targetUid}`, JSON.stringify(updated));
            return updated;
          });
        }
      } catch (e) {}
    }

    // 2. Fetch authoritative profile from Google Sheets
    const activeSpreadsheetId = localStorage.getItem('app_spreadsheet_id') || (import.meta as any).env?.VITE_SPREADSHEET_ID;
    const activeToken = localStorage.getItem('app_access_token');
    if (!activeSpreadsheetId) return;

    try {
      const res = await getSheetValues(activeToken, activeSpreadsheetId, 'users!A:ZZ');
      const rows = res.values as any[][];
      if (rows && rows.length > 1) {
        const headers = rows[0];
        let idIdx = headers.indexOf('id');
        if (idIdx === -1) idIdx = headers.indexOf('uid');
        let emailIdx = headers.indexOf('email');
        const currentUserEmail = auth?.currentUser?.email || user?.email;

        const userRow = rows.slice(1).find(r => 
          (idIdx !== -1 && String(r[idIdx]) === targetUid) ||
          (emailIdx !== -1 && currentUserEmail && String(r[emailIdx]).toLowerCase() === currentUserEmail.toLowerCase())
        );

        if (userRow) {
          const remoteObj: any = {};
          headers.forEach((h, i) => {
            let val = userRow[i];
            if (typeof val === 'string') {
              if (val.toLowerCase() === 'true') val = true;
              if (val.toLowerCase() === 'false') val = false;
            }
            remoteObj[h] = val;
          });
          setProfile((prev: any) => {
            if (prev) {
              const hasDiff = Object.keys(remoteObj).some(k => prev[k] !== remoteObj[k]);
              if (!hasDiff) return prev;
            }
            const updated = { ...(prev || {}), ...remoteObj };
            localStorage.setItem(`user_profile_${targetUid}`, JSON.stringify(updated));
            return updated;
          });
        }
      }
    } catch (e) {
      console.warn('Sync profile from sheet warning:', e);
    }
  }, [auth, user]);

  useEffect(() => {
    const initFirebase = async () => {
      try {
        let firebaseConfig: any = null;
        try {
          const configResponse = await fetch('/firebase-applet-config.json');
          if (configResponse.ok) {
            firebaseConfig = await configResponse.json();
          }
        } catch (e) {
          // Network fetch failed, will use fallback
        }

        if (!firebaseConfig) {
          firebaseConfig = {
            projectId: "gen-lang-client-0786382859",
            appId: "1:877596231331:web:bd68513c77a523d235a116",
            apiKey: "AIzaSyCBRAeVjFpslb7v4ck_VEehjjWjBt49LUI",
            authDomain: "gen-lang-client-0786382859.firebaseapp.com",
            storageBucket: "gen-lang-client-0786382859.firebasestorage.app",
            messagingSenderId: "877596231331",
            measurementId: "",
            oAuthClientId: "877596231331-2nc9as436ivah450g39rc9mrpgp1lhpt.apps.googleusercontent.com",
            recaptchaSiteKey: ""
          };
        }
        
        let app;
        if (!getApps().length) {
          app = initializeApp(firebaseConfig);
        } else {
          app = getApp();
        }

        const authInstance = getAuth(app);
        setAuth(authInstance);

        const unsubscribe = onAuthStateChanged(authInstance, (currentUser) => {
          setUser(currentUser);
          if (currentUser) {
            // Restore persistent access token if not present
            const storedToken = localStorage.getItem('app_access_token');
            if (storedToken) {
              setAccessToken(storedToken);
            }

            // Restore or build profile
            const savedProfile = localStorage.getItem(`user_profile_${currentUser.uid}`);
            const isOwnerEmail = currentUser.email && (
              currentUser.email.toLowerCase() === 'travelio11111@gmail.com' ||
              currentUser.email.toLowerCase().startsWith('admin')
            );

            if (savedProfile) {
              try {
                const parsed = JSON.parse(savedProfile);
                if (isOwnerEmail) {
                  parsed.role = 'admin';
                  parsed.isVerified = true;
                }
                setProfile(parsed);
              } catch (e) {
                setProfile({
                  uid: currentUser.uid,
                  email: currentUser.email || '',
                  displayName: currentUser.displayName || 'Admin',
                  role: isOwnerEmail ? 'admin' : 'pengurus',
                  isVerified: Boolean(isOwnerEmail),
                  createdAt: Date.now()
                });
              }
            } else {
              // Check if cache_users has this user
              let initialVerified = Boolean(isOwnerEmail);
              let initialRole: any = isOwnerEmail ? 'admin' : 'pengurus';
              let initialName = currentUser.displayName || (isOwnerEmail ? 'Administrator' : 'Pengguna');
              try {
                const cachedUsers = JSON.parse(localStorage.getItem('cache_users') || '[]');
                const found = cachedUsers.find((u: any) => 
                  String(u.id || u.uid) === currentUser.uid ||
                  (currentUser.email && u.email && u.email.toLowerCase() === currentUser.email.toLowerCase())
                );
                if (found) {
                  initialVerified = isOwnerEmail || found.isVerified === true || found.isVerified === 'true' || found.isVerified === 'TRUE';
                  initialRole = isOwnerEmail ? 'admin' : (found.role || 'pengurus');
                  initialName = found.displayName || initialName;
                }
              } catch (e) {}

              const defaultProfile = {
                uid: currentUser.uid,
                email: currentUser.email || '',
                displayName: initialName,
                role: initialRole,
                isVerified: initialVerified,
                createdAt: Date.now()
              };
              setProfile(defaultProfile);
              localStorage.setItem(`user_profile_${currentUser.uid}`, JSON.stringify(defaultProfile));
            }

            // Sync authoritative profile data from Google Sheets
            syncProfileFromSheet(currentUser.uid);
          } else {
            // User logged out: clear individual user profile, but KEEP persistent database token!
            setUser(null);
            setProfile(null);
            const persistentToken = localStorage.getItem('app_access_token');
            if (persistentToken) {
              setAccessToken(persistentToken);
            }
          }
          setLoading(false);
        });

        return () => unsubscribe();
      } catch (error) {
        console.error('Error initializing Firebase Auth:', error);
        setLoading(false);
      }
    };

    initFirebase();
  }, []);

  // Sync back remote updates from Google Sheets when 'users' collection updates
  useEffect(() => {
    if (!user?.uid) return;

    const handleUpdate = (e: Event) => {
      const customEvt = e as CustomEvent;
      if (!customEvt.detail || !customEvt.detail.collectionName || customEvt.detail.collectionName === 'users') {
        syncProfileFromSheet(user.uid);
      }
    };

    window.addEventListener('data_updated', handleUpdate);
    return () => window.removeEventListener('data_updated', handleUpdate);
  }, [user?.uid, syncProfileFromSheet]);

  return (
    <FirebaseContext.Provider value={{ 
      user, 
      loading, 
      db: null, 
      auth, 
      accessToken, 
      googleSignIn, 
      refreshAccessToken, 
      profile, 
      spreadsheetId, 
      setSpreadsheetId,
      updateProfile,
      syncProfileFromSheet 
    }}>
      {children}
    </FirebaseContext.Provider>
  );
};

export const useFirebase = () => useContext(FirebaseContext);
