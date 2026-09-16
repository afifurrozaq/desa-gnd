import React, { createContext, useContext, useEffect, useState } from 'react';
import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth, onAuthStateChanged, User, Auth } from 'firebase/auth';
import { getFirestore, Firestore, doc, getDocFromCache, getDocFromServer } from 'firebase/firestore';

interface FirebaseContextType {
  user: User | null;
  loading: boolean;
  db: Firestore | null;
  auth: Auth | null;
}

const FirebaseContext = createContext<FirebaseContextType>({ 
  user: null, 
  loading: true, 
  db: null,
  auth: null 
});

async function testConnection(db: Firestore) {
  try {
    // Testing connection as per skill requirement
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.error("Please check your Firebase configuration.");
    }
  }
}

export const FirebaseProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [db, setDb] = useState<Firestore | null>(null);
  const [auth, setAuth] = useState<Auth | null>(null);

  useEffect(() => {
    const initFirebase = async () => {
      try {
        const firebaseConfig = {
          apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
          authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
          projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
          storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
          messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
          appId: import.meta.env.VITE_FIREBASE_APP_ID,
          measurementId: import.meta.env.VITE_FIREBASE_MEASUREMENT_ID || '',
        };

        if (!firebaseConfig.apiKey || !firebaseConfig.projectId || !firebaseConfig.appId) {
          console.warn('Firebase config not found. Please set up your .env file.');
          setLoading(false);
          return;
        }

        let app;
        if (!getApps().length) {
          app = initializeApp(firebaseConfig);
        } else {
          app = getApp();
        }

        const authInstance = getAuth(app);
        const firestoreDatabaseId = import.meta.env.VITE_FIREBASE_FIRESTORE_DATABASE_ID;
        const firestore = firestoreDatabaseId
          ? getFirestore(app, firestoreDatabaseId)
          : getFirestore(app);

        setAuth(authInstance);
        setDb(firestore);
        testConnection(firestore);

        const unsubscribe = onAuthStateChanged(authInstance, (user) => {
          setUser(user);
          setLoading(false);
        });

        return () => unsubscribe();
      } catch (error) {
        console.error('Error initializing Firebase:', error);
        setLoading(false);
      }
    };

    initFirebase();
  }, []);

  return (
    <FirebaseContext.Provider value={{ user, loading, db, auth }}>
      {children}
    </FirebaseContext.Provider>
  );
};

export const useFirebase = () => useContext(FirebaseContext);
