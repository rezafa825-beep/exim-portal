import { getApp, getApps, initializeApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';

// Firebase Web configuration.
// These values identify the Firebase project; authorization must be enforced server-side.
const firebaseConfig = {
  apiKey: 'AIzaSyCTizvApcTs_WiIvStv-Tp01EllXX6jPgM',
  authDomain: 'reza-system-authenticati-5badc.firebaseapp.com',
  projectId: 'reza-system-authenticati-5badc',
  storageBucket: 'reza-system-authenticati-5badc.firebasestorage.app',
  messagingSenderId: '1052494346845',
  appId: '1:1052494346845:web:87729b7208126027872bb4',
};

export const firebaseApp = getApps().length ? getApp() : initializeApp(firebaseConfig);
export const auth = getAuth(firebaseApp);
