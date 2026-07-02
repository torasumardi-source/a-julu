import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { getFirestore } from "firebase/firestore";

export const firebaseConfig = {
  apiKey: "AIzaSyAPEvDpMfVNEWGkNtSWGZXGVvQ6mJQKyv4",
  authDomain: "a-julu.firebaseapp.com",
  projectId: "a-julu",
  storageBucket: "a-julu.firebasestorage.app",
  messagingSenderId: "769136506347",
  appId: "1:769136506347:web:1e7e3b08244d0193afd5ae",
};

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app);