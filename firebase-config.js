// ── Firebase Configuration ──────────────────────────
// Replace the values below with your Firebase project config.
// Get these from: Firebase Console → Project Settings → General → Your apps → Web app
const firebaseConfig = {
  apiKey: "AIzaSyBrGdtV_U4vzf8kbl8yeO7Oadq9nfQvgoA",
  authDomain: "canteen-manager-fda66.firebaseapp.com",
  projectId: "canteen-manager-fda66",
  storageBucket: "canteen-manager-fda66.firebasestorage.app",
  messagingSenderId: "874701270392",
  appId: "1:874701270392:web:4c886c610d7b35a3676c8d"
};

// Initialize Firebase
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();
const auth = firebase.auth();
