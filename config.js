// Firebase project field-leader-1915 (Firebase console > Project settings > Your apps).
// This web key is meant to be public; access is controlled by firestore.rules.
export const firebaseConfig = {
  apiKey: "AIzaSyBGLUx4Tce6Wi22rA0xuNKtABIJUu-EauA",
  authDomain: "field-leader-1915.firebaseapp.com",
  projectId: "field-leader-1915",
  storageBucket: "field-leader-1915.firebasestorage.app",
  messagingSenderId: "614410178584",
  appId: "1:614410178584:web:1a39eca8b7d06e4aec736d"
};

// The owner is always an admin, even before any logins exist. Must match firestore.rules.
export const OWNER_EMAIL = "fpina@1915south.com";
// Only work emails can create an account.
export const EMAIL_DOMAIN = "1915south.com";
