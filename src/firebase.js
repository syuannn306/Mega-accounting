import { initializeApp } from "firebase/app";
import { getDatabase, ref, onValue, set, get } from "firebase/database";

const firebaseConfig = {
  apiKey: "AIzaSyB7ekviPcETr0_ypapWKwswoeTMbdUvwhU",
  authDomain: "mega-c1409.firebaseapp.com",
  databaseURL: "https://mega-c1409-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "mega-c1409",
  storageBucket: "mega-c1409.firebasestorage.app",
  messagingSenderId: "1025053696548",
  appId: "1:1025053696548:web:a79a53914740d93d4e6eb1"
};

const app = initializeApp(firebaseConfig);
export const db = getDatabase(app);
export { ref, onValue, set, get };
