// seed-kelas.js
// Script sekali jalan untuk mengisi data kelas ke Firestore.
// Jalankan dengan: node seed-kelas.js

import { initializeApp, cert } from "firebase-admin/app";
import { getFirestore } from "firebase-admin/firestore";
import { readFileSync } from "fs";

// Baca service account key
const serviceAccount = JSON.parse(
  readFileSync("./serviceAccountKey.json", "utf-8")
);

initializeApp({
  credential: cert(serviceAccount),
});

const db = getFirestore();

const tingkat = ["X", "XI", "XII"];
const huruf = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];

async function seedKelas() {
  const batch = db.batch();
  let count = 0;

  for (const t of tingkat) {
    for (const h of huruf) {
      const id = `${t}-${h}`;
      const ref = db.collection("kelas").doc(id);
      batch.set(ref, {
        nama: id,
        qrId: id,
      });
      count++;
    }
  }

  await batch.commit();
  console.log(`Selesai. ${count} kelas berhasil dibuat (X-A s.d XII-J).`);
}

seedKelas().catch((err) => {
  console.error("Gagal seed data:", err);
});