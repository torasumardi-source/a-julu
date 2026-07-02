import { useState, useEffect, useRef } from "react";
import { BrowserRouter, Routes, Route, Navigate, useNavigate, useParams, Link } from "react-router-dom";
import { initializeApp, deleteApp } from "firebase/app";
import {
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  createUserWithEmailAndPassword,
  getAuth,
  signOut as signOutSecondary,
} from "firebase/auth";
import {
  doc,
  getDoc,
  setDoc,
  addDoc,
  deleteDoc,
  collection,
  query,
  where,
  getDocs,
} from "firebase/firestore";
import { Html5QrcodeScanner } from "html5-qrcode";
import { QRCodeSVG } from "qrcode.react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { auth, db, firebaseConfig } from "./firebase";

/* ========================================================================
   DATA DASAR
======================================================================== */
const HARI_LIST = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
const KELAS_LIST = [];
for (const t of ["X", "XI", "XII"]) {
  for (const h of ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]) {
    KELAS_LIST.push(`${t}-${h}`);
  }
}

async function buatAkunTanpaLogout(email, password) {
  const secondaryApp = initializeApp(firebaseConfig, "Secondary" + Date.now());
  const secondaryAuth = getAuth(secondaryApp);
  const cred = await createUserWithEmailAndPassword(secondaryAuth, email, password);
  await signOutSecondary(secondaryAuth);
  await deleteApp(secondaryApp);
  return cred.user.uid;
}

/* ========================================================================
   DESAIN — warna teal/amber, tema gelap, kartu kaca (glassmorphism)
======================================================================== */
const colors = {
  bg: "#0a1210",
  panel: "rgba(15,32,28,0.85)",
  panelAlt: "rgba(255,255,255,0.04)",
  border: "rgba(255,255,255,0.08)",
  textPrimary: "#f0fdf9",
  textSecondary: "#a7c4bb",
  muted: "#6b8f83",
  teal: "#14b8a6",
  tealSoft: "rgba(20,184,166,0.15)",
  amber: "#f59e0b",
  green: "#22c55e",
  red: "#ef4444",
  fontMono: "'JetBrains Mono', monospace",
};

const GLOBAL_STYLE = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
  * { box-sizing: border-box; }
  html, body { margin:0; padding:0; background:${colors.bg}; font-family:'Inter',sans-serif; color:${colors.textPrimary}; }
  #root { min-height:100vh; }
  body {
    background:
      radial-gradient(circle at 15% 10%, rgba(20,184,166,0.10), transparent 40%),
      radial-gradient(circle at 85% 90%, rgba(245,158,11,0.08), transparent 40%),
      ${colors.bg};
  }
  .aj-card {
    background:${colors.panel}; border:1px solid ${colors.border}; border-radius:18px;
    backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px);
    box-shadow: 0 8px 32px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.05);
  }
  .aj-input, .aj-select {
    background:rgba(255,255,255,0.05); color:${colors.textPrimary};
    border:1px solid ${colors.border}; border-radius:10px; padding:11px 14px;
    font-size:14px; font-family:inherit; width:100%; outline:none;
    transition: border-color .15s, background .15s;
  }
  .aj-input::placeholder { color:${colors.muted}; }
  .aj-input:focus, .aj-select:focus {
    border-color:${colors.teal}; background:rgba(20,184,166,0.08);
    box-shadow:0 0 0 3px rgba(20,184,166,0.18);
  }
  .aj-select option { background:#0f201c; color:${colors.textPrimary}; }
  .aj-btn {
    background:linear-gradient(135deg, ${colors.teal}, #0d9488);
    color:#04140f; font-weight:700; border:none; border-radius:10px;
    padding:12px 20px; font-size:14px; cursor:pointer;
    transition: transform .15s, box-shadow .15s, filter .15s;
  }
  .aj-btn:hover:not(:disabled) { transform:translateY(-1px); box-shadow:0 8px 20px rgba(20,184,166,0.35); }
  .aj-btn:disabled { opacity:.5; cursor:not-allowed; }
  .aj-btn-ghost {
    background:rgba(255,255,255,0.05); color:${colors.textSecondary};
    border:1px solid ${colors.border}; border-radius:10px; padding:10px 16px;
    font-size:13px; cursor:pointer; transition:.15s;
  }
  .aj-btn-ghost:hover { background:rgba(20,184,166,0.12); border-color:${colors.teal}; color:${colors.textPrimary}; }
  .aj-tab {
    background:transparent; border:none; color:${colors.muted}; font-weight:600;
    font-size:13px; padding:10px 16px; border-radius:999px; cursor:pointer; transition:.15s;
  }
  .aj-tab.active { background:${colors.tealSoft}; color:${colors.teal}; }
  .aj-table { width:100%; border-collapse:collapse; font-size:13px; }
  .aj-table th {
    text-align:left; padding:10px 12px; background:rgba(20,184,166,0.15); color:${colors.teal};
    font-weight:700; text-transform:uppercase; font-size:11px; letter-spacing:.05em;
  }
  .aj-table td { padding:10px 12px; border-bottom:1px solid ${colors.border}; color:${colors.textSecondary}; }
  .aj-badge { display:inline-block; padding:3px 10px; border-radius:999px; font-size:12px; font-weight:700; }
  ::-webkit-scrollbar { width:6px; }
  ::-webkit-scrollbar-thumb { background:rgba(20,184,166,0.35); border-radius:999px; }
  @keyframes ajPulse { 0%,100%{opacity:1;} 50%{opacity:.55;} }
  .aj-pulse { animation: ajPulse 2s ease-in-out infinite; }
`;

function GlobalStyle() {
  return <style>{GLOBAL_STYLE}</style>;
}

function Footer() {
  return (
    <footer
      style={{
        textAlign: "center",
        padding: "24px 16px 32px",
        color: colors.muted,
        fontSize: 12,
        lineHeight: 1.6,
      }}
    >
      <div style={{ opacity: 0.6, marginBottom: 4 }}>A-JULU — Sistem Absensi Digital</div>
      <div>SMA Negeri 1 Lumbanjulu • Dikembangkan oleh Restuadi G. Sinaga, S.Kom</div>
    </footer>
  );
}

function Logo() {
  return (
    <div style={{ textAlign: "center", marginBottom: 4 }}>
      <div
        style={{
          display: "inline-flex",
          alignItems: "center",
          justifyContent: "center",
          width: 56,
          height: 56,
          borderRadius: 16,
          background: `linear-gradient(135deg, ${colors.teal}, #0d9488)`,
          fontSize: 24,
          marginBottom: 10,
          boxShadow: "0 8px 24px rgba(20,184,166,0.35)",
        }}
      >
        ✓
      </div>
      <h1 style={{ margin: 0, fontSize: 26, fontWeight: 800, letterSpacing: "-0.02em" }}>
        A-JULU
      </h1>
      <p style={{ margin: "4px 0 0", color: colors.textSecondary, fontSize: 13 }}>
        Sistem Absensi Digital SMA N 1 Lumbanjulu
      </p>
    </div>
  );
}

/* ========================================================================
   LOGIN SISWA
======================================================================== */
function LoginSiswa() {
  const [nisn, setNisn] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const navigate = useNavigate();

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      await signInWithEmailAndPassword(auth, `${nisn}@ajulu.app`, nisn);
      navigate("/siswa");
    } catch (err) {
      console.error(err);
      setError("NISN tidak ditemukan atau salah. Coba lagi.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="aj-card" style={{ maxWidth: 380, margin: "0 auto", padding: 32, width: "100%" }}>
        <Logo />
        <p style={{ textAlign: "center", color: colors.textSecondary, fontSize: 13, marginTop: 20, marginBottom: 20 }}>
          Login Siswa
        </p>
        <form onSubmit={handleLogin}>
          <input
            className="aj-input"
            type="text"
            placeholder="Masukkan NISN"
            value={nisn}
            onChange={(e) => setNisn(e.target.value)}
            style={{ marginBottom: 12 }}
            required
          />
          <button className="aj-btn" type="submit" disabled={loading} style={{ width: "100%" }}>
            {loading ? "Memproses..." : "Masuk"}
          </button>
        </form>
        {error && <p style={{ color: colors.red, fontSize: 13, marginTop: 12, textAlign: "center" }}>{error}</p>}
        <p style={{ textAlign: "center", marginTop: 20 }}>
          <Link to="/login-guru" style={{ color: colors.muted, fontSize: 12 }}>
            Login sebagai guru / admin →
          </Link>
        </p>
      </div>
      <Footer />
    </div>
  );
}

/* ========================================================================
   LOGIN GURU / ADMIN
======================================================================== */
function LoginGuru() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const navigate = useNavigate();

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    setLoading(true);
    try {
      await signInWithEmailAndPassword(auth, email, password);
      navigate("/guru");
    } catch (err) {
      console.error(err);
      setError("Email atau password salah.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="aj-card" style={{ maxWidth: 380, margin: "0 auto", padding: 32, width: "100%" }}>
        <Logo />
        <p style={{ textAlign: "center", color: colors.textSecondary, fontSize: 13, marginTop: 20, marginBottom: 20 }}>
          Login Guru / Admin
        </p>
        <form onSubmit={handleLogin}>
          <input
            className="aj-input"
            type="email"
            placeholder="Email"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            style={{ marginBottom: 10 }}
            required
          />
          <input
            className="aj-input"
            type="password"
            placeholder="Password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            style={{ marginBottom: 12 }}
            required
          />
          <button className="aj-btn" type="submit" disabled={loading} style={{ width: "100%" }}>
            {loading ? "Memproses..." : "Masuk"}
          </button>
        </form>
        {error && <p style={{ color: colors.red, fontSize: 13, marginTop: 12, textAlign: "center" }}>{error}</p>}
        <p style={{ textAlign: "center", marginTop: 20 }}>
          <Link to="/login-siswa" style={{ color: colors.muted, fontSize: 12 }}>
            ← Login sebagai siswa
          </Link>
        </p>
      </div>
      <Footer />
    </div>
  );
}

/* ========================================================================
   DASHBOARD SISWA
======================================================================== */
const HARI_ARR = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

function DashboardSiswa() {
  const [siswa, setSiswa] = useState(null);
  const [mapelList, setMapelList] = useState([]);
  const [showMapel, setShowMapel] = useState(false);
  const [loading, setLoading] = useState(true);
  const navigate = useNavigate();

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (user) => {
      if (!user) return navigate("/login-siswa");
      const nisn = user.email.split("@")[0];
      const snap = await getDoc(doc(db, "siswa", nisn));
      if (snap.exists()) setSiswa({ nisn, ...snap.data() });
      setLoading(false);
    });
    return () => unsub();
  }, [navigate]);

  async function handleBukaAbsensi() {
    if (!siswa) return;
    const hariIni = HARI_ARR[new Date().getDay()];
    const q = query(collection(db, "jadwal"), where("kelas", "==", siswa.kelas), where("hari", "==", hariIni));
    const snap = await getDocs(q);
    setMapelList(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    setShowMapel(true);
  }

  function handleLogout() {
    signOut(auth).then(() => navigate("/login-siswa"));
  }

  if (loading) return <div style={{ padding: 40, textAlign: "center", color: colors.muted }}>Memuat...</div>;

  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="aj-card" style={{ maxWidth: 420, margin: "0 auto", padding: 28, width: "100%" }}>
        <Logo />
        <div style={{ textAlign: "center", marginTop: 20, marginBottom: 8 }}>
          <div style={{ fontSize: 16, fontWeight: 700 }}>{siswa?.nama}</div>
          <span className="aj-badge" style={{ background: colors.tealSoft, color: colors.teal, marginTop: 6 }}>
            Kelas {siswa?.kelas}
          </span>
        </div>

        {!showMapel && (
          <button
            className="aj-btn"
            onClick={handleBukaAbsensi}
            style={{ width: "100%", marginTop: 24, padding: 18, fontSize: 16 }}
          >
            📍 Absensi
          </button>
        )}

        {showMapel && (
          <div style={{ marginTop: 20 }}>
            <p style={{ color: colors.textSecondary, fontSize: 13, marginBottom: 10 }}>
              Pilih mata pelajaran hari ini:
            </p>
            {mapelList.length === 0 && (
              <p style={{ color: colors.muted, fontSize: 13, textAlign: "center", padding: 20 }}>
                Tidak ada jadwal hari ini.
              </p>
            )}
            {mapelList.map((m) => (
              <button
                key={m.id}
                onClick={() => navigate(`/siswa/scan/${m.id}`)}
                className="aj-btn-ghost"
                style={{ width: "100%", textAlign: "left", marginBottom: 8, display: "flex", justifyContent: "space-between" }}
              >
                <span>{m.mapel}</span>
                <span style={{ color: colors.amber, fontFamily: colors.fontMono }}>{m.jamMulai}</span>
              </button>
            ))}
          </div>
        )}

        <button onClick={handleLogout} className="aj-btn-ghost" style={{ width: "100%", marginTop: 20 }}>
          Logout
        </button>
      </div>
      <Footer />
    </div>
  );
}

/* ========================================================================
   SCAN QR
======================================================================== */
function ScanQR() {
  const { jadwalId } = useParams();
  const navigate = useNavigate();
  const [siswa, setSiswa] = useState(null);
  const [jadwal, setJadwal] = useState(null);
  const [status, setStatus] = useState("memuat");
  const [pesan, setPesan] = useState("");
  const scannerRef = useRef(null);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (user) => {
      if (!user) return navigate("/login-siswa");
      const nisn = user.email.split("@")[0];
      const siswaSnap = await getDoc(doc(db, "siswa", nisn));
      const jadwalSnap = await getDoc(doc(db, "jadwal", jadwalId));
      if (siswaSnap.exists() && jadwalSnap.exists()) {
        setSiswa({ nisn, ...siswaSnap.data() });
        setJadwal(jadwalSnap.data());
        setStatus("scanning");
      } else {
        setPesan("Data siswa atau jadwal tidak ditemukan.");
        setStatus("gagal");
      }
    });
    return () => unsub();
  }, [jadwalId, navigate]);

  useEffect(() => {
    if (status !== "scanning") return;
    const scanner = new Html5QrcodeScanner("qr-reader", { fps: 10, qrbox: 250 }, false);
    scannerRef.current = scanner;
    scanner.render((decodedText) => handleScanSukses(decodedText), () => {});
    return () => { scanner.clear().catch(() => {}); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [status]);

  async function handleScanSukses(decodedText) {
    if (!scannerRef.current) return;
    await scannerRef.current.clear().catch(() => {});

    if (decodedText.trim() !== jadwal.kelas) {
      setPesan("QR ini bukan untuk kelas Anda. Pastikan scan QR di kelas yang benar.");
      setStatus("gagal");
      return;
    }
    const [h, m] = jadwal.jamMulai.split(":").map(Number);
    const now = new Date();
    const jamMulaiDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), h, m, 0);
    const selisihMenit = (now - jamMulaiDate) / 60000;

    if (selisihMenit < 0) {
      setPesan("Belum waktunya absensi untuk mata pelajaran ini.");
      setStatus("gagal");
      return;
    }
    if (selisihMenit > 5) {
      setPesan("Waktu absensi sudah lewat (lebih dari 5 menit dari jam mulai).");
      setStatus("gagal");
      return;
    }

    const tanggal = now.toISOString().split("T")[0];
    const absensiId = `${siswa.nisn}_${jadwalId}_${tanggal}`;
    const absensiRef = doc(db, "absensi", absensiId);
    const existing = await getDoc(absensiRef);
    if (existing.exists()) {
      setPesan("Anda sudah absen untuk mata pelajaran ini hari ini.");
      setStatus("sudah");
      return;
    }

    await setDoc(absensiRef, {
      nisn: siswa.nisn, nama: siswa.nama, kelas: siswa.kelas, mapel: jadwal.mapel,
      tanggal, jamScan: now.toTimeString().slice(0, 5), status: "Hadir",
    });
    setStatus("sukses");
  }

  return (
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", justifyContent: "center", padding: 20 }}>
      <GlobalStyle />
      <div className="aj-card" style={{ maxWidth: 420, margin: "0 auto", padding: 28, width: "100%" }}>
        <Logo />

        {status === "memuat" && <p style={{ textAlign: "center", color: colors.muted, marginTop: 20 }}>Memuat data...</p>}

        {status === "scanning" && (
          <>
            <p style={{ textAlign: "center", color: colors.textSecondary, fontSize: 13, marginTop: 20 }}>
              {jadwal?.mapel} — arahkan kamera ke QR di dinding kelas
            </p>
            <div id="qr-reader" style={{ marginTop: 14, borderRadius: 14, overflow: "hidden" }}></div>
          </>
        )}

        {status === "sukses" && (
          <div style={{ textAlign: "center", marginTop: 24 }}>
            <div className="aj-pulse" style={{ fontSize: 56, color: colors.green }}>✓</div>
            <h2 style={{ color: colors.green, margin: "8px 0" }}>Anda Hadir</h2>
            <p style={{ color: colors.textSecondary, fontSize: 13 }}>{jadwal?.mapel} — tercatat hadir</p>
            <button className="aj-btn" onClick={() => navigate("/siswa")} style={{ marginTop: 20 }}>Kembali</button>
          </div>
        )}

        {(status === "gagal" || status === "sudah") && (
          <div style={{ textAlign: "center", marginTop: 24 }}>
            <div style={{ fontSize: 48, color: status === "sudah" ? colors.amber : colors.red }}>
              {status === "sudah" ? "◷" : "✕"}
            </div>
            <h2 style={{ color: status === "sudah" ? colors.amber : colors.red, margin: "8px 0" }}>
              {status === "sudah" ? "Sudah Absen" : "Gagal"}
            </h2>
            <p style={{ color: colors.textSecondary, fontSize: 13 }}>{pesan}</p>
            <button className="aj-btn-ghost" onClick={() => navigate("/siswa")} style={{ marginTop: 20 }}>Kembali</button>
          </div>
        )}
      </div>
      <Footer />
    </div>
  );
}

/* ========================================================================
   REKAP ABSENSI
======================================================================== */
function RekapAbsensi() {
  const [kelas, setKelas] = useState(KELAS_LIST[0]);
  const [mapel, setMapel] = useState("");
  const [tanggal, setTanggal] = useState(new Date().toISOString().split("T")[0]);
  const [hasil, setHasil] = useState(null);
  const [loading, setLoading] = useState(false);

  async function handleCari() {
    setLoading(true);
    const siswaSnap = await getDocs(query(collection(db, "siswa"), where("kelas", "==", kelas)));
    const semuaSiswa = siswaSnap.docs.map((d) => ({ nisn: d.id, ...d.data() }));
    const absensiSnap = await getDocs(query(collection(db, "absensi"), where("kelas", "==", kelas), where("mapel", "==", mapel), where("tanggal", "==", tanggal)));
    const hadirMap = {};
    absensiSnap.docs.forEach((d) => { hadirMap[d.data().nisn] = d.data(); });
    setHasil(semuaSiswa.map((s) => ({
      nisn: s.nisn, nama: s.nama,
      status: hadirMap[s.nisn] ? "Hadir" : "Absen",
      jamScan: hadirMap[s.nisn]?.jamScan || "-",
    })));
    setLoading(false);
  }

  async function handleDownloadExcel() {
    if (!hasil) return;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Rekap Absensi");
    ws.columns = [
      { header: "NISN", key: "nisn", width: 15 },
      { header: "Nama", key: "nama", width: 25 },
      { header: "Status", key: "status", width: 12 },
      { header: "Jam Scan", key: "jamScan", width: 12 },
    ];
    hasil.forEach((row) => ws.addRow(row));
    ws.getRow(1).font = { bold: true };
    const buffer = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buffer], { type: "application/octet-stream" }), `Rekap_${kelas}_${mapel}_${tanggal}.xlsx`);
  }

  return (
    <div className="aj-card" style={{ padding: 24 }}>
      <h3 style={{ marginTop: 0 }}>Rekap Absensi</h3>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
        <select className="aj-select" style={{ width: "auto" }} value={kelas} onChange={(e) => setKelas(e.target.value)}>
          {KELAS_LIST.map((k) => <option key={k} value={k}>{k}</option>)}
        </select>
        <input className="aj-input" style={{ width: "auto" }} type="text" placeholder="Mata Pelajaran" value={mapel} onChange={(e) => setMapel(e.target.value)} />
        <input className="aj-input" style={{ width: "auto" }} type="date" value={tanggal} onChange={(e) => setTanggal(e.target.value)} />
        <button className="aj-btn" onClick={handleCari} disabled={loading}>{loading ? "Mencari..." : "Cari"}</button>
      </div>

      {hasil && (
        <>
          <button className="aj-btn-ghost" onClick={handleDownloadExcel} style={{ marginBottom: 14 }}>⬇ Download Excel</button>
          <table className="aj-table">
            <thead><tr><th>NISN</th><th>Nama</th><th>Status</th><th>Jam Scan</th></tr></thead>
            <tbody>
              {hasil.map((r) => (
                <tr key={r.nisn}>
                  <td>{r.nisn}</td><td>{r.nama}</td>
                  <td>
                    <span className="aj-badge" style={{ background: r.status === "Hadir" ? "rgba(34,197,94,0.15)" : "rgba(239,68,68,0.15)", color: r.status === "Hadir" ? colors.green : colors.red }}>
                      {r.status}
                    </span>
                  </td>
                  <td>{r.jamScan}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
    </div>
  );
}

/* ========================================================================
   KELOLA JADWAL
======================================================================== */
function KelolaJadwal() {
  const [kelas, setKelas] = useState(KELAS_LIST[0]);
  const [mapel, setMapel] = useState("");
  const [hari, setHari] = useState(HARI_LIST[0]);
  const [jamMulai, setJamMulai] = useState("");
  const [daftar, setDaftar] = useState([]);
  const [pesan, setPesan] = useState("");

  async function muatDaftar() {
    const snap = await getDocs(collection(db, "jadwal"));
    setDaftar(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
  }
  useEffect(() => { muatDaftar(); }, []);

  async function handleTambah(e) {
    e.preventDefault();
    if (!mapel || !jamMulai) return setPesan("Isi mata pelajaran dan jam mulai.");
    await addDoc(collection(db, "jadwal"), { kelas, mapel, hari, jamMulai });
    setMapel(""); setJamMulai(""); setPesan("Jadwal ditambahkan.");
    muatDaftar();
  }
  async function handleHapus(id) { await deleteDoc(doc(db, "jadwal", id)); muatDaftar(); }

  return (
    <div className="aj-card" style={{ padding: 24 }}>
      <h3 style={{ marginTop: 0 }}>Kelola Jadwal</h3>
      <form onSubmit={handleTambah} style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
        <select className="aj-select" style={{ width: "auto" }} value={kelas} onChange={(e) => setKelas(e.target.value)}>
          {KELAS_LIST.map((k) => <option key={k} value={k}>{k}</option>)}
        </select>
        <input className="aj-input" style={{ width: "auto" }} type="text" placeholder="Mata Pelajaran" value={mapel} onChange={(e) => setMapel(e.target.value)} />
        <select className="aj-select" style={{ width: "auto" }} value={hari} onChange={(e) => setHari(e.target.value)}>
          {HARI_LIST.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
        <input className="aj-input" style={{ width: "auto" }} type="time" value={jamMulai} onChange={(e) => setJamMulai(e.target.value)} />
        <button className="aj-btn" type="submit">Tambah Jadwal</button>
      </form>
      {pesan && <p style={{ color: colors.teal, fontSize: 13 }}>{pesan}</p>}
      <table className="aj-table">
        <thead><tr><th>Kelas</th><th>Mapel</th><th>Hari</th><th>Jam Mulai</th><th></th></tr></thead>
        <tbody>
          {daftar.map((j) => (
            <tr key={j.id}>
              <td>{j.kelas}</td><td>{j.mapel}</td><td>{j.hari}</td><td>{j.jamMulai}</td>
              <td><button className="aj-btn-ghost" onClick={() => handleHapus(j.id)}>Hapus</button></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ========================================================================
   KELOLA SISWA
======================================================================== */
function KelolaSiswa() {
  const [nisn, setNisn] = useState("");
  const [nama, setNama] = useState("");
  const [kelas, setKelas] = useState(KELAS_LIST[0]);
  const [pesan, setPesan] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleTambah(e) {
    e.preventDefault();
    if (!nisn || !nama) return setPesan("Isi NISN dan nama.");
    setLoading(true);
    try {
      await buatAkunTanpaLogout(`${nisn}@ajulu.app`, nisn);
      await setDoc(doc(db, "siswa", nisn), { nisn, nama, kelas });
      setPesan(`Siswa ${nama} (${nisn}) berhasil ditambahkan.`);
      setNisn(""); setNama("");
    } catch (err) {
      console.error(err);
      setPesan("Gagal menambahkan siswa. NISN mungkin sudah terdaftar.");
    } finally { setLoading(false); }
  }

  return (
    <div className="aj-card" style={{ padding: 24 }}>
      <h3 style={{ marginTop: 0 }}>Kelola Siswa</h3>
      <form onSubmit={handleTambah} style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <input className="aj-input" style={{ width: "auto" }} type="text" placeholder="NISN" value={nisn} onChange={(e) => setNisn(e.target.value)} />
        <input className="aj-input" style={{ width: "auto" }} type="text" placeholder="Nama" value={nama} onChange={(e) => setNama(e.target.value)} />
        <select className="aj-select" style={{ width: "auto" }} value={kelas} onChange={(e) => setKelas(e.target.value)}>
          {KELAS_LIST.map((k) => <option key={k} value={k}>{k}</option>)}
        </select>
        <button className="aj-btn" type="submit" disabled={loading}>{loading ? "Memproses..." : "Tambah Siswa"}</button>
      </form>
      {pesan && <p style={{ color: colors.teal, fontSize: 13 }}>{pesan}</p>}
      <p style={{ color: colors.muted, fontSize: 12 }}>Catatan: password login siswa otomatis sama dengan NISN.</p>
    </div>
  );
}

/* ========================================================================
   KELOLA GURU
======================================================================== */
function KelolaGuru() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [nama, setNama] = useState("");
  const [role, setRole] = useState("guru");
  const [pesan, setPesan] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleTambah(e) {
    e.preventDefault();
    if (!email || !password || !nama) return setPesan("Lengkapi semua field.");
    setLoading(true);
    try {
      const uid = await buatAkunTanpaLogout(email, password);
      await setDoc(doc(db, "guru", uid), { nama, email, role });
      setPesan(`Akun ${role} ${nama} berhasil dibuat.`);
      setEmail(""); setPassword(""); setNama("");
    } catch (err) {
      console.error(err);
      setPesan("Gagal membuat akun. Cek apakah email sudah terdaftar.");
    } finally { setLoading(false); }
  }

  return (
    <div className="aj-card" style={{ padding: 24 }}>
      <h3 style={{ marginTop: 0 }}>Kelola Guru / Admin</h3>
      <form onSubmit={handleTambah} style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <input className="aj-input" style={{ width: "auto" }} type="text" placeholder="Nama" value={nama} onChange={(e) => setNama(e.target.value)} />
        <input className="aj-input" style={{ width: "auto" }} type="email" placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} />
        <input className="aj-input" style={{ width: "auto" }} type="password" placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} />
        <select className="aj-select" style={{ width: "auto" }} value={role} onChange={(e) => setRole(e.target.value)}>
          <option value="guru">Guru</option>
          <option value="admin">Admin</option>
        </select>
        <button className="aj-btn" type="submit" disabled={loading}>{loading ? "Memproses..." : "Tambah Akun"}</button>
      </form>
      {pesan && <p style={{ color: colors.teal, fontSize: 13 }}>{pesan}</p>}
    </div>
  );
}

/* ========================================================================
   GENERATE QR
======================================================================== */
function GenerateQR() {
  return (
    <div className="aj-card" style={{ padding: 24 }}>
      <h3 style={{ marginTop: 0 }}>QR Code Kelas (cetak &amp; tempel di dinding)</h3>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 16 }}>
        {KELAS_LIST.map((k) => (
          <div key={k} style={{ textAlign: "center", background: "#fff", borderRadius: 12, padding: 12 }}>
            <QRCodeSVG value={k} size={110} />
            <p style={{ fontWeight: 700, marginTop: 8, color: "#0a1210" }}>{k}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ========================================================================
   DASHBOARD GURU (KONTAINER)
======================================================================== */
function DashboardGuru() {
  const [profil, setProfil] = useState(null);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("rekap");
  const navigate = useNavigate();

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (user) => {
      if (!user) return navigate("/login-guru");
      const snap = await getDoc(doc(db, "guru", user.uid));
      setProfil(snap.exists() ? snap.data() : { nama: user.email, role: "guru" });
      setLoading(false);
    });
    return () => unsub();
  }, [navigate]);

  function handleLogout() { signOut(auth).then(() => navigate("/login-guru")); }

  if (loading) return <div style={{ padding: 40, textAlign: "center", color: colors.muted }}>Memuat...</div>;
  const isAdmin = profil?.role === "admin";

  return (
    <div style={{ minHeight: "100vh", padding: "24px 20px" }}>
      <GlobalStyle />
      <div style={{ maxWidth: 960, margin: "0 auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 40, height: 40, borderRadius: 12, background: `linear-gradient(135deg, ${colors.teal}, #0d9488)`, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800 }}>✓</div>
            <div>
              <div style={{ fontWeight: 800, fontSize: 16 }}>A-JULU — {profil?.nama}</div>
              <span className="aj-badge" style={{ background: colors.tealSoft, color: colors.teal, fontSize: 11 }}>{profil?.role}</span>
            </div>
          </div>
          <button className="aj-btn-ghost" onClick={handleLogout}>Logout</button>
        </div>

        <div className="aj-card" style={{ padding: 8, marginBottom: 20, display: "flex", gap: 6, flexWrap: "wrap" }}>
          <button className={`aj-tab ${tab === "rekap" ? "active" : ""}`} onClick={() => setTab("rekap")}>Rekap Absensi</button>
          {isAdmin && <button className={`aj-tab ${tab === "jadwal" ? "active" : ""}`} onClick={() => setTab("jadwal")}>Kelola Jadwal</button>}
          {isAdmin && <button className={`aj-tab ${tab === "siswa" ? "active" : ""}`} onClick={() => setTab("siswa")}>Kelola Siswa</button>}
          {isAdmin && <button className={`aj-tab ${tab === "guru" ? "active" : ""}`} onClick={() => setTab("guru")}>Kelola Guru</button>}
          {isAdmin && <button className={`aj-tab ${tab === "qr" ? "active" : ""}`} onClick={() => setTab("qr")}>Generate QR</button>}
        </div>

        {tab === "rekap" && <RekapAbsensi />}
        {tab === "jadwal" && isAdmin && <KelolaJadwal />}
        {tab === "siswa" && isAdmin && <KelolaSiswa />}
        {tab === "guru" && isAdmin && <KelolaGuru />}
        {tab === "qr" && isAdmin && <GenerateQR />}
      </div>
      <Footer />
    </div>
  );
}

/* ========================================================================
   APP / ROUTING
======================================================================== */
function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<Navigate to="/login-siswa" />} />
        <Route path="/login-siswa" element={<LoginSiswa />} />
        <Route path="/login-guru" element={<LoginGuru />} />
        <Route path="/siswa" element={<DashboardSiswa />} />
        <Route path="/siswa/scan/:jadwalId" element={<ScanQR />} />
        <Route path="/guru" element={<DashboardGuru />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
