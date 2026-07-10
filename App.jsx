import React, { useState, useEffect, useMemo, useCallback, useRef, createContext, useContext } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
} from "recharts";

/* ============================================================
   GANATRA CLINIC — API-connected build.
   Same interface as the offline prototype, but every read and
   write now goes through the Express/PostgreSQL backend.
   ============================================================ */

const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;700;800&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');";

const BLOCKS = [
  { name: "Building — Residential", rate: 5 },
  { name: "Building — Clinic / Commercial Premises", rate: 10 },
  { name: "Furniture & Fittings", rate: 10 },
  { name: "Medical & Diagnostic Equipment (Plant & Machinery)", rate: 15 },
  { name: "Motor Vehicles (own use)", rate: 15 },
  { name: "Computers, Billing Software & Peripherals", rate: 40 },
  { name: "Intangible Assets (software licences, know-how)", rate: 25 },
];
const EXPENSE_CATEGORIES = [
  "Nursing Staff Salary", "Electricity / Light Bill", "Housekeeping Expenses",
  "Rent", "Medicine Bills", "Repair & Maintenance", "Miscellaneous Expenses", "Staff Welfare",
];
const SHIFTS = ["Morning", "Evening"];
const PAY_TYPES = ["Daily", "Monthly"];
const COLLECTION_MODES = ["Cash", "UPI", "Card", "Other"];
const REFERRAL_TYPES = ["Lab Test", "Hospital"];
const COLORS = ["#0B4F4A", "#C9A227", "#1F8A5F", "#B3423A", "#5B6B69", "#7FB3AB", "#E8C468"];

const inr = (n) => "₹" + (Number(n) || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });
const todayISO = () => new Date().toISOString().slice(0, 10);
const d10 = (v) => (v ? String(v).slice(0, 10) : "");
function fyOf(dateStr) {
  const d = new Date(dateStr); const y = d.getFullYear(), m = d.getMonth() + 1;
  const startYear = m >= 4 ? y : y - 1;
  return `${startYear}-${String((startYear + 1) % 100).padStart(2, "0")}`;
}
function fyRange(fy) { const startYear = parseInt(fy.split("-")[0], 10); return { start: `${startYear}-04-01`, end: `${startYear + 1}-03-31` }; }
function last4FYs() {
  const cur = fyOf(todayISO()); const startYear = parseInt(cur.split("-")[0], 10);
  return [0, 1, 2, 3].map((i) => { const y = startYear - i; return `${y}-${String((y + 1) % 100).padStart(2, "0")}`; });
}

/* -------- API client -------- */
const ApiContext = createContext(null);
const useApi = () => useContext(ApiContext);

async function apiFetch(origin, token, path, { method = "GET", body, isForm = false } = {}) {
  if (!origin) throw new Error("Set the API server URL first.");
  const headers = {};
  if (token) headers.Authorization = `Bearer ${token}`;
  let payload;
  if (body !== undefined) {
    if (isForm) payload = body;
    else { headers["Content-Type"] = "application/json"; payload = JSON.stringify(body); }
  }
  let res;
  try {
    res = await fetch(`${origin}/api${path}`, { method, headers, body: payload });
  } catch {
    throw new Error(`Couldn't reach ${origin} — check the URL and that the server is running.`);
  }
  let data = null;
  if (res.status !== 204) { try { data = await res.json(); } catch { data = null; } }
  if (!res.ok) throw new Error((data && data.error) || `Request failed (${res.status})`);
  return data;
}

async function storeGet(key) { try { const r = await window.storage.get(key, false); return r ? JSON.parse(r.value) : null; } catch { return null; } }
async function storeSet(key, value) { try { await window.storage.set(key, JSON.stringify(value), false); } catch (e) { console.error(e); } }

function exportExcel(filename, sheetsObj) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheetsObj).forEach(([name, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ "No data": "—" }]);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, `${filename}.xlsx`);
}
function ExportRow({ onExcel }) {
  return (
    <div className="export-row no-print">
      <button className="btn secondary small" onClick={onExcel} type="button">⬇ Export Excel</button>
      <button className="btn secondary small" onClick={() => window.print()} type="button">⎙ Export PDF</button>
    </div>
  );
}
function ErrorNote({ msg }) { return msg ? <div className="err-note">{msg}</div> : null; }

/* -------- row → app-shape mappers (backend is snake_case) -------- */
const mapDoctor = (r) => ({ id: r.id, name: r.name, shift: r.shift, payType: r.pay_type, rate: Number(r.rate) });
const mapCase = (r) => ({
  id: r.id, caseNo: r.case_no, date: d10(r.case_date), patientName: r.patient_name, phone: r.phone,
  briefHistory: r.brief_history, doctorId: r.doctor_id, doctorName: r.doctor_name, shift: r.shift,
  externalPrescription: r.external_prescription, image: r.image_url,
  medicines: (r.medicines || []).map((m) => ({ name: m.medicine_name, qty: Number(m.qty), price: Number(m.unit_price) })),
});
const mapCollection = (r) => ({
  id: r.id, caseId: r.case_id, caseNo: r.case_no, patientName: r.patient_name, phone: r.phone,
  date: d10(r.collection_date), amountDue: Number(r.amount_due), amountCollected: Number(r.amount_collected),
  balance: Number(r.balance), mode: r.mode, image: r.image_url,
});
const mapDoctorPay = (r) => ({ id: r.id, date: d10(r.pay_date), doctorId: r.doctor_id, doctorName: r.doctor_name, amount: Number(r.amount) });
const mapReferral = (r) => ({ id: r.id, date: d10(r.referral_date), patientName: r.patient_name, referralType: r.referral_type, referredTo: r.referred_to, amount: Number(r.amount), notes: r.notes });
const mapGift = (r) => ({ id: r.id, date: d10(r.gift_date), repName: r.rep_name, company: r.company, gift: r.gift_description, doctorId: r.doctor_id, doctorName: r.doctor_name });
const mapExpense = (r) => ({ id: r.id, date: d10(r.expense_date), category: r.category, amount: Number(r.amount), narration: r.narration, image: r.image_url });
const mapAsset = (r) => ({ id: r.id, name: r.name, block: r.block, rate: Number(r.rate), purchaseDate: d10(r.purchase_date), cost: Number(r.cost) });
const mapCapital = (r) => ({ id: r.id, date: d10(r.txn_date), type: r.txn_type, amount: Number(r.amount), note: r.note });
const mapSettings = (r) => ({ clinicName: r.clinic_name || "Ganatra Clinic", proprietor: r.proprietor || "Dr. Bhavisha Pratik Ganatra", address: r.address || "", phone: r.phone || "" });

/* -------- Image capture: upload, rotate to align, attach via API -------- */
function ImageCapture({ value, onChange }) {
  const { call, origin } = useApi();
  const [raw, setRaw] = useState(null);
  const [rotation, setRotation] = useState(0);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const inputRef = useRef();

  const onFile = (e) => {
    const f = e.target.files[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = () => { setRaw(reader.result); setRotation(0); setErr(""); };
    reader.readAsDataURL(f); e.target.value = "";
  };

  const attach = () => {
    setErr("");
    const img = new Image();
    img.onload = () => {
      const swap = rotation % 180 !== 0;
      const maxDim = 1000;
      let w = img.width, h = img.height;
      const scale = Math.min(1, maxDim / Math.max(w, h));
      w = Math.round(w * scale); h = Math.round(h * scale);
      const canvas = document.createElement("canvas");
      canvas.width = swap ? h : w; canvas.height = swap ? w : h;
      const ctx = canvas.getContext("2d");
      ctx.translate(canvas.width / 2, canvas.height / 2);
      ctx.rotate((rotation * Math.PI) / 180);
      ctx.drawImage(img, -w / 2, -h / 2, w, h);
      canvas.toBlob(async (blob) => {
        setBusy(true);
        try {
          const fd = new FormData();
          fd.append("photo", blob, "stationery.jpg");
          const result = await call("/upload", { method: "POST", body: fd, isForm: true });
          onChange(result.url);
          setRaw(null); setRotation(0);
        } catch (e) { setErr(e.message); }
        setBusy(false);
      }, "image/jpeg", 0.75);
    };
    img.src = raw;
  };

  if (value && !raw) {
    return (
      <div className="imgcap">
        <img src={`${origin}${value}`} alt="attached stationery" className="imgcap-thumb" />
        <button className="btn secondary small" type="button" onClick={() => onChange(null)}>Remove image</button>
      </div>
    );
  }
  if (raw) {
    return (
      <div className="imgcap">
        <div className="imgcap-preview-wrap"><img src={raw} alt="preview" style={{ transform: `rotate(${rotation}deg)` }} className="imgcap-preview" /></div>
        <div className="imgcap-controls">
          <button className="btn secondary small" type="button" onClick={() => setRotation((r) => (r - 90 + 360) % 360)}>⟲ Rotate</button>
          <button className="btn secondary small" type="button" onClick={() => setRotation((r) => (r + 90) % 360)}>⟳ Rotate</button>
          <button className="btn small" type="button" disabled={busy} onClick={attach}>{busy ? "Uploading…" : "Attach to record"}</button>
          <button className="btn danger small" type="button" onClick={() => setRaw(null)}>Cancel</button>
        </div>
        <ErrorNote msg={err} />
      </div>
    );
  }
  return (
    <div className="imgcap">
      <input ref={inputRef} type="file" accept="image/*" capture="environment" onChange={onFile} className="imgcap-input" />
      <label className="btn secondary small imgcap-label" onClick={() => inputRef.current && inputRef.current.click()}>📎 Upload stationery photo</label>
    </div>
  );
}

/* ============================== AUTH ============================== */
function AuthScreen({ onLogin, origin, setOrigin }) {
  const [mode, setMode] = useState("login");
  const [name, setName] = useState("");
  const [userId, setUserId] = useState("");
  const [password, setPassword] = useState("");
  const [role, setRole] = useState("Reception");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const localCall = (path, opts) => apiFetch(origin, null, path, opts);

  const submit = async (e) => {
    e.preventDefault(); setError("");
    if (!origin.trim()) { setError("Enter your API server URL first."); return; }
    if (!userId.trim() || !password) { setError("Enter a user ID and password."); return; }
    setBusy(true);
    try {
      const path = mode === "register" ? "/auth/register" : "/auth/login";
      const body = mode === "register" ? { userId, password, name: name || userId, role } : { userId, password };
      const data = await localCall(path, { method: "POST", body });
      await storeSet("clinic:apiOrigin", origin);
      onLogin({ token: data.token, userId: data.user.userId, name: data.user.name, role: data.user.role });
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const demo = async () => {
    if (!origin.trim()) { setError("Enter your API server URL first."); return; }
    setError(""); setBusy(true);
    try {
      let data;
      try {
        data = await localCall("/auth/register", { method: "POST", body: { userId: "demo", password: "demo1234", name: "Front Desk", role: "Reception" } });
      } catch {
        data = await localCall("/auth/login", { method: "POST", body: { userId: "demo", password: "demo1234" } });
      }
      const token = data.token;
      const call2 = (p, o) => apiFetch(origin, token, p, o);
      const existingDoctors = await call2("/doctors");
      let doctors = existingDoctors;
      if (existingDoctors.length === 0) {
        const d1 = await call2("/doctors", { method: "POST", body: { name: "Dr. Bhavisha Pratik Ganatra", shift: "Morning", payType: "Monthly", rate: 60000 } });
        const d2 = await call2("/doctors", { method: "POST", body: { name: "Dr. Kunal Shah", shift: "Evening", payType: "Daily", rate: 1500 } });
        doctors = [d1, d2];
        const t = todayISO();
        const c1 = await call2("/cases", { method: "POST", body: { date: t, patientName: "Ramesh Patel", phone: "9825012345", briefHistory: "Fever, body ache — 3 days", doctorId: d1.id, shift: "Morning", externalPrescription: "Azithromycin 500 (from medical store)", medicines: [{ name: "Paracetamol 650", qty: 10, price: 2 }] } });
        const c2 = await call2("/cases", { method: "POST", body: { date: t, patientName: "Sunita Mehta", phone: "9825098765", briefHistory: "Routine antenatal checkup", doctorId: d2.id, shift: "Evening", medicines: [{ name: "Folic Acid", qty: 30, price: 1.5 }] } });
        await call2("/collections", { method: "POST", body: { caseId: c1.id, caseNo: c1.case_no, patientName: "Ramesh Patel", phone: "9825012345", date: t, amountDue: 500, amountCollected: 500, mode: "Cash" } });
        await call2("/collections", { method: "POST", body: { caseId: c2.id, caseNo: c2.case_no, patientName: "Sunita Mehta", phone: "9825098765", date: t, amountDue: 800, amountCollected: 500, mode: "UPI" } });
        await call2("/doctor-pays", { method: "POST", body: { doctorId: d2.id, date: t, amount: 1500 } });
        await call2("/referrals", { method: "POST", body: { date: t, patientName: "Ramesh Patel", referralType: "Lab Test", referredTo: "City Diagnostics", amount: 200, notes: "CBC + LFT" } });
        await call2("/gifts", { method: "POST", body: { date: t, repName: "Anil Kapoor", company: "MedLife Pharma", gift: "Diary set", doctorId: d1.id } });
        await call2("/expenses", { method: "POST", body: { date: t, category: "Rent", amount: 25000, narration: "Monthly clinic rent" } });
        await call2("/expenses", { method: "POST", body: { date: t, category: "Electricity / Light Bill", amount: 3200, narration: "Light bill" } });
        await call2("/expenses", { method: "POST", body: { date: t, category: "Nursing Staff Salary", amount: 15000, narration: "Nursing staff salary" } });
        await call2("/assets", { method: "POST", body: { name: "Digital BP monitor & diagnostic set", block: "Medical & Diagnostic Equipment (Plant & Machinery)", rate: 15, purchaseDate: fyRange(fyOf(t)).start, cost: 65000 } });
        await call2("/assets", { method: "POST", body: { name: "Reception computer & billing software", block: "Computers, Billing Software & Peripherals", rate: 40, purchaseDate: fyRange(fyOf(t)).start, cost: 55000 } });
        await call2("/capital", { method: "POST", body: { date: fyRange(fyOf(t)).start, type: "Introduced", amount: 300000, note: "Opening capital" } });
      }
      await storeSet("clinic:apiOrigin", origin);
      onLogin({ token, userId: data.user.userId, name: data.user.name, role: data.user.role });
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  return (
    <div className="auth-wrap">
      <style>{`
        ${FONT_IMPORT}
        .auth-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;background:radial-gradient(circle at 25% 15%,#123f3a 0%,#081d1b 65%);font-family:'Inter',sans-serif;padding:24px;}
        .auth-card{width:100%;max-width:420px;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 30px 70px rgba(0,0,0,.45);}
        .auth-header{background:linear-gradient(135deg,#0B4F4A,#082F2C);padding:30px 26px 22px;text-align:center;}
        .auth-header .brand{font-family:'Plus Jakarta Sans',sans-serif;font-weight:800;color:#fff;font-size:26px;letter-spacing:.5px;margin-top:6px;}
        .auth-header .sub{color:#B9D8D2;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin-top:4px;}
        .pulse-path{stroke:#C9A227;stroke-width:2.5;fill:none;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:340;stroke-dashoffset:340;animation:draw 1.4s ease forwards .2s;}
        @keyframes draw{to{stroke-dashoffset:0;}}
        @media (prefers-reduced-motion:reduce){.pulse-path{animation:none;stroke-dashoffset:0;}}
        .auth-body{padding:26px 26px 24px;}
        .tabs{display:flex;gap:6px;margin-bottom:18px;}
        .tab-btn{flex:1;padding:8px 0;font-size:12.5px;font-weight:700;border:1px solid #E1E8E6;background:#F5F8F7;color:#5B6B69;border-radius:6px;cursor:pointer;}
        .tab-btn.active{background:#0B4F4A;color:#fff;border-color:#0B4F4A;}
        .field{margin-bottom:14px;}
        .field label{display:block;font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#5B6B69;font-weight:700;margin-bottom:5px;}
        .field input,.field select{width:100%;border:1.5px solid #E1E8E6;border-radius:8px;padding:10px 11px;font-size:14.5px;font-family:'IBM Plex Mono',monospace;outline:none;box-sizing:border-box;}
        .field input:focus,.field select:focus{border-color:#0B4F4A;}
        .err{color:#B3423A;font-size:12.5px;margin-bottom:10px;}
        .submit-btn{width:100%;background:#C9A227;color:#2A2103;font-weight:700;border:none;padding:12px;border-radius:8px;font-size:14.5px;cursor:pointer;box-shadow:0 4px 0 #96791b;}
        .submit-btn:active{transform:translateY(2px);box-shadow:0 2px 0 #96791b;}
        .demo-link{display:block;width:100%;text-align:center;margin-top:14px;font-size:12.5px;color:#5B6B69;text-decoration:underline;background:none;border:none;cursor:pointer;}
        .note{font-size:11px;color:#8a9a97;margin-top:12px;line-height:1.6;text-align:center;}
        .api-field{background:#F5F8F7;border:1px dashed #C9A227;border-radius:8px;padding:10px 12px;margin-bottom:16px;}
        .api-field label{display:block;font-size:10.5px;text-transform:uppercase;letter-spacing:1px;color:#8a6d0a;font-weight:700;margin-bottom:5px;}
        .api-field input{width:100%;border:1px solid #E1E8E6;border-radius:6px;padding:8px 9px;font-size:13px;font-family:'IBM Plex Mono',monospace;box-sizing:border-box;}
      `}</style>
      <div className="auth-card">
        <div className="auth-header">
          <svg viewBox="0 0 200 40" width="180" height="36" style={{ margin: "0 auto", display: "block" }}>
            <path className="pulse-path" d="M0,20 L55,20 L65,4 L75,36 L85,20 L200,20" />
          </svg>
          <div className="brand">GANATRA CLINIC</div>
          <div className="sub">Practice &amp; Accounts Manager</div>
        </div>
        <div className="auth-body">
          <div className="api-field">
            <label>API server URL</label>
            <input value={origin} onChange={(e) => setOrigin(e.target.value)} placeholder="https://your-backend.onrender.com" />
          </div>
          <div className="tabs">
            <button className={"tab-btn" + (mode === "login" ? " active" : "")} onClick={() => setMode("login")} type="button">Log in</button>
            <button className={"tab-btn" + (mode === "register" ? " active" : "")} onClick={() => setMode("register")} type="button">Register</button>
          </div>
          <form onSubmit={submit}>
            {mode === "register" && (
              <>
                <div className="field"><label>Your name</label><input value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Front desk / Dr. Ganatra" /></div>
                <div className="field"><label>Role</label><select value={role} onChange={(e) => setRole(e.target.value)}>{["Doctor", "Reception", "Nurse", "Admin"].map((r) => <option key={r} value={r}>{r}</option>)}</select></div>
              </>
            )}
            <div className="field"><label>User ID</label><input value={userId} onChange={(e) => setUserId(e.target.value)} autoCapitalize="none" placeholder="your-user-id" /></div>
            <div className="field"><label>Password</label><input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" /></div>
            {error && <div className="err">{error}</div>}
            <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Please wait…" : mode === "login" ? "Log in" : "Create account"}</button>
          </form>
          <button className="demo-link" onClick={demo} type="button" disabled={busy}>Skip — seed &amp; explore sample clinic data (user: demo)</button>
          <div className="note">Same user ID and password sign in on web and mobile — both talk to the same API server. Staff accounts and clinic records are shared across everyone who logs in, not siloed per device.</div>
        </div>
      </div>
    </div>
  );
}

/* ============================== NAV / SHELL ============================== */
const NAV = [
  { key: "dashboard", label: "Dashboard" }, { key: "cases", label: "Case Records" },
  { key: "collections", label: "Collections" }, { key: "doctors", label: "Doctor Shifts & Pay" },
  { key: "referrals", label: "Referral Income" }, { key: "gifts", label: "Gifts Register" },
  { key: "expenses", label: "Expenses" }, { key: "assets", label: "Fixed Assets" },
  { key: "statements", label: "Financial Statements" }, { key: "settings", label: "Settings" },
];

export default function App() {
  const [origin, setOrigin] = useState("");
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [view, setView] = useState("dashboard");
  const [fy, setFy] = useState(fyOf(todayISO()));

  const [settings, setSettings] = useState({ clinicName: "Ganatra Clinic", proprietor: "Dr. Bhavisha Pratik Ganatra", address: "", phone: "" });
  const [doctors, setDoctors] = useState([]);
  const [cases, setCases] = useState([]);
  const [collections, setCollections] = useState([]);
  const [doctorPays, setDoctorPays] = useState([]);
  const [referrals, setReferrals] = useState([]);
  const [gifts, setGifts] = useState([]);
  const [expenses, setExpenses] = useState([]);
  const [assets, setAssets] = useState([]);
  const [capital, setCapital] = useState([]);

  useEffect(() => { storeGet("clinic:apiOrigin").then((o) => { if (o) setOrigin(o); }); }, []);

  const call = useCallback((path, opts) => apiFetch(origin, session?.token, path, opts), [origin, session]);
  const apiValue = useMemo(() => ({ call, origin, token: session?.token }), [call, origin, session]);

  const reloadAll = useCallback(async () => {
    setLoading(true); setLoadError("");
    try {
      const [s, d, c, col, dp, ref, g, exp, ast, cap] = await Promise.all([
        call("/settings"), call("/doctors"), call("/cases"), call("/collections"),
        call("/doctor-pays"), call("/referrals"), call("/gifts"), call("/expenses"),
        call("/assets"), call("/capital"),
      ]);
      setSettings(mapSettings(s || {})); setDoctors((d || []).map(mapDoctor)); setCases((c || []).map(mapCase));
      setCollections((col || []).map(mapCollection)); setDoctorPays((dp || []).map(mapDoctorPay));
      setReferrals((ref || []).map(mapReferral)); setGifts((g || []).map(mapGift));
      setExpenses((exp || []).map(mapExpense)); setAssets((ast || []).map(mapAsset)); setCapital((cap || []).map(mapCapital));
    } catch (e) { setLoadError(e.message); }
    setLoading(false);
  }, [call]);

  useEffect(() => { if (session) reloadAll(); }, [session]); // eslint-disable-line

  // create/delete helpers — call the API, then patch local state
  const addDoctor = useCallback(async (body) => { const r = await call("/doctors", { method: "POST", body }); setDoctors((p) => [...p, mapDoctor(r)]); }, [call]);
  const removeDoctor = useCallback(async (id) => { await call(`/doctors/${id}`, { method: "DELETE" }); setDoctors((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addCase = useCallback(async (body) => { const r = await call("/cases", { method: "POST", body }); const doc = doctors.find((d) => d.id === body.doctorId); setCases((p) => [{ ...mapCase(r), doctorName: doc?.name }, ...p]); return r; }, [call, doctors]);
  const removeCase = useCallback(async (id) => { await call(`/cases/${id}`, { method: "DELETE" }); setCases((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addCollection = useCallback(async (body) => { const r = await call("/collections", { method: "POST", body }); setCollections((p) => [mapCollection(r), ...p]); }, [call]);
  const removeCollection = useCallback(async (id) => { await call(`/collections/${id}`, { method: "DELETE" }); setCollections((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addDoctorPay = useCallback(async (body) => { const r = await call("/doctor-pays", { method: "POST", body }); const doc = doctors.find((d) => d.id === body.doctorId); setDoctorPays((p) => [{ ...mapDoctorPay(r), doctorName: doc?.name }, ...p]); }, [call, doctors]);
  const removeDoctorPay = useCallback(async (id) => { await call(`/doctor-pays/${id}`, { method: "DELETE" }); setDoctorPays((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addReferral = useCallback(async (body) => { const r = await call("/referrals", { method: "POST", body }); setReferrals((p) => [mapReferral(r), ...p]); }, [call]);
  const removeReferral = useCallback(async (id) => { await call(`/referrals/${id}`, { method: "DELETE" }); setReferrals((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addGift = useCallback(async (body) => { const r = await call("/gifts", { method: "POST", body }); const doc = doctors.find((d) => d.id === body.doctorId); setGifts((p) => [{ ...mapGift(r), doctorName: doc?.name }, ...p]); }, [call, doctors]);
  const removeGift = useCallback(async (id) => { await call(`/gifts/${id}`, { method: "DELETE" }); setGifts((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addExpense = useCallback(async (body) => { const r = await call("/expenses", { method: "POST", body }); setExpenses((p) => [mapExpense(r), ...p]); }, [call]);
  const removeExpense = useCallback(async (id) => { await call(`/expenses/${id}`, { method: "DELETE" }); setExpenses((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addAsset = useCallback(async (body) => { const r = await call("/assets", { method: "POST", body }); setAssets((p) => [mapAsset(r), ...p]); }, [call]);
  const removeAsset = useCallback(async (id) => { await call(`/assets/${id}`, { method: "DELETE" }); setAssets((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addCapital = useCallback(async (body) => { const r = await call("/capital", { method: "POST", body }); setCapital((p) => [mapCapital(r), ...p]); }, [call]);
  const removeCapital = useCallback(async (id) => { await call(`/capital/${id}`, { method: "DELETE" }); setCapital((p) => p.filter((x) => x.id !== id)); }, [call]);
  const updateSettings = useCallback(async (body) => { const r = await call("/settings", { method: "PUT", body }); setSettings(mapSettings(r)); }, [call]);

  if (!session) return <AuthScreen onLogin={setSession} origin={origin} setOrigin={setOrigin} />;

  return (
    <ApiContext.Provider value={apiValue}>
      <div className="app-root">
        <style>{`
          ${FONT_IMPORT}
          .app-root{--primary:#0B4F4A;--primary-dark:#082F2C;--accent:#C9A227;--accent-soft:#F3E3A8;--bg:#F5F8F7;--surface:#FFFFFF;
            --ink:#142524;--ink-soft:#5B6B69;--border:#E1E8E6;--income:#1F8A5F;--expense:#B3423A;
            min-height:100vh;background:var(--bg);font-family:'Inter',sans-serif;color:var(--ink);display:flex;}
          .sidebar{width:230px;background:linear-gradient(180deg,var(--primary),var(--primary-dark));color:#EAF3F1;flex-shrink:0;padding:22px 0;display:flex;flex-direction:column;}
          .sidebar .brand{font-family:'Plus Jakarta Sans',sans-serif;font-weight:800;font-size:20px;padding:0 20px 2px;}
          .sidebar .biz{padding:0 20px 16px;font-size:11.5px;color:#B9D8D2;border-bottom:1px solid rgba(255,255,255,.15);margin-bottom:8px;line-height:1.5;word-break:break-all;}
          .nav-item{text-align:left;background:none;border:none;color:#EAF3F1;padding:10px 20px;font-size:13.5px;font-weight:500;cursor:pointer;border-left:3px solid transparent;opacity:.82;}
          .nav-item.active{background:rgba(0,0,0,.22);border-left-color:var(--accent);opacity:1;font-weight:700;}
          .nav-item:hover{opacity:1;}
          .logout{margin-top:auto;padding:12px 20px;font-size:12px;color:#E9CFCF;background:none;border:none;text-align:left;cursor:pointer;text-decoration:underline;}
          .main{flex:1;min-width:0;}
          .topbar{background:var(--surface);border-bottom:2px solid var(--accent);padding:14px 26px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;}
          .topbar h1{font-family:'Plus Jakarta Sans',sans-serif;font-size:19px;margin:0;}
          .fy-select{font-family:'IBM Plex Mono',monospace;background:#fff;border:1px solid var(--border);border-radius:6px;padding:6px 10px;font-size:13px;}
          .content{padding:22px 26px 60px;max-width:1150px;}
          .card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:18px 20px;margin-bottom:18px;box-shadow:0 1px 4px rgba(10,40,36,.05);}
          .card h2{font-family:'Plus Jakarta Sans',sans-serif;font-size:16.5px;margin:0 0 12px;color:var(--primary-dark);}
          .grid-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin-bottom:20px;}
          .stat{background:var(--surface);border:1px solid var(--border);border-left:4px solid var(--accent);border-radius:10px;padding:14px 16px;}
          .stat .label{font-size:10.5px;text-transform:uppercase;letter-spacing:.8px;color:var(--ink-soft);font-weight:700;}
          .stat .value{font-family:'IBM Plex Mono',monospace;font-size:20px;font-weight:600;margin-top:5px;}
          table{width:100%;border-collapse:collapse;font-size:13px;}
          th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;color:var(--ink-soft);border-bottom:2px solid var(--accent);padding:6px 8px;}
          td{padding:7px 8px;border-bottom:1px solid var(--border);font-family:'IBM Plex Mono',monospace;vertical-align:top;}
          tr:last-child td{border-bottom:none;}
          .num{text-align:right;}
          .btn{background:var(--primary);color:#fff;border:none;padding:9px 16px;border-radius:7px;font-size:13.5px;font-weight:600;cursor:pointer;}
          .btn.secondary{background:transparent;color:var(--primary);border:1px solid var(--primary);}
          .btn.small{padding:5px 10px;font-size:12px;}
          .btn.danger{background:var(--expense);}
          .btn:disabled{opacity:.55;cursor:default;}
          select,input[type=text],input[type=date],input[type=number],input[type=tel]{font-family:'Inter',sans-serif;border:1px solid var(--border);border-radius:6px;padding:7px 9px;font-size:13.5px;}
          .form-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:12px;}
          .form-grid label{display:block;font-size:10.5px;font-weight:700;color:var(--ink-soft);margin-bottom:4px;text-transform:uppercase;letter-spacing:.4px;}
          .empty{color:var(--ink-soft);font-size:13px;padding:10px 0;}
          .pill{display:inline-block;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:700;}
          .pill.morning{background:#FFF2CE;color:#8a6d0a;} .pill.evening{background:#E4DCF7;color:#5b3fa3;}
          .balance-tag{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;padding:4px 10px;border-radius:6px;display:inline-block;}
          .balance-tag.ok{background:#e3efe6;color:var(--income);} .balance-tag.bad{background:#f6e4e4;color:var(--expense);}
          .export-row{display:flex;gap:8px;margin-top:12px;}
          .imgcap{margin-top:4px;} .imgcap-thumb{max-width:160px;max-height:120px;border-radius:8px;border:1px solid var(--border);display:block;margin-bottom:6px;}
          .imgcap-preview-wrap{width:100%;max-width:260px;height:170px;overflow:hidden;border:1px dashed var(--border);border-radius:8px;display:flex;align-items:center;justify-content:center;margin-bottom:8px;background:#fafcfb;}
          .imgcap-preview{max-width:90%;max-height:90%;transition:transform .15s;}
          .imgcap-controls{display:flex;gap:6px;flex-wrap:wrap;} .imgcap-input{display:none;} .imgcap-label{cursor:pointer;display:inline-block;}
          .note-box{font-size:12px;color:var(--ink-soft);line-height:1.6;background:#F5F8F7;border-left:3px solid var(--accent);padding:10px 12px;border-radius:6px;margin-top:10px;}
          .err-note{font-size:12px;color:var(--expense);background:#f6e4e4;padding:8px 10px;border-radius:6px;margin-top:8px;}
          @media (max-width:820px){
            .app-root{flex-direction:column;} .sidebar{width:100%;flex-direction:row;flex-wrap:wrap;padding:14px 12px;align-items:center;}
            .sidebar .brand{padding:0 10px 0 0;font-size:17px;} .sidebar .biz{display:none;}
            .nav-item{padding:7px 9px;font-size:11.5px;border-left:none;border-bottom:3px solid transparent;}
            .nav-item.active{border-left:none;border-bottom-color:var(--accent);} .logout{margin-top:0;margin-left:auto;padding:8px 10px;}
            .content{padding:16px 12px 50px;}
          }
          @media print{ .sidebar,.topbar .fy-select,.no-print{display:none !important;} .content{padding:0;max-width:100%;} body{background:#fff;} .card{box-shadow:none;border:1px solid #ccc;} }
        `}</style>

        <nav className="sidebar">
          <div className="brand">GANATRA CLINIC</div>
          <div className="biz">{settings.proprietor}<br />{session.name} ({session.userId}) · {session.role}<br /><span style={{ opacity: .7 }}>{origin}</span></div>
          {NAV.map((n) => (<button key={n.key} className={"nav-item" + (view === n.key ? " active" : "")} onClick={() => setView(n.key)}>{n.label}</button>))}
          <button className="logout" onClick={() => setSession(null)}>Log out</button>
        </nav>

        <div className="main">
          <div className="topbar">
            <h1>{NAV.find((n) => n.key === view)?.label}</h1>
            <select className="fy-select no-print" value={fy} onChange={(e) => setFy(e.target.value)}>{last4FYs().map((f) => <option key={f} value={f}>FY {f}</option>)}</select>
          </div>
          <div className="content">
            {loading ? <div className="empty">Loading clinic records from the server…</div> : loadError ? (
              <div className="card"><h2>Couldn't load your data</h2><ErrorNote msg={loadError} /><button className="btn" style={{ marginTop: 10 }} onClick={reloadAll} type="button">Retry</button></div>
            ) : (
              <>
                {view === "dashboard" && <Dashboard settings={settings} collections={collections} referrals={referrals} expenses={expenses} doctorPays={doctorPays} fy={fy} />}
                {view === "cases" && <CaseRecords cases={cases} addCase={addCase} removeCase={removeCase} doctors={doctors} />}
                {view === "collections" && <Collections collections={collections} addCollection={addCollection} removeCollection={removeCollection} cases={cases} fy={fy} />}
                {view === "doctors" && <DoctorShifts doctors={doctors} addDoctor={addDoctor} removeDoctor={removeDoctor} doctorPays={doctorPays} addDoctorPay={addDoctorPay} removeDoctorPay={removeDoctorPay} />}
                {view === "referrals" && <Referrals referrals={referrals} addReferral={addReferral} removeReferral={removeReferral} fy={fy} />}
                {view === "gifts" && <Gifts gifts={gifts} addGift={addGift} removeGift={removeGift} doctors={doctors} />}
                {view === "expenses" && <Expenses expenses={expenses} addExpense={addExpense} removeExpense={removeExpense} fy={fy} />}
                {view === "assets" && <FixedAssets assets={assets} addAsset={addAsset} removeAsset={removeAsset} fy={fy} />}
                {view === "statements" && <FinancialStatements fy={fy} settings={settings} />}
                {view === "settings" && <SettingsPage settings={settings} updateSettings={updateSettings} session={session} origin={origin} capital={capital} addCapital={addCapital} removeCapital={removeCapital} />}
              </>
            )}
          </div>
        </div>
      </div>
    </ApiContext.Provider>
  );
}

/* ============================== DASHBOARD ============================== */
function Dashboard({ settings, collections, referrals, expenses, doctorPays, fy }) {
  const { call } = useApi();
  const [income, setIncome] = useState(null);
  useEffect(() => { call(`/statements/income?fy=${fy}`).then(setIncome).catch(() => setIncome(null)); }, [call, fy]);

  const t = todayISO();
  const sumDue = (arr, s, e) => arr.filter((c) => c.date >= s && c.date <= e).reduce((a, c) => a + Number(c.amountDue || 0), 0);
  const weekStart = new Date(); weekStart.setDate(weekStart.getDate() - 6);
  const monthStart = t.slice(0, 8) + "01";
  const today = sumDue(collections, t, t);
  const week = sumDue(collections, weekStart.toISOString().slice(0, 10), t);
  const month = sumDue(collections, monthStart, t);
  const outstanding = collections.reduce((a, c) => a + Number(c.balance || 0), 0);
  const netProfit = income ? income.netProfit : null;

  const last30 = [];
  for (let i = 29; i >= 0; i--) { const dt = new Date(); dt.setDate(dt.getDate() - i); const iso = dt.toISOString().slice(0, 10); last30.push({ date: iso.slice(5), amount: sumDue(collections, iso, iso) }); }
  const range = fyRange(fy);
  const expenseByCat = {};
  expenses.filter((e) => e.date >= range.start && e.date <= range.end).forEach((e) => { expenseByCat[e.category] = (expenseByCat[e.category] || 0) + Number(e.amount); });
  const pieData = Object.entries(expenseByCat).map(([name, value]) => ({ name, value }));

  return (
    <div>
      <div className="grid-cards">
        <div className="stat"><div className="label">Today's Collection</div><div className="value">{inr(today)}</div></div>
        <div className="stat"><div className="label">Last 7 Days</div><div className="value">{inr(week)}</div></div>
        <div className="stat"><div className="label">This Month</div><div className="value">{inr(month)}</div></div>
        <div className="stat"><div className="label">Outstanding Dues</div><div className="value">{inr(outstanding)}</div></div>
        <div className="stat" style={{ borderLeftColor: (netProfit ?? 0) >= 0 ? "#1F8A5F" : "#B3423A" }}>
          <div className="label">Net Profit (FY {fy})</div>
          <div className="value" style={{ color: (netProfit ?? 0) >= 0 ? "#1F8A5F" : "#B3423A" }}>{netProfit === null ? "…" : inr(netProfit)}</div>
        </div>
      </div>
      <div className="card">
        <h2>Collection trend — last 30 days</h2>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={last30}>
            <defs><linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#0B4F4A" stopOpacity={0.35} /><stop offset="100%" stopColor="#0B4F4A" stopOpacity={0.02} /></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#E1E8E6" /><XAxis dataKey="date" fontSize={11} /><YAxis fontSize={11} />
            <Tooltip formatter={(v) => inr(v)} /><Area type="monotone" dataKey="amount" stroke="#0B4F4A" fill="url(#cg)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
      <div className="card">
        <h2>Expense breakdown — FY {fy}</h2>
        {pieData.length === 0 ? <div className="empty">No expenses logged yet for this financial year.</div> : (
          <ResponsiveContainer width="100%" height={240}>
            <PieChart><Pie data={pieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>{pieData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}</Pie><Legend /><Tooltip formatter={(v) => inr(v)} /></PieChart>
          </ResponsiveContainer>
        )}
      </div>
    </div>
  );
}

/* ============================== CASE RECORDS ============================== */
function CaseRecords({ cases, addCase, removeCase, doctors }) {
  const blank = { date: todayISO(), patientName: "", phone: "", briefHistory: "", doctorId: doctors[0]?.id || "", shift: "Morning", externalPrescription: "", image: null };
  const [form, setForm] = useState(blank);
  const [meds, setMeds] = useState([{ name: "", qty: "", price: "" }]);
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);

  const addMedRow = () => setMeds([...meds, { name: "", qty: "", price: "" }]);
  const updMed = (i, f, v) => { const n = [...meds]; n[i] = { ...n[i], [f]: v }; setMeds(n); };
  const rmMed = (i) => setMeds(meds.filter((_, idx) => idx !== i));

  const save = async () => {
    setErr("");
    if (!form.patientName.trim()) { setErr("Enter the patient's name."); return; }
    const medicines = meds.filter((m) => m.name.trim()).map((m) => ({ name: m.name, qty: Number(m.qty) || 0, price: Number(m.price) || 0 }));
    setBusy(true);
    try {
      await addCase({ date: form.date, patientName: form.patientName, phone: form.phone, briefHistory: form.briefHistory, doctorId: form.doctorId || null, shift: form.shift, externalPrescription: form.externalPrescription, imageUrl: form.image, medicines });
      setForm(blank); setMeds([{ name: "", qty: "", price: "" }]);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeCase(id); } catch (e) { setErr(e.message); } };
  const medValue = (c) => (c.medicines || []).reduce((s, m) => s + m.qty * m.price, 0);

  return (
    <div>
      <div className="card">
        <h2>New case paper entry</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Patient name</label><input type="text" value={form.patientName} onChange={(e) => setForm({ ...form, patientName: e.target.value })} /></div>
          <div><label>Phone</label><input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></div>
          <div><label>Doctor</label><select value={form.doctorId} onChange={(e) => setForm({ ...form, doctorId: e.target.value })}>
            {doctors.length === 0 && <option value="">Add a doctor first</option>}
            {doctors.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
          </select></div>
          <div><label>Shift</label><select value={form.shift} onChange={(e) => setForm({ ...form, shift: e.target.value })}>{SHIFTS.map((s) => <option key={s} value={s}>{s}</option>)}</select></div>
        </div>
        <div className="form-grid">
          <div style={{ gridColumn: "span 2" }}><label>Brief medical history</label><input type="text" style={{ width: "100%" }} value={form.briefHistory} onChange={(e) => setForm({ ...form, briefHistory: e.target.value })} placeholder="e.g. Fever, body ache — 3 days" /></div>
          <div style={{ gridColumn: "span 2" }}><label>Prescribed — buy from medical store</label><input type="text" style={{ width: "100%" }} value={form.externalPrescription} onChange={(e) => setForm({ ...form, externalPrescription: e.target.value })} /></div>
        </div>
        <label style={{ display: "block", fontSize: 10.5, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", marginBottom: 6 }}>Medicines dispensed loose (clinical record only)</label>
        {meds.map((m, i) => (
          <div key={i} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr auto", gap: 8, marginBottom: 6 }}>
            <input type="text" placeholder="Medicine name" value={m.name} onChange={(e) => updMed(i, "name", e.target.value)} />
            <input type="number" placeholder="Qty" value={m.qty} onChange={(e) => updMed(i, "qty", e.target.value)} />
            <input type="number" placeholder="Indicative unit price" value={m.price} onChange={(e) => updMed(i, "price", e.target.value)} />
            <button className="btn danger small" type="button" onClick={() => rmMed(i)}>✕</button>
          </div>
        ))}
        <button className="btn secondary small" type="button" onClick={addMedRow}>+ Add medicine</button>
        <div style={{ marginTop: 14 }}>
          <label style={{ display: "block", fontSize: 10.5, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", marginBottom: 6 }}>Case paper photo</label>
          <ImageCapture value={form.image} onChange={(img) => setForm({ ...form, image: img })} />
        </div>
        <button className="btn" style={{ marginTop: 16 }} type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : "Save case record"}</button>
        <ErrorNote msg={err} />
        <div className="note-box">Loose-medicine values are for clinical/inventory reference only — not posted as a separate expense, since "Medicine Bills" already covers what the clinic pays its supplier.</div>
      </div>
      <div className="card">
        <h2>Case register</h2>
        {cases.length === 0 ? <div className="empty">No case papers recorded yet.</div> : (
          <table>
            <thead><tr><th>Case No.</th><th>Date</th><th>Patient</th><th>Doctor</th><th>Shift</th><th>History</th><th className="num">Meds Value</th><th>Photo</th><th></th></tr></thead>
            <tbody>{[...cases].sort((a, b) => (a.date < b.date ? 1 : -1)).map((c) => (
              <tr key={c.id}><td>{c.caseNo}</td><td>{c.date}</td><td>{c.patientName}</td><td>{c.doctorName || "—"}</td>
                <td><span className={"pill " + (c.shift || "").toLowerCase()}>{c.shift}</span></td>
                <td style={{ maxWidth: 180 }}>{c.briefHistory}</td><td className="num">{inr(medValue(c))}</td>
                <td>{c.image ? "📎" : "—"}</td><td><button className="btn danger small" type="button" onClick={() => remove(c.id)}>Delete</button></td></tr>
            ))}</tbody>
          </table>
        )}
        <ExportRow onExcel={() => exportExcel("case-records", { Cases: cases.map((c) => ({ CaseNo: c.caseNo, Date: c.date, Patient: c.patientName, Phone: c.phone, Doctor: c.doctorName, Shift: c.shift, History: c.briefHistory, ExternalPrescription: c.externalPrescription, MedicinesDispensedValue: medValue(c) })) })} />
      </div>
    </div>
  );
}

/* ============================== COLLECTIONS ============================== */
function Collections({ collections, addCollection, removeCollection, cases, fy }) {
  const { call } = useApi();
  const blank = { caseId: "", caseNo: "", patientName: "", phone: "", date: todayISO(), amountDue: "", amountCollected: "", mode: "Cash", image: null };
  const [form, setForm] = useState(blank);
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [rollups, setRollups] = useState({ daily: [], weekly: [], monthly: [] });
  const range = fyRange(fy);

  useEffect(() => {
    const q = `from=${range.start}&to=${range.end}`;
    Promise.all(["daily", "weekly", "monthly"].map((p) => call(`/collections/rollup?period=${p}&${q}`)))
      .then(([d, w, m]) => setRollups({
        daily: d.map((r) => ({ period: d10(r.period), due: Number(r.due), collected: Number(r.collected) })),
        weekly: w.map((r) => ({ period: d10(r.period), due: Number(r.due), collected: Number(r.collected) })),
        monthly: m.map((r) => ({ period: String(r.period).slice(0, 7), due: Number(r.due), collected: Number(r.collected) })),
      })).catch(() => {});
  }, [call, fy]); // eslint-disable-line

  const pickCase = (caseId) => { const c = cases.find((x) => x.id === caseId); setForm({ ...form, caseId, caseNo: c ? c.caseNo : "", patientName: c ? c.patientName : form.patientName, phone: c ? c.phone : form.phone }); };
  const save = async () => {
    setErr(""); if (!form.patientName.trim() || form.amountDue === "") { setErr("Enter patient name and amount due."); return; }
    setBusy(true);
    try {
      await addCollection({ caseId: form.caseId || null, caseNo: form.caseNo || null, patientName: form.patientName, phone: form.phone, date: form.date, amountDue: Number(form.amountDue), amountCollected: Number(form.amountCollected) || 0, mode: form.mode, imageUrl: form.image });
      setForm(blank);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeCollection(id); } catch (e) { setErr(e.message); } };

  return (
    <div>
      <div className="card">
        <h2>New collection entry</h2>
        <div className="form-grid">
          <div><label>Linked case no.</label><select value={form.caseId} onChange={(e) => pickCase(e.target.value)}><option value="">— Not linked —</option>{cases.map((c) => <option key={c.id} value={c.id}>{c.caseNo} — {c.patientName}</option>)}</select></div>
          <div><label>Patient name</label><input type="text" value={form.patientName} onChange={(e) => setForm({ ...form, patientName: e.target.value })} /></div>
          <div><label>Phone</label><input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></div>
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Amount due (₹)</label><input type="number" value={form.amountDue} onChange={(e) => setForm({ ...form, amountDue: e.target.value })} /></div>
          <div><label>Amount collected (₹)</label><input type="number" value={form.amountCollected} onChange={(e) => setForm({ ...form, amountCollected: e.target.value })} /></div>
          <div><label>Mode</label><select value={form.mode} onChange={(e) => setForm({ ...form, mode: e.target.value })}>{COLLECTION_MODES.map((m) => <option key={m} value={m}>{m}</option>)}</select></div>
        </div>
        <ImageCapture value={form.image} onChange={(img) => setForm({ ...form, image: img })} />
        <button className="btn" style={{ marginTop: 14 }} type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : "Save collection entry"}</button>
        <ErrorNote msg={err} />
      </div>
      <div className="card">
        <h2>Collection register</h2>
        {collections.length === 0 ? <div className="empty">No collections recorded yet.</div> : (
          <table>
            <thead><tr><th>Date</th><th>Case No.</th><th>Patient</th><th>Phone</th><th className="num">Due</th><th className="num">Collected</th><th className="num">Balance</th><th>Mode</th><th></th></tr></thead>
            <tbody>{[...collections].sort((a, b) => (a.date < b.date ? 1 : -1)).map((c) => (
              <tr key={c.id}><td>{c.date}</td><td>{c.caseNo || "—"}</td><td>{c.patientName}</td><td>{c.phone}</td>
                <td className="num">{inr(c.amountDue)}</td><td className="num">{inr(c.amountCollected)}</td>
                <td className="num" style={{ color: c.balance > 0 ? "var(--expense)" : "var(--income)" }}>{inr(c.balance)}</td>
                <td>{c.mode}</td><td><button className="btn danger small" type="button" onClick={() => remove(c.id)}>Delete</button></td></tr>
            ))}</tbody>
          </table>
        )}
        <ExportRow onExcel={() => exportExcel("collections", { Collections: collections, Daily: rollups.daily, Weekly: rollups.weekly, Monthly: rollups.monthly })} />
      </div>
      <div className="card">
        <h2>Daily / weekly / monthly collection — FY {fy}</h2>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(300px,1fr))", gap: 16 }}>
          <RollupTable title="Daily" rows={rollups.daily.slice(0, 10)} />
          <RollupTable title="Weekly" rows={rollups.weekly.slice(0, 8)} />
          <RollupTable title="Monthly" rows={rollups.monthly} />
        </div>
      </div>
    </div>
  );
}
function RollupTable({ title, rows }) {
  return (
    <div>
      <div style={{ fontWeight: 700, fontSize: 12.5, marginBottom: 6, color: "var(--primary-dark)" }}>{title}</div>
      {rows.length === 0 ? <div className="empty">No data.</div> : (
        <table><thead><tr><th>Period</th><th className="num">Due</th><th className="num">Collected</th></tr></thead>
          <tbody>{rows.map((r, i) => (<tr key={i}><td>{r.period}</td><td className="num">{inr(r.due)}</td><td className="num">{inr(r.collected)}</td></tr>))}</tbody>
        </table>
      )}
    </div>
  );
}

/* ============================== DOCTOR SHIFTS & PAY ============================== */
function DoctorShifts({ doctors, addDoctor, removeDoctor, doctorPays, addDoctorPay, removeDoctorPay }) {
  const { call } = useApi();
  const [d, setD] = useState({ name: "", shift: "Morning", payType: "Daily", rate: "" });
  const [p, setP] = useState({ date: todayISO(), doctorId: "", amount: "" });
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [last14, setLast14] = useState([]);

  useEffect(() => {
    const to = todayISO(); const fromD = new Date(); fromD.setDate(fromD.getDate() - 13);
    const from = fromD.toISOString().slice(0, 10);
    call(`/doctor-pays/daily-net?from=${from}&to=${to}`)
      .then((rows) => setLast14(rows.map((r) => ({ date: d10(r.date), collection: Number(r.collection), pay: Number(r.doctor_pay), otherExp: Number(r.other_expense), net: Number(r.net) }))))
      .catch(() => {});
  }, [call, doctorPays]);

  const addDoc = async () => { setErr(""); if (!d.name.trim()) return; setBusy(true); try { await addDoctor({ name: d.name, shift: d.shift, payType: d.payType, rate: Number(d.rate) || 0 }); setD({ name: "", shift: "Morning", payType: "Daily", rate: "" }); } catch (e) { setErr(e.message); } setBusy(false); };
  const delDoc = async (id) => { try { await removeDoctor(id); } catch (e) { setErr(e.message); } };
  const addPay = async () => { setErr(""); if (!p.doctorId || !p.amount) return; setBusy(true); try { await addDoctorPay({ doctorId: p.doctorId, date: p.date, amount: Number(p.amount) }); setP({ date: todayISO(), doctorId: "", amount: "" }); } catch (e) { setErr(e.message); } setBusy(false); };
  const delPay = async (id) => { try { await removeDoctorPay(id); } catch (e) { setErr(e.message); } };

  return (
    <div>
      <div className="card">
        <h2>Doctors on roster</h2>
        <div className="form-grid">
          <div><label>Name</label><input type="text" value={d.name} onChange={(e) => setD({ ...d, name: e.target.value })} placeholder="Dr. …" /></div>
          <div><label>Shift</label><select value={d.shift} onChange={(e) => setD({ ...d, shift: e.target.value })}>{SHIFTS.map((s) => <option key={s} value={s}>{s}</option>)}</select></div>
          <div><label>Pay type</label><select value={d.payType} onChange={(e) => setD({ ...d, payType: e.target.value })}>{PAY_TYPES.map((s) => <option key={s} value={s}>{s}</option>)}</select></div>
          <div><label>Rate (₹)</label><input type="number" value={d.rate} onChange={(e) => setD({ ...d, rate: e.target.value })} /></div>
        </div>
        <button className="btn" type="button" disabled={busy} onClick={addDoc}>Add doctor</button>
        <table style={{ marginTop: 14 }}><thead><tr><th>Name</th><th>Shift</th><th>Pay type</th><th className="num">Rate</th><th></th></tr></thead>
          <tbody>{doctors.map((x) => (<tr key={x.id}><td>{x.name}</td><td><span className={"pill " + x.shift.toLowerCase()}>{x.shift}</span></td><td>{x.payType}</td><td className="num">{inr(x.rate)}</td><td><button className="btn danger small" type="button" onClick={() => delDoc(x.id)}>Delete</button></td></tr>))}</tbody>
        </table>
        <ErrorNote msg={err} />
      </div>
      <div className="card">
        <h2>Log a pay entry</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={p.date} onChange={(e) => setP({ ...p, date: e.target.value })} /></div>
          <div><label>Doctor</label><select value={p.doctorId} onChange={(e) => { const doc = doctors.find((x) => x.id === e.target.value); setP({ ...p, doctorId: e.target.value, amount: doc ? doc.rate : p.amount }); }}><option value="">Select doctor</option>{doctors.map((x) => <option key={x.id} value={x.id}>{x.name} ({x.shift})</option>)}</select></div>
          <div><label>Amount (₹)</label><input type="number" value={p.amount} onChange={(e) => setP({ ...p, amount: e.target.value })} /></div>
        </div>
        <button className="btn" type="button" disabled={busy} onClick={addPay}>Log pay entry</button>
        <table style={{ marginTop: 14 }}><thead><tr><th>Date</th><th>Doctor</th><th className="num">Amount</th><th></th></tr></thead>
          <tbody>{[...doctorPays].sort((a, b) => (a.date < b.date ? 1 : -1)).map((x) => (<tr key={x.id}><td>{x.date}</td><td>{x.doctorName}</td><td className="num">{inr(x.amount)}</td><td><button className="btn danger small" type="button" onClick={() => delPay(x.id)}>Delete</button></td></tr>))}</tbody>
        </table>
        <ExportRow onExcel={() => exportExcel("doctor-pay", { DoctorPay: doctorPays.map((x) => ({ Date: x.date, Doctor: x.doctorName, Amount: x.amount })) })} />
      </div>
      <div className="card">
        <h2>Daily net result — last 14 days</h2>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={last14}><CartesianGrid strokeDasharray="3 3" stroke="#E1E8E6" /><XAxis dataKey="date" tickFormatter={(v) => v.slice(5)} fontSize={11} /><YAxis fontSize={11} /><Tooltip formatter={(v) => inr(v)} /><Bar dataKey="net" fill="#0B4F4A" radius={[4, 4, 0, 0]} /></BarChart>
        </ResponsiveContainer>
        <table style={{ marginTop: 10 }}><thead><tr><th>Date</th><th className="num">Collection</th><th className="num">Doctor Pay</th><th className="num">Other Expenses</th><th className="num">Net</th></tr></thead>
          <tbody>{last14.slice().reverse().map((r) => (<tr key={r.date}><td>{r.date}</td><td className="num">{inr(r.collection)}</td><td className="num">{inr(r.pay)}</td><td className="num">{inr(r.otherExp)}</td><td className="num" style={{ color: r.net >= 0 ? "var(--income)" : "var(--expense)" }}>{inr(r.net)}</td></tr>))}</tbody>
        </table>
        <div className="note-box">Monthly-rate pay shows in full on the date it's logged, so that day may dip sharply — log monthly pay in equal daily portions, or use the Collections page's weekly/monthly rollups instead.</div>
      </div>
    </div>
  );
}

/* ============================== REFERRALS ============================== */
function Referrals({ referrals, addReferral, removeReferral, fy }) {
  const blank = { date: todayISO(), patientName: "", referralType: "Lab Test", referredTo: "", amount: "", notes: "" };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const save = async () => { setErr(""); if (!form.patientName.trim() || !form.amount) { setErr("Enter patient name and amount."); return; } setBusy(true); try { await addReferral({ ...form, amount: Number(form.amount) }); setForm(blank); } catch (e) { setErr(e.message); } setBusy(false); };
  const remove = async (id) => { try { await removeReferral(id); } catch (e) { setErr(e.message); } };
  const range = fyRange(fy);
  const total = referrals.filter((r) => r.date >= range.start && r.date <= range.end).reduce((a, r) => a + r.amount, 0);
  return (
    <div>
      <div className="card">
        <h2>New referral</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Patient</label><input type="text" value={form.patientName} onChange={(e) => setForm({ ...form, patientName: e.target.value })} /></div>
          <div><label>Referral type</label><select value={form.referralType} onChange={(e) => setForm({ ...form, referralType: e.target.value })}>{REFERRAL_TYPES.map((r) => <option key={r} value={r}>{r}</option>)}</select></div>
          <div><label>Referred to</label><input type="text" value={form.referredTo} onChange={(e) => setForm({ ...form, referredTo: e.target.value })} /></div>
          <div><label>Amount received (₹)</label><input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></div>
          <div><label>Notes</label><input type="text" value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></div>
        </div>
        <button className="btn" type="button" disabled={busy} onClick={save}>Save referral</button>
        <ErrorNote msg={err} />
      </div>
      <div className="card">
        <h2>Referral income — FY {fy} total: {inr(total)}</h2>
        {referrals.length === 0 ? <div className="empty">No referrals recorded yet.</div> : (
          <table><thead><tr><th>Date</th><th>Patient</th><th>Type</th><th>Referred To</th><th className="num">Amount</th><th>Notes</th><th></th></tr></thead>
            <tbody>{[...referrals].sort((a, b) => (a.date < b.date ? 1 : -1)).map((r) => (<tr key={r.id}><td>{r.date}</td><td>{r.patientName}</td><td>{r.referralType}</td><td>{r.referredTo}</td><td className="num">{inr(r.amount)}</td><td>{r.notes}</td><td><button className="btn danger small" type="button" onClick={() => remove(r.id)}>Delete</button></td></tr>))}</tbody>
          </table>
        )}
        <ExportRow onExcel={() => exportExcel("referral-income", { Referrals: referrals })} />
      </div>
    </div>
  );
}

/* ============================== GIFTS REGISTER ============================== */
function Gifts({ gifts, addGift, removeGift, doctors }) {
  const blank = { date: todayISO(), repName: "", company: "", gift: "", doctorId: "" };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const save = async () => { setErr(""); if (!form.repName.trim()) { setErr("Enter the medical rep's name."); return; } setBusy(true); try { await addGift(form); setForm(blank); } catch (e) { setErr(e.message); } setBusy(false); };
  const remove = async (id) => { try { await removeGift(id); } catch (e) { setErr(e.message); } };
  return (
    <div>
      <div className="card">
        <h2>Log a gift received</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Medical rep</label><input type="text" value={form.repName} onChange={(e) => setForm({ ...form, repName: e.target.value })} /></div>
          <div><label>Company</label><input type="text" value={form.company} onChange={(e) => setForm({ ...form, company: e.target.value })} /></div>
          <div><label>Gift description</label><input type="text" value={form.gift} onChange={(e) => setForm({ ...form, gift: e.target.value })} /></div>
          <div><label>Doctor</label><select value={form.doctorId} onChange={(e) => setForm({ ...form, doctorId: e.target.value })}><option value="">— Select —</option>{doctors.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
        </div>
        <button className="btn" type="button" disabled={busy} onClick={save}>Save entry</button>
        <ErrorNote msg={err} />
        <div className="note-box">Disclosure log only — never posted to income or expenses. Gifting from pharmaceutical reps to doctors is restricted under India's medical ethics rules (e.g. the Uniform Code for Pharmaceuticals Marketing Practices and NMC's professional-conduct regulations); check current guidance with your professional body if in doubt.</div>
      </div>
      <div className="card">
        <h2>Gifts register</h2>
        {gifts.length === 0 ? <div className="empty">No entries yet.</div> : (
          <table><thead><tr><th>Date</th><th>Rep</th><th>Company</th><th>Gift</th><th>Doctor</th><th></th></tr></thead>
            <tbody>{[...gifts].sort((a, b) => (a.date < b.date ? 1 : -1)).map((g) => (<tr key={g.id}><td>{g.date}</td><td>{g.repName}</td><td>{g.company}</td><td>{g.gift}</td><td>{g.doctorName || "—"}</td><td><button className="btn danger small" type="button" onClick={() => remove(g.id)}>Delete</button></td></tr>))}</tbody>
          </table>
        )}
        <ExportRow onExcel={() => exportExcel("gifts-register", { Gifts: gifts })} />
      </div>
    </div>
  );
}

/* ============================== EXPENSES ============================== */
function Expenses({ expenses, addExpense, removeExpense, fy }) {
  const blank = { date: todayISO(), category: EXPENSE_CATEGORIES[0], amount: "", narration: "", image: null };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const save = async () => { setErr(""); if (!form.amount) { setErr("Enter an amount."); return; } setBusy(true); try { await addExpense({ date: form.date, category: form.category, amount: Number(form.amount), narration: form.narration, imageUrl: form.image }); setForm(blank); } catch (e) { setErr(e.message); } setBusy(false); };
  const remove = async (id) => { try { await removeExpense(id); } catch (e) { setErr(e.message); } };
  const range = fyRange(fy);
  const inFY = expenses.filter((e) => e.date >= range.start && e.date <= range.end);
  const byCat = {}; inFY.forEach((e) => { byCat[e.category] = (byCat[e.category] || 0) + e.amount; });
  const total = inFY.reduce((a, e) => a + e.amount, 0);
  return (
    <div>
      <div className="card">
        <h2>New expense</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Category</label><select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>{EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}</select></div>
          <div><label>Amount (₹)</label><input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></div>
          <div style={{ gridColumn: "span 2" }}><label>Narration</label><input type="text" style={{ width: "100%" }} value={form.narration} onChange={(e) => setForm({ ...form, narration: e.target.value })} /></div>
        </div>
        <ImageCapture value={form.image} onChange={(img) => setForm({ ...form, image: img })} />
        <button className="btn" style={{ marginTop: 14 }} type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : "Save expense"}</button>
        <ErrorNote msg={err} />
      </div>
      <div className="card">
        <h2>Category totals — FY {fy} (total: {inr(total)})</h2>
        <table><thead><tr><th>Category</th><th className="num">Amount</th></tr></thead><tbody>{EXPENSE_CATEGORIES.map((c) => (<tr key={c}><td>{c}</td><td className="num">{inr(byCat[c] || 0)}</td></tr>))}</tbody></table>
      </div>
      <div className="card">
        <h2>All expense entries</h2>
        {expenses.length === 0 ? <div className="empty">No expenses logged yet.</div> : (
          <table><thead><tr><th>Date</th><th>Category</th><th>Narration</th><th className="num">Amount</th><th>Photo</th><th></th></tr></thead>
            <tbody>{[...expenses].sort((a, b) => (a.date < b.date ? 1 : -1)).map((e) => (<tr key={e.id}><td>{e.date}</td><td>{e.category}</td><td>{e.narration}</td><td className="num">{inr(e.amount)}</td><td>{e.image ? "📎" : "—"}</td><td><button className="btn danger small" type="button" onClick={() => remove(e.id)}>Delete</button></td></tr>))}</tbody>
          </table>
        )}
        <ExportRow onExcel={() => exportExcel("expenses", { Expenses: expenses, CategoryTotals: EXPENSE_CATEGORIES.map((c) => ({ Category: c, Amount: byCat[c] || 0 })) })} />
      </div>
    </div>
  );
}

/* ============================== FIXED ASSETS ============================== */
function FixedAssets({ assets, addAsset, removeAsset, fy }) {
  const { call } = useApi();
  const blank = { name: "", block: BLOCKS[0].name, rate: BLOCKS[0].rate, purchaseDate: todayISO(), cost: "" };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [dep, setDep] = useState({ rows: [], totalDep: 0 });

  useEffect(() => { call(`/assets/depreciation?fy=${fy}`).then((d) => setDep({ rows: (d.rows || []).map((r) => ({ ...mapAsset(r), ...r })), totalDep: Number(d.totalDep) })).catch(() => {}); }, [call, fy, assets]);

  const add = async () => { setErr(""); if (!form.name.trim() || !form.cost) { setErr("Enter a description and cost."); return; } setBusy(true); try { await addAsset({ ...form, cost: Number(form.cost) }); setForm(blank); } catch (e) { setErr(e.message); } setBusy(false); };
  const remove = async (id) => { try { await removeAsset(id); } catch (e) { setErr(e.message); } };

  return (
    <div>
      <div className="card">
        <h2>Add asset (CAPEX)</h2>
        <div className="form-grid">
          <div><label>Description</label><input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="e.g. ECG machine" /></div>
          <div><label>Block of assets</label><select value={form.block} onChange={(e) => { const b = BLOCKS.find((x) => x.name === e.target.value); setForm({ ...form, block: b.name, rate: b.rate }); }}>{BLOCKS.map((b) => <option key={b.name} value={b.name}>{b.name} ({b.rate}%)</option>)}</select></div>
          <div><label>Rate %</label><input type="number" value={form.rate} onChange={(e) => setForm({ ...form, rate: Number(e.target.value) })} /></div>
          <div><label>Date put to use</label><input type="date" value={form.purchaseDate} onChange={(e) => setForm({ ...form, purchaseDate: e.target.value })} /></div>
          <div><label>Cost (₹)</label><input type="number" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} /></div>
        </div>
        <button className="btn" type="button" disabled={busy} onClick={add}>Add to register</button>
        <ErrorNote msg={err} />
      </div>
      <div className="card">
        <h2>Depreciation register — FY {fy}</h2>
        {dep.rows.length === 0 ? <div className="empty">No fixed assets on the register yet.</div> : (
          <table>
            <thead><tr><th>Asset</th><th>Block</th><th className="num">Rate</th><th className="num">Cost</th><th className="num">Opening WDV</th><th className="num">Depreciation</th><th className="num">Closing WDV</th><th></th></tr></thead>
            <tbody>{dep.rows.map((r) => (
              <tr key={r.id}><td>{r.name}</td><td>{r.block}</td><td className="num">{r.rate}%</td><td className="num">{inr(r.cost)}</td>
                <td className="num">{r.applicable ? inr(r.wdvStart) : "—"}</td><td className="num">{r.applicable ? inr(r.dep) : "—"}</td><td className="num">{r.applicable ? inr(r.wdvEnd) : "—"}</td>
                <td><button className="btn danger small" type="button" onClick={() => remove(r.id)}>Delete</button></td></tr>
            ))}</tbody>
            <tfoot><tr style={{ fontWeight: 700 }}><td colSpan={5}>Total depreciation for FY {fy}</td><td className="num">{inr(dep.totalDep)}</td><td colSpan={2}></td></tr></tfoot>
          </table>
        )}
        <ExportRow onExcel={() => exportExcel("fixed-assets", { DepreciationFY: dep.rows.map((r) => ({ Asset: r.name, Block: r.block, Rate: r.rate, Cost: r.cost, OpeningWDV: r.applicable ? r.wdvStart : 0, Depreciation: r.applicable ? r.dep : 0, ClosingWDV: r.applicable ? r.wdvEnd : 0 })) })} />
        <div className="note-box">Depreciation follows the Section 32 written-down-value method (half rate if used under 180 days in the year of purchase), computed server-side so every screen agrees with the ledger.</div>
      </div>
    </div>
  );
}

/* ============================== FINANCIAL STATEMENTS ============================== */
function FinancialStatements({ fy, settings }) {
  const { call } = useApi();
  const [income, setIncome] = useState(null);
  const [capAcct, setCapAcct] = useState(null);
  const [bs, setBs] = useState(null);
  const [err, setErr] = useState("");

  useEffect(() => {
    setErr("");
    Promise.all([call(`/statements/income?fy=${fy}`), call(`/statements/capital-account?fy=${fy}`), call(`/statements/balance-sheet?fy=${fy}`)])
      .then(([i, c, b]) => { setIncome(i); setCapAcct(c); setBs(b); })
      .catch((e) => setErr(e.message));
  }, [call, fy]);

  if (err) return <div className="card"><h2>Couldn't load statements</h2><ErrorNote msg={err} /></div>;
  if (!income || !capAcct || !bs) return <div className="empty">Loading statements…</div>;

  const chartData = [{ name: "Income", value: income.income.total }, { name: "Expense", value: income.totalExpense }];

  return (
    <div>
      <div className="card">
        <h2>{settings.clinicName} — Income Statement, FY {fy}</h2>
        <p style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: -6 }}>{settings.proprietor}</p>
        <table><thead><tr><th>Income</th><th className="num">Amount</th></tr></thead>
          <tbody>
            <tr><td>Patient Collection Income</td><td className="num">{inr(income.income.patientCollection)}</td></tr>
            <tr><td>Referral Income</td><td className="num">{inr(income.income.referral)}</td></tr>
            <tr style={{ fontWeight: 700 }}><td>Total Income</td><td className="num">{inr(income.income.total)}</td></tr>
          </tbody>
        </table>
        <table style={{ marginTop: 14 }}><thead><tr><th>Expenses</th><th className="num">Amount</th></tr></thead>
          <tbody>
            {income.expenses.map((r) => (<tr key={r.name}><td>{r.name}</td><td className="num">{inr(r.amount)}</td></tr>))}
            <tr style={{ fontWeight: 700 }}><td>Total Expense</td><td className="num">{inr(income.totalExpense)}</td></tr>
            <tr style={{ fontWeight: 700, borderTop: "2px solid var(--accent)" }}><td>{income.netProfit >= 0 ? "Net Profit" : "Net Loss"}</td><td className="num" style={{ color: income.netProfit >= 0 ? "var(--income)" : "var(--expense)" }}>{inr(Math.abs(income.netProfit))}</td></tr>
          </tbody>
        </table>
        <ResponsiveContainer width="100%" height={180}>
          <BarChart data={chartData} layout="vertical"><XAxis type="number" fontSize={11} /><YAxis type="category" dataKey="name" fontSize={12} width={70} /><Tooltip formatter={(v) => inr(v)} /><Bar dataKey="value" radius={[0, 6, 6, 0]}><Cell fill="#1F8A5F" /><Cell fill="#B3423A" /></Bar></BarChart>
        </ResponsiveContainer>
        <ExportRow onExcel={() => exportExcel("income-statement", { Income: [{ Head: "Patient Collection Income", Amount: income.income.patientCollection }, { Head: "Referral Income", Amount: income.income.referral }], Expenses: income.expenses.map((r) => ({ Head: r.name, Amount: r.amount })), Summary: [{ TotalIncome: income.income.total, TotalExpense: income.totalExpense, NetProfit: income.netProfit }] })} />
      </div>
      <div className="card">
        <h2>Capital Account — as of {capAcct.asOf}</h2>
        <table><tbody>
          <tr><td>Capital introduced (cumulative)</td><td className="num">{inr(capAcct.capIntroduced)}</td></tr>
          <tr><td>Less: Drawings (cumulative)</td><td className="num">-{inr(capAcct.drawings)}</td></tr>
          <tr><td>Add: Cumulative net profit (all periods to date)</td><td className="num">{inr(capAcct.cumulativeNetProfit)}</td></tr>
          <tr style={{ fontWeight: 700, borderTop: "2px solid var(--accent)" }}><td>Closing Capital — {settings.proprietor}</td><td className="num">{inr(capAcct.closingCapital)}</td></tr>
        </tbody></table>
        <ExportRow onExcel={() => exportExcel("capital-account", { CapitalAccount: [{ CapitalIntroduced: capAcct.capIntroduced, Drawings: capAcct.drawings, CumulativeNetProfit: capAcct.cumulativeNetProfit, ClosingCapital: capAcct.closingCapital }] })} />
      </div>
      <div className="card">
        <h2>Balance Sheet — as of {bs.asOf}</h2>
        <table><thead><tr><th>Liabilities & Capital</th><th className="num">Amount</th><th>Assets</th><th className="num">Amount</th></tr></thead>
          <tbody>
            <tr><td>Capital Account</td><td className="num">{inr(bs.capital.closingCapital)}</td><td>Cash & Bank</td><td className="num">{inr(bs.assets.cashBank)}</td></tr>
            <tr><td></td><td></td><td>Sundry Debtors (patient dues)</td><td className="num">{inr(bs.assets.debtors)}</td></tr>
            <tr><td></td><td></td><td>Fixed Assets (net of depreciation)</td><td className="num">{inr(bs.assets.fixedAssetsNet)}</td></tr>
            <tr style={{ fontWeight: 700, borderTop: "2px solid var(--accent)" }}><td>Total</td><td className="num">{inr(bs.capital.closingCapital)}</td><td>Total</td><td className="num">{inr(bs.assets.total)}</td></tr>
          </tbody>
        </table>
        <div style={{ marginTop: 10 }}><span className={"balance-tag " + (bs.ties ? "ok" : "bad")}>{bs.ties ? "Balance sheet ties out" : `Off by ${inr(Math.abs(bs.capital.closingCapital - bs.assets.total))}`}</span></div>
        <ExportRow onExcel={() => exportExcel("balance-sheet", { BalanceSheet: [{ Capital: bs.capital.closingCapital, CashBank: bs.assets.cashBank, Debtors: bs.assets.debtors, FixedAssetsNet: bs.assets.fixedAssetsNet, TotalAssets: bs.assets.total }] })} />
        <div className="note-box">Model assumes the clinic settles expenses directly out of collections (no supplier credit tracked), so Assets equal Capital by construction. Add a Sundry Creditors flow if medicine purchases start going on credit.</div>
      </div>
    </div>
  );
}

/* ============================== SETTINGS ============================== */
function SettingsPage({ settings, updateSettings, session, origin, capital, addCapital, removeCapital }) {
  const [form, setForm] = useState(settings);
  useEffect(() => setForm(settings), [settings]);
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const canEditSettings = session.role === "Admin" || session.role === "Doctor";

  const [cap, setCap] = useState({ date: todayISO(), type: "Introduced", amount: "", note: "" });
  const [capErr, setCapErr] = useState("");

  const save = async () => { setErr(""); setBusy(true); try { await updateSettings(form); } catch (e) { setErr(e.message); } setBusy(false); };
  const saveCap = async () => { setCapErr(""); if (!cap.amount) { setCapErr("Enter an amount."); return; } try { await addCapital({ ...cap, amount: Number(cap.amount) }); setCap({ date: todayISO(), type: "Introduced", amount: "", note: "" }); } catch (e) { setCapErr(e.message); } };

  return (
    <div>
      <div className="card">
        <h2>Clinic profile</h2>
        <div className="form-grid">
          <div><label>Clinic name</label><input type="text" value={form.clinicName} onChange={(e) => setForm({ ...form, clinicName: e.target.value })} disabled={!canEditSettings} /></div>
          <div><label>Proprietor / Head doctor</label><input type="text" value={form.proprietor} onChange={(e) => setForm({ ...form, proprietor: e.target.value })} disabled={!canEditSettings} /></div>
          <div><label>Address</label><input type="text" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} disabled={!canEditSettings} /></div>
          <div><label>Phone</label><input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} disabled={!canEditSettings} /></div>
        </div>
        {canEditSettings ? <button className="btn" type="button" disabled={busy} onClick={save}>Save profile</button> : <div className="note-box">Only Admin or Doctor roles can edit the clinic profile.</div>}
        <ErrorNote msg={err} />
      </div>

      <div className="card">
        <h2>Capital account entries</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={cap.date} onChange={(e) => setCap({ ...cap, date: e.target.value })} /></div>
          <div><label>Type</label><select value={cap.type} onChange={(e) => setCap({ ...cap, type: e.target.value })}><option value="Introduced">Capital Introduced</option><option value="Drawings">Drawings</option></select></div>
          <div><label>Amount (₹)</label><input type="number" value={cap.amount} onChange={(e) => setCap({ ...cap, amount: e.target.value })} /></div>
          <div><label>Note</label><input type="text" value={cap.note} onChange={(e) => setCap({ ...cap, note: e.target.value })} /></div>
        </div>
        <button className="btn" type="button" onClick={saveCap}>Save entry</button>
        <ErrorNote msg={capErr} />
        <table style={{ marginTop: 14 }}><thead><tr><th>Date</th><th>Type</th><th className="num">Amount</th><th>Note</th><th></th></tr></thead>
          <tbody>{[...capital].sort((a, b) => (a.date < b.date ? 1 : -1)).map((c) => (<tr key={c.id}><td>{c.date}</td><td>{c.type}</td><td className="num">{inr(c.amount)}</td><td>{c.note}</td><td><button className="btn danger small" type="button" onClick={() => removeCapital(c.id)}>Delete</button></td></tr>))}</tbody>
        </table>
      </div>

      <div className="card">
        <h2>Connection</h2>
        <p style={{ fontSize: 13, color: "var(--ink-soft)" }}>API server: <code>{origin}</code></p>
        <p style={{ fontSize: 13, color: "var(--ink-soft)" }}>Logged in as {session.name} ({session.userId}) — role: {session.role}</p>
        <div className="note-box">To add more staff logins, use the Register tab on the sign-in screen — there's no separate admin user list in this build yet. Same user ID and password work on both the web and mobile views of this app, since both call the same API server.</div>
      </div>
    </div>
  );
}
