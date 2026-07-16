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
const COLORS = ["#714B67", "#C9A227", "#1F8A5F", "#B3423A", "#5B6B69", "#7FB3AB", "#E8C468"];

const PERMISSION_MODULES = [
  { key: "cases", label: "Case Records" },
  { key: "collections", label: "Collections" },
  { key: "doctorPay", label: "Doctor Shifts & Pay" },
  { key: "referrals", label: "Referral Income" },
  { key: "gifts", label: "Gifts Register" },
  { key: "expenses", label: "Expenses" },
  { key: "assets", label: "Fixed Assets" },
  { key: "statements", label: "Financial Statements" },
  { key: "auditLog", label: "User Access Report" },
];
const LEVELS = ["none", "view", "write", "edit", "delete"];
const LEVEL_LABELS = { none: "No access", view: "View only", write: "View + Write", edit: "View + Write + Edit", delete: "Full (View/Write/Edit/Delete)" };
const PASSWORD_RULE = /^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$/;
const PASSWORD_HINT = "At least 8 characters, with a letter, a number, and a special character.";

function makeCan(session) {
  return (moduleKey, level) => {
    if (!session) return false;
    if (session.role === "Admin") return true;
    const p = session.permissions?.[moduleKey];
    if (!p) return false;
    if (level === "export") return !!p.export;
    return LEVELS.indexOf(p.level || "none") >= LEVELS.indexOf(level);
  };
}

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
/** Escapes text before it's inserted into a print view's innerHTML — without
 *  this, a patient name or note containing "<" or "&" could be interpreted
 *  as HTML/script rather than displayed as plain text. */
function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch]));
}

/** Reusable column sorting: useSortableRows(rows, columns, defaultKey)
 *  takes a { key: accessorFn } map declared upfront (so sorting is correct
 *  from the very first render) and returns the sorted array plus a <Th>
 *  component for clickable, arrow-indicating headers. */
function useSortableRows(rows, columns, defaultKey, defaultDir = "asc") {
  const [sortKey, setSortKey] = useState(defaultKey);
  const [sortDir, setSortDir] = useState(defaultDir);

  const sorted = useMemo(() => {
    const acc = columns[sortKey];
    if (!acc) return rows;
    const withVals = rows.map((r) => ({ r, v: acc(r) }));
    withVals.sort((a, b) => {
      const av = a.v, bv = b.v;
      let cmp;
      if (typeof av === "number" && typeof bv === "number") cmp = av - bv;
      else cmp = String(av ?? "").localeCompare(String(bv ?? ""), undefined, { numeric: true, sensitivity: "base" });
      return sortDir === "asc" ? cmp : -cmp;
    });
    return withVals.map((x) => x.r);
  }, [rows, columns, sortKey, sortDir]);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    else { setSortKey(key); setSortDir("asc"); }
  };

  const Th = ({ sortKeyName, children, className }) => {
    const active = sortKey === sortKeyName;
    return (
      <th className={className} onClick={() => toggleSort(sortKeyName)} style={{ cursor: "pointer", userSelect: "none", whiteSpace: "nowrap" }}>
        {children}{active ? (sortDir === "asc" ? " ▲" : " ▼") : <span style={{ opacity: 0.25 }}> ⇅</span>}
      </th>
    );
  };

  return { sorted, Th };
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

/** Date-range export control: pick From/To, then Export Excel/PDF filters
 *  `rows` by `dateField` before handing off. Used on every module page. */
function CustomExport({ rows, dateField, buildSheets, filenameBase, printTitle, printColumns, canExport = true }) {
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [open, setOpen] = useState(false);

  if (!canExport) return null;

  const filtered = () => {
    if (!from && !to) return rows;
    return rows.filter((r) => {
      const d = r[dateField];
      if (!d) return false;
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
  };

  const doExcel = () => exportExcel(filenameBase, buildSheets(filtered()));

  const doPrint = () => {
    const data = filtered();
    const rangeLabel = from || to ? `${from || "start"} to ${to || "today"}` : "all dates";
    const win = document.getElementById("print-root");
    if (!win) { window.print(); return; }
    const rowsHtml = data.map((r) => `<tr>${printColumns.map((c) => `<td>${escapeHtml(c.value(r))}</td>`).join("")}</tr>`).join("");
    win.innerHTML = `
      <h2>${escapeHtml(printTitle)}</h2>
      <p style="color:#5B6B69;font-size:12px;">Range: ${escapeHtml(rangeLabel)} — ${data.length} record(s)</p>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr>${printColumns.map((c) => `<th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">${escapeHtml(c.label)}</th>`).join("")}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    document.body.classList.add("printing-custom");
    window.print();
    setTimeout(() => { document.body.classList.remove("printing-custom"); win.innerHTML = ""; }, 300);
  };

  return (
    <div className="custom-export no-print">
      <button className="btn secondary small" type="button" onClick={() => setOpen((o) => !o)}>📅 Custom range export</button>
      {open && (
        <div className="custom-export-panel">
          <div><label>From</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></div>
          <div><label>To</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></div>
          <button className="btn small" type="button" onClick={doExcel}>⬇ Excel</button>
          <button className="btn small" type="button" onClick={doPrint}>⎙ PDF</button>
        </div>
      )}
    </div>
  );
}
function PrintRoot() { return <div id="print-root" className="print-root"></div>; }

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
const mapGift = (r) => ({ id: r.id, date: d10(r.gift_date), repName: r.rep_name, company: r.company, gift: r.gift_description, doctorId: r.doctor_id, doctorName: r.doctor_name, amount: Number(r.amount) || 0 });
const mapExpense = (r) => ({ id: r.id, date: d10(r.expense_date), category: r.category, amount: Number(r.amount), narration: r.narration, image: r.image_url });
const mapAsset = (r) => ({ id: r.id, name: r.name, block: r.block, rate: Number(r.rate), purchaseDate: d10(r.purchase_date), cost: Number(r.cost) });
const mapCapital = (r) => ({ id: r.id, date: d10(r.txn_date), type: r.txn_type, amount: Number(r.amount), note: r.note });
const mapSettings = (r) => ({ clinicName: r.clinic_name || "Ganatra Clinic", proprietor: r.proprietor || "Dr. Bhavisha Pratik Ganatra", address: r.address || "", phone: r.phone || "" });

/* -------- Image capture: upload, rotate to align, attach via API -------- */
function ImageCapture({ value, onChange }) {
  const { call, origin, token } = useApi();
  const [raw, setRaw] = useState(null);
  const [rotation, setRotation] = useState(0);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [displayUrl, setDisplayUrl] = useState(null);
  const inputRef = useRef();

  useEffect(() => {
    let objectUrl;
    if (value && !raw) {
      fetch(`${origin}${value}`, { headers: { Authorization: `Bearer ${token}` } })
        .then((r) => (r.ok ? r.blob() : Promise.reject(new Error("Could not load image"))))
        .then((blob) => { objectUrl = URL.createObjectURL(blob); setDisplayUrl(objectUrl); })
        .catch(() => setDisplayUrl(null));
    } else {
      setDisplayUrl(null);
    }
    return () => { if (objectUrl) URL.revokeObjectURL(objectUrl); };
  }, [value, raw, origin, token]);

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
        {displayUrl ? <img src={displayUrl} alt="attached stationery" className="imgcap-thumb" /> : <div className="empty">Loading photo…</div>}
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
  const [stage, setStage] = useState("form"); // form | adminOtp | forgotRequest | forgotReset | useridRequest | useridVerify
  const [mode, setMode] = useState("login");
  const [name, setName] = useState("");
  const [userId, setUserId] = useState("");
  const [password, setPassword] = useState("");
  const [email, setEmail] = useState("");
  const [mobile, setMobile] = useState("");
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");
  const [busy, setBusy] = useState(false);

  // admin OTP stage
  const [otpCode, setOtpCode] = useState("");
  const [devCode, setDevCode] = useState("");

  // forgot-password stage
  const [fpUserId, setFpUserId] = useState("");
  const [fpCode, setFpCode] = useState("");
  const [fpPassword, setFpPassword] = useState("");

  // forgot-User ID stage
  const [fuMobile, setFuMobile] = useState("");
  const [fuCode, setFuCode] = useState("");
  const [fuRevealedUserId, setFuRevealedUserId] = useState("");

  const localCall = (path, opts) => apiFetch(origin, null, path, opts);
  const isAdminBootstrapAttempt = userId.trim() === "pratik";

  const submit = async (e) => {
    e.preventDefault(); setError(""); setInfo("");
    if (!origin.trim()) { setError("Enter your API server URL first."); return; }
    if (!userId.trim() || !password) { setError("Enter a user ID and password."); return; }
    if (mode === "register" && !PASSWORD_RULE.test(password)) { setError(PASSWORD_HINT); return; }
    if (mode === "register" && !isAdminBootstrapAttempt && !mobile.trim()) { setError("Enter a mobile number — it's required for password resets."); return; }
    setBusy(true);
    try {
      if (mode === "login") {
        const data = await localCall("/auth/login", { method: "POST", body: { userId, password } });
        await storeSet("clinic:apiOrigin", origin);
        onLogin({ token: data.token, userId: data.user.userId, name: data.user.name, role: data.user.role, permissions: data.user.permissions });
      } else {
        const data = await localCall("/auth/register", { method: "POST", body: { userId, password, name: name || userId, email: email || undefined, mobile: mobile || undefined } });
        if (data.requiresOtp) {
          setDevCode(data.devCode || "");
          setInfo(data.message);
          setStage("adminOtp");
        } else {
          setInfo(data.message || "Registered — waiting for admin approval.");
          setMode("login"); setPassword("");
        }
      }
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const submitAdminOtp = async (e) => {
    e.preventDefault(); setError("");
    if (!otpCode.trim()) { setError("Enter the 6-digit code."); return; }
    setBusy(true);
    try {
      const data = await localCall("/auth/verify-admin-otp", { method: "POST", body: { userId, code: otpCode.trim() } });
      await storeSet("clinic:apiOrigin", origin);
      onLogin({ token: data.token, userId: data.user.userId, name: data.user.name, role: data.user.role, permissions: data.user.permissions });
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const requestReset = async (e) => {
    e.preventDefault(); setError(""); setInfo("");
    if (!fpUserId.trim()) { setError("Enter your user ID."); return; }
    setBusy(true);
    try {
      const data = await localCall("/auth/forgot-password", { method: "POST", body: { userId: fpUserId.trim() } });
      setInfo(data.message); setDevCode(data.devCode || "");
      setStage("forgotReset");
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const submitReset = async (e) => {
    e.preventDefault(); setError("");
    if (!PASSWORD_RULE.test(fpPassword)) { setError(PASSWORD_HINT); return; }
    setBusy(true);
    try {
      await localCall("/auth/reset-password", { method: "POST", body: { userId: fpUserId.trim(), code: fpCode.trim(), newPassword: fpPassword } });
      setInfo("Password updated — log in with your new password.");
      setStage("form"); setMode("login"); setUserId(fpUserId); setPassword("");
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const requestUserId = async (e) => {
    e.preventDefault(); setError(""); setInfo("");
    if (!fuMobile.trim()) { setError("Enter your registered mobile number."); return; }
    setBusy(true);
    try {
      const data = await localCall("/auth/forgot-userid", { method: "POST", body: { mobile: fuMobile.trim() } });
      setInfo(data.message); setDevCode(data.devCode || "");
      setStage("useridVerify");
    } catch (e) { setError(e.message); }
    setBusy(false);
  };

  const verifyUserId = async (e) => {
    e.preventDefault(); setError("");
    if (!fuCode.trim()) { setError("Enter the 6-digit code."); return; }
    setBusy(true);
    try {
      const data = await localCall("/auth/verify-userid-otp", { method: "POST", body: { mobile: fuMobile.trim(), code: fuCode.trim() } });
      setFuRevealedUserId(data.userId); setInfo(""); setDevCode("");
    } catch (e) { setError(e.message); }
    setBusy(false);
  };


  return (
    <div className="auth-wrap">
      <style>{`
        ${FONT_IMPORT}
        .auth-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;background:radial-gradient(circle at 25% 15%,#123f3a 0%,#081d1b 65%);font-family:'Inter',sans-serif;padding:24px;}
        .auth-card{width:100%;max-width:440px;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 30px 70px rgba(0,0,0,.45);}
        .auth-header{background:linear-gradient(135deg,#714B67,#4A2F44);padding:30px 26px 22px;text-align:center;}
        .auth-header .brand{font-family:'Plus Jakarta Sans',sans-serif;font-weight:800;color:#fff;font-size:26px;letter-spacing:.5px;margin-top:6px;}
        .auth-header .sub{color:#B9D8D2;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin-top:4px;}
        .pulse-path{stroke:#C9A227;stroke-width:2.5;fill:none;stroke-linecap:round;stroke-linejoin:round;stroke-dasharray:340;stroke-dashoffset:340;animation:draw 1.4s ease forwards .2s;}
        @keyframes draw{to{stroke-dashoffset:0;}}
        @media (prefers-reduced-motion:reduce){.pulse-path{animation:none;stroke-dashoffset:0;}}
        .auth-body{padding:26px 26px 24px;}
        .tabs{display:flex;gap:6px;margin-bottom:18px;}
        .tab-btn{flex:1;padding:8px 0;font-size:12.5px;font-weight:700;border:1px solid #E1E8E6;background:#F5F8F7;color:#5B6B69;border-radius:6px;cursor:pointer;}
        .tab-btn.active{background:#714B67;color:#fff;border-color:#714B67;}
        .field{margin-bottom:14px;}
        .field label{display:block;font-size:11px;text-transform:uppercase;letter-spacing:1px;color:#5B6B69;font-weight:700;margin-bottom:5px;}
        .field input,.field select{width:100%;border:1.5px solid #E1E8E6;border-radius:8px;padding:10px 11px;font-size:14.5px;font-family:'IBM Plex Mono',monospace;outline:none;box-sizing:border-box;}
        .field input:focus,.field select:focus{border-color:#714B67;}
        .field .hint{font-size:10.5px;color:#8a9a97;margin-top:4px;}
        .err{color:#B3423A;font-size:12.5px;margin-bottom:10px;}
        .info{color:#714B67;font-size:12.5px;margin-bottom:10px;background:#EAF3F1;padding:8px 10px;border-radius:6px;}
        .dev-code{font-family:'IBM Plex Mono',monospace;font-size:22px;letter-spacing:4px;text-align:center;background:#F3E3A8;color:#5b4a06;padding:10px;border-radius:8px;margin-bottom:14px;}
        .submit-btn{width:100%;background:#C9A227;color:#2A2103;font-weight:700;border:none;padding:12px;border-radius:8px;font-size:14.5px;cursor:pointer;box-shadow:0 4px 0 #96791b;}
        .submit-btn:active{transform:translateY(2px);box-shadow:0 2px 0 #96791b;}
        .note{font-size:11px;color:#8a9a97;margin-top:12px;line-height:1.6;text-align:center;}
        .api-field{background:#F5F8F7;border:1px dashed #C9A227;border-radius:8px;padding:10px 12px;margin-bottom:16px;}
        .api-field label{display:block;font-size:10.5px;text-transform:uppercase;letter-spacing:1px;color:#8a6d0a;font-weight:700;margin-bottom:5px;}
        .api-field input{width:100%;border:1px solid #E1E8E6;border-radius:6px;padding:8px 9px;font-size:13px;font-family:'IBM Plex Mono',monospace;box-sizing:border-box;}
        .back-link{display:block;text-align:center;margin-top:10px;font-size:12px;color:#5B6B69;background:none;border:none;cursor:pointer;text-decoration:underline;}
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

          {stage === "form" && (
            <>
              <div className="tabs">
                <button className={"tab-btn" + (mode === "login" ? " active" : "")} onClick={() => { setMode("login"); setError(""); setInfo(""); }} type="button">Log in</button>
                <button className={"tab-btn" + (mode === "register" ? " active" : "")} onClick={() => { setMode("register"); setError(""); setInfo(""); }} type="button">Register</button>
              </div>
              <form onSubmit={submit}>
                {mode === "register" && (
                  <div className="field"><label>Your name</label><input value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Front desk / Dr. Ganatra" /></div>
                )}
                <div className="field"><label>User ID</label><input value={userId} onChange={(e) => setUserId(e.target.value)} autoCapitalize="none" placeholder="your-user-id" /></div>
                {mode === "register" && (
                  <>
                    <div className="field">
                      <label>Email {isAdminBootstrapAttempt ? "(required for admin sign-up)" : "(optional)"}</label>
                      <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@example.com" />
                    </div>
                    {!isAdminBootstrapAttempt && (
                      <div className="field"><label>Mobile number</label><input type="tel" value={mobile} onChange={(e) => setMobile(e.target.value)} placeholder="+91 98765 43210" /></div>
                    )}
                  </>
                )}
                <div className="field">
                  <label>Password</label>
                  <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" />
                  {mode === "register" && <div className="hint">{PASSWORD_HINT}</div>}
                </div>
                {error && <div className="err">{error}</div>}
                {info && <div className="info">{info}</div>}
                <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Please wait…" : mode === "login" ? "Log in" : isAdminBootstrapAttempt ? "Create admin account" : "Register — pending approval"}</button>
              </form>
              {mode === "login" && (
                <div style={{ display: "flex", justifyContent: "center", gap: 14 }}>
                  <button className="back-link" type="button" onClick={() => { setStage("forgotRequest"); setError(""); setInfo(""); }}>Forgot your password?</button>
                  <button className="back-link" type="button" onClick={() => { setStage("useridRequest"); setError(""); setInfo(""); setFuRevealedUserId(""); }}>Forgot your User ID?</button>
                </div>
              )}
              <div className="note">Real staff accounts need admin approval before they can log in. Same user ID and password work on web and mobile — both talk to the same API server.</div>
            </>
          )}

          {stage === "adminOtp" && (
            <form onSubmit={submitAdminOtp}>
              {info && <div className="info">{info}</div>}
              {devCode && <div className="dev-code">{devCode}</div>}
              <div className="field"><label>6-digit code</label><input value={otpCode} onChange={(e) => setOtpCode(e.target.value)} placeholder="000000" maxLength={6} /></div>
              {error && <div className="err">{error}</div>}
              <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Verifying…" : "Verify & activate admin account"}</button>
              <button className="back-link" type="button" onClick={() => { setStage("form"); setError(""); setInfo(""); }}>Back</button>
            </form>
          )}

          {stage === "forgotRequest" && (
            <form onSubmit={requestReset}>
              <div className="field"><label>Your user ID</label><input value={fpUserId} onChange={(e) => setFpUserId(e.target.value)} placeholder="your-user-id" /></div>
              {error && <div className="err">{error}</div>}
              <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Sending…" : "Send code to my mobile"}</button>
              <button className="back-link" type="button" onClick={() => { setStage("form"); setError(""); }}>Back to log in</button>
            </form>
          )}

          {stage === "forgotReset" && (
            <form onSubmit={submitReset}>
              {info && <div className="info">{info}</div>}
              {devCode && <div className="dev-code">{devCode}</div>}
              <div className="field"><label>6-digit code</label><input value={fpCode} onChange={(e) => setFpCode(e.target.value)} placeholder="000000" maxLength={6} /></div>
              <div className="field">
                <label>New password</label>
                <input type="password" value={fpPassword} onChange={(e) => setFpPassword(e.target.value)} placeholder="••••••••" />
                <div className="hint">{PASSWORD_HINT}</div>
              </div>
              {error && <div className="err">{error}</div>}
              <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Saving…" : "Set new password"}</button>
              <button className="back-link" type="button" onClick={() => { setStage("form"); setError(""); }}>Back to log in</button>
            </form>
          )}

          {stage === "useridRequest" && (
            <form onSubmit={requestUserId}>
              <div className="field"><label>Your registered mobile number</label><input type="tel" value={fuMobile} onChange={(e) => setFuMobile(e.target.value)} placeholder="+91 98765 43210" /></div>
              {error && <div className="err">{error}</div>}
              <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Sending…" : "Send code to this number"}</button>
              <button className="back-link" type="button" onClick={() => { setStage("form"); setError(""); }}>Back to log in</button>
            </form>
          )}

          {stage === "useridVerify" && (
            <div>
              {!fuRevealedUserId ? (
                <form onSubmit={verifyUserId}>
                  {info && <div className="info">{info}</div>}
                  {devCode && <div className="dev-code">{devCode}</div>}
                  <div className="field"><label>6-digit code</label><input value={fuCode} onChange={(e) => setFuCode(e.target.value)} placeholder="000000" maxLength={6} /></div>
                  {error && <div className="err">{error}</div>}
                  <button className="submit-btn" type="submit" disabled={busy}>{busy ? "Verifying…" : "Reveal my User ID"}</button>
                  <button className="back-link" type="button" onClick={() => { setStage("form"); setError(""); }}>Back to log in</button>
                </form>
              ) : (
                <div>
                  <div className="info">Your User ID is:</div>
                  <div className="dev-code">{fuRevealedUserId}</div>
                  <button className="submit-btn" type="button" onClick={() => { setUserId(fuRevealedUserId); setMode("login"); setStage("form"); }}>Log in now</button>
                  <button className="back-link" type="button" onClick={() => { setFpUserId(fuRevealedUserId); setStage("forgotRequest"); setError(""); setInfo(""); }}>I also need to reset my password</button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================== NAV / SHELL ============================== */
const NAV = [
  { key: "dashboard", label: "Dashboard", icon: "🏠", grad: ["#F4A340", "#E85D3D"] }, { key: "cases", label: "Case Records", module: "cases", icon: "📋", grad: ["#3FA9D9", "#1F6FA8"] },
  { key: "patientMaster", label: "Patient Master", module: "cases", icon: "🧑‍🤝‍🧑", grad: ["#B45FC7", "#7C3AA6"] },
  { key: "patients", label: "Patient History", module: "cases", icon: "🕒", grad: ["#4A90D9", "#1F5FA8"] },
  { key: "collections", label: "Collections", module: "collections", icon: "💰", grad: ["#3FB86E", "#1F8A5F"] }, { key: "doctors", label: "Doctor Shifts & Pay", module: "doctorPay", icon: "🩺", grad: ["#E8557D", "#B3336B"] },
  { key: "referrals", label: "Referral Income", module: "referrals", icon: "🔗", grad: ["#5C7FE8", "#3B4FC7"] }, { key: "gifts", label: "Gifts Register", module: "gifts", icon: "🎁", grad: ["#E85D8A", "#C7336B"] },
  { key: "expenses", label: "Expenses", module: "expenses", icon: "🧾", grad: ["#E8823D", "#B3423A"] }, { key: "assets", label: "Fixed Assets", module: "assets", icon: "🏢", grad: ["#6C7B95", "#3E4A63"] },
  { key: "statements", label: "Financial Statements", module: "statements", icon: "📊", grad: ["#7C6FD9", "#4F3FA8"] }, { key: "settings", label: "Settings", icon: "⚙️", grad: ["#8A97A8", "#5B6B78"] },
  { key: "auditLog", label: "User Access Report", module: "auditLog", icon: "🔍", grad: ["#2FADA0", "#1B6E7A"] },
  { key: "admin", label: "User Approvals", adminOnly: true, icon: "✅", grad: ["#4FB88A", "#1F8A5F"] },
];

export default function App() {
  const [origin, setOrigin] = useState("https://ganatra-clinic.onrender.com");
  const [session, setSession] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [view, setView] = useState("launcher");
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
  const [loans, setLoans] = useState([]);
  const [patientsMaster, setPatientsMaster] = useState([]);
  const [deposits, setDeposits] = useState([]);

  useEffect(() => { storeGet("clinic:apiOrigin").then((o) => { if (o) setOrigin(o); }); }, []);

  const call = useCallback((path, opts) => apiFetch(origin, session?.token, path, opts), [origin, session]);
  const apiValue = useMemo(() => ({ call, origin, token: session?.token }), [call, origin, session]);

  const reloadAll = useCallback(async () => {
    setLoading(true); setLoadError("");
    const can = makeCan(session);
    const safeCall = async (path, moduleKey) => {
      if (moduleKey && !can(moduleKey, "view")) return [];
      try { return await call(path); } catch { return []; }
    };
    try {
      const [s, d, c, col, dp, ref, g, exp, ast, cap, other, pm] = await Promise.all([
        call("/settings").catch(() => ({})),
        safeCall("/doctors", "doctorPay"), safeCall("/cases", "cases"), safeCall("/collections", "collections"),
        safeCall("/doctor-pays", "doctorPay"), safeCall("/referrals", "referrals"), safeCall("/gifts", "gifts"), safeCall("/expenses", "expenses"),
        safeCall("/assets", "assets"), safeCall("/capital", "statements"), safeCall("/other-balance", "statements"), safeCall("/patient-master", "cases"),
      ]);
      setSettings(mapSettings(s || {})); setDoctors((d || []).map(mapDoctor)); setCases((c || []).map(mapCase));
      setCollections((col || []).map(mapCollection)); setDoctorPays((dp || []).map(mapDoctorPay));
      setReferrals((ref || []).map(mapReferral)); setGifts((g || []).map(mapGift));
      setExpenses((exp || []).map(mapExpense)); setAssets((ast || []).map(mapAsset)); setCapital((cap || []).map(mapCapital));
      const otherMapped = (other || []).map(mapOther);
      setLoans(otherMapped.filter((x) => x.category === "unsecured_loan"));
      setDeposits(otherMapped.filter((x) => x.category === "security_deposit"));
      setPatientsMaster((pm || []).map(mapPatientMaster));
    } catch (e) { setLoadError(e.message); }
    setLoading(false);
  }, [call, session]);

  useEffect(() => { if (session) reloadAll(); }, [session]); // eslint-disable-line

  // create/delete helpers — call the API, then patch local state
  const addDoctor = useCallback(async (body) => { const r = await call("/doctors", { method: "POST", body }); setDoctors((p) => [...p, mapDoctor(r)]); }, [call]);
  const updateDoctor = useCallback(async (id, body) => { const r = await call(`/doctors/${id}`, { method: "PUT", body }); setDoctors((p) => p.map((x) => (x.id === id ? mapDoctor(r) : x))); }, [call]);
  const removeDoctor = useCallback(async (id) => { await call(`/doctors/${id}`, { method: "DELETE" }); setDoctors((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addCase = useCallback(async (body) => {
    const r = await call("/cases", { method: "POST", body });
    const doc = doctors.find((d) => d.id === body.doctorId);
    setCases((p) => [{ ...mapCase(r), doctorName: doc?.name }, ...p]);
    if (r.collection) setCollections((p) => [mapCollection(r.collection), ...p]);
    return r;
  }, [call, doctors]);
  const updateCase = useCallback(async (id, body) => {
    const r = await call(`/cases/${id}`, { method: "PUT", body });
    const doc = doctors.find((d) => d.id === body.doctorId);
    setCases((p) => p.map((x) => (x.id === id ? { ...mapCase(r), doctorName: doc?.name } : x)));
    setCollections((p) => p.map((c) => (c.caseId === id ? { ...c, patientName: body.patientName, phone: body.phone, caseNo: r.case_no } : c)));
  }, [call, doctors]);
  const removeCase = useCallback(async (id) => { await call(`/cases/${id}`, { method: "DELETE" }); setCases((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addCollection = useCallback(async (body) => { const r = await call("/collections", { method: "POST", body }); setCollections((p) => [mapCollection(r), ...p]); }, [call]);
  const updateCollection = useCallback(async (id, body) => { const r = await call(`/collections/${id}`, { method: "PUT", body }); setCollections((p) => p.map((x) => (x.id === id ? mapCollection(r) : x))); }, [call]);
  const removeCollection = useCallback(async (id) => { await call(`/collections/${id}`, { method: "DELETE" }); setCollections((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addDoctorPay = useCallback(async (body) => { const r = await call("/doctor-pays", { method: "POST", body }); const doc = doctors.find((d) => d.id === body.doctorId); setDoctorPays((p) => [{ ...mapDoctorPay(r), doctorName: doc?.name }, ...p]); }, [call, doctors]);
  const updateDoctorPay = useCallback(async (id, body) => { const r = await call(`/doctor-pays/${id}`, { method: "PUT", body }); const doc = doctors.find((d) => d.id === body.doctorId); setDoctorPays((p) => p.map((x) => (x.id === id ? { ...mapDoctorPay(r), doctorName: doc?.name } : x))); }, [call, doctors]);
  const removeDoctorPay = useCallback(async (id) => { await call(`/doctor-pays/${id}`, { method: "DELETE" }); setDoctorPays((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addReferral = useCallback(async (body) => { const r = await call("/referrals", { method: "POST", body }); setReferrals((p) => [mapReferral(r), ...p]); }, [call]);
  const updateReferral = useCallback(async (id, body) => { const r = await call(`/referrals/${id}`, { method: "PUT", body }); setReferrals((p) => p.map((x) => (x.id === id ? mapReferral(r) : x))); }, [call]);
  const removeReferral = useCallback(async (id) => { await call(`/referrals/${id}`, { method: "DELETE" }); setReferrals((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addGift = useCallback(async (body) => { const r = await call("/gifts", { method: "POST", body }); const doc = doctors.find((d) => d.id === body.doctorId); setGifts((p) => [{ ...mapGift(r), doctorName: doc?.name }, ...p]); }, [call, doctors]);
  const updateGift = useCallback(async (id, body) => { const r = await call(`/gifts/${id}`, { method: "PUT", body }); const doc = doctors.find((d) => d.id === body.doctorId); setGifts((p) => p.map((x) => (x.id === id ? { ...mapGift(r), doctorName: doc?.name } : x))); }, [call, doctors]);
  const removeGift = useCallback(async (id) => { await call(`/gifts/${id}`, { method: "DELETE" }); setGifts((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addExpense = useCallback(async (body) => { const r = await call("/expenses", { method: "POST", body }); setExpenses((p) => [mapExpense(r), ...p]); }, [call]);
  const updateExpense = useCallback(async (id, body) => { const r = await call(`/expenses/${id}`, { method: "PUT", body }); setExpenses((p) => p.map((x) => (x.id === id ? mapExpense(r) : x))); }, [call]);
  const removeExpense = useCallback(async (id) => { await call(`/expenses/${id}`, { method: "DELETE" }); setExpenses((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addAsset = useCallback(async (body) => { const r = await call("/assets", { method: "POST", body }); setAssets((p) => [mapAsset(r), ...p]); }, [call]);
  const updateAsset = useCallback(async (id, body) => { const r = await call(`/assets/${id}`, { method: "PUT", body }); setAssets((p) => p.map((x) => (x.id === id ? mapAsset(r) : x))); }, [call]);
  const removeAsset = useCallback(async (id) => { await call(`/assets/${id}`, { method: "DELETE" }); setAssets((p) => p.filter((x) => x.id !== id)); }, [call]);
  const addCapital = useCallback(async (body) => { const r = await call("/capital", { method: "POST", body }); setCapital((p) => [mapCapital(r), ...p]); }, [call]);
  const removeCapital = useCallback(async (id) => { await call(`/capital/${id}`, { method: "DELETE" }); setCapital((p) => p.filter((x) => x.id !== id)); }, [call]);
  const mapOther = (r) => ({ id: r.id, category: r.category, txnType: r.txn_type, partyName: r.party_name, amount: Number(r.amount), date: d10(r.txn_date), note: r.note });
  const addOtherBalance = useCallback(async (body) => { const r = await call("/other-balance", { method: "POST", body }); const item = mapOther(r); (item.category === "unsecured_loan" ? setLoans : setDeposits)((p) => [item, ...p]); }, [call]);
  const updateOtherBalance = useCallback(async (id, body) => { const r = await call(`/other-balance/${id}`, { method: "PUT", body }); const item = mapOther(r); (item.category === "unsecured_loan" ? setLoans : setDeposits)((p) => p.map((x) => (x.id === id ? item : x))); }, [call]);
  const removeOtherBalance = useCallback(async (id, category) => { await call(`/other-balance/${id}`, { method: "DELETE" }); (category === "unsecured_loan" ? setLoans : setDeposits)((p) => p.filter((x) => x.id !== id)); }, [call]);
  const mapPatientMaster = (r) => ({ id: r.id, name: r.name, mobile: r.mobile || "", gender: r.gender || "", dob: d10(r.dob), address: r.address || "" });
  const addPatientMaster = useCallback(async (body) => { const r = await call("/patient-master", { method: "POST", body }); setPatientsMaster((p) => [mapPatientMaster(r), ...p]); }, [call]);
  const updatePatientMaster = useCallback(async (id, body) => { const r = await call(`/patient-master/${id}`, { method: "PUT", body }); setPatientsMaster((p) => p.map((x) => (x.id === id ? mapPatientMaster(r) : x))); }, [call]);
  const removePatientMaster = useCallback(async (id) => { await call(`/patient-master/${id}`, { method: "DELETE" }); setPatientsMaster((p) => p.filter((x) => x.id !== id)); }, [call]);
  const updateSettings = useCallback(async (body) => { const r = await call("/settings", { method: "PUT", body }); setSettings(mapSettings(r)); }, [call]);

  if (!session) return <AuthScreen onLogin={setSession} origin={origin} setOrigin={setOrigin} />;

  const can = makeCan(session);

  return (
    <ApiContext.Provider value={apiValue}>
      <div className="app-root">
        <style>{`
          ${FONT_IMPORT}
          .app-root{--primary:#714B67;--primary-dark:#4A2F44;--accent:#C9A227;--accent-soft:#F3E3A8;--bg:#F5F8F7;--surface:#FFFFFF;
            --ink:#142524;--ink-soft:#5B6B69;--border:#E1E8E6;--income:#1F8A5F;--expense:#B3423A;
            min-height:100vh;background:var(--bg);font-family:'Inter',sans-serif;color:var(--ink);display:flex;}
          .sidebar{width:230px;background:linear-gradient(180deg,var(--primary),var(--primary-dark));color:#EAF3F1;flex-shrink:0;padding:22px 0;display:flex;flex-direction:column;}
          .sidebar .brand{font-family:'Plus Jakarta Sans',sans-serif;font-weight:800;font-size:20px;padding:0 20px 2px;}
          .sidebar .biz{padding:0 20px 16px;font-size:11.5px;color:#B9D8D2;border-bottom:1px solid rgba(255,255,255,.15);margin-bottom:8px;line-height:1.5;word-break:break-all;}
          .nav-item{text-align:left;background:none;border:none;color:#EAF3F1;padding:10px 20px;font-size:13.5px;font-weight:500;cursor:pointer;border-left:3px solid transparent;opacity:.82;display:flex;align-items:center;gap:10px;}
          .nav-icon{font-size:15px;line-height:1;width:18px;text-align:center;flex-shrink:0;}
          .nav-icon-chip{font-size:13px;line-height:1;width:26px;height:26px;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 1px 3px rgba(0,0,0,.25);}
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
          .launcher-wrap{background:linear-gradient(160deg,#EFEBFA 0%,#F6F3FC 45%,#FDFBFF 100%);border-radius:16px;padding:36px 28px;min-height:calc(100vh - 100px);}
          .launcher-header{margin-bottom:28px;}
          .launcher-header h2{font-family:'Plus Jakarta Sans',sans-serif;font-weight:800;font-size:26px;color:var(--primary-dark);margin:0 0 6px;}
          .launcher-header p{color:var(--ink-soft);font-size:14px;margin:0;}
          .launcher-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:18px;}
          .launcher-tile{display:flex;flex-direction:column;align-items:center;gap:10px;background:#fff;border:1px solid #EAE6F5;border-radius:14px;padding:20px 10px;cursor:pointer;transition:transform .12s ease,box-shadow .12s ease;box-shadow:0 2px 6px rgba(80,60,120,.06);}
          .launcher-tile:hover{transform:translateY(-3px);box-shadow:0 8px 20px rgba(80,60,120,.14);}
          .launcher-icon{width:52px;height:52px;border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:24px;box-shadow:0 3px 8px rgba(0,0,0,.15);}
          .launcher-label{font-size:12.5px;font-weight:600;color:var(--ink);text-align:center;line-height:1.3;}
          @media(max-width:640px){ .launcher-grid{grid-template-columns:repeat(auto-fill,minmax(96px,1fr));gap:12px;} .launcher-wrap{padding:22px 14px;} }
          .period-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-bottom:18px;}
          .week-picker{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin:-4px 0 10px;}
          .week-picker input[type="month"],.week-picker select{font-size:11.5px;padding:4px 6px;border-radius:6px;border:1px solid var(--border);background:#fff;color:var(--ink);}
          .week-range-label{font-size:11px;color:var(--ink-soft);width:100%;}
          .kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:20px;}
          .kpi-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 18px;box-shadow:0 1px 4px rgba(10,40,36,.05);}
          .kpi-card.clickable{cursor:pointer;transition:transform .12s ease,box-shadow .12s ease;}
          .kpi-card.clickable:hover{transform:translateY(-2px);box-shadow:0 6px 16px rgba(10,40,36,.12);}
          .kpi-top{display:flex;align-items:center;gap:8px;margin-bottom:10px;}
          .kpi-icon{font-size:20px;line-height:1;}
          .kpi-label{font-size:10.5px;text-transform:uppercase;letter-spacing:.8px;color:var(--ink-soft);font-weight:700;}
          .kpi-bottom{display:flex;align-items:center;justify-content:space-between;gap:10px;}
          .kpi-value{font-family:'IBM Plex Mono',monospace;font-size:22px;font-weight:700;white-space:nowrap;}
          .section-header{font-family:'Plus Jakarta Sans',sans-serif;font-size:13px;font-weight:800;letter-spacing:.6px;text-transform:uppercase;color:var(--primary-dark);margin:26px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--accent-soft);}
          .section-header:first-child{margin-top:0;}
          .insight-banner{background:linear-gradient(135deg,var(--primary),var(--primary-dark));color:#EAF3F1;border-radius:12px;padding:16px 20px;margin-bottom:20px;font-size:14px;line-height:1.6;display:flex;align-items:center;gap:12px;}
          .insight-banner .insight-icon{font-size:22px;}
          .mode-pill{display:inline-flex;align-items:center;gap:5px;}
          .period-card{margin-bottom:0;}
          .period-card table{font-size:12.5px;}
          .period-card td{padding:5px 4px;}
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
          .info{font-size:12px;color:var(--income);background:#e3efe6;padding:8px 10px;border-radius:6px;margin-top:8px;}
          @media (max-width:820px){
            .app-root{flex-direction:column;} .sidebar{width:100%;flex-direction:row;flex-wrap:wrap;padding:14px 12px;align-items:center;}
            .sidebar .brand{padding:0 10px 0 0;font-size:17px;} .sidebar .biz{display:none;}
            .nav-item{padding:7px 9px;font-size:11.5px;border-left:none;border-bottom:3px solid transparent;}
            .nav-item.active{border-left:none;border-bottom-color:var(--accent);} .logout{margin-top:0;margin-left:auto;padding:8px 10px;}
            .content{padding:16px 12px 50px;}
          }
          .custom-export{margin-top:10px;}
          .suggest-list{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid var(--border);border-radius:8px;margin-top:4px;box-shadow:0 8px 20px rgba(10,40,36,.12);z-index:10;max-height:220px;overflow-y:auto;}
          .suggest-item{padding:9px 12px;font-size:13.5px;cursor:pointer;border-bottom:1px solid var(--border);}
          .suggest-item:last-child{border-bottom:none;}
          .suggest-item:hover{background:#F5F8F7;}
          .custom-export-panel{display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap;margin-top:8px;padding:10px;background:#F5F8F7;border-radius:8px;}
          .custom-export-panel label{display:block;font-size:10px;text-transform:uppercase;color:var(--ink-soft);font-weight:700;margin-bottom:3px;}
          .print-root{display:none;}
          @media print{
            .sidebar,.topbar .fy-select,.no-print{display:none !important;} .content{padding:0;max-width:100%;} body{background:#fff;} .card{box-shadow:none;border:1px solid #ccc;}
            body.printing-custom .content > *{display:none !important;}
            body.printing-custom .print-root{display:block !important;}
          }
        `}</style>

        <nav className="sidebar">
          <div className="brand" style={{ cursor: "pointer" }} onClick={() => setView("launcher")} title="Back to app launcher">GANATRA CLINIC</div>
          <div className="biz">{settings.proprietor}<br />{session.name} ({session.userId}) · {session.role}<br /><span style={{ opacity: .7 }}>{origin}</span></div>
          <button className={"nav-item" + (view === "launcher" ? " active" : "")} onClick={() => setView("launcher")}><span className="nav-icon-chip" style={{ background: `linear-gradient(135deg, #F4A340, #E85D3D)` }}>🔷</span>Home</button>
          {NAV.filter((n) => !n.adminOnly || session.role === "Admin").filter((n) => !n.module || can(n.module, "view")).map((n) => (<button key={n.key} className={"nav-item" + (view === n.key ? " active" : "")} onClick={() => setView(n.key)}><span className="nav-icon-chip" style={{ background: `linear-gradient(135deg, ${n.grad[0]}, ${n.grad[1]})` }}>{n.icon}</span>{n.label}</button>))}
          <button className="logout" onClick={() => setSession(null)}>Log out</button>
        </nav>

        <div className="main">
          <div className="topbar" style={view === "launcher" ? { visibility: "hidden", height: 0, overflow: "hidden", padding: 0, margin: 0, border: "none" } : undefined}>
            <h1>{NAV.find((n) => n.key === view)?.label || "Home"}</h1>
            <select className="fy-select no-print" value={fy} onChange={(e) => setFy(e.target.value)}>{last4FYs().map((f) => <option key={f} value={f}>FY {f}</option>)}</select>
          </div>
          <div className="content">
            {loading ? <div className="empty">Loading clinic records from the server…</div> : loadError ? (
              <div className="card"><h2>Couldn't load your data</h2><ErrorNote msg={loadError} /><button className="btn" style={{ marginTop: 10 }} onClick={reloadAll} type="button">Retry</button></div>
            ) : (
              <>
                {view === "launcher" && <LauncherGrid settings={settings} session={session} can={can} setView={setView} />}
                {view === "dashboard" && <Dashboard settings={settings} collections={collections} referrals={referrals} expenses={expenses} doctorPays={doctorPays} cases={cases} fy={fy} setView={setView} />}
                {view === "cases" && can("cases", "view") && <CaseRecords cases={cases} addCase={addCase} updateCase={updateCase} removeCase={removeCase} doctors={doctors} patientsMaster={patientsMaster} can={can} />}
                {view === "patientMaster" && can("cases", "view") && <PatientMaster can={can} patients={patientsMaster} addPatient={addPatientMaster} updatePatient={updatePatientMaster} removePatient={removePatientMaster} />}
                {view === "patients" && can("cases", "view") && <PatientHistory can={can} updateCase={updateCase} updateCollection={updateCollection} cases={cases} doctors={doctors} />}
                {view === "collections" && can("collections", "view") && <Collections collections={collections} addCollection={addCollection} updateCollection={updateCollection} removeCollection={removeCollection} cases={cases} fy={fy} can={can} />}
                {view === "doctors" && can("doctorPay", "view") && <DoctorShifts doctors={doctors} addDoctor={addDoctor} updateDoctor={updateDoctor} removeDoctor={removeDoctor} doctorPays={doctorPays} addDoctorPay={addDoctorPay} updateDoctorPay={updateDoctorPay} removeDoctorPay={removeDoctorPay} can={can} />}
                {view === "referrals" && can("referrals", "view") && <Referrals referrals={referrals} addReferral={addReferral} updateReferral={updateReferral} removeReferral={removeReferral} fy={fy} can={can} />}
                {view === "gifts" && can("gifts", "view") && <Gifts gifts={gifts} addGift={addGift} updateGift={updateGift} removeGift={removeGift} doctors={doctors} can={can} />}
                {view === "expenses" && can("expenses", "view") && <Expenses expenses={expenses} addExpense={addExpense} updateExpense={updateExpense} removeExpense={removeExpense} fy={fy} can={can} />}
                {view === "assets" && can("assets", "view") && <FixedAssets assets={assets} addAsset={addAsset} updateAsset={updateAsset} removeAsset={removeAsset} fy={fy} can={can} />}
                {view === "statements" && can("statements", "view") && <FinancialStatements fy={fy} settings={settings} can={can} />}
                {view === "auditLog" && can("auditLog", "view") && <AccessReport can={can} />}
                {view === "settings" && <SettingsPage settings={settings} updateSettings={updateSettings} session={session} origin={origin} capital={capital} addCapital={addCapital} removeCapital={removeCapital} loans={loans} deposits={deposits} addOtherBalance={addOtherBalance} updateOtherBalance={updateOtherBalance} removeOtherBalance={removeOtherBalance} can={can} />}
                {view === "admin" && session.role === "Admin" && <AdminUsers />}
              </>
            )}
            <PrintRoot />
          </div>
        </div>
      </div>
    </ApiContext.Provider>
  );
}

/* ============================== DASHBOARD ============================== */
/** Computes mode-wise collection, outstanding dues, expenses, and net
 *  profit for whatever date range [start,end] is passed in. Net profit here
 *  = total revenue billed (collected + still outstanding) minus expenses
 *  logged in that same window — a quick operational view, not the full
 *  Income Statement (which also nets off doctor pay and depreciation). */
function periodSummary(collections, expenses, start, end) {
  const rows = collections.filter((c) => c.date >= start && c.date <= end);
  const modeTotals = {};
  COLLECTION_MODES.forEach((m) => { modeTotals[m] = 0; });
  let outstanding = 0;
  rows.forEach((c) => {
    modeTotals[c.mode || "Other"] = (modeTotals[c.mode || "Other"] || 0) + Number(c.amountCollected || 0);
    outstanding += Number(c.balance || 0);
  });
  const expenseTotal = expenses.filter((e) => e.date >= start && e.date <= end).reduce((s, e) => s + Number(e.amount || 0), 0);
  const collectedTotal = Object.values(modeTotals).reduce((s, v) => s + v, 0);
  const netProfit = collectedTotal + outstanding - expenseTotal;
  return { modeTotals, outstanding, expenseTotal, netProfit, collectedTotal };
}

const MODE_ICONS = { Cash: "💵", UPI: "📱", Card: "💳", Other: "🧾" };

/** A tiny inline trend line — no axes, no grid, just the shape of the last
 *  N values. Uses Recharts (already a dependency) so this adds no new
 *  packages and no new build risk. */
function Sparkline({ data, color = "#714B67", height = 34, width = 90 }) {
  if (!data || data.length < 2) return <div style={{ width, height }} />;
  return (
    <ResponsiveContainer width={width} height={height}>
      <LineChart data={data}>
        <Line type="monotone" dataKey="v" stroke={color} strokeWidth={2} dot={false} isAnimationActive={false} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function KpiCard({ icon, label, value, color, sparkline, sparklineColor, hint, onClick }) {
  return (
    <div className={"kpi-card" + (onClick ? " clickable" : "")} title={onClick ? `${hint || ""} — click for a detailed, exportable report`.trim() : hint} onClick={onClick} role={onClick ? "button" : undefined}>
      <div className="kpi-top">
        <span className="kpi-icon">{icon}</span>
        <span className="kpi-label">{label}</span>
      </div>
      <div className="kpi-bottom">
        <span className="kpi-value" style={{ color }}>{value}</span>
        {sparkline && <Sparkline data={sparkline} color={sparklineColor || color} />}
      </div>
    </div>
  );
}

function formatDate(iso) {
  if (!iso) return "";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}

function PeriodCard({ title, summary, start, end, onDrill, headerExtra }) {
  const Row = ({ label, value, kind, mode, bold, borderTop, color }) => (
    <tr
      style={{ fontWeight: bold ? 700 : 400, borderTop: borderTop ? `${borderTop} solid var(--border)` : undefined, cursor: "pointer" }}
      onClick={() => onDrill({ kind, mode, start, end, label: title })}
    >
      <td style={{ textDecoration: "underline", textDecorationStyle: "dotted", textDecorationColor: "var(--ink-soft)" }}>{mode ? <span className="mode-pill">{MODE_ICONS[mode]} {label}</span> : label}</td>
      <td className="num" style={{ color }}>{inr(value)}</td>
    </tr>
  );
  return (
    <div className="card period-card">
      <h2>{title}</h2>
      {headerExtra}
      <table>
        <tbody>
          {COLLECTION_MODES.map((m) => (
            <Row key={m} label={m} value={summary.modeTotals[m] || 0} kind="mode" mode={m} />
          ))}
          <Row label="Collected total" value={summary.collectedTotal} kind="collected" bold borderTop="1px" />
          <Row label="+ Outstanding due" value={summary.outstanding} kind="outstanding" />
          <Row label="− Expenses" value={summary.expenseTotal} kind="expenses" color="var(--expense)" />
          <tr style={{ fontWeight: 700, borderTop: "2px solid var(--accent)" }}>
            <td>{summary.netProfit >= 0 ? "📈" : "📉"} Net Profit</td>
            <td className="num" style={{ color: summary.netProfit >= 0 ? "var(--income)" : "var(--expense)" }}>{inr(summary.netProfit)}</td>
          </tr>
        </tbody>
      </table>
      <div style={{ fontSize: 10.5, color: "var(--ink-soft)", marginTop: 8 }}>Click any row for a detailed, exportable report.</div>
    </div>
  );
}

/** The drill-down report panel: shows the actual records behind whatever
 *  row was clicked, with its own Excel and PDF export. */
function DrillDownPanel({ drill, onClose }) {
  const { origin, token } = useApi();
  if (!drill) return null;
  const isExpense = drill.kind === "expenses";
  const columns = isExpense
    ? [{ label: "Date", value: (r) => r.date }, { label: "Category", value: (r) => r.category }, { label: "Narration", value: (r) => r.narration }, { label: "Amount", value: (r) => inr(r.amount) }]
    : [{ label: "Date", value: (r) => r.date }, { label: "Case No.", value: (r) => r.caseNo || "—" }, { label: "Patient", value: (r) => r.patientName }, { label: "Shift", value: (r) => r.shift || "—" }, { label: "Doctor", value: (r) => r.doctorName || "—" }, { label: "Mode", value: (r) => r.mode }, { label: "Due", value: (r) => inr(r.amountDue) }, { label: "Collected", value: (r) => inr(r.amountCollected) }, { label: "Balance", value: (r) => inr(r.balance) }];
  const numericLabels = ["Amount", "Due", "Collected", "Balance"];
  const total = isExpense ? drill.rows.reduce((s, r) => s + Number(r.amount || 0), 0) : drill.rows.reduce((s, r) => s + Number(r.amountCollected || 0), 0);

  const viewImage = async (path) => {
    if (!path) return;
    try {
      const res = await fetch(`${origin}${path}`, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) throw new Error("not found");
      const blob = await res.blob();
      window.open(URL.createObjectURL(blob), "_blank");
    } catch { alert("Could not load that photo."); }
  };

  const doExcel = () => exportExcel(`drilldown-${drill.title.replace(/[^a-z0-9]+/gi, "-").toLowerCase()}`, {
    Detail: drill.rows.map((r) => Object.fromEntries(columns.map((c) => [c.label, c.value(r)]))),
  });
  const doPrint = () => {
    const win = document.getElementById("print-root");
    if (!win) { window.print(); return; }
    const rowsHtml = drill.rows.map((r) => `<tr>${columns.map((c) => `<td>${escapeHtml(c.value(r))}</td>`).join("")}</tr>`).join("");
    win.innerHTML = `
      <h2>${escapeHtml(drill.title)}</h2>
      <p style="color:#5B6B69;font-size:12px;">${drill.rows.length} record(s) — total ${inr(total)}</p>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr>${columns.map((c) => `<th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">${escapeHtml(c.label)}</th>`).join("")}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    document.body.classList.add("printing-custom");
    window.print();
    setTimeout(() => { document.body.classList.remove("printing-custom"); win.innerHTML = ""; }, 300);
  };

  return (
    <div className="card" style={{ borderColor: "var(--accent)", borderWidth: 2 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <h2>{drill.title} — {drill.rows.length} record(s), total {inr(total)}</h2>
        <button className="btn secondary small no-print" type="button" onClick={onClose}>✕ Close</button>
      </div>
      {drill.rows.length === 0 ? <div className="empty">No records in this range.</div> : (
        <table>
          <thead><tr>{columns.map((c) => <th key={c.label} className={numericLabels.includes(c.label) ? "num" : ""}>{c.label}</th>)}<th className="no-print">Photo</th></tr></thead>
          <tbody>{drill.rows.map((r, i) => (
            <tr key={i}>
              {columns.map((c) => <td key={c.label} className={numericLabels.includes(c.label) ? "num" : ""}>{c.value(r)}</td>)}
              <td className="no-print">{r.image ? <button className="btn secondary small" type="button" onClick={() => viewImage(r.image)}>📎 View</button> : "—"}</td>
            </tr>
          ))}</tbody>
        </table>
      )}
      <div className="export-row no-print">
        <button className="btn secondary small" type="button" onClick={doExcel}>⬇ Export Excel</button>
        <button className="btn secondary small" type="button" onClick={doPrint}>⎙ Export PDF</button>
      </div>
    </div>
  );
}

/** Day-wise Morning vs Evening collection chart, with a customizable date
 *  range and its own Excel/PDF report. Shift comes from the linked case
 *  record (collections don't carry shift directly); anything not linked to
 *  a case shows as "Unlinked" rather than being silently dropped. */
function ShiftCollectionChart({ collections, cases, fy }) {
  const monthStart = todayISO().slice(0, 8) + "01";
  const [from, setFrom] = useState(monthStart);
  const [to, setTo] = useState(todayISO());
  const [open, setOpen] = useState(false);

  const caseById = useMemo(() => Object.fromEntries(cases.map((c) => [c.id, c])), [cases]);

  const dayRows = useMemo(() => {
    const map = {};
    let d = new Date(from + "T00:00:00");
    const end = new Date(to + "T00:00:00");
    while (d <= end) { map[d.toISOString().slice(0, 10)] = { date: d.toISOString().slice(5, 10), Morning: 0, Evening: 0, Unlinked: 0 }; d.setDate(d.getDate() + 1); }
    collections.filter((c) => c.date >= from && c.date <= to).forEach((c) => {
      const row = map[c.date];
      if (!row) return;
      const linked = c.caseId ? caseById[c.caseId] : null;
      const key = linked?.shift === "Morning" ? "Morning" : linked?.shift === "Evening" ? "Evening" : "Unlinked";
      row[key] += Number(c.amountCollected || 0);
    });
    return Object.values(map);
  }, [collections, caseById, from, to]);

  const totals = dayRows.reduce((s, r) => ({ Morning: s.Morning + r.Morning, Evening: s.Evening + r.Evening, Unlinked: s.Unlinked + r.Unlinked }), { Morning: 0, Evening: 0, Unlinked: 0 });

  const doExcel = () => exportExcel("morning-evening-collection", { DayWise: dayRows.map((r) => ({ Date: r.date, Morning: r.Morning, Evening: r.Evening, Unlinked: r.Unlinked, Total: r.Morning + r.Evening + r.Unlinked })) });
  const doPrint = () => {
    const win = document.getElementById("print-root");
    if (!win) { window.print(); return; }
    const rowsHtml = dayRows.map((r) => `<tr><td>${r.date}</td><td>${inr(r.Morning)}</td><td>${inr(r.Evening)}</td><td>${inr(r.Unlinked)}</td><td>${inr(r.Morning + r.Evening + r.Unlinked)}</td></tr>`).join("");
    win.innerHTML = `
      <h2>Morning vs Evening Collection</h2>
      <p style="color:#5B6B69;font-size:12px;">${from} to ${to} — Morning ${inr(totals.Morning)}, Evening ${inr(totals.Evening)}, Unlinked ${inr(totals.Unlinked)}</p>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Date</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Morning</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Evening</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Unlinked</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Total</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    document.body.classList.add("printing-custom");
    window.print();
    setTimeout(() => { document.body.classList.remove("printing-custom"); win.innerHTML = ""; }, 300);
  };

  return (
    <div className="card">
      <h2>Morning vs Evening Collection — {from} to {to}</h2>
      <p style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: -8 }}>Morning {inr(totals.Morning)} · Evening {inr(totals.Evening)} · Unlinked {inr(totals.Unlinked)} (collections not tied to a case record, so shift is unknown)</p>
      <ResponsiveContainer width="100%" height={260}>
        <BarChart data={dayRows}>
          <CartesianGrid strokeDasharray="3 3" stroke="#E1E8E6" />
          <XAxis dataKey="date" fontSize={10} interval={dayRows.length > 20 ? Math.floor(dayRows.length / 20) : 0} />
          <YAxis fontSize={11} />
          <Tooltip formatter={(v) => inr(v)} />
          <Legend />
          <Bar dataKey="Morning" fill="#C9A227" radius={[3, 3, 0, 0]} />
          <Bar dataKey="Evening" fill="#714B67" radius={[3, 3, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
      <button className="btn secondary small no-print" type="button" onClick={() => setOpen((o) => !o)} style={{ marginTop: 10 }}>📅 Customise date range</button>
      {open && (
        <div className="custom-export-panel">
          <div><label>From</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></div>
          <div><label>To</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></div>
        </div>
      )}
      <div className="export-row no-print">
        <button className="btn secondary small" type="button" onClick={doExcel}>⬇ Export Excel</button>
        <button className="btn secondary small" type="button" onClick={doPrint}>⎙ Export PDF</button>
      </div>
    </div>
  );
}

/** A colorful grid-tile home screen (Odoo-style app launcher) — each module
 *  gets its own gradient-colored icon tile. Purely a navigation surface; it
 *  doesn't fetch or show any data itself, so it carries no security surface
 *  beyond the same view-permission filtering already applied everywhere else. */
function LauncherGrid({ settings, session, can, setView }) {
  const tiles = NAV.filter((n) => !n.adminOnly || session.role === "Admin").filter((n) => !n.module || can(n.module, "view"));
  return (
    <div className="launcher-wrap">
      <div className="launcher-header">
        <h2>{settings.clinicName || "Ganatra Clinic"}</h2>
        <p>Welcome back, {session.name.split(" ")[0]} — pick where you want to go.</p>
      </div>
      <div className="launcher-grid">
        {tiles.map((n) => (
          <button key={n.key} className="launcher-tile" onClick={() => setView(n.key)}>
            <span className="launcher-icon" style={{ background: `linear-gradient(135deg, ${n.grad[0]}, ${n.grad[1]})` }}>{n.icon}</span>
            <span className="launcher-label">{n.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

function Dashboard({ settings, collections, referrals, expenses, doctorPays, cases, fy, setView }) {
  const { call } = useApi();
  const [income, setIncome] = useState(null);
  useEffect(() => { call(`/statements/income?fy=${fy}`).then(setIncome).catch(() => setIncome(null)); }, [call, fy]);
  const [drill, setDrill] = useState(null);

  const t = todayISO();
  const yest = (() => { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().slice(0, 10); })();
  const twoDaysAgo = (() => { const d = new Date(); d.setDate(d.getDate() - 2); return d.toISOString().slice(0, 10); })();
  const monthStart = t.slice(0, 8) + "01";
  const monthLabel = new Date(t + "T00:00:00").toLocaleDateString("en-IN", { month: "long", year: "numeric" });
  const range = fyRange(fy);

  // Week-of-month selector: "Week 1" = days 1–7, "Week 2" = days 8–14, and so
  // on, for whichever month the person picks — not a fixed rolling 7 days.
  const [weekMonth, setWeekMonth] = useState(t.slice(0, 7)); // "YYYY-MM"
  const [weekNum, setWeekNum] = useState(Math.min(5, Math.ceil(Number(t.slice(8, 10)) / 7)));
  const daysInWeekMonth = new Date(Number(weekMonth.slice(0, 4)), Number(weekMonth.slice(5, 7)), 0).getDate();
  const weekOptions = Array.from({ length: Math.ceil(daysInWeekMonth / 7) }, (_, i) => i + 1);
  const weekStartDay = (weekNum - 1) * 7 + 1;
  const weekEndDay = Math.min(weekNum * 7, daysInWeekMonth);
  const weekStart = `${weekMonth}-${String(weekStartDay).padStart(2, "0")}`;
  const weekEnd = `${weekMonth}-${String(weekEndDay).padStart(2, "0")}`;
  const weekMonthLabel = new Date(`${weekMonth}-01T00:00:00`).toLocaleDateString("en-IN", { month: "short", year: "numeric" });

  const todaySummary = periodSummary(collections, expenses, t, t);
  const yestSummary = periodSummary(collections, expenses, yest, yest);
  const twoDaysAgoSummary = periodSummary(collections, expenses, twoDaysAgo, twoDaysAgo);
  const weekSummary = periodSummary(collections, expenses, weekStart, weekEnd);
  // Always the true "last 7 days," independent of the Week-of-month dropdown above —
  // used only for the top insight banner and KPI cards, which should never silently
  // change meaning just because someone picked a different week to look back at.
  const rollingWeekStart = (() => { const d = new Date(); d.setDate(d.getDate() - 6); return d.toISOString().slice(0, 10); })();
  const rollingWeekSummary = periodSummary(collections, expenses, rollingWeekStart, t);
  const monthSummary = periodSummary(collections, expenses, monthStart, t);
  const yearSummary = periodSummary(collections, expenses, range.start, range.end);

  const caseById = useMemo(() => Object.fromEntries(cases.map((c) => [c.id, c])), [cases]);
  const enrichCollection = (c) => { const linked = c.caseId ? caseById[c.caseId] : null; return { ...c, shift: linked?.shift || "", doctorName: linked?.doctorName || "" }; };

  const openDrill = ({ kind, mode, start, end, label }) => {
    const inRange = (d) => (!start || d >= start) && (!end || d <= end);
    if (kind === "expenses") {
      setDrill({ title: `Expenses — ${label}`, kind: "expenses", rows: expenses.filter((e) => inRange(e.date)) });
      return;
    }
    let rows = collections.filter((c) => inRange(c.date)).map(enrichCollection);
    let title = `All Collections — ${label}`;
    if (kind === "mode") { rows = rows.filter((c) => (c.mode || "Other") === mode); title = `${mode} Collections — ${label}`; }
    else if (kind === "outstanding") { rows = rows.filter((c) => Number(c.balance || 0) > 0); title = `Outstanding Due — ${label}`; }
    setDrill({ title, kind: "collections", rows });
  };

  const netProfit = income ? income.netProfit : null;
  const outstanding = collections.reduce((a, c) => a + Number(c.balance || 0), 0);

  const last30 = [];
  for (let i = 29; i >= 0; i--) { const dt = new Date(); dt.setDate(dt.getDate() - i); const iso = dt.toISOString().slice(0, 10); const s = periodSummary(collections, [], iso, iso); last30.push({ date: iso.slice(5), amount: s.collectedTotal + s.outstanding }); }
  const expenseByCat = {};
  expenses.filter((e) => e.date >= range.start && e.date <= range.end).forEach((e) => { expenseByCat[e.category] = (expenseByCat[e.category] || 0) + Number(e.amount); });
  const pieData = Object.entries(expenseByCat).map(([name, value]) => ({ name, value }));

  const collectionPieData = [
    ...COLLECTION_MODES.map((m) => ({ name: m, value: yearSummary.modeTotals[m] || 0 })),
    { name: "Outstanding Due", value: yearSummary.outstanding },
  ].filter((x) => x.value > 0);

  // Previous 7 days (before this week), for the "up/down vs last week" insight
  const prevWeekEnd = (() => { const d = new Date(); d.setDate(d.getDate() - 7); return d.toISOString().slice(0, 10); })();
  const prevWeekStart = (() => { const d = new Date(); d.setDate(d.getDate() - 13); return d.toISOString().slice(0, 10); })();
  const prevWeekSummary = periodSummary(collections, expenses, prevWeekStart, prevWeekEnd);
  const weekChangePct = prevWeekSummary.collectedTotal > 0
    ? Math.round(((rollingWeekSummary.collectedTotal - prevWeekSummary.collectedTotal) / prevWeekSummary.collectedTotal) * 100)
    : null;

  // Sparkline series: last 14 days of collected+expense, for the KPI cards
  const last14Spark = [];
  for (let i = 13; i >= 0; i--) { const dt = new Date(); dt.setDate(dt.getDate() - i); const iso = dt.toISOString().slice(0, 10); const s = periodSummary(collections, expenses, iso, iso); last14Spark.push({ collected: s.collectedTotal, expense: s.expenseTotal, net: s.netProfit }); }
  const sparkCollected = last14Spark.map((s) => ({ v: s.collected }));
  const sparkExpense = last14Spark.map((s) => ({ v: s.expense }));
  const sparkOutstanding = last14Spark.map(() => ({ v: outstanding })); // flat — overall figure, shown for shape consistency

  // Doctor-wise revenue for the FY, via each collection's linked case
  const doctorRevenue = useMemo(() => {
    const totals = {};
    collections.filter((c) => c.date >= range.start && c.date <= range.end).forEach((c) => {
      const linked = c.caseId ? caseById[c.caseId] : null;
      const name = linked?.doctorName || "Unassigned";
      totals[name] = (totals[name] || 0) + Number(c.amountDue || 0);
    });
    return Object.entries(totals).sort((a, b) => b[1] - a[1]);
  }, [collections, caseById, range.start, range.end]);

  return (
    <div>
      <div className="insight-banner">
        <span className="insight-icon">{weekChangePct === null ? "👋" : weekChangePct >= 0 ? "📈" : "📉"}</span>
        <span>
          {weekChangePct === null
            ? <>Welcome back — here's how {settings.clinicName || "the clinic"} is doing right now.</>
            : <>Collections are <b>{weekChangePct >= 0 ? "up" : "down"} {Math.abs(weekChangePct)}%</b> this week vs last week.</>}
          {" "}Outstanding dues stand at <b>{inr(outstanding)}</b>.
          {expenseByCat && Object.keys(expenseByCat).length > 0 && <> Expenses this FY: <b>{inr(Object.values(expenseByCat).reduce((s, v) => s + v, 0))}</b>.</>}
        </span>
      </div>

      <div className="section-header">Key Numbers</div>
      <div className="kpi-grid">
        <KpiCard icon="⚠️" label="Outstanding Dues" value={inr(outstanding)} color="var(--accent)" hint="Total unpaid balance across every collection entry, all time" onClick={() => openDrill({ kind: "outstanding", label: "All Time" })} />
        <KpiCard icon={(netProfit ?? 0) >= 0 ? "📈" : "📉"} label={`Net Profit (FY ${fy})`} value={netProfit === null ? "…" : inr(netProfit)} color={(netProfit ?? 0) >= 0 ? "#1F8A5F" : "#B3423A"} sparkline={last14Spark.map((s) => ({ v: s.net }))} sparklineColor={(netProfit ?? 0) >= 0 ? "#1F8A5F" : "#B3423A"} hint="Full P&L net profit for the financial year — click for the full Income Statement, with its own custom-range PDF/Excel export" onClick={() => setView("statements")} />
        <KpiCard icon="💰" label="This Week's Collection" value={inr(rollingWeekSummary.collectedTotal)} color="var(--primary-dark)" sparkline={sparkCollected} hint="Cash actually collected in the last 7 days" onClick={() => openDrill({ kind: "collected", start: rollingWeekStart, end: t, label: "This Week" })} />
        <KpiCard icon="🧾" label="This Week's Expenses" value={inr(rollingWeekSummary.expenseTotal)} color="var(--expense)" sparkline={sparkExpense} sparklineColor="var(--expense)" hint="Expenses logged in the last 7 days" onClick={() => openDrill({ kind: "expenses", start: rollingWeekStart, end: t, label: "This Week" })} />
      </div>

      {drill && <DrillDownPanel drill={drill} onClose={() => setDrill(null)} />}

      <div className="section-header">Daily &amp; Periodic Snapshot</div>
      <div className="period-grid">
        <PeriodCard title={formatDate(t)} summary={todaySummary} start={t} end={t} onDrill={openDrill} />
        <PeriodCard title={formatDate(yest)} summary={yestSummary} start={yest} end={yest} onDrill={openDrill} />
        <PeriodCard title={formatDate(twoDaysAgo)} summary={twoDaysAgoSummary} start={twoDaysAgo} end={twoDaysAgo} onDrill={openDrill} />
        <PeriodCard
          title={`Week ${weekNum} — ${weekMonthLabel}`} summary={weekSummary} start={weekStart} end={weekEnd} onDrill={openDrill}
          headerExtra={
            <div className="week-picker no-print">
              <input type="month" value={weekMonth} onChange={(e) => { setWeekMonth(e.target.value); setWeekNum(1); }} />
              <select value={weekNum} onChange={(e) => setWeekNum(Number(e.target.value))}>
                {weekOptions.map((w) => <option key={w} value={w}>Week {w}</option>)}
              </select>
              <span className="week-range-label">{formatDate(weekStart)} – {formatDate(weekEnd)}</span>
            </div>
          }
        />
        <PeriodCard title={monthLabel} summary={monthSummary} start={monthStart} end={t} onDrill={openDrill} />
        <PeriodCard title={`FY ${fy}`} summary={yearSummary} start={range.start} end={range.end} onDrill={openDrill} />
      </div>

      <div className="section-header">Trends</div>
      <div className="card">
        <h2>📈 Collection trend — last 30 days</h2>
        <ResponsiveContainer width="100%" height={220}>
          <AreaChart data={last30}>
            <defs><linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#714B67" stopOpacity={0.35} /><stop offset="100%" stopColor="#714B67" stopOpacity={0.02} /></linearGradient></defs>
            <CartesianGrid strokeDasharray="3 3" stroke="#E1E8E6" /><XAxis dataKey="date" fontSize={11} /><YAxis fontSize={11} />
            <Tooltip formatter={(v) => inr(v)} /><Area type="monotone" dataKey="amount" stroke="#714B67" fill="url(#cg)" strokeWidth={2} />
          </AreaChart>
        </ResponsiveContainer>
      </div>
      <ShiftCollectionChart collections={collections} cases={cases} fy={fy} />

      <div className="section-header">Breakdowns</div>
      <div className="card">
        <h2>🥧 Collection breakdown — FY {fy} (mode &amp; outstanding dues)</h2>
        {collectionPieData.length === 0 ? <div className="empty">No collections logged yet for this financial year.</div> : (
          <ResponsiveContainer width="100%" height={240}>
            <PieChart><Pie data={collectionPieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>{collectionPieData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}</Pie><Legend /><Tooltip formatter={(v) => inr(v)} /></PieChart>
          </ResponsiveContainer>
        )}
      </div>
      <div className="card">
        <h2>🧾 Expense breakdown — FY {fy}</h2>
        {pieData.length === 0 ? <div className="empty">No expenses logged yet for this financial year.</div> : (
          <ResponsiveContainer width="100%" height={240}>
            <PieChart><Pie data={pieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={2}>{pieData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}</Pie><Legend /><Tooltip formatter={(v) => inr(v)} /></PieChart>
          </ResponsiveContainer>
        )}
      </div>

      <div className="section-header">Doctor Performance</div>
      <div className="card">
        <h2>👩‍⚕️ Revenue by doctor — FY {fy}</h2>
        {doctorRevenue.length === 0 ? <div className="empty">No collections linked to a doctor yet this financial year.</div> : (
          <table>
            <thead><tr><th>Doctor</th><th className="num">Billed revenue</th></tr></thead>
            <tbody>{doctorRevenue.map(([name, amt]) => (<tr key={name}><td>{name}</td><td className="num">{inr(amt)}</td></tr>))}</tbody>
          </table>
        )}
        <div className="note-box">Based on Amount Due for collections linked to a case (shift/doctor comes from the linked case record) — "Unassigned" means the collection isn't linked to a case, or the case has no doctor selected.</div>
      </div>
    </div>
  );
}

/* ============================== CASE RECORDS ============================== */
function CaseRecords({ cases, addCase, updateCase, removeCase, doctors, patientsMaster, can }) {
  const blank = { date: todayISO(), patientName: "", phone: "", briefHistory: "", doctorId: doctors[0]?.id || "", shift: "Morning", externalPrescription: "", image: null };
  const [form, setForm] = useState(blank);
  const [meds, setMeds] = useState([{ name: "", qty: "", price: "" }]);
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [nameSuggestOpen, setNameSuggestOpen] = useState(false);
  const [phoneSuggestOpen, setPhoneSuggestOpen] = useState(false);
  const nameSuggestions = form.patientName.trim().length >= 2
    ? patientsMaster.filter((p) => p.name.toLowerCase().includes(form.patientName.toLowerCase())).slice(0, 8) : [];
  const phoneSuggestions = form.phone.trim().length >= 3
    ? patientsMaster.filter((p) => p.mobile.includes(form.phone)).slice(0, 8) : [];
  const applyPatient = (p) => { setForm({ ...form, patientName: p.name, phone: p.mobile }); setNameSuggestOpen(false); setPhoneSuggestOpen(false); };

  const addMedRow = () => setMeds([...meds, { name: "", qty: "", price: "" }]);
  const updMed = (i, f, v) => { const n = [...meds]; n[i] = { ...n[i], [f]: v }; setMeds(n); };
  const rmMed = (i) => setMeds(meds.filter((_, idx) => idx !== i));

  const [editingId, setEditingId] = useState(null);
  const startEdit = (c) => {
    setEditingId(c.id);
    setForm({ date: c.date, patientName: c.patientName, phone: c.phone || "", briefHistory: c.briefHistory || "", doctorId: c.doctorId || "", shift: c.shift || "Morning", externalPrescription: c.externalPrescription || "", image: c.image || null });
    setMeds(c.medicines && c.medicines.length ? c.medicines.map((m) => ({ name: m.name, qty: String(m.qty), price: String(m.price) })) : [{ name: "", qty: "", price: "" }]);
    setErr("");
  };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setMeds([{ name: "", qty: "", price: "" }]); setErr(""); };
  const save = async () => {
    setErr("");
    if (!form.patientName.trim()) { setErr("Enter the patient's name."); return; }
    const medicines = meds.filter((m) => m.name.trim()).map((m) => ({ name: m.name, qty: Number(m.qty) || 0, price: Number(m.price) || 0 }));
    setBusy(true);
    const payload = { date: form.date, patientName: form.patientName, phone: form.phone, briefHistory: form.briefHistory, doctorId: form.doctorId || null, shift: form.shift, externalPrescription: form.externalPrescription, imageUrl: form.image, medicines };
    try {
      if (editingId) { await updateCase(editingId, payload); setEditingId(null); }
      else { await addCase(payload); }
      setForm(blank); setMeds([{ name: "", qty: "", price: "" }]);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeCase(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };
  const medValue = (c) => (c.medicines || []).reduce((s, m) => s + m.qty * m.price, 0);
  const caseColumns = { caseNo: (c) => c.caseNo, date: (c) => c.date, patientName: (c) => c.patientName, doctorName: (c) => c.doctorName || "", shift: (c) => c.shift || "", briefHistory: (c) => c.briefHistory || "", medValue: (c) => medValue(c) };
  const { sorted: sortedCases, Th: CaseTh } = useSortableRows(cases, caseColumns, "date", "desc");

  // ---- Bulk upload: download a template, then parse a filled-in copy ----
  const TEMPLATE_HEADERS = ["Date (YYYY-MM-DD)", "Patient Name", "Phone", "Doctor", "Shift (Morning/Evening)", "Brief Medical History", "Medicine Name", "Quantity", "Indicative Unit Price", "Amount Due (₹, optional)", "Mode of Payment (Cash/UPI/Card/Other, optional)"];
  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    const example = [
      ["2026-07-13", "Ramesh Patel", "9825012345", doctors[0]?.name || "Dr. Bhavisha Pratik Ganatra", "Morning", "Fever, body ache — 3 days", "Paracetamol 650", 10, 2, 500, "Cash"],
      ["2026-07-13", "Ramesh Patel", "9825012345", doctors[0]?.name || "Dr. Bhavisha Pratik Ganatra", "Morning", "Fever, body ache — 3 days", "ORS sachets", 5, 10, "", ""],
      ["2026-07-13", "Sunita Mehta", "9825098765", doctors[0]?.name || "Dr. Bhavisha Pratik Ganatra", "Evening", "Routine checkup", "", "", "", "", ""],
    ];
    const ws = XLSX.utils.aoa_to_sheet([TEMPLATE_HEADERS, ...example]);
    XLSX.utils.book_append_sheet(wb, ws, "Case Records");
    XLSX.writeFile(wb, "case-records-upload-template.xlsx");
  };

  const normalizeDate = (v) => {
    if (!v) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const s = String(v).trim();
    const d = new Date(s);
    return isNaN(d.getTime()) ? "" : d.toISOString().slice(0, 10);
  };

  const [bulkBusy, setBulkBusy] = useState(false);
  const [bulkResult, setBulkResult] = useState(null);
  const bulkInputRef = useRef();

  const handleBulkUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkBusy(true); setBulkResult(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const groups = new Map();
      let skipped = 0;
      rows.forEach((row) => {
        const date = normalizeDate(row["Date (YYYY-MM-DD)"] || row["Date"]);
        const patientName = String(row["Patient Name"] || "").trim();
        if (!date || !patientName) { skipped++; return; }
        const phone = String(row["Phone"] || "").trim();
        const doctorName = String(row["Doctor"] || "").trim();
        const shiftRaw = String(row["Shift (Morning/Evening)"] || row["Shift"] || "Morning").trim();
        const shift = ["Morning", "Evening"].includes(shiftRaw) ? shiftRaw : "Morning";
        const history = String(row["Brief Medical History"] || "").trim();
        const key = [date, patientName, phone, doctorName, shift, history].join("||");
        if (!groups.has(key)) {
          const amountDueRaw = row["Amount Due (₹, optional)"] ?? row["Amount Due"];
          const modeRaw = String(row["Mode of Payment (Cash/UPI/Card/Other, optional)"] || row["Mode of Payment"] || row["Mode"] || "").trim();
          groups.set(key, {
            date, patientName, phone, doctorName, shift, briefHistory: history, medicines: [],
            amountDue: amountDueRaw !== undefined && amountDueRaw !== "" ? Number(amountDueRaw) || 0 : undefined,
            mode: COLLECTION_MODES.includes(modeRaw) ? modeRaw : undefined,
          });
        }
        const medName = String(row["Medicine Name"] || "").trim();
        if (medName) {
          groups.get(key).medicines.push({ name: medName, qty: Number(row["Quantity"]) || 0, price: Number(row["Indicative Unit Price"]) || 0 });
        }
      });

      let created = 0, failed = 0;
      for (const g of groups.values()) {
        const doctor = doctors.find((d) => d.name.toLowerCase() === g.doctorName.toLowerCase());
        try {
          await addCase({ date: g.date, patientName: g.patientName, phone: g.phone, briefHistory: g.briefHistory, doctorId: doctor ? doctor.id : null, shift: g.shift, externalPrescription: "", imageUrl: null, medicines: g.medicines, amountDue: g.amountDue, mode: g.mode });
          created++;
        } catch { failed++; }
      }
      setBulkResult({ created, failed, skipped, unmatchedDoctors: [...groups.values()].some((g) => g.doctorName && !doctors.find((d) => d.name.toLowerCase() === g.doctorName.toLowerCase())) });
    } catch (err) {
      setBulkResult({ error: "Couldn't read that file — make sure it's the .xlsx template, not edited into a different format." });
    }
    setBulkBusy(false);
    e.target.value = "";
  };

  return (
    <div>
      {can("cases", "write") && (
        <div className="card">
          <h2>Bulk upload case records</h2>
          <p style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: -8 }}>Download the template, fill it in offline, then upload it back to create many case records at once.</p>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button className="btn secondary" type="button" onClick={downloadTemplate}>⬇ Download upload template</button>
            <input ref={bulkInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkUpload} style={{ display: "none" }} />
            <button className="btn" type="button" disabled={bulkBusy} onClick={() => bulkInputRef.current && bulkInputRef.current.click()}>{bulkBusy ? "Uploading…" : "📤 Upload filled template"}</button>
          </div>
          {bulkResult && (
            bulkResult.error ? <ErrorNote msg={bulkResult.error} /> : (
              <div className="info" style={{ marginTop: 10 }}>
                Created {bulkResult.created} case record(s).{bulkResult.failed > 0 ? ` ${bulkResult.failed} row(s) failed to save.` : ""}{bulkResult.skipped > 0 ? ` ${bulkResult.skipped} row(s) skipped (missing date or patient name).` : ""}
                {bulkResult.unmatchedDoctors ? " Some doctor names didn't match anyone on your roster exactly — those cases were created without a doctor assigned; edit them individually to fix." : ""}
              </div>
            )
          )}
          <div className="note-box">
            The template has one row per medicine. To give one patient's visit multiple medicines, repeat the same Date, Patient Name, Phone, Doctor, Shift, and Brief Medical History on consecutive rows, changing only the Medicine Name/Quantity/Price columns — matching rows are automatically grouped into a single case. The Doctor column must match a name already on your roster exactly (see "Doctor Shifts & Pay") or it will be left unassigned. Every case — from this upload or the form above — automatically gets a matching Collections entry: fill in Amount Due and Mode of Payment on this sheet to pre-fill it, or leave them blank and it's created as "Pending," ready to complete later on the Collections page.
          </div>
        </div>
      )}
      <div className="card">
        <h2>{editingId ? "Edit case paper entry" : "New case paper entry"}</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div style={{ position: "relative" }}>
            <label>Patient name</label>
            <input type="text" value={form.patientName} onChange={(e) => { setForm({ ...form, patientName: e.target.value }); setNameSuggestOpen(true); }} onFocus={() => setNameSuggestOpen(true)} />
            {nameSuggestOpen && nameSuggestions.length > 0 && (
              <div className="suggest-list">{nameSuggestions.map((p) => (<div key={p.id} className="suggest-item" onClick={() => applyPatient(p)}>{p.name}{p.mobile ? ` — ${p.mobile}` : ""}</div>))}</div>
            )}
          </div>
          <div style={{ position: "relative" }}>
            <label>Phone</label>
            <input type="tel" value={form.phone} onChange={(e) => { setForm({ ...form, phone: e.target.value }); setPhoneSuggestOpen(true); }} onFocus={() => setPhoneSuggestOpen(true)} />
            {phoneSuggestOpen && phoneSuggestions.length > 0 && (
              <div className="suggest-list">{phoneSuggestions.map((p) => (<div key={p.id} className="suggest-item" onClick={() => applyPatient(p)}>{p.mobile} — {p.name}</div>))}</div>
            )}
          </div>
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
        <div style={{ marginTop: 16, display: "flex", gap: 8 }}>
          {(editingId ? can("cases", "edit") : can("cases", "write")) && <button className="btn" type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : editingId ? "Update case record" : "Save case record"}</button>}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
        <ErrorNote msg={err} />
        <div className="note-box">Loose-medicine values are for clinical/inventory reference only — not posted as a separate expense, since "Medicine Bills" already covers what the clinic pays its supplier.</div>
      </div>
      <div className="card">
        <h2>Case register</h2>
        {cases.length === 0 ? <div className="empty">No case papers recorded yet.</div> : (
          <table>
            <thead><tr><CaseTh sortKeyName="caseNo">Case No.</CaseTh><CaseTh sortKeyName="date">Date</CaseTh><CaseTh sortKeyName="patientName">Patient</CaseTh><CaseTh sortKeyName="doctorName">Doctor</CaseTh><CaseTh sortKeyName="shift">Shift</CaseTh><CaseTh sortKeyName="briefHistory">History</CaseTh><CaseTh sortKeyName="medValue" className="num">Meds Value</CaseTh><th>Photo</th><th></th></tr></thead>
            <tbody>{sortedCases.map((c) => (
              <tr key={c.id}><td>{c.caseNo}</td><td>{c.date}</td><td>{c.patientName}</td><td>{c.doctorName || "—"}</td>
                <td><span className={"pill " + (c.shift || "").toLowerCase()}>{c.shift}</span></td>
                <td style={{ maxWidth: 180 }}>{c.briefHistory}</td><td className="num">{inr(medValue(c))}</td>
                <td>{c.image ? "📎" : "—"}</td><td style={{ display: "flex", gap: 6 }}>{can("cases", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(c)}>Edit</button>}{can("cases", "delete") && <button className="btn danger small" type="button" onClick={() => remove(c.id)}>Delete</button>}</td></tr>
            ))}</tbody>
          </table>
        )}
        <CustomExport
          rows={sortedCases} dateField="date" filenameBase="case-records" printTitle="Case Records" canExport={can("cases", "export")}
          buildSheets={(rows) => ({ Cases: rows.map((c) => ({ CaseNo: c.caseNo, Date: c.date, Patient: c.patientName, Phone: c.phone, Doctor: c.doctorName, Shift: c.shift, History: c.briefHistory, ExternalPrescription: c.externalPrescription, MedicinesDispensedValue: medValue(c) })) })}
          printColumns={[
            { label: "Case No.", value: (c) => c.caseNo }, { label: "Date", value: (c) => c.date }, { label: "Patient", value: (c) => c.patientName },
            { label: "Doctor", value: (c) => c.doctorName }, { label: "Shift", value: (c) => c.shift }, { label: "History", value: (c) => c.briefHistory },
          ]}
        />
      </div>
    </div>
  );
}

/** Exact age in years, months, and days — not just a rounded year count. */
function exactAge(dobIso) {
  if (!dobIso) return "—";
  const dob = new Date(dobIso + "T00:00:00");
  const today = new Date();
  if (dob > today) return "—";
  let years = today.getFullYear() - dob.getFullYear();
  let months = today.getMonth() - dob.getMonth();
  let days = today.getDate() - dob.getDate();
  if (days < 0) { months -= 1; days += new Date(today.getFullYear(), today.getMonth(), 0).getDate(); }
  if (months < 0) { years -= 1; months += 12; }
  return `${years}y ${months}m ${days}d`;
}

const GENDERS = ["Male", "Female", "Other"];

/* ============================== PATIENT MASTER ============================== */
function PatientMaster({ can, patients, addPatient, updatePatient, removePatient }) {
  const blank = { name: "", mobile: "", gender: "", dob: "", address: "" };
  const [form, setForm] = useState(blank);
  const [editingId, setEditingId] = useState(null);
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [query, setQuery] = useState("");

  const startEdit = (p) => { setEditingId(p.id); setForm({ name: p.name, mobile: p.mobile, gender: p.gender, dob: p.dob, address: p.address }); setErr(""); };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setErr(""); };
  const save = async () => {
    setErr(""); if (!form.name.trim()) { setErr("Enter the patient's name."); return; }
    setBusy(true);
    try {
      if (editingId) { await updatePatient(editingId, form); setEditingId(null); }
      else { await addPatient(form); }
      setForm(blank);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removePatient(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };

  const filtered = query.trim()
    ? patients.filter((p) => p.name.toLowerCase().includes(query.toLowerCase()) || p.mobile.includes(query))
    : patients;

  const columns = { name: (p) => p.name, mobile: (p) => p.mobile, gender: (p) => p.gender, dob: (p) => p.dob, address: (p) => p.address };
  const { sorted, Th } = useSortableRows(filtered, columns, "name");

  const doExcel = () => exportExcel("patient-master", { Patients: sorted.map((p) => ({ Name: p.name, Mobile: p.mobile, Gender: p.gender, DOB: p.dob, Age: exactAge(p.dob), Address: p.address })) });
  const doPrint = () => {
    const win = document.getElementById("print-root");
    if (!win) { window.print(); return; }
    const rowsHtml = sorted.map((p) => `<tr><td>${escapeHtml(p.name)}</td><td>${escapeHtml(p.mobile)}</td><td>${escapeHtml(p.gender)}</td><td>${p.dob}</td><td>${escapeHtml(exactAge(p.dob))}</td><td>${escapeHtml(p.address)}</td></tr>`).join("");
    win.innerHTML = `
      <h2>Patient Master</h2>
      <p style="color:#5B6B69;font-size:12px;">${sorted.length} patient(s)${query ? ` — filtered by "${escapeHtml(query)}"` : ""}</p>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Name</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Mobile</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Gender</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">DOB</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Age</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Address</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
    document.body.classList.add("printing-custom");
    window.print();
    setTimeout(() => { document.body.classList.remove("printing-custom"); win.innerHTML = ""; }, 300);
  };

  return (
    <div>
      <div className="card">
        <h2>{editingId ? "Edit patient" : "Add patient"}</h2>
        <div className="form-grid">
          <div><label>Patient name</label><input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} /></div>
          <div><label>Mobile number</label><input type="tel" value={form.mobile} onChange={(e) => setForm({ ...form, mobile: e.target.value })} /></div>
          <div><label>Gender</label><select value={form.gender} onChange={(e) => setForm({ ...form, gender: e.target.value })}><option value="">— Select —</option>{GENDERS.map((g) => <option key={g} value={g}>{g}</option>)}</select></div>
          <div><label>Date of birth</label><input type="date" value={form.dob} onChange={(e) => setForm({ ...form, dob: e.target.value })} /></div>
          <div style={{ gridColumn: "span 2" }}><label>Address</label><input type="text" style={{ width: "100%" }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></div>
        </div>
        {form.dob && <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 10 }}>Current age: <b>{exactAge(form.dob)}</b></div>}
        <div style={{ display: "flex", gap: 8 }}>
          {(editingId ? can("cases", "edit") : can("cases", "write")) && <button className="btn" type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : editingId ? "Update patient" : "Save patient"}</button>}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
        <ErrorNote msg={err} />
      </div>

      <div className="card">
        <h2>Patient register ({filtered.length})</h2>
        <input type="text" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search by name or mobile number" style={{ width: "100%", maxWidth: 360, marginBottom: 12 }} />
        {sorted.length === 0 ? <div className="empty">No patients found.</div> : (
          <table>
            <thead><tr><Th sortKeyName="name">Name</Th><Th sortKeyName="mobile">Mobile</Th><Th sortKeyName="gender">Gender</Th><Th sortKeyName="dob">DOB</Th><th>Age</th><Th sortKeyName="address">Address</Th><th></th></tr></thead>
            <tbody>{sorted.map((p) => (
              <tr key={p.id}>
                <td>{p.name}</td><td>{p.mobile}</td><td>{p.gender}</td><td>{p.dob}</td><td>{exactAge(p.dob)}</td><td style={{ maxWidth: 220 }}>{p.address}</td>
                <td style={{ display: "flex", gap: 6 }}>
                  {can("cases", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(p)}>Edit</button>}
                  {can("cases", "delete") && <button className="btn danger small" type="button" onClick={() => remove(p.id)}>Delete</button>}
                </td>
              </tr>
            ))}</tbody>
          </table>
        )}
        {can("cases", "export") && (
          <div className="export-row no-print">
            <button className="btn secondary small" type="button" onClick={doExcel}>⬇ Export Excel</button>
            <button className="btn secondary small" type="button" onClick={doPrint}>⎙ Export PDF</button>
          </div>
        )}
        <div className="note-box">Export respects your current search and column sort — sort by any header, filter with the search box, then export exactly what you see.</div>
      </div>
    </div>
  );
}

/* ============================== PATIENT HISTORY ============================== */
function PatientHistory({ can, updateCase, updateCollection, cases, doctors }) {
  const { call } = useApi();
  const [editCase, setEditCase] = useState(null);
  const [editColl, setEditColl] = useState(null);
  const [editErr, setEditErr] = useState("");
  const [query, setQuery] = useState("");
  const [suggestions, setSuggestions] = useState([]);
  const [selected, setSelected] = useState(null);
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState("");

  useEffect(() => {
    if (!query.trim() || selected) { setSuggestions([]); return; }
    const t = setTimeout(() => {
      call(`/patients/search?q=${encodeURIComponent(query)}`).then(setSuggestions).catch(() => {});
    }, 300);
    return () => clearTimeout(t);
  }, [query, call, selected]);

  const selectPatient = (p) => { setSelected(p); setQuery(`${p.patient_name}${p.phone ? " — " + p.phone : ""}`); setSuggestions([]); };

  const load = () => {
    if (!selected) return;
    setLoading(true); setErr("");
    const params = new URLSearchParams({ name: selected.patient_name, phone: selected.phone || "" });
    if (from) params.set("from", from);
    if (to) params.set("to", to);
    call(`/patients/history?${params.toString()}`).then(setData).catch((e) => setErr(e.message)).finally(() => setLoading(false));
  };
  useEffect(() => { if (selected) load(); }, [selected]); // eslint-disable-line

  const startEditCase = (c) => {
    setEditCase({ id: c.id, date: d10(c.case_date), patientName: c.patient_name, phone: c.phone || "", briefHistory: c.brief_history || "", doctorId: c.doctor_id || "", shift: c.shift || "Morning", externalPrescription: c.external_prescription || "", image: c.image_url || null, medicines: (c.medicines || []).map((m) => ({ name: m.medicine_name, qty: String(m.qty), price: String(m.unit_price) })) });
    setEditErr("");
  };
  const saveEditCase = async () => {
    setEditErr("");
    const medicines = editCase.medicines.filter((m) => m.name.trim()).map((m) => ({ name: m.name, qty: Number(m.qty) || 0, price: Number(m.price) || 0 }));
    try {
      await updateCase(editCase.id, { date: editCase.date, patientName: editCase.patientName, phone: editCase.phone, briefHistory: editCase.briefHistory, doctorId: editCase.doctorId || null, shift: editCase.shift, externalPrescription: editCase.externalPrescription, imageUrl: editCase.image, medicines });
      setEditCase(null); load();
    } catch (e) { setEditErr(e.message); }
  };

  const startEditColl = (c) => {
    setEditColl({ id: c.id, caseId: c.case_id || "", caseNo: c.case_no || "", patientName: c.patient_name, phone: c.phone || "", date: d10(c.collection_date), amountDue: String(c.amount_due), amountCollected: String(c.amount_collected), mode: c.mode || "Cash", image: c.image_url || null });
    setEditErr("");
  };
  const saveEditColl = async () => {
    setEditErr("");
    try {
      await updateCollection(editColl.id, { caseId: editColl.caseId || null, caseNo: editColl.caseNo || null, patientName: editColl.patientName, phone: editColl.phone, date: editColl.date, amountDue: Number(editColl.amountDue), amountCollected: Number(editColl.amountCollected) || 0, mode: editColl.mode, imageUrl: editColl.image });
      setEditColl(null); load();
    } catch (e) { setEditErr(e.message); }
  };

  const medValue = (c) => (c.medicines || []).reduce((s, m) => s + Number(m.qty) * Number(m.unit_price), 0);

  const doExcel = () => {
    if (!data) return;
    exportExcel(`patient-history-${data.patient.name}`, {
      Visits: data.cases.map((c) => ({ CaseNo: c.case_no, Date: d10(c.case_date), Doctor: c.doctor_name, Shift: c.shift, History: c.brief_history, ExternalPrescription: c.external_prescription, MedicinesValue: medValue(c) })),
      Payments: data.collections.map((c) => ({ Date: d10(c.collection_date), CaseNo: c.case_no, Due: Number(c.amount_due), Collected: Number(c.amount_collected), Balance: Number(c.balance), Mode: c.mode })),
    });
  };

  const doPrint = () => {
    if (!data) return;
    const win = document.getElementById("print-root");
    if (!win) { window.print(); return; }
    const visitRows = data.cases.map((c) => `<tr><td>${escapeHtml(c.case_no)}</td><td>${d10(c.case_date)}</td><td>${escapeHtml(c.doctor_name || "")}</td><td>${escapeHtml(c.brief_history || "")}</td></tr>`).join("");
    const payRows = data.collections.map((c) => `<tr><td>${d10(c.collection_date)}</td><td>${escapeHtml(c.case_no || "")}</td><td>${inr(c.amount_due)}</td><td>${inr(c.amount_collected)}</td><td>${inr(c.balance)}</td></tr>`).join("");
    win.innerHTML = `
      <h2>Patient History — ${escapeHtml(data.patient.name)}${data.patient.phone ? " (" + escapeHtml(data.patient.phone) + ")" : ""}</h2>
      <p style="color:#5B6B69;font-size:12px;">${escapeHtml(from || "start")} to ${escapeHtml(to || "today")}</p>
      <h3 style="margin-top:16px;">Visit History</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px;">Case No.</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px;">Date</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px;">Doctor</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px;">History</th></tr></thead>
        <tbody>${visitRows || "<tr><td colspan=4>No visits</td></tr>"}</tbody>
      </table>
      <h3 style="margin-top:16px;">Payment History</h3>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px;">Date</th><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px;">Case No.</th><th style="text-align:right;border-bottom:2px solid #C9A227;padding:5px;">Due</th><th style="text-align:right;border-bottom:2px solid #C9A227;padding:5px;">Collected</th><th style="text-align:right;border-bottom:2px solid #C9A227;padding:5px;">Balance</th></tr></thead>
        <tbody>${payRows || "<tr><td colspan=5>No payments</td></tr>"}</tbody>
      </table>`;
    document.body.classList.add("printing-custom");
    window.print();
    setTimeout(() => { document.body.classList.remove("printing-custom"); win.innerHTML = ""; }, 300);
  };

  return (
    <div>
      <div className="card">
        <h2>Find a patient</h2>
        <div style={{ position: "relative", maxWidth: 360 }}>
          <input type="text" value={query} onChange={(e) => { setQuery(e.target.value); setSelected(null); setData(null); }} placeholder="Search by name or mobile number" style={{ width: "100%" }} />
          {suggestions.length > 0 && (
            <div className="suggest-list">
              {suggestions.map((p, i) => (<div key={i} className="suggest-item" onClick={() => selectPatient(p)}>{p.patient_name}{p.phone ? ` — ${p.phone}` : ""}</div>))}
            </div>
          )}
        </div>
        {query.trim() && !selected && suggestions.length === 0 && <div className="empty">No matching patients yet — keep typing, or check spelling.</div>}
      </div>

      {editCase && (
        <div className="card" style={{ borderColor: "var(--accent)" }}>
          <h2>Edit visit — {editCase.patientName}</h2>
          <div className="form-grid">
            <div><label>Date</label><input type="date" value={editCase.date} onChange={(e) => setEditCase({ ...editCase, date: e.target.value })} /></div>
            <div><label>Patient name</label><input type="text" value={editCase.patientName} onChange={(e) => setEditCase({ ...editCase, patientName: e.target.value })} /></div>
            <div><label>Phone</label><input type="tel" value={editCase.phone} onChange={(e) => setEditCase({ ...editCase, phone: e.target.value })} /></div>
            <div><label>Doctor</label><select value={editCase.doctorId} onChange={(e) => setEditCase({ ...editCase, doctorId: e.target.value })}>{doctors.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
            <div><label>Shift</label><select value={editCase.shift} onChange={(e) => setEditCase({ ...editCase, shift: e.target.value })}>{SHIFTS.map((s) => <option key={s} value={s}>{s}</option>)}</select></div>
          </div>
          <div className="form-grid">
            <div style={{ gridColumn: "span 2" }}><label>Brief medical history</label><input type="text" style={{ width: "100%" }} value={editCase.briefHistory} onChange={(e) => setEditCase({ ...editCase, briefHistory: e.target.value })} /></div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="btn" type="button" onClick={saveEditCase}>Update visit</button>
            <button className="btn secondary" type="button" onClick={() => setEditCase(null)}>Cancel</button>
          </div>
          <ErrorNote msg={editErr} />
        </div>
      )}
      {editColl && (
        <div className="card" style={{ borderColor: "var(--accent)" }}>
          <h2>Edit payment — {editColl.patientName}</h2>
          <div className="form-grid">
            <div><label>Date</label><input type="date" value={editColl.date} onChange={(e) => setEditColl({ ...editColl, date: e.target.value })} /></div>
            <div><label>Amount due (₹)</label><input type="number" value={editColl.amountDue} onChange={(e) => setEditColl({ ...editColl, amountDue: e.target.value })} /></div>
            <div><label>Amount collected (₹)</label><input type="number" value={editColl.amountCollected} onChange={(e) => setEditColl({ ...editColl, amountCollected: e.target.value })} /></div>
            <div><label>Mode</label><select value={editColl.mode} onChange={(e) => setEditColl({ ...editColl, mode: e.target.value })}>{COLLECTION_MODES.map((m) => <option key={m} value={m}>{m}</option>)}</select></div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button className="btn" type="button" onClick={saveEditColl}>Update payment</button>
            <button className="btn secondary" type="button" onClick={() => setEditColl(null)}>Cancel</button>
          </div>
          <ErrorNote msg={editErr} />
        </div>
      )}

      {selected && (
        <>
          <div className="card">
            <h2>{selected.patient_name}{selected.phone ? ` — ${selected.phone}` : ""}</h2>
            <div className="form-grid" style={{ maxWidth: 340 }}>
              <div><label>From</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></div>
              <div><label>To</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></div>
            </div>
            <button className="btn" type="button" onClick={load}>Apply range</button>
            <ErrorNote msg={err} />
          </div>

          {loading ? <div className="empty">Loading history…</div> : data && (
            <>
              <div className="card">
                <h2>Visit history ({data.cases.length})</h2>
                {data.cases.length === 0 ? <div className="empty">No visits in this range.</div> : (
                  <table>
                    <thead><tr><th>Case No.</th><th>Date</th><th>Doctor</th><th>Shift</th><th>History</th><th className="num">Meds Value</th><th></th></tr></thead>
                    <tbody>{data.cases.map((c) => (
                      <tr key={c.id}><td>{c.case_no}</td><td>{d10(c.case_date)}</td><td>{c.doctor_name || "—"}</td>
                        <td><span className={"pill " + (c.shift || "").toLowerCase()}>{c.shift}</span></td>
                        <td style={{ maxWidth: 200 }}>{c.brief_history}</td><td className="num">{inr(medValue(c))}</td><td>{can("cases", "edit") && <button className="btn secondary small" type="button" onClick={() => startEditCase(c)}>Edit</button>}</td></tr>
                    ))}</tbody>
                  </table>
                )}
              </div>
              <div className="card">
                <h2>Payment history ({data.collections.length})</h2>
                {data.collections.length === 0 ? <div className="empty">No payments in this range.</div> : (
                  <table>
                    <thead><tr><th>Date</th><th>Case No.</th><th className="num">Due</th><th className="num">Collected</th><th className="num">Balance</th><th>Mode</th><th></th></tr></thead>
                    <tbody>{data.collections.map((c) => (
                      <tr key={c.id}><td>{d10(c.collection_date)}</td><td>{c.case_no || "—"}</td>
                        <td className="num">{inr(c.amount_due)}</td><td className="num">{inr(c.amount_collected)}</td>
                        <td className="num" style={{ color: c.balance > 0 ? "var(--expense)" : "var(--income)" }}>{inr(c.balance)}</td><td>{c.mode}</td><td>{can("collections", "edit") && <button className="btn secondary small" type="button" onClick={() => startEditColl(c)}>Edit</button>}</td></tr>
                    ))}</tbody>
                  </table>
                )}
                {can("cases", "export") && (
                  <div className="export-row no-print">
                    <button className="btn secondary small" type="button" onClick={doExcel}>⬇ Export Excel</button>
                    <button className="btn secondary small" type="button" onClick={doPrint}>⎙ Export PDF</button>
                  </div>
                )}
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}

/* ============================== COLLECTIONS ============================== */
function Collections({ collections, addCollection, updateCollection, removeCollection, cases, fy, can }) {
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

  const [editingId, setEditingId] = useState(null);
  const [caseQuery, setCaseQuery] = useState("");
  const [caseSuggestOpen, setCaseSuggestOpen] = useState(false);
  const caseSuggestions = caseQuery.trim()
    ? cases.filter((c) => c.caseNo.toLowerCase().includes(caseQuery.toLowerCase()) || c.patientName.toLowerCase().includes(caseQuery.toLowerCase())).slice(0, 15)
    : [];
  const pickCase = (c) => {
    setForm({ ...form, caseId: c.id, caseNo: c.caseNo, patientName: c.patientName, phone: c.phone });
    setCaseQuery(`${c.caseNo} — ${c.patientName}`);
    setCaseSuggestOpen(false);
  };
  const clearCase = () => { setForm({ ...form, caseId: "", caseNo: "" }); setCaseQuery(""); };
  const startEdit = (c) => {
    setEditingId(c.id);
    setForm({ caseId: c.caseId || "", caseNo: c.caseNo || "", patientName: c.patientName, phone: c.phone || "", date: c.date, amountDue: String(c.amountDue), amountCollected: String(c.amountCollected), mode: c.mode || "Cash", image: c.image || null });
    setCaseQuery(c.caseNo ? `${c.caseNo} — ${c.patientName}` : "");
    setErr("");
  };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setCaseQuery(""); setErr(""); };
  const save = async () => {
    setErr(""); if (!form.patientName.trim() || form.amountDue === "") { setErr("Enter patient name and amount due."); return; }
    setBusy(true);
    const payload = { caseId: form.caseId || null, caseNo: form.caseNo || null, patientName: form.patientName, phone: form.phone, date: form.date, amountDue: Number(form.amountDue), amountCollected: Number(form.amountCollected) || 0, mode: form.mode, imageUrl: form.image };
    try {
      if (editingId) { await updateCollection(editingId, payload); setEditingId(null); }
      else { await addCollection(payload); }
      setForm(blank); setCaseQuery("");
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeCollection(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };

  // ---- Cases with no Collections entry yet — find them and let staff book them, one at a time or all at once ----
  const linkedCaseIds = useMemo(() => new Set(collections.map((c) => c.caseId).filter(Boolean)), [collections]);
  const orphanCases = useMemo(() => [...cases].filter((c) => !linkedCaseIds.has(c.id)).sort((a, b) => (a.date < b.date ? 1 : -1)), [cases, linkedCaseIds]);
  const [orphanRows, setOrphanRows] = useState({}); // caseId -> { amountDue, amountCollected, mode }
  const [selected, setSelected] = useState(new Set());
  const [bookBusy, setBookBusy] = useState(false);
  const [bookErr, setBookErr] = useState("");

  const orphanRow = (c) => orphanRows[c.id] || { amountDue: "", amountCollected: "0", mode: "Cash" };
  const setOrphanField = (id, field, value) => setOrphanRows((p) => ({ ...p, [id]: { ...(p[id] || { amountDue: "", amountCollected: "0", mode: "Cash" }), [field]: value } }));
  const toggleSelected = (id) => setSelected((p) => { const n = new Set(p); n.has(id) ? n.delete(id) : n.add(id); return n; });
  const toggleSelectAll = () => setSelected((p) => (p.size === orphanCases.length ? new Set() : new Set(orphanCases.map((c) => c.id))));

  const bookOne = async (c) => {
    const row = orphanRow(c);
    if (row.amountDue === "" || row.amountDue === undefined) throw new Error(`Enter Amount Due for ${c.caseNo} before booking.`);
    await addCollection({ caseId: c.id, caseNo: c.caseNo, patientName: c.patientName, phone: c.phone, date: c.date, amountDue: Number(row.amountDue), amountCollected: Number(row.amountCollected) || 0, mode: row.mode || "Cash", imageUrl: null });
  };
  const bookSelected = async () => {
    setBookErr(""); setBookBusy(true);
    const targets = orphanCases.filter((c) => selected.has(c.id));
    try {
      for (const c of targets) await bookOne(c);
      setSelected(new Set());
    } catch (e) { setBookErr(e.message); }
    setBookBusy(false);
  };
  const bookSingle = async (c) => {
    setBookErr("");
    try { await bookOne(c); } catch (e) { setBookErr(e.message); }
  };

  return (
    <div>
      <div className="card">
        <h2>{editingId ? "Edit collection entry" : "New collection entry"}</h2>
        <div className="form-grid">
          <div style={{ position: "relative" }}>
            <label>Linked case no.</label>
            <input
              type="text" value={caseQuery} placeholder="Search by case no. or patient name"
              onChange={(e) => { setCaseQuery(e.target.value); setCaseSuggestOpen(true); if (!e.target.value.trim()) clearCase(); }}
              onFocus={() => setCaseSuggestOpen(true)}
              style={{ width: "100%" }}
            />
            {caseSuggestOpen && caseSuggestions.length > 0 && (
              <div className="suggest-list">
                {caseSuggestions.map((c) => (<div key={c.id} className="suggest-item" onClick={() => pickCase(c)}>{c.caseNo} — {c.patientName}</div>))}
              </div>
            )}
            {form.caseId && <div style={{ fontSize: 11, color: "var(--income)", marginTop: 4 }}>✓ Linked to {form.caseNo}</div>}
          </div>
          <div><label>Patient name</label><input type="text" value={form.patientName} onChange={(e) => setForm({ ...form, patientName: e.target.value })} /></div>
          <div><label>Phone</label><input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></div>
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Amount due (₹)</label><input type="number" value={form.amountDue} onChange={(e) => setForm({ ...form, amountDue: e.target.value })} /></div>
          <div><label>Amount collected (₹)</label><input type="number" value={form.amountCollected} onChange={(e) => setForm({ ...form, amountCollected: e.target.value })} /></div>
          <div><label>Mode</label><select value={form.mode} onChange={(e) => setForm({ ...form, mode: e.target.value })}>{COLLECTION_MODES.map((m) => <option key={m} value={m}>{m}</option>)}</select></div>
        </div>
        <ImageCapture value={form.image} onChange={(img) => setForm({ ...form, image: img })} />
        <div style={{ marginTop: 14, display: "flex", gap: 8 }}>
          {(editingId ? can("collections", "edit") : can("collections", "write")) && (
            <button className="btn" type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : editingId ? "Update collection entry" : "Save collection entry"}</button>
          )}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
        <ErrorNote msg={err} />
      </div>

      {can("collections", "write") && orphanCases.length > 0 && (
        <div className="card" style={{ borderColor: "var(--accent)" }}>
          <h2>Cases without a Collection entry ({orphanCases.length})</h2>
          <p style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: -8 }}>Every new case now gets a Collections entry automatically — this list is for older cases from before that, or any case that somehow ended up without one. Amount Due starts blank on purpose; Mode defaults to Cash and Amount Collected to ₹0, both editable before booking.</p>
          <table>
            <thead><tr>
              <th><input type="checkbox" checked={selected.size === orphanCases.length && orphanCases.length > 0} onChange={toggleSelectAll} /> All</th>
              <th>Date</th><th>Case No.</th><th>Patient</th><th className="num">Amount Due (₹)</th><th className="num">Amount Collected (₹)</th><th>Mode</th><th></th>
            </tr></thead>
            <tbody>{orphanCases.map((c) => {
              const row = orphanRow(c);
              return (
                <tr key={c.id}>
                  <td><input type="checkbox" checked={selected.has(c.id)} onChange={() => toggleSelected(c.id)} /></td>
                  <td>{c.date}</td><td>{c.caseNo}</td><td>{c.patientName}</td>
                  <td className="num"><input type="number" style={{ width: 90 }} value={row.amountDue} placeholder="Required" onChange={(e) => setOrphanField(c.id, "amountDue", e.target.value)} /></td>
                  <td className="num"><input type="number" style={{ width: 90 }} value={row.amountCollected} onChange={(e) => setOrphanField(c.id, "amountCollected", e.target.value)} /></td>
                  <td><select value={row.mode} onChange={(e) => setOrphanField(c.id, "mode", e.target.value)}>{COLLECTION_MODES.map((m) => <option key={m} value={m}>{m}</option>)}</select></td>
                  <td><button className="btn small" type="button" onClick={() => bookSingle(c)}>Book</button></td>
                </tr>
              );
            })}</tbody>
          </table>
          <button className="btn" style={{ marginTop: 12 }} type="button" disabled={bookBusy || selected.size === 0} onClick={bookSelected}>{bookBusy ? "Booking…" : `Book ${selected.size} selected`}</button>
          <ErrorNote msg={bookErr} />
        </div>
      )}

      <div className="card">
        <h2>Collection register</h2>
        {collections.length === 0 ? <div className="empty">No collections recorded yet.</div> : (
          <table>
            <thead><tr><th>Date</th><th>Case No.</th><th>Patient</th><th>Phone</th><th className="num">Due</th><th className="num">Collected</th><th className="num">Balance</th><th>Mode</th><th></th></tr></thead>
            <tbody>{[...collections].sort((a, b) => (a.date < b.date ? 1 : -1)).map((c) => {
              const pending = !c.mode && Number(c.amountDue || 0) === 0 && Number(c.amountCollected || 0) === 0;
              return (
              <tr key={c.id} style={pending ? { background: "#FFF8E8" } : undefined}><td>{c.date}</td><td>{c.caseNo || "—"}</td><td>{c.patientName}</td><td>{c.phone}</td>
                <td className="num">{inr(c.amountDue)}</td><td className="num">{inr(c.amountCollected)}</td>
                <td className="num" style={{ color: c.balance > 0 ? "var(--expense)" : "var(--income)" }}>{inr(c.balance)}</td>
                <td>{c.mode || <span style={{ color: "var(--accent)", fontWeight: 700 }}>Pending</span>}</td><td style={{ display: "flex", gap: 6 }}>{can("collections", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(c)}>Edit</button>}{can("collections", "delete") && <button className="btn danger small" type="button" onClick={() => remove(c.id)}>Delete</button>}</td></tr>
            );})}</tbody>
          </table>
        )}
        <CustomExport
          rows={collections} dateField="date" filenameBase="collections" printTitle="Collections" canExport={can("collections", "export")}
          buildSheets={(rows) => ({ Collections: rows, Daily: rollups.daily, Weekly: rollups.weekly, Monthly: rollups.monthly })}
          printColumns={[
            { label: "Date", value: (r) => r.date }, { label: "Case No.", value: (r) => r.caseNo }, { label: "Patient", value: (r) => r.patientName },
            { label: "Due", value: (r) => inr(r.amountDue) }, { label: "Collected", value: (r) => inr(r.amountCollected) }, { label: "Balance", value: (r) => inr(r.balance) },
          ]}
        />
        {can("collections", "export") && <div className="note-box">Leave both dates blank in "Custom range export" and it exports every collection on record — a full downloadable database, not just a date range.</div>}
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
function DoctorShifts({ doctors, addDoctor, updateDoctor, removeDoctor, doctorPays, addDoctorPay, updateDoctorPay, removeDoctorPay, can }) {
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

  const [editingDocId, setEditingDocId] = useState(null);
  const [editingPayId, setEditingPayId] = useState(null);
  const startEditDoc = (x) => { setEditingDocId(x.id); setD({ name: x.name, shift: x.shift, payType: x.payType, rate: String(x.rate) }); setErr(""); };
  const cancelEditDoc = () => { setEditingDocId(null); setD({ name: "", shift: "Morning", payType: "Daily", rate: "" }); setErr(""); };
  const addDoc = async () => {
    setErr(""); if (!d.name.trim()) return; setBusy(true);
    const payload = { name: d.name, shift: d.shift, payType: d.payType, rate: Number(d.rate) || 0 };
    try {
      if (editingDocId) { await updateDoctor(editingDocId, payload); setEditingDocId(null); }
      else { await addDoctor(payload); }
      setD({ name: "", shift: "Morning", payType: "Daily", rate: "" });
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const delDoc = async (id) => { try { await removeDoctor(id); if (editingDocId === id) cancelEditDoc(); } catch (e) { setErr(e.message); } };
  const startEditPay = (x) => { setEditingPayId(x.id); setP({ date: x.date, doctorId: x.doctorId, amount: String(x.amount) }); setErr(""); };
  const cancelEditPay = () => { setEditingPayId(null); setP({ date: todayISO(), doctorId: "", amount: "" }); setErr(""); };
  const addPay = async () => {
    setErr(""); if (!p.doctorId || !p.amount) return; setBusy(true);
    const payload = { doctorId: p.doctorId, date: p.date, amount: Number(p.amount) };
    try {
      if (editingPayId) { await updateDoctorPay(editingPayId, payload); setEditingPayId(null); }
      else { await addDoctorPay(payload); }
      setP({ date: todayISO(), doctorId: "", amount: "" });
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const delPay = async (id) => { try { await removeDoctorPay(id); if (editingPayId === id) cancelEditPay(); } catch (e) { setErr(e.message); } };

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
        <div style={{ display: "flex", gap: 8 }}>
          {(editingDocId ? can("doctorPay", "edit") : can("doctorPay", "write")) && <button className="btn" type="button" disabled={busy} onClick={addDoc}>{editingDocId ? "Update doctor" : "Add doctor"}</button>}
          {editingDocId && <button className="btn secondary" type="button" onClick={cancelEditDoc}>Cancel edit</button>}
        </div>
        <table style={{ marginTop: 14 }}><thead><tr><th>Name</th><th>Shift</th><th>Pay type</th><th className="num">Rate</th><th></th></tr></thead>
          <tbody>{doctors.map((x) => (<tr key={x.id}><td>{x.name}</td><td><span className={"pill " + x.shift.toLowerCase()}>{x.shift}</span></td><td>{x.payType}</td><td className="num">{inr(x.rate)}</td><td style={{ display: "flex", gap: 6 }}>{can("doctorPay", "edit") && <button className="btn secondary small" type="button" onClick={() => startEditDoc(x)}>Edit</button>}{can("doctorPay", "delete") && <button className="btn danger small" type="button" onClick={() => delDoc(x.id)}>Delete</button>}</td></tr>))}</tbody>
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
        <div style={{ display: "flex", gap: 8 }}>
          {(editingPayId ? can("doctorPay", "edit") : can("doctorPay", "write")) && <button className="btn" type="button" disabled={busy} onClick={addPay}>{editingPayId ? "Update pay entry" : "Log pay entry"}</button>}
          {editingPayId && <button className="btn secondary" type="button" onClick={cancelEditPay}>Cancel edit</button>}
        </div>
        <table style={{ marginTop: 14 }}><thead><tr><th>Date</th><th>Doctor</th><th className="num">Amount</th><th></th></tr></thead>
          <tbody>{[...doctorPays].sort((a, b) => (a.date < b.date ? 1 : -1)).map((x) => (<tr key={x.id}><td>{x.date}</td><td>{x.doctorName}</td><td className="num">{inr(x.amount)}</td><td style={{ display: "flex", gap: 6 }}>{can("doctorPay", "edit") && <button className="btn secondary small" type="button" onClick={() => startEditPay(x)}>Edit</button>}{can("doctorPay", "delete") && <button className="btn danger small" type="button" onClick={() => delPay(x.id)}>Delete</button>}</td></tr>))}</tbody>
        </table>
        <CustomExport
          rows={doctorPays} dateField="date" filenameBase="doctor-pay" printTitle="Doctor Shifts & Pay" canExport={can("doctorPay", "export")}
          buildSheets={(rows) => ({ DoctorPay: rows.map((x) => ({ Date: x.date, Doctor: x.doctorName, Amount: x.amount })) })}
          printColumns={[{ label: "Date", value: (x) => x.date }, { label: "Doctor", value: (x) => x.doctorName }, { label: "Amount", value: (x) => inr(x.amount) }]}
        />
      </div>
      <div className="card">
        <h2>Daily net result — last 14 days</h2>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={last14}><CartesianGrid strokeDasharray="3 3" stroke="#E1E8E6" /><XAxis dataKey="date" tickFormatter={(v) => v.slice(5)} fontSize={11} /><YAxis fontSize={11} /><Tooltip formatter={(v) => inr(v)} /><Bar dataKey="net" fill="#714B67" radius={[4, 4, 0, 0]} /></BarChart>
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
function Referrals({ referrals, addReferral, updateReferral, removeReferral, fy, can }) {
  const blank = { date: todayISO(), patientName: "", referralType: "Lab Test", referredTo: "", amount: "", notes: "" };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const startEdit = (r) => { setEditingId(r.id); setForm({ date: r.date, patientName: r.patientName, referralType: r.referralType, referredTo: r.referredTo || "", amount: String(r.amount), notes: r.notes || "" }); setErr(""); };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setErr(""); };
  const save = async () => {
    setErr(""); if (!form.patientName.trim() || !form.amount) { setErr("Enter patient name and amount."); return; }
    setBusy(true);
    const payload = { ...form, amount: Number(form.amount) };
    try {
      if (editingId) { await updateReferral(editingId, payload); setEditingId(null); }
      else { await addReferral(payload); }
      setForm(blank);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeReferral(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };
  const range = fyRange(fy);
  const total = referrals.filter((r) => r.date >= range.start && r.date <= range.end).reduce((a, r) => a + r.amount, 0);
  return (
    <div>
      <div className="card">
        <h2>{editingId ? "Edit referral" : "New referral"}</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Patient</label><input type="text" value={form.patientName} onChange={(e) => setForm({ ...form, patientName: e.target.value })} /></div>
          <div><label>Referral type</label><select value={form.referralType} onChange={(e) => setForm({ ...form, referralType: e.target.value })}>{REFERRAL_TYPES.map((r) => <option key={r} value={r}>{r}</option>)}</select></div>
          <div><label>Referred to</label><input type="text" value={form.referredTo} onChange={(e) => setForm({ ...form, referredTo: e.target.value })} /></div>
          <div><label>Amount received (₹)</label><input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></div>
          <div><label>Notes</label><input type="text" value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {(editingId ? can("referrals", "edit") : can("referrals", "write")) && <button className="btn" type="button" disabled={busy} onClick={save}>{editingId ? "Update referral" : "Save referral"}</button>}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
        <ErrorNote msg={err} />
      </div>
      <div className="card">
        <h2>Referral income — FY {fy} total: {inr(total)}</h2>
        {referrals.length === 0 ? <div className="empty">No referrals recorded yet.</div> : (
          <table><thead><tr><th>Date</th><th>Patient</th><th>Type</th><th>Referred To</th><th className="num">Amount</th><th>Notes</th><th></th></tr></thead>
            <tbody>{[...referrals].sort((a, b) => (a.date < b.date ? 1 : -1)).map((r) => (<tr key={r.id}><td>{r.date}</td><td>{r.patientName}</td><td>{r.referralType}</td><td>{r.referredTo}</td><td className="num">{inr(r.amount)}</td><td>{r.notes}</td><td style={{ display: "flex", gap: 6 }}>{can("referrals", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(r)}>Edit</button>}{can("referrals", "delete") && <button className="btn danger small" type="button" onClick={() => remove(r.id)}>Delete</button>}</td></tr>))}</tbody>
          </table>
        )}
        <CustomExport
          rows={referrals} dateField="date" filenameBase="referral-income" printTitle="Referral Income" canExport={can("referrals", "export")}
          buildSheets={(rows) => ({ Referrals: rows })}
          printColumns={[{ label: "Date", value: (r) => r.date }, { label: "Patient", value: (r) => r.patientName }, { label: "Type", value: (r) => r.referralType }, { label: "Referred To", value: (r) => r.referredTo }, { label: "Amount", value: (r) => inr(r.amount) }]}
        />
      </div>
    </div>
  );
}

/* ============================== GIFTS REGISTER ============================== */
function Gifts({ gifts, addGift, updateGift, removeGift, doctors, can }) {
  const blank = { date: todayISO(), repName: "", company: "", gift: "", doctorId: "", amount: "" };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const startEdit = (g) => { setEditingId(g.id); setForm({ date: g.date, repName: g.repName, company: g.company || "", gift: g.gift || "", doctorId: g.doctorId || "", amount: g.amount ? String(g.amount) : "" }); setErr(""); };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setErr(""); };
  const save = async () => {
    setErr(""); if (!form.repName.trim()) { setErr("Enter the medical rep's name."); return; }
    setBusy(true);
    const payload = { ...form, amount: Number(form.amount) || 0 };
    try {
      if (editingId) { await updateGift(editingId, payload); setEditingId(null); }
      else { await addGift(payload); }
      setForm(blank);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeGift(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };
  return (
    <div>
      <div className="card">
        <h2>{editingId ? "Edit gift entry" : "Log a gift received"}</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Medical rep</label><input type="text" value={form.repName} onChange={(e) => setForm({ ...form, repName: e.target.value })} /></div>
          <div><label>Company</label><input type="text" value={form.company} onChange={(e) => setForm({ ...form, company: e.target.value })} /></div>
          <div><label>Gift description</label><input type="text" value={form.gift} onChange={(e) => setForm({ ...form, gift: e.target.value })} /></div>
          <div><label>Doctor</label><select value={form.doctorId} onChange={(e) => setForm({ ...form, doctorId: e.target.value })}><option value="">— Select —</option>{doctors.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></div>
          <div><label>Amount (₹, optional)</label><input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} placeholder="Leave blank for non-monetary gifts" /></div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {(editingId ? can("gifts", "edit") : can("gifts", "write")) && <button className="btn" type="button" disabled={busy} onClick={save}>{editingId ? "Update entry" : "Save entry"}</button>}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
        <ErrorNote msg={err} />
        <div className="note-box">The rep, company, and gift description are always disclosure-only. If you enter an Amount, that value now flows into the Income Statement as "Gift Income" and into the Balance Sheet/Capital Account — leave it blank for non-monetary gifts (samples, stationery, etc.) that shouldn't affect the books. Gifting from pharmaceutical reps to doctors is restricted under India's medical ethics rules (e.g. the Uniform Code for Pharmaceuticals Marketing Practices and NMC's professional-conduct regulations); check current guidance with your professional body if in doubt.</div>
      </div>
      <div className="card">
        <h2>Gifts register</h2>
        {gifts.length === 0 ? <div className="empty">No entries yet.</div> : (
          <table><thead><tr><th>Date</th><th>Rep</th><th>Company</th><th>Gift</th><th>Doctor</th><th className="num">Amount</th><th></th></tr></thead>
            <tbody>{[...gifts].sort((a, b) => (a.date < b.date ? 1 : -1)).map((g) => (<tr key={g.id}><td>{g.date}</td><td>{g.repName}</td><td>{g.company}</td><td>{g.gift}</td><td>{g.doctorName || "—"}</td><td className="num">{g.amount ? inr(g.amount) : "—"}</td><td style={{ display: "flex", gap: 6 }}>{can("gifts", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(g)}>Edit</button>}{can("gifts", "delete") && <button className="btn danger small" type="button" onClick={() => remove(g.id)}>Delete</button>}</td></tr>))}</tbody>
          </table>
        )}
        <CustomExport
          rows={gifts} dateField="date" filenameBase="gifts-register" printTitle="Gifts Register" canExport={can("gifts", "export")}
          buildSheets={(rows) => ({ Gifts: rows.map((g) => ({ Date: g.date, Rep: g.repName, Company: g.company, Gift: g.gift, Doctor: g.doctorName, Amount: g.amount })) })}
          printColumns={[{ label: "Date", value: (g) => g.date }, { label: "Rep", value: (g) => g.repName }, { label: "Company", value: (g) => g.company }, { label: "Gift", value: (g) => g.gift }, { label: "Doctor", value: (g) => g.doctorName }, { label: "Amount", value: (g) => (g.amount ? inr(g.amount) : "—") }]}
        />
      </div>
    </div>
  );
}

/* ============================== EXPENSES ============================== */
function Expenses({ expenses, addExpense, updateExpense, removeExpense, fy, can }) {
  const blank = { date: todayISO(), category: EXPENSE_CATEGORIES[0], amount: "", narration: "", image: null };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const startEdit = (e) => { setEditingId(e.id); setForm({ date: e.date, category: e.category, amount: String(e.amount), narration: e.narration || "", image: e.image || null }); setErr(""); };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setErr(""); };
  const save = async () => {
    setErr(""); if (!form.amount) { setErr("Enter an amount."); return; }
    setBusy(true);
    const payload = { date: form.date, category: form.category, amount: Number(form.amount), narration: form.narration, imageUrl: form.image };
    try {
      if (editingId) { await updateExpense(editingId, payload); setEditingId(null); }
      else { await addExpense(payload); }
      setForm(blank);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeExpense(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };
  const range = fyRange(fy);
  const inFY = expenses.filter((e) => e.date >= range.start && e.date <= range.end);
  const byCat = {}; inFY.forEach((e) => { byCat[e.category] = (byCat[e.category] || 0) + e.amount; });
  const total = inFY.reduce((a, e) => a + e.amount, 0);
  return (
    <div>
      <div className="card">
        <h2>{editingId ? "Edit expense" : "New expense"}</h2>
        <div className="form-grid">
          <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
          <div><label>Category</label><select value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>{EXPENSE_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}</select></div>
          <div><label>Amount (₹)</label><input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></div>
          <div style={{ gridColumn: "span 2" }}><label>Narration</label><input type="text" style={{ width: "100%" }} value={form.narration} onChange={(e) => setForm({ ...form, narration: e.target.value })} /></div>
        </div>
        <ImageCapture value={form.image} onChange={(img) => setForm({ ...form, image: img })} />
        <div style={{ marginTop: 14, display: "flex", gap: 8 }}>
          {(editingId ? can("expenses", "edit") : can("expenses", "write")) && <button className="btn" type="button" disabled={busy} onClick={save}>{busy ? "Saving…" : editingId ? "Update expense" : "Save expense"}</button>}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
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
            <tbody>{[...expenses].sort((a, b) => (a.date < b.date ? 1 : -1)).map((e) => (<tr key={e.id}><td>{e.date}</td><td>{e.category}</td><td>{e.narration}</td><td className="num">{inr(e.amount)}</td><td>{e.image ? "📎" : "—"}</td><td style={{ display: "flex", gap: 6 }}>{can("expenses", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(e)}>Edit</button>}{can("expenses", "delete") && <button className="btn danger small" type="button" onClick={() => remove(e.id)}>Delete</button>}</td></tr>))}</tbody>
          </table>
        )}
        <CustomExport
          rows={expenses} dateField="date" filenameBase="expenses" printTitle="Expenses" canExport={can("expenses", "export")}
          buildSheets={(rows) => ({ Expenses: rows, CategoryTotals: EXPENSE_CATEGORIES.map((c) => ({ Category: c, Amount: rows.filter((r) => r.category === c).reduce((s, r) => s + r.amount, 0) })) })}
          printColumns={[{ label: "Date", value: (e) => e.date }, { label: "Category", value: (e) => e.category }, { label: "Narration", value: (e) => e.narration }, { label: "Amount", value: (e) => inr(e.amount) }]}
        />
      </div>
    </div>
  );
}

/* ============================== FIXED ASSETS ============================== */
function FixedAssets({ assets, addAsset, updateAsset, removeAsset, fy, can }) {
  const { call } = useApi();
  const blank = { name: "", block: BLOCKS[0].name, rate: BLOCKS[0].rate, purchaseDate: todayISO(), cost: "" };
  const [form, setForm] = useState(blank); const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const [dep, setDep] = useState({ rows: [], totalDep: 0 });

  useEffect(() => { call(`/assets/depreciation?fy=${fy}`).then((d) => setDep({ rows: (d.rows || []).map((r) => ({ ...mapAsset(r), ...r })), totalDep: Number(d.totalDep) })).catch(() => {}); }, [call, fy, assets]);

  const [editingId, setEditingId] = useState(null);
  const startEdit = (a) => { setEditingId(a.id); setForm({ name: a.name, block: a.block, rate: a.rate, purchaseDate: a.purchaseDate, cost: String(a.cost) }); setErr(""); };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setErr(""); };
  const add = async () => {
    setErr(""); if (!form.name.trim() || !form.cost) { setErr("Enter a description and cost."); return; }
    setBusy(true);
    const payload = { ...form, cost: Number(form.cost) };
    try {
      if (editingId) { await updateAsset(editingId, payload); setEditingId(null); }
      else { await addAsset(payload); }
      setForm(blank);
    } catch (e) { setErr(e.message); }
    setBusy(false);
  };
  const remove = async (id) => { try { await removeAsset(id); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };

  return (
    <div>
      <div className="card">
        <h2>{editingId ? "Edit asset" : "Add asset (CAPEX)"}</h2>
        <div className="form-grid">
          <div><label>Description</label><input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} placeholder="e.g. ECG machine" /></div>
          <div><label>Block of assets</label><select value={form.block} onChange={(e) => { const b = BLOCKS.find((x) => x.name === e.target.value); setForm({ ...form, block: b.name, rate: b.rate }); }}>{BLOCKS.map((b) => <option key={b.name} value={b.name}>{b.name} ({b.rate}%)</option>)}</select></div>
          <div><label>Rate %</label><input type="number" value={form.rate} onChange={(e) => setForm({ ...form, rate: Number(e.target.value) })} /></div>
          <div><label>Date put to use</label><input type="date" value={form.purchaseDate} onChange={(e) => setForm({ ...form, purchaseDate: e.target.value })} /></div>
          <div><label>Cost (₹)</label><input type="number" value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} /></div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {(editingId ? can("assets", "edit") : can("assets", "write")) && <button className="btn" type="button" disabled={busy} onClick={add}>{editingId ? "Update asset" : "Add to register"}</button>}
          {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
        </div>
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
                <td style={{ display: "flex", gap: 6 }}>{can("assets", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(r)}>Edit</button>}{can("assets", "delete") && <button className="btn danger small" type="button" onClick={() => remove(r.id)}>Delete</button>}</td></tr>
            ))}</tbody>
            <tfoot><tr style={{ fontWeight: 700 }}><td colSpan={5}>Total depreciation for FY {fy}</td><td className="num">{inr(dep.totalDep)}</td><td colSpan={2}></td></tr></tfoot>
          </table>
        )}
        <CustomExport
          rows={dep.rows} dateField="purchaseDate" filenameBase="fixed-assets" printTitle="Fixed Assets — Depreciation Register" canExport={can("assets", "export")}
          buildSheets={(rows) => ({ DepreciationFY: rows.map((r) => ({ Asset: r.name, Block: r.block, Rate: r.rate, Cost: r.cost, OpeningWDV: r.applicable ? r.wdvStart : 0, Depreciation: r.applicable ? r.dep : 0, ClosingWDV: r.applicable ? r.wdvEnd : 0 })) })}
          printColumns={[{ label: "Asset", value: (r) => r.name }, { label: "Block", value: (r) => r.block }, { label: "Cost", value: (r) => inr(r.cost) }, { label: "Depreciation", value: (r) => (r.applicable ? inr(r.dep) : "—") }, { label: "Closing WDV", value: (r) => (r.applicable ? inr(r.wdvEnd) : "—") }]}
        />
        <div className="note-box">Depreciation follows the Section 32 written-down-value method (half rate if used under 180 days in the year of purchase), computed server-side so every screen agrees with the ledger.</div>
      </div>
    </div>
  );
}

/* ============================== FINANCIAL STATEMENTS ============================== */
function FinancialStatements({ fy, settings, can }) {
  const { call } = useApi();
  const [income, setIncome] = useState(null);
  const [capAcct, setCapAcct] = useState(null);
  const [bs, setBs] = useState(null);
  const [err, setErr] = useState("");
  const [customFrom, setCustomFrom] = useState("");
  const [customTo, setCustomTo] = useState("");
  const [customOpen, setCustomOpen] = useState(false);
  const [customIncome, setCustomIncome] = useState(null);
  const [customBusy, setCustomBusy] = useState(false);

  const fetchCustom = async () => {
    if (!customFrom || !customTo) return;
    setCustomBusy(true);
    try { setCustomIncome(await call(`/statements/income?from=${customFrom}&to=${customTo}`)); }
    catch (e) { setErr(e.message); }
    setCustomBusy(false);
  };
  const printCustom = () => {
    if (!customIncome) return;
    const win = document.getElementById("print-root");
    if (!win) { window.print(); return; }
    const rows = [
      { label: "Patient Collection Income", amount: customIncome.income.patientCollection },
      { label: "Referral Income", amount: customIncome.income.referral },
      { label: "Gift Income", amount: customIncome.income.gift },
      { label: "Total Income", amount: customIncome.income.total },
      ...customIncome.expenses.map((e) => ({ label: e.name, amount: e.amount })),
      { label: "Total Expense", amount: customIncome.totalExpense },
      { label: customIncome.netProfit >= 0 ? "Net Profit" : "Net Loss", amount: Math.abs(customIncome.netProfit) },
    ];
    win.innerHTML = `<h2>${escapeHtml(settings.clinicName)} — Income Statement</h2>
      <p style="color:#5B6B69;font-size:12px;">${customFrom} to ${customTo}</p>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead><tr><th style="text-align:left;border-bottom:2px solid #C9A227;padding:5px 6px;">Head</th><th style="text-align:right;border-bottom:2px solid #C9A227;padding:5px 6px;">Amount</th></tr></thead>
        <tbody>${rows.map((r) => `<tr><td style="padding:4px 6px;">${escapeHtml(r.label)}</td><td style="padding:4px 6px;text-align:right;">${inr(r.amount)}</td></tr>`).join("")}</tbody>
      </table>`;
    document.body.classList.add("printing-custom");
    window.print();
    setTimeout(() => { document.body.classList.remove("printing-custom"); win.innerHTML = ""; }, 300);
  };

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
            <tr><td>Gift Income</td><td className="num">{inr(income.income.gift)}</td></tr>
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
        {can("statements", "export") && <ExportRow onExcel={() => exportExcel("income-statement", { Income: [{ Head: "Patient Collection Income", Amount: income.income.patientCollection }, { Head: "Referral Income", Amount: income.income.referral }, { Head: "Gift Income", Amount: income.income.gift }], Expenses: income.expenses.map((r) => ({ Head: r.name, Amount: r.amount })), Summary: [{ TotalIncome: income.income.total, TotalExpense: income.totalExpense, NetProfit: income.netProfit }] })} />}
        {can("statements", "export") && (
          <div className="custom-export no-print">
            <button className="btn secondary small" type="button" onClick={() => setCustomOpen((o) => !o)}>📅 Custom range income statement</button>
            {customOpen && (
              <div className="custom-export-panel">
                <div><label>From</label><input type="date" value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} /></div>
                <div><label>To</label><input type="date" value={customTo} onChange={(e) => setCustomTo(e.target.value)} /></div>
                <button className="btn small" type="button" disabled={customBusy} onClick={fetchCustom}>{customBusy ? "Loading…" : "Fetch"}</button>
                {customIncome && (
                  <>
                    <button className="btn small" type="button" onClick={() => exportExcel("income-statement-custom", { Income: [{ Head: "Patient Collection Income", Amount: customIncome.income.patientCollection }, { Head: "Referral Income", Amount: customIncome.income.referral }, { Head: "Gift Income", Amount: customIncome.income.gift }], Expenses: customIncome.expenses.map((r) => ({ Head: r.name, Amount: r.amount })), Summary: [{ TotalIncome: customIncome.income.total, TotalExpense: customIncome.totalExpense, NetProfit: customIncome.netProfit }] })}>⬇ Excel</button>
                    <button className="btn small" type="button" onClick={printCustom}>⎙ PDF</button>
                  </>
                )}
              </div>
            )}
          </div>
        )}
      </div>
      <div className="card">
        <h2>Capital Account — as of {capAcct.asOf}</h2>
        <table><tbody>
          <tr><td>Capital introduced (cumulative)</td><td className="num">{inr(capAcct.capIntroduced)}</td></tr>
          <tr><td>Less: Drawings (cumulative)</td><td className="num">-{inr(capAcct.drawings)}</td></tr>
          <tr><td>Add: Cumulative net profit (all periods to date)</td><td className="num">{inr(capAcct.cumulativeNetProfit)}</td></tr>
          <tr style={{ fontWeight: 700, borderTop: "2px solid var(--accent)" }}><td>Closing Capital — {settings.proprietor}</td><td className="num">{inr(capAcct.closingCapital)}</td></tr>
        </tbody></table>
        {can("statements", "export") && <ExportRow onExcel={() => exportExcel("capital-account", { CapitalAccount: [{ CapitalIntroduced: capAcct.capIntroduced, Drawings: capAcct.drawings, CumulativeNetProfit: capAcct.cumulativeNetProfit, ClosingCapital: capAcct.closingCapital }] })} />}
      </div>
      <div className="card">
        <h2>Balance Sheet — as of {bs.asOf}</h2>
        <table><thead><tr><th>Liabilities & Capital</th><th className="num">Amount</th><th>Assets</th><th className="num">Amount</th></tr></thead>
          <tbody>
            <tr><td>Unsecured Loans</td><td className="num">{inr(bs.liabilities.unsecuredLoan)}</td><td>Cash & Bank</td><td className="num">{inr(bs.assets.cashBank)}</td></tr>
            <tr><td>Capital Account</td><td className="num">{inr(bs.capital.closingCapital)}</td><td>Sundry Debtors (patient dues)</td><td className="num">{inr(bs.assets.debtors)}</td></tr>
            <tr><td></td><td></td><td>Security Deposits Given</td><td className="num">{inr(bs.assets.securityDeposit)}</td></tr>
            <tr><td></td><td></td><td>Fixed Assets (net of depreciation)</td><td className="num">{inr(bs.assets.fixedAssetsNet)}</td></tr>
            <tr style={{ fontWeight: 700, borderTop: "2px solid var(--accent)" }}><td>Total</td><td className="num">{inr(bs.liabilities.unsecuredLoan + bs.capital.closingCapital)}</td><td>Total</td><td className="num">{inr(bs.assets.total)}</td></tr>
          </tbody>
        </table>
        <div style={{ marginTop: 10 }}><span className={"balance-tag " + (bs.ties ? "ok" : "bad")}>{bs.ties ? "Balance sheet ties out" : `Off by ${inr(Math.abs(bs.liabilities.unsecuredLoan + bs.capital.closingCapital - bs.assets.total))}`}</span></div>
        {can("statements", "export") && <ExportRow onExcel={() => exportExcel("balance-sheet", { BalanceSheet: [{ UnsecuredLoans: bs.liabilities.unsecuredLoan, Capital: bs.capital.closingCapital, CashBank: bs.assets.cashBank, Debtors: bs.assets.debtors, SecurityDeposits: bs.assets.securityDeposit, FixedAssetsNet: bs.assets.fixedAssetsNet, TotalAssets: bs.assets.total }] })} />}
        <div className="note-box">Model assumes the clinic settles expenses directly out of collections (no supplier credit tracked), so Liabilities + Capital equal Assets by construction. Add a Sundry Creditors flow if medicine purchases start going on credit.</div>
      </div>
    </div>
  );
}

/* ============================== USER ACCESS REPORT ============================== */
function AccessReport({ can }) {
  const { call } = useApi();
  const today = todayISO();
  const monthAgo = (() => { const d = new Date(); d.setDate(d.getDate() - 30); return d.toISOString().slice(0, 10); })();
  const [from, setFrom] = useState(monthAgo);
  const [to, setTo] = useState(today);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");

  const load = () => {
    setLoading(true); setErr("");
    call(`/audit-log?from=${from}&to=${to}`).then(setRows).catch((e) => setErr(e.message)).finally(() => setLoading(false));
  };
  useEffect(load, [call]); // eslint-disable-line

  const moduleLabel = (m) => PERMISSION_MODULES.find((x) => x.key === m)?.label || m;
  const stamp = (r) => new Date(r.created_at).toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
  const exportRows = rows.map((r) => ({ ...r, date: String(r.created_at).slice(0, 10) }));

  return (
    <div>
      <div className="card">
        <h2>User Access Report</h2>
        <p style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: -8 }}>Who accessed which module, and when — logged automatically as staff use the app.</p>
        <div className="form-grid" style={{ maxWidth: 340 }}>
          <div><label>From</label><input type="date" value={from} onChange={(e) => setFrom(e.target.value)} /></div>
          <div><label>To</label><input type="date" value={to} onChange={(e) => setTo(e.target.value)} /></div>
        </div>
        <button className="btn" type="button" onClick={load}>Apply range</button>
        <ErrorNote msg={err} />
        {loading ? <div className="empty">Loading…</div> : rows.length === 0 ? <div className="empty">No access recorded in this range.</div> : (
          <table>
            <thead><tr><th>Date &amp; Time</th><th>User</th><th>Module</th><th>Action</th></tr></thead>
            <tbody>{rows.map((r) => (<tr key={r.id}><td>{stamp(r)}</td><td>{r.user_label}</td><td>{moduleLabel(r.module)}</td><td style={{ textTransform: "capitalize" }}>{r.action}</td></tr>))}</tbody>
          </table>
        )}
        <CustomExport
          rows={exportRows} dateField="date" filenameBase="user-access-report" printTitle="User Access Report" canExport={can("auditLog", "export")}
          buildSheets={(data) => ({ AccessLog: data.map((r) => ({ DateTime: stamp(r), User: r.user_label, Module: moduleLabel(r.module), Action: r.action })) })}
          printColumns={[{ label: "Date & Time", value: (r) => stamp(r) }, { label: "User", value: (r) => r.user_label }, { label: "Module", value: (r) => moduleLabel(r.module) }, { label: "Action", value: (r) => r.action }]}
        />
      </div>
    </div>
  );
}

/* ============================== SETTINGS ============================== */
function SettingsPage({ settings, updateSettings, session, origin, capital, addCapital, removeCapital, loans, deposits, addOtherBalance, updateOtherBalance, removeOtherBalance, can }) {
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
        {can("statements", "write") && <button className="btn" type="button" onClick={saveCap}>Save entry</button>}
        <ErrorNote msg={capErr} />
        <table style={{ marginTop: 14 }}><thead><tr><th>Date</th><th>Type</th><th className="num">Amount</th><th>Note</th><th></th></tr></thead>
          <tbody>{[...capital].sort((a, b) => (a.date < b.date ? 1 : -1)).map((c) => (<tr key={c.id}><td>{c.date}</td><td>{c.type}</td><td className="num">{inr(c.amount)}</td><td>{c.note}</td><td>{can("statements", "delete") && <button className="btn danger small" type="button" onClick={() => removeCapital(c.id)}>Delete</button>}</td></tr>))}</tbody>
        </table>
      </div>

      <LoanDepositSection
        title="Unsecured loans" category="unsecured_loan" types={["Taken", "Repaid"]} partyLabel="Lender's name"
        items={loans} addOtherBalance={addOtherBalance} updateOtherBalance={updateOtherBalance} removeOtherBalance={removeOtherBalance} can={can}
        note='Shows as "Unsecured Loans" on the liability side of the Balance Sheet. "Taken" adds to cash and the loan balance; "Repaid" reduces both.'
      />
      <LoanDepositSection
        title="Security deposits given" category="security_deposit" types={["Given", "Refunded"]} partyLabel="Given to (e.g. property owner)"
        items={deposits} addOtherBalance={addOtherBalance} updateOtherBalance={updateOtherBalance} removeOtherBalance={removeOtherBalance} can={can}
        note='Shows as a Current Asset on the Balance Sheet. "Given" moves the amount out of cash and into this deposit asset; "Refunded" moves it back.'
      />

      <div className="card">
        <h2>Connection</h2>
        <p style={{ fontSize: 13, color: "var(--ink-soft)" }}>API server: <code>{origin}</code></p>
        <p style={{ fontSize: 13, color: "var(--ink-soft)" }}>Logged in as {session.name} ({session.userId}) — role: {session.role}</p>
        <div className="note-box">{session.role === "Admin" ? "Manage staff accounts and permissions under \"User Approvals\" in the sidebar." : "New staff accounts need an Admin to approve them and set access under \"User Approvals\" before they can log in."} Same user ID and password work on both the web and mobile views of this app, since both call the same API server.</div>
      </div>
    </div>
  );
}

function LoanDepositSection({ title, category, types, partyLabel, items, addOtherBalance, updateOtherBalance, removeOtherBalance, can, note }) {
  const blank = { txnType: types[0], partyName: "", amount: "", date: todayISO(), note: "" };
  const [form, setForm] = useState(blank);
  const [editingId, setEditingId] = useState(null);
  const [err, setErr] = useState("");

  const startEdit = (item) => { setEditingId(item.id); setForm({ txnType: item.txnType, partyName: item.partyName || "", amount: String(item.amount), date: item.date, note: item.note || "" }); setErr(""); };
  const cancelEdit = () => { setEditingId(null); setForm(blank); setErr(""); };
  const save = async () => {
    setErr(""); if (!form.amount) { setErr("Enter an amount."); return; }
    const payload = { category, ...form, amount: Number(form.amount) };
    try {
      if (editingId) { await updateOtherBalance(editingId, payload); setEditingId(null); }
      else { await addOtherBalance(payload); }
      setForm(blank);
    } catch (e) { setErr(e.message); }
  };
  const remove = async (id) => { try { await removeOtherBalance(id, category); if (editingId === id) cancelEdit(); } catch (e) { setErr(e.message); } };

  const balance = items.reduce((s, x) => s + (x.txnType === types[0] ? x.amount : -x.amount), 0);

  return (
    <div className="card">
      <h2>{title} — current balance: {inr(balance)}</h2>
      <div className="form-grid">
        <div><label>Date</label><input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></div>
        <div><label>Type</label><select value={form.txnType} onChange={(e) => setForm({ ...form, txnType: e.target.value })}>{types.map((t) => <option key={t} value={t}>{t}</option>)}</select></div>
        <div><label>{partyLabel}</label><input type="text" value={form.partyName} onChange={(e) => setForm({ ...form, partyName: e.target.value })} /></div>
        <div><label>Amount (₹)</label><input type="number" value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></div>
        <div><label>Note</label><input type="text" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></div>
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        {(editingId ? can("statements", "edit") : can("statements", "write")) && <button className="btn" type="button" onClick={save}>{editingId ? "Update entry" : "Save entry"}</button>}
        {editingId && <button className="btn secondary" type="button" onClick={cancelEdit}>Cancel edit</button>}
      </div>
      <ErrorNote msg={err} />
      {items.length > 0 && (
        <table style={{ marginTop: 14 }}><thead><tr><th>Date</th><th>Type</th><th>{partyLabel}</th><th className="num">Amount</th><th>Note</th><th></th></tr></thead>
          <tbody>{[...items].sort((a, b) => (a.date < b.date ? 1 : -1)).map((x) => (
            <tr key={x.id}><td>{x.date}</td><td>{x.txnType}</td><td>{x.partyName}</td><td className="num">{inr(x.amount)}</td><td>{x.note}</td>
              <td style={{ display: "flex", gap: 6 }}>
                {can("statements", "edit") && <button className="btn secondary small" type="button" onClick={() => startEdit(x)}>Edit</button>}
                {can("statements", "delete") && <button className="btn danger small" type="button" onClick={() => remove(x.id)}>Delete</button>}
              </td></tr>
          ))}</tbody>
        </table>
      )}
      <div className="note-box">{note}</div>
    </div>
  );
}

/* ============================== ADMIN: USER APPROVALS ============================== */
function AdminUsers() {
  const { call } = useApi();
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState("");
  const [drafts, setDrafts] = useState({}); // userId -> { role, permissions }

  const load = () => {
    setLoading(true); setErr("");
    call("/admin/users").then((rows) => {
      setUsers(rows);
      const d = {};
      rows.forEach((u) => { d[u.id] = { role: u.role || "Staff", permissions: u.permissions && Object.keys(u.permissions).length ? u.permissions : Object.fromEntries(PERMISSION_MODULES.map((m) => [m.key, { level: "none", export: false }])) }; });
      setDrafts(d);
    }).catch((e) => setErr(e.message)).finally(() => setLoading(false));
  };
  useEffect(load, [call]);

  const setLevel = (userId, moduleKey, level) => {
    setDrafts((p) => ({ ...p, [userId]: { ...p[userId], permissions: { ...p[userId].permissions, [moduleKey]: { ...p[userId].permissions[moduleKey], level } } } }));
  };
  const setExport = (userId, moduleKey, exp) => {
    setDrafts((p) => ({ ...p, [userId]: { ...p[userId], permissions: { ...p[userId].permissions, [moduleKey]: { ...p[userId].permissions[moduleKey], export: exp } } } }));
  };
  const setRole = (userId, role) => setDrafts((p) => ({ ...p, [userId]: { ...p[userId], role } }));

  const save = async (userId, activate) => {
    setErr("");
    try {
      await call(`/admin/users/${userId}/permissions`, { method: "PUT", body: { role: drafts[userId].role, permissions: drafts[userId].permissions, activate } });
      load();
    } catch (e) { setErr(e.message); }
  };
  const deactivate = async (userId) => {
    try { await call(`/admin/users/${userId}/deactivate`, { method: "PUT" }); load(); } catch (e) { setErr(e.message); }
  };

  if (loading) return <div className="empty">Loading users…</div>;
  const pending = users.filter((u) => u.status === "pending_approval");
  const others = users.filter((u) => u.status !== "pending_approval");

  const UserCard = ({ u }) => {
    const d = drafts[u.id] || { role: "Staff", permissions: {} };
    return (
      <div className="card" key={u.id}>
        <h2>{u.name} <span style={{ fontWeight: 400, fontSize: 13, color: "var(--ink-soft)" }}>({u.user_id})</span></h2>
        <p style={{ fontSize: 12.5, color: "var(--ink-soft)", marginTop: -8 }}>
          Status: <b>{u.status === "pending_approval" ? "Awaiting approval" : u.status === "pending_otp" ? "Awaiting email verification" : "Active"}</b>
          {u.email ? ` · ${u.email}` : ""}{u.mobile ? ` · ${u.mobile}` : ""}
        </p>
        <div className="form-grid" style={{ maxWidth: 260 }}>
          <div><label>Role label</label><input type="text" value={d.role} onChange={(e) => setRole(u.id, e.target.value)} placeholder="e.g. Nurse, Reception" /></div>
        </div>
        <table>
          <thead><tr><th>Module</th><th>Access level</th><th>Export</th></tr></thead>
          <tbody>
            {PERMISSION_MODULES.map((m) => (
              <tr key={m.key}>
                <td>{m.label}</td>
                <td>
                  <select value={d.permissions[m.key]?.level || "none"} onChange={(e) => setLevel(u.id, m.key, e.target.value)}>
                    {LEVELS.map((lv) => <option key={lv} value={lv}>{LEVEL_LABELS[lv]}</option>)}
                  </select>
                </td>
                <td><input type="checkbox" checked={!!d.permissions[m.key]?.export} onChange={(e) => setExport(u.id, m.key, e.target.checked)} /></td>
              </tr>
            ))}
          </tbody>
        </table>
        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          {u.status === "pending_approval" && <button className="btn" type="button" onClick={() => save(u.id, true)}>Approve &amp; activate</button>}
          {u.status === "active" && <button className="btn secondary" type="button" onClick={() => save(u.id, false)}>Save changes</button>}
          {u.status === "active" && u.role !== "Admin" && <button className="btn danger" type="button" onClick={() => deactivate(u.id)}>Deactivate</button>}
        </div>
      </div>
    );
  };

  return (
    <div>
      <ErrorNote msg={err} />
      {pending.length > 0 && (
        <>
          <div className="card"><h2>Pending approval ({pending.length})</h2><p style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>These accounts registered themselves and can't log in until you set their access below.</p></div>
          {pending.map((u) => <UserCard key={u.id} u={u} />)}
        </>
      )}
      <div className="card"><h2>All accounts</h2></div>
      {others.map((u) => <UserCard key={u.id} u={u} />)}
    </div>
  );
}
