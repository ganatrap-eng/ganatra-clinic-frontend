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
