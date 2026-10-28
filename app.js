"use strict";

/* ============================================================
   State
   ============================================================ */
const state = {
  ar: { rows: [], headers: [], map: {}, file: null, buffer: null, saved: false },
  ap: { rows: [], headers: [], map: {}, file: null, buffer: null, saved: false },
  cf: { rows: [], file: null, buffer: null, saved: false, sheet: "", col: "" },
  ds: { rows: [], file: null, buffer: null, saved: false },
  // Cash Out sources. byDate = { "YYYY-MM-DD": amount } (cash going out); from = "file" (uploaded now) or "saved" (from the last saved summary)
  ln: { byDate: {}, file: null, buffer: null, saved: false, info: "", from: "" },   // Loan Payment Schedule
  uc: { byDate: {}, file: null, buffer: null, saved: false, info: "", from: "" },   // Unreleased Checks
  po: { byDate: {}, file: null, buffer: null, saved: false, info: "", from: "" },   // Projected Cash Outflow   // Disposal Schedule (Proceeds from Disposal)   // Cash Flow Report (Retail trend)
};
// Opening cash balances (beginning balance of Week 1). USD amounts are converted at 62.50 PHP per USD.
const USD_RATE = 62.5;
const accounts = [
  { bank: "LDB", no: "", purpose: "", bal: 0 },
  { bank: "BPI", no: "", purpose: "", bal: 0 },
  { bank: "UB", no: "", purpose: "", bal: 0 },
  { bank: "BDO", no: "", purpose: "", bal: 0 },
  { bank: "PBCom", no: "", purpose: "", bal: 0 },
  { bank: "CBC", no: "", purpose: "", bal: 0 },
  { bank: "AUB", no: "", purpose: "", bal: 0 },
  { bank: "MBTC", no: "", purpose: "USD account (× 62.50)", bal: 0 },
];
let overrides = {};          // { rowKey: { "YYYY-MM-DD" (week start): number } } typed amounts, kept by calendar week
let editing = false;
let customRows = [];         // rows added by the user: { id, name, sign (-1 deduct / +1 add), vals: { "YYYY-MM-DD" (week start): amount } }
let nextRowId = 1;
let report = null;           // last computed report
let saved = false;           // true once the summary has been saved this session
let lastSaved = null;        // { name, sig } of the newest file in "Cashflow Summary"
let rootDir = null;          // FileSystemDirectoryHandle chosen by the user
let baseDir = null;          // the "FILES" folder inside it (or the chosen folder itself when it is called FILES)

const $ = (id) => document.getElementById(id);
const DAY = 86400000;
const UPLOAD_DIR = "Uploaded Files";
const SUMMARY_DIR = "Cashflow Summary";

/* ============================================================
   Helpers
   ============================================================ */
function mondayOf(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() - ((x.getDay() + 6) % 7));
  return x;
}
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function toISO(d) { return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; }
function fmtDate(d) { return d.toLocaleDateString("en-PH", { month: "short", day: "numeric" }); }
function fmtNum(n) {
  if (!n) return "-";
  const s = Math.abs(n).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return n < 0 ? `(${s})` : s;
}
function stamp() {
  const d = new Date(), p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}
function escapeHtml(t) { return String(t).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }
function sumObj(o) { return Object.values(o).reduce((a, b) => a + b, 0); }
function parseNumber(v) {
  if (typeof v === "number") return v;
  if (v == null) return NaN;
  let s = String(v).trim();
  if (!s || s === "-") return NaN;
  const neg = /^\(.*\)$/.test(s) || s.startsWith("-");
  const n = parseFloat(s.replace(/[^0-9.]/g, ""));
  return isNaN(n) ? NaN : neg ? -n : n;
}
function parseDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === "number" && v > 20000 && v < 80000) {
    const d = new Date(Math.round((v - 25569) * DAY));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  if (typeof v === "string" && v.trim()) {
    const d = new Date(v.trim());
    if (!isNaN(d) && d.getFullYear() > 1990) return d;
  }
  return null;
}
let toastTimer;
function toast(html, isErr = false, ms = 5000) {
  const t = $("toast");
  t.innerHTML = html; t.className = "toast" + (isErr ? " err" : "");
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.add("hidden"), ms);
}

/* ============================================================
   Folder access (File System Access API — Chrome / Edge)
   ============================================================ */
const HAS_FS = "showDirectoryPicker" in window;
// why folder access may be unavailable (shown to the user instead of a dead button)
function fsUnavailableReason() {
  if (HAS_FS) return "";
  if (!window.isSecureContext) return `This page was opened as ${location.protocol}//${location.host || ""} — browsers turn off folder access on plain "http" addresses. Open the system from its file path instead (e.g. Z:\\Hertz\\V3.6\\index.html or \\\\SERVER\\Hertz\\V3.6\\index.html), or use https.`;
  return "This browser cannot save into folders. Use Google Chrome or Microsoft Edge (up to date).";
}

function idb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open("hertz-cashflow", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
}
async function idbGet(k) { let db; try { db = await idb(); } catch (_) { return null; } return new Promise((res) => { const q = db.transaction("kv").objectStore("kv").get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(null); }); }
async function idbSet(k, v) { let db; try { db = await idb(); } catch (_) { return; } return new Promise((res) => { const tx = db.transaction("kv", "readwrite"); tx.objectStore("kv").put(v, k); tx.oncomplete = res; tx.onerror = res; }); }

// "FILES" holds the "Uploaded Files" and "Cashflow Summary" folders. Choose either the project folder or FILES itself.
async function resolveBase() {
  if (/^files$/i.test(rootDir.name)) return rootDir;
  try { return await rootDir.getDirectoryHandle("FILES"); } catch (_) {}
  try { await rootDir.getDirectoryHandle(SUMMARY_DIR); return rootDir; } catch (_) {}      // the FILES folder under another name
  if (!confirm(`"${rootDir.name}" has no FILES folder.

OK = create FILES inside "${rootDir.name}"
Cancel = choose another folder (pick the shared FILES folder, or the Hertz folder that contains it)`)) {
    const e = new Error("No FILES folder chosen."); e.name = "AbortError"; throw e;
  }
  return rootDir.getDirectoryHandle("FILES", { create: true });
}

function showFolder(status) {
  const dot = $("folder-dot"), txt = $("folder-text"), btn = $("btn-folder"), lbl = $("folder-label");
  dot.className = "dot" + (status === "ok" ? " ok" : "");
  if (status === "ok") { txt.textContent = `Saving to: ${(baseDir || rootDir).name}`; lbl.textContent = "Change folder"; }
  else if (status === "reconnect") { txt.textContent = `Folder "${rootDir.name}" needs permission`; lbl.textContent = "Reconnect"; }
  else { txt.textContent = HAS_FS ? "No save folder connected" : fsUnavailableReason(); lbl.textContent = HAS_FS ? "Connect folder" : "Folder access off — why?"; }
  btn.disabled = false;                                    // clicking always explains what to do
  btn.title = txt.textContent;
}

// The folder window can take a few seconds to appear on a network drive (or open behind the browser):
// a second click while it is open must not start another one.
let pickerOpen = false;
async function connectFolder() {
  if (pickerOpen) { toast("The folder window is already open<small>On the network drive it can take a few seconds to appear, or it may be behind this browser window (check the taskbar). Choose the FILES folder there.</small>", false, 7000); return; }
  pickerOpen = true; $("btn-folder").disabled = true; $("folder-label").textContent = "Choosing folder…";
  try { await connectFolderInner(); }
  finally { pickerOpen = false; $("btn-folder").disabled = false; if ($("folder-label").textContent === "Choosing folder…") showFolder(rootDir && baseDir ? "ok" : rootDir ? "reconnect" : undefined); }
}
async function connectFolderInner() {
  try {
    if (rootDir && (await rootDir.queryPermission({ mode: "readwrite" })) !== "granted") {
      if ((await rootDir.requestPermission({ mode: "readwrite" })) === "granted") { await onFolderReady(); return; }
    }
    if (!HAS_FS) { toast("Folder access is not available<small>" + escapeHtml(fsUnavailableReason()) + "</small>", true, 15000); return; }
    rootDir = await window.showDirectoryPicker({ mode: "readwrite", id: "hertz-cashflow" });
    try { await idbSet("root", rootDir); } catch (_) { /* remembering is optional */ }
    await onFolderReady();
  } catch (e) {
    if (e.name === "AbortError") return;
    if (/already active/i.test(e.message || "")) { toast("The folder window is already open<small>It may be behind this browser window (check the taskbar). Choose the FILES folder there.</small>", false, 7000); return; }
    const tips = {
      SecurityError: "The browser or a company policy blocks folder access on this page. Open the system from its file path in Chrome/Edge (not an http address), or ask IT to allow the File System Access setting.",
      NotAllowedError: "Permission was not given. Click Connect folder again and choose Allow / Edit files.",
      NotFoundError: "The folder was moved, renamed or the network drive is not reachable. Check the LAN connection / mapped drive and choose the folder again.",
      InvalidModificationError: "The folder could not be changed (read-only?). Ask IT for write access to the FILES folder on the server.",
      NoModificationAllowedError: "The folder is read-only for your account. Ask IT for read + write access to the FILES folder on the server.",
      TypeError: "Choose a folder inside the share (e.g. Z:\\Hertz\\FILES), not the top of the drive.",
    };
    toast(`Could not connect the folder<small>${escapeHtml(e.name)}: ${escapeHtml(e.message)}<br>${escapeHtml(tips[e.name] || "Choose the FILES folder (inside the shared Hertz folder), not the top of the drive, and check you have write access.")}</small>`, true, 15000);
  }
}
async function onFolderReady() {
  baseDir = await resolveBase();
  await baseDir.getDirectoryHandle(UPLOAD_DIR, { create: true });
  await baseDir.getDirectoryHandle(SUMMARY_DIR, { create: true });
  showFolder("ok");
  await openShared();
  toast(`Folder connected<small>"${UPLOAD_DIR}" and "${SUMMARY_DIR}" are ready inside ${escapeHtml(baseDir.name)}.</small>`);
  for (const k of ["ar", "ap"]) if (state[k].file && !state[k].saved) await saveOriginal(k); // flush files uploaded before connecting
}
async function restoreFolder() {
  showFolder();
  if (!HAS_FS) return;
  try {
    const h = await idbGet("root");
    if (!h) return;
    rootDir = h;
    if ((await h.queryPermission({ mode: "readwrite" })) === "granted") { await onFolderReady_silent(); }
    else showFolder("reconnect");
  } catch (_) { showFolder(); }
}
async function onFolderReady_silent() {
  baseDir = await resolveBase();
  await baseDir.getDirectoryHandle(UPLOAD_DIR, { create: true });
  await baseDir.getDirectoryHandle(SUMMARY_DIR, { create: true });
  showFolder("ok");
  await openShared();
}
// Load everyone's latest parts (FILES/Current). The first time (nothing shared yet) the last saved summary is used as before.
async function openShared() {
  await baseDir.getDirectoryHandle(CURRENT_DIR, { create: true });
  showProgress("Loading everyone's latest updates…");
  try { sharedFound = await loadShared(); }
  catch (e) { toast("Could not load the shared updates:<small>" + escapeHtml(e.message) + "</small>", true, 10000); }
  showProgress("Checking the last saved summary…");
  await loadLatestSummary();
  if (sharedFound) { renderNews(); if (report) showTab("summary"); } else showNothingShared();
  clearInterval(pollTimer);
  pollTimer = setInterval(checkNews, 60000);
  updateSaveState();
}
async function openDir(path) {
  let dir = baseDir;
  for (const part of [].concat(path)) dir = await dir.getDirectoryHandle(part, { create: true });
  return dir;
}
async function writeTo(dirName, fileName, data) {
  const dir = await openDir(dirName);
  const fh = await dir.getFileHandle(fileName, { create: true });
  const w = await fh.createWritable();
  await w.write(data); await w.close();
}
// Returns fileName, or "name (2).ext", "name (3).ext"... if a file with that name already exists in the folder.
async function uniqueName(dirName, fileName) {
  const dir = await openDir(dirName);
  const dot = fileName.lastIndexOf("."), base = dot > 0 ? fileName.slice(0, dot) : fileName, ext = dot > 0 ? fileName.slice(dot) : "";
  for (let n = 1; ; n++) {
    const candidate = n === 1 ? fileName : `${base} (${n})${ext}`;
    try { await dir.getFileHandle(candidate); } catch (_) { return candidate; }
  }
}
async function folderReady() {
  return rootDir && baseDir && (await rootDir.queryPermission({ mode: "readwrite" })) === "granted";
}

/* ============================================================
   File upload + parsing
   ============================================================ */
const FIELD_LABELS = { party: "Customer / Vendor", due: "Due date", amount: "Amount (invoice)", type: "Invoice type", vat: "VAT type", inv: "Invoice no.", proj: "Projected date", rem: "Remarks" };
const AR_ONLY_FIELDS = ["type", "vat", "inv", "proj", "rem"];
const KIND_LABEL = { ar: "AR_Aging", ap: "AP_Aging", cf: "CashFlow_Report", ds: "Disposal_Schedule", ln: "Loan_Payment_Schedule", uc: "Unreleased_Checks", po: "Projected_Cash_Outflow" };

// Reads the file, waits for its handler, then marks the part as changed by this user (unless it came from FILES/Current).
function readInto(reader, file, kind) {
  return new Promise((res) => {
    const h = reader.onload;
    reader.onload = async (e) => {
      try { await h(e); } finally {
        const s = state[kind];
        if (s && s.buffer) { if (sharedLoading) Object.assign(s, { dirty: false, copySaved: true, saved: true, handle: null }); else { s.dirty = true; s.sharedBy = null; } }
        res();
      }
    };
    reader.onerror = () => res();
    reader.readAsArrayBuffer(file);
  });
}
function handleFile(kind, file) {
  const info = $("info-" + kind);
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
      // The AR report has Summary / Aging by Customer / Detail sheets: the invoices are in "Detail".
      const sheetName = wb.SheetNames.find((n) => /^detail$/i.test(n)) || wb.SheetNames[0];
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
      const { headerIdx, headers } = findHeader(grid);
      const r0 = XLSX.utils.decode_range(wb.Sheets[sheetName]["!ref"] || "A1").s;
      grid.forEach((row, i) => (row.__r = r0.r + i));                       // sheet row of each line (to know which cells were edited)
      const rows = grid.slice(headerIdx + 1).filter((r) => r.some((c) => c != null && c !== ""));
      state[kind].c0 = r0.c;
      Object.assign(state[kind], { rows, headers, file, sheet: sheetName, buffer: e.target.result, saved: false, saveErr: null, map: guessMapping(headers) });
      $("drop-" + kind).classList.add("loaded");
      renderMapping(kind);
      saved = false;
      await saveOriginal(kind);
      refresh();
      showTab("summary");
    } catch (err) {
      info.innerHTML = `<div class="err">Could not read this file: ${escapeHtml(err.message)}</div>`;
    }
  };
  return readInto(reader, file, kind);
}

/* ============================================================
   Bank Balances file -> Setup > Cash balances (Beginning Balance of Week 1)
   Layout: a heading row with "Bank" and "Balance" (also "Account no." / "Purpose"), one row per account, then "Total".
   ============================================================ */
let bankUpload = null;          // { file, buffer, copied } — the uploaded Bank Balances file (copied to Uploaded Files on Save)
function parseBankBalances(wb) {
  for (const name of wb.SheetNames) {
    const g = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const t = (c) => (typeof c === "string" ? c.trim().toLowerCase() : "");
    const hi = g.slice(0, 30).findIndex((r) => r.some((c) => t(c) === "bank") && r.some((c) => /balance/.test(t(c))));
    if (hi < 0) continue;
    const H = g[hi].map(t), col = (re) => H.findIndex((x) => re.test(x));
    const iB = col(/^bank$/), iN = col(/account/), iP = col(/purpose/), iV = col(/balance/);
    const rows = [];
    for (const r of g.slice(hi + 1)) {
      const bank = r[iB] == null ? "" : String(r[iB]).trim();
      if (/^(grand\s*)?total/i.test(bank)) break;
      const v = typeof r[iV] === "number" ? r[iV] : parseNumber(r[iV]);
      if (!bank && (v == null || isNaN(v))) continue;
      rows.push({ bank, no: iN >= 0 && r[iN] != null ? String(r[iN]).trim() : "", purpose: iP >= 0 && r[iP] != null ? String(r[iP]).trim() : "", bal: isNaN(v) || v == null ? 0 : Math.round(v * 100) / 100 });
    }
    if (rows.length) return { rows, sheet: name };
  }
  throw new Error('No table with "Bank" and "Balance" headings was found.');
}
function handleBankFile(file) {
  const reader = new FileReader();
  return new Promise((res) => {
    reader.onload = (e) => {
      try {
        const res2 = parseBankBalances(XLSX.read(e.target.result, { type: "array", cellDates: true }));
        accounts.splice(0, accounts.length, ...res2.rows);
        bankUpload = { file, buffer: e.target.result, copied: false };
        saveAccounts();                                    // shared with everyone on Save summary
        renderAccounts();
        $("drop-bk").classList.add("loaded");
        $("mstat-bk").innerHTML = `<b>${escapeHtml(file.name)}</b><br>${res2.rows.length} account${res2.rows.length === 1 ? "" : "s"} · total ${fmtNum(acctTotal())} · sheet "${escapeHtml(res2.sheet)}"`;
        saved = false;
        if (report || hasUploads()) refresh(); else updateSaveState();
        toast(`Bank balances loaded from ${escapeHtml(file.name)}<small>${res2.rows.length} accounts · Beginning Balance of Week 1 = ${fmtNum(acctTotal())}. Click Save summary to share them.</small>`, false, 7000);
      } catch (err) {
        $("mstat-bk").innerHTML = `<span class="err">Could not read this file: ${escapeHtml(err.message)}</span>`;
      }
      res();
    };
    reader.onerror = () => res();
    reader.readAsArrayBuffer(file);
  });
}

/* ============================================================
   Cash Flow Report -> Retail trend ("Actual Retail Collection")
   ============================================================ */
function parseCashflowReport(wb) {
  const name = wb.SheetNames.find((n) => /^summary$/i.test(n)) || wb.SheetNames[0];
  const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
  const isTxt = (c, re) => typeof c === "string" && re.test(c.trim());
  const hi = grid.findIndex((r) => r.some((c) => isTxt(c, /^date$/i)) && r.some((c) => isTxt(c, /actual\s*retail/i)));
  if (hi < 0) throw new Error('Could not find the "Date" and "Actual Retail Collection" headings on the Summary sheet.');
  const h = grid[hi].map((c) => (typeof c === "string" ? c.trim().toLowerCase() : ""));
  const iDate = h.findIndex((x) => x === "date"), iAct = h.findIndex((x) => /actual\s*retail/.test(x)), iProj = h.findIndex((x) => /projected\s*retail/.test(x));
  const o = XLSX.utils.decode_range(wb.Sheets[name]["!ref"] || "A1").s, rows = [];
  grid.slice(hi + 1).forEach((r, j) => {
    const d = parseDate(r[iDate]);
    if (!d) return;
    const a = r[iAct], p = iProj >= 0 ? r[iProj] : null;
    rows.push({ date: new Date(d.getFullYear(), d.getMonth(), d.getDate()), actual: typeof a === "number" ? a : null, projected: typeof p === "number" ? p : null, r: o.r + hi + 1 + j });
  });
  return { rows, sheet: name, col: XLSX.utils.encode_col(o.c + iAct), hdr: o.r + hi, cAct: o.c + iAct };
}

function handleCashflowReport(file) {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      let wb = XLSX.read(e.target.result, { type: "array", cellDates: true, sheets: ["Summary"] });   // only the Summary sheet: much faster
      if (!wb.SheetNames.length || !wb.Sheets[wb.SheetNames[0]]) wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
      const res = parseCashflowReport(wb);
      Object.assign(state.cf, { rows: res.rows, file, sheet: res.sheet, col: res.col, buffer: e.target.result, saved: false, saveErr: null });
      // Weekly rhythm: Week 1 is the NEXT week after the report's last actual day (Fri Oct 9 -> Mon Oct 12).
      const lastAct = res.rows.filter((x) => x.actual != null).reduce((m, x) => (x.date > m ? x.date : m), new Date(0));
      if (lastAct.getTime() > 0) {
        const next = addDays(lastAct, (8 - lastAct.getDay()) % 7 || 7);
        $("opt-start").value = toISO(next); $("opt-asof").value = toISO(next);
        $("opt-start").title = "Set automatically: the Monday after the Cash Flow Report's last actual date";
      }
      $("drop-cf").classList.add("loaded");
      await saveOriginal("cf");
      $("review").classList.remove("hidden"); $("empty").classList.add("hidden"); $("tabs").classList.remove("hidden");
      renderRetail();
      refresh();
      showTab("summary");
    } catch (err) {
      $("mstat-cf").innerHTML = `<span class="err">Could not read this file: ${escapeHtml(err.message)}</span>`;
    }
  };
  return readInto(reader, file, "cf");
}

const compactPeso = (n) => (Math.abs(n) >= 1e6 ? (n / 1e6).toFixed(2) + "M" : Math.abs(n) >= 1e3 ? (n / 1e3).toFixed(0) + "K" : n.toFixed(0));
function retailWeeks() {
  const map = new Map();
  for (const x of state.cf.rows) {
    if (x.actual == null) continue;
    const m = mondayOf(x.date), key = toISO(m);
    if (!map.has(key)) map.set(key, { key, monday: m, days: [], actual: 0, projected: 0 });
    const w = map.get(key);
    w.days.push(x); w.actual += x.actual; w.projected += x.projected || 0;
  }
  return [...map.values()].sort((a, b) => a.monday - b.monday);
}

// Average per business day over the window chosen on the Retail Trend tab, and the report's last actual date.
function retailStats() {
  const weeks = retailWeeks();
  if (!weeks.length) return null;
  const sel = $("rt-weeks").value, n = sel === "all" ? weeks.length : Math.min(+sel, weeks.length);
  const days = weeks.slice(-n).flatMap((w) => w.days);
  const avg = days.reduce((s, x) => s + x.actual, 0) / days.length;
  const last = weeks[weeks.length - 1].days.reduce((m, x) => (x.date > m ? x.date : m), new Date(0));
  return { avg, n, days: days.length, last, weeks };
}
// Retail (Credit Card) row: actual weekly total for weeks the report already covers, the average for the weeks after.
function retailForecast(wk) {
  const out = wk.map(() => 0);
  if (!state.cf.rows.length) return out;
  const st = retailStats();
  if (!st) return out;
  const basis = $("rf-basis").value, M = basis === "trend" ? retailModel() : null;
  const perWeek = basis === "daily" ? st.avg : st.avg * 5;
  const byKey = new Map(st.weeks.map((w) => [w.key, w]));
  let given = false;
  wk.forEach((w, i) => {
    if (w.start <= st.last) { const a = byKey.get(toISO(w.start)); out[i] = a ? a.actual : 0; }
    else if (M) out[i] = M.predict(M.weekIndex(w.start));                 // weekly trend, every week after the last actual date
    else if ($("rf-apply").value === "all" || !given) { out[i] = perWeek; given = true; }
  });
  return out;
}

// Weekly history (date range) -> model used both by the trend card and by the Summary's Retail row.
function retailModel() {
  if (!state.cf.rows.length) return null;
  const from = parseDate($("rx-from").value), to = parseDate($("rx-to").value);
  if (!from || !to || to < from) return null;
  const all = retailWeeks();
  const hist = all.map((w) => ({ ...w, days: w.days.filter((x) => x.date >= from && x.date <= to) })).filter((w) => w.days.length)
    .map((w) => ({ ...w, actual: w.days.reduce((s, x) => s + x.actual, 0) }));
  if (hist.length < 2) return null;
  const ys = hist.map((w) => w.actual), n = ys.length, mean = ys.reduce((s, v) => s + v, 0) / n, xm = (n - 1) / 2;
  const sxy = ys.reduce((s, v, i) => s + (i - xm) * (v - mean), 0), sxx = ys.reduce((s, _, i) => s + (i - xm) ** 2, 0);
  const slope = sxy / sxx, method = $("rx-method").value;
  // i = weeks after the first history week
  const predict = (i) => (method === "pattern" ? ys[((i - n) % n + n) % n] : Math.max(0, method === "linear" ? mean + slope * (i - xm) : mean));
  const weekIndex = (monday) => Math.round((mondayOf(monday) - hist[0].monday) / (7 * DAY));
  return { all, from, to, hist, ys, n, mean, xm, slope, method, predict, weekIndex };
}

// Projected weekly amounts for the 13 weeks after the history, plus the history itself.
function renderRetailExpand() {
  const chart = $("rx-chart");
  const empty = (msg) => { chart.innerHTML = `<p class="muted" style="padding:16px 4px">${msg}</p>`; $("rx-chips").innerHTML = ""; $("rx-table").innerHTML = ""; };
  if (!state.cf.rows.length) return empty("Upload the Cash Flow Report to see the trend.");
  const M = retailModel();
  if (!M) return empty("Choose a valid history range with at least 2 weeks of Actual Retail Collection data.");
  const { all, from, to, hist, ys, n, mean, slope, method } = M;
  const start = mondayOf(addDays(hist[hist.length - 1].monday, 7));          // first Monday after the history
  const byKey = new Map(all.map((w) => [w.key, w]));
  const proj = Array.from({ length: 13 }, (_, k) => {
    const m = addDays(start, k * 7), a = byKey.get(toISO(m));
    return { monday: m, value: M.predict(n + k), actual: a ? a.actual : null };
  });
  const fmtD = (d) => d.toLocaleDateString("en-PH", { month: "short", day: "numeric" });
  $("rx-sub").textContent = `History: ${n} weeks, ${fmtD(from)} – ${fmtD(to)}, ${to.getFullYear()} → projected ${fmtD(start)} – ${fmtD(addDays(start, 12 * 7 + 4))}, ${addDays(start, 12 * 7 + 4).getFullYear()} (13 weeks, Mon–Fri).`;
  const sgn = (v) => (v >= 0 ? "+" : "") + fmtNum(v);
  $("rx-chips").innerHTML = [["History total", fmtNum(ys.reduce((s, v) => s + v, 0))], ["Average per week", fmtNum(mean)],
    [method === "pattern" ? "Pattern repeats every" : "Trend (change per week)", method === "pattern" ? `${n} weeks` : sgn(slope)], ["Lowest projected week", fmtNum(Math.min(...proj.map((p) => p.value)))], ["Highest projected week", fmtNum(Math.max(...proj.map((p) => p.value)))],
    ["13-week projected total", fmtNum(proj.reduce((s, p) => s + p.value, 0))]].map(([k, v]) => `<div class="bd-chip"><span>${k}</span><b>${v}</b></div>`).join("");

  // chart: history bars (dark) then 13 projected bars (yellow); dark tick = actual already reported for a projected week
  const items = [...hist.map((w) => ({ m: w.monday, v: w.actual, kind: "h" })), ...proj.map((p) => ({ m: p.monday, v: p.value, kind: "p", a: p.actual }))];
  const W = 960, H = 280, L = 60, R = 16, T = 24, B = 46, max = Math.max(...items.map((i) => Math.max(i.v, i.a || 0))) * 1.12 || 1;
  const bw = (W - L - R) / items.length, y = (v) => T + (H - T - B) * (1 - v / max);
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Weekly retail trend">`;
  for (let t = 0; t <= 4; t++) { const v = (max * t) / 4; svg += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#e5e8ee"/><text x="${L - 8}" y="${y(v) + 4}" text-anchor="end" class="ax">${compactPeso(v)}</text>`; }
  items.forEach((it, i) => {
    const x = L + i * bw + bw * 0.16, w = bw * 0.68, yy = y(it.v), lab = `${it.m.getMonth() + 1}/${it.m.getDate()}`;
    svg += `<rect x="${x}" y="${yy}" width="${w}" height="${H - B - yy}" rx="4" fill="${it.kind === "h" ? "#1d2330" : "#ffd100"}"><title>Week of ${lab}: ${fmtNum(it.v)}${it.kind === "p" ? " (projected)" : ""}</title></rect>`;
    if (it.a != null) svg += `<line x1="${x - 2}" x2="${x + w + 2}" y1="${y(it.a)}" y2="${y(it.a)}" stroke="#c0392b" stroke-width="3"><title>Actual so far: ${fmtNum(it.a)}</title></line>`;
    svg += `<text x="${x + w / 2}" y="${yy - 5}" text-anchor="middle" class="vl">${compactPeso(it.v)}</text>`;
    svg += `<text x="${x + w / 2}" y="${H - B + 17}" text-anchor="middle" class="ax">${lab}</text>`;
  });
  // divider between history and projection
  const dx = L + n * bw;
  svg += `<line x1="${dx}" x2="${dx}" y1="${T - 8}" y2="${H - B}" stroke="#9aa1b2" stroke-dasharray="4 4"/><text x="${dx + 6}" y="${T - 10}" class="ax">projection →</text></svg>`;
  chart.innerHTML = `<div class="rt-legend"><span><i style="background:#1d2330"></i>Actual weekly total (history)</span><span><i style="background:#ffd100"></i>Projected (${method === "pattern" ? "history pattern repeated" : method === "linear" ? "straight-line trend" : "flat average"})</span><span><i style="background:#c0392b"></i>Actual already reported for a projected week</span></div>${svg}`;

  $("rx-table").innerHTML = `<table><thead><tr><th>Week (Mon–Fri)</th><th>Basis</th><th>Weekly amount</th><th>Actual already reported</th></tr></thead><tbody>
    ${hist.map((w) => `<tr><td class="l">${fmtD(w.monday)} – ${fmtD(addDays(w.monday, 4))}</td><td class="l">History (actual)</td><td><b>${fmtNum(w.actual)}</b></td><td>—</td></tr>`).join("")}
    ${proj.map((p, k) => `<tr class="in"><td class="l">${fmtD(p.monday)} – ${fmtD(addDays(p.monday, 4))}</td><td class="l">Projected · week ${k + 1} of 13</td><td><b>${fmtNum(p.value)}</b></td><td>${p.actual != null ? fmtNum(p.actual) : "—"}</td></tr>`).join("")}
    <tr class="grand"><td class="l" colspan="2">13-week projected total</td><td>${fmtNum(proj.reduce((s, p) => s + p.value, 0))}</td><td></td></tr></tbody></table>`;
}

function renderRetail() {
  renderRetailExpand();
  const chart = $("rt-chart");
  if (!state.cf.rows.length) {
    chart.innerHTML = '<p class="muted" style="padding:16px 4px">No Cash Flow Report yet. Click <b>Upload</b> and add it in the Cash Flow Report slot.</p>';
    $("rt-chips").innerHTML = ""; $("rt-days").innerHTML = ""; $("rt-weekly").innerHTML = "";
    return;
  }
  const weeks = retailWeeks();
  if (!weeks.length) { chart.innerHTML = '<p class="muted" style="padding:16px 4px">The report has no Actual Retail Collection amounts.</p>'; return; }
  const sel = $("rt-weeks").value, n = sel === "all" ? weeks.length : Math.min(+sel, weeks.length);
  const win = weeks.slice(-n), prev = weeks.slice(Math.max(0, weeks.length - 2 * n), weeks.length - n);
  const days = win.flatMap((w) => w.days);
  const total = days.reduce((s, x) => s + x.actual, 0), proj = days.reduce((s, x) => s + (x.projected || 0), 0);
  const avg = total / days.length;
  const hi = days.reduce((a, b) => (b.actual > a.actual ? b : a)), lo = days.reduce((a, b) => (b.actual < a.actual ? b : a));
  const dShort = (d) => d.toLocaleDateString("en-PH", { month: "short", day: "numeric" });
  const dName = (d) => d.toLocaleDateString("en-PH", { weekday: "short" });
  const first = win[0].days[0].date, last = win[win.length - 1].days[win[win.length - 1].days.length - 1].date;
  $("rt-sub").textContent = `Source: ${state.cf.file.name} · Summary sheet · "Actual Retail Collection" (column ${state.cf.col}) · ${dShort(first)} – ${dShort(last)}, ${last.getFullYear()}`;

  const chips = [["Total collected", fmtNum(total)], ["Business days", String(days.length)], ["Average per business day", fmtNum(avg)],
    [`Highest day (${dName(hi.date)} ${dShort(hi.date)})`, fmtNum(hi.actual)], [`Lowest day (${dName(lo.date)} ${dShort(lo.date)})`, fmtNum(lo.actual)]];
  if (proj) chips.push(["Your report's projection", fmtNum(proj)], ["Actual vs projection", ((total / proj - 1) * 100 >= 0 ? "+" : "") + ((total / proj - 1) * 100).toFixed(1) + "%"]);
  if (prev.length) {
    const pd = prev.flatMap((w) => w.days), pa = pd.reduce((s, x) => s + x.actual, 0) / pd.length;
    chips.push([`vs previous ${n === 1 ? "week" : n + " weeks"} (avg/day)`, (avg / pa - 1 >= 0 ? "+" : "") + ((avg / pa - 1) * 100).toFixed(1) + "%"]);
  }
  $("rt-chips").innerHTML = chips.map(([k, v]) => `<div class="bd-chip"><span>${k}</span><b>${v}</b></div>`).join("");

  // Chart: daily bars for short windows, weekly bars for longer ones. Yellow tick = the report's own projection.
  const daily = days.length <= 25;
  const items = daily
    ? days.map((x) => ({ label: `${dName(x.date)} ${x.date.getMonth() + 1}/${x.date.getDate()}`, value: x.actual, proj: x.projected }))
    : win.map((w) => ({ label: `${w.monday.getMonth() + 1}/${w.monday.getDate()}`, value: w.actual, proj: w.projected }));
  const W = 920, H = 270, L = 56, R = 16, T = 22, B = 44, max = Math.max(...items.map((i) => Math.max(i.value, i.proj || 0))) * 1.12 || 1;
  const bw = (W - L - R) / items.length, y = (v) => T + (H - T - B) * (1 - v / max);
  let svg = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Retail collections">`;
  for (let t = 0; t <= 4; t++) {
    const v = (max * t) / 4;
    svg += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" stroke="#e5e8ee"/><text x="${L - 8}" y="${y(v) + 4}" text-anchor="end" class="ax">${compactPeso(v)}</text>`;
  }
  items.forEach((it, i) => {
    const x = L + i * bw + bw * 0.18, w = bw * 0.64, yy = y(it.value);
    svg += `<rect x="${x}" y="${yy}" width="${w}" height="${H - B - yy}" rx="4" fill="#1d2330"><title>${it.label}: ${fmtNum(it.value)}</title></rect>`;
    if (it.proj) svg += `<line x1="${x - 3}" x2="${x + w + 3}" y1="${y(it.proj)}" y2="${y(it.proj)}" stroke="#ffd100" stroke-width="3"><title>Projected: ${fmtNum(it.proj)}</title></line>`;
    if (items.length <= 14) svg += `<text x="${x + w / 2}" y="${yy - 6}" text-anchor="middle" class="vl">${compactPeso(it.value)}</text>`;
    if (items.length <= 14 || i % Math.ceil(items.length / 14) === 0) svg += `<text x="${x + w / 2}" y="${H - B + 18}" text-anchor="middle" class="ax">${it.label}</text>`;
  });
  svg += "</svg>";
  chart.innerHTML = `<div class="rt-legend"><span><i style="background:#1d2330"></i>Actual retail collection (${daily ? "per business day" : "per week"})</span>${proj ? '<span><i style="background:#ffd100"></i>Your report\'s projection</span>' : ""}</div>${svg}`;

  // Day table for the window: actual collection per business day, with total and average at the bottom
  $("rt-days").innerHTML = `<table><thead><tr><th>Date</th><th>Day</th><th>Actual retail collection</th></tr></thead><tbody>
    ${days.map((x) => `<tr><td class="l">${x.date.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" })}</td><td class="l">${dName(x.date)}</td><td><b>${fmtNum(x.actual)}</b></td></tr>`).join("")}
    <tr class="grand"><td class="l" colspan="2">Total (${days.length} business days)</td><td>${fmtNum(total)}</td></tr>
    <tr class="grand"><td class="l" colspan="2">Average per business day</td><td>${fmtNum(avg)}</td></tr></tbody></table>`;

  // Weekly totals for the whole report, newest first
  const inWin = new Set(win.map((w) => w.key));
  $("rt-weekly").innerHTML = `<table><thead><tr><th>Week (Mon–Sun)</th><th>Business days</th><th>Total collected</th><th>Average per day</th><th>Projected (report)</th><th>Actual vs projected</th></tr></thead><tbody>
    ${[...weeks].reverse().map((w) => `<tr class="${inWin.has(w.key) ? "in" : ""}"><td class="l">${dShort(w.monday)} – ${dShort(addDays(w.monday, 6))}${inWin.has(w.key) ? " ●" : ""}</td><td>${w.days.length}</td><td><b>${fmtNum(w.actual)}</b></td>
      <td>${fmtNum(w.actual / w.days.length)}</td><td>${w.projected ? fmtNum(w.projected) : "—"}</td><td>${w.projected ? ((w.actual / w.projected - 1) * 100).toFixed(1) + "%" : "—"}</td></tr>`).join("")}</tbody></table>`;
}

/* ============================================================
   Unit Disposals file -> "Proceeds from Disposal" row
   Follows the file's own schedule: the "Disposal proceeds" row of the Weekly Cash Flow sheet (with From / To rows).
   Each column goes to the Monday–Sunday week that holds most of its days (e.g. Nov 1–8 -> week of Nov 2).
   ============================================================ */
function mostDaysMonday(f, t) {
  const days = {};
  for (let d = new Date(f.getFullYear(), f.getMonth(), f.getDate()); d <= t; d = addDays(d, 1)) { const m = toISO(mondayOf(d)); days[m] = (days[m] || 0) + 1; }
  const iso = Object.entries(days).sort((x, y) => y[1] - x[1] || (x[0] < y[0] ? -1 : 1))[0][0];
  return new Date(iso + "T00:00:00");
}
function parseUnitDisposals(wb) {
  const names = [...wb.SheetNames].sort((x, y) => /weekly/i.test(y) - /weekly/i.test(x));   // look at "Weekly Cash Flow" first
  for (const name of names) {
    const g = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    const t = (c, re) => typeof c === "string" && re.test(c.trim());
    const rFrom = g.findIndex((r) => t(r[0], /^from$/i)), rTo = g.findIndex((r) => t(r[0], /^to$/i));
    const rAmt = g.findIndex((r) => t(r[0], /^disposal proceeds$/i)), rUnits = g.findIndex((r) => t(r[0], /^units\b.*cash received/i));
    if (rFrom < 0 || rTo < 0 || rAmt < 0) continue;
    const o = XLSX.utils.decode_range(wb.Sheets[name]["!ref"] || "A1").s, rows = [];
    for (let col = 1; col < g[rFrom].length; col++) {
      const f = parseDate(g[rFrom][col]), to = parseDate(g[rTo][col]);
      if (!f || !to) continue;                                  // the Total column has no dates
      const amount = typeof g[rAmt][col] === "number" ? g[rAmt][col] : 0;
      const units = rUnits >= 0 && typeof g[rUnits][col] === "number" ? g[rUnits][col] : null;
      rows.push({ from: f, to, amount, units, monday: mostDaysMonday(f, to), c: o.c + col });
    }
    if (rows.length) return { rows, sheet: name, amtRow: o.r + rAmt };
  }
  throw new Error('No "Disposal proceeds" row with From / To week rows (Weekly Cash Flow sheet of the Unit Disposals file).');
}
function disposalForecast(wk, start) {
  const out = wk.map(() => 0);
  for (const x of state.ds.rows) {
    const i = Math.round((x.monday - start) / (7 * DAY));
    if (i >= 0 && i < wk.length) out[i] += x.amount;
  }
  return out;
}
function renderDisposalMonths() {
  const box = $("ds-months");
  if (!state.ds.rows.length) { box.textContent = "No Unit Disposals file uploaded yet."; return; }
  const srcLine = state.ds.from === "saved" ? `<p class="muted">From the last saved summary (${escapeHtml(state.ds.savedFrom || "")}). Upload a new Unit Disposals file to replace it.</p>`
    : state.ds.file ? `<p class="muted">Uploaded: ${escapeHtml(state.ds.file.name)}</p>` : "";
  const f = (d) => d.toLocaleDateString("en-PH", { month: "short", day: "numeric" });
  const tot = state.ds.rows.reduce((s, x) => s + x.amount, 0), units = state.ds.rows.reduce((s, x) => s + (x.units || 0), 0);
  box.innerHTML = srcLine + `<table><thead><tr><th>File week</th><th>Units</th><th>Disposal proceeds</th><th>Summary week (Mon–Sun)</th></tr></thead><tbody>${state.ds.rows.map((x) =>
    `<tr><td class="l">${f(x.from)} – ${f(x.to)}</td><td>${x.units == null ? "" : x.units}</td><td><b>${fmtNum(x.amount)}</b></td><td class="l">${x.amount ? f(x.monday) + " – " + f(addDays(x.monday, 6)) : ""}</td></tr>`).join("")}
    <tr class="grand"><td class="l">Total</td><td>${units || ""}</td><td>${fmtNum(tot)}</td><td></td></tr></tbody></table>`;
}

function handleDisposal(file) {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
      const res = parseUnitDisposals(wb), rows = res.rows;
      Object.assign(state.ds, { rows, file, buffer: e.target.result, saved: false, saveErr: null, sheet: res.sheet, from: "file" });
      $("drop-ds").classList.add("loaded");
      await saveOriginal("ds");
      const tot = rows.reduce((s, x) => s + x.amount, 0), units = rows.reduce((s, x) => s + (x.units || 0), 0);
      $("mstat-ds").innerHTML = `<b>${escapeHtml(file.name)}</b><br>${units ? units + " units · " : ""}${fmtNum(tot)} disposal proceeds · sheet "${escapeHtml(res.sheet)}"`;
      renderDisposalMonths();
      saved = false;
      refresh(); renderFileList();
      $("review").classList.remove("hidden"); $("empty").classList.add("hidden"); $("tabs").classList.remove("hidden");
      showTab("summary");
    } catch (err) {
      $("mstat-ds").innerHTML = `<span class="err">Could not read this file: ${escapeHtml(err.message)}</span>`;
    }
  };
  return readInto(reader, file, "ds");
}

/* ============================================================
   Cash Out sources: Loan Payment Schedule, Unreleased Checks, Projected Cash Outflow
   ============================================================ */
function sumDates(byDate, from, to) {
  const a = toISO(from), b = toISO(to);
  let s = 0;
  for (const [d, v] of Object.entries(byDate)) if (d >= a && d <= b) s += v;
  return s;
}
const isTxt = (c, re) => typeof c === "string" && re.test(c.trim());
function gridOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }); }

// Loans: the sheet with "Due date" + "Net amount" (Detail). Net = principal + interest - EWT.
function parseLoans(wb) {
  for (const name of wb.SheetNames) {
    const g = gridOf(wb.Sheets[name]);
    const hi = g.slice(0, 20).findIndex((r) => r.some((c) => isTxt(c, /^due date$/i)) && r.some((c) => isTxt(c, /^net amount$/i)));
    if (hi < 0) continue;
    const iD = g[hi].findIndex((c) => isTxt(c, /^due date$/i)), iA = g[hi].findIndex((c) => isTxt(c, /^net amount$/i));
    const o = XLSX.utils.decode_range(wb.Sheets[name]["!ref"] || "A1").s, lines = [];
    const byDate = {}; let n = 0, total = 0;
    g.slice(hi + 1).forEach((r, j) => {
      const d = parseDate(r[iD]), a = r[iA];
      if (!d || typeof a !== "number") return;
      const k = toISO(d); byDate[k] = (byDate[k] || 0) + a; n++; total += a;
      lines.push({ r: o.r + hi + 1 + j, date: d });
    });
    if (n) return { byDate, n, total, what: "payments", loc: { sheet: name, hdr: o.r + hi, cAmt: o.c + iA, lines } };
  }
  throw new Error('No sheet with "Due date" and "Net amount" headings.');
}
// Unreleased checks: "NET AMOUNT" + "Actual/Projected Date Cleared".
function parseChecks(wb) {
  for (const name of wb.SheetNames) {
    const g = gridOf(wb.Sheets[name]);
    const hi = g.slice(0, 20).findIndex((r) => r.some((c) => isTxt(c, /date cleared/i)) && r.some((c) => isTxt(c, /^net amount$/i)));
    if (hi < 0) continue;
    const iD = g[hi].findIndex((c) => isTxt(c, /date cleared/i)), iA = g[hi].findIndex((c) => isTxt(c, /^net amount$/i));
    const o = XLSX.utils.decode_range(wb.Sheets[name]["!ref"] || "A1").s, lines = [];
    const byDate = {}; let n = 0, total = 0, undated = 0;
    g.slice(hi + 1).forEach((r, j) => {
      const a = r[iA];
      if (typeof a !== "number") return;
      const d = parseDate(r[iD]);
      lines.push({ r: o.r + hi + 1 + j, date: d });
      if (!d) { undated += a; return; }
      const k = toISO(d); byDate[k] = (byDate[k] || 0) + a; n++; total += a;
    });
    if (n) return { byDate, n, total, undated, what: "checks", loc: { sheet: name, hdr: o.r + hi, cAmt: o.c + iA, lines } };
  }
  throw new Error('No sheet with "NET AMOUNT" and "Date Cleared" headings.');
}
// Projected outflow: weekly columns with "From"/"To" dates and a "TOTAL" row. Each column goes to the
// Monday–Sunday week that holds most of its days (e.g. Nov 1–8 -> week of Nov 2), dated on that Monday.
function parseProjected(wb) {
  for (const name of wb.SheetNames) {
    const g = gridOf(wb.Sheets[name]);
    const rFrom = g.findIndex((r) => isTxt(r[0], /^from$/i)), rTo = g.findIndex((r) => isTxt(r[0], /^to$/i));
    if (rFrom < 0 || rTo < 0) continue;
    const rTot = g.findIndex((r, i) => i > rTo && isTxt(r[0], /^total$/i));
    const o = XLSX.utils.decode_range(wb.Sheets[name]["!ref"] || "A1").s, cols = [];
    const byDate = {}; let n = 0, total = 0;
    for (let col = 1; col < g[rFrom].length; col++) {
      const f = parseDate(g[rFrom][col]), t = parseDate(g[rTo][col]);
      if (!f || !t) continue;                                   // month-total / grand-total columns have no dates
      let amt = 0;
      if (rTot >= 0 && typeof g[rTot][col] === "number") amt = g[rTot][col];
      else for (let r = rTo + 1; r < g.length; r++) if (typeof g[r][col] === "number") amt += g[r][col];
      const days = {};
      for (let d = new Date(f.getFullYear(), f.getMonth(), f.getDate()); d <= t; d = addDays(d, 1)) { const m = toISO(mondayOf(d)); days[m] = (days[m] || 0) + 1; }
      const target = Object.entries(days).sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1))[0][0];
      byDate[target] = (byDate[target] || 0) + amt; n++; total += amt;
      cols.push({ c: o.c + col, monday: new Date(target + "T00:00:00"), amt });
    }
    if (n) return { byDate, n, total, what: "weekly columns", loc: { sheet: name, totRow: rTot >= 0 ? o.r + rTot : -1, cols } };
  }
  throw new Error('No sheet with "From" / "To" week rows.');
}

const CO_META = { ln: { label: "Loan payments", parse: parseLoans }, uc: { label: "Unreleased checks", parse: parseChecks }, po: { label: "Projected payments", parse: parseProjected } };
function coInfo(res) {
  const ds = Object.keys(res.byDate).sort();
  const f = (s) => new Date(s + "T00:00:00").toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" });
  return `${res.n.toLocaleString("en-PH")} ${res.what} · ${fmtNum(res.total)} · ${ds.length ? f(ds[0]) + " – " + f(ds[ds.length - 1]) : ""}`;
}
function handleCashOut(key, file) {
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
      const res = CO_META[key].parse(wb);
      Object.assign(state[key], { byDate: res.byDate, file, buffer: e.target.result, saved: false, saveErr: null, info: coInfo(res), from: "file", rows: [1] });
      $("drop-" + key).classList.add("loaded");
      await saveOriginal(key);
      $("mstat-" + key).innerHTML = `<b>${escapeHtml(file.name)}</b><br>${state[key].info}` + (res.undated ? `<br><span class="warn">${fmtNum(res.undated)} without a clearing date (not placed)</span>` : "");
      renderCashOutSources();
      saved = false;
      refresh(); renderFileList();
      $("review").classList.remove("hidden"); $("empty").classList.add("hidden"); $("tabs").classList.remove("hidden");
      showTab("summary");
    } catch (err) {
      $("mstat-" + key).innerHTML = `<span class="err">Could not read this file: ${escapeHtml(err.message)}</span>`;
    }
  };
  return readInto(reader, file, key);
}
function renderCashOutSources() {
  const rows = Object.entries(CO_META).map(([k, m]) => {
    const s = state[k], has = Object.keys(s.byDate).length > 0;
    const src = !has ? '<span class="muted">Not loaded</span>' : s.from === "file" ? `Uploaded: ${escapeHtml(s.file.name)}` : `From the last saved summary${s.savedFrom ? " (" + escapeHtml(s.savedFrom) + ")" : ""}`;
    return `<tr><td class="l"><b>${m.label}</b></td><td class="l">${src}</td><td class="l">${has ? s.info : ""}</td>
      <td>${has ? `<button type="button" class="btn ghost small" data-coclear="${k}">Remove</button>` : ""}</td></tr>`;
  }).join("");
  $("co-sources").innerHTML = `<table><thead><tr><th>Cash Out row</th><th>Source</th><th>Contents</th><th></th></tr></thead><tbody>${rows}</tbody></table>`;
  $("co-sources").querySelectorAll("[data-coclear]").forEach((b) => (b.onclick = () => {
    const k = b.dataset.coclear;
    Object.assign(state[k], { byDate: {}, file: null, buffer: null, info: "", from: "", rows: [], dirty: true, removed: true });
    $("drop-" + k).classList.remove("loaded"); $("mstat-" + k).textContent = "No file chosen";
    renderCashOutSources(); saved = false; if (report) refresh();
  }));
}

// Copy the uploaded file into the "Uploaded Files" folder.
let uploadBatch = null;      // name of the dated folder inside "Uploaded Files" for the current Upload pop-up session
function batchFolderName(d = new Date()) {
  const p = (n) => String(n).padStart(2, "0"), h12 = d.getHours() % 12 || 12;
  return `${p(d.getMonth() + 1)}-${p(d.getDate())}-${d.getFullYear()} - ${h12}.${p(d.getMinutes())}${d.getHours() < 12 ? "am" : "pm"}`;
}

async function saveOriginal(kind) {
  const s = state[kind];
  if (!s.file) return;
  if ($("opt-autosave").checked && (await folderReady())) {
    try {
      if (!uploadBatch) uploadBatch = batchFolderName();
      const path = [UPLOAD_DIR, uploadBatch];
      const name = await uniqueName(path, s.file.name);          // original file name, kept as uploaded
      await writeTo(path, name, s.buffer);
      s.saved = true; s.savedAs = name; s.savedIn = `${UPLOAD_DIR}/${uploadBatch}`;
    } catch (e) { s.saved = false; s.saveErr = e.message; }
  }
  renderFileInfo(kind);
}
function renderFileInfo(kind) {
  const s = state[kind];
  if (!s.file) return;
  const n = s.rows.length;
  let line;
  if (s.saved) line = `<div class="ok">✓ Copy saved to "${s.savedIn || UPLOAD_DIR}"</div>`;
  else if (s.saveErr) line = `<div class="err">Could not save a copy: ${escapeHtml(s.saveErr)}</div>`;
  else if (!$("opt-autosave").checked) line = `<div class="warn">Auto-save is off — file was read but not copied to "${UPLOAD_DIR}"</div>`;
  else line = `<div class="warn">Not saved yet — connect a folder (bottom of the sidebar) to keep a copy in "${UPLOAD_DIR}"</div>`;
  if ($("info-" + kind)) $("info-" + kind).innerHTML = `<div class="name">${escapeHtml(s.file.name)}</div><div class="muted">${n} rows read from sheet "${escapeHtml(s.sheet || "")}"</div>${line}`;
  const ok = s.saved ? `<span class="ok">✓ Copy saved to "${s.savedIn || UPLOAD_DIR}"</span>` : s.saveErr ? `<span class="err">Could not save a copy</span>` : !$("opt-autosave").checked ? `<span class="warn">Auto-save is off — not copied</span>` : `<span class="warn">Not saved yet — connect a folder first</span>`;
  $("mstat-" + kind).innerHTML = `<b>${escapeHtml(s.file.name)}</b> · ${n} rows<br>${ok}`;
}

function findHeader(grid) {
  let best = 0, bestScore = -1;
  grid.slice(0, 30).forEach((r, i) => {
    const score = r.reduce((n, c) => n + (typeof c === "string" && /due|balance|amount|total|customer|vendor|supplier|invoice|name|current|days/i.test(c) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; best = i; }
  });
  return { headerIdx: best, headers: (grid[best] || []).map((h, i) => (h == null || h === "" ? `Column ${i + 1}` : String(h).trim())) };
}
function guessMapping(headers) {
  const find = (...pats) => { for (const p of pats) { const i = headers.findIndex((h) => p.test(h)); if (i >= 0) return i; } return -1; };
  const map = {};
  map.due = find(/^due\s*date$/i, /due\s*date/i, /^due$/i);
  map.party = find(/^company$/i, /customer|client|vendor|supplier|payee|payor|company/i, /name|account/i);
  map.amount = find(/^invoice\s*amount$/i, /outstanding|balance/i, /amount\s*due|open/i, /^total/i, /amount/i);
  map.type = find(/invoice\s*type/i, /^type$/i, /service/i);
  map.vat = find(/vat\s*type/i, /^vat$/i);
  map.inv = find(/invoice\s*(no|num|#)/i, /^invoice$/i, /^si\b/i);
  map.proj = find(/projected\s*(collection\s*)?date/i);
  map.rem = find(/^remarks?$/i, /remark/i);
  if (map.party < 0) map.party = 0;
  return map;
}
function renderMapping(kind) {
  const s = state[kind], box = $("map-" + kind);
  box.innerHTML = "";
  for (const f of Object.keys(FIELD_LABELS)) {
    if (kind === "ap" && AR_ONLY_FIELDS.includes(f)) continue;
    const label = document.createElement("label");
    label.textContent = FIELD_LABELS[f];
    const sel = document.createElement("select");
    sel.innerHTML = `<option value="-1">— none —</option>` + s.headers.map((h, i) => `<option value="${i}" ${s.map[f] === i ? "selected" : ""}>${escapeHtml(h)}</option>`).join("");
    sel.onchange = () => { s.map[f] = parseInt(sel.value, 10); refresh(); };
    label.appendChild(sel); box.appendChild(label);
  }
  $("mapbox-" + kind).classList.remove("hidden");
}

/* ============================================================
   Report logic (kept simple for now — logic will be refined next)
   ============================================================ */
// Projected Date column: a date, "-" (no collection expected) or blank.
function readProjected(v) {
  if (v == null || v === "") return { proj: null, projDash: false };
  if (typeof v === "string" && /^\s*[-–—]+\s*$/.test(v)) return { proj: null, projDash: true };
  const d = parseDate(v);
  return { proj: d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null, projDash: false };
}
function remarkText(v) {
  if (v == null) return "";
  if (v instanceof Date) return v.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" });
  return String(v).trim();
}
// Remark groups — used only to show / filter remarks (they do not change the forecast).
const REMARK_GROUPS = [
  { key: "paid", label: "Paid / collected", re: /\b(paid|collected|prepaid)\b/i },
  { key: "cancel", label: "Cancelled", re: /cancel/i },
  { key: "offset", label: "Offset", re: /offset/i },
  { key: "cm", label: "For CM", re: /\bcm\b|credit\s*memo/i },
  { key: "check", label: "Checking / revision / follow up", re: /check|revis|follow/i },
];
function remarkGroup(s) { if (!s) return "none"; const g = REMARK_GROUPS.find((x) => x.re.test(s)); return g ? g.key : "other"; }
const REMARK_FILTERS = [["all", "All"], ["any", "With a remark"], ...REMARK_GROUPS.map((g) => [g.key, g.label]), ["other", "Other remarks"], ["none", "No remark"]];
function remarkMatch(x, f) { const g = remarkGroup(x.rem); return f === "all" || (f === "any" ? g !== "none" : g === f); }

// true when the Projected Date cell of this AR line was changed in the file editor
function projWasEdited(kind, r) {
  const s = state[kind], m = s.map;
  if (!m || m.proj < 0 || r.__r == null) return false;
  const ed = s.edits && s.edits[s.sheet];
  return !!ed && XLSX.utils.encode_cell({ r: r.__r, c: (s.c0 || 0) + m.proj }) in ed;
}

function extractItems(kind) {
  const { rows, map } = state[kind];
  const items = []; let skipped = 0, undated = 0;
  if (map.amount < 0) return { items, skipped, undated };
  for (const r of rows) {
    const amount = parseNumber(r[map.amount]);
    const party = map.party >= 0 ? r[map.party] : "";
    if (isNaN(amount) || amount === 0) { skipped++; continue; }
    // Skip footer rows only when a cell is exactly "Total" / "Grand Total" (a customer like "Total (Philippines) Corp." is real).
    if (r.some((c) => typeof c === "string" && /^\s*(grand\s*)?(sub\s*)?total\s*:?\s*$/i.test(c))) { skipped++; continue; }
    const due = map.due >= 0 ? parseDate(r[map.due]) : null;
    if (!due) undated++;
    items.push({
      party: party == null || party === "" ? "(unnamed)" : String(party).trim(),
      due, amount, row: r.__r,
      type: map.type >= 0 && r[map.type] != null ? String(r[map.type]) : "",
      vat: map.vat >= 0 && r[map.vat] != null ? String(r[map.vat]) : "",
      inv: map.inv >= 0 && r[map.inv] != null ? String(r[map.inv]) : "",
      ...readProjected(map.proj >= 0 ? r[map.proj] : null),
      projEdited: projWasEdited(kind, r),
      rem: map.rem >= 0 ? remarkText(r[map.rem]) : "",
    });
  }
  return { items, skipped, undated };
}

/* ---- AR collection rules (see Setup > AR collection rules) ---- */
function numVal(id, d) { const v = parseFloat($(id).value); return isNaN(v) ? d : v; }
function weekVal(id) { const v = parseInt($(id).value, 10); return isNaN(v) || v < 1 ? null : v; }
function arRules() {
  const asOf = parseDate($("opt-asof").value) || new Date();
  return {
    asOf: new Date(asOf.getFullYear(), asOf.getMonth(), asOf.getDate()),
    allow: { allin: numVal("al-allin", 7), other: numVal("al-other", 14), fixed: numVal("al-fixed", 7) },
    buckets: [{ max: 30, wk: weekVal("bk-30") }, { max: 60, wk: weekVal("bk-60") }, { max: 90, wk: weekVal("bk-90") }, { max: 120, wk: weekVal("bk-120") }, { max: Infinity, wk: weekVal("bk-max") }],
    ewt: { on: $("ewt-on").checked, rental: numVal("ewt-rental", 5) / 100, other: numVal("ewt-other", 2) / 100, vat: numVal("ewt-vat", 12) / 100 },
    exempt: new Set($("ewt-exempt").value.split("\n").map((s) => s.trim().toLowerCase()).filter(Boolean)),
  };
}
const normType = (s) => String(s || "").toLowerCase().replace(/[\u2013\u2014]/g, "-").replace(/\s+/g, " ").trim();
// Fixed Monthly listed first = advance invoicing; "All In - ST" anywhere = +7; otherwise +14 (matches the timing file).
function typeKind(type) {
  const parts = normType(type).split(",").map((p) => p.trim()).filter(Boolean);
  const first = parts[0] || "";
  const kind = first === "fixed monthly" ? "fixed" : parts.includes("all in - st") ? "allin" : "other";
  const rental = first === "all in - st" || first === "fixed monthly"; // 5% EWT; every other first type = 2%
  return { kind, rental };
}

// Turns AR invoices into expected cash per forecast week.
function computeAR(items, start, weeks, rules) {
  const weekly = Array(weeks).fill(0);
  const out = { weekly, rule: 0, later: 0, nodue: 0, dash: 0, gross: 0, ewt: 0, cash: 0, count: 0, byParty: new Map(), mix: { allin: 0, other: 0, fixed: 0 }, lines: [] };
  const BK_LABEL = ["1–30", "31–60", "61–90", "91–120", "Over 120"];
  const ALLOW_LABEL = { allin: "All In – ST", other: "other charges", fixed: "Fixed Monthly" };
  const WEEK = 7 * DAY;
  for (const it of items) {
    const tk = typeKind(it.type);
    let ewt = 0, ewtRate = 0;
    const nonVat = normType(it.vat) === "non vat";
    if (rules.ewt.on && !rules.exempt.has(it.party.toLowerCase())) {
      const base = nonVat ? it.amount : it.amount / (1 + rules.ewt.vat);
      ewtRate = tk.rental ? rules.ewt.rental : rules.ewt.other;
      ewt = base * ewtRate;
    }
    const cash = it.amount - ewt;
    out.count++; out.gross += it.amount; out.ewt += ewt; out.cash += cash; out.mix[tk.kind] += cash;

    // Where does it land? week index 0..weeks-1, or an excluded bucket.
    let slot, week = null, why = "", dpd = null;
    if (!it.due) { slot = "nodue"; why = "No due date — not in the forecast"; }
    else {
      const due = new Date(it.due.getFullYear(), it.due.getMonth(), it.due.getDate());
      dpd = Math.round((rules.asOf - due) / DAY);
      const allow = rules.allow[tk.kind], exp = addDays(due, allow);
      if (it.projEdited && (it.proj || it.projDash)) {   // Projected Date changed in the system: always used
        if (it.projDash) { slot = "dash"; why = `Projected Date set to "-" in the system — not in the forecast`; }
        else { const raw = Math.floor((it.proj - start) / WEEK) + 1; week = Math.max(1, raw);
          why = `Projected Date edited in the system → ${fmtDate(it.proj)}, ${it.proj.getFullYear()}` + (raw < 1 ? " — before Week 1, so Week 1" : ""); }
      } else if (exp >= rules.asOf) {                    // due date + allowance not passed yet
        const raw = Math.floor((exp - start) / WEEK) + 1;
        week = Math.max(1, raw);
        why = (dpd <= 0 ? "Not yet due" : `${dpd} days past due, still within the allowance`) +
          `: due date + ${allow} days allowance (${ALLOW_LABEL[tk.kind]}) = ${fmtDate(exp)}` + (raw < 1 ? " — before Week 1, so Week 1" : "");
      } else if (it.projDash) {                          // late, and the AR team marked "-": no collection expected
        slot = "dash"; why = `Late (due date + ${allow} days = ${fmtDate(exp)} has passed); Projected Date is "-" — not in the forecast`;
      } else if (it.proj) {                              // late: use the AR team's Projected Date
        const raw = Math.floor((it.proj - start) / WEEK) + 1;
        week = Math.max(1, raw);
        why = `Late (due date + ${allow} days = ${fmtDate(exp)} has passed) → Projected Date ${fmtDate(it.proj)}, ${it.proj.getFullYear()}` + (raw < 1 ? " — before Week 1, so Week 1" : "");
      } else {                                           // late, no Projected Date: overdue bucket rule
        const bi = rules.buckets.findIndex((x) => dpd <= x.max), b = rules.buckets[bi];
        if (!b || b.wk == null) { slot = "rule"; why = `Late, no Projected Date: ${dpd} days overdue (${BK_LABEL[bi]}) — not in the forecast`; }
        else { week = b.wk; why = `Late, no Projected Date: ${dpd} days overdue (${BK_LABEL[bi]} bucket) → Week ${week}`; }
      }
      if (week != null) { slot = week > weeks ? "later" : week - 1; if (slot === "later") why += ` — after Week ${weeks}, not in the forecast`; }
    }
    if (typeof slot === "number") weekly[slot] += cash; else out[slot] += cash;
    out.lines.push({ party: it.party, inv: it.inv, type: it.type, due: it.due, dpd, amount: it.amount, ewt, cash, slot, why, proj: it.proj, projDash: it.projDash, rem: it.rem, row: it.row, ewtRate, nonVat });

    if (!out.byParty.has(it.party)) out.byParty.set(it.party, { weekly: Array(weeks).fill(0), rule: 0, later: 0, nodue: 0, dash: 0 });
    const p = out.byParty.get(it.party);
    if (typeof slot === "number") p.weekly[slot] += cash; else p[slot] += cash;
  }
  return out;
}

const wkIso = (i) => toISO(report.wk[i].start);     // calendar key of forecast column i
function resetEdits() { overrides = {}; customRows.forEach((r) => (r.vals = {})); }
function hasEdits() { return Object.keys(overrides).length > 0 || customRows.some((r) => Object.keys(r.vals).length); }

function acctTotal() { return accounts.reduce((s, a) => s + (a.bal || 0), 0); }

// Raw (unedited) numbers from the uploaded files
function buildBase() {
  const start = mondayOf(new Date(...$("opt-start").value.split("-").map((v, i) => (i === 1 ? v - 1 : +v))));
  const weeks = Math.max(1, Math.min(52, parseInt($("opt-weeks").value, 10) || 13));
  const rules = arRules();
  const ar = extractItems("ar"), ap = extractItems("ap");
  const arRes = computeAR(ar.items, start, weeks, rules);
  const wk = [];
  for (let i = 0; i < weeks; i++) { const s = addDays(start, i * 7); wk.push({ n: i + 1, start: s, end: addDays(s, 6) }); }
  // AR Corporate comes from the AR Aging. AP and the other rows stay empty (logic not built yet).
  const base = {};
  [...CASH_IN, ...CASH_OUT].forEach(([k]) => (base[k] = wk.map(() => 0)));
  base.arCorp = arRes.weekly.slice();
  base.retail = retailForecast(wk);
  base.disposal = disposalForecast(wk, start);
  for (const [k, src] of Object.entries(CO_SRC)) base[k] = wk.map((w) => -sumDates(state[src].byDate, w.start, w.end));
  return { start, weeks, wk, ar, ap, arRes, rules, base };
}

// Apply manual edits and derive totals / balances
function compute(r) {
  const n = r.weeks, v = {};
  for (const k of Object.keys(r.base)) v[k] = r.base[k].map((x, i) => { const o = overrides[k], d = toISO(r.wk[i].start); return o && o[d] !== undefined ? (CO_SRC[k] ? -Math.abs(o[d]) : o[d]) : x; });
  customRows.forEach((row) => { v["c:" + row.id] = Array.from({ length: n }, (_, i) => row.sign * Math.abs(row.vals[toISO(r.wk[i].start)] || 0)); });
  v.cashIn = [], v.cashOut = [], v.net = [], v.begin = [], v.end = [];
  let run = acctTotal();
  for (let i = 0; i < n; i++) {
    v.cashIn.push(rowOrder.in.reduce((s, k) => s + v[k][i], 0));
    v.cashOut.push(rowOrder.out.reduce((s, k) => s + v[k][i], 0));
    v.net.push(v.cashIn[i] + v.cashOut[i]);
    v.begin.push(run); run += v.net[i]; v.end.push(run);
  }
  return { v, opening: acctTotal(), closing: run };
}

// Row lists follow the template's Summary sheet. Cash In rows were given by the user; Cash Out rows are still pending.
const CASH_IN = [
  ["arCorp", "AR Corporate"],
  ["retail", "Retail (Credit Card)"],
  ["loan", "Loan Proceeds"],
  ["disposal", "Proceeds from Disposal"],
  ["affiliate", "Advances from Affiliates"],
];
const CASH_OUT = [["loans", "Loan payments"], ["checks", "Unreleased checks"], ["projout", "Projected payments"]];
const CO_SRC = { loans: "ln", checks: "uc", projout: "po" };   // Cash Out row -> source in state
// Display order of the rows inside each section. Template rows are fixed; rows added by the user ("c:<id>")
// can be dragged to any position, in either section. Amounts keep their own sign (Deduct -, Add +).
const rowOrder = { in: CASH_IN.map(([k]) => k), out: CASH_OUT.map(([k]) => k) };
function rowDefFor(k, side) {
  if (k.startsWith("c:")) {
    const r = customRows.find((x) => "c:" + x.id === k);
    return { k, id: r.id, label: r.name, cls: side + " indent custom", custom: true };
  }
  const f = CASH_IN.find(([kk]) => kk === k) || CASH_OUT.find(([kk]) => kk === k);
  return { k, label: f[1], cls: side + " indent", edit: true };
}
function rowsDef() {
  return [
    { k: "begin", label: "Beginning Balance", cls: "begin", total: "opening" },
    { section: "Cash In", drop: "in:start" },
    ...rowOrder.in.map((k) => rowDefFor(k, "in")),
    { k: "cashIn", label: "Total Cash In", cls: "total", drop: "in:end" },
    { section: "Cash Out", drop: "out:start" },
    ...(rowOrder.out.length ? rowOrder.out.map((k) => rowDefFor(k, "out")) : [{ note: "Cash Out line items to be defined", drop: "out:end" }]),
    { k: "cashOut", label: "Total Cash Out", cls: "total", drop: "out:end" },
    { k: "net", label: "Net Cash Flow", cls: "total" },
    { k: "end", label: "Ending Cash Balance", cls: "end", total: "closing" },
  ];
}
function rowTotal(row, c) { return row.total ? c[row.total] : c.v[row.k].reduce((a, b) => a + b, 0); }

/* ============================================================
   Rendering
   ============================================================ */
// The Summary is shown as soon as ANY source file is uploaded (AR Aging is not required first).
function hasUploads() {
  return !!(state.ar.rows.length || state.ap.rows.length || ["cf", "ds", "ln", "uc", "po"].some((k) => state[k].buffer)
    || state.ds.rows.length || ["ln", "uc", "po"].some((k) => Object.keys(state[k].byDate).length));   // also schedules kept from the last summary
}
function refresh() {
  if (!$("opt-start").value || !hasUploads()) return;
  report = buildBase();
  $("review").classList.remove("hidden");
  $("empty").classList.add("hidden");
  $("tabs").classList.remove("hidden");
  renderSummary();
  renderDetail();
  renderPastDue();
  renderNotes();
  $("bal-total").textContent = fmtNum(acctTotal());
  updateSaveState();
  if (typeof renderFileList === "function") renderFileList();
  if (typeof renderOverrideWarning === "function" && fe.k && !$("fe-modal").classList.contains("hidden")) renderOverrideWarning();
}

// Top-bar subtitle and the key figures shown in the summary cards
function renderCards(c) {
  const r = report, sum = (a) => a.reduce((x, y) => x + y, 0);
  const money = (el, v) => { $(el).textContent = fmtNum(v) === "-" ? "0.00" : fmtNum(v); $(el).classList.toggle("neg", v < 0); };
  money("card-begin", c.opening);
  money("card-in", sum(c.v.cashIn));
  money("card-out", sum(c.v.cashOut));
  money("card-net", sum(c.v.net));
  money("card-end", c.closing);
  $("card-in-sub").textContent = `${r.weeks} weeks`;
  $("card-out-sub").textContent = `${r.weeks} weeks`;
  $("card-net-sub").textContent = sum(c.v.net) >= 0 ? "Cash increases" : "Cash decreases";
  let low = 0;
  c.v.end.forEach((v, i) => { if (v < c.v.end[low]) low = i; });
  $("card-end-sub").textContent = `Lowest: ${fmtNum(c.v.end[low]) === "-" ? "0.00" : fmtNum(c.v.end[low])} (Week ${low + 1})`;
}

function renderSummary() {
  const r = report, c = compute(r), wk = r.wk, ncol = wk.length + 2;
  $("rpt-company").textContent = $("opt-company").value;
  $("rpt-weeks").textContent = `${r.weeks} Week Cash Flow Forecast`;
  const last = r.wk[r.wk.length - 1];
  const range = `${fmtDate(r.wk[0].start)} – ${fmtDate(last.end)}, ${last.end.getFullYear()}`;
  $("rpt-range").textContent = range;
  $("top-sub").textContent = `${$("opt-company").value} · ${range}`;
  renderCards(c);
  const head = `<tr><th>Week #</th>${wk.map((w) => `<th>${w.n}</th>`).join("")}<th>Total</th></tr>
    <tr class="dates"><th>Monday</th>${wk.map((w) => `<th>${fmtDate(w.start)}</th>`).join("")}<th></th></tr>
    <tr class="dates"><th>Sunday</th>${wk.map((w) => `<th>${fmtDate(w.end)}</th>`).join("")}<th></th></tr>`;
  const body = rowsDef().map((row) => {
    if (row.section) return `<tr class="section" data-drop="${row.drop}"><td colspan="${ncol}">${row.section}</td></tr>`;
    if (row.note) return `<tr class="pending" data-drop="${row.drop}"><td colspan="${ncol}">${row.note}</td></tr>`;
    if (row.custom) {
      const cr = customRows.find((x) => x.id === row.id);
      const labelCell = editing
        ? `<td class="indent custom-label"><span class="cr-handle" title="Drag to move this row">⋮⋮</span><input type="text" class="cr-name" data-id="${cr.id}" value="${escapeHtml(cr.name)}" placeholder="Row name (e.g. EWT adjustment)">
             <select class="cr-sign" data-id="${cr.id}"><option value="-1" ${cr.sign < 0 ? "selected" : ""}>Deduct (−)</option><option value="1" ${cr.sign > 0 ? "selected" : ""}>Add (+)</option></select>
             <button type="button" class="cr-del" data-id="${cr.id}" title="Remove this row">✕</button></td>`
        : `<td class="indent"><span class="cr-handle" title="Drag to move this row">⋮⋮</span> ${escapeHtml(cr.name || "(unnamed row)")}${cr.sign < 0 ? ' <span class="muted">(deduction)</span>' : ""}</td>`;
      const cells = c.v[row.k].map((val, i) => editing
        ? `<td><input type="number" step="0.01" min="0" value="${cr.vals[wkIso(i)] ?? ""}" data-custom="${cr.id}" data-i="${i}" placeholder="0.00"></td>`
        : `<td data-cell="${row.k}:${i}" class="${val < 0 ? "neg" : ""}"${cr.vals[wkIso(i)] != null ? ` title="${escapeHtml(typedTip(row.k, wkIso(i)))}"` : ""}>${fmtNum(val)}</td>`).join("");
      const t = rowTotal(row, c);
      return `<tr class="${row.cls}" data-key="${row.k}">${labelCell}${cells}<td data-cell="${row.k}:t" class="${t < 0 ? "neg" : ""}">${fmtNum(t)}</td></tr>`;
    }
    const cells = c.v[row.k].map((val, i) => {
      if (row.edit && editing) {
        const edited = overrides[row.k] && overrides[row.k][wkIso(i)] !== undefined;
        return `<td class="${edited ? "edited" : ""}"><input type="number" step="0.01" value="${Math.round(val * 100) / 100}" data-row="${row.k}" data-i="${i}"></td>`;
      }
      const edited = row.edit && overrides[row.k] && overrides[row.k][wkIso(i)] !== undefined;
      const click = row.k === "arCorp" ? " clickable" : "";
      const tip = edited ? typedTip(row.k, wkIso(i)) : click ? "Click to see the invoices behind this amount" : "";
      return `<td data-cell="${row.k}:${i}" ${tip ? `title="${escapeHtml(tip)}"` : ""} class="${val < 0 ? "neg" : ""} ${edited ? "edited" : ""}${click}">${fmtNum(val)}</td>`;
    }).join("");
    const t = rowTotal(row, c);
    return `<tr class="${row.cls}" ${row.drop ? `data-drop="${row.drop}"` : `data-key="${row.k}"`}><td class="${row.cls.includes("indent") ? "indent" : ""}">${row.label}</td>${cells}<td data-cell="${row.k}:t" ${row.k === "arCorp" ? 'title="Click to see all invoices"' : ""} class="${t < 0 ? "neg" : ""}${row.k === "arCorp" ? " clickable" : ""}">${fmtNum(t)}</td></tr>`;
  }).join("");
  const wrap = $("table-summary");
  wrap.classList.toggle("editing", editing);
  wrap.innerHTML = `<table><thead>${head}</thead><tbody>${body}</tbody></table>`;
  wrap.querySelectorAll("input[data-row]").forEach((inp) => (inp.oninput = onCellEdit));
  wrap.querySelectorAll("input[data-custom]").forEach((inp) => (inp.oninput = () => {
    const cr = customRows.find((x) => x.id === +inp.dataset.custom), v = parseFloat(inp.value);
    const dk = wkIso(+inp.dataset.i); if (isNaN(v)) delete cr.vals[dk]; else cr.vals[dk] = v;
    saved = false; refreshDerived();
  }));
  wrap.querySelectorAll(".cr-name").forEach((inp) => (inp.oninput = () => { customRows.find((x) => x.id === +inp.dataset.id).name = inp.value; saved = false; updateSaveState(); }));
  wrap.querySelectorAll(".cr-sign").forEach((sel) => (sel.onchange = () => { customRows.find((x) => x.id === +sel.dataset.id).sign = +sel.value; saved = false; renderSummary(); updateSaveState(); }));
  wrap.querySelectorAll(".cr-del").forEach((btn) => (btn.onclick = () => { const dk = "c:" + btn.dataset.id; customRows = customRows.filter((x) => x.id !== +btn.dataset.id); rowOrder.in = rowOrder.in.filter((k) => k !== dk); rowOrder.out = rowOrder.out.filter((k) => k !== dk); saved = false; renderSummary(); updateSaveState(); }));
  $("btn-reset").classList.toggle("hidden", !hasEdits());
  wireRowDrag(wrap);
  wrap.onclick = (e) => {
    const td = e.target.closest("td.clickable");
    if (!td) return;
    const idx = td.dataset.cell.split(":")[1];
    openBreakdown(idx === "t" ? null : +idx);
  };
}

/* ---- AR Corporate breakdown pop-up (invoice list behind a clicked amount) ---- */
let bdWeek = null;
function openBreakdown(weekIdx) {
  bdWeek = weekIdx;
  const r = report, a = r.arRes, c = compute(r);
  const lines = a.lines.filter((x) => (weekIdx == null ? true : x.slot === weekIdx) && remarkMatch(x, $("bd-rem").value));
  lines.sort((x, y) => y.cash - x.cash);
  const sum = (k) => lines.reduce((s, x) => s + x[k], 0);
  const inForecast = lines.filter((x) => typeof x.slot === "number");
  const w = weekIdx == null ? null : r.wk[weekIdx];
  $("bd-title").textContent = weekIdx == null ? "AR Corporate — all weeks" : `AR Corporate — Week ${w.n} (${fmtDate(w.start)} – ${fmtDate(w.end)})`;
  const shown = weekIdx == null ? c.v.arCorp.reduce((s, v) => s + v, 0) : c.v.arCorp[weekIdx];
  const calc = weekIdx == null ? a.weekly.reduce((s, v) => s + v, 0) : a.weekly[weekIdx];
  const edited = Math.abs(shown - calc) > 0.005;
  const chips = [
    ["Invoices", String(lines.length)], ["Invoice amount", fmtNum(sum("amount"))], ["EWT deducted", fmtNum(-sum("ewt"))],
    [weekIdx == null ? "Expected cash (all invoices)" : "Expected cash", fmtNum(sum("cash"))],
  ];
  if (weekIdx == null) chips.push(["Within the forecast weeks", fmtNum(inForecast.reduce((s, x) => s + x.cash, 0))]);
  $("bd-chips").innerHTML = chips.map(([k, v]) => `<div class="bd-chip"><span>${k}</span><b>${v}</b></div>`).join("");
  $("bd-note").innerHTML = edited
    ? `⚠ This amount was edited manually: calculated <b>${fmtNum(calc)}</b>, shown in the table <b>${fmtNum(shown)}</b>. The list below is the calculated amount.`
    : (weekIdx == null ? "Includes invoices that are not in the forecast (Projected Date "-", overdue over 120 days without a Projected Date, after the last week, or no due date), marked in the last column." : "");
  $("bd-note").classList.toggle("hidden", !$("bd-note").innerHTML);
  const head = "<tr><th>Customer</th><th>Invoice</th><th>Type</th><th>Due date</th><th>Days overdue</th><th>Projected date</th><th>Invoice amount</th><th>EWT</th><th>Cash expected</th><th>Why it lands here</th><th>Remarks</th></tr>";
  const pd = (x) => (x.projDash ? "-" : x.proj ? x.proj.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" }) : "");
  const body = lines.map((x) => {
    const wk = typeof x.slot === "number" ? `Week ${x.slot + 1}` : "—";
    return `<tr><td class="l">${escapeHtml(x.party)}</td><td class="l">${escapeHtml(x.inv || "")}</td><td class="l">${escapeHtml(x.type || "")}</td>
      <td>${x.due ? x.due.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" }) : "—"}</td>
      <td>${x.dpd == null ? "—" : x.dpd > 0 ? x.dpd : "not due"}</td><td>${pd(x)}</td><td>${fmtNum(x.amount)}</td><td>${fmtNum(-x.ewt)}</td><td><b>${fmtNum(x.cash)}</b></td>
      <td class="l why">${weekIdx == null ? `<b>${wk}</b> · ` : ""}${escapeHtml(x.why)}</td><td class="l rem">${escapeHtml(x.rem || "")}</td></tr>`;
  }).join("");
  const foot = `<tr class="grand"><td class="l" colspan="6">Total (${lines.length} invoices)</td><td>${fmtNum(sum("amount"))}</td><td>${fmtNum(-sum("ewt"))}</td><td>${fmtNum(sum("cash"))}</td><td></td><td></td></tr>`;
  $("bd-table").innerHTML = `<table><thead>${head}</thead><tbody>${body || '<tr><td class="l" colspan="11">No invoices match.</td></tr>'}${body ? foot : ""}</tbody></table>`;
  $("bd-modal").classList.remove("hidden");
}
function closeBreakdown() { $("bd-modal").classList.add("hidden"); }

/* ---- drag the added rows to any position in Cash In / Cash Out ---- */
let dragKey = null;
function clearDropMarks() { document.querySelectorAll("tr.drop-before, tr.drop-after").forEach((t) => t.classList.remove("drop-before", "drop-after")); }
// Where would the dragged row land if dropped on this <tr>?  -> { sec, index, before }
function dropTarget(tr, e) {
  const d = tr.dataset.drop;
  if (d) {
    const [sec, where] = d.split(":");
    return { sec, index: where === "start" ? 0 : rowOrder[sec].length, before: where === "end" };
  }
  const k = tr.dataset.key;
  const sec = rowOrder.in.includes(k) ? "in" : rowOrder.out.includes(k) ? "out" : null;
  if (!sec) return null;
  const rect = tr.getBoundingClientRect(), before = e.clientY < rect.top + rect.height / 2;
  const idx = rowOrder[sec].indexOf(k);
  return { sec, index: before ? idx : idx + 1, before };
}
function moveRow(key, t) {
  const from = rowOrder.in.includes(key) ? "in" : "out";
  const fi = rowOrder[from].indexOf(key);
  rowOrder[from].splice(fi, 1);
  let ti = t.index;
  if (from === t.sec && fi < ti) ti--;
  rowOrder[t.sec].splice(ti, 0, key);
  saved = false;
  renderSummary();
  updateSaveState();
}
function wireRowDrag(wrap) {
  wrap.querySelectorAll(".cr-handle").forEach((h) => {
    const tr = h.closest("tr");
    h.addEventListener("mousedown", () => (tr.draggable = true));
    h.addEventListener("mouseup", () => (tr.draggable = false));
    tr.addEventListener("dragstart", (e) => {
      dragKey = tr.dataset.key;
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", dragKey);
      tr.classList.add("dragging");
    });
    tr.addEventListener("dragend", () => { tr.draggable = false; tr.classList.remove("dragging"); clearDropMarks(); dragKey = null; });
  });
  wrap.querySelectorAll("tbody tr").forEach((tr) => {
    tr.addEventListener("dragover", (e) => {
      if (!dragKey) return;
      const t = dropTarget(tr, e);
      if (!t) return;
      e.preventDefault();
      clearDropMarks();
      tr.classList.add(t.before ? "drop-before" : "drop-after");
    });
    tr.addEventListener("drop", (e) => {
      if (!dragKey) return;
      e.preventDefault();
      const t = dropTarget(tr, e);
      clearDropMarks();
      if (t) moveRow(dragKey, t);
    });
  });
}

// Live-update the derived cells while the user types, without rebuilding inputs (keeps focus).
function refreshDerived() {
  const c = compute(report);
  rowsDef().forEach((row) => {
    if (row.section || row.note) return;
    c.v[row.k].forEach((v, j) => setCell(row.k, j, v));
    setCell(row.k, "t", rowTotal(row, c));
  });
  renderCards(c);
  $("btn-reset").classList.toggle("hidden", !hasEdits());
  updateSaveState();
}
function onCellEdit(e) {
  const inp = e.target, k = inp.dataset.row, i = +inp.dataset.i;
  const val = parseFloat(inp.value);
  overrides[k] = overrides[k] || {};
  overrides[k][wkIso(i)] = isNaN(val) ? 0 : val;
  inp.parentElement.classList.add("edited");
  saved = false;
  refreshDerived();
}
function setCell(k, i, v) {
  const td = document.querySelector(`[data-cell="${k}:${i}"]`);
  if (!td) return;
  td.textContent = fmtNum(v);
  td.classList.toggle("neg", v < 0);
}

function detailRows(r) {
  return [...r.arRes.byParty.entries()].map(([name, p]) => {
    const excl = p.rule + p.later + p.nodue + p.dash;
    return { name, weekly: p.weekly, excl, total: p.weekly.reduce((s, v) => s + v, 0) + excl };
  }).sort((a, b) => b.total - a.total);
}
function renderDetail() {
  const r = report, rows = detailRows(r), n = r.weeks;
  const tw = Array(n).fill(0); let te = 0, tt = 0;
  rows.forEach((x) => { x.weekly.forEach((v, i) => (tw[i] += v)); te += x.excl; tt += x.total; });
  const head = `<tr><th>Customer (AR Corporate)</th>${r.wk.map((w) => `<th>Wk ${w.n}</th>`).join("")}<th>Not in forecast</th><th>Total cash</th></tr>`;
  const body = rows.map((x) => `<tr class="in"><td>${escapeHtml(x.name)}</td>${x.weekly.map((v) => `<td>${fmtNum(v)}</td>`).join("")}<td>${fmtNum(x.excl)}</td><td>${fmtNum(x.total)}</td></tr>`).join("");
  const foot = `<tr class="grand"><td>Total (${rows.length} customers)</td>${tw.map((v) => `<td>${fmtNum(v)}</td>`).join("")}<td>${fmtNum(te)}</td><td>${fmtNum(tt)}</td></tr>`;
  $("table-detail").innerHTML = `<table><thead>${head}</thead><tbody>${body}${foot}</tbody></table>`;
}
/* ---- Past Due tab: every overdue invoice, grouped by aging bucket ---- */
const PD_BUCKETS = [
  { label: "1–30 days overdue", short: "1–30 days", max: 30 }, { label: "31–60 days overdue", short: "31–60 days", max: 60 }, { label: "61–90 days overdue", short: "61–90 days", max: 90 },
  { label: "91–120 days overdue", short: "91–120 days", max: 120 }, { label: "Over 120 days overdue", short: "Over 120 days", max: Infinity },
];
let pdBucket = "all";
let pdRemark = "all";
function renderPastDue() {
  if (!report) return;
  const q = $("pd-search").value.trim().toLowerCase();
  const all = report.arRes.lines.filter((x) => x.dpd != null && x.dpd > 0);
  const bucketOf = (d) => PD_BUCKETS.findIndex((b) => d <= b.max);
  const bySearch = all.filter((x) => !q || x.party.toLowerCase().includes(q) || (x.inv || "").toLowerCase().includes(q) || (x.rem || "").toLowerCase().includes(q));
  const inBucket = (x) => pdBucket === "all" || bucketOf(x.dpd) === +pdBucket;
  const rows = bySearch.filter((x) => inBucket(x) && remarkMatch(x, pdRemark)).sort((a, c) => c.dpd - a.dpd || c.amount - a.amount);
  const sum = (a) => a.reduce((s, x) => s + x.amount, 0);
  const custs = new Set(rows.map((x) => x.party)).size;
  $("pd-sub").textContent = `Open invoices past their due date as of ${report.rules.asOf.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" })}. Oldest first.`;
  $("pd-chips").innerHTML = [["Past due invoices", rows.length.toLocaleString("en-PH")], ["Customers", custs.toLocaleString("en-PH")], ["Total past due", fmtNum(sum(rows))],
    ["Share of all open AR", report.arRes.gross ? (sum(rows) / report.arRes.gross * 100).toFixed(1) + "%" : "-"]]
    .map(([k, v]) => `<div class="bd-chip"><span>${k}</span><b>${v}</b></div>`).join("");

  // days-bucket buttons (counts follow the search box and the remarks filter)
  const forBuckets = bySearch.filter((x) => remarkMatch(x, pdRemark));
  const cnt = PD_BUCKETS.map((_, i) => forBuckets.filter((x) => bucketOf(x.dpd) === i).length);
  const btn = (key, label, n, cur, attr) => `<button type="button" class="pd-fbtn${String(cur) === String(key) ? " active" : ""}" data-${attr}="${key}">${label}<span>${n.toLocaleString("en-PH")}</span></button>`;
  $("pd-filter").innerHTML = '<span class="pd-flabel">Days overdue</span>' + btn("all", "All", forBuckets.length, pdBucket, "b") + PD_BUCKETS.map((b, i) => btn(i, b.short, cnt[i], pdBucket, "b")).join("");
  $("pd-filter").querySelectorAll(".pd-fbtn").forEach((el) => (el.onclick = () => { pdBucket = el.dataset.b; renderPastDue(); }));
  // remarks buttons (counts follow the search box and the days-bucket filter)
  const forRemarks = bySearch.filter(inBucket);
  $("pd-remark").innerHTML = REMARK_FILTERS.map(([k, label]) =>
    `<option value="${k}"${k === pdRemark ? " selected" : ""}>${label} (${forRemarks.filter((x) => remarkMatch(x, k)).length.toLocaleString("en-PH")})</option>`).join("");

  if (!rows.length) { $("pd-list").innerHTML = `<p class="muted" style="padding:16px 4px">${all.length ? "No past due invoices match the filter." : "No past due invoices as of this date."}</p>`; return; }
  const dt = (d) => d.toLocaleDateString("en-PH", { month: "short", day: "numeric", year: "numeric" });
  $("pd-list").innerHTML = `<table><thead><tr><th>Customer</th><th>Invoice</th><th>Type</th><th>Due date</th><th>Days overdue</th><th>Bucket</th><th>Projected date</th><th>Invoice amount</th><th>Remarks</th></tr></thead><tbody>
    ${rows.map((x) => `<tr><td class="l">${escapeHtml(x.party)}</td><td class="l">${escapeHtml(x.inv || "")}</td><td class="l">${escapeHtml(x.type || "")}</td>
      <td>${dt(x.due)}</td><td>${x.dpd}</td><td class="l">${PD_BUCKETS[bucketOf(x.dpd)].short}</td><td>${x.projDash ? "-" : x.proj ? dt(x.proj) : ""}</td><td>${fmtNum(x.amount)}</td><td class="l rem">${escapeHtml(x.rem || "")}</td></tr>`).join("")}
    </tbody><tfoot><tr class="grand"><td class="l" colspan="7">Total — ${rows.length.toLocaleString("en-PH")} invoices</td><td>${fmtNum(sum(rows))}</td><td></td></tr></tfoot></table>`;
}

function renderNotes() {
  const { ar, ap, arRes: a, weeks } = report, tot = (x) => x.items.reduce((s, it) => s + it.amount, 0), notes = [];
  if (ar.items.length) {
    notes.push(`AR Corporate: ${a.count} invoices, gross ${fmtNum(a.gross)}, EWT deducted ${fmtNum(a.ewt)}, expected cash ${fmtNum(a.cash)}.`);
    notes.push(`Not in the forecast: overdue-rule exclusions ${fmtNum(a.rule)}, after Week ${weeks} ${fmtNum(a.later)}, no due date ${fmtNum(a.nodue)}.`);
  }
  if (ap.items.length) notes.push(`AP file read: ${ap.items.length} items, total ${fmtNum(tot(ap))} — AP logic is not built yet, so Cash Out is empty.`);
  if (state.cf.rows.length) {
    const st = retailStats(), RM = $("rf-basis").value === "trend" ? retailModel() : null;
    if (RM) notes.push(`Retail (Credit Card): weekly trend from the Cash Flow Report (${RM.n} weeks of history, ${RM.method === "pattern" ? "pattern repeated" : RM.method === "linear" ? "straight-line trend" : "flat average"}) for every week after the last actual date (${st.last.toLocaleDateString("en-PH", { month: "short", day: "numeric" })}). The row can be edited (Edit).`);
    else if (st) notes.push(`Retail (Credit Card): from the Cash Flow Report — average ${fmtNum(st.avg)} per business day over the last ${st.n === 1 ? "week" : st.n + " weeks"}; ${$("rf-basis").value === "daily" ? "daily average as shown" : "× 5 business days"}, ${$("rf-apply").value === "all" ? "every week" : "Week 1 only"} after the last actual date (${st.last.toLocaleDateString("en-PH", { month: "short", day: "numeric" })}). The row can be edited (Edit).`);
  }
  if (state.ds.rows.length) {
    const inWin = state.ds.rows.filter((x) => { const i = Math.round((x.monday - report.start) / (7 * DAY)); return i >= 0 && i < weeks; });
    notes.push(`Proceeds from Disposal (Unit Disposals file, Weekly Cash Flow "Disposal proceeds"): ${fmtNum(inWin.reduce((s, x) => s + x.amount, 0))} in these ${weeks} weeks.`);
  }
  notes.push("Loan and affiliate rows are typed in with Edit.");
  $("notes").textContent = notes.join(" ");
}

// Bank balances typed in Setup are remembered by this browser (each computer keeps its own) and filled in next time.
const ACCT_KEY = "cf-accounts";
function saveAccounts() { bankDirty = true; try { localStorage.setItem(ACCT_KEY, JSON.stringify(accounts)); } catch (_) {} renderBankInfo(); }
function restoreAccounts() {
  try {
    const a = JSON.parse(localStorage.getItem(ACCT_KEY) || "null");
    if (Array.isArray(a) && a.length) accounts.splice(0, accounts.length, ...a.map((x) => ({ bank: String(x.bank || ""), no: String(x.no || ""), purpose: String(x.purpose || ""), bal: +x.bal || 0 })));
  } catch (_) {}
}
function renderAccounts() {
  const tb = document.querySelector("#tbl-accounts tbody");
  tb.innerHTML = "";
  accounts.forEach((a, i) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input type="text" value="${escapeHtml(a.bank)}" data-k="bank"></td><td><input type="text" value="${escapeHtml(a.no)}" data-k="no"></td>
      <td><input type="text" value="${escapeHtml(a.purpose)}" data-k="purpose"></td><td><input type="number" step="0.01" value="${a.bal}" data-k="bal"></td>
      <td><button type="button" class="btn ghost small" data-del>✕</button></td>`;
    tr.querySelectorAll("input").forEach((inp) => (inp.oninput = () => {
      a[inp.dataset.k] = inp.dataset.k === "bal" ? parseFloat(inp.value) || 0 : inp.value;
      saveAccounts(); saved = false; refresh();
    }));
    tr.querySelector("[data-del]").onclick = () => { accounts.splice(i, 1); saveAccounts(); renderAccounts(); refresh(); };
    tb.appendChild(tr);
  });
  $("bal-total").textContent = fmtNum(acctTotal());
}

/* ============================================================
   Edit / Save
   ============================================================ */
// Fingerprint of everything that ends up in the saved file. Used to block saving an unchanged summary.
function hash53(str) {
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < str.length; i++) { const ch = str.charCodeAt(i); h1 = Math.imul(h1 ^ ch, 2654435761); h2 = Math.imul(h2 ^ ch, 1597334677); }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}
function summarySignature() {
  if (!report) return "";
  const c = compute(report), r2 = (n) => Math.round(n * 100) / 100;
  return hash53(JSON.stringify({
    company: $("opt-company").value, start: toISO(report.start), weeks: report.weeks,
    accounts: accounts.map((a) => [a.bank, a.no, a.purpose, r2(a.bal)]),
    rows: rowsDef().filter((x) => x.k).map((x) => [x.k, x.custom ? x.label : "", c.v[x.k].map(r2)]),
    fileEdits: typeof FILE_KINDS === "undefined" ? [] : FILE_KINDS.map(([k]) => (hasOriginal(k) ? editCount(k) : 0)),
  }));
}
function updateSaveState() {
  const btn = $("btn-save");
  if (typeof saving !== "undefined" && saving) return;
  $("btn-edit").disabled = !report;
  $("btn-add-row").disabled = !report;
  if (!report) { btn.disabled = true; return; }
  const pending = typeof SHARED_KINDS !== "undefined" && (SHARED_KINDS.some((k) => state[k].dirty || state[k].syncOriginal) || bankDirty || (typedBase && typedChanges(typedNow(), typedBase).length > 0));
  const same = !pending && lastSaved && lastSaved.sig === summarySignature();
  btn.disabled = !!same;
  btn.textContent = same ? "Saved ✓ (no changes)" : "Save summary";
  btn.title = same ? "Nothing changed since the last save. Upload a new file or make an edit to save again." : "Saves the parts you changed for everyone, and updates Cashflow_Summary_LATEST.xlsx with everyone's latest work.";
  if ($("btn-final")) $("btn-final").disabled = !report;
}

// Newest file in "Cashflow Summary" (names sort by date), plus the fingerprint stored inside it.
async function loadLatestSummary() {
  lastSaved = null;
  try {
    const dir = await baseDir.getDirectoryHandle(SUMMARY_DIR, { create: true });
    const names = [];
    for await (const [n, h] of dir.entries()) if (h.kind === "file" && /^Cashflow_Summary_.*\.xlsx$/i.test(n)) names.push(n);
    names.sort();
    const latest = names.includes(LATEST_NAME) ? LATEST_NAME : names.pop();
    if (latest) {
      const f = await (await dir.getFileHandle(latest)).getFile();
      const ab = await f.arrayBuffer();
      let wb = XLSX.read(ab, { type: "array", bookProps: true });
      if (!/cashflow-sig:/.test((wb.Props && wb.Props.Comments) || "")) wb = XLSX.read(ab, { type: "array", sheets: ["Meta"] });   // older summaries
      const m = /cashflow-sig:(\S+)/.exec((wb.Props && wb.Props.Comments) || "");
      const meta = wb.Sheets && wb.Sheets["Meta"];
      lastSaved = { name: latest, sig: m ? m[1] : meta && meta.B1 ? String(meta.B1.v) : null };
      // carry the hand-entered items of the newest saved summary into this session (matched by week date)
      const ks = /cf-state:(\S+)/.exec((wb.Props && wb.Props.Keywords) || "");
      if (ks && !sharedFound) {
        try {
          const st = decodeState(ks[1]);
          const a = applyCarryState(st), b = applyCashOutState(st.cashOut, latest), u = applyDisposalState(st.unitDisposals, latest);
          if (a) typedBase = emptyTyped();                       // carried amounts are shared on the first Save
          if (a || b || u) refresh();
        } catch (_) { /* unreadable state: ignore */ }
      }
    }
  } catch (_) { lastSaved = null; }
  $("last-saved").textContent = lastSaved ? (lastSaved.name === LATEST_NAME ? "Latest summary in FILES" : `Last saved: ${lastSaved.name}`) : "";
  $("last-saved").title = lastSaved ? lastSaved.name : "";
  updateSaveState();
}

function toggleEdit() {
  editing = !editing;
  $("btn-edit").classList.toggle("on", editing);
  $("btn-edit").textContent = editing ? "✓ Done editing" : "✎ Edit";
  $("edit-hint").textContent = editing
    ? "Edit mode — change any Cash In amount, or use + Add row for deductions and other lines. Totals update as you type."
    : "Preview only — nothing is saved until you click Save summary.";
  $("edit-hint").classList.toggle("on", editing);
  if (report) renderSummary();
}

// Builds the saved file in the format of "Cash Forecast-EEL Template Blank.xls" (single "Summary" sheet).
// Totals / balances are live Excel formulas, with the computed values cached so every viewer shows numbers.
const XL_BLUE = "FF0066CC", XL_NAVY = "FF333399", XL_NUM = "#,##0.00";
// Excel stores dates without a timezone: use UTC midnight so Oct 5 stays Oct 5 regardless of the PC timezone.
const xlDate = (d) => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
function xlCol(n) { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function xlLabel(k) {
  if (k.startsWith("c:")) { const x = customRows.find((r) => "c:" + r.id === k); return (x.name || "(unnamed row)") + (x.sign < 0 ? " (deduction)" : ""); }
  return (CASH_IN.find(([kk]) => kk === k) || CASH_OUT.find(([kk]) => kk === k))[1];
}

// Hand-entered items, kept in the saved file's document properties (no extra sheet) so next week can pick them up.
function carryState() {
  const cutoff = report ? addDays(report.start, -28) : null;               // forget weeks that are long past
  const keep = (m) => Object.fromEntries(Object.entries(m || {}).filter(([d]) => !cutoff || new Date(d + "T00:00:00") >= cutoff));
  return {
    v: 1,
    rows: customRows.map((r) => ({ id: r.id, name: r.name, sign: r.sign, vals: keep(r.vals) })),
    order: { in: rowOrder.in.slice(), out: rowOrder.out.slice() },
    overrides: Object.fromEntries(Object.entries(overrides).map(([k, m]) => [k, keep(m)])),
    unitDisposals: state.ds.rows.length ? { name: state.ds.file ? state.ds.file.name : state.ds.savedFrom || "", sheet: state.ds.sheet || "",
      rows: state.ds.rows.map((x) => ({ from: toISO(x.from), to: toISO(x.to), amount: x.amount, units: x.units, monday: toISO(x.monday) })) } : null,
    cashOut: Object.fromEntries(Object.keys(CO_META).filter((k) => Object.keys(state[k].byDate).length).map((k) => [k, { byDate: keep(state[k].byDate), info: state[k].info, name: state[k].file ? state[k].file.name : state[k].savedFrom || "" }])),
  };
}
const encodeState = (o) => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
const decodeState = (s) => JSON.parse(decodeURIComponent(escape(atob(s))));

function applyCarryState(st) {
  if (!st || !st.rows || hasEdits() || customRows.length) return false;      // never overwrite what is on screen
  const idMap = {};
  st.rows.forEach((r) => { const id = nextRowId++; idMap[r.id] = id; customRows.push({ id, uid: r.uid || newUid(), name: r.name, sign: r.sign, vals: { ...r.vals } }); });
  const fixed = new Set([...CASH_IN, ...CASH_OUT].map(([k]) => k));
  const map = (k) => (k.startsWith("c:") ? (idMap[k.slice(2)] ? "c:" + idMap[k.slice(2)] : null) : fixed.has(k) ? k : null);
  const ord = (a) => (a || []).map(map).filter(Boolean);
  const order = { in: ord(st.order && st.order.in), out: ord(st.order && st.order.out) };
  CASH_IN.forEach(([k]) => { if (!order.in.includes(k) && !order.out.includes(k)) order.in.push(k); });
  CASH_OUT.forEach(([k]) => { if (!order.in.includes(k) && !order.out.includes(k)) order.out.push(k); });
  customRows.forEach((r) => { const k = "c:" + r.id; if (!order.in.includes(k) && !order.out.includes(k)) order.in.push(k); });
  rowOrder.in = order.in; rowOrder.out = order.out;
  for (const [k, m] of Object.entries(st.overrides || {})) overrides[k] = { ...m };
  return true;
}

// Unit Disposals schedule from the last saved summary, used when no Unit Disposals file was uploaded in this session.
function applyDisposalState(ud, fileName) {
  if (!ud || !ud.rows || !ud.rows.length || state.ds.rows.length) return false;
  const D = (s) => new Date(s + "T00:00:00");
  state.ds.rows = ud.rows.map((x) => ({ from: D(x.from), to: D(x.to), amount: x.amount, units: x.units, monday: D(x.monday) }));
  Object.assign(state.ds, { file: null, sheet: ud.sheet || "", savedFrom: ud.name ? `${ud.name}, via ${fileName}` : fileName, from: "saved" });
  const tot = state.ds.rows.reduce((s, x) => s + x.amount, 0), units = state.ds.rows.reduce((s, x) => s + (x.units || 0), 0);
  $("drop-ds").classList.add("loaded");
  $("mstat-ds").innerHTML = `Using the schedule kept in the last saved summary<br><span class="muted">${units ? units + " units · " : ""}${fmtNum(tot)} disposal proceeds</span>`;
  renderDisposalMonths();
  if (typeof renderFileList === "function") renderFileList();
  return true;
}

// Cash Out sources from the last saved summary are used when no file was uploaded for them in this session.
function applyCashOutState(co, fileName) {
  if (!co) return false;
  let used = false;
  for (const [k, v] of Object.entries(co)) {
    if (!CO_META[k] || Object.keys(state[k].byDate).length) continue;
    Object.assign(state[k], { byDate: { ...v.byDate }, info: v.info || "", from: "saved", savedFrom: v.name ? `${v.name}, via ${fileName}` : fileName, rows: [1] });
    $("drop-" + k).classList.add("loaded");
    $("mstat-" + k).innerHTML = `Using the amounts kept in the last saved summary<br><span class="muted">${escapeHtml(v.info || "")}</span>`;
    used = true;
  }
  if (used) { renderCashOutSources(); if (typeof renderFileList === "function") renderFileList(); }
  return used;
}

/* ============================================================
   Detail sheets of the saved summary: a copy of each source sheet as uploaded (formatting kept), with columns
   (or a row, for week-across layouts) added at the right end / bottom: the Summary week of each line, and for
   AR the EWT and net cash. The Summary amounts are formulas (SUMIFS) that point to those added cells.
   ============================================================ */
const qs = (n) => `'${n.replace(/'/g, "''")}'`;
const ADD_HEAD = { font: { name: "Calibri", size: 11, bold: true, color: { argb: "FF1F4E79" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } } };
const ADD_CELL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F7FC" } };
const SLOT_TXT = { nodue: "Not in forecast (no due date)", dash: 'Not in forecast ("-")', rule: "Not in forecast (overdue rule)", later: "After the last forecast week" };

function addHead(ws, r, c, text, width) {
  const x = ws.getCell(r, c);
  x.value = text; x.font = ADD_HEAD.font; x.fill = ADD_HEAD.fill; x.alignment = { wrapText: true, vertical: "middle" };
  if (width) ws.getColumn(c).width = width;
}
function addCell(ws, r, c, value, fmt) {
  const x = ws.getCell(r, c);
  x.value = value instanceof Date ? xlDate(value) : value; x.fill = ADD_CELL;
  x.numFmt = fmt || (value instanceof Date ? "mm/dd/yyyy" : typeof value === "number" || (value && value.formula) ? XL_NUM : "General");
  return x;
}
// value of a copied cell: formulas that point to other sheets / files keep their last value
function copyValue(cell) {
  const v = cell.value;
  if (v && typeof v === "object" && ("formula" in v || "sharedFormula" in v)) {
    let f = null; try { f = cell.formula; } catch (_) {}
    const res = v.result === undefined ? null : v.result;
    return !f || /!|\[/.test(f) ? res : { formula: f, result: v.result };
  }
  return v;
}
function copySheet(dst, src, name) {
  const ws = dst.addWorksheet(name);
  try { if (src.views && src.views.length) ws.views = JSON.parse(JSON.stringify(src.views)).map((v) => Object.assign(v, { tabSelected: false })); } catch (_) {}
  if (src.properties && src.properties.defaultRowHeight) ws.properties.defaultRowHeight = src.properties.defaultRowHeight;
  for (let c = 1; c <= src.columnCount; c++) { const a = src.getColumn(c), b = ws.getColumn(c); if (a.width) b.width = a.width; if (a.hidden) b.hidden = true; }
  src.eachRow({ includeEmpty: true }, (row, rn) => {
    const d = ws.getRow(rn);
    if (row.height) d.height = row.height;
    if (row.hidden) d.hidden = true;
    row.eachCell({ includeEmpty: true }, (cell, cn) => {
      const t = d.getCell(cn);
      if (cell.type !== ExcelJS.ValueType.Merge) t.value = copyValue(cell);
      if (cell.style) {
        t.style = JSON.parse(JSON.stringify(cell.style));
        const fg = t.style.fill && t.style.fill.fgColor && t.style.fill.fgColor.argb;
        if (fg === "FFFFF2CC") delete t.style.fill;                    // the system's "edited" yellow is not shown in the summary file
      }
    });
  });
  Object.values(src._merges || {}).forEach((m) => { try { ws.mergeCells(m.range); } catch (_) {} });
  return ws;
}
// values only (when the file cannot be opened with formatting, e.g. an old .xls)
function plainSheet(dst, sw, name) {
  const ws = dst.addWorksheet(name), rg = XLSX.utils.decode_range(sw["!ref"] || "A1");
  for (let r = rg.s.r; r <= rg.e.r; r++) for (let c = rg.s.c; c <= rg.e.c; c++) {
    const x = sw[XLSX.utils.encode_cell({ r, c })];
    if (!x || x.v == null) continue;
    const t = ws.getCell(r + 1, c + 1);
    t.value = x.v instanceof Date ? xlDate(x.v) : x.v;
    if (x.v instanceof Date) t.numFmt = "mm/dd/yyyy"; else if (x.z) t.numFmt = x.z;
  }
  (sw["!cols"] || []).forEach((c, i) => { if (c && c.wch) ws.getColumn(i + 1).width = c.wch; });
  return ws;
}
async function sourceSheet(dst, k, sjsWs, sheet, name, notes) {
  try {
    const wb = await excelOf(k), src = wb.getWorksheet(sheet);
    if (src) return copySheet(dst, src, name);
  } catch (e) { notes.push(`${name}: copied as values only (${e.message})`); }
  return plainSheet(dst, sjsWs, name);
}
// a simple table (sources remembered from the last saved summary: no file to copy)
function tableSheet(dst, name, title, heads, rows) {
  const ws = dst.addWorksheet(name);
  ws.getCell("A1").value = title; ws.getCell("A1").font = { bold: true, size: 12 };
  heads.forEach(([h, w], i) => addHead(ws, 3, i + 1, h, w));
  rows.forEach((r, j) => r.forEach((v, i) => { if (v != null) addCell(ws, 4 + j, i + 1, v); }));
  return { ws, first: 4, last: 3 + Math.max(1, rows.length) };
}
const sumifs = (name, amt, crit, L) => `SUMIFS(${qs(name)}!${amt},${qs(name)}!${crit},${L}$9)`;
const colRange = (c, r1, r2) => `$${xlCol(c)}$${r1}:$${xlCol(c)}$${r2}`;
const rowRange = (r, c1, c2) => `$${xlCol(c1)}$${r}:$${xlCol(c2)}$${r}`;

// Retail basis of each forecast week (mirrors retailForecast)
function retailBasis(wk) {
  const st = retailStats();
  if (!st) return wk.map(() => ({ actual: false, text: "" }));
  const basis = $("rf-basis").value, M = basis === "trend" ? retailModel() : null;
  let given = false;
  return wk.map((w) => {
    if (w.start <= st.last) return { actual: true, text: "Actual (Cash Flow Report)" };
    if (M) {
      if (M.method === "pattern") { const i = ((M.weekIndex(w.start) - M.n) % M.n + M.n) % M.n; return { text: `History pattern: same as the week of ${fmtDate(M.hist[i].monday)}` }; }
      return { text: M.method === "linear" ? "Straight-line trend of the history weeks" : "Average of the history weeks" };
    }
    if ($("rf-apply").value === "all" || !given) { given = true; return { text: basis === "daily" ? `Average per business day (last ${st.n} weeks)` : `Average per business day × 5 (last ${st.n} weeks)` }; }
    return { text: "" };
  });
}

function bankSheet(dst, r) {
  const name = "Bank Balances", ws = dst.addWorksheet(name);
  [14, 22, 34, 18].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  ws.getCell("A1").value = "Bank balances — Beginning Balance of Week 1"; ws.getCell("A1").font = { bold: true, size: 12 };
  ws.getCell("A2").value = `Week 1 starts ${r.wk[0].start.toLocaleDateString("en-PH", { month: "long", day: "numeric", year: "numeric" })}`;
  ["Bank", "Account no.", "Purpose", "Balance (PHP)"].forEach((h, i) => addHead(ws, 4, i + 1, h));
  accounts.forEach((a, j) => {
    const row = 5 + j, usd = /USD\s*([\d,.]+)\s*[×x*]\s*([\d.]+)/i.exec(a.purpose || "");
    ws.getCell(row, 1).value = a.bank; ws.getCell(row, 2).value = a.no; ws.getCell(row, 3).value = a.purpose;
    const b = ws.getCell(row, 4);
    b.value = usd ? { formula: `${usd[1].replace(/,/g, "")}*${usd[2]}`, result: a.bal || 0 } : a.bal || 0; b.numFmt = XL_NUM;
  });
  const tot = 5 + accounts.length;
  ws.getCell(tot, 1).value = "Total"; ws.getCell(tot, 1).font = { bold: true };
  const t = ws.getCell(tot, 4);
  t.value = { formula: accounts.length ? `SUM(D5:D${tot - 1})` : "0", result: acctTotal() }; t.numFmt = XL_NUM; t.font = { bold: true }; t.border = { top: { style: "thin" }, bottom: { style: "double" } };
  return `${qs(name)}!$D$${tot}`;
}

// Builds the detail sheets; returns { begin, links: { rowKey: (L) => formula } , notes }
async function addDetailSheets(dst, r) {
  const links = {}, notes = [], wk = r.wk, cache = {};
  const begin = bankSheet(dst, r);
  const sjs = (k) => editWorkbook(k);

  // AR Corporate: copy of the Detail sheet + Summary week / EWT / Net cash / How it was placed
  if (state.ar.buffer && r.arRes.lines.length) {
    try {
      const s = state.ar, name = "AR Corporate", sw = sjs("ar").Sheets[s.sheet], rg = XLSX.utils.decode_range(sw["!ref"] || "A1");
      const hdr = rg.s.r + findHeader(gridOf(sw)).headerIdx + 1;
      const ws = await sourceSheet(dst, "ar", sw, s.sheet, name, notes);
      const W = rg.e.c + 2, E = W + 1, N = W + 2, Y = W + 3, aCol = rg.s.c + s.map.amount, A = xlCol(aCol + 1);
      addHead(ws, hdr, W, "Summary week (Monday)", 14); addHead(ws, hdr, E, "EWT", 13); addHead(ws, hdr, N, "Net cash", 15); addHead(ws, hdr, Y, "How it was placed", 60);
      let first = Infinity, last = 0;
      for (const ln of r.arRes.lines) {
        if (ln.row == null) continue;
        const row = ln.row + 1, amtCell = sw[XLSX.utils.encode_cell({ r: ln.row, c: aCol })], isNum = amtCell && amtCell.t === "n";
        first = Math.min(first, row); last = Math.max(last, row);
        addCell(ws, row, W, typeof ln.slot === "number" ? wk[ln.slot].start : SLOT_TXT[ln.slot] || "");
        addCell(ws, row, E, ln.ewtRate && isNum ? { formula: ln.nonVat ? `${A}${row}*${ln.ewtRate}` : `${A}${row}/${1 + r.rules.ewt.vat}*${ln.ewtRate}`, result: ln.ewt } : ln.ewt, XL_NUM);
        addCell(ws, row, N, isNum ? { formula: `${A}${row}-${xlCol(E)}${row}`, result: ln.cash } : ln.cash, XL_NUM);
        addCell(ws, row, Y, ln.why);
      }
      if (last) links.arCorp = (L) => sumifs(name, colRange(N, first, last), colRange(W, first, last), L);
    } catch (e) { notes.push("AR Corporate sheet not added: " + e.message); }
  }

  // Retail: copy of the Cash Flow Report Summary sheet + Summary week per day + a weekly block used by the Summary
  if (state.cf.buffer && state.cf.rows.length) {
    try {
      const s = state.cf, name = "Cash Flow Report";
      const swb = s.wbEdit || XLSX.read(s.buffer, { type: "array", cellDates: true, sheets: [s.sheet] }), sw = swb.Sheets[s.sheet];
      const res = parseCashflowReport(swb), rg = XLSX.utils.decode_range(sw["!ref"] || "A1"), hdr = res.hdr + 1;
      const ws = await sourceSheet(dst, "cf", sw, s.sheet, name, notes);
      const W = rg.e.c + 2, B = W + 2, act = res.cAct + 1;
      addHead(ws, hdr, W, "Summary week (Monday)", 14);
      let first = Infinity, last = 0;
      for (const x of res.rows) { if (x.actual == null) continue; const row = x.r + 1; first = Math.min(first, row); last = Math.max(last, row); addCell(ws, row, W, mondayOf(x.date)); }
      addHead(ws, hdr, B, "Summary week (Monday)", 14); addHead(ws, hdr, B + 1, "Retail (Credit Card)", 17); addHead(ws, hdr, B + 2, "Basis", 46);
      const bas = retailBasis(wk);
      wk.forEach((w, i) => {
        const row = hdr + 1 + i;
        addCell(ws, row, B, w.start);
        addCell(ws, row, B + 1, bas[i].actual && last ? { formula: `SUMIFS(${colRange(act, first, last)},${colRange(W, first, last)},${xlCol(B)}${row})`, result: r.base.retail[i] } : r.base.retail[i], XL_NUM);
        addCell(ws, row, B + 2, bas[i].text);
      });
      links.retail = (L, i) => `${qs(name)}!$${xlCol(B + 1)}$${hdr + 1 + i}`;
    } catch (e) { notes.push("Cash Flow Report sheet not added: " + e.message); }
  }

  // Unit Disposals: copy of the Weekly Cash Flow sheet + a "Summary week" row under it
  if (state.ds.rows.length) {
    try {
      const s = state.ds, name = "Unit Disposals";
      if (s.buffer) {
        const res = parseUnitDisposals(sjs("ds")), sw = sjs("ds").Sheets[res.sheet], rg = XLSX.utils.decode_range(sw["!ref"] || "A1");
        const ws = await sourceSheet(dst, "ds", sw, res.sheet, name, notes), R = rg.e.r + 3;
        addHead(ws, R, rg.s.c + 1, "Summary week (Monday)");
        res.rows.forEach((x) => addCell(ws, R, x.c + 1, x.monday));
        const c1 = Math.min(...res.rows.map((x) => x.c)) + 1, c2 = Math.max(...res.rows.map((x) => x.c)) + 1;
        links.disposal = (L) => sumifs(name, rowRange(res.amtRow + 1, c1, c2), rowRange(R, c1, c2), L);
      } else {
        const t = tableSheet(dst, name, `Unit Disposals — kept in the last saved summary${s.savedFrom ? ` (${s.savedFrom})` : ""}`,
          [["From", 12], ["To", 12], ["Units", 8], ["Disposal proceeds", 18], ["Summary week (Monday)", 14]],
          s.rows.map((x) => [x.from, x.to, x.units, x.amount, x.monday]));
        links.disposal = (L) => sumifs(name, colRange(4, t.first, t.last), colRange(5, t.first, t.last), L);
      }
    } catch (e) { notes.push("Unit Disposals sheet not added: " + e.message); }
  }

  // Cash Out: one sheet per file (amounts are cash going out: the Summary shows them negative)
  const CO_SHEET = { loans: ["ln", "Loan Payments"], checks: ["uc", "Unreleased Checks"], projout: ["po", "Projected Payments"] };
  for (const [rk, [k, name]] of Object.entries(CO_SHEET)) {
    const s = state[k];
    if (!Object.keys(s.byDate).length) continue;
    try {
      if (s.buffer) {
        const swb = sjs(k), res = CO_META[k].parse(swb), loc = res.loc, sw = swb.Sheets[loc.sheet], rg = XLSX.utils.decode_range(sw["!ref"] || "A1");
        const ws = await sourceSheet(dst, k, sw, loc.sheet, name, notes);
        if (k === "po") {                                       // weeks across: a "Summary week" row (and the amount used, if there is no TOTAL row)
          const R = rg.e.r + 3, c1 = Math.min(...loc.cols.map((x) => x.c)) + 1, c2 = Math.max(...loc.cols.map((x) => x.c)) + 1;
          addHead(ws, R, rg.s.c + 1, "Summary week (Monday)");
          loc.cols.forEach((x) => addCell(ws, R, x.c + 1, x.monday));
          let amtRow = loc.totRow + 1;
          if (loc.totRow < 0) { amtRow = R + 1; addHead(ws, amtRow, rg.s.c + 1, "Amount used"); loc.cols.forEach((x) => addCell(ws, amtRow, x.c + 1, x.amt, XL_NUM)); }
          links[rk] = (L) => "-" + sumifs(name, rowRange(amtRow, c1, c2), rowRange(R, c1, c2), L);
        } else {                                                // one line per payment / check: a "Summary week" column
          const W = rg.e.c + 2, hdr = loc.hdr + 1;
          addHead(ws, hdr, W, "Summary week (Monday)", 14);
          let first = Infinity, last = 0;
          for (const x of loc.lines) { const row = x.r + 1; first = Math.min(first, row); last = Math.max(last, row); addCell(ws, row, W, x.date ? mondayOf(x.date) : "No clearing date"); }
          if (last) links[rk] = (L) => "-" + sumifs(name, colRange(loc.cAmt + 1, first, last), colRange(W, first, last), L);
        }
      } else {
        const days = Object.keys(s.byDate).sort();
        const t = tableSheet(dst, name, `${CO_META[k].label} — kept in the last saved summary${s.savedFrom ? ` (${s.savedFrom})` : ""}`,
          [["Date", 12], ["Amount", 16], ["Summary week (Monday)", 14]],
          days.map((d) => { const dt = new Date(d + "T00:00:00"); return [dt, s.byDate[d], mondayOf(dt)]; }));
        links[rk] = (L) => "-" + sumifs(name, colRange(2, t.first, t.last), colRange(3, t.first, t.last), L);
      }
    } catch (e) { notes.push(`${name} sheet not added: ` + e.message); }
  }
  return { begin, links, notes };
}

async function buildWorkbookXlsx(sig) {
  const r = report, c = compute(r), wk = r.wk, n = wk.length;
  const FIRST = 6, LAST = FIRST + n - 1, TOT = LAST + 1;               // week columns start at F; Total after the last week
  const wb = new ExcelJS.Workbook();
  wb.creator = "Hertz Cashflow Summary"; wb.lastModifiedBy = userName() || "Hertz Cashflow Summary"; wb.created = new Date(); wb.modified = new Date();
  wb.description = "cashflow-sig:" + (sig || "");
  wb.keywords = "cf-state:" + encodeState(carryState());
  const ws = wb.addWorksheet("Summary", { views: [{ state: "frozen", xSplit: 5, ySplit: 10, zoomScale: 75, showGridLines: false }] });
  const det = await addDetailSheets(wb, r);
  // amount cell of a row: a formula into its detail sheet; a typed change is kept as "+ (adjustment)"
  const amountCell = (k, i, L) => {
    const link = det.links[k], val = c.v[k][i];
    if (!link) return { v: val };
    const adj = Math.round((val - r.base[k][i]) * 100) / 100, f = link(L, i);
    return Math.abs(adj) >= 0.01 ? { v: { formula: `${f}+(${adj})`, result: val }, adj: true } : { v: { formula: f, result: val } };
  };

  // column widths as in the template
  [8.1, 29, 6.3, 9.1, 9.1].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  for (let i = FIRST; i <= LAST; i++) ws.getColumn(i).width = 16;
  ws.getColumn(TOT).width = 17;

  const base = { name: "Calibri", size: 12 };
  const put = (cell, value, o = {}) => {
    const x = ws.getCell(cell);
    x.value = value;
    x.font = { ...base, bold: !!o.bold, color: o.color ? { argb: o.color } : undefined };
    if (o.fmt) x.numFmt = o.fmt;
    if (o.align) x.alignment = { horizontal: o.align };
    if (o.border) x.border = o.border;
    return x;
  };

  // --- title block ---
  put("A1", $("opt-company").value, { bold: true });
  put("A2", "Cash Flow Forecast", { bold: true });
  put("A3", "Summary", { bold: true });
  put("A4", `${n} Week Forecast`, { bold: true });

  // --- Week # / Monday / Sunday header ---
  put("E8", "Week #", { bold: true, color: XL_BLUE, align: "right" });
  put("E9", "Monday", { bold: true, color: XL_BLUE, align: "right" });
  put("A10", "In Philippine Pesos", { bold: true, color: XL_NAVY });
  put("E10", "Sunday", { bold: true, color: XL_BLUE, align: "right" });
  wk.forEach((w, i) => {
    const L = xlCol(FIRST + i);
    put(`${L}8`, w.n, { bold: true, color: XL_BLUE, align: "center" });
    put(`${L}9`, xlDate(w.start), { bold: true, fmt: "mm/dd/yy;@", align: "center" });
    put(`${L}10`, xlDate(w.end), { bold: true, fmt: "mm/dd/yy;@", align: "center", border: { bottom: { style: "medium" } } });
  });
  put(`${xlCol(TOT)}8`, "Total", { bold: true, color: XL_BLUE, align: "center" });

  // --- rows ---
  const weekCells = (row, fn) => wk.forEach((_, i) => fn(xlCol(FIRST + i), i));
  const totalFormula = (row) => `SUM(${xlCol(FIRST)}${row}:${xlCol(LAST)}${row})`;
  const sum = (a) => a.reduce((x, y) => x + y, 0);
  let row = 12;

  // Beginning Balance
  const rBegin = row;
  put(`A${row}`, "Beginning Balance", { bold: true });
  // (formulas for weeks 2..n are filled after the Ending row is known)
  row += 2;

  // Cash In
  put(`A${row}`, "Cash In", { bold: true }); row++;
  const inFirst = row;
  for (const [k, label] of rowOrder.in.map((k) => [k, xlLabel(k)])) {
    put(`B${row}`, label);
    weekCells(row, (L, i) => { put(`${L}${row}`, amountCell(k, i, L).v, { fmt: XL_NUM }); });
    put(`${xlCol(TOT)}${row}`, { formula: totalFormula(row), result: sum(c.v[k]) }, { fmt: XL_NUM, bold: true });
    row++;
  }
  const inLast = row - 1, rTotIn = row;
  put(`A${row}`, "Total Cash In", { bold: true });
  weekCells(row, (L, i) => put(`${L}${row}`, { formula: `SUM(${L}${inFirst}:${L}${inLast})`, result: c.v.cashIn[i] }, { bold: true, fmt: XL_NUM, border: { top: { style: "thin" } } }));
  put(`${xlCol(TOT)}${row}`, { formula: totalFormula(row), result: sum(c.v.cashIn) }, { bold: true, fmt: XL_NUM, border: { top: { style: "thin" } } });
  row += 2;

  // Cash Out
  put(`A${row}`, "Cash Out", { bold: true }); row++;
  const outFirst = row;
  for (const [k, label] of rowOrder.out.map((k) => [k, xlLabel(k)])) {
    put(`B${row}`, label);
    weekCells(row, (L, i) => { put(`${L}${row}`, amountCell(k, i, L).v, { fmt: XL_NUM }); });
    put(`${xlCol(TOT)}${row}`, { formula: totalFormula(row), result: sum(c.v[k]) }, { fmt: XL_NUM, bold: true });
    row++;
  }
  const outLast = row - 1, rTotOut = row;
  put(`A${row}`, "Total Cash Out", { bold: true });
  weekCells(row, (L, i) => put(`${L}${row}`, rowOrder.out.length ? { formula: `SUM(${L}${outFirst}:${L}${outLast})`, result: c.v.cashOut[i] } : 0, { bold: true, fmt: XL_NUM, border: { top: { style: "thin" } } }));
  put(`${xlCol(TOT)}${row}`, { formula: totalFormula(row), result: sum(c.v.cashOut) }, { bold: true, fmt: XL_NUM, border: { top: { style: "thin" } } });
  row += 2;

  // Net Cash Flow
  const rNet = row;
  put(`A${row}`, "Net Cash Flow", { bold: true });
  weekCells(row, (L, i) => put(`${L}${row}`, { formula: `${L}${rTotIn}+${L}${rTotOut}`, result: c.v.net[i] }, { fmt: XL_NUM }));
  put(`${xlCol(TOT)}${row}`, { formula: totalFormula(row), result: sum(c.v.net) }, { bold: true, fmt: XL_NUM });
  row += 2;

  // Ending Cash Balance
  const rEnd = row;
  put(`A${row}`, "Ending Cash Balance", { bold: true });
  weekCells(row, (L, i) => put(`${L}${row}`, { formula: `${L}${rBegin}+${L}${rNet}`, result: c.v.end[i] },
    { bold: true, fmt: XL_NUM, align: "center", border: { top: { style: "thin" }, bottom: { style: "double" } } }));
  put(`${xlCol(TOT)}${row}`, { formula: `${xlCol(LAST)}${rEnd}`, result: c.closing }, { bold: true, fmt: XL_NUM, align: "center", border: { top: { style: "thin" }, bottom: { style: "double" } } });

  // Beginning Balance row: week 1 = opening cash, then previous week's ending balance
  weekCells(rBegin, (L, i) => {
    const v = i === 0 ? { formula: det.begin, result: c.opening } : { formula: `${xlCol(FIRST + i - 1)}${rEnd}`, result: c.v.begin[i] };
    put(`${L}${rBegin}`, v, { bold: true, fmt: XL_NUM, align: "center" });
  });
  put(`${xlCol(TOT)}${rBegin}`, { formula: `${xlCol(FIRST)}${rBegin}`, result: c.opening }, { bold: true, fmt: XL_NUM, align: "center" });

  if (det.notes.length) toast("Summary detail sheets:<small>" + det.notes.map(escapeHtml).join("<br>") + "</small>", true, 10000);
  wb.calcProperties = Object.assign(wb.calcProperties || {}, { fullCalcOnLoad: true });   // Excel recalculates the links on open
  return wb.xlsx.writeBuffer();
}

function downloadBuffer(buf, name) {
  const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// Save summary also keeps a copy of every file uploaded in this session (as it is now, incl. edits written back)
async function saveUploadedCopies() {
  const done = [];
  for (const [k] of FILE_KINDS) {
    const s = state[k];
    if (!s.file || !s.buffer || s.copySaved) continue;
    if (!uploadBatch) uploadBatch = batchFolderName();
    const path = [UPLOAD_DIR, uploadBatch];
    const name = await uniqueName(path, s.file.name);
    await writeTo(path, name, s.buffer);
    s.copySaved = true; s.saved = true; s.savedIn = `${UPLOAD_DIR}/${uploadBatch}`;
    done.push(name);
  }
  if (bankUpload && !bankUpload.copied) {
    if (!uploadBatch) uploadBatch = batchFolderName();
    const name = await uniqueName([UPLOAD_DIR, uploadBatch], bankUpload.file.name);
    await writeTo([UPLOAD_DIR, uploadBatch], name, bankUpload.buffer);
    bankUpload.copied = true; done.push(name);
  }
  return done;
}
let saving = false;
function saveStep(text) { const b = $("btn-save"); b.disabled = true; b.innerHTML = `<span class="spin"></span>${escapeHtml(text)}`; }
async function saveSummary() {
  if (!report || saving) return;
  saving = true;
  saveStep("Saving…");
  try { await saveSummaryInner(); }
  catch (e) { console.error(e); toast("Not saved — something went wrong:<small>" + escapeHtml(e && e.message ? e.message : String(e)) + "</small>", true, 12000); }
  finally { saving = false; $("btn-save").innerHTML = ""; updateSaveState(); }
}
async function saveSummaryInner(finalize) {
  // ask first (while the click is fresh) for permission to update original files
  for (const k of FILE_KINDS.map(([x]) => x).filter((x) => hasOriginal(x) && (editCount(x) || state[x].syncOriginal) && state[x].handle)) {
    try { const hd = state[k].handle; if ((await hd.queryPermission({ mode: "readwrite" })) !== "granted") await hd.requestPermission({ mode: "readwrite" }); } catch (_) {}
  }
  if (editing) toggleEdit();
  clearTimeout(reingestTimer); for (const k of Object.keys(state)) if (state[k].wbEdit && editCount(k)) await reingest(k);   // edits still being recalculated
  saveStep("Checking the save folder…");
  if (HAS_FS && !(await folderReady())) {
    toast("Choose the shared FILES folder first — then the summary is saved there.", false, 5000);
    await connectFolder();
    if (!(await folderReady())) { toast("Not saved — no save folder connected. Click Connect folder (bottom of the sidebar) and try again.", true, 7000); return; }
  }
  if (!(await folderReady())) {                                   // no shared folder: download only
    const buf = await buildWorkbookXlsx(summarySignature()), name = `Cashflow_Summary_${toISO(new Date())}_${stamp().slice(9)}.xlsx`;
    downloadBuffer(buf, name);
    toast(`No folder connected — downloaded instead<small>Connect the shared FILES folder (bottom of the sidebar) so everyone's work is combined.</small>`, false, 7000);
    return;
  }
  if (!userName() && !askUserName()) { toast("Not saved — type your name first (it is recorded with what you save).", true, 6000); return; }
  saveStep("Waiting for other saves…");
  if (!(await acquireSaveLock())) { toast("Not saved — another save is taking too long. Try again in a minute.", true, 8000); return; }
  try {
    // edits made in the uploaded files are written into their originals first (with a backup)
    let wr;
    saveStep("Updating edited files…");
    try { wr = await writeEditedOriginals(); } catch (e) { toast("Could not update the original file(s): " + escapeHtml(e.message), true, 8000); return; }
    if (!wr.ok) { toast(escapeHtml(wr.msg), true, 8000); return; }
    saveStep("Combining with everyone's updates…");
    const shared = await saveShared();
    renderFileList();
    if (!$("fe-modal").classList.contains("hidden")) renderEditor();
    if (!report) { toast(`✓ Saved<small>${shared.map(escapeHtml).join(", ") || "Nothing changed"}</small>`); return; }
    saveStep("Building the summary + detail sheets…");
    const sig = summarySignature();
    const buf = await buildWorkbookXlsx(sig);
    saveStep("Writing the summary…");
    await writeTo(SUMMARY_DIR, LATEST_NAME, buf);
    let dated = null;
    if (finalize) { dated = await uniqueName(SUMMARY_DIR, `Cashflow_Summary_${toISO(new Date())}_${stamp().slice(9)}.xlsx`); await writeTo(SUMMARY_DIR, dated, buf); }
    lastSaved = { name: LATEST_NAME, sig };
    saveStep("Saving copies of the uploaded files…");
    let copies = [];
    try { copies = await saveUploadedCopies(); } catch (e) { toast("Saved, but copies of the uploaded files could not be kept:<small>" + escapeHtml(e.message) + "</small>", true, 10000); }
    $("last-saved").textContent = `Last saved: ${whenText(new Date().toISOString())}`;
    $("last-saved").title = `${SUMMARY_DIR}/${LATEST_NAME}`;
    const chip = $("last-saved"); chip.classList.add("flash"); setTimeout(() => chip.classList.remove("flash"), 4000);
    toast(`✓ ${finalize ? "Week finalized" : "Saved"} at ${new Date().toLocaleTimeString("en-PH", { hour: "numeric", minute: "2-digit" })}<small>${shared.length ? "Shared: " + shared.map(escapeHtml).join(", ") + "<br>" : ""}Summary with everyone's latest work: ${SUMMARY_DIR}/${LATEST_NAME}${dated ? `<br>Dated copy kept: ${SUMMARY_DIR}/${escapeHtml(dated)}` : ""}${copies.length ? `<br>Uploaded files copied to ${UPLOAD_DIR}/${escapeHtml(uploadBatch)}: ${copies.map(escapeHtml).join(", ")}` : ""}${editedCopies.length ? `<br>Edited file(s) kept in ${UPLOAD_DIR}/${escapeHtml(uploadBatch)}: ${editedCopies.map(escapeHtml).join(", ")}` : ""}${wr.written.length ? "<br>Updated original file(s): " + wr.written.map(escapeHtml).join("; ") : ""}${wr.skipped && wr.skipped.length ? "<br>Not changed (kept in the system only): " + wr.skipped.map(escapeHtml).join("; ") : ""}</small>`, false, 10000);
    saved = true;
  } finally { await releaseSaveLock(); }
}
async function finalizeWeek() {
  if (!report || saving) return;
  if (!confirm("Finalize week?\n\nThis saves everything (like Save summary) and also keeps a dated copy of the summary in Cashflow Summary as this week's record.")) return;
  saving = true; saveStep("Finalizing…");
  try { await saveSummaryInner(true); }
  catch (e) { console.error(e); toast("Not saved — something went wrong:<small>" + escapeHtml(e && e.message ? e.message : String(e)) + "</small>", true, 12000); }
  finally { saving = false; $("btn-save").innerHTML = ""; updateSaveState(); }
}

/* ============================================================
   Several people, one shared cashflow (LAN folder FILES)
   FILES/Current/ holds the latest version of every part, each saved separately:
     part-<ar|cf|ds|ln|uc|po>.json (+ <k>-file.<ext>)  source files with their edits
     part-bank.json   bank balances (Beginning Balance of Week 1)
     part-typed.json  typed amounts / added rows, cell by cell, with who typed them
   Opening the system loads all parts. Save re-reads them, writes only the parts you changed, and rebuilds
   Cashflow_Summary_LATEST.xlsx from everyone's latest work. "Finalize week" keeps a dated copy.
   ============================================================ */
const CURRENT_DIR = "Current", LATEST_NAME = "Cashflow_Summary_LATEST.xlsx", SAVE_LOCK = "_saving.json";
var SHARED_KINDS = ["ar", "cf", "ds", "ln", "uc", "po"];
const PART_LABEL = { ar: "AR Aging", cf: "Cash Flow Report", ds: "Unit Disposals", ln: "Loan Payment Schedule", uc: "Unreleased Checks", po: "Projected Cash Outflow", bank: "Bank balances", typed: "Typed amounts / added rows" };
const SESSION_ID = Math.random().toString(36).slice(2) + Date.now().toString(36);
var sharedLoading = false;      // true while parts from FILES/Current are being read (not a change by this user)
var sharedFound = false;        // FILES/Current had parts when the page connected
const sharedSeen = {};          // part -> "at" of the version on screen
var editedCopies = [], bankDirty = false, bankInfo = null, typedBase = null, typedMeta = { ov: {}, rows: {} }, pollTimer = null, pendingNews = [];
const newUid = () => "r" + Math.random().toString(36).slice(2, 10);

function userName() { try { return localStorage.getItem("cf-user") || ""; } catch (_) { return ""; } }
function setUserName(n) { try { localStorage.setItem("cf-user", n); } catch (_) {} renderUser(); }
function renderUser() { const n = userName(); $("user-label").textContent = n || "Set your name"; $("btn-user").title = n ? `Signed in as ${n} — your saves are recorded under this name. Click to change.` : "Your name (recorded with every part you save)"; }
function askUserName() {
  const n = prompt("Your name (recorded with every part you save, so others see who updated what):", userName() || "");
  if (n && n.trim()) setUserName(n.trim());
  return userName();
}
const hhmm = (iso) => new Date(iso).toLocaleTimeString("en-PH", { hour: "numeric", minute: "2-digit" });
function whenText(iso) {
  if (!iso) return "";
  const d = new Date(iso), today = new Date();
  return d.toDateString() === today.toDateString() ? `today ${hhmm(iso)}` : `${d.toLocaleDateString("en-PH", { month: "short", day: "numeric" })} ${hhmm(iso)}`;
}
function typedTip(k, iso) {
  const m = k.startsWith("c:") ? (typedMeta.rows[(customRows.find((r) => "c:" + r.id === k) || {}).uid] || {})[iso] : (typedMeta.ov[k] || {})[iso];
  const base = typedBase && (k.startsWith("c:") ? null : typedBase.overrides[k]);
  const cur = k.startsWith("c:") ? null : overrides[k] && overrides[k][iso];
  if (!m || (base && cur !== base[iso])) return "Typed by you — not saved yet";
  return `Typed by ${m.by} (${whenText(m.at)})`;
}
function renderBankInfo() {
  const el = $("bank-by");
  if (!el) return;
  el.textContent = bankDirty ? `Changed by you${bankUpload && !bankUpload.copied ? ` (from ${bankUpload.file.name})` : ""} — not saved yet (Save summary shares it with everyone).`
    : bankInfo ? `Shared balances — updated by ${bankInfo.by}, ${whenText(bankInfo.at)}${bankInfo.source ? ` (from ${bankInfo.source})` : ""}.` : "";
}

// ---- reading / writing FILES/Current
async function curDir() { return openDir(CURRENT_DIR); }
async function readJson(dir, name) { try { return JSON.parse(await (await (await dir.getFileHandle(name)).getFile()).text()); } catch (_) { return null; } }
async function writeJson(name, obj) { await writeTo(CURRENT_DIR, name, new Blob([JSON.stringify(obj)], { type: "application/json" })); }
const stamp2 = () => ({ by: userName() || "Someone", at: new Date().toISOString(), session: SESSION_ID });

// ---- typed amounts / added rows: compare with what was loaded, merge into the latest shared version
function typedNow() {
  customRows.forEach((r) => { if (!r.uid) r.uid = newUid(); });
  const uidOf = (k) => (k.startsWith("c:") ? "c:" + (customRows.find((r) => "c:" + r.id === k) || {}).uid : k);
  return {
    overrides: Object.fromEntries(Object.entries(overrides).map(([k, m]) => [k, { ...m }])),
    rows: customRows.map((r) => ({ uid: r.uid, name: r.name, sign: r.sign, vals: { ...r.vals } })),
    order: { in: rowOrder.in.map(uidOf), out: rowOrder.out.map(uidOf) },
  };
}
const emptyTyped = () => ({ overrides: {}, rows: [], order: null });
function typedChanges(now, base) {
  const ch = [];
  const keys = new Set([...Object.keys(now.overrides), ...Object.keys(base.overrides)]);
  for (const k of keys) {
    const a = now.overrides[k] || {}, b = base.overrides[k] || {};
    for (const d of new Set([...Object.keys(a), ...Object.keys(b)])) if (a[d] !== b[d]) ch.push(a[d] === undefined ? { t: "ovDel", k, d } : { t: "ov", k, d, v: a[d] });
  }
  const bRows = new Map(base.rows.map((r) => [r.uid, r])), nRows = new Map(now.rows.map((r) => [r.uid, r]));
  for (const r of now.rows) {
    const o = bRows.get(r.uid);
    if (!o) { ch.push({ t: "rowAdd", row: r }); continue; }
    if (o.name !== r.name || o.sign !== r.sign) ch.push({ t: "rowSet", uid: r.uid, name: r.name, sign: r.sign });
    for (const d of new Set([...Object.keys(r.vals), ...Object.keys(o.vals)])) if (r.vals[d] !== o.vals[d]) ch.push({ t: "cell", uid: r.uid, d, v: r.vals[d] });
  }
  for (const r of base.rows) if (!nRows.has(r.uid)) ch.push({ t: "rowDel", uid: r.uid });
  if (JSON.stringify(now.order) !== JSON.stringify(base.order)) ch.push({ t: "order", order: now.order });
  return ch;
}
function applyTypedChanges(remote, ch, who) {
  const R = remote ? JSON.parse(JSON.stringify(remote)) : { overrides: {}, rows: [], order: null };
  R.overrides = R.overrides || {}; R.rows = R.rows || [];
  const cellV = (x) => (x && typeof x === "object" ? x : { v: x });
  for (const c of ch) {
    if (c.t === "ov") { R.overrides[c.k] = R.overrides[c.k] || {}; R.overrides[c.k][c.d] = { v: c.v, by: who.by, at: who.at }; }
    else if (c.t === "ovDel") { if (R.overrides[c.k]) delete R.overrides[c.k][c.d]; }
    else if (c.t === "rowAdd") { if (!R.rows.some((r) => r.uid === c.row.uid)) R.rows.push({ uid: c.row.uid, name: c.row.name, sign: c.row.sign, by: who.by, at: who.at, vals: Object.fromEntries(Object.entries(c.row.vals).map(([d, v]) => [d, { v, by: who.by, at: who.at }])) }); }
    else if (c.t === "rowDel") { R.rows = R.rows.filter((r) => r.uid !== c.uid); }
    else if (c.t === "rowSet") { const r = R.rows.find((x) => x.uid === c.uid); if (r) Object.assign(r, { name: c.name, sign: c.sign }); }
    else if (c.t === "cell") { const r = R.rows.find((x) => x.uid === c.uid); if (r) { if (c.v === undefined) delete r.vals[c.d]; else r.vals[c.d] = { v: c.v, by: who.by, at: who.at }; } }
    else if (c.t === "order") R.order = c.order;
  }
  R.rows.forEach((r) => Object.keys(r.vals).forEach((d) => (r.vals[d] = cellV(r.vals[d]))));
  Object.values(R.overrides).forEach((m) => Object.keys(m).forEach((d) => (m[d] = cellV(m[d]))));
  return R;
}
// put a shared typed version on screen (and remember it as the base for the next comparison)
function adoptTyped(R) {
  R = R || { overrides: {}, rows: [], order: null };
  overrides = {}; typedMeta = { ov: {}, rows: {} };
  for (const [k, m] of Object.entries(R.overrides || {})) for (const [d, x] of Object.entries(m)) {
    (overrides[k] = overrides[k] || {})[d] = x.v; (typedMeta.ov[k] = typedMeta.ov[k] || {})[d] = { by: x.by, at: x.at };
  }
  customRows = (R.rows || []).map((r) => {
    typedMeta.rows[r.uid] = Object.fromEntries(Object.entries(r.vals || {}).map(([d, x]) => [d, { by: x.by, at: x.at }]));
    return { id: nextRowId++, uid: r.uid, name: r.name, sign: r.sign, vals: Object.fromEntries(Object.entries(r.vals || {}).map(([d, x]) => [d, x.v])) };
  });
  const fixed = new Set([...CASH_IN, ...CASH_OUT].map(([k]) => k)), byUid = new Map(customRows.map((r) => [r.uid, "c:" + r.id]));
  const map = (k) => (k.startsWith("c:") ? byUid.get(k.slice(2)) || null : fixed.has(k) ? k : null);
  const ord = { in: ((R.order && R.order.in) || []).map(map).filter(Boolean), out: ((R.order && R.order.out) || []).map(map).filter(Boolean) };
  CASH_IN.forEach(([k]) => { if (!ord.in.includes(k) && !ord.out.includes(k)) ord.in.push(k); });
  CASH_OUT.forEach(([k]) => { if (!ord.in.includes(k) && !ord.out.includes(k)) ord.out.push(k); });
  customRows.forEach((r) => { const k = "c:" + r.id; if (!ord.in.includes(k) && !ord.out.includes(k)) ord.in.push(k); });
  rowOrder.in = ord.in; rowOrder.out = ord.out;
  typedBase = typedNow();
}

// ---- one source part
const INGEST = { ar: (f) => handleFile("ar", f), cf: (f) => handleCashflowReport(f), ds: (f) => handleDisposal(f), ln: (f) => handleCashOut("ln", f), uc: (f) => handleCashOut("uc", f), po: (f) => handleCashOut("po", f) };
function clearPart(k) {
  const s = state[k];
  if (k === "ds") Object.assign(s, { rows: [], file: null, buffer: null, from: "" });
  else if (CO_META[k]) Object.assign(s, { byDate: {}, file: null, buffer: null, info: "", from: "", rows: [] });
  else Object.assign(s, { rows: [], file: null, buffer: null });
  resetFileEdits(k);
  const d = $("drop-" + k); if (d) d.classList.remove("loaded");
}
async function loadPart(dir, k, p) {
  const s = state[k];
  clearPart(k);
  if (!p.empty) {
    if (p.file) {
      const f = await (await dir.getFileHandle(p.file)).getFile();
      sharedLoading = true;
      try { await INGEST[k](new File([await f.arrayBuffer()], p.fileName || p.file, { type: f.type })); } finally { sharedLoading = false; }
      if (k === "ar" && p.map) s.map = p.map;
      await reattachOriginal(k, p.fileName || "");
    } else if (p.data && k === "ds") applyDisposalState(p.data, p.from || "shared");
    else if (p.data && CO_META[k]) applyCashOutState({ [k]: p.data }, p.from || "shared");
  }
  Object.assign(s, { dirty: false, removed: false, sharedBy: p.by, sharedAt: p.at, copySaved: true });
  sharedSeen[k] = p.at;
  if (k === "ds") renderDisposalMonths();
  if (CO_META[k]) renderCashOutSources();
}
async function writePart(k) {
  const s = state[k], who = stamp2(), p = { kind: k, ...who };
  if (s.removed || !hasFileData(k)) p.empty = true;
  else if (s.buffer) {
    let buf = s.buffer;
    if (editCount(k) && /\.(xlsx|xlsm)$/i.test(s.file ? s.file.name : "")) {      // put the edits into the shared copy
      s.editedSave = true;
      const wb = await excelOf(k);
      wb.calcProperties = Object.assign(wb.calcProperties || {}, { fullCalcOnLoad: true });
      const out = await wb.xlsx.writeBuffer();
      buf = out instanceof ArrayBuffer ? out : out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
      for (const [sheet, cells] of Object.entries(s.edits || {})) { s.written[sheet] = s.written[sheet] || {}; Object.keys(cells).forEach((a) => (s.written[sheet][a] = true)); }
      Object.assign(s, { buffer: buf, edits: {}, recalc: {}, ops: [] });
    }
    const name = s.file ? s.file.name : k + ".xlsx", ext = (/\.[^.]+$/.exec(name) || [".xlsx"])[0];
    p.file = `${k}-file${ext}`; p.fileName = name;
    await writeTo(CURRENT_DIR, p.file, buf);
    if (k === "ar") p.map = s.map;
  } else if (k === "ds") {
    p.data = { name: s.savedFrom || "", sheet: s.sheet || "", rows: s.rows.map((x) => ({ from: toISO(x.from), to: toISO(x.to), amount: x.amount, units: x.units, monday: toISO(x.monday) })) };
  } else if (CO_META[k]) p.data = { byDate: s.byDate, info: s.info, name: s.savedFrom || "" };
  await writeJson(`part-${k}.json`, p);
  Object.assign(s, { dirty: false, removed: false, sharedBy: p.by, sharedAt: p.at, remEdited: {} });
  sharedSeen[k] = p.at;
}

// ---- load everything others saved (parts I have changed but not saved are kept, unless force)
async function loadShared(opts = {}) {
  if (!baseDir) return false;
  const dir = await curDir();
  let found = false;
  for (const k of SHARED_KINDS) {
    const p = await readJson(dir, `part-${k}.json`);
    if (!p) continue;
    found = true;
    if (p.at === sharedSeen[k] || (state[k].dirty && !opts.force)) continue;
    showProgress(`Loading ${PART_LABEL[k]}${p.by ? ` (saved by ${p.by}, ${whenText(p.at)})` : ""}…`);
    try { await loadPart(dir, k, p); } catch (e) { toast(`Could not load the shared ${PART_LABEL[k]}:<small>${escapeHtml(e.message)}</small>`, true, 8000); }
  }
  const bank = await readJson(dir, "part-bank.json");
  if (bank) {
    found = true;
    if (bank.at !== sharedSeen.bank && (!bankDirty || opts.force)) {
      accounts.splice(0, accounts.length, ...bank.accounts.map((a) => ({ bank: a.bank || "", no: a.no || "", purpose: a.purpose || "", bal: +a.bal || 0 })));
      bankInfo = bank; bankDirty = false; sharedSeen.bank = bank.at;
      try { localStorage.setItem(ACCT_KEY, JSON.stringify(accounts)); } catch (_) {}
      renderAccounts();
    }
  }
  const typed = await readJson(dir, "part-typed.json");
  if (typed) {
    found = true;
    if (typed.at !== sharedSeen.typed) {
      const mine = typedBase ? typedChanges(typedNow(), typedBase) : [];
      adoptTyped(typed); sharedSeen.typed = typed.at;
      if (mine.length && !opts.force) { const R = applyTypedChanges(typed, mine, stamp2()); const keepBase = typedBase; adoptTyped(R); typedBase = keepBase; }   // keep my unsaved typing on top
    }
  }
  if (typedBase === null) typedBase = typedNow();
  renderBankInfo(); renderFileList(); renderAccounts();
  if (found) refresh();
  pendingNews = []; renderNews();
  return found;
}
// progress / status line at the top of the page
function showProgress(text) {
  const b = $("lock-banner");
  b.className = "lock-banner mine";
  b.innerHTML = `<span class="spin"></span> ${escapeHtml(text)}`;
}
function showNothingShared() {
  const b = $("lock-banner");
  b.className = "lock-banner mine";
  b.innerHTML = `ℹ️ <b>Nothing has been shared yet</b> in ${escapeHtml(baseDir.name)}/${CURRENT_DIR}. Upload the files once and click <b>Save summary</b> — after that, everyone who opens the system sees the latest work automatically.${report ? " Showing what was kept in the last saved summary for now." : ""}
    <button type="button" class="btn small" id="ns-upload">Upload files</button>`;
  $("ns-upload").onclick = () => { b.className = "lock-banner hidden"; openUpload(); };
}

// ---- who changed what since this page loaded it (checked every minute)
async function checkNews() {
  if (!baseDir || saving) return;
  try {
    const dir = await curDir(), news = [];
    for (const k of [...SHARED_KINDS, "bank", "typed"]) {
      const p = await readJson(dir, `part-${k}.json`);
      if (p && p.at !== sharedSeen[k] && p.session !== SESSION_ID) news.push({ k, by: p.by, at: p.at });
    }
    pendingNews = news; renderNews();
  } catch (_) {}
}
function renderNews() {
  const b = $("lock-banner");
  if (!pendingNews.length) { b.className = "lock-banner hidden"; return; }
  const who = [...new Set(pendingNews.map((x) => x.by))].join(", "), what = pendingNews.map((x) => `${PART_LABEL[x.k]} (${hhmm(x.at)})`).join(", ");
  b.className = "lock-banner other";
  b.innerHTML = `🔄 New updates from <b>${escapeHtml(who)}</b>: ${escapeHtml(what)}. <button type="button" class="btn small" id="news-load">Load them</button>`;
  $("news-load").onclick = async () => { await loadShared(); toast("Loaded the latest updates" + (SHARED_KINDS.some((k) => state[k].dirty) || bankDirty ? "<small>Parts you changed and have not saved yet were kept as yours.</small>" : "")); };
}

// ---- short lock while a save is being written (a few seconds), so two saves never mix
async function acquireSaveLock() {
  const dir = await curDir();
  for (let n = 0; n < 30; n++) {
    const l = await readJson(dir, SAVE_LOCK);
    if (!l || l.session === SESSION_ID || Date.now() - new Date(l.at).getTime() > 60000) {
      await writeJson(SAVE_LOCK, stamp2());
      const chk = await readJson(dir, SAVE_LOCK);
      if (chk && chk.session === SESSION_ID) return true;
    }
    saveStep(`Waiting for ${l ? l.by : "another save"} to finish saving…`);
    await new Promise((r) => setTimeout(r, 1000));
  }
  return false;
}
async function releaseSaveLock() { try { const dir = await curDir(), l = await readJson(dir, SAVE_LOCK); if (l && l.session === SESSION_ID) await dir.removeEntry(SAVE_LOCK); } catch (_) {} }

// ---- Save: my changed parts + everyone's latest -> Current + Cashflow_Summary_LATEST.xlsx
async function saveShared() {
  const dir = await curDir();
  // someone saved the same part after I loaded it?
  for (const k of [...SHARED_KINDS, "bank"]) {
    const mineChanged = k === "bank" ? bankDirty : state[k].dirty;
    if (!mineChanged) continue;
    const p = await readJson(dir, `part-${k}.json`);
    if (p && p.at !== sharedSeen[k] && p.session !== SESSION_ID && !confirm(`${p.by} also saved "${PART_LABEL[k]}" (${whenText(p.at)}) after you opened it.\n\nOK = save yours (replaces theirs)\nCancel = keep theirs (your changes to this part are dropped)`)) {
      if (k === "bank") bankDirty = false; else state[k].dirty = false;
      sharedSeen[k] = null;                                     // load theirs below
    }
  }
  await loadShared();                                           // everyone else's latest (my changed parts are kept)
  const done = []; editedCopies = [];
  // a part nobody has shared yet (first use of V3.6, e.g. schedules carried from the last summary) is shared as it is
  for (const k of SHARED_KINDS) if (!state[k].dirty && hasFileData(k) && !(await readJson(dir, `part-${k}.json`))) state[k].dirty = true;
  for (const k of SHARED_KINDS) if (state[k].dirty) {
    saveStep(`Sharing ${PART_LABEL[k]}…`); await writePart(k); done.push(PART_LABEL[k]);
    if (state[k].editedSave && state[k].buffer) {                // also keep the edited file in Uploaded Files (history)
      try {
        const s = state[k], nm = s.file ? s.file.name : k + ".xlsx", dot = nm.lastIndexOf(".");
        if (!uploadBatch) uploadBatch = batchFolderName();
        const cp = await uniqueName([UPLOAD_DIR, uploadBatch], `${nm.slice(0, dot)} (edited by ${userName() || "someone"})${nm.slice(dot)}`);
        await writeTo([UPLOAD_DIR, uploadBatch], cp, s.buffer);
        editedCopies.push(cp);
      } catch (e) { toast(`Saved, but the edited copy of ${escapeHtml(PART_LABEL[k])} could not be kept in ${UPLOAD_DIR}:<small>${escapeHtml(e.message)}</small>`, true, 8000); }
      state[k].editedSave = false;
    }
  }
  if (bankDirty || (acctTotal() !== 0 && !(await readJson(dir, "part-bank.json")))) {   // blank balances (e.g. the website copy) are never shared unless typed
    const p = { accounts: accounts.map((a) => ({ bank: a.bank, no: a.no, purpose: a.purpose, bal: a.bal })), source: bankUpload ? bankUpload.file.name : "", ...stamp2() };
    await writeJson("part-bank.json", p); bankInfo = p; bankDirty = false; sharedSeen.bank = p.at;
    done.push(PART_LABEL.bank);
  }
  const ch = typedChanges(typedNow(), typedBase || emptyTyped());
  if (ch.length) {
    const remote = await readJson(dir, "part-typed.json"), R = Object.assign(applyTypedChanges(remote, ch, stamp2()), stamp2());
    await writeJson("part-typed.json", R); adoptTyped(R); sharedSeen.typed = R.at;
    done.push(PART_LABEL.typed);
  }
  renderBankInfo(); renderFileList();
  return done;
}

/* ============================================================
   Init
   ============================================================ */
$("opt-start").value = "2026-10-12"; // Week 1 starts Monday, October 12, 2026 (set by the user)
$("opt-asof").value = $("opt-start").value; // aging as-of date follows the Week 1 start (editable in Setup)
for (const kind of ["ar", "ap"]) {
  const drop = $("drop-" + kind);
  $("file-" + kind).onchange = (e) => e.target.files[0] && handleFile(kind, e.target.files[0]);
  ["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => e.dataTransfer.files[0] && handleFile(kind, e.dataTransfer.files[0]));
}
["opt-start", "opt-weeks", "opt-asof", "al-allin", "al-other", "al-fixed", "bk-30", "bk-60", "bk-90", "bk-120", "bk-max", "ewt-on", "ewt-rental", "ewt-other", "ewt-vat", "ewt-exempt"]
  .forEach((id) => ($(id).onchange = () => { if (id === "opt-start") $("opt-asof").value = $("opt-start").value; saved = false; refresh(); }));
$("opt-autosave").onchange = async () => { for (const k of ["ar", "ap"]) { if (state[k].file && !state[k].saved && $("opt-autosave").checked) await saveOriginal(k); else renderFileInfo(k); } };
$("opt-company").oninput = () => { if (report) { $("rpt-company").textContent = $("opt-company").value; $("top-sub").textContent = `${$("opt-company").value} · ${$("rpt-range").textContent}`; updateSaveState(); } };
$("btn-folder").onclick = connectFolder;
$("btn-edit").onclick = toggleEdit;
$("pd-search").oninput = renderPastDue;
$("btn-add-row").onclick = () => {
  const nr = { id: nextRowId++, uid: newUid(), name: "", sign: -1, vals: {} };
  customRows.push(nr);
  rowOrder.in.push("c:" + nr.id);
  saved = false;
  if (!editing) toggleEdit(); else renderSummary();
  updateSaveState();
  const names = document.querySelectorAll(".cr-name");
  if (names.length) names[names.length - 1].focus();
};
$("btn-reset").onclick = () => { resetEdits(); renderSummary(); updateSaveState(); };
$("btn-save").onclick = saveSummary;
function showTab(name) {
  document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("active", t.dataset.tab === name));
  document.querySelectorAll(".pane").forEach((p) => p.classList.toggle("hidden", p.id !== "pane-" + name));
}
document.querySelectorAll(".tab").forEach((t) => (t.onclick = () => showTab(t.dataset.tab)));
$("btn-bal").onclick = () => { showTab("setup"); $("bal-panel").scrollIntoView({ behavior: "smooth", block: "center" }); };
$("btn-add-acct").onclick = () => { accounts.push({ bank: "", no: "", purpose: "", bal: 0 }); saveAccounts(); renderAccounts(); };
restoreAccounts();
renderAccounts();
restoreFolder();
{
  const drop = $("drop-cf");
  $("file-cf").onchange = (e) => e.target.files[0] && handleCashflowReport(e.target.files[0]);
  ["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => e.dataTransfer.files[0] && handleCashflowReport(e.dataTransfer.files[0]));
  ["rx-from", "rx-to", "rx-method"].forEach((id) => ($(id).onchange = () => { renderRetailExpand(); saved = false; refresh(); }));
  for (const key of Object.keys(CO_META)) {
    const d = $("drop-" + key);
    $("file-" + key).onchange = (e) => e.target.files[0] && handleCashOut(key, e.target.files[0]);
    ["dragenter", "dragover"].forEach((ev) => d.addEventListener(ev, (e) => { e.preventDefault(); d.classList.add("over"); }));
    ["dragleave", "drop"].forEach((ev) => d.addEventListener(ev, (e) => { e.preventDefault(); d.classList.remove("over"); }));
    d.addEventListener("drop", (e) => e.dataTransfer.files[0] && handleCashOut(key, e.dataTransfer.files[0]));
  }
  renderCashOutSources();
  {
    const dropDs = $("drop-ds");
    $("file-ds").onchange = (e) => e.target.files[0] && handleDisposal(e.target.files[0]);
    $("file-bk").onchange = (e) => { if (e.target.files[0]) handleBankFile(e.target.files[0]); e.target.value = ""; };
    const dropBk = $("drop-bk");
    ["dragenter", "dragover"].forEach((ev) => dropBk.addEventListener(ev, (e) => { e.preventDefault(); dropBk.classList.add("over"); }));
    ["dragleave", "drop"].forEach((ev) => dropBk.addEventListener(ev, (e) => { e.preventDefault(); dropBk.classList.remove("over"); }));
    dropBk.addEventListener("drop", (e) => e.dataTransfer.files[0] && handleBankFile(e.dataTransfer.files[0]));
    ["dragenter", "dragover"].forEach((ev) => dropDs.addEventListener(ev, (e) => { e.preventDefault(); dropDs.classList.add("over"); }));
    ["dragleave", "drop"].forEach((ev) => dropDs.addEventListener(ev, (e) => { e.preventDefault(); dropDs.classList.remove("over"); }));
    dropDs.addEventListener("drop", (e) => e.dataTransfer.files[0] && handleDisposal(e.dataTransfer.files[0]));
  }
  $("rt-weeks").onchange = () => { renderRetail(); saved = false; refresh(); };
  ["rf-basis", "rf-apply"].forEach((id) => ($(id).onchange = () => { saved = false; refresh(); }));
  renderRetail();
}

/* ============================================================
   Uploaded files: list in the sidebar + editor (V3.4)
   - Files picked with the system file picker (or dropped) keep a handle, so edits can be written back
     into the ORIGINAL file on "Save summary" (a backup is kept first in FILES/Uploaded Files/<date>/).
   - Files remembered from the last saved summary (no original this session) are edited as their schedule;
     those edits live in the saved summary only.
   ============================================================ */
const FILE_KINDS = [
  ["ar", "AR Aging"], ["ap", "AP Aging"], ["cf", "Cash Flow Report"], ["ds", "Unit Disposals"],
  ["ln", "Loan Payment Schedule"], ["uc", "Unreleased Checks"], ["po", "Projected Cash Outflow"],
];
const KIND_NAME = Object.fromEntries(FILE_KINDS);
const startWork = () => {};      // V3.6: no "in use" lock — everyone can work at the same time
const DISPATCH = {
  ar: (f) => (startWork(), handleFile("ar", f)), ap: (f) => (startWork(), handleFile("ap", f)), cf: (f) => (startWork(), handleCashflowReport(f)), ds: (f) => (startWork(), handleDisposal(f)),
  ln: (f) => (startWork(), handleCashOut("ln", f)), uc: (f) => (startWork(), handleCashOut("uc", f)), po: (f) => (startWork(), handleCashOut("po", f)),
};
function resetFileEdits(k) { Object.assign(state[k], { edits: {}, wbEdit: null, written: {}, backedUp: false, remEdited: {}, recalc: {}, ops: [], copySaved: false }); }
const hasOriginal = (k) => !!(state[k].buffer);
const isRemembered = (k) => !hasOriginal(k) && (k === "ds" ? state.ds.rows.length > 0 : CO_META[k] ? Object.keys(state[k].byDate).length > 0 : false);
const hasFileData = (k) => hasOriginal(k) || isRemembered(k);
function editCount(k) {
  const s = state[k];
  if (isRemembered(k)) return Object.keys(s.remEdited || {}).length;
  return Object.values(s.edits || {}).reduce((n, m) => n + Object.keys(m).length, 0) + (s.ops || []).length;
}

// sidebar list
function renderFileList() {
  const ks = FILE_KINDS.filter(([k]) => hasFileData(k));
  $("side-files").classList.toggle("hidden", !ks.length);
  $("file-list").innerHTML = ks.map(([k, label]) => {
    const n = editCount(k), rem = isRemembered(k), s = state[k];
    const fname = rem ? "Kept in the last saved summary" : s.file ? s.file.name : "";
    const by = s.dirty ? "Changed by you · not saved yet" : s.sharedBy ? `Saved by ${s.sharedBy} · ${whenText(s.sharedAt)}` : "Not shared yet — Save summary shares it";
    const where = (s.handle ? `Original on this computer: ${s.handle.name} (updated on Save)\n` : "Original not linked on this computer\n") + (s.sharedBy && !s.dirty ? `Shared copy: ${baseDir ? baseDir.name : "FILES"}/${CURRENT_DIR}/${k}-file${(/\.[^.]+$/.exec(s.file ? s.file.name : "") || [".xlsx"])[0]}` : "");
    const tip = [KIND_NAME[k] + (fname ? ": " + fname : ""), by, where, "Click to view and edit"].filter(Boolean).join("\n");
    return `<button type="button" class="file-item${s.dirty ? " fi-dirty" : ""}" data-k="${k}" title="${escapeHtml(tip)}">
      <span class="fi-name">${label}</span>${n ? `<span class="fi-badge">${n} edit${n > 1 ? "s" : ""}</span>` : ""}
      <span class="fi-sub">${escapeHtml(fname)}</span>
      <span class="fi-by">${escapeHtml(by)}</span></button>`;
  }).join("");
  $("file-list").querySelectorAll(".file-item").forEach((b) => (b.onclick = () => openEditor(b.dataset.k)));
}

// ---- picking files with a handle (system picker / drag & drop), so the original can be updated later
const XL_TYPES = [{ description: "Excel workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx", ".xlsm"], "application/vnd.ms-excel": [".xls"] } }];
function wirePickers() {
  for (const [k] of FILE_KINDS) {
    const label = document.querySelector(`label[for="file-${k}"]`), input = $("file-" + k), drop = $("drop-" + k);
    if (!label || !input || !drop) continue;
    // plain file input (no handle): reset edits; the original cannot be written back
    input.addEventListener("change", () => { resetFileEdits(k); state[k].handle = null; startWork(); }, true);
    if (window.showOpenFilePicker) {
      label.addEventListener("click", async (e) => {
        e.preventDefault();
        let hnd;
        try { [hnd] = await window.showOpenFilePicker({ types: XL_TYPES, excludeAcceptAllOption: false, multiple: false }); } catch (_) { return; }
        const file = await hnd.getFile();
        resetFileEdits(k); state[k].handle = hnd; state[k].origSeen = file.lastModified;
        rememberOriginal(k, hnd, file.lastModified);
        DISPATCH[k](file);
      });
    }
    // drag & drop: grab the handle while the drop event is still active
    drop.addEventListener("drop", (e) => {
      resetFileEdits(k); state[k].handle = null; startWork();
      const it = e.dataTransfer && e.dataTransfer.items && e.dataTransfer.items[0];
      if (it && it.getAsFileSystemHandle) it.getAsFileSystemHandle().then(async (hnd) => { if (hnd && hnd.kind === "file") { const f = await hnd.getFile(); state[k].handle = hnd; state[k].origSeen = f.lastModified; rememberOriginal(k, hnd, f.lastModified); } }).catch(() => {});
    }, true);
  }
}

// ---- the ORIGINAL file of each part, remembered on this computer (so later saves can update it, also after reopening)
const origKey = (k) => "orig-" + k;
async function rememberOriginal(k, hnd, lastModified) {
  try { await idbSet(origKey(k), { handle: hnd, name: hnd.name, lastModified: lastModified || 0 }); } catch (_) {}
}
// after a part is loaded from FILES/Current: reconnect it to the original this computer uploaded / linked (same file name)
async function reattachOriginal(k, fileName) {
  try {
    const o = await idbGet(origKey(k));
    if (o && o.handle && o.name === fileName) { state[k].handle = o.handle; state[k].origSeen = o.lastModified || 0; }
  } catch (_) {}
}
async function linkOriginal(k) {
  if (!window.showOpenFilePicker) { toast("Linking an original file needs Chrome or Edge.", true, 6000); return; }
  let hnd;
  try { [hnd] = await window.showOpenFilePicker({ types: XL_TYPES, excludeAcceptAllOption: false, multiple: false }); } catch (_) { return; }
  const s = state[k], cur = s.file ? s.file.name : "";
  if (cur && hnd.name !== cur && !confirm(`Link "${hnd.name}" as the original of ${KIND_NAME[k]}?

The system's file is named "${cur}". On the next Save summary, "${hnd.name}" is replaced with the system's version (a backup is kept first).`)) return;
  const f = await hnd.getFile();
  s.handle = hnd; s.origSeen = f.lastModified; s.syncOriginal = true;      // next Save writes the current version into it
  await rememberOriginal(k, hnd, f.lastModified);
  if (fe.k === k) openEditorSub(k);
  renderFileList(); updateSaveState();
  toast(`Linked: ${escapeHtml(hnd.name)}<small>Click Save summary to update it with the system's version (a backup is kept first). Later saves keep it up to date.</small>`, false, 8000);
}

// ---- editor state
const fe = { k: null, sheet: null, page: 0, size: 100, q: "", grid: null, filters: {}, sel: null, pageRows: [], pinned: 0, dragging: false };
function editWorkbook(k) {
  const s = state[k];
  if (!s.wbEdit) s.wbEdit = XLSX.read(s.buffer, { type: "array", cellDates: true });
  return s.wbEdit;
}
const showVal = (v) => (v == null ? "" : v instanceof Date ? toISO(v) : typeof v === "number" ? String(Math.round(v * 1e6) / 1e6) : String(v));
function sheetGrid(ws) {
  if (!ws || !ws["!ref"]) return { rows: [], r0: 0, c0: 0, cols: 0 };
  const rg = XLSX.utils.decode_range(ws["!ref"]), rows = [];
  for (let r = rg.s.r; r <= rg.e.r; r++) {
    const row = [];
    for (let cc = rg.s.c; cc <= rg.e.c; cc++) { const cell = ws[XLSX.utils.encode_cell({ r, c: cc })]; row.push(cell ? cell.v : null); }
    rows.push({ r, vals: row });
  }
  return { rows, r0: rg.s.r, c0: rg.s.c, cols: rg.e.c - rg.s.c + 1 };
}

function openEditorSub(k) {
  const rem = isRemembered(k), s = state[k];
  $("fe-sub").textContent = rem
    ? "Kept in the last saved summary (no file). Upload the file to edit all its sheets."
    : `${s.file ? s.file.name : ""} · ${s.handle ? `Save summary updates the shared copy AND the original file "${s.handle.name}" on this computer (a backup is kept first)` : "Save summary updates the shared copy. This computer is not linked to the original file — click Link original file to update it too"}`;
  const b = $("fe-link");
  b.classList.toggle("hidden", rem || !window.showOpenFilePicker);
  b.textContent = s.handle ? "Change original file…" : "Link original file…";
  b.onclick = () => linkOriginal(k);
}
function openEditor(k) {
  fe.k = k; fe.page = 0; fe.q = ""; fe.filters = {}; fe.sel = null; $("fe-search").value = ""; $("fe-pop").classList.add("hidden");
  const rem = isRemembered(k);
  $("fe-title").textContent = KIND_NAME[k];
  openEditorSub(k);
  if (rem) { fe.sheet = null; $("fe-tabs").innerHTML = ""; }
  else {
    const wb = editWorkbook(k);
    const preferred = { ar: /^detail$/i, ap: /^detail$/i, cf: /^summary$/i, ds: /weekly/i, ln: /^detail$/i, uc: /unreleased/i, po: /./ }[k];
    fe.sheet = wb.SheetNames.find((n) => preferred.test(n)) || wb.SheetNames[0];
    renderFeTabs();
  }
  renderEditor();
  $("fe-modal").classList.remove("hidden");
}
function renderFeTabs() {
  const wb = editWorkbook(fe.k);
  $("fe-tabs").innerHTML = wb.SheetNames.map((n) => {
    const e = Object.keys((state[fe.k].edits || {})[n] || {}).length;
    return `<button type="button" class="fe-tab${n === fe.sheet ? " active" : ""}" data-s="${escapeHtml(n)}">${escapeHtml(n)}${e ? ` <span>${e}</span>` : ""}</button>`;
  }).join("");
  $("fe-tabs").querySelectorAll(".fe-tab").forEach((b) => (b.onclick = () => { fe.sheet = b.dataset.s; fe.page = 0; fe.sel = null; $("fe-pop").classList.add("hidden"); renderFeTabs(); renderEditor(); }));
}
// row index (in fe.grid.rows) of the sheet's header row; rows up to it stay visible and are not filtered
function headerRowIdx(grid) {
  const rows = grid.rows.slice(0, 30).map((x) => x.vals);
  if (!rows.length) return -1;
  const { headerIdx } = findHeader(rows);
  const strings = (rows[headerIdx] || []).filter((v) => typeof v === "string" && v.trim()).length;
  return strings >= 2 ? headerIdx : -1;
}
const fkey = (v) => showVal(v);
// shown in the grid: 550000 -> 550,000 · 14352.54 -> 14,352.54 (years like 2021 in a "Year" column stay as is)
// Columns the file itself shows WITHOUT thousands commas (IDs, check / pay numbers, codes): keep them that way.
function plainCols(ws) {
  const st = {};
  for (const [a, x] of Object.entries(ws)) {
    if (a[0] === "!" || !x || x.t !== "n" || typeof x.w !== "string" || Math.abs(x.v) < 1000) continue;
    const c = XLSX.utils.decode_cell(a).c, s = (st[c] = st[c] || { comma: 0, plain: 0 });
    if (x.w.includes(",")) s.comma++; else if (/^\s*-?\d+(\.\d+)?\s*$/.test(x.w)) s.plain++;
  }
  return new Set(Object.entries(st).filter(([, s]) => s.plain && !s.comma).map(([c]) => +c));
}
function dispVal(v, yearCol, plain) {
  if (typeof v !== "number") return showVal(v);
  if (yearCol && Number.isInteger(v) && v >= 1900 && v <= 2100) return String(v);
  if (plain) return String(Math.round(v * 1e6) / 1e6);
  return v.toLocaleString("en-PH", { maximumFractionDigits: 6 });
}                                   // value as shown = value used by the filter
// Summary row fed by each file; typed amounts on that row (Edit mode / carried from the last save) hide file changes
const KIND_ROW = { ar: ["arCorp", "AR Corporate"], cf: ["retail", "Retail (Credit Card)"], ds: ["disposal", "Proceeds from Disposal"], ln: ["loans", "Loan payments"], uc: ["checks", "Unreleased checks"], po: ["projout", "Projected payments"] };
function renderOverrideWarning() {
  const box = $("fe-warn"), kr = KIND_ROW[fe.k], ov = kr && overrides[kr[0]];
  const ws = ov && report ? report.wk.filter((w) => ov[toISO(w.start)] !== undefined) : [];
  if (!ws.length) { box.classList.add("hidden"); return; }
  box.innerHTML = `⚠ The Summary row <b>${kr[1]}</b> has amounts typed by hand in ${ws.length} week(s) (${ws.map((w) => "Week " + w.n).join(", ")}). Typed amounts win over the file, so your changes here do not show in those weeks.
    <button type="button" class="btn small" id="fe-unov">Use the file values instead</button>`;
  box.classList.remove("hidden");
  $("fe-unov").onclick = () => { delete overrides[kr[0]]; saved = false; refresh(); renderOverrideWarning(); toast(`Typed amounts removed from ${kr[1]} — the Summary now follows the file.`); };
}
function renderEditor() {
  renderOverrideWarning();
  const k = fe.k, s = state[k];
  if (isRemembered(k)) return renderRememberedEditor();
  const ws = editWorkbook(k).Sheets[fe.sheet];
  fe.grid = sheetGrid(ws);
  const hi = headerRowIdx(fe.grid);
  const pinned = fe.grid.rows.slice(0, hi + 1), all = fe.grid.rows.slice(hi + 1);
  const flt = fe.filters[fe.sheet] || {}, q = fe.q.toLowerCase();
  const hit = (v) => showVal(v).toLowerCase().includes(q) || (typeof v === "number" && dispVal(v).toLowerCase().includes(q));
  const rows = all.filter((x) => (!q || x.vals.some(hit))
    && Object.entries(flt).every(([col, set]) => set.has(fkey(x.vals[+col - fe.grid.c0]))));
  const pages = Math.max(1, Math.ceil(rows.length / fe.size));
  fe.page = Math.min(fe.page, pages - 1);
  const part = rows.slice(fe.page * fe.size, (fe.page + 1) * fe.size);
  fe.pinned = pinned.length;
  fe.pageRows = [...pinned, ...part];
  const nf = Object.keys(flt).length;
  $("fe-range").textContent = (rows.length ? `Rows ${fe.page * fe.size + 1}–${fe.page * fe.size + part.length} of ${rows.length.toLocaleString("en-PH")}` : "No rows")
    + (nf || q ? ` (filtered from ${all.length.toLocaleString("en-PH")})` : "");
  $("fe-prev").disabled = fe.page === 0; $("fe-next").disabled = fe.page >= pages - 1;
  $("fe-clearf").classList.toggle("hidden", !nf);
  const ed = (s.edits || {})[fe.sheet] || {}, wr = (s.written || {})[fe.sheet] || {};
  const colHead = Array.from({ length: fe.grid.cols }, (_, i) => {
    const col = fe.grid.c0 + i;
    return `<th data-col="${col}"><span>${XLSX.utils.encode_col(col)}</span><button type="button" class="fe-fb${flt[col] ? " on" : ""}" data-fcol="${col}" title="Filter this column">▾</button></th>`;
  }).join("");
  const hdrRow = hi >= 0 ? fe.grid.rows[hi].vals : [];
  const yearCols = new Set(hdrRow.map((h, ci) => (typeof h === "string" && /year/i.test(h) ? fe.grid.c0 + ci : -1)).filter((x) => x >= 0));
  const plain = plainCols(ws);
  const body = fe.pageRows.map((x, i) => `<tr class="${i < fe.pinned ? "pin" : ""}"><th class="rn" data-ri="${i}">${x.r + 1}</th>${x.vals.map((v, ci) => {
    const col = fe.grid.c0 + ci, addr = XLSX.utils.encode_cell({ r: x.r, c: col }), cls = addr in ed ? " edited" : addr in wr ? " written" : "";
    const fcell = ws[addr] && ws[addr].f;
    return `<td class="fe-c${cls}${typeof v === "number" ? " num" : ""}${fcell ? " fx" : ""}" data-a="${addr}" data-i="${i}" data-c="${col}"${fcell ? ` title="=${escapeHtml(ws[addr].f)}"` : ""}>${escapeHtml(dispVal(v, yearCols.has(col), plain.has(col)))}</td>`;
  }).join("")}</tr>`).join("");
  $("fe-grid").innerHTML = `<table><thead><tr><th class="rn"></th>${colHead}</tr></thead><tbody>${body}</tbody></table>`;
  if (fe.sel && (Math.max(fe.sel.a[0], fe.sel.b[0]) >= fe.pageRows.length)) fe.sel = null;
  paintSel();
  $("fe-note").textContent = `${editCount(k)} change(s)${(s.ops || []).length ? ` (incl. ${s.ops.length} inserted row/column)` : ""}. The Summary updates right away${s.handle ? "; the original file is updated on Save summary." : "."}`;
  $("fe-undo").disabled = !editCount(k);
  $("fe-find").disabled = false; $("fe-addrow").disabled = false; $("fe-addcol").disabled = false;
}

// ---- range selection
const selBox = () => fe.sel && { i1: Math.min(fe.sel.a[0], fe.sel.b[0]), i2: Math.max(fe.sel.a[0], fe.sel.b[0]), c1: Math.min(fe.sel.a[1], fe.sel.b[1]), c2: Math.max(fe.sel.a[1], fe.sel.b[1]) };
function paintSel() {
  const bx = selBox();
  $("fe-grid").querySelectorAll("td.fe-c[data-a]").forEach((td) => {
    const i = +td.dataset.i, col = +td.dataset.c;
    const on = !!bx && i >= bx.i1 && i <= bx.i2 && col >= bx.c1 && col <= bx.c2;
    td.classList.toggle("sel", on);
    td.classList.toggle("anchor", on && fe.sel.a[0] === i && fe.sel.a[1] === col);
  });
  const n = bx ? (bx.i2 - bx.i1 + 1) * (bx.c2 - bx.c1 + 1) : 0;
  if (n > 1) $("fe-note").textContent = `${n.toLocaleString("en-PH")} cells selected — type a value and press Enter to fill them all, or Delete to clear them.`;
}
function selectedAddrs() {
  const bx = selBox(); if (!bx) return [];
  const out = [];
  for (let i = bx.i1; i <= bx.i2; i++) for (let col = bx.c1; col <= bx.c2; col++) out.push(XLSX.utils.encode_cell({ r: fe.pageRows[i].r, c: col }));
  return out;
}
const anchorTd = () => fe.sel && $("fe-grid").querySelector(`td.fe-c[data-i="${fe.sel.a[0]}"][data-c="${fe.sel.a[1]}"]`);

// edit: type into the active cell; Enter applies to every selected cell (or just the cell when one is selected)
function beginCellEdit(td, initial) {
  if (!td || td.querySelector("input")) return;
  const addr = td.dataset.a, ws = editWorkbook(fe.k).Sheets[fe.sheet], cell = ws[addr];
  const many = selectedAddrs().length > 1;
  const inp = document.createElement("input");
  inp.value = initial != null ? initial : cell ? (cell.f ? "=" + cell.f : showVal(cell.v)) : ""; inp.className = "fe-in";
  td.textContent = ""; td.appendChild(inp); inp.focus();
  if (initial == null) inp.select(); else inp.setSelectionRange(inp.value.length, inp.value.length);
  let done = false;
  const finish = (keep) => {
    if (done) return; done = true;
    if (!keep) { renderEditor(); $("fe-grid").focus(); return; }
    if (many) applyToSelection(inp.value); else commitCell(addr, inp.value, cell);
    $("fe-grid").focus();
  };
  inp.onkeydown = (e) => { e.stopPropagation(); if (e.key === "Enter" || e.key === "Tab") { e.preventDefault(); finish(true); } if (e.key === "Escape") { e.preventDefault(); finish(false); } };
  inp.onblur = () => finish(true);
}
function parseTyped(text, old) {
  const t = text.trim();
  if (t === "") return null;
  if (t[0] === "=" && t.length > 1) return { formula: t.slice(1) };            // a formula, e.g. =SUM(B6:B29)
  if ((old && old.t === "d") || /^\d{4}-\d{2}-\d{2}$/.test(t)) { const d = parseDate(t); if (d) return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  const n = Number(t.replace(/,/g, ""));
  if (((old && old.t === "n") || /^-?[\d,]*\.?\d+$/.test(t)) && !isNaN(n)) return n;
  return t;
}
// set one cell (no re-render); returns true when the value changed
// set one cell (no re-render); returns true when the value changed.
// opts.keepFormulas: skip cells that hold a formula (multi-cell fill / delete) — returns "skip"
function applyCellValue(addr, text, old, opts = {}) {
  const s = state[fe.k], ws = editWorkbook(fe.k).Sheets[fe.sheet];
  const val = parseTyped(text, old), isF = val && typeof val === "object" && !(val instanceof Date) && val.formula;
  if (old && old.f && !isF) {
    if (opts.keepFormulas) return "skip";
    if (!confirm(`${addr} holds a formula (=${old.f}).\n\nOK = replace the formula with the fixed value "${text}"\nCancel = keep the formula`)) return false;
  }
  if (isF) {
    if (old && old.f && old.f.replace(/\s/g, "") === val.formula.replace(/\s/g, "")) return false;
    let v = old ? old.v : 0;
    try { v = evalSimpleFormula(val.formula, ws, editWorkbook(fe.k)); } catch (_) { /* Excel will calculate it */ }
    s.edits = s.edits || {}; s.edits[fe.sheet] = s.edits[fe.sheet] || {};
    s.edits[fe.sheet][addr] = { formula: val.formula };
    ws[addr] = { t: "n", f: val.formula, v };
  } else {
    const before = old ? old.v : null;
    const same = (x, y) => (x instanceof Date && y instanceof Date ? toISO(x) === toISO(y) : x === y);
    if (same(val, before) && !(old && old.f)) return false;
    s.edits = s.edits || {}; s.edits[fe.sheet] = s.edits[fe.sheet] || {};
    s.edits[fe.sheet][addr] = val;
    if (val == null) delete ws[addr];
    else ws[addr] = val instanceof Date ? { t: "d", v: val } : typeof val === "number" ? { t: "n", v: val } : { t: "s", v: val };
  }
  const rg = XLSX.utils.decode_range(ws["!ref"] || "A1"), cc = XLSX.utils.decode_cell(addr);
  if (cc.r > rg.e.r || cc.c > rg.e.c) { rg.e.r = Math.max(rg.e.r, cc.r); rg.e.c = Math.max(rg.e.c, cc.c); ws["!ref"] = XLSX.utils.encode_range(rg); }
  return true;
}
function commitCell(addr, text, old) {
  if (applyCellValue(addr, text, old) === true) afterFileEdit(fe.k); else renderEditor();
}
// fill / clear every selected cell — totals and other formula cells are left alone (they recalculate)
function applyToSelection(text) {
  const ws = editWorkbook(fe.k).Sheets[fe.sheet];
  let n = 0, kept = 0;
  const formula = text.trim()[0] === "=";
  for (const addr of selectedAddrs()) {
    const r = applyCellValue(addr, text, ws[addr], { keepFormulas: !formula });
    if (r === true) n++; else if (r === "skip") kept++;
  }
  if (n) afterFileEdit(fe.k); else renderEditor();
  if (kept) toast(`${kept} formula cell${kept > 1 ? "s were" : " was"} kept (e.g. totals with =SUM) — they recalculate from the new amounts.`, false, 5000);
}

/* ---- insert rows / columns (like Excel's Insert) ----
   The new line copies the previous line's format (on save) and its formulas (shifted); every formula in the
   workbook that points below / right of it is shifted, and totals right below / right of a range grow to include it. */
const colNum = (L) => L.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colLet = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const REF_RE = /((?:'[^']+'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\w(!])/g;
// apply fn(sheetName|null, a, b|null) -> [a', b'] | null to every A1 reference outside "string literals"
function mapRefs(f, fn) {
  return String(f).split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(REF_RE, (m, pre, ca, c1, ra, r1, cb, c2, rb, r2, off, whole) => {
    const prev = whole[off - 1];
    if (!pre && prev && /[\w.$]/.test(prev)) return m;                       // part of a longer name
    const sn = pre ? pre.slice(0, -1).replace(/^'|'$/g, "").replace(/''/g, "'") : null;
    const a = { col: colNum(c1), row: +r1, ca: !!ca, ra: !!ra }, b = c2 ? { col: colNum(c2), row: +r2, ca: !!cb, ra: !!rb } : null;
    const res = fn(sn, a, b);
    if (!res) return m;
    const w = (x) => `${x.ca ? "$" : ""}${colLet(x.col)}${x.ra ? "$" : ""}${x.row}`;
    return (pre || "") + w(res[0]) + (res[1] ? ":" + w(res[1]) : "");
  }))).join("");
}
// a line was inserted at 0-based index `at` (axis "r" = row, "c" = column) in sheet `tgt`;
// the formula sits in sheet `own` at 0-based index `pos` on that axis (after the insert)
function shiftRefs(f, own, tgt, axis, at, pos) {
  const k = axis === "r" ? "row" : "col", A = at + 1;
  return mapRefs(f, (sn, a, b) => {
    if ((sn || own) !== tgt) return null;
    const x = { ...a }, y = b ? { ...b } : null;
    if (!y) { if (x[k] >= A) x[k]++; return [x]; }
    const lo = Math.min(x[k], y[k]), hi = Math.max(x[k], y[k]), up = (z) => z[k]++;
    if (lo >= A) { up(x); up(y); }
    else if (hi >= A) up(y[k] >= x[k] ? y : x);
    else if (hi === A - 1 && own === tgt && pos >= A) up(y[k] >= x[k] ? y : x);   // a total right after the range: include the new line
    return [x, y];
  });
}
// copy a formula to a cell dr rows / dc columns away (relative parts move, $absolute parts stay)
function moveFormula(f, dr, dc) {
  return mapRefs(f, (sn, a, b) => { const t = (x) => ({ ...x, row: x.ra ? x.row : x.row + dr, col: x.ca ? x.col : x.col + dc }); return b ? [t(a), t(b)] : [t(a)]; });
}
function insertLine(k, sheet, axis, at) {
  const s = state[k], wb = editWorkbook(k), ws = wb.Sheets[sheet], key = axis === "r" ? "r" : "c";
  const dec = (a) => XLSX.utils.decode_cell(a), enc = (o) => XLSX.utils.encode_cell(o);
  const bump = (o) => (o[key] >= at ? { ...o, [key]: o[key] + 1 } : o);
  // 1. move the cells
  Object.keys(ws).filter((a) => a[0] !== "!").map((a) => [a, dec(a)]).filter(([, o]) => o[key] >= at)
    .sort((x, y) => y[1][key] - x[1][key]).forEach(([a, o]) => { ws[enc(bump(o))] = ws[a]; delete ws[a]; });
  const rg = XLSX.utils.decode_range(ws["!ref"] || "A1"); rg.e[key]++; ws["!ref"] = XLSX.utils.encode_range(rg);
  (ws["!merges"] || []).forEach((m) => { if (m.s[key] >= at) { m.s[key]++; m.e[key]++; } else if (m.e[key] >= at) m.e[key]++; });
  // 2. shift every formula in the workbook (and formulas typed in this session)
  for (const sn of wb.SheetNames) {
    const w2 = wb.Sheets[sn];
    for (const a of Object.keys(w2)) if (a[0] !== "!" && w2[a] && w2[a].f) w2[a].f = shiftRefs(w2[a].f, sn, sheet, axis, at, dec(a)[key]);
  }
  const remap = (obj) => { const out = {}; for (const [a, v] of Object.entries(obj || {})) out[enc(bump(dec(a)))] = v; return out; };
  s.edits[sheet] = remap(s.edits[sheet]); s.written[sheet] = remap(s.written[sheet]);
  if (s.recalc && s.recalc[sheet]) s.recalc[sheet] = new Set([...s.recalc[sheet]].map((a) => enc(bump(dec(a)))));
  for (const [sn, cells] of Object.entries(s.edits)) for (const [a, v] of Object.entries(cells))
    if (v && typeof v === "object" && !(v instanceof Date) && v.formula) cells[a] = { formula: shiftRefs(v.formula, sn, sheet, axis, at, dec(a)[key]) };
  // 3. the new line takes the previous line's formulas (shifted by one), e.g. its own =SUM(B17:D17) month total
  s.edits[sheet] = s.edits[sheet] || {};
  const other = axis === "r" ? "c" : "r";
  for (let i = rg.s[other]; i <= rg.e[other]; i++) {
    const prevAddr = enc(axis === "r" ? { r: at - 1, c: i } : { r: i, c: at - 1 }), prev = ws[prevAddr];
    if (!prev || !prev.f) continue;
    const f = moveFormula(prev.f, axis === "r" ? 1 : 0, axis === "r" ? 0 : 1), na = enc(axis === "r" ? { r: at, c: i } : { r: i, c: at });
    ws[na] = { t: "n", f, v: 0 };
    s.edits[sheet][na] = { formula: f };
  }
  s.ops.push({ axis, sheet, at });
}
function addLine(axis) {
  if (isRemembered(fe.k)) return;
  if (!fe.sel) { toast("Select a cell first — the new " + (axis === "r" ? "row goes below it" : "column goes right of it") + " and copies its format.", true, 5000); return; }
  const r = fe.pageRows[fe.sel.a[0]].r, c = fe.sel.a[1];
  insertLine(fe.k, fe.sheet, axis, axis === "r" ? r + 1 : c + 1);
  fe.sel = null;
  afterFileEdit(fe.k);
  toast(axis === "r" ? `Row ${r + 2} inserted below row ${r + 1}` : `Column ${colLet(c + 2)} inserted right of column ${colLet(c + 1)}`,
    false, 5000);
}

// ---- Find & Replace (Ctrl+H): e.g. every "BPI Tokyo" -> "BPI Philippines"
// rows of a sheet that pass the current filters + search box (all pages)
function shownRows(sheet) {
  const g = sheetGrid(editWorkbook(fe.k).Sheets[sheet]), hi = headerRowIdx(g), flt = (fe.filters[sheet] || {}), q = fe.q.toLowerCase();
  return g.rows.slice(hi + 1).filter((x) => (!q || x.vals.some((v) => showVal(v).toLowerCase().includes(q)))
    && Object.entries(flt).every(([col, set]) => set.has(fkey(x.vals[+col - g.c0])))).map((x) => x.r);
}
function findMatches(o) {
  const wb = editWorkbook(fe.k), out = [];
  if (!o.find) return out;
  const esc = o.find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(o.whole ? `^${esc}$` : esc, o.matchCase ? "g" : "gi");
  const sheets = o.scope === "all" ? wb.SheetNames : [fe.sheet];
  for (const sh of sheets) {
    const ws = wb.Sheets[sh];
    if (!ws || !ws["!ref"]) continue;
    const only = o.scope === "shown" ? new Set(shownRows(sh)) : null;
    for (const addr of Object.keys(ws)) {
      if (addr[0] === "!") continue;
      const cl = ws[addr];
      if (!cl || typeof cl.v !== "string") continue;                    // text cells only (names, remarks, …)
      if (only && !only.has(XLSX.utils.decode_cell(addr).r)) continue;
      re.lastIndex = 0;
      if (!re.test(cl.v)) continue;
      re.lastIndex = 0;
      out.push({ sheet: sh, addr, old: cl.v, now: cl.v.replace(re, o.repl) });
    }
  }
  return out;
}
function openReplace() {
  if (isRemembered(fe.k)) return;
  const pop = $("fe-pop");
  pop.innerHTML = `<div class="fp-head"><b>Find &amp; Replace</b><span class="muted">${escapeHtml(KIND_NAME[fe.k])}</span></div>
    <label class="fr-l">Find what<input type="text" class="fr-find" placeholder="e.g. BPI Tokyo"></label>
    <label class="fr-l">Replace with<input type="text" class="fr-repl" placeholder="e.g. BPI Philippines"></label>
    <label class="fp-all"><input type="checkbox" class="fr-whole"> Match entire cell contents</label>
    <label class="fp-all"><input type="checkbox" class="fr-case"> Match case</label>
    <label class="fr-l">Look in
      <select class="fr-scope"><option value="sheet">This sheet (${escapeHtml(fe.sheet)})</option><option value="shown">Only the rows shown (filters / search)</option><option value="all">All sheets of the file</option></select></label>
    <div class="fr-res muted">Type what to find.</div>
    <div class="fp-foot"><button type="button" class="btn ghost small fr-close">Close</button><button type="button" class="btn primary small fr-go" disabled>Replace all</button></div>`;
  pop.style.left = "auto"; pop.style.right = "24px"; pop.style.top = "120px";
  pop.classList.remove("hidden");
  const opts = () => ({ find: pop.querySelector(".fr-find").value, repl: pop.querySelector(".fr-repl").value, whole: pop.querySelector(".fr-whole").checked, matchCase: pop.querySelector(".fr-case").checked, scope: pop.querySelector(".fr-scope").value });
  const preview = () => {
    const m = findMatches(opts()), res = pop.querySelector(".fr-res");
    pop.querySelector(".fr-go").disabled = !m.length;
    if (!opts().find) { res.textContent = "Type what to find."; return; }
    if (!m.length) { res.textContent = "No cells match."; return; }
    const ex = m.slice(0, 4).map((x) => `<div class="fr-ex"><span>${escapeHtml(x.sheet)}!${x.addr}</span> ${escapeHtml(x.old)} → <b>${escapeHtml(x.now)}</b></div>`).join("");
    res.innerHTML = `<b>${m.length.toLocaleString("en-PH")} cell${m.length > 1 ? "s" : ""}</b> will change${m.length > 4 ? " — for example:" : ":"}${ex}`;
  };
  pop.querySelectorAll("input, select").forEach((el) => (el.oninput = el.onchange = preview));
  pop.querySelector(".fr-close").onclick = () => pop.classList.add("hidden");
  pop.querySelector(".fr-go").onclick = () => {
    const m = findMatches(opts());
    if (!m.length) return;
    const keep = fe.sheet, ws = editWorkbook(fe.k).Sheets;
    let n = 0;
    for (const x of m) { fe.sheet = x.sheet; if (applyCellValue(x.addr, x.now, ws[x.sheet][x.addr])) n++; }
    fe.sheet = keep;
    pop.classList.add("hidden");
    if (n) afterFileEdit(fe.k); else renderEditor();
    toast(`Replaced in ${n.toLocaleString("en-PH")} cell${n === 1 ? "" : "s"}<small>Edited cells are yellow; the Summary is updated. Use "Undo all edits in this file" to go back.</small>`, false, 6000);
  };
  pop.querySelector(".fr-find").focus();
}

// ---- Excel-style column filter
function openFilter(col, btn) {
  const pop = $("fe-pop"), flt = fe.filters[fe.sheet] || {}, ci = col - fe.grid.c0;
  const hi = headerRowIdx(fe.grid), q = fe.q.toLowerCase();
  // values from rows that pass the OTHER filters and the search box (like Excel)
  const base = fe.grid.rows.slice(hi + 1).filter((x) => (!q || x.vals.some((v) => showVal(v).toLowerCase().includes(q)))
    && Object.entries(flt).every(([c2, set]) => +c2 === col || set.has(fkey(x.vals[+c2 - fe.grid.c0]))));
  const counts = new Map();
  base.forEach((x) => { const v = fkey(x.vals[ci]); counts.set(v, (counts.get(v) || 0) + 1); });
  const vals = [...counts.keys()].sort((a, b) => (a === "" ? 1 : b === "" ? -1 : a.localeCompare(b, undefined, { numeric: true })));
  const cur = flt[col], plain = plainCols(editWorkbook(fe.k).Sheets[fe.sheet]).has(col);
  pop.innerHTML = `<div class="fp-head"><b>Filter column ${XLSX.utils.encode_col(col)}</b><span class="muted">${vals.length.toLocaleString("en-PH")} values</span></div>
    <input type="text" class="fp-search" placeholder="Search values…">
    <label class="fp-all"><input type="checkbox" class="fp-allbox" ${!cur ? "checked" : ""}> (Select all)</label>
    <div class="fp-list">${vals.slice(0, 2000).map((v, i) => `<label class="fp-item" data-t="${escapeHtml(v.toLowerCase())}"><input type="checkbox" data-v="${i}" ${!cur || cur.has(v) ? "checked" : ""}> ${v === "" ? "<i>(Blanks)</i>" : escapeHtml(/^-?\d+(\.\d+)?$/.test(v) ? dispVal(+v, false, plain) : v)} <span class="muted">${counts.get(v)}</span></label>`).join("")}</div>
    <div class="fp-foot"><button type="button" class="btn ghost small fp-clear">Clear filter</button><button type="button" class="btn ghost small fp-cancel">Cancel</button><button type="button" class="btn primary small fp-ok">OK</button></div>`;
  const r = btn.getBoundingClientRect(), m = pop.parentElement.getBoundingClientRect();
  pop.style.left = Math.max(8, Math.min(r.left - m.left - 220, m.width - 330)) + "px";
  pop.style.top = (r.bottom - m.top + 4) + "px";
  pop.classList.remove("hidden");
  const boxes = () => [...pop.querySelectorAll(".fp-item input")];
  pop.querySelector(".fp-search").oninput = (e) => { const t = e.target.value.toLowerCase(); pop.querySelectorAll(".fp-item").forEach((l) => l.classList.toggle("hidden", !!t && !l.dataset.t.includes(t))); };
  pop.querySelector(".fp-allbox").onchange = (e) => boxes().forEach((b) => { if (!b.parentElement.classList.contains("hidden")) b.checked = e.target.checked; });
  const close = () => pop.classList.add("hidden");
  pop.querySelector(".fp-cancel").onclick = close;
  pop.querySelector(".fp-clear").onclick = () => { delete flt[col]; fe.filters[fe.sheet] = flt; close(); fe.page = 0; fe.sel = null; renderEditor(); };
  pop.querySelector(".fp-ok").onclick = () => {
    const keep = new Set(boxes().filter((b) => b.checked).map((b) => vals[+b.dataset.v]));
    if (keep.size === vals.length && !cur) delete flt[col];
    else if (keep.size === vals.length && cur) { delete flt[col]; }
    else flt[col] = keep;
    fe.filters[fe.sheet] = flt; close(); fe.page = 0; fe.sel = null; renderEditor();
  };
  pop.querySelector(".fp-search").focus();
}

let reingestTimer = null;
function afterFileEdit(k) {
  state[k].dirty = true;
  renderFeTabs(); renderEditor(); renderFileList();
  saved = false;
  clearTimeout(reingestTimer);
  reingestTimer = setTimeout(() => reingest(k), 150);
}
// re-read the edited workbook so the Summary follows the edit
/* ---- recalculation after an edit (the browser has no Excel engine) ----
   1. simple formulas on edited sheets: numbers, cell references, + - * /, ( ), SUM(...)   e.g. =SUM(B6:B29), =E2+F2-G2
   2. the Unit Disposal plan's own rules (ready-to-sell date + lag, then weekly SUMIFS / COUNTIFS)
   Anything else (IF, SUMIFS, other sheets, dates/text in arithmetic) keeps its last value. */
// Small formula engine for the file editor: numbers, cell / range references (also on other sheets and whole
// columns like 'Unreleased Checks'!$D:$D), + - * / &, SUM, SUMIFS and COUNTIFS. Anything else throws, and that
// formula keeps its last value (Excel recalculates it when the file is opened).
const XL_EPOCH = Date.UTC(1899, 11, 30);
const serialOf = (d) => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - XL_EPOCH) / DAY;
function evalSimpleFormula(f, ws, wb, cache) {
  const src = String(f);
  let i = 0;
  const sp = () => { while (src[i] === " ") i++; };
  const sheetOf = (name) => { if (name == null) return ws; const s = wb && wb.Sheets[name]; if (!s) throw 0; return s; };
  const cellV = (sh, addr) => { const cl = sh[addr]; return !cl || cl.v == null ? null : cl.v; };
  const num = (v) => { if (v == null || v === "") return 0; if (typeof v === "number") return v; if (typeof v === "boolean") return +v; throw 0; };
  // reference at the cursor: [sheet!]A1, [sheet!]A1:B9 or [sheet!]D:D
  const REF = /^(?:'((?:[^']|'')+)'!|([A-Za-z0-9_.]+)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![\w(])/i;
  function readRef() {
    const m = REF.exec(src.slice(i));
    if (!m) return null;
    i += m[0].length;
    const sh = sheetOf(m[1] != null ? m[1].replace(/''/g, "'") : m[2] != null ? m[2] : null), ref = m[3].replace(/\$/g, "").toUpperCase();
    if (!ref.includes(":")) return { sh, cell: ref };
    let [p, q] = ref.split(":");
    if (/^[A-Z]+$/.test(p)) {                                     // whole column: used rows of that sheet
      const rg = XLSX.utils.decode_range(sh["!ref"] || "A1");
      p += rg.s.r + 1; q += rg.e.r + 1;
    }
    return { sh, range: XLSX.utils.decode_range(p + ":" + q) };
  }
  const cells = (r) => {
    let m = null, key = "";
    if (cache) { m = cache.get(r.sh); if (!m) cache.set(r.sh, (m = new Map())); key = r.range.s.r + ":" + r.range.s.c + ":" + r.range.e.r + ":" + r.range.e.c; const hit = m.get(key); if (hit) return hit; }
    const out = [], cols = [];
    for (let C = r.range.s.c; C <= r.range.e.c; C++) cols.push(XLSX.utils.encode_col(C));
    for (let R = r.range.s.r; R <= r.range.e.r; R++) { const rn = R + 1; for (const cl of cols) { const x = r.sh[cl + rn]; out.push(!x || x.v == null ? null : x.v); } }
    if (m) m.set(key, out);
    return out;
  };
  const args = () => {                                           // after "NAME(": list of refs / values
    const out = [];
    for (;;) {
      sp();
      const save = i, r = readRef();
      sp();
      if (r && (src[i] === "," || src[i] === ")")) out.push(r); else { i = save; out.push({ val: concat() }); }
      sp();
      if (src[i] === ",") { i++; continue; }
      if (src[i] === ")") { i++; return out; }
      throw 0;
    }
  };
  // Excel criteria: ">=46307", "<>0", "AGENCY", 5, a date ...
  const crit = (c) => {
    let op = "=", s = c;
    if (typeof c === "string") { const m = /^(>=|<=|<>|>|<|=)?(.*)$/s.exec(c); op = m[1] || "="; s = m[2]; }
    const n = c instanceof Date ? serialOf(c) : typeof s === "number" ? s : s !== "" && !isNaN(+s) ? +s : null;
    const re = typeof s === "string" && n == null ? new RegExp("^" + s.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i") : null;
    return (v) => {
      const x = v instanceof Date ? serialOf(v) : v;
      if (n != null) {
        if (typeof x !== "number") return op === "<>";
        return op === "=" ? x === n : op === "<>" ? x !== n : op === ">" ? x > n : op === "<" ? x < n : op === ">=" ? x >= n : x <= n;
      }
      const t = x == null ? "" : String(x), hit = s === "" ? t === "" : re.test(t);
      return op === "<>" ? !hit : op === "=" ? hit : false;
    };
  };
  const ifs = (list, sumR) => {                                   // [rangeRef, criterion] pairs
    if (list.length % 2 || list.some((x, k) => (k % 2 === 0 ? !x.range : false))) throw 0;
    const cols = [];
    for (let k = 0; k < list.length; k += 2) cols.push([cells(list[k]), crit(list[k + 1].val !== undefined ? list[k + 1].val : cellV(list[k + 1].sh, list[k + 1].cell))]);
    const vals = sumR ? cells(sumR) : null, n = cols[0][0].length;
    if (cols.some(([v]) => v.length !== n) || (vals && vals.length !== n)) throw 0;
    let s = 0;
    for (let r = 0; r < n; r++) if (cols.every(([v, ok]) => ok(v[r]))) s += vals ? (typeof vals[r] === "number" ? vals[r] : 0) : 1;
    return s;
  };
  function primary() {
    sp();
    const rest = src.slice(i);
    if (src[i] === "(") { i++; const v = concat(); sp(); if (src[i] !== ")") throw 0; i++; return v; }
    if (src[i] === "-") { i++; return -num(primary()); }
    if (src[i] === "+") { i++; return primary(); }
    if (src[i] === '"') { const m = /^"((?:[^"]|"")*)"/.exec(rest); if (!m) throw 0; i += m[0].length; return m[1].replace(/""/g, '"'); }
    let m = /^\d+(\.\d+)?([eE][-+]?\d+)?/.exec(rest);
    if (m && !/^\d+:/.test(rest)) { i += m[0].length; return parseFloat(m[0]); }
    m = /^(SUM|SUMIFS|COUNTIFS)\(/i.exec(rest);
    if (m) {
      i += m[0].length;
      const fn = m[1].toUpperCase(), a = args();
      if (fn === "SUM") return a.reduce((s, x) => s + (x.range ? cells(x).reduce((t, v) => t + (typeof v === "number" ? v : 0), 0) : x.cell ? num(cellV(x.sh, x.cell)) : num(x.val)), 0);
      if (fn === "SUMIFS") { if (!a[0].range) throw 0; return ifs(a.slice(1), a[0]); }
      return ifs(a, null);
    }
    const r = readRef();
    if (r && r.cell) return cellV(r.sh, r.cell);
    throw 0;
  }
  const valOf = (v) => (v instanceof Date ? serialOf(v) : v);
  function term() { let v = primary(); for (;;) { sp(); if (src[i] === "*") { i++; v = num(v) * num(primary()); } else if (src[i] === "/") { i++; v = num(v) / num(primary()); } else return v; } }
  function expr() { let v = term(); for (;;) { sp(); if (src[i] === "+") { i++; v = num(v) + num(term()); } else if (src[i] === "-") { i++; v = num(v) - num(term()); } else return v; } }
  function concat() { let v = expr(); for (;;) { sp(); if (src[i] === "&") { i++; const w = expr(); v = String(valOf(v) ?? "") + String(valOf(w) ?? ""); } else return v; } }
  const v = concat(); sp();
  if (i < src.length) throw 0;
  const out = v == null ? 0 : v;
  if (typeof out !== "number" || !isFinite(out)) throw 0;          // dates / text results: leave the cell as it is
  return out;
}
function recalcSheet(ws, track, wb) {
  if (!ws) return false;
  const fs = Object.keys(ws).filter((a) => a[0] !== "!" && ws[a] && ws[a].f);
  let any = false;
  for (let pass = 0; pass < 8; pass++) {                      // a few passes so formulas that use other formulas settle
    let changed = false;
    for (const a of fs) {
      let v;
      try { v = evalSimpleFormula(ws[a].f, ws, wb); } catch (_) { continue; }
      if (Math.abs((typeof ws[a].v === "number" ? ws[a].v : NaN) - v) > 1e-9 || typeof ws[a].v !== "number") { ws[a].v = v; ws[a].t = "n"; delete ws[a].w; changed = true; any = true; if (track) track.add(a); }
    }
    if (!changed) break;
  }
  return any;
}
// Unit Disposal plan: Expected cash-in = MAX(projected month, orig ORCR date, plan date) + lag (14 if ORCR pending, else 7);
// Weekly "Disposal proceeds" = SUMIFS(Proceeds, cash-in within From–To); "Units — cash received" = COUNTIFS(...).
function recalcUnitPlan(wb, edits, onPut) {
  const nm = (re) => wb.SheetNames.find((n) => re.test(n));
  const schN = nm(/disposal schedule/i), wkN = nm(/weekly/i), asN = nm(/assumption/i);
  if (!schN || !wkN) return;
  const sch = wb.Sheets[schN], wk = wb.Sheets[wkN], as = asN ? wb.Sheets[asN] : null;
  const isT = (c, re) => typeof c === "string" && re.test(c.trim());
  const userSet = (sheet, addr) => !!(edits && edits[sheet] && addr in edits[sheet]);
  const put = (sheet, ws, addr, v) => {
    if (userSet(sheet, addr)) return;
    const old = ws[addr], same = old && (old.v instanceof Date && v instanceof Date ? toISO(old.v) === toISO(v) : old.v === v);
    if (same) return;
    ws[addr] = Object.assign({}, old && old.f ? { f: old.f } : {}, v instanceof Date ? { t: "d", v } : { t: "n", v });   // keep the formula
    if (onPut) onPut(sheet, addr);
  };
  const ag = as ? XLSX.utils.sheet_to_json(as, { header: 1, raw: true, defval: null }) : [];
  const asm = (re, d) => { const r = ag.find((x) => isT(x[0], re)); return r && r[1] != null && r[1] !== "" ? r[1] : d; };
  const plan = parseDate(asm(/plan date/i, null)), lagIn = +asm(/orcr in hand/i, 7), lagPend = +asm(/orcr still pending/i, 14);
  const g = XLSX.utils.sheet_to_json(sch, { header: 1, raw: true, defval: null, blankrows: true });
  const r0 = XLSX.utils.decode_range(sch["!ref"]).s.r;
  const hi = g.findIndex((r) => r.some((c) => isT(c, /^expected cash-in$/i)) && r.some((c) => isT(c, /^proceeds/i)));
  if (hi < 0) return;
  const H = g[hi].map((c) => (typeof c === "string" ? c.trim().toLowerCase() : ""));
  const col = (re) => H.findIndex((x) => re.test(x));
  const cI = col(/^orcr status/), cJ = col(/^orig orcr available/), cK = col(/^projected disposal month/), cL = col(/^proceeds/), cM = col(/^ready-to-sell/), cN = col(/^lag/), cO = col(/^expected cash-in$/);
  const c0 = XLSX.utils.decode_range(sch["!ref"]).s.c, A = (r, c) => XLSX.utils.encode_cell({ r: r0 + r, c: c0 + c });
  const units = [];
  for (let r = hi + 1; r < g.length; r++) {
    const row = g[r];
    if (isT(row[0], /^total$/i)) break;
    if (!row[0] || typeof row[cL] !== "number") continue;
    let ready = [parseDate(row[cK]), parseDate(row[cJ]), plan].filter(Boolean).reduce((m, d) => (!m || d > m ? d : m), null);
    if (userSet(schN, A(r, cM)) && parseDate(row[cM])) ready = parseDate(row[cM]);
    let lag = /^pending/i.test(String(row[cI] || "")) ? lagPend : lagIn;
    if (userSet(schN, A(r, cN)) && typeof row[cN] === "number") lag = row[cN];
    let cash = ready ? addDays(new Date(ready.getFullYear(), ready.getMonth(), ready.getDate()), lag) : null;
    if (userSet(schN, A(r, cO)) && parseDate(row[cO])) cash = parseDate(row[cO]);
    if (ready && cM >= 0) put(schN, sch, A(r, cM), new Date(ready.getFullYear(), ready.getMonth(), ready.getDate()));
    if (cN >= 0) put(schN, sch, A(r, cN), lag);
    if (cash && cO >= 0) put(schN, sch, A(r, cO), cash);
    units.push({ cash: cash ? toISO(cash) : null, amt: row[cL] });
  }
  // weekly rows
  const wg = XLSX.utils.sheet_to_json(wk, { header: 1, raw: true, defval: null, blankrows: true });
  const w0 = XLSX.utils.decode_range(wk["!ref"]).s;
  const rf = wg.findIndex((r) => isT(r[0], /^from$/i)), rt = wg.findIndex((r) => isT(r[0], /^to$/i));
  const ra = wg.findIndex((r) => isT(r[0], /^disposal proceeds$/i)), ru = wg.findIndex((r) => isT(r[0], /^units\b.*cash received/i));
  if (rf < 0 || rt < 0 || ra < 0) return;
  for (let c = 1; c < wg[rf].length; c++) {
    const f = parseDate(wg[rf][c]), t = parseDate(wg[rt][c]);
    if (!f || !t) continue;
    const a = toISO(f), b = toISO(t), inWk = units.filter((u) => u.cash && u.cash >= a && u.cash <= b);
    put(wkN, wk, XLSX.utils.encode_cell({ r: w0.r + ra, c: w0.c + c }), inWk.reduce((s, u) => s + u.amt, 0));
    if (ru >= 0) put(wkN, wk, XLSX.utils.encode_cell({ r: w0.r + ru, c: w0.c + c }), inWk.length);
  }
}
// Cells each formula reads: [{ sheet, r1, c1, r2, c2 }] (0-based; whole columns = all rows). Text in quotes is ignored.
const refCache = new Map();
function formulaRefs(f, own) {
  const key = own + "\u0001" + f;
  let out = refCache.get(key);
  if (out) return out;
  out = [];
  const src = String(f).replace(/"(?:[^"]|"")*"/g, '""');
  const RE = /(?<![\w.$'])(?:'((?:[^']|'')+)'!|([A-Za-z0-9_.]+)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})(?![\w(!])/gi;
  for (const m of src.matchAll(RE)) {
    const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2] != null ? m[2] : own, ref = m[3].replace(/\$/g, "").toUpperCase();
    try {
      if (/^[A-Z]+:[A-Z]+$/.test(ref)) { const [p, q] = ref.split(":"); out.push({ sheet, r1: 0, r2: 1e9, c1: XLSX.utils.decode_col(p), c2: XLSX.utils.decode_col(q) }); }
      else { const rg = XLSX.utils.decode_range(ref.includes(":") ? ref : ref + ":" + ref); out.push({ sheet, r1: rg.s.r, r2: rg.e.r, c1: rg.s.c, c2: rg.e.c }); }
    } catch (_) {}
  }
  if (refCache.size > 50000) refCache.clear();
  refCache.set(key, out);
  return out;
}
// every formula cell of the workbook (kept until the formulas change)
function formulaIndex(wb, stampKey) {
  if (wb.__fx && wb.__fxKey === stampKey) return wb.__fx;
  const all = [];
  for (const n of wb.SheetNames) {
    const ws = wb.Sheets[n];
    for (const a of Object.keys(ws)) {
      if (a[0] === "!" || !ws[a] || !ws[a].f) continue;
      const rc = XLSX.utils.decode_cell(a);
      all.push({ sheet: n, addr: a, refs: formulaRefs(ws[a].f, n), r: rc.r, c: rc.c });
    }
  }
  wb.__fx = all; wb.__fxKey = stampKey;
  return all;
}
const tick = () => new Promise((r) => setTimeout(r, 0));
const nowMs = () => (typeof performance !== "undefined" ? performance.now() : Date.now());
// After an edit: recalculate only the formulas that (directly or through other formulas) read an edited cell.
// Works in small steps so the page never freezes, even on big files (e.g. AR Aging's "Aging by Customer" SUMIFS).
async function recalcEdited(k) {
  const s = state[k], wb = s.wbEdit;
  if (!wb) return;
  s.recalc = s.recalc || {};
  const tr = (n) => (s.recalc[n] = s.recalc[n] || new Set());
  let front = [];
  for (const [sheet, cells] of Object.entries(s.edits || {})) for (const a of Object.keys(cells)) { const rc = XLSX.utils.decode_cell(a); front.push({ sheet, r: rc.r, c: rc.c }); }
  if (k === "ds") recalcUnitPlan(wb, s.edits, (sheet, addr) => { tr(sheet).add(addr); const rc = XLSX.utils.decode_cell(addr); front.push({ sheet, r: rc.r, c: rc.c }); });
  const fEdits = Object.values(s.edits || {}).reduce((n, m) => n + Object.values(m).filter((v) => v && typeof v === "object" && v.formula).length, 0);
  const all = formulaIndex(wb, `${(s.ops || []).length}:${fEdits}`);
  let t = nowMs();
  for (let round = 0; front.length && round < 30; round++) {
    const bySheet = new Map();
    for (const d of front) { if (!bySheet.has(d.sheet)) bySheet.set(d.sheet, []); bySheet.get(d.sheet).push(d); }
    const reads = (x) => x.refs.some((rf) => { const ds = bySheet.get(rf.sheet); return ds && ds.some((d) => d.r >= rf.r1 && d.r <= rf.r2 && d.c >= rf.c1 && d.c <= rf.c2); });
    const cache = new Map(), next = [];
    for (let i = 0; i < all.length; i++) {
      if (nowMs() - t > 30) { await tick(); t = nowMs(); }          // let the page breathe
      const x = all[i];
      if (!reads(x)) continue;
      const ws = wb.Sheets[x.sheet], cell = ws[x.addr];
      if (!cell || !cell.f) continue;
      let v;
      try { v = evalSimpleFormula(cell.f, ws, wb, cache); } catch (_) { continue; }
      if (typeof cell.v !== "number" || Math.abs(cell.v - v) > 1e-9) {
        cell.v = v; cell.t = "n"; delete cell.w;
        tr(x.sheet).add(x.addr);
        next.push({ sheet: x.sheet, r: x.r, c: x.c });
      }
    }
    front = next;
  }
}

const reingestJobs = {};
function reingest(k) {
  if (reingestJobs[k]) { reingestJobs[k].again = true; return reingestJobs[k].p; }
  const job = (reingestJobs[k] = { again: false });
  job.p = (async () => {
    try { do { job.again = false; await reingestOnce(k); } while (job.again); }
    finally { delete reingestJobs[k]; }
  })();
  return job.p;
}
async function reingestOnce(k) {
  const s = state[k], wb = s.wbEdit;
  try {
    await recalcEdited(k);
    if (fe.k === k && !$("fe-modal").classList.contains("hidden")) renderEditor();   // show recalculated totals
    if (k === "ar" || k === "ap") {
      const sheetName = wb.SheetNames.find((n) => /^detail$/i.test(n)) || wb.SheetNames[0];
      const grid = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
      const { headerIdx, headers } = findHeader(grid);
      const r0 = XLSX.utils.decode_range(wb.Sheets[sheetName]["!ref"] || "A1").s;
      grid.forEach((row, i) => (row.__r = r0.r + i));
      s.rows = grid.slice(headerIdx + 1).filter((r) => r.some((x) => x != null && x !== ""));
      s.c0 = r0.c;
      s.headers = headers;
    } else if (k === "cf") { const res = parseCashflowReport(wb); s.rows = res.rows; renderRetail(); }
    else if (k === "ds") { s.rows = parseUnitDisposals(wb).rows; renderDisposalMonths(); }
    else if (CO_META[k]) { const res = CO_META[k].parse(wb); s.byDate = res.byDate; s.info = coInfo(res); renderCashOutSources(); }
    refresh();
  } catch (e) { toast(`${KIND_NAME[k]}: the edited file can no longer be read — ${escapeHtml(e.message)}`, true, 7000); }
}

// remembered sources (no original): edit the schedule itself
function renderRememberedEditor() {
  const k = fe.k, s = state[k], re = s.remEdited || (s.remEdited = {});
  $("fe-range").textContent = ""; $("fe-prev").disabled = $("fe-next").disabled = true;
  let head, body;
  if (k === "ds") {
    head = "<tr><th>File week</th><th>Units</th><th>Disposal proceeds</th></tr>";
    body = s.rows.map((x, i) => `<tr><td class="l">${toISO(x.from)} – ${toISO(x.to)}</td>
      <td class="fe-c num${re["u" + i] ? " edited" : ""}" data-rem="u${i}">${x.units == null ? "" : x.units}</td>
      <td class="fe-c num${re["a" + i] ? " edited" : ""}" data-rem="a${i}">${showVal(x.amount)}</td></tr>`).join("");
  } else {
    const ds = Object.keys(s.byDate).sort();
    head = "<tr><th>Date</th><th>Amount</th></tr>";
    body = ds.map((d) => `<tr><td class="l">${d}</td><td class="fe-c num${re[d] ? " edited" : ""}" data-rem="${d}">${showVal(s.byDate[d])}</td></tr>`).join("");
  }
  $("fe-grid").innerHTML = `<table class="rem"><thead>${head}</thead><tbody>${body}</tbody></table>`;
  $("fe-note").textContent = `${editCount(k)} edited value(s). Kept in the saved summary only (no original file this session).`;
  $("fe-undo").disabled = true;
  $("fe-find").disabled = true; $("fe-addrow").disabled = true; $("fe-addcol").disabled = true;
}
function beginRememberedEdit(td) {
  if (td.querySelector("input")) return;
  const k = fe.k, s = state[k], key = td.dataset.rem;
  const inp = document.createElement("input"); inp.className = "fe-in"; inp.value = td.textContent.trim();
  td.textContent = ""; td.appendChild(inp); inp.focus(); inp.select();
  let done = false;
  const finish = (keep) => {
    if (done) return; done = true;
    if (keep) {
      const n = Number(inp.value.replace(/,/g, ""));
      if (!isNaN(n)) {
        if (k === "ds") { const i = +key.slice(1); if (key[0] === "a") s.rows[i].amount = n; else s.rows[i].units = n; renderDisposalMonths(); }
        else { s.byDate[key] = n; renderCashOutSources(); }
        s.remEdited[key] = true; s.dirty = true; saved = false; renderFileList();
        if (report) refresh();
      }
    }
    renderEditor();
  };
  inp.onkeydown = (e) => { if (e.key === "Enter" || e.key === "Tab") { e.preventDefault(); finish(true); } if (e.key === "Escape") finish(false); };
  inp.onblur = () => finish(true);
}

// Inserted rows / columns, replayed on the real file: same style as the previous line, merges and formulas shifted.
function replayInserts(wb, ops) {
  const eachFormula = (fn) => wb.eachSheet((w) => w.eachRow({ includeEmpty: false }, (row) => row.eachCell({ includeEmpty: false }, (cell) => fn(w, cell))));
  // shared formulas -> plain formulas first, so moving rows cannot break them
  eachFormula((w, cell) => { const v = cell.value; if (v && typeof v === "object" && v.sharedFormula) cell.value = { formula: cell.formula, result: v.result }; });
  for (const op of ops) {
    const ws = wb.getWorksheet(op.sheet);
    if (!ws) continue;
    const key = op.axis === "r" ? "r" : "c";
    const merges = Object.values(ws._merges || {}).map((m) => m.range).filter(Boolean);
    merges.forEach((m) => { try { ws.unMergeCells(m); } catch (_) {} });
    if (op.axis === "r") {
      ws.insertRow(op.at + 1, [], "i");                               // style inherited from the row above
      const prev = ws.getRow(op.at);
      if (prev.height) ws.getRow(op.at + 1).height = prev.height;
    } else {
      ws.spliceColumns(op.at + 1, 0, []);
      const pc = ws.getColumn(op.at), nc = ws.getColumn(op.at + 1);
      if (pc.width) nc.width = pc.width;
      ws.eachRow({ includeEmpty: true }, (row) => { const src = row.getCell(op.at); if (src && src.style) row.getCell(op.at + 1).style = JSON.parse(JSON.stringify(src.style)); });
    }
    merges.forEach((m) => { const r = XLSX.utils.decode_range(m); if (r.s[key] >= op.at) { r.s[key]++; r.e[key]++; } else if (r.e[key] >= op.at) r.e[key]++; try { ws.mergeCells(XLSX.utils.encode_range(r)); } catch (_) {} });
    eachFormula((w, cell) => {
      const v = cell.value;
      if (!(v && typeof v === "object" && v.formula)) return;
      const nf = shiftRefs(v.formula, w.name, op.sheet, op.axis, op.at, op.axis === "r" ? cell.row - 1 : cell.col - 1);
      if (nf !== v.formula) cell.value = { formula: nf, result: v.result };
    });
  }
}

// Excel features found in a workbook file (zip part names). Used to warn when a rewrite would drop some.
function xlFeatures(buf) {
  const cat = (k) => /pivot/i.test(k) ? "pivot tables" : /\/charts\//i.test(k) ? "charts" : /threadedComment/i.test(k) ? "threaded comments"
    : /comments\d*\.xml/i.test(k) ? "comments" : /\/drawings\/drawing/i.test(k) ? "drawings / shapes" : /vbaProject/i.test(k) ? "macros" : /persons|documenttasks/i.test(k) ? "tasks / @mentions" : null;
  try { return new Set((XLSX.read(buf, { type: "array", bookFiles: true, sheetRows: 1 }).keys || []).map(cat).filter(Boolean)); } catch (_) { return new Set(); }
}

// The uploaded file as it is now in the system (inserted lines and edits included; edited cells yellow), as ExcelJS.
async function excelOf(k) {
  const s = state[k];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(s.buffer);
  if (s.ops && s.ops.length) replayInserts(wb, s.ops);
  for (const [sheet, cells] of Object.entries(s.edits)) {
    const ws = wb.getWorksheet(sheet);
    if (!ws) continue;
    for (const [addr, val] of Object.entries(cells)) {
      const cell = ws.getCell(addr);
      if (val && typeof val === "object" && !(val instanceof Date) && val.formula) {
        const now = s.wbEdit && s.wbEdit.Sheets[sheet] && s.wbEdit.Sheets[sheet][addr];
        cell.value = { formula: val.formula, result: now && typeof now.v === "number" ? now.v : undefined };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
        continue;
      }
      cell.value = val instanceof Date ? xlDate(val) : val;
      if (val instanceof Date && !/[dmy]/i.test(cell.numFmt || "")) cell.numFmt = "yyyy-mm-dd";
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };
    }
  }
  // totals and other formulas that the system recalculated: keep the formula, store the new result
  for (const [sheet, set] of Object.entries(s.recalc || {})) {
    const ws = wb.getWorksheet(sheet), src = s.wbEdit && s.wbEdit.Sheets[sheet];
    if (!ws || !src) continue;
    for (const addr of set) {
      if (s.edits[sheet] && addr in s.edits[sheet]) continue;
      const cell = ws.getCell(addr), now = src[addr];
      if (!now) continue;
      const res = now.v instanceof Date ? xlDate(now.v) : now.v, cv = cell.value;
      if (cv && typeof cv === "object" && (cv.formula || cv.sharedFormula)) cell.value = Object.assign({}, cv, { result: res });
    }
  }
  return wb;
}

// ---- write the edits into the ORIGINAL files (called from Save summary)
async function writeEditedOriginals() {
  let todo = FILE_KINDS.map(([x]) => x).filter((x) => hasOriginal(x) && (editCount(x) || state[x].syncOriginal) && state[x].handle);
  if (!todo.length) return { ok: true, written: [] };
  const written = [], skipped = [];
  // 1. permission to change each original (asked first, while the click is still fresh); refused = only that file is skipped
  const allowed = [];
  for (const k of todo) {
    const hnd = state[k].handle;
    let ok = false;
    try { ok = (await hnd.queryPermission({ mode: "readwrite" })) === "granted" || (await hnd.requestPermission({ mode: "readwrite" })) === "granted"; } catch (_) {}
    if (ok) allowed.push(k); else skipped.push(`${hnd.name} (no permission to change it, or it was moved)`);
  }
  todo = allowed;
  // 2. a backup needs the save folder
  if (todo.length && !(await folderReady())) return { ok: false, msg: "Connect the save folder first (bottom of the sidebar) so a backup of each original file can be kept before it is updated." };
  for (const k of todo) {
    const s = state[k], hnd = s.handle;
    // changed in Excel after it was uploaded / linked? then ask before replacing it
    try {
      const now = (await hnd.getFile()).lastModified;
      if (s.origSeen && now > s.origSeen + 2000 && !confirm(`"${hnd.name}" was changed outside the system (${new Date(now).toLocaleString("en-PH")}) after it was uploaded or linked.\n\nOK = replace it with the system's version (a backup of it is kept first)\nCancel = leave the original as it is`)) { skipped.push(`${hnd.name} (changed in Excel — left as it is)`); continue; }
    } catch (_) { skipped.push(`${hnd.name} (could not be opened — moved or renamed?)`); continue; }
    if (!/\.(xlsx|xlsm)$/i.test(hnd.name)) { toast(`"${escapeHtml(hnd.name)}" is not an .xlsx file — its edits stay in the system only.`, true, 7000); continue; }
    // rebuild the file with ExcelJS: only the edited cells change (and turn yellow)
    const wb = await excelOf(k);
    wb.calcProperties = Object.assign(wb.calcProperties || {}, { fullCalcOnLoad: true });   // Excel also recalculates on open
    const out = await wb.xlsx.writeBuffer();
    // warn when the rewrite would drop Excel features (pivot tables, charts, …) and let the user decide
    const had = xlFeatures(s.buffer), kept = xlFeatures(out instanceof ArrayBuffer ? out : out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength));
    const lost = [...had].filter((x) => !kept.has(x));
    if (lost.length && !confirm(`"${hnd.name}" contains ${lost.join(", ")} that will be LOST if the system updates this file.\n\nA backup of the original is kept in FILES/${UPLOAD_DIR}.\n\nOK = update the original anyway\nCancel = keep the edits in the system only (the file is not changed)`)) {
      skipped.push(hnd.name);
      continue;
    }
    // backup once per session, then overwrite the original
    if (s.backedUp !== hnd.name) {
      if (!uploadBatch) uploadBatch = batchFolderName();
      const dot = hnd.name.lastIndexOf("."), bname = await uniqueName([UPLOAD_DIR, uploadBatch], `${hnd.name.slice(0, dot)} (before edit)${hnd.name.slice(dot)}`);
      let orig = s.buffer;
      try { orig = await (await hnd.getFile()).arrayBuffer(); } catch (_) {}
      await writeTo([UPLOAD_DIR, uploadBatch], bname, orig);
      s.backedUp = hnd.name; s.backupAt = `${UPLOAD_DIR}/${uploadBatch}/${bname}`;
    }
    const w = await hnd.createWritable(); await w.write(out); await w.close();
    try { s.origSeen = (await hnd.getFile()).lastModified; await rememberOriginal(k, hnd, s.origSeen); } catch (_) {}
    s.syncOriginal = false;
    // the updated file is now the original; keep the edited cells marked as written
    for (const [sheet, cells] of Object.entries(s.edits)) { s.written[sheet] = s.written[sheet] || {}; Object.keys(cells).forEach((a) => (s.written[sheet][a] = true)); }
    s.buffer = out instanceof ArrayBuffer ? out : out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
    s.edits = {}; s.recalc = {}; s.ops = []; s.editedSave = true;   // the edited copy in Uploaded Files is kept by the shared save
    written.push(`${hnd.name}${s.backupAt ? " (backup: " + s.backupAt + ")" : ""}`);
  }
  return { ok: true, written, skipped };
}

// Upload pop-up
const modal = $("upload-modal");
function openUpload() { uploadBatch = null; modal.classList.remove("hidden"); }   // each time the pop-up is opened, a new dated folder is used
function closeUpload() { modal.classList.add("hidden"); }
$("btn-upload").onclick = openUpload;
$("bd-rem").innerHTML = REMARK_FILTERS.map(([k, l]) => `<option value="${k}">${l}</option>`).join("");
$("bd-rem").onchange = () => openBreakdown(bdWeek);
$("pd-remark").onchange = () => { pdRemark = $("pd-remark").value; renderPastDue(); };
$("bd-close").onclick = closeBreakdown;
$("bd-done").onclick = closeBreakdown;
$("bd-modal").addEventListener("mousedown", (e) => { if (e.target === $("bd-modal")) closeBreakdown(); });
$("um-close").onclick = closeUpload;
$("um-done").onclick = closeUpload;
modal.addEventListener("mousedown", (e) => { if (e.target === modal) closeUpload(); });
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape") return;
  if (!modal.classList.contains("hidden")) closeUpload();
  if (!$("bd-modal").classList.contains("hidden")) closeBreakdown();
});

// V3.4 wiring: file editor
FILE_KINDS.forEach(([k]) => resetFileEdits(k));
wirePickers();
$("fe-close").onclick = $("fe-done").onclick = () => $("fe-modal").classList.add("hidden");
$("fe-modal").addEventListener("mousedown", (e) => { if (e.target === $("fe-modal")) $("fe-modal").classList.add("hidden"); });
$("fe-grid").addEventListener("mousedown", (e) => {
  const fb = e.target.closest(".fe-fb");
  if (fb) { e.preventDefault(); e.stopPropagation(); openFilter(+fb.dataset.fcol, fb); return; }
  const td = e.target.closest("td.fe-c");
  if (td && td.dataset.rem) { beginRememberedEdit(td); return; }
  if (e.target.closest("input")) return;
  const th = e.target.closest("thead th[data-col]"), rn = e.target.closest("th.rn[data-ri]");
  if (th) { e.preventDefault(); const col = +th.dataset.col; fe.sel = { a: [fe.pinned, col], b: [Math.max(fe.pinned, fe.pageRows.length - 1), col] }; paintSel(); $("fe-grid").focus(); return; }
  if (rn) { e.preventDefault(); const i = +rn.dataset.ri, last = fe.grid.c0 + fe.grid.cols - 1; fe.sel = { a: [i, fe.grid.c0], b: [i, last] }; paintSel(); $("fe-grid").focus(); return; }
  if (!td || !td.dataset.a) return;
  e.preventDefault();
  const p = [+td.dataset.i, +td.dataset.c];
  if (e.shiftKey && fe.sel) fe.sel.b = p; else fe.sel = { a: p, b: p };
  fe.dragging = true; paintSel(); $("fe-grid").focus();
});
$("fe-grid").addEventListener("mouseover", (e) => {
  if (!fe.dragging) return;
  const td = e.target.closest("td.fe-c[data-a]");
  if (td) { fe.sel.b = [+td.dataset.i, +td.dataset.c]; paintSel(); }
});
document.addEventListener("mouseup", () => (fe.dragging = false));
$("fe-grid").addEventListener("dblclick", (e) => { const td = e.target.closest("td.fe-c[data-a]"); if (td) { fe.sel = { a: [+td.dataset.i, +td.dataset.c], b: [+td.dataset.i, +td.dataset.c] }; paintSel(); beginCellEdit(td); } });
$("fe-grid").addEventListener("keydown", (e) => {
  if (!fe.sel || e.target.closest("input") || isRemembered(fe.k)) return;
  const bx = selBox();
  if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); applyToSelection(""); return; }
  if (e.key === "F2" || e.key === "Enter") { e.preventDefault(); beginCellEdit(anchorTd()); return; }
  const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
  if (mv) {
    e.preventDefault();
    const tgt = e.shiftKey ? fe.sel.b : fe.sel.a;
    const n = [Math.min(fe.pageRows.length - 1, Math.max(0, tgt[0] + mv[0])), Math.min(fe.grid.c0 + fe.grid.cols - 1, Math.max(fe.grid.c0, tgt[1] + mv[1]))];
    if (e.shiftKey) fe.sel.b = n; else fe.sel = { a: n, b: n };
    paintSel(); return;
  }
  if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); beginCellEdit(anchorTd(), e.key); }
});
document.addEventListener("mousedown", (e) => { const pop = $("fe-pop"); if (!pop.classList.contains("hidden") && !pop.contains(e.target) && !e.target.closest(".fe-fb")) pop.classList.add("hidden"); });
$("fe-find").onclick = openReplace;
$("fe-addrow").onclick = () => addLine("r");
$("fe-addcol").onclick = () => addLine("c");
document.addEventListener("keydown", (e) => { if ((e.ctrlKey || e.metaKey) && (e.key === "h" || e.key === "H") && !$("fe-modal").classList.contains("hidden")) { e.preventDefault(); openReplace(); } });
$("fe-clearf").onclick = () => { fe.filters[fe.sheet] = {}; fe.page = 0; fe.sel = null; renderEditor(); };
$("fe-search").oninput = () => { fe.q = $("fe-search").value; fe.page = 0; fe.sel = null; renderEditor(); };
$("fe-prev").onclick = () => { fe.page--; fe.sel = null; renderEditor(); };
$("fe-next").onclick = () => { fe.page++; fe.sel = null; renderEditor(); };
$("fe-undo").onclick = () => {
  const k = fe.k;
  if (!confirm(`Undo all ${editCount(k)} edit(s) in ${KIND_NAME[k]}?`)) return;
  state[k].edits = {}; state[k].wbEdit = null; state[k].recalc = {}; state[k].ops = [];          // back to the file as uploaded
  editWorkbook(k); renderFeTabs(); afterFileEdit(k);
};
document.addEventListener("keydown", (e) => {
  if (e.key !== "Escape" || $("fe-modal").classList.contains("hidden") || (document.activeElement && document.activeElement.classList.contains("fe-in"))) return;
  if (!$("fe-pop").classList.contains("hidden")) { $("fe-pop").classList.add("hidden"); return; }
  $("fe-modal").classList.add("hidden");
});

// V3.6 several people at the same time
renderUser();
$("btn-user").onclick = () => askUserName();
$("btn-final").onclick = finalizeWeek;
$("btn-refresh").onclick = async () => { if (!baseDir) { toast("Connect the shared FILES folder first (bottom of the sidebar).", true, 5000); return; } await loadShared(); toast("Showing everyone's latest updates" + (SHARED_KINDS.some((k) => state[k].dirty) || bankDirty ? "<small>Parts you changed and have not saved yet were kept as yours.</small>" : "")); };
renderBankInfo();

// show unexpected errors instead of failing silently
window.addEventListener("error", (e) => { if (e && e.message) toast("Something went wrong:<small>" + escapeHtml(e.message) + "</small>", true, 10000); });
window.addEventListener("unhandledrejection", (e) => { const m = e && e.reason ? (e.reason.message || String(e.reason)) : "unknown error"; toast("Something went wrong:<small>" + escapeHtml(m) + "</small>", true, 10000); });
