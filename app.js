/**
 * Dress Rental Demo — in-memory "database" and overlap logic.
 *
 * A dress is NOT available for [newFrom, newTo] if there exists a booking
 * for the same code where: newFrom <= existing.to AND newTo >= existing.from
 * (inclusive date comparison on ISO date strings YYYY-MM-DD).
 */

// --- Seed data (replaced when you load a CSV from a URL) ---
const SEED_BOOKINGS = [
  { code: "D100", from: "2026-03-30", to: "2026-04-02", name: "Sara" },
  { code: "D100", from: "2026-04-05", to: "2026-04-07", name: "Mona" },
];

/** @type {{ code: string, from: string, to: string, name: string }[]} */
let bookings = [...SEED_BOOKINGS];

/**
 * Your live sheet (CSV export). Edit: File → Share must allow access for export to work.
 * If data is on another tab, change gid= to the number in the URL (#gid=…).
 */
const DEFAULT_GOOGLE_SHEET_CSV =
  "https://docs.google.com/spreadsheets/d/1sNSPfX90oVsfX1uL5MGNv9T9kQilB1bHhcEaOltdTvw/export?format=csv&gid=0";

/**
 * Client mode (default): clean UI, no URL fields. Set your Apps Script Web App /exec URL here before sharing.
 * Dev mode: open the site with ?dev=1 to edit CSV URL and script URL in the page.
 */
const IS_DEV_MODE = new URLSearchParams(window.location.search).get("dev") === "1";

/** Paste the Web App /exec URL here for client-facing builds (when ?dev=1 is not used). */
const CLIENT_SHEET_APPEND_URL =
  "https://script.google.com/macros/s/AKfycbx3zH45OFa0ANhJ2BJeyMvsfBNL24FPfpYESxgy3xefFTt7RV5hd9CO8KhUU6HLaWTx_A/exec";

const MIN_CHECK_MS = 450; // artificial delay so loading state is visible

// --- DOM ---
const checkForm = document.getElementById("check-form");
const checkCode = document.getElementById("check-code");
const checkFrom = document.getElementById("check-from");
const checkTo = document.getElementById("check-to");
const checkBtn = document.getElementById("check-btn");
const checkResult = document.getElementById("check-result");

const bookForm = document.getElementById("book-form");
const bookCode = document.getElementById("book-code");
const bookName = document.getElementById("book-name");
const bookFrom = document.getElementById("book-from");
const bookTo = document.getElementById("book-to");
const bookBtn = document.getElementById("book-btn");
const bookMessage = document.getElementById("book-message");

const filterCode = document.getElementById("filter-code");
const bookingsTbody = document.getElementById("bookings-tbody");
const bookingsEmpty = document.getElementById("bookings-empty");

const importForm = document.getElementById("import-form");
const importUrl = document.getElementById("import-url");
const importBtn = document.getElementById("import-btn");
const syncStatus = document.getElementById("sync-status");
const sheetAppendUrlInput = document.getElementById("sheet-append-url");
const refreshBookingsBtn = document.getElementById("refresh-bookings-btn");

const LS_SHEET_APPEND_URL = "royal-dress-sheet-append-url";

// --- Helpers ---

function normalizeCode(value) {
  return String(value || "").trim();
}

/**
 * Compare ISO date strings (YYYY-MM-DD) as orderable strings — valid for calendar dates.
 */
function parseOrder(iso) {
  return iso.replace(/-/g, "");
}

/**
 * Returns overlapping bookings for the same dress code (inclusive overlap rule).
 */
function findConflicts(code, from, to) {
  const c = normalizeCode(code);
  return bookings.filter(
    (b) =>
      normalizeCode(b.code).toLowerCase() === c.toLowerCase() &&
      from <= b.to &&
      to >= b.from
  );
}

function isDateRangeValid(from, to) {
  if (!from || !to) return false;
  return parseOrder(from) <= parseOrder(to);
}

function setButtonLoading(btn, loading) {
  const loader = btn.querySelector(".btn-loader");
  btn.disabled = loading;
  btn.classList.toggle("is-loading", loading);
  if (loader) loader.hidden = !loading;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function clearNode(el) {
  while (el.firstChild) el.removeChild(el.firstChild);
}

function renderCheckResult(available, conflicts) {
  clearNode(checkResult);
  if (available) {
    const div = document.createElement("div");
    div.className = "status-available";
    div.textContent = "Available ✅";
    checkResult.appendChild(div);
    return;
  }
  const wrap = document.createElement("div");
  wrap.className = "status-unavailable";
  wrap.innerHTML = "<strong>Not Available ❌</strong>";
  if (conflicts.length) {
    const sub = document.createElement("p");
    sub.className = "conflicts-title";
    sub.textContent = "Conflicting booking dates:";
    wrap.appendChild(sub);
    const ul = document.createElement("ul");
    ul.className = "conflicts-list";
    conflicts.forEach((b) => {
      const li = document.createElement("li");
      li.textContent = `${b.from} → ${b.to} (${b.name})`;
      ul.appendChild(li);
    });
    wrap.appendChild(ul);
  }
  checkResult.appendChild(wrap);
}

function showBookMessage(type, text) {
  clearNode(bookMessage);
  const div = document.createElement("div");
  div.className = type === "success" ? "msg-success" : "msg-error";
  div.textContent = text;
  bookMessage.appendChild(div);
}

function getFilteredBookings() {
  const q = normalizeCode(filterCode.value).toLowerCase();
  if (!q) return [...bookings].sort(sortBookings);
  return bookings
    .filter((b) => normalizeCode(b.code).toLowerCase().includes(q))
    .sort(sortBookings);
}

function sortBookings(a, b) {
  const byFrom = a.from.localeCompare(b.from);
  if (byFrom !== 0) return byFrom;
  return normalizeCode(a.code).localeCompare(normalizeCode(b.code));
}

function renderBookingsTable() {
  const rows = getFilteredBookings();
  clearNode(bookingsTbody);
  bookingsEmpty.hidden = rows.length > 0;

  rows.forEach((b) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><span class="code-badge">${escapeHtml(b.code)}</span></td>
      <td>${escapeHtml(b.name)}</td>
      <td>${escapeHtml(b.from)}</td>
      <td>${escapeHtml(b.to)}</td>
    `;
    bookingsTbody.appendChild(tr);
  });
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

// --- CSV import (Excel “Save As CSV” or Google Sheets export URL) ---

/** Split one CSV line respecting double-quoted cells. */
function splitCsvLine(line) {
  const cells = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cur += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ",") {
      cells.push(cur.trim());
      cur = "";
    } else {
      cur += c;
    }
  }
  cells.push(cur.trim().replace(/\r$/, ""));
  return cells;
}

function headerKey(cell) {
  return String(cell || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/**
 * Map header label → column index for: code, from, to, name
 */
function mapHeaderIndices(headerCells) {
  const aliases = {
    code: ["code", "dress code", "dress_code", "dresscode", "sku"],
    from: ["from", "from date", "from_date", "start", "start date"],
    to: ["to", "to date", "to_date", "end", "end date"],
    name: ["name", "customer", "customer name", "client"],
  };
  const idx = { code: -1, from: -1, to: -1, name: -1 };
  headerCells.forEach((raw, i) => {
    const h = headerKey(raw);
    for (const field of Object.keys(aliases)) {
      if (idx[field] !== -1) continue;
      if (aliases[field].includes(h)) idx[field] = i;
    }
  });
  return idx;
}

/** Normalize cell date to YYYY-MM-DD for overlap logic. */
function normalizeDateCell(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  // ISO date
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // Excel sometimes exports M/D/YYYY or D/M/YYYY — try native parse (locale-dependent)
  const t = Date.parse(s.replace(/(\d{1,2})\/(\d{1,2})\/(\d{4})/, "$2/$1/$3"));
  if (!Number.isNaN(t)) {
    const d = new Date(t);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  const d2 = new Date(s);
  if (!Number.isNaN(d2.getTime())) {
    const y = d2.getFullYear();
    const m = String(d2.getMonth() + 1).padStart(2, "0");
    const day = String(d2.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  return null;
}

/**
 * Parse CSV text into booking objects. Expects a header row.
 */
function bookingsFromCsvText(text) {
  const lines = text.split(/\r?\n/).filter((ln) => ln.trim().length > 0);
  if (lines.length < 2) {
    throw new Error("CSV must have a header row and at least one data row.");
  }
  const header = splitCsvLine(lines[0]);
  const col = mapHeaderIndices(header);
  if (col.code < 0 || col.from < 0 || col.to < 0 || col.name < 0) {
    throw new Error(
      'Need columns for dress code, from date, to date, and name (e.g. headers: code,from,to,name).'
    );
  }
  const out = [];
  for (let r = 1; r < lines.length; r++) {
    const cells = splitCsvLine(lines[r]);
    if (cells.every((c) => !c)) continue;
    const code = normalizeCode(cells[col.code]);
    const from = normalizeDateCell(cells[col.from]);
    const to = normalizeDateCell(cells[col.to]);
    const name = String(cells[col.name] || "").trim();
    if (!code || !from || !to || !name) continue;
    if (!isDateRangeValid(from, to)) continue;
    out.push({ code, from, to, name });
  }
  if (out.length === 0) {
    throw new Error("No valid data rows found (check dates and column names).");
  }
  return out;
}

async function fetchBookingsRowsFromUrl(url) {
  const res = await fetch(url, { method: "GET", cache: "no-store" });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} — check the URL and that the file is public.`);
  }
  const text = await res.text();
  return bookingsFromCsvText(text);
}

/** Avoids stale CDN/browser cache after a row was just appended in the Sheet. */
function withCacheBuster(url) {
  try {
    const u = new URL(url);
    u.searchParams.set("_", String(Date.now()));
    return u.toString();
  } catch {
    const sep = url.includes("?") ? "&" : "?";
    return `${url}${sep}_=${Date.now()}`;
  }
}

function sameBooking(a, b) {
  return (
    normalizeCode(a.code).toLowerCase() === normalizeCode(b.code).toLowerCase() &&
    a.from === b.from &&
    a.to === b.to &&
    String(a.name || "").trim() === String(b.name || "").trim()
  );
}

/**
 * After Apps Script appends a row, Sheets’ CSV export often lags or is cached.
 * We refetch with a cache-buster and, if the new row is still missing, merge it in.
 */
async function refreshBookingsAfterAppend(savedRow) {
  const url = withCacheBuster(getReadCsvUrl());
  try {
    await delay(400);
    let rows = await fetchBookingsRowsFromUrl(url);
    if (!rows.some((r) => sameBooking(r, savedRow))) {
      rows = [...rows, savedRow];
    }
    rows.sort(sortBookings);
    bookings = rows;
  } catch {
    bookings.push(savedRow);
    bookings.sort(sortBookings);
  }
  renderBookingsTable();
}

async function fetchAndApplyBookings(url) {
  const rows = await fetchBookingsRowsFromUrl(url);
  bookings = rows;
  renderBookingsTable();
  return rows.length;
}

function getReadCsvUrl() {
  if (IS_DEV_MODE && importUrl) {
    return String(importUrl.value || "").trim() || DEFAULT_GOOGLE_SHEET_CSV;
  }
  return DEFAULT_GOOGLE_SHEET_CSV;
}

function getSheetAppendUrl() {
  const configured = String(CLIENT_SHEET_APPEND_URL || "").trim();
  if (!IS_DEV_MODE && configured) return configured;
  if (IS_DEV_MODE && sheetAppendUrlInput) {
    const typed = String(sheetAppendUrlInput.value || "").trim();
    if (typed) return typed;
  }
  if (configured) return configured;
  if (!sheetAppendUrlInput) return "";
  return String(sheetAppendUrlInput.value || "").trim();
}

/**
 * POST one booking to your bound Apps Script web app (appends a row to the first sheet).
 * Uses text/plain body so the browser avoids a JSON preflight that Apps Script may not answer.
 */
function validateAppsScriptPostUrl(url) {
  const u = url.toLowerCase();
  // Common mistake: same link as “read CSV” — that URL returns HTML/CSV, not JSON.
  if (
    u.includes("docs.google.com/spreadsheets") &&
    (u.includes("export") || u.includes("format=csv") || u.includes("format%3dcsv"))
  ) {
    throw new Error(
      'Wrong URL: that is your Sheet’s CSV export (for “Load CSV” only). For saving bookings, paste the Web app URL from Apps Script — it starts with https://script.google.com/macros/s/…/exec'
    );
  }
  if (!u.includes("script.google.com/macros/")) {
    throw new Error(
      "Save-bookings URL must come from Apps Script: Deploy → Web app → copy the link (script.google.com/macros/s/…/exec)."
    );
  }
}

async function appendRowToGoogleSheet(booking) {
  const url = getSheetAppendUrl();
  if (!url) {
    throw new Error(
      "Add your Apps Script Web App URL in “Apps Script URL (save new bookings)” below."
    );
  }
  validateAppsScriptPostUrl(url);
  let res;
  try {
    res = await fetch(url, {
      method: "POST",
      mode: "cors",
      credentials: "omit",
      redirect: "follow",
      cache: "no-store",
      body: JSON.stringify(booking),
      headers: { "Content-Type": "text/plain;charset=utf-8" },
    });
  } catch (err) {
    if (err instanceof TypeError && err.message === "Failed to fetch") {
      throw new Error(
        "Could not reach the script (network or browser blocked the request). If DevTools shows 401 Unauthorized, fix the Web app deployment (see message below for 401)."
      );
    }
    throw err;
  }

  const text = await res.text();

  if (res.status === 401 || res.status === 403) {
    throw new Error(
      "Google returned " +
        res.status +
        " Unauthorized — the Web app does not allow anonymous access. In Apps Script: Deploy → Manage deployments → ✏️ Edit → set “Who has access” to **Anyone** (not “Anyone with Google account”), Execute as: **Me** → Deploy **new version**. If you use Google Workspace, an admin may block public web apps."
    );
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    if (/docs\.google\.com|text\/html|<!doctype html/i.test(text.slice(0, 200))) {
      throw new Error(
        "Got a web page instead of JSON — the save URL is probably not your /exec Web app link, or deploy access is not set to “Anyone”."
      );
    }
    throw new Error(
      "Unexpected response from Google Script — use the Deploy → Web app URL ending in /exec, then redeploy after code edits."
    );
  }
  if (!parsed || parsed.ok !== true) {
    throw new Error((parsed && parsed.error) || "Sheet script did not confirm the row.");
  }
}

function showSyncStatus(type, text) {
  if (!syncStatus) return;
  clearNode(syncStatus);
  const div = document.createElement("div");
  div.className = type === "success" ? "msg-success" : "msg-error";
  div.textContent = text;
  syncStatus.appendChild(div);
}

// --- Events: Check availability ---
checkForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const code = normalizeCode(checkCode.value);
  const from = checkFrom.value;
  const to = checkTo.value;

  clearNode(checkResult);
  if (!code || !from || !to) {
    const div = document.createElement("div");
    div.className = "msg-error";
    div.textContent = "Please fill in dress code and both dates.";
    checkResult.appendChild(div);
    return;
  }
  if (!isDateRangeValid(from, to)) {
    const div = document.createElement("div");
    div.className = "msg-error";
    div.textContent = "“To date” must be on or after “From date”.";
    checkResult.appendChild(div);
    return;
  }

  setButtonLoading(checkBtn, true);
  await delay(MIN_CHECK_MS);

  const conflicts = findConflicts(code, from, to);
  renderCheckResult(conflicts.length === 0, conflicts);
  setButtonLoading(checkBtn, false);
});

// --- Events: Book ---
bookForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const code = normalizeCode(bookCode.value);
  const name = String(bookName.value || "").trim();
  const from = bookFrom.value;
  const to = bookTo.value;

  clearNode(bookMessage);

  if (!code || !name || !from || !to) {
    showBookMessage("error", "Please fill in all fields.");
    return;
  }
  if (!isDateRangeValid(from, to)) {
    showBookMessage("error", "“To date” must be on or after “From date”.");
    return;
  }

  setButtonLoading(bookBtn, true);
  await delay(MIN_CHECK_MS);

  const conflicts = findConflicts(code, from, to);
  if (conflicts.length > 0) {
    const names = conflicts.map((c) => `${c.from} → ${c.to}`).join("; ");
    showBookMessage(
      "error",
      `This dress is not available for those dates. Conflicts: ${names}`
    );
    setButtonLoading(bookBtn, false);
    return;
  }

  const payload = { code, from, to, name };
  try {
    await appendRowToGoogleSheet(payload);
  } catch (err) {
    showBookMessage("error", err.message || String(err));
    setButtonLoading(bookBtn, false);
    return;
  }

  await refreshBookingsAfterAppend(payload);

  showBookMessage(
    "success",
    IS_DEV_MODE
      ? `Saved to Google Sheet for ${name} — ${code} from ${from} to ${to}.`
      : `Booking saved for ${name} — dress ${code}, ${from} through ${to}.`
  );
  setButtonLoading(bookBtn, false);
});

filterCode.addEventListener("input", () => renderBookingsTable());

// --- Events: Import CSV (dev mode UI) ---
if (importForm && IS_DEV_MODE) {
  importForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const url = String(importUrl.value || "").trim();
    clearNode(syncStatus);
    if (!url) {
      showSyncStatus("error", "Enter a public CSV URL.");
      return;
    }
    setButtonLoading(importBtn, true);
    try {
      const n = await fetchAndApplyBookings(url);
      showSyncStatus("success", `Loaded ${n} booking(s). List replaced.`);
    } catch (err) {
      const msg =
        err instanceof TypeError && err.message === "Failed to fetch"
          ? "Network/CORS blocked. Serve this page over http(s) (e.g. VS Code Live Server), not file://."
          : err.message || String(err);
      showSyncStatus("error", msg);
    } finally {
      setButtonLoading(importBtn, false);
    }
  });
}

if (refreshBookingsBtn) {
  refreshBookingsBtn.addEventListener("click", async () => {
    clearNode(syncStatus);
    setButtonLoading(refreshBookingsBtn, true);
    try {
      const n = await fetchAndApplyBookings(getReadCsvUrl());
      showSyncStatus(
        "success",
        IS_DEV_MODE ? `Reloaded ${n} row(s) from URL.` : `Showing ${n} reservation(s).`
      );
    } catch (err) {
      const isNet = err instanceof TypeError && err.message === "Failed to fetch";
      showSyncStatus(
        "error",
        isNet
          ? IS_DEV_MODE
            ? "Could not fetch (check URL and use http(s) not file://)."
            : "Could not refresh. Check your connection or try again."
          : err.message || String(err)
      );
    } finally {
      setButtonLoading(refreshBookingsBtn, false);
    }
  });
}

/**
 * On load: fetch your Google Sheet CSV (or `?csv=URL` to use another file). On failure, sample data stays.
 */
function initRemoteBookings() {
  const queryCsv = IS_DEV_MODE ? new URLSearchParams(window.location.search).get("csv") : null;
  const url = (queryCsv && queryCsv.trim()) || DEFAULT_GOOGLE_SHEET_CSV;
  if (importUrl) importUrl.value = url;

  const showLoadOnImportButton = IS_DEV_MODE && importBtn;
  if (showLoadOnImportButton) setButtonLoading(importBtn, true);
  fetchAndApplyBookings(url)
    .then((n) => {
      if (IS_DEV_MODE) {
        const label = queryCsv?.trim() ? "Loaded from custom URL" : "Connected to Google Sheet";
        showSyncStatus("success", `${label}: ${n} booking(s).`);
      } else {
        showSyncStatus("success", `Loaded ${n} reservation(s).`);
      }
    })
    .catch((err) => {
      const isNet = err instanceof TypeError && err.message === "Failed to fetch";
      showSyncStatus(
        "error",
        isNet
          ? IS_DEV_MODE
            ? "Could not reach the sheet (use http(s) not file://, and ensure the sheet is shared)."
            : "Could not load reservations. Use the hosted link your team sent you, or try Refresh list."
          : err.message || String(err)
      );
    })
    .finally(() => {
      if (showLoadOnImportButton) setButtonLoading(importBtn, false);
    });
}

function initSheetAppendUrlField() {
  if (!sheetAppendUrlInput) return;
  const saved = localStorage.getItem(LS_SHEET_APPEND_URL);
  if (saved) sheetAppendUrlInput.value = saved;
  const persist = () => {
    localStorage.setItem(LS_SHEET_APPEND_URL, sheetAppendUrlInput.value.trim());
  };
  sheetAppendUrlInput.addEventListener("change", persist);
  sheetAppendUrlInput.addEventListener("blur", persist);
}

/** Browsers block fetching Google URLs when the page is opened as file:// — need http(s). */
function initFileProtocolBanner() {
  if (window.location.protocol !== "file:") return;
  const bar = document.createElement("div");
  bar.className = "file-protocol-banner";
  bar.setAttribute("role", "alert");
  if (IS_DEV_MODE) {
    bar.innerHTML =
      "<strong>This page is opened as a file (file://).</strong> The browser will not load remote data. " +
      "Run <code>npm start</code> in the project folder and open the <code>http://</code> address it prints, or use <strong>Live Server</strong>.";
  } else {
    bar.innerHTML =
      "<strong>This page was opened from your computer as a file.</strong> Open the same demo using the <strong>web link</strong> you were given (it should start with <code>http://</code> or <code>https://</code>).";
  }
  document.body.insertBefore(bar, document.body.firstChild);
}

function initPresentationMode() {
  document.body.classList.add(IS_DEV_MODE ? "mode-dev" : "mode-client");
}

// --- Init ---
initPresentationMode();
initFileProtocolBanner();
initSheetAppendUrlField();
renderBookingsTable();
initRemoteBookings();
