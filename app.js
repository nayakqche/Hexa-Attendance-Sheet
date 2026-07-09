// Hexa Climate — Leave reconciliation between Attendance Sheet and Timesheet.
// Anything in LEAVE_CODES is a leave; anything else in a cell that has a value
// is treated as Present.

const LEAVE_CODES = ["SL", "PL", "UL", "FL", "SPL", "ML", "PT", "BL", "CL"];
const LEAVE_SET = new Set(LEAVE_CODES);

const state = { attendance: null, timesheet: null };
// Every uploaded file, with parse details, for the "Uploaded files" table.
const uploadedFiles = { attendance: [], timesheet: [] };

const els = {
  attendanceFile: document.getElementById("attendanceFile"),
  timesheetFile: document.getElementById("timesheetFile"),
  attendanceInfo: document.getElementById("attendanceInfo"),
  timesheetInfo: document.getElementById("timesheetInfo"),
  attendanceZone: document.getElementById("attendanceZone"),
  timesheetZone: document.getElementById("timesheetZone"),
  reconcileBtn: document.getElementById("reconcileBtn"),
  summary: document.getElementById("summary"),
  results: document.getElementById("results"),
  filesPanel: document.getElementById("filesPanel"),
  filesTable: document.getElementById("filesTable"),
  statEmployees: document.getElementById("statEmployees"),
  statDates: document.getElementById("statDates"),
  statMismatches: document.getElementById("statMismatches"),
  statOnlyIn: document.getElementById("statOnlyIn"),
  genPanel: document.getElementById("genPanel"),
  genList: document.getElementById("genList"),
  genSearch: document.getElementById("genSearch"),
  genSelectAll: document.getElementById("genSelectAll"),
  genClear: document.getElementById("genClear"),
  genProject: document.getElementById("genProject"),
  genBtn: document.getElementById("genBtn"),
  genStatus: document.getElementById("genStatus"),
};

els.attendanceFile.addEventListener("change", (e) => handleFile(e, "attendance"));
els.timesheetFile.addEventListener("change", (e) => handleFile(e, "timesheet"));
els.reconcileBtn.addEventListener("click", runReconciliation);
els.genSearch.addEventListener("input", () => renderGenList());
els.genSelectAll.addEventListener("click", () => toggleAllGen(true));
els.genClear.addEventListener("click", () => toggleAllGen(false));
els.genBtn.addEventListener("click", generateSelected);

async function handleFile(evt, kind) {
  const files = Array.from(evt.target.files || []);
  if (!files.length) return;
  const zone = kind === "attendance" ? els.attendanceZone : els.timesheetZone;
  const info = kind === "attendance" ? els.attendanceInfo : els.timesheetInfo;
  try {
    for (const file of files) {
      const wb = await readWorkbook(file);
      const parsed = parseWorkbook(wb, file.name);
      // Replace any earlier upload with the same name (avoid double counting).
      const dup = uploadedFiles[kind].findIndex((f) => f.name === file.name);
      if (dup >= 0) uploadedFiles[kind].splice(dup, 1);
      // Accumulate across every upload for this side (don't discard earlier files).
      uploadedFiles[kind].push({
        name: file.name,
        size: file.size,
        sheets: wb.SheetNames.length,
        sheetNames: wb.SheetNames.join(", "),
        employees: parsed.rows.length,
        dates: parsed.dates.length,
        dateRange: parsed.dates.length
          ? `${parsed.dates[0]} → ${parsed.dates[parsed.dates.length - 1]}`
          : "—",
        parsed,
      });
    }
    // Rebuild state from ALL files uploaded for this side so far.
    const merged = mergeParsed(uploadedFiles[kind].map((f) => f.parsed));
    state[kind] = { parsed: merged };
    zone.classList.add("has-file");
    info.innerHTML =
      `<b>${uploadedFiles[kind].length} file${uploadedFiles[kind].length > 1 ? "s" : ""} loaded</b> · ` +
      `${merged.rows.length} employees · ${merged.dates.length} dates`;
  } catch (err) {
    console.error(err);
    info.innerHTML = `<span class="err">Parse error: ${err.message}</span>`;
    state[kind] = null;
  }
  renderFilesTable();
  if (kind === "attendance" && state.attendance) populateGenerator();
  els.reconcileBtn.disabled = !(state.attendance && state.timesheet);
}

function fmtSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function renderFilesTable() {
  const all = [
    ...uploadedFiles.attendance.map((f) => ({ ...f, kind: "Attendance" })),
    ...uploadedFiles.timesheet.map((f) => ({ ...f, kind: "Timesheet" })),
  ];
  if (!all.length) {
    els.filesPanel.classList.add("hidden");
    return;
  }
  els.filesPanel.classList.remove("hidden");
  const head =
    `<thead><tr>` +
    ["#", "Type", "File name", "Size", "Sheets", "Employees", "Dates", "Date range", "Status"]
      .map((h) => `<th>${h}</th>`).join("") +
    `</tr></thead>`;
  const body = all.map((f, i) => {
    const ok = f.employees > 0 && f.dates > 0;
    const status = ok
      ? `<span class="tag present">Parsed</span>`
      : `<span class="tag mismatch">No data — check layout</span>`;
    return (
      `<tr>` +
      `<td>${i + 1}</td>` +
      `<td><span class="tag ${f.kind === "Attendance" ? "leave" : "missing"}">${f.kind}</span></td>` +
      `<td title="Sheets: ${f.sheetNames}">${f.name}</td>` +
      `<td>${fmtSize(f.size)}</td>` +
      `<td>${f.sheets}</td>` +
      `<td>${f.employees}</td>` +
      `<td>${f.dates}</td>` +
      `<td>${f.dateRange}</td>` +
      `<td>${status}</td>` +
      `</tr>`
    );
  }).join("");
  els.filesTable.innerHTML = head + `<tbody>${body}</tbody>`;
}

function mergeParsed(parsedList) {
  const empMap = new Map();
  const dateSet = new Set();
  for (const p of parsedList) {
    for (const { employee, name, marks, raw, title } of p.rows) {
      if (!empMap.has(employee)) empMap.set(employee, { name: "", marks: {}, raw: {}, title: "" });
      const e = empMap.get(employee);
      Object.assign(e.marks, marks);
      if (raw) Object.assign(e.raw, raw);
      if (title && !e.title) e.title = title;
      if (name && !e.name) e.name = name;
    }
    for (const d of p.dates) dateSet.add(d);
  }
  return {
    rows: [...empMap.entries()].map(([employee, e]) => ({
      employee, name: e.name || employee, marks: e.marks, raw: e.raw, title: e.title,
    })),
    dates: [...dateSet].sort(),
  };
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error("File read failed"));
    reader.readAsArrayBuffer(file);
  });
}

// Extracts { rows: [{employee, marks: { 'YYYY-MM-DD': 'SL'|'P'|... }}], dates: [...] }
// Handles three common shapes:
//   (A) Wide: employee-per-row, dates across columns
//   (B) Long: one row per employee-date with an explicit Date column
//   (C) Vertical single-employee timesheet: name in a header cell, dates down
//       column A, leave code in the Start/End Time columns
function parseWorkbook(wb, fileName = "") {
  const allRows = [];
  const allDatesSet = new Set();

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    if (!aoa.length) continue;

    // Shape C — try the vertical single-employee timesheet first.
    const vertical = parseVerticalTimesheet(aoa, fileName);
    if (vertical && Object.keys(vertical.marks).length) {
      allRows.push({ employee: vertical.employee, marks: vertical.marks });
      Object.keys(vertical.marks).forEach((d) => allDatesSet.add(d));
      continue;
    }

    const { headerRowIdx, headers } = findHeaderRow(aoa);
    if (headerRowIdx < 0) continue;

    const empColIdx = findEmployeeCol(headers);
    const titleColIdx = findTitleCol(headers);
    const dateCols = findDateCols(headers);
    const longDateColIdx = findLongDateCol(headers);
    const longValueColIdx = findLongValueCol(headers);

    if (empColIdx < 0) continue;

    if (dateCols.length > 0) {
      // WIDE shape
      for (let r = headerRowIdx + 1; r < aoa.length; r++) {
        const row = aoa[r];
        const emp = normEmp(row[empColIdx]);
        if (!emp) continue;
        const marks = {};
        const raw = {};
        for (const dc of dateCols) {
          const val = row[dc.colIdx];
          const mark = classify(val);
          if (mark !== null) marks[dc.iso] = mark;
          raw[dc.iso] = String(val == null ? "" : val).trim();
        }
        const title = titleColIdx >= 0 ? String(row[titleColIdx] || "").trim() : "";
        const name = String(row[empColIdx] || "").trim();
        if (Object.keys(marks).length) {
          allRows.push({ employee: emp, name, marks, raw, title });
          Object.keys(marks).forEach((d) => allDatesSet.add(d));
        }
      }
    } else if (longDateColIdx >= 0 && longValueColIdx >= 0) {
      // LONG shape
      const byEmp = new Map();
      for (let r = headerRowIdx + 1; r < aoa.length; r++) {
        const row = aoa[r];
        const emp = normEmp(row[empColIdx]);
        const iso = toISODate(row[longDateColIdx]);
        if (!emp || !iso) continue;
        const mark = classify(row[longValueColIdx]);
        if (mark === null) continue;
        if (!byEmp.has(emp)) byEmp.set(emp, {});
        byEmp.get(emp)[iso] = mark;
        allDatesSet.add(iso);
      }
      for (const [emp, marks] of byEmp) allRows.push({ employee: emp, marks });
    }
  }

  // Merge duplicate employee rows across sheets (e.g. one sheet per month)
  const merged = new Map();
  for (const { employee, name, marks, raw, title } of allRows) {
    if (!merged.has(employee)) merged.set(employee, { name: "", marks: {}, raw: {}, title: "" });
    const e = merged.get(employee);
    Object.assign(e.marks, marks);
    if (raw) Object.assign(e.raw, raw);
    if (title && !e.title) e.title = title;
    if (name && !e.name) e.name = name;
  }

  return {
    rows: [...merged.entries()].map(([employee, e]) => ({
      employee, name: e.name || employee, marks: e.marks, raw: e.raw, title: e.title,
    })),
    dates: [...allDatesSet].sort(),
  };
}

function findTitleCol(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toLowerCase();
    if (/^(job\s*title|designation|title|role)$/.test(h)) return i;
  }
  return -1;
}

// Shape C: a single-employee monthly timesheet.
//   - Employee name lives in a cell next to an "Employee Name" label (or, failing
//     that, is derived from the file name).
//   - A header row contains "Date" plus "Day"/"Start Time"/"Total Hours".
//   - Dates run down the Date column; the leave code (e.g. SL) appears in one of
//     the time/hours cells on that row. Anything that isn't a leave code = Present.
function parseVerticalTimesheet(aoa, fileName) {
  const scan = Math.min(aoa.length, 12);

  // 1) Employee name from an "Employee Name" label cell.
  let employee = "";
  for (let r = 0; r < scan && !employee; r++) {
    const row = aoa[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (/employee\s*name/i.test(String(row[c] || ""))) {
        for (let k = c + 1; k < row.length; k++) {
          const v = String(row[k] || "").trim();
          if (v) { employee = v; break; }
        }
      }
      if (employee) break;
    }
  }
  if (!employee) employee = empNameFromFile(fileName);
  if (!employee) return null;

  // 2) Header row with a "Date" column.
  let headerRowIdx = -1, dateColIdx = -1, dayColIdx = -1;
  for (let r = 0; r < Math.min(aoa.length, 15); r++) {
    const row = (aoa[r] || []).map((c) => String(c || "").trim().toLowerCase());
    const dc = row.findIndex((c) => c === "date");
    const looksLikeTimesheet = row.some((c) => /day|start time|end time|total hours/.test(c));
    if (dc >= 0 && looksLikeTimesheet) {
      headerRowIdx = r;
      dateColIdx = dc;
      dayColIdx = row.findIndex((c) => c === "day");
      break;
    }
  }
  if (headerRowIdx < 0 || dateColIdx < 0) return null;

  // 3) Walk the date rows; a leave code in any non-date/day cell wins.
  const marks = {};
  for (let r = headerRowIdx + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    const iso = toISODate(row[dateColIdx]);
    if (!iso) continue;
    let leave = null;
    for (let c = 0; c < row.length; c++) {
      if (c === dateColIdx || c === dayColIdx) continue;
      const code = leaveCodeOf(row[c]);
      if (code) { leave = code; break; }
    }
    marks[iso] = leave || "P";
  }
  return { employee: normEmp(employee), marks };
}

// Returns a leave code if the cell holds one, otherwise null (never "P").
function leaveCodeOf(v) {
  const m = classify(v);
  return LEAVE_SET.has(m) ? m : null;
}

// Best-effort employee name from a file name like "Apoorv Sohoni April 2025 (1).xlsx".
function empNameFromFile(fileName) {
  if (!fileName) return "";
  let s = fileName.replace(/\.[^.]+$/, "");           // drop extension
  s = s.replace(/\(\d+\)/g, " ");                       // drop "(1)"
  s = s.replace(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/gi, " ");
  s = s.replace(/\b(19|20)\d{2}\b/g, " ");              // drop years
  s = s.replace(/\b(time\s*sheet|timesheet|attendance)\b/gi, " ");
  return s.replace(/[_\-]+/g, " ").replace(/\s+/g, " ").trim();
}

// Scan first ~15 rows for the one that contains an employee-like header.
function findHeaderRow(aoa) {
  const scanLimit = Math.min(aoa.length, 15);
  for (let i = 0; i < scanLimit; i++) {
    const row = aoa[i].map((c) => String(c || "").trim().toLowerCase());
    const looksLikeHeader = row.some((c) => /employee|emp\s*(id|name|code)|name|staff/.test(c));
    if (looksLikeHeader) return { headerRowIdx: i, headers: aoa[i] };
  }
  // Fallback: first non-empty row
  for (let i = 0; i < scanLimit; i++) {
    if (aoa[i].some((c) => String(c || "").trim())) return { headerRowIdx: i, headers: aoa[i] };
  }
  return { headerRowIdx: -1, headers: [] };
}

function findEmployeeCol(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toLowerCase();
    if (/^(employee|emp)(\s*(id|code|name|no))?$|^name$|^staff/.test(h)) return i;
  }
  return -1;
}

function findDateCols(headers) {
  const out = [];
  for (let i = 0; i < headers.length; i++) {
    const iso = toISODate(headers[i]);
    if (iso) out.push({ colIdx: i, iso });
  }
  return out;
}

function findLongDateCol(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toLowerCase();
    if (/^date$/.test(h)) return i;
  }
  return -1;
}
function findLongValueCol(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim().toLowerCase();
    if (/status|attendance|leave|mark|type/.test(h)) return i;
  }
  return -1;
}

function normEmp(v) {
  if (v == null) return "";
  return String(v).trim().replace(/\s+/g, " ").toUpperCase();
}

// Convert any header/cell into an ISO date, or return null.
function toISODate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return isoOf(v);
  if (typeof v === "number" && v > 20000 && v < 80000) {
    // Excel serial
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${pad(d.y)}-${pad(d.m)}-${pad(d.d)}`;
  }
  const s = String(v).trim();
  // dd/mm/yyyy or dd-mm-yyyy
  let m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
  if (m) {
    let [_, d, mo, y] = m;
    if (y.length === 2) y = "20" + y;
    return `${y}-${pad(mo)}-${pad(d)}`;
  }
  // yyyy-mm-dd
  m = s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  // Try native
  const d = new Date(s);
  if (!isNaN(d) && /\d/.test(s) && s.length >= 6) return isoOf(d);
  return null;
}
const isoOf = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const pad = (n) => String(n).padStart(2, "0");

// Returns a leave code string (e.g. "SL"), "P" for present, or null for empty.
function classify(v) {
  if (v == null || v === "") return null;
  const s = String(v).trim().toUpperCase();
  if (!s || s === "-" || s === "N/A") return null;
  // Sometimes cells are like "SL - Sick Leave" or "1 SL". Extract the code.
  for (const code of LEAVE_CODES) {
    const re = new RegExp(`(^|[^A-Z])${code}([^A-Z]|$)`);
    if (re.test(s)) return code;
  }
  return "P";
}

function runReconciliation() {
  els.results.innerHTML = "";
  const A = indexParsed(state.attendance.parsed);
  const T = indexParsed(state.timesheet.parsed);

  // Only reconcile the employees whose timesheets were uploaded — not everyone
  // in the attendance sheet.
  const mismatches = [];
  const missing = [];
  const unmatched = [];
  const dateSet = new Set();
  const comparedEmployees = new Set();

  for (const [emp, tMarks] of T) {
    const matchKey = matchAttendance(emp, A);
    if (!matchKey) {
      unmatched.push({ employee: emp });
      continue;
    }
    comparedEmployees.add(emp);
    const aMarks = A.get(matchKey);
    const dates = new Set([...Object.keys(aMarks), ...Object.keys(tMarks)]);
    for (const d of dates) {
      dateSet.add(d);
      const a = aMarks[d];
      const t = tMarks[d];
      if (a == null && t != null) {
        missing.push({ employee: emp, date: d, attendance: "—", timesheet: label(t), reason: "Missing in Attendance" });
      } else if (t == null && a != null) {
        missing.push({ employee: emp, date: d, attendance: label(a), timesheet: "—", reason: "Missing in Timesheet" });
      } else if (a != null && t != null && a !== t) {
        mismatches.push({
          employee: emp, date: d,
          attendance: label(a), timesheet: label(t),
          note: mismatchNote(a, t),
        });
      }
    }
  }

  els.summary.classList.remove("hidden");
  els.statEmployees.textContent = comparedEmployees.size;
  els.statDates.textContent = dateSet.size;
  els.statMismatches.textContent = mismatches.length;
  els.statOnlyIn.textContent = missing.length;

  renderTable("Leave Mismatches", ["Employee", "Date", "Attendance Sheet", "Timesheet", "Where it's missing"],
    mismatches.map((m) => [m.employee, m.date, tag(m.attendance), tag(m.timesheet), m.note]),
    mismatches, "mismatches.csv");

  renderTable("Missing Dates (matched employees only)", ["Employee", "Date", "Attendance Sheet", "Timesheet", "Note"],
    missing.map((m) => [m.employee, m.date, tag(m.attendance), tag(m.timesheet), m.reason]),
    missing, "missing.csv");

  if (unmatched.length) {
    renderTable("Timesheet employees not found in Attendance", ["Employee"],
      unmatched.map((m) => [m.employee]), unmatched, "unmatched.csv");
  }
}

// Find the attendance key for a timesheet employee. Tries, in order:
//   1. exact (normalized) match
//   2. whole-string fuzzy match ("Aapoorv" vs "Apoorv")
//   3. token-subset match ("Karunakaran" ⊆ "Karunakaran Thangaraj") — only when
//      exactly one attendance name qualifies, to avoid grabbing the wrong person.
function matchAttendance(emp, A) {
  if (A.has(emp)) return emp;
  const norm = (s) => s.replace(/[^A-Z0-9]/gi, "").toUpperCase();
  const target = norm(emp);
  if (!target) return null;

  // 2) whole-string fuzzy
  let best = null, bestDist = Infinity;
  for (const key of A.keys()) {
    const k = norm(key);
    if (k === target) return key;
    const d = levenshtein(target, k);
    if (d < bestDist) { bestDist = d; best = key; }
  }
  if (best && bestDist <= Math.max(1, Math.floor(target.length * 0.12))) return best;

  // 3) token-subset (unique only)
  const tokensOf = (s) => String(s).toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
  const tTokens = tokensOf(emp);
  if (tTokens.length) {
    const candidates = [];
    for (const key of A.keys()) {
      const aTokens = tokensOf(key);
      const allIn = tTokens.every((tt) =>
        aTokens.some((at) => at === tt || levenshtein(at, tt) <= 1));
      if (allIn) candidates.push(key);
    }
    if (candidates.length === 1) return candidates[0];
  }
  return null;
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  const dp = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return dp[m][n];
}

function indexParsed(parsed) {
  const m = new Map();
  for (const { employee, marks } of parsed.rows) m.set(employee, marks);
  return m;
}

function label(code) {
  if (code === "P") return "Present";
  return code;
}

// Explain a mismatch: which file is missing the leave, or that the codes differ.
// a = attendance mark, t = timesheet mark ("P" = present, else a leave code).
function mismatchNote(a, t) {
  const aLeave = a !== "P";
  const tLeave = t !== "P";
  if (tLeave && !aLeave) return `${t} not marked in Attendance Sheet`;
  if (aLeave && !tLeave) return `${a} not marked in Timesheet`;
  // Both are leaves but different codes.
  return `Different leave codes (Attendance ${a} vs Timesheet ${t})`;
}

function tag(text) {
  const cls = text === "Present" ? "present" : text === "—" ? "missing" : "leave";
  return `<span class="tag ${cls}">${text}</span>`;
}

function renderTable(title, headers, rows, rawData, downloadName) {
  const container = document.createElement("div");
  container.style.marginBottom = "20px";
  container.className = "results-block";
  container.style.background = "#fff";
  container.style.border = "1px solid #e5e8ef";
  container.style.borderRadius = "10px";
  container.style.overflow = "hidden";
  container.innerHTML = `<h3>${title} <span style="color:#6b7383;font-weight:400;font-size:13px;">(${rows.length})</span></h3>`;

  if (rows.length === 0) {
    container.innerHTML += `<div style="padding:18px;color:#6b7383;font-size:13px;">None found.</div>`;
    els.results.appendChild(container);
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr>`;
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = row.map((c) => `<td>${c}</td>`).join("");
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  container.appendChild(table);

  const dl = document.createElement("div");
  dl.className = "download-row";
  const btn = document.createElement("button");
  btn.textContent = "Download CSV";
  btn.onclick = () => downloadCSV(rawData, downloadName);
  dl.appendChild(btn);
  container.appendChild(dl);

  els.results.appendChild(container);
}

function downloadCSV(rows, name) {
  if (!rows.length) return;
  const keys = Object.keys(rows[0]);
  const csv = [keys.join(",")].concat(
    rows.map((r) => keys.map((k) => `"${String(r[k]).replace(/"/g, '""')}"`).join(","))
  ).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
}

/* ------------------------------------------------------------------ *
 *  Timesheet generator: build a per-employee monthly timesheet .xlsx
 *  that replicates the company template, with leaves from attendance.
 * ------------------------------------------------------------------ */

// Attendance codes that mean a week-off or a holiday (leaves are handled
// separately via leaveCodeOf).
const WEEKOFF_RE = /^(w\.?\s*o\.?|w\.?\s*off|wo|week\s*off|weekly\s*off|off)$/i;
const HOLIDAY_RE = /^(h|hol|holiday|ph|public\s*holiday)$/i;

function populateGenerator() {
  const rows = state.attendance.parsed.rows;
  if (!rows.length) { els.genPanel.classList.add("hidden"); return; }
  els.genPanel.classList.remove("hidden");
  renderGenList();
}

function renderGenList() {
  const rows = state.attendance.parsed.rows.slice().sort((a, b) =>
    (a.name || a.employee).localeCompare(b.name || b.employee));
  const q = (els.genSearch.value || "").trim().toLowerCase();
  // Preserve currently-checked employees across re-renders.
  const checked = new Set(
    [...els.genList.querySelectorAll("input:checked")].map((c) => c.value));
  const shown = rows.filter((r) =>
    !q || (r.name || r.employee).toLowerCase().includes(q));
  els.genList.innerHTML = shown.map((r) => {
    const key = r.employee;
    const disp = r.name || r.employee;
    const on = checked.has(key) ? "checked" : "";
    return `<label class="gen-item"><input type="checkbox" value="${escapeAttr(key)}" ${on}/> ${escapeHtml(disp)}</label>`;
  }).join("") || `<div class="gen-empty">No employees match.</div>`;
  els.genList.querySelectorAll("input").forEach((c) =>
    c.addEventListener("change", updateGenButton));
  updateGenButton();
}

function toggleAllGen(on) {
  els.genList.querySelectorAll("input").forEach((c) => { c.checked = on; });
  updateGenButton();
}

function updateGenButton() {
  const n = els.genList.querySelectorAll("input:checked").length;
  els.genBtn.disabled = n === 0;
  els.genBtn.textContent = n > 1 ? `Generate & Download (${n})` : "Generate & Download";
}

async function generateSelected() {
  const keys = [...els.genList.querySelectorAll("input:checked")].map((c) => c.value);
  if (!keys.length) return;
  const byKey = new Map(state.attendance.parsed.rows.map((r) => [r.employee, r]));
  const month = detectMonth(state.attendance.parsed.dates);
  if (!month) { els.genStatus.textContent = "Couldn't detect the month from attendance dates."; return; }
  const project = (els.genProject.value || "").trim();
  els.genBtn.disabled = true;
  let done = 0;
  for (const key of keys) {
    const emp = byKey.get(key);
    if (!emp) continue;
    els.genStatus.textContent = `Generating ${done + 1} / ${keys.length}…`;
    try {
      downloadTimesheet(emp, month.year, month.month0, project);
      done++;
      await sleep(350); // let each download start before the next
    } catch (err) {
      console.error(err);
      els.genStatus.innerHTML = `<span class="err">Failed on ${escapeHtml(emp.name || key)}: ${err.message}</span>`;
    }
  }
  els.genStatus.textContent = `Done — generated ${done} timesheet${done === 1 ? "" : "s"} for ${monthLong(month.year, month.month0)} ${month.year}.`;
  updateGenButton();
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// Pick the (year, month) that most of the attendance dates fall in.
function detectMonth(dates) {
  const counts = {};
  for (const d of dates) counts[d.slice(0, 7)] = (counts[d.slice(0, 7)] || 0) + 1;
  let ym = null, best = -1;
  for (const [k, c] of Object.entries(counts)) if (c > best) { best = c; ym = k; }
  if (!ym) return null;
  const [y, m] = ym.split("-").map(Number);
  return { year: y, month0: m - 1 };
}

function dayInfo(rawCode, date) {
  const s = String(rawCode == null ? "" : rawCode).trim();
  const leave = leaveCodeOf(s);
  if (leave) return { kind: "leave", code: leave };
  if (HOLIDAY_RE.test(s)) return { kind: "holiday" };
  const dow = date.getDay();
  if (dow === 0 || dow === 6 || WEEKOFF_RE.test(s)) return { kind: "weekoff" };
  return { kind: "work" };
}

function rowCells(info) {
  if (info.kind === "leave")   return { s: info.code, e: info.code, reg: "0:00", ot: "0:00", tot: "-", act: "-", hrs: 0 };
  if (info.kind === "holiday") return { s: "Holiday", e: "Holiday", reg: "0:00", ot: "0:00", tot: "-", act: "-", hrs: 0 };
  if (info.kind === "weekoff") return { s: "W. Off", e: "W. Off", reg: "0:00", ot: "0:00", tot: "-", act: "-", hrs: 0 };
  return { s: "9:30 AM", e: "6:30 AM", reg: "09:00", ot: "0:00", tot: "9.00", act: "9.00", hrs: 9 };
}

const longDate = (d) => d.toLocaleDateString("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
const weekdayName = (d) => d.toLocaleDateString("en-US", { weekday: "long" });
const monthShortLabel = (y, m0) => new Date(y, m0, 1).toLocaleDateString("en-US", { month: "short" }) + "-" + String(y).slice(-2);
const monthLong = (y, m0) => new Date(y, m0, 1).toLocaleDateString("en-US", { month: "long" });

function buildTimesheetAOA(emp, year, m0, project) {
  const projHeader = project ? "Project " + project : "Project";
  const aoa = [
    ["Employee Name:", emp.name || emp.employee],
    ["Title:", emp.title || ""],
    ["Month:", monthShortLabel(year, m0)],
    [],
    ["Date", "Day", "Start Time", "End Time", "Regular Hours", "Overtime Hours", "Total Hours", projHeader],
    ["", "", "", "", "", "", "", "On an actual basis"],
  ];
  const days = new Date(year, m0 + 1, 0).getDate();
  let sum = 0;
  for (let d = 1; d <= days; d++) {
    const date = new Date(year, m0, d);
    const iso = `${year}-${pad(m0 + 1)}-${pad(d)}`;
    const c = rowCells(dayInfo(emp.raw ? emp.raw[iso] : "", date));
    sum += c.hrs;
    aoa.push([longDate(date), weekdayName(date), c.s, c.e, c.reg, c.ot, c.tot, c.act]);
  }
  const tot = sum.toFixed(2);
  aoa.push(["", "", "", "", "", "", tot, tot]);
  return aoa;
}

function downloadTimesheet(emp, year, m0, project) {
  const aoa = buildTimesheetAOA(emp, year, m0, project);
  const XL = window.__XLSX_STYLE || XLSX; // styled writer if available
  const ws = XL.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 28 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 15 }, { wch: 12 }, { wch: 16 }];

  const lastRow = aoa.length - 1;
  ws["!merges"] = [];
  for (let c = 0; c <= 6; c++) ws["!merges"].push({ s: { r: 4, c }, e: { r: 5, c } });
  ws["!merges"].push({ s: { r: lastRow, c: 0 }, e: { r: lastRow, c: 5 } });

  applyTimesheetStyles(ws, aoa, XL);

  const wb = XL.utils.book_new();
  XL.utils.book_append_sheet(wb, ws, "Timesheet");
  const fname = `${emp.name || emp.employee}_${monthLong(year, m0)} ${year}.xlsx`;
  XL.writeFile(wb, fname);
}

// Borders, gray header, bold labels — ignored gracefully if the styled writer
// isn't loaded (data + merges still come through).
function applyTimesheetStyles(ws, aoa, XL) {
  const thin = { style: "thin", color: { rgb: "000000" } };
  const border = { top: thin, bottom: thin, left: thin, right: thin };
  const setStyle = (r, c, style) => {
    const ref = XL.utils.encode_cell({ r, c });
    if (!ws[ref]) ws[ref] = { t: "s", v: "" };
    ws[ref].s = Object.assign({}, ws[ref].s, style);
  };
  const lastRow = aoa.length - 1;
  for (let r = 4; r <= lastRow; r++) {
    for (let c = 0; c <= 7; c++) {
      const header = r === 4 || r === 5;
      setStyle(r, c, {
        border,
        alignment: { horizontal: "center", vertical: "center", wrapText: true },
        ...(header ? { font: { bold: true }, fill: { fgColor: { rgb: "D9D9D9" } } } : {}),
      });
    }
  }
  [0, 1, 2].forEach((r) => setStyle(r, 0, { font: { bold: true } }));
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}
function escapeAttr(s) { return escapeHtml(s); }
