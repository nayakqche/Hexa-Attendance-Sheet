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
};

els.attendanceFile.addEventListener("change", (e) => handleFile(e, "attendance"));
els.timesheetFile.addEventListener("change", (e) => handleFile(e, "timesheet"));
els.reconcileBtn.addEventListener("click", runReconciliation);

async function handleFile(evt, kind) {
  const files = Array.from(evt.target.files || []);
  if (!files.length) return;
  const zone = kind === "attendance" ? els.attendanceZone : els.timesheetZone;
  const info = kind === "attendance" ? els.attendanceInfo : els.timesheetInfo;
  try {
    const perFile = [];
    for (const file of files) {
      const wb = await readWorkbook(file);
      const parsed = parseWorkbook(wb);
      perFile.push({ name: file.name, parsed });
      // Record full details for the uploaded-files table.
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
      });
    }
    const merged = mergeParsed(perFile.map((p) => p.parsed));
    state[kind] = { files, parsed: merged };
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
    for (const { employee, marks } of p.rows) {
      if (!empMap.has(employee)) empMap.set(employee, {});
      Object.assign(empMap.get(employee), marks);
    }
    for (const d of p.dates) dateSet.add(d);
  }
  return {
    rows: [...empMap.entries()].map(([employee, marks]) => ({ employee, marks })),
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
// Handles two common shapes:
//   (A) Wide: employee-per-row, dates across columns
//   (B) Long: one row per employee-date with an explicit Date column
function parseWorkbook(wb) {
  const allRows = [];
  const allDatesSet = new Set();

  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    if (!aoa.length) continue;

    const { headerRowIdx, headers } = findHeaderRow(aoa);
    if (headerRowIdx < 0) continue;

    const empColIdx = findEmployeeCol(headers);
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
        for (const dc of dateCols) {
          const val = row[dc.colIdx];
          const mark = classify(val);
          if (mark !== null) marks[dc.iso] = mark;
        }
        if (Object.keys(marks).length) {
          allRows.push({ employee: emp, marks });
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
  for (const { employee, marks } of allRows) {
    if (!merged.has(employee)) merged.set(employee, {});
    Object.assign(merged.get(employee), marks);
  }

  return {
    rows: [...merged.entries()].map(([employee, marks]) => ({ employee, marks })),
    dates: [...allDatesSet].sort(),
  };
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

  const allEmployees = new Set([...A.keys(), ...T.keys()]);
  const mismatches = [];
  const missing = [];
  const dateUnion = new Set();
  let comparedCells = 0;

  for (const emp of allEmployees) {
    const aMarks = A.get(emp) || {};
    const tMarks = T.get(emp) || {};
    const dates = new Set([...Object.keys(aMarks), ...Object.keys(tMarks)]);
    for (const d of dates) {
      dateUnion.add(d);
      const a = aMarks[d];
      const t = tMarks[d];
      if (a == null && t != null) {
        missing.push({ employee: emp, date: d, attendance: "—", timesheet: label(t), reason: "Missing in Attendance" });
      } else if (t == null && a != null) {
        missing.push({ employee: emp, date: d, attendance: label(a), timesheet: "—", reason: "Missing in Timesheet" });
      } else if (a != null && t != null) {
        comparedCells++;
        if (a !== t) {
          mismatches.push({ employee: emp, date: d, attendance: label(a), timesheet: label(t) });
        }
      }
    }
  }

  els.summary.classList.remove("hidden");
  els.statEmployees.textContent = allEmployees.size;
  els.statDates.textContent = dateUnion.size;
  els.statMismatches.textContent = mismatches.length;
  els.statOnlyIn.textContent = missing.length;

  renderTable("Leave Mismatches", ["Employee", "Date", "Attendance Sheet", "Timesheet"],
    mismatches.map((m) => [m.employee, m.date, tag(m.attendance), tag(m.timesheet)]),
    mismatches, "mismatches.csv");

  renderTable("Missing in One File", ["Employee", "Date", "Attendance Sheet", "Timesheet", "Note"],
    missing.map((m) => [m.employee, m.date, tag(m.attendance), tag(m.timesheet), m.reason]),
    missing, "missing.csv");
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
