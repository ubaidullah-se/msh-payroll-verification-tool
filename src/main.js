import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════════════════════
let FILES = [],
  parsedData = {},
  audit = {},
  logs = [];

// Mapping state: keyed by normalised TS name
// tsNormKey → { payrollKey, dept }
const mappingStore = new Map();

// ═══════════════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════════════
const ROLES = [
  { v: "week1", l: "Week 1 Timesheet" },
  { v: "week2", l: "Week 2 Timesheet" },
  { v: "invoice_msh", l: "MSH Invoice (client invoice)" },
  { v: "invoice_ys", l: "Yellowstone Invoice" },
  { v: "payroll", l: "Internal Payroll Summary" },
  { v: "adp", l: "ADP Export (CSV)" },
  { v: "ignore", l: "— Ignore —" },
];

const DEPT_NORM = {
  housekeeping: "housekeeper",
  "housekeeping scottsdale": "housekeeper",
  administrador: "administrador",
  proyec: "yellowstone",
};

const INV_DESC_MAP = {
  housekeeper: "housekeeper",
  "housekeeper ot": "housekeeper_ot",
  cook: "cook",
  "cook ot": "cook_ot",
  dishwasher: "dishwasher",
  maintenance: "maintenance",
  "pool attendant": "pool_attendant",
  "pool attendant ot": "pool_attendant_ot",
  "housekeeper inspector": "housekeeper_inspector",
  "housekeeper inspector ot": "housekeeper_inspector_ot",
  casida: "housekeeper",
  "casida ot": "housekeeper_ot",
  lobby: "housekeeper",
  "lobby ot": "housekeeper_ot",
  villa: "housekeeper",
  "villa ot": "housekeeper_ot",
};

const YELLOWSTONE_DEPTS = new Set(["yellowstone", "proyec"]);

// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════
const log = (lvl, msg) => logs.push({ lvl, msg });
const toNum = (v) => {
  const n = parseFloat(String(v ?? "").replace(/[$,]/g, ""));
  return isNaN(n) ? 0 : n;
};
const norm = (s) =>
  String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^a-z0-9 ]/g, "");
const fmtH = (n) =>
  n === null || n === undefined ? "—" : Number(n).toFixed(2);
const diffCls = (d) =>
  d === null ? "" : Math.abs(d) < 0.05 ? "dz" : d > 0 ? "dp" : "dn";
const capName = (s) =>
  String(s || "")
    .split(" ")
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");

function detectRole(name) {
  const n = name.toLowerCase();
  if (n.includes("epi0") || n.endsWith(".csv") || n.includes("adp"))
    return "adp";
  if (n.includes("buttes_pay") || n.includes("payroll")) return "payroll";
  if (n.includes("yellowstone")) return "invoice_ys";
  if (
    n.includes("invoice_69") ||
    n.includes("invoice_msh") ||
    (n.includes("invoice") && n.includes("msh"))
  )
    return "invoice_msh";
  if (
    n.includes("week_1") ||
    n.includes("_hk") ||
    n.includes("hospitality_week")
  )
    return "week1";
  if (n.includes("proyec") || n.includes("hospital")) return "week2";
  return "ignore";
}

function getRoleMap() {
  const rm = {};
  FILES.forEach((f, i) => {
    const role = document.getElementById(`r${i}`).value;
    if (role !== "ignore") {
      if (!rm[role]) rm[role] = [];
      rm[role].push(f);
    }
  });
  return rm;
}

function readBuf(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = (e) => res(e.target.result);
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}
async function toRaw(file) {
  const buf = await readBuf(file);
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: null,
    raw: true,
  });
}

function setStep(n) {
  [1, 2, 3].forEach((i) => {
    const el = document.getElementById(`stp${i}`);
    el.className = "step" + (i < n ? " done" : i === n ? " active" : "");
  });
}

// ═══════════════════════════════════════════════════════════════
// FILE UPLOAD / ROLE UI  (Step 1)
// ═══════════════════════════════════════════════════════════════
const dz = document.getElementById("dropZone");
const fi = document.getElementById("fileInput");
dz.addEventListener("dragover", (e) => {
  e.preventDefault();
  dz.classList.add("drag-over");
});
dz.addEventListener("dragleave", () => dz.classList.remove("drag-over"));
dz.addEventListener("drop", (e) => {
  e.preventDefault();
  dz.classList.remove("drag-over");
  handleFiles([...e.dataTransfer.files]);
});
fi.addEventListener("change", () => handleFiles([...fi.files]));

function handleFiles(files) {
  FILES = files;
  const grid = document.getElementById("roleGrid");
  grid.innerHTML = "";
  files.forEach((f, i) => {
    const role = detectRole(f.name);
    const ext = f.name.split(".").pop().toUpperCase();
    const item = document.createElement("div");
    item.className = "role-item";
    item.innerHTML = `<span style="font-size:18px">${ext === "CSV" ? "📄" : "📊"}</span>
      <span class="fname" title="${f.name}">${f.name}</span>
      <select id="r${i}">${ROLES.map((r) => `<option value="${r.v}"${r.v === role ? " selected" : ""}>${r.l}</option>`).join("")}</select>`;
    grid.appendChild(item);
  });
  document.getElementById("rolePanel").style.display = "block";
  setStep(1);
}

function resetAll() {
  FILES = [];
  parsedData = {};
  audit = {};
  logs = [];
  mappingStore.clear();
  fi.value = "";
  ["rolePanel", "mappingPanel", "results"].forEach(
    (id) => (document.getElementById(id).style.display = "none"),
  );
  document.getElementById("roleGrid").innerHTML = "";
  document.getElementById("mapRows").innerHTML = "";
  document.getElementById("alertBanners").innerHTML = "";
  setStep(1);
}

// ═══════════════════════════════════════════════════════════════
// MAPPING UI  (Step 2)
// ═══════════════════════════════════════════════════════════════
async function showMapping() {
  const btn = document.getElementById("toMappingBtn");
  btn.disabled = true;
  btn.textContent = "⏳ Loading…";

  try {
    logs = [];
    const rm = getRoleMap();

    // Parse timesheets to collect all employee names
    const tsEmps = new Map(); // normKey → displayName
    const parseAndCollect = async (files, label) => {
      for (const f of files || []) {
        const r = parseTimesheet(await toRaw(f), `${label}:${f.name}`);
        Object.entries(r.employees).forEach(([k, v]) => {
          if (!tsEmps.has(k)) tsEmps.set(k, v.name);
        });
      }
    };
    await parseAndCollect(rm["week1"], "W1");
    await parseAndCollect(rm["week2"], "W2");

    // Parse payroll to get name options and depts
    let payrollEmps = {}; // normKey → {name, dept, ...}
    for (const f of rm["payroll"] || []) {
      const pd = parsePayroll(await toRaw(f), `Payroll:${f.name}`);
      Object.assign(payrollEmps, pd.employees);
    }

    const payKeys = Object.keys(payrollEmps);
    const allDepts = [
      ...new Set(payKeys.map((k) => payrollEmps[k].dept).filter(Boolean)),
    ].sort();

    // Build mapping rows
    const rowsEl = document.getElementById("mapRows");
    rowsEl.innerHTML = "";
    mappingStore.clear();

    // Use index-based DOM ids to avoid special-char issues in selectors
    let idx = 0;
    tsEmps.forEach((displayName, tsKey) => {
      const rowId = idx++;
      const autoPayKey =
        payKeys.find((pk) => norm(pk) === tsKey) ||
        payKeys.find((pk) => norm(pk).includes(tsKey.split(" ")[0])) ||
        "";
      const autoDept = autoPayKey
        ? payrollEmps[autoPayKey].dept
        : allDepts[0] || "";

      // Persist initial values
      mappingStore.set(tsKey, { payrollKey: autoPayKey, dept: autoDept });

      const payOpts =
        `<option value="">— Not in payroll —</option>` +
        payKeys
          .map(
            (pk) =>
              `<option value="${pk}"${pk === autoPayKey ? " selected" : ""}>${capName(pk)}</option>`,
          )
          .join("");
      const deptOpts = allDepts
        .map(
          (d) =>
            `<option value="${d}"${d === autoDept ? " selected" : ""}>${d}</option>`,
        )
        .join("");

      const row = document.createElement("div");
      row.className = "map-row";
      row.dataset.tskey = tsKey;
      row.dataset.rowid = rowId;
      row.innerHTML = `
        <div>
          <div class="map-name">${capName(displayName)}</div>
          <div class="map-hint">${tsKey}</div>
        </div>
        <select id="mpay_${rowId}" data-tskey="${tsKey}">${payOpts}</select>
        <select id="mdept_${rowId}" data-tskey="${tsKey}">${deptOpts}</select>`;

      // When payroll name changes → auto-update dept
      row.querySelector(`#mpay_${rowId}`).addEventListener("change", (e) => {
        const pk = e.target.value;
        const dept = pk && payrollEmps[pk] ? payrollEmps[pk].dept : "";
        const deptSel = document.getElementById(`mdept_${rowId}`);
        if (dept) deptSel.value = dept;
        mappingStore.set(tsKey, {
          payrollKey: pk,
          dept: deptSel.value,
        });
      });
      row.querySelector(`#mdept_${rowId}`).addEventListener("change", (e) => {
        const cur = mappingStore.get(tsKey) || {};
        mappingStore.set(tsKey, { ...cur, dept: e.target.value });
      });

      rowsEl.appendChild(row);
    });

    document.getElementById("rolePanel").style.display = "none";
    document.getElementById("mappingPanel").style.display = "block";
    setStep(2);
  } catch (err) {
    alert("Error loading mapping: " + err.message);
    console.error(err);
  }

  btn.disabled = false;
  btn.textContent = "➡ Proceed to Employee Mapping";
}

function backToRoles() {
  document.getElementById("mappingPanel").style.display = "none";
  document.getElementById("rolePanel").style.display = "block";
  setStep(1);
}

function backToMapping() {
  document.getElementById("results").style.display = "none";
  document.getElementById("mappingPanel").style.display = "block";
  setStep(2);
}

// Collect current mapping UI state into mappingStore before audit
function collectMappings() {
  document.querySelectorAll(".map-row").forEach((row) => {
    const tsKey = row.dataset.tskey;
    const rowId = row.dataset.rowid;
    const payrollKey = document.getElementById(`mpay_${rowId}`).value;
    const dept = document.getElementById(`mdept_${rowId}`).value;
    mappingStore.set(tsKey, { payrollKey, dept });
  });
}

// ═══════════════════════════════════════════════════════════════
// TIMESHEET PARSER — dynamic column detection
// ═══════════════════════════════════════════════════════════════
function parseTimesheet(raw, label) {
  const employees = {};
  let weekStart = null,
    weekEnd = null,
    company = null;

  // Scan header rows for dates / company
  for (let r = 0; r < Math.min(15, raw.length); r++) {
    const row = raw[r] || [];
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || "").trim();
      if (v.includes("Week starting")) {
        for (let dc = 1; dc <= 4; dc++) {
          const dv = String(row[c + dc] || "").trim();
          if (dv && dv !== "nan" && dv.length >= 6) {
            weekStart = dv.substring(0, 10);
            break;
          }
        }
      }
      if (v.includes("Week ending")) {
        for (let dc = 1; dc <= 4; dc++) {
          const dv = String(row[c + dc] || "").trim();
          if (dv && dv !== "nan" && dv.length >= 6) {
            weekEnd = dv.substring(0, 10);
            break;
          }
        }
      }
      if (v.includes("COMPANY NAME"))
        company = v.replace(/COMPANY NAME[:]/, "").trim();
    }
  }

  // Dynamically find nameCol, timeInKeyCol, totalHrsCol
  let nameCol = 0,
    timeInKeyCol = 3,
    totalHrsCol = 25;
  for (let r = 0; r < Math.min(20, raw.length); r++) {
    const row = raw[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c] || "").trim() === "ASSOCIATE NAME:") {
        nameCol = c;
        for (let dr = 1; dr <= 8; dr++) {
          const nrow = raw[r + dr] || [];
          for (let nc = 0; nc < nrow.length; nc++) {
            if (String(nrow[nc] || "").includes("TOTAL HRS TO PAY")) {
              totalHrsCol = nc;
              for (let tc2 = nameCol + 1; tc2 < nrow.length; tc2++) {
                if (String(nrow[tc2] || "").trim() === "Time In") {
                  timeInKeyCol = tc2;
                  break;
                }
              }
              break;
            }
          }
          if (totalHrsCol !== 25 || r > 5) break;
        }
        log(
          "INFO",
          `[${label}] Layout: nameCol=${nameCol} timeInKeyCol=${timeInKeyCol} totalHrsCol=${totalHrsCol}`,
        );
      }
    }
  }

  // Extract employees
  let currentName = null;
  const SKIP = new Set([
    "ASSOCIATE NAME:",
    "MSH Hospitality",
    "",
    "nan",
    "NaN",
    "NaT",
    "0",
  ]);
  for (let r = 0; r < raw.length; r++) {
    const row = raw[r] || [];
    const c0 = String(row[nameCol] || "").trim();
    const isTiming = ["Time In", "Time Out", "BREAK", "Hours"].some((k) =>
      c0.includes(k),
    );
    const isHeader =
      c0.startsWith("COMPANY") ||
      c0.startsWith("Manager") ||
      c0.startsWith("Week") ||
      c0.includes("TOTAL");
    if (
      c0 &&
      !SKIP.has(c0) &&
      !isTiming &&
      !isHeader &&
      isNaN(Number(c0)) &&
      c0.length > 1
    ) {
      currentName = c0;
    }
    const timeKeyVal = String(row[timeInKeyCol] || "").trim();
    if (timeKeyVal === "Time Out" && currentName) {
      let total = toNum(row[totalHrsCol]);
      if (total === 0) total = toNum(row[totalHrsCol + 1]);
      if (total === 0) {
        for (let c = row.length - 1; c >= row.length - 6; c--) {
          const v = toNum(row[c]);
          if (v > 0 && v !== 0.5) {
            total = v;
            break;
          }
        }
      }
      if (total > 0) {
        const key = norm(currentName);
        const prev = employees[key] || { name: currentName, hours: 0 };
        prev.hours += total;
        employees[key] = prev;
        log(
          "OK",
          `[${label}] ${currentName}: +${total.toFixed(2)} hrs → total=${prev.hours.toFixed(2)}`,
        );
      }
    }
  }
  log(
    "INFO",
    `[${label}] Period: ${weekStart || "?"} → ${weekEnd || "?"} | Found: ${Object.keys(employees).length} employees`,
  );
  return { employees, weekStart, weekEnd, company };
}

// ═══════════════════════════════════════════════════════════════
// PAYROLL SUMMARY PARSER
// ═══════════════════════════════════════════════════════════════
function parsePayroll(raw, label) {
  const employees = {},
    depts = {};
  let currentName = null,
    currentDept = null,
    period = null;

  for (const row of raw) {
    if (!row) continue;
    const v = String(row[4] || "").trim();
    if (v.match(/\d{2}\.\d{2}.{0,5}\d{2}\.\d{2}/)) {
      period = v;
      break;
    }
  }

  const DEPT_HDR =
    /^(HOUSEKEEPING|HOUSEKEEPING SCOTTSDALE|ADMINISTRADOR|PROYEC|CHEF|MAINTENANCE|POOL|COOK)$/i;

  for (let r = 0; r < raw.length; r++) {
    const row = raw[r] || [];
    const col1 = String(row[1] || "").trim();
    if (!col1 || col1 === "nan" || col1 === "NaT") continue;

    const dataInRow = [4, 5, 6, 7, 8, 9, 10, 11, 12].some(
      (c) =>
        toNum(row[c]) > 0 ||
        (String(row[c] || "").trim() && String(row[c] || "").trim() !== "nan"),
    );
    if (DEPT_HDR.test(col1) && !dataInRow) {
      currentDept = col1.trim().toUpperCase();
      log("INFO", `[${label}] Dept section: ${currentDept}`);
      continue;
    }

    const structural = [
      "MSH Hospitality",
      "ASSOCIATE NAME:",
      "Regular Hours",
      "Over Time",
      "0GG",
      "nan",
      "NaN",
      "",
    ];
    if (structural.some((s) => col1.toLowerCase() === s.toLowerCase()))
      continue;
    if (col1.match(/^\d{2}\.\d{2}/)) continue;

    const nextRow = raw[r + 1] || [];
    const hasNextHours =
      toNum(nextRow[4]) > 0 ||
      toNum(nextRow[5]) > 0 ||
      toNum(nextRow[6]) > 0 ||
      toNum(nextRow[7]) > 0;
    const hasCurrentHours = toNum(row[4]) > 0 || toNum(row[6]) > 0;
    const hasFileNum = String(row[8] || "")
      .trim()
      .match(/^\d+$/);

    if (col1 && isNaN(Number(col1)) && !DEPT_HDR.test(col1)) currentName = col1;

    if (currentName && (hasCurrentHours || hasNextHours || hasFileNum)) {
      const dataRow = hasCurrentHours ? row : hasNextHours ? nextRow : row;
      const w1reg = toNum(dataRow[4]),
        w1ot = toNum(dataRow[5]);
      const w2reg = toNum(dataRow[6]),
        w2ot = toNum(dataRow[7]);
      const fileNum = String(dataRow[8] || "").trim();
      const adpReg = toNum(dataRow[11]),
        adpOT = toNum(dataRow[12]);
      const totalReg = w1reg + w2reg,
        totalOT = w1ot + w2ot;
      if (totalReg > 0 || totalOT > 0 || fileNum.match(/^\d{3,}/)) {
        const key = norm(currentName);
        employees[key] = {
          name: currentName,
          dept: currentDept || "UNKNOWN",
          w1reg,
          w1ot,
          w2reg,
          w2ot,
          totalReg,
          totalOT,
          totalHours: totalReg + totalOT,
          fileNum,
          adpReg,
          adpOT,
        };
        depts[currentDept || "UNKNOWN"] =
          (depts[currentDept || "UNKNOWN"] || 0) + totalReg + totalOT;
        log(
          "OK",
          `[${label}] ${currentName} [${currentDept}] W1=${w1reg}+${w1ot}OT W2=${w2reg}+${w2ot}OT File#${fileNum}`,
        );
      }
    }
  }
  log(
    "INFO",
    `[${label}] Period: ${period} | Employees: ${Object.keys(employees).length}`,
  );
  return { employees, depts, period };
}

// ═══════════════════════════════════════════════════════════════
// MSH INVOICE PARSER
// ═══════════════════════════════════════════════════════════════
function parseInvoiceMSH(raw, label) {
  const lines = [];
  let invoiceNum = null,
    period = null;
  for (const row of raw) {
    if (!row) continue;
    if (String(row[4] || "").trim() === "INVOICE #")
      invoiceNum = String(row[5] || "").trim();
    const c0 = String(row[0] || "").trim();
    if (c0.match(/^\d{2}\.\d{2}/)) period = c0;
    const hours = toNum(row[4]),
      rate = toNum(row[3]),
      amount = toNum(row[5]);
    const descKey = norm(c0);
    if (c0 && (hours > 0 || rate > 0) && INV_DESC_MAP[descKey] !== undefined) {
      lines.push({
        description: c0,
        descKey,
        canonKey: INV_DESC_MAP[descKey],
        rate,
        hours,
        amount,
      });
      log("OK", `[${label}] Line: "${c0}" ${hours}h @ $${rate}`);
    }
  }
  log(
    "INFO",
    `[${label}] Invoice #${invoiceNum} Period: ${period} Lines: ${lines.length}`,
  );
  return { lines, invoiceNum, period };
}

// ═══════════════════════════════════════════════════════════════
// YELLOWSTONE INVOICE PARSER
// ═══════════════════════════════════════════════════════════════
function parseInvoiceYS(raw, label) {
  let amount = 0,
    invoiceNum = null,
    period = null;
  for (const row of raw) {
    if (!row) continue;
    if (String(row[2] || "").trim() === "INVOICE #")
      invoiceNum = String(row[3] || "").trim();
    const c2 = String(row[2] || "").trim(),
      c3 = String(row[3] || "");
    if (c2 === "DATE" || c2.includes("DATE")) period = c3.trim();
    if (c2.includes("TOTAL")) {
      const v = toNum(row[3]);
      if (v > 0) amount = v;
    }
  }
  if (amount === 0) {
    for (const row of raw) {
      const v = toNum(row?.[3] || 0);
      if (v > amount) amount = v;
    }
  }
  log(
    "INFO",
    `[${label}] YS Invoice #${invoiceNum} Period: ${period} Amount: $${amount}`,
  );
  return { amount, invoiceNum, period };
}

// ═══════════════════════════════════════════════════════════════
// ADP PARSER
// ═══════════════════════════════════════════════════════════════
function parseADP(raw, label) {
  const records = {};
  if (!raw.length) return records;
  let hdr = 0;
  for (let r = 0; r < Math.min(5, raw.length); r++) {
    if (raw[r] && String(raw[r][0] || "").includes("Co Code")) {
      hdr = r;
      break;
    }
  }
  const headers = (raw[hdr] || []).map((h) => String(h || "").trim());
  const fi2 = headers.findIndex((h) => h === "File #");
  const ri = headers.findIndex((h) => h === "Reg Hours");
  const oi = headers.findIndex((h) => h === "O/T Hours");
  for (let r = hdr + 1; r < raw.length; r++) {
    const row = raw[r] || [];
    const fn = String(row[fi2] || "").trim();
    const reg = toNum(row[ri]),
      ot = toNum(row[oi]);
    if (fn && (reg || ot)) {
      records[fn] = { fileNum: fn, reg, ot, total: reg + ot };
      log("OK", `[${label}] File #${fn}: Reg=${reg} OT=${ot}`);
    }
  }
  log("INFO", `[${label}] ADP records: ${Object.keys(records).length}`);
  return records;
}

// ═══════════════════════════════════════════════════════════════
// MAIN AUDIT  (Step 3)
// ═══════════════════════════════════════════════════════════════
async function runAudit() {
  // Snapshot current mapping UI state
  collectMappings();

  logs = [];
  parsedData = {};
  audit = {};
  const btn = document.getElementById("auditBtn");
  btn.disabled = true;
  btn.textContent = "⏳ Running audit…";

  try {
    const rm = getRoleMap();

    // ── Parse timesheets ──
    const combineTS = async (files, weekLabel) => {
      const combined = {};
      let start = null,
        end = null,
        company = null;
      for (const f of files || []) {
        const r = parseTimesheet(await toRaw(f), `${weekLabel}:${f.name}`);
        Object.entries(r.employees).forEach(([k, v]) => {
          combined[k] = combined[k]
            ? { ...v, hours: combined[k].hours + v.hours }
            : { ...v };
        });
        if (!start) {
          start = r.weekStart;
          end = r.weekEnd;
          company = r.company;
        }
      }
      return { emps: combined, start, end, company };
    };
    const w1 = await combineTS(rm["week1"], "W1");
    const w2 = await combineTS(rm["week2"], "W2");
    parsedData.w1 = w1;
    parsedData.w2 = w2;

    // ── Parse payroll ──
    let payrollData = { employees: {}, depts: {}, period: null };
    for (const f of rm["payroll"] || []) {
      payrollData = parsePayroll(await toRaw(f), `Payroll:${f.name}`);
    }
    parsedData.payroll = payrollData;

    // ── Build biweekly using mapping ──
    //  mappingStore: tsNormKey → {payrollKey, dept}
    //  The "biweekly" key is the payrollKey (if mapped) or the tsNormKey (unmapped)
    const allTSKeys = new Set([
      ...Object.keys(w1.emps),
      ...Object.keys(w2.emps),
    ]);
    const biweekly = {};

    allTSKeys.forEach((tsKey) => {
      const e1 = w1.emps[tsKey] || { name: tsKey, hours: 0 };
      const e2 = w2.emps[tsKey] || { name: tsKey, hours: 0 };
      const map = mappingStore.get(tsKey) || {};
      const resolvedKey =
        map.payrollKey && map.payrollKey !== "" ? map.payrollKey : tsKey;
      const dept = (
        map.dept ||
        payrollData.employees[resolvedKey]?.dept ||
        "UNKNOWN"
      ).toUpperCase();
      const deptNorm = norm(dept);
      const isYS =
        YELLOWSTONE_DEPTS.has(deptNorm) ||
        ["proyec", "rishabh"].some((y) => tsKey.includes(y));
      const displayName = capName(
        payrollData.employees[resolvedKey]?.name || e1.name || e2.name,
      );

      if (biweekly[resolvedKey]) {
        biweekly[resolvedKey].w1 += e1.hours;
        biweekly[resolvedKey].w2 += e2.hours;
        biweekly[resolvedKey].total += e1.hours + e2.hours;
      } else {
        biweekly[resolvedKey] = {
          name: displayName,
          dept,
          isYS,
          w1: e1.hours,
          w2: e2.hours,
          total: e1.hours + e2.hours,
        };
      }
    });
    parsedData.biweekly = biweekly;

    // ── Parse invoices ──
    const invoiceLines = [];
    for (const f of rm["invoice_msh"] || []) {
      invoiceLines.push(parseInvoiceMSH(await toRaw(f), `MSHInv:${f.name}`));
    }
    parsedData.invoices = invoiceLines;
    const invByCanon = {};
    invoiceLines.forEach((inv) => {
      inv.lines.forEach((l) => {
        if (!invByCanon[l.canonKey])
          invByCanon[l.canonKey] = {
            desc: l.description,
            hours: 0,
            amount: 0,
            rate: l.rate,
          };
        invByCanon[l.canonKey].hours += l.hours;
        invByCanon[l.canonKey].amount += l.amount;
      });
    });
    parsedData.invByCanon = invByCanon;

    const ysInvoices = [];
    for (const f of rm["invoice_ys"] || []) {
      ysInvoices.push(parseInvoiceYS(await toRaw(f), `YSInv:${f.name}`));
    }
    parsedData.ysInvoices = ysInvoices;

    let adpRecords = {};
    for (const f of rm["adp"] || []) {
      adpRecords = parseADP(await toRaw(f), `ADP:${f.name}`);
    }
    parsedData.adp = adpRecords;

    // ── Audit 1: Employee overview ──
    audit.employees = Object.values(biweekly)
      .filter((e) => e.total > 0)
      .sort((a, b) => b.total - a.total);

    // ── Audit 2: Dept vs Invoice ──
    const deptHoursTS = {};
    Object.values(biweekly).forEach((e) => {
      if (e.isYS) return;
      const canon = DEPT_NORM[norm(e.dept)] || norm(e.dept);
      deptHoursTS[canon] = (deptHoursTS[canon] || 0) + e.total;
    });
    const allDeptKeys = new Set([
      ...Object.keys(deptHoursTS),
      ...Object.keys(invByCanon),
    ]);
    audit.deptComparison = [];
    allDeptKeys.forEach((k) => {
      if (k === "yellowstone" || k === "administrador") return;
      const tsH = deptHoursTS[k] || null,
        invH = invByCanon[k] ? invByCanon[k].hours : null;
      const diff = tsH !== null && invH !== null ? tsH - invH : null;
      audit.deptComparison.push({
        key: k,
        desc: invByCanon[k] ? invByCanon[k].desc : k,
        tsHours: tsH,
        invHours: invH,
        diff,
        rate: invByCanon[k]?.rate || null,
        amount: invByCanon[k]?.amount || null,
        mismatch: diff !== null && Math.abs(diff) >= 0.05,
      });
    });
    audit.deptComparison.sort(
      (a, b) => Math.abs(b.diff || 0) - Math.abs(a.diff || 0),
    );

    // ── Audit 3: TS vs Payroll per employee (100% match) ──
    const allEmpKeys = new Set([
      ...Object.keys(biweekly),
      ...Object.keys(payrollData.employees),
    ]);
    audit.payComparisons = [];
    allEmpKeys.forEach((k) => {
      const ts = biweekly[k],
        pay = payrollData.employees[k];
      const tsTotal = ts ? ts.total : null,
        payTotal = pay ? pay.totalHours : null;
      const diff =
        tsTotal !== null && payTotal !== null ? tsTotal - payTotal : null;
      const isYS = ts
        ? ts.isYS
        : (pay && YELLOWSTONE_DEPTS.has(norm(pay.dept || ""))) ||
          k.includes("rishabh");
      const name = ts ? ts.name : pay ? capName(pay.name) : k;
      audit.payComparisons.push({
        name,
        key: k,
        dept: ts ? ts.dept : pay ? pay.dept : "UNKNOWN",
        isYS,
        tsW1: ts ? ts.w1 : null,
        tsW2: ts ? ts.w2 : null,
        tsTotal,
        payW1: pay ? pay.w1reg + pay.w1ot : null,
        payW2: pay ? pay.w2reg + pay.w2ot : null,
        payTotal,
        diff,
        mismatch: diff !== null && Math.abs(diff) >= 0.05,
        missingTS: tsTotal === null,
        missingPay: payTotal === null,
        fileNum: pay ? pay.fileNum : null,
      });
    });
    audit.payComparisons.sort((a, b) => {
      if (a.mismatch !== b.mismatch) return a.mismatch ? -1 : 1;
      return Math.abs(b.diff || 0) - Math.abs(a.diff || 0);
    });

    // ── Audit 4: Payroll vs ADP (100% match) ──
    audit.adpComparisons = [];
    Object.values(payrollData.employees).forEach((p) => {
      const adp = adpRecords[p.fileNum];
      const regDiff = adp ? p.totalReg - adp.reg : null,
        otDiff = adp ? p.totalOT - adp.ot : null;
      const isYS = YELLOWSTONE_DEPTS.has(
        norm(p.dept || "") || norm(p.name).includes("rishabh"),
      );
      audit.adpComparisons.push({
        name: capName(p.name),
        fileNum: p.fileNum,
        dept: p.dept,
        isYS,
        payReg: p.totalReg,
        payOT: p.totalOT,
        adpReg: adp ? adp.reg : null,
        adpOT: adp ? adp.ot : null,
        regDiff,
        otDiff,
        mismatch:
          (regDiff !== null && Math.abs(regDiff) >= 0.05) ||
          (otDiff !== null && Math.abs(otDiff) >= 0.05),
        missingADP: !adp && !!p.fileNum,
      });
    });
    Object.values(adpRecords).forEach((a) => {
      if (
        !Object.values(payrollData.employees).find(
          (p) => p.fileNum === a.fileNum,
        )
      ) {
        audit.adpComparisons.push({
          name: "—",
          fileNum: a.fileNum,
          dept: "UNKNOWN",
          isYS: false,
          payReg: null,
          payOT: null,
          adpReg: a.reg,
          adpOT: a.ot,
          regDiff: null,
          otDiff: null,
          mismatch: true,
          missingPayroll: true,
        });
      }
    });
    audit.adpComparisons.sort((a, b) => {
      if (a.mismatch !== b.mismatch) return a.mismatch ? -1 : 1;
      return Math.abs(b.regDiff || 0) - Math.abs(a.regDiff || 0);
    });

    renderResults();
    document.getElementById("mappingPanel").style.display = "none";
    document.getElementById("results").style.display = "block";
    setStep(3);
    document.getElementById("results").scrollIntoView({ behavior: "smooth" });
  } catch (err) {
    log("ERR", "Fatal: " + err.message);
    console.error(err);
    alert(
      "Processing error: " + err.message + "\n\nSee Parse Log for details.",
    );
  }
  btn.disabled = false;
  btn.textContent = "▶ Run Payroll Audit";
}

// ═══════════════════════════════════════════════════════════════
// RENDER
// ═══════════════════════════════════════════════════════════════
function diffBadge(d) {
  if (d === null) return '<span class="badge b-miss">N/A</span>';
  if (Math.abs(d) < 0.05) return '<span class="badge b-ok">✓ Match</span>';
  const s = (d > 0 ? "+" : "") + d.toFixed(2);
  if (Math.abs(d) < 1) return `<span class="badge b-warn">⚠ ${s} hrs</span>`;
  return `<span class="badge b-err">✗ ${s} hrs</span>`;
}

function renderResults() {
  const p = parsedData,
    a = audit;

  // Period bar
  document.getElementById("pbar").innerHTML = `
    <div>📅 W1: <span>${p.w1.start || "?"}</span> → <span>${p.w1.end || "?"}</span></div>
    <div>📅 W2: <span>${p.w2.start || "?"}</span> → <span>${p.w2.end || "?"}</span></div>
    <div>🏢 <span>${p.w1.company || p.w2.company || "The Buttes"}</span></div>
    <div>👥 Employees: <span>${a.employees.length}</span></div>
    <div>⏱ Total TS hrs: <span>${a.employees.reduce((s, e) => s + e.total, 0).toFixed(2)}</span></div>
    <div>🔵 Yellowstone employees: <span>${a.employees.filter((e) => e.isYS).length}</span></div>`;

  // Summary cards
  const payMM = a.payComparisons.filter((r) => r.mismatch).length;
  const adpMM = a.adpComparisons.filter((r) => r.mismatch).length;
  const deptMM = a.deptComparison.filter((r) => r.mismatch).length;
  const missPay = a.payComparisons.filter(
    (r) => r.missingTS || r.missingPay,
  ).length;
  const total = payMM + adpMM + deptMM;
  document.getElementById("cards").innerHTML = `
    <div class="card ${total === 0 ? "ok" : "err"}"><div class="lbl">Total Issues</div><div class="val">${total}</div><div class="sub">${total === 0 ? "All checks passed ✓" : "Requires review"}</div></div>
    <div class="card ${deptMM === 0 ? "ok" : "warn"}"><div class="lbl">Dept vs Invoice</div><div class="val">${deptMM}</div><div class="sub">dept-level mismatches</div></div>
    <div class="card ${payMM === 0 ? "ok" : "err"}"><div class="lbl">TS vs Payroll</div><div class="val">${payMM}</div><div class="sub">employee-level mismatches</div></div>
    <div class="card ${adpMM === 0 ? "ok" : "err"}"><div class="lbl">Payroll vs ADP</div><div class="val">${adpMM}</div><div class="sub">employee-level mismatches</div></div>
    <div class="card ${missPay === 0 ? "ok" : "warn"}"><div class="lbl">Missing Records</div><div class="val">${missPay}</div><div class="sub">employees in one source only</div></div>`;

  // Alert banners
  const banners = [];
  const ysEmps = a.employees.filter((e) => e.isYS);
  if (ysEmps.length > 0) {
    const ysInvTotal = p.ysInvoices.reduce((s, i) => s + i.amount, 0);
    banners.push(`<div class="alert alert-ys">🔵 <div><strong>Yellowstone Activity LLC (${ysEmps.length} employee${ysEmps.length > 1 ? "s" : ""}):</strong>
      ${ysEmps.map((e) => `<strong>${e.name}</strong>`).join(", ")} — On hotel timesheets but billed separately via Yellowstone invoices
      totalling <strong>$${ysInvTotal.toFixed(2)}</strong> (${p.ysInvoices.map((i) => `Invoice #${i.invoiceNum}`).join(", ")}).
      Excluded from dept-vs-invoice comparison. Must still match 100% in TS↔Payroll and Payroll↔ADP.</div></div>`);
  }
  if (payMM > 0) {
    const names = a.payComparisons
      .filter((r) => r.mismatch)
      .map(
        (r) =>
          `<strong>${r.name}</strong> (Δ ${r.diff != null ? (r.diff > 0 ? "+" : "") + r.diff.toFixed(2) : "?"} hrs)`,
      )
      .join(", ");
    banners.push(
      `<div class="alert alert-err">✗ <div><strong>Timesheet vs Payroll Mismatches (${payMM}):</strong> ${names}</div></div>`,
    );
  }
  if (adpMM > 0) {
    const names = a.adpComparisons
      .filter((r) => r.mismatch)
      .map((r) => `<strong>${r.name}</strong> File#${r.fileNum}`)
      .join(", ");
    banners.push(
      `<div class="alert alert-err">✗ <div><strong>Payroll vs ADP Mismatches (${adpMM}):</strong> ${names}</div></div>`,
    );
  }
  if (deptMM > 0) {
    const names = a.deptComparison
      .filter((r) => r.mismatch)
      .map(
        (r) =>
          `<strong>${r.desc}</strong> (Δ ${r.diff != null ? (r.diff > 0 ? "+" : "") + r.diff.toFixed(2) : "?"} hrs)`,
      )
      .join(", ");
    banners.push(
      `<div class="alert alert-warn">⚠ <div><strong>Dept/Invoice Hour Differences (${deptMM}):</strong> ${names}</div></div>`,
    );
  }
  document.getElementById("alertBanners").innerHTML = banners.join("");

  // Tab badges
  document.getElementById("tb0").textContent = a.employees.length;
  document.getElementById("tb1").textContent = deptMM > 0 ? `${deptMM} ⚠` : "✓";
  document.getElementById("tb2").textContent = payMM > 0 ? `${payMM} ✗` : "✓";
  document.getElementById("tb3").textContent = adpMM > 0 ? `${adpMM} ✗` : "✓";

  // Tab 0: Employees
  document.getElementById("empTable").innerHTML = `
    <thead><tr><th>#</th><th>Employee Name</th><th>Department</th><th>Entity</th><th>Week 1 Hrs</th><th>Week 2 Hrs</th><th>Biweekly Total</th></tr></thead>
    <tbody>${
      a.employees
        .map(
          (e, i) => `<tr class="${e.isYS ? "tr-ys" : ""}">
      <td style="color:var(--muted)">${i + 1}</td>
      <td class="emp-name">${e.name}</td>
      <td><span class="badge b-dept">${e.dept}</span></td>
      <td>${e.isYS ? '<span class="badge b-ys">🔵 Yellowstone LLC</span>' : '<span class="badge b-info">MSH Hospitality</span>'}</td>
      <td class="num">${fmtH(e.w1)}</td>
      <td class="num">${fmtH(e.w2)}</td>
      <td class="num"><strong>${fmtH(e.total)}</strong></td>
    </tr>`,
        )
        .join("") ||
      '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:28px">No timesheet data parsed</td></tr>'
    }</tbody>`;

  // Tab 1: Dept vs Invoice
  const ysInvTotal2 = p.ysInvoices.reduce((s, i) => s + i.amount, 0);
  document.getElementById("tc1-inner").innerHTML = `
    ${
      ysEmps.length > 0
        ? `<div class="alert alert-ys" style="margin-bottom:14px">🔵 <div>
      Yellowstone LLC employees (<strong>${ysEmps.map((e) => e.name).join(", ")}</strong>) are excluded.
      Billed separately: <strong>$${ysInvTotal2.toFixed(2)}</strong> (${p.ysInvoices.map((i) => `Invoice #${i.invoiceNum}`).join(", ")}).
    </div></div>`
        : ""
    }
    <div class="twrap"><table>
      <thead><tr><th>Dept / Role</th><th>TS Hours (biweekly)</th><th>Invoice Hours (biweekly)</th><th>Δ Diff</th><th>Rate/hr</th><th>Invoiced Amount</th><th>Status</th></tr></thead>
      <tbody>${
        a.deptComparison
          .map(
            (r) => `<tr class="${r.mismatch ? "tr-mismatch" : ""}">
        <td><strong>${r.desc}</strong></td>
        <td class="num">${fmtH(r.tsHours)}</td>
        <td class="num">${fmtH(r.invHours)}</td>
        <td class="num ${diffCls(r.diff)}">${r.diff !== null ? (r.diff > 0 ? "+" : "") + r.diff.toFixed(2) : "—"}</td>
        <td class="num">${r.rate ? "$" + r.rate.toFixed(2) : "—"}</td>
        <td class="num">${r.amount ? "$" + r.amount.toFixed(2) : "—"}</td>
        <td>${diffBadge(r.diff)}</td>
      </tr>`,
          )
          .join("") ||
        '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:28px">No invoice data — upload MSH Invoice files</td></tr>'
      }</tbody>
    </table></div>
    <div style="margin-top:14px;font-size:12px;color:var(--muted)">
      <strong style="color:var(--text)">Note:</strong> Dept totals aggregate all non-Yellowstone employees by department.
      "Housekeeper" on the invoice = all employees in the <em>HOUSEKEEPING</em> department on timesheets.
    </div>`;

  // Tab 2: TS vs Payroll
  document.getElementById("payTable").innerHTML = `
    <thead><tr><th>Employee</th><th>Dept</th><th>Entity</th><th>TS W1</th><th>TS W2</th><th>TS Total</th><th>Pay W1</th><th>Pay W2</th><th>Pay Total</th><th>Δ Diff</th><th>Status</th></tr></thead>
    <tbody>${
      a.payComparisons
        .map(
          (
            r,
          ) => `<tr class="${r.mismatch ? "tr-mismatch" : r.isYS ? "tr-ys" : ""}">
      <td class="emp-name">${r.name}</td>
      <td><span class="badge b-dept">${r.dept}</span></td>
      <td>${r.isYS ? '<span class="badge b-ys">🔵 YS LLC</span>' : '<span class="badge b-info">MSH</span>'}</td>
      <td class="num">${r.tsW1 !== null ? fmtH(r.tsW1) : '<span style="color:var(--muted)">—</span>'}</td>
      <td class="num">${r.tsW2 !== null ? fmtH(r.tsW2) : '<span style="color:var(--muted)">—</span>'}</td>
      <td class="num"><strong>${r.tsTotal !== null ? fmtH(r.tsTotal) : '<span class="badge b-miss">Missing</span>'}</strong></td>
      <td class="num">${r.payW1 !== null ? fmtH(r.payW1) : '<span style="color:var(--muted)">—</span>'}</td>
      <td class="num">${r.payW2 !== null ? fmtH(r.payW2) : '<span style="color:var(--muted)">—</span>'}</td>
      <td class="num"><strong>${r.payTotal !== null ? fmtH(r.payTotal) : '<span class="badge b-miss">Missing</span>'}</strong></td>
      <td class="num ${diffCls(r.diff)}">${r.diff !== null ? (r.diff > 0 ? "+" : "") + r.diff.toFixed(2) : "—"}</td>
      <td>${r.missingTS ? '<span class="badge b-warn">⚠ Not in TS</span>' : r.missingPay ? '<span class="badge b-warn">⚠ Not in Payroll</span>' : r.mismatch ? `<span class="badge b-err">✗ MISMATCH ${(r.diff > 0 ? "+" : "") + r.diff.toFixed(2)} hrs</span>` : '<span class="badge b-ok">✓ Match</span>'}</td>
    </tr>`,
        )
        .join("") ||
      '<tr><td colspan="11" style="text-align:center;color:var(--muted);padding:28px">No payroll data — upload Internal Payroll Summary</td></tr>'
    }</tbody>`;

  // Tab 3: Payroll vs ADP
  document.getElementById("adpTable").innerHTML = `
    <thead><tr><th>Employee</th><th>Dept</th><th>ADP File #</th><th>Entity</th><th>Pay Reg</th><th>Pay OT</th><th>ADP Reg</th><th>ADP OT</th><th>Reg Δ</th><th>OT Δ</th><th>Status</th></tr></thead>
    <tbody>${
      a.adpComparisons
        .map(
          (
            r,
          ) => `<tr class="${r.mismatch ? "tr-mismatch" : r.isYS ? "tr-ys" : ""}">
      <td class="emp-name">${r.name}</td>
      <td><span class="badge b-dept">${r.dept || "?"}</span></td>
      <td><span class="badge b-info">${r.fileNum || "—"}</span></td>
      <td>${r.isYS ? '<span class="badge b-ys">🔵 YS LLC</span>' : '<span class="badge b-info">MSH</span>'}</td>
      <td class="num">${fmtH(r.payReg)}</td>
      <td class="num">${fmtH(r.payOT)}</td>
      <td class="num">${r.adpReg !== null ? fmtH(r.adpReg) : '<span class="badge b-err">Not Found</span>'}</td>
      <td class="num">${r.adpOT !== null ? fmtH(r.adpOT) : "—"}</td>
      <td class="num ${diffCls(r.regDiff)}">${r.regDiff !== null ? (r.regDiff > 0 ? "+" : "") + r.regDiff.toFixed(2) : "—"}</td>
      <td class="num ${diffCls(r.otDiff)}">${r.otDiff !== null ? (r.otDiff > 0 ? "+" : "") + r.otDiff.toFixed(2) : "—"}</td>
      <td>${r.missingPayroll ? '<span class="badge b-warn">⚠ Not in Payroll</span>' : r.missingADP ? '<span class="badge b-err">✗ Missing in ADP</span>' : r.mismatch ? '<span class="badge b-err">✗ MISMATCH</span>' : '<span class="badge b-ok">✓ Match</span>'}</td>
    </tr>`,
        )
        .join("") ||
      '<tr><td colspan="11" style="text-align:center;color:var(--muted);padding:28px">No data — upload Payroll Summary and ADP Export</td></tr>'
    }</tbody>`;

  // Tab 4: Log
  document.getElementById("logList").innerHTML = logs
    .map(
      (e) => `
    <li><span class="loglvl ${e.lvl}">${e.lvl}</span><span>${e.msg}</span></li>`,
    )
    .join("");
}

// ═══════════════════════════════════════════════════════════════
// TABS
// ═══════════════════════════════════════════════════════════════
function switchTab(idx) {
  document
    .querySelectorAll(".tcontent")
    .forEach((t, i) => t.classList.toggle("active", i === idx));
  document
    .querySelectorAll(".tab")
    .forEach((t, i) => t.classList.toggle("active", i === idx));
}

// ═══════════════════════════════════════════════════════════════
// EXPORT CSV
// ═══════════════════════════════════════════════════════════════
function exportCSV() {
  const rows = [
    [
      "Employee",
      "Dept",
      "Entity",
      "TS W1",
      "TS W2",
      "TS Total",
      "Pay Total",
      "TS-Pay Diff",
      "ADP File#",
      "ADP Reg",
      "Pay Reg",
      "Reg Diff",
      "ADP OT",
      "Pay OT",
      "OT Diff",
      "Overall Status",
    ],
  ];
  const payMap = {};
  audit.payComparisons.forEach((r) => {
    payMap[r.key] = r;
  });
  const adpMap = {};
  audit.adpComparisons.forEach((r) => {
    adpMap[r.fileNum] = r;
  });
  audit.employees.forEach((e) => {
    const pay = payMap[e.key] || payMap[norm(e.name)] || {};
    const adp = adpMap[pay.fileNum] || {};
    rows.push([
      e.name,
      e.dept,
      e.isYS ? "Yellowstone LLC" : "MSH Hospitality",
      fmtH(e.w1),
      fmtH(e.w2),
      fmtH(e.total),
      fmtH(pay.payTotal),
      pay.diff != null ? (pay.diff > 0 ? "+" : "") + pay.diff.toFixed(2) : "—",
      pay.fileNum || "—",
      fmtH(adp.adpReg),
      fmtH(adp.payReg),
      adp.regDiff != null
        ? (adp.regDiff > 0 ? "+" : "") + adp.regDiff.toFixed(2)
        : "—",
      fmtH(adp.adpOT),
      fmtH(adp.payOT),
      adp.otDiff != null
        ? (adp.otDiff > 0 ? "+" : "") + adp.otDiff.toFixed(2)
        : "—",
      pay.mismatch || adp.mismatch ? "MISMATCH" : "OK",
    ]);
  });
  const csv = rows.map((r) => r.map((v) => `"${v}"`).join(",")).join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = "payroll_audit_report_v3.csv";
  a.click();
}

const actions = {
  showMapping,
  resetAll,
  runAudit,
  backToRoles,
  resetAll,
  exportCSV,
  switchTab,
  switchTab,
  switchTab,
  switchTab,
  switchTab,
};

Object.assign(window, actions);
