/* ecoism-qc static form-filling app
   Runs entirely client-side. No backend, no data leaves the browser
   (except the template files fetched from this same site).
*/

const DATA_DIR = "data/";
const XLSX_TEMPLATE_DIR = "templates_xlsx/";
const DOCX_TEMPLATE_DIR = "templates_docx/";

let fieldMap = {};
let docxForms = {};
let locations = [];
let personnel = {};
let materials = [];
let pipeNames = [];

const XLSX_FILES = {
  "Final1_ACC": { file: "20260729-Final1_ACC-Template.xlsx", label: "Acceptance Minutes of Material" },
  "Final11_Rec": { file: "20260729-Final11_Rec_Template.xlsx", label: "Pipe Jacking Record (PJ-03)" },
  "Final12-13_Ins": { file: "Final12-13_Ins_Template.xlsx", label: "Internal Ins / Inspection (Civil Work)" },
  "Final5-10_ITP": { file: "20260729-Final5-10_ITP_Template.xlsx", label: "ITP Checksheet PJ-01~07 + 02a/02b" },
};

const SHEET_LABELS = {
  "lần 2": "Acceptance Minutes",
  "PJ-03": "PJ-03 - Pipe Jacking Record",
  "Internal Ins": "Internal Inspection Minutes",
  "Inspection": "Inspection Sheet of Civil Work",
  "PJ-01": "PJ-01 - Input Machinery",
  "PJ-02": "PJ-02 - Preparation Work",
  "PJ-02a": "PJ-02a - Temporary Facilities",
  "PJ-02b": "PJ-02b",
  "PJ-03a": "PJ-03a - Cushion Ring Width",
  "PJ-04": "PJ-04 - Coordinate & Elevation",
  "PJ-05": "PJ-05 - Grout Injection",
  "PJ-06": "PJ-06 - Leakage Test",
  "PJ-07": "PJ-07 - Completed Jacking Pipes",
};

async function loadJSON(path) {
  const r = await fetch(path);
  return r.json();
}

async function init() {
  [fieldMap, docxForms, locations, personnel, materials, pipeNames] = await Promise.all([
    loadJSON(DATA_DIR + "field_map.json"),
    loadJSON(DATA_DIR + "docx_forms.json"),
    loadJSON(DATA_DIR + "locations.json"),
    loadJSON(DATA_DIR + "personnel.json"),
    loadJSON(DATA_DIR + "materials.json"),
    loadJSON(DATA_DIR + "pipe_names.json"),
  ]);
  renderPicker();
  document.getElementById("backBtn").addEventListener("click", showPicker);
  document.getElementById("exportBtn").addEventListener("click", doExport);
}

let currentSelection = null; // {kind:'xlsx', fileKey, sheet} or {kind:'docx', formKey}

function renderPicker() {
  const container = document.getElementById("formList");
  container.innerHTML = "";

  // xlsx forms (one card per sheet)
  for (const [fileKey, info] of Object.entries(XLSX_FILES)) {
    const sheets = fieldMap[fileKey] || {};
    for (const sheetName of Object.keys(sheets)) {
      const card = document.createElement("div");
      card.className = "form-card";
      card.innerHTML = `
        <div class="code">Excel · ${fileKey}</div>
        <div class="name">${SHEET_LABELS[sheetName] || sheetName}</div>
        <span class="type">${sheets[sheetName].length} trường</span>
      `;
      card.onclick = () => selectForm({ kind: "xlsx", fileKey, sheet: sheetName });
      container.appendChild(card);
    }
  }

  // docx forms
  for (const [formKey, info] of Object.entries(docxForms)) {
    const card = document.createElement("div");
    card.className = "form-card";
    card.innerHTML = `
      <div class="code">Word · ${formKey}</div>
      <div class="name">${info.title}</div>
      <span class="type">${info.fields.length} trường</span>
    `;
    card.onclick = () => selectForm({ kind: "docx", formKey });
    container.appendChild(card);
  }
}

function showPicker() {
  document.getElementById("picker").classList.remove("hidden");
  document.getElementById("formArea").classList.add("hidden");
  currentSelection = null;
}

function selectForm(sel) {
  currentSelection = sel;
  document.getElementById("picker").classList.add("hidden");
  document.getElementById("formArea").classList.remove("hidden");
  const formEl = document.getElementById("dataForm");
  formEl.innerHTML = "";
  document.getElementById("formStatus")?.remove();

  if (sel.kind === "xlsx") {
    const fields = fieldMap[sel.fileKey][sel.sheet];
    document.getElementById("formTitle").textContent =
      `2. Nhập dữ liệu — ${SHEET_LABELS[sel.sheet] || sel.sheet}`;
    renderXlsxFields(formEl, fields, sel.fileKey, sel.sheet);
  } else {
    const info = docxForms[sel.formKey];
    document.getElementById("formTitle").textContent = `2. Nhập dữ liệu — ${info.title}`;
    renderDocxFields(formEl, info.fields);
  }
}

/* ---------- Manual-entry override rules ----------
   Some sheets must ALWAYS be plain manual text input, never auto-detected
   dropdowns, because they hold equipment spec/model/certificate values that
   only coincidentally resemble other patterns (location/pipe/personnel).
   PJ-01's actual data columns are H,L,P,T (4 merged groups H:K/L:O/P:S/T:W)
   plus F/I/M/R/V/X/Z (STT + checkbox columns) - applying to the whole sheet
   is simplest and safe since none of these should ever match a location,
   pipe-name, or personnel dropdown anyway. */
const MANUAL_ENTRY_SHEETS = {
  "Final5-10_ITP": ["PJ-01"],
};

function sheetIsManualEntryOnly(sel) {
  if (!sel) return false;
  const list = MANUAL_ENTRY_SHEETS[sel.fileKey];
  return !!list && list.includes(sel.sheet);
}

/* Cell types that must be forced to TEXT on export, never auto-numeric/date.
   Applies to any field whose value matches these patterns, regardless of sheet -
   protects values like "42.15 ~ 42.18" or equipment codes that look numeric. */
function mustForceText(cellRef, sel, sample) {
  if (isLocationLike(sample)) return true;
  if (sheetIsManualEntryOnly(sel)) return true;
  return false;
}


function isLocationLike(v) {
  return typeof v === "string" && v.includes("~") || (typeof v === "string" && /shaft|giếng|span|nhịp/i.test(v));
}
function isPipeLike(v) {
  return typeof v === "string" && /^S\d+$/i.test(v.trim());
}
function isPersonLike(v) {
  if (typeof v !== "string" || !v.trim()) return false;
  // Case-insensitive: signature-block samples are routinely typed in ALL
  // CAPS in the real filled documents (e.g. "VŨ KIÊN CHUNG") while
  // personnel.json stores names in normal case ("Vũ Kiên Chung") - an
  // exact-match check silently missed every one of these until fixed
  // 2026-08-08 (confirmed against a real PJ-06 export where the dropdown
  // never appeared for this exact reason).
  const all = allPersonnelNames().map(n => n.toLowerCase());
  return all.includes(v.trim().toLowerCase());
}
function allPersonnelNames() {
  return [].concat(...Object.values(personnel).map(g => g.map(p => p.name))).filter(Boolean);
}

/* ---------- Name / Position signature-block dropdowns (all sheets) ----------
   Nearly every sheet ends with a "Name/ Tên:" + "Position/ Vị trí:"
   signature block. Detecting these by CONTEXT (the fixed label text) is
   far more reliable than matching the field's sample value: sample-based
   matching missed fields whose real-world example used a name/role not
   yet listed in personnel.json, or one that only failed the old
   case-sensitive isPersonLike check (audited across all sheets
   2026-08-08 - 22 Name-context and 30 Position-context fields total).
   Matched on the ENGLISH label prefix only ("^Name/", "^Position/") -
   "Location/ Vị trí:" also contains the Vietnamese word "Vị trí" and
   would otherwise be a false positive. */
function isNameContextField(f) {
  return typeof f.context === "string" && /^Name\//.test(f.context.trim());
}
function isPositionContextField(f) {
  return typeof f.context === "string" && /^Position\//.test(f.context.trim());
}

/* Canonical "role_en/role_vi" combos from personnel.json, plus a couple
   of real position titles seen across historical exports that have no
   matching structured personnel.json entry (the site supervision
   consultant's title varies by document - "Inspector/TVGS" in some,
   "Surveyor Engineer/KS TĐ" in PJ-04's). Anything else not listed here is
   still reachable via "Khác (nhập tay)...". */
const EXTRA_POSITION_OPTIONS = ["Inspector/TVGS", "Surveyor Engineer/KS TĐ"];
function allPositionOptions() {
  const fromPersonnel = [].concat(...Object.values(personnel).map(g => g))
    .filter(p => p.role_en && p.role_vi)
    .map(p => `${p.role_en}/${p.role_vi}`);
  return [...new Set([...fromPersonnel, ...EXTRA_POSITION_OPTIONS])];
}

/* ---------- Accept/Reject checkbox dropdown ----------
The X/Z columns (Wingdings 2 font) across the ITP checksheets are visual
checkboxes: typing the exact glyph character renders as a checked/unchecked
box in Excel. Free-text entry makes this easy to mistype (wrong case, wrong
character entirely - it fails silently, the cell just shows a blank or wrong
glyph with no error). Most sheets use "R" for Accept and "£" for Reject, but
PJ-02a uses "*" for Reject instead - preserve that distinction exactly rather
than unifying it, since changing PJ-02a's glyph would visually break its
checkbox rendering (verified against each sheet's actual Wingdings 2 cells,
2026-08-06). */
const CHECKBOX_SYMBOLS = {
  "PJ-01": { accept: "R", reject: "£" },
  "PJ-02": { accept: "R", reject: "£" },
  "PJ-02a": { accept: "R", reject: "*" },
  "PJ-04": { accept: "R", reject: "£" },
  "PJ-07": { accept: "R", reject: "£" },
  "PJ-02b": { accept: "R", reject: "£" },
};

function buildAcceptRejectSelect(sheet) {
  const symbols = CHECKBOX_SYMBOLS[sheet] || { accept: "R", reject: "£" };
  const sel = document.createElement("select");
  const blank = document.createElement("option");
  blank.value = "";
  blank.textContent = "-- để trống --";
  sel.appendChild(blank);
  const accept = document.createElement("option");
  accept.value = symbols.accept;
  accept.textContent = "Đạt/ Accept";
  sel.appendChild(accept);
  const reject = document.createElement("option");
  reject.value = symbols.reject;
  reject.textContent = "Không đạt/ Reject";
  sel.appendChild(reject);
  return sel;
}

/* ---------- PJ-06 Leakage Test: computed PI / TI / Duration ----------
   Per XuBi's confirmed answer to Q5 (Xac_nhan_Q1-Q9_29072026.xlsx): PI and
   TI are each a single template cell holding a whole sentence ("PI =
   0.5liters x D x L / 1 hour = ... (lit/h)", "TI = V/T = ... (lit/h)") -
   the APP must compute the number from the real D/L/V/duration and splice
   it in, and must NEVER split them into separate label/value cells or
   leave them as free-text fields the user retypes by hand. Confirmed
   against a real export on 2026-08-07: leaving them as plain text fields
   let the field-map's sample hint ("85.20") get typed in verbatim and
   shipped as if it were real data for an unrelated pipe/length. Duration
   (X29) has no field-map entry at all - the raw template hardcodes a
   stale "1 hour" there - so it's recomputed here from the same start/
   finish times as TI and patched directly on export (no UI field). */
const PJ06_CELLS = {
  diameter: "P24",  // free text, e.g. "800 mm"
  length: "P25",    // free text, e.g. "213.006 m"
  pi: "O26",
  timeStart: "O28", // <input type="time">
  timeFinish: "O29",
  duration: "X29",  // not in field_map - patched directly, no UI field
  volume: "P30",    // free text, e.g. "0 liter"
  ti: "O31",
};

function isPJ06ComputedCell(sel, cell) {
  return sel.fileKey === "Final5-10_ITP" && sel.sheet === "PJ-06" &&
    (cell === PJ06_CELLS.pi || cell === PJ06_CELLS.ti);
}

/* P30's raw template cell already contains "     liter" (spaces + the
   fixed unit word) - a number-only field patched into the blank before
   that suffix, same shape as O24/O25's "D = "/"L =" prefix cells but with
   the fixed text AFTER the value instead of before. A plain free-text
   field here made the unit easy to forget entirely (typing just "30"
   loses "liter" on export, since the whole cell gets overwritten) -
   confirmed against a real PJ-06 export on 2026-08-08. */
function isPJ06VolumeCell(sel, cell) {
  return sel.fileKey === "Final5-10_ITP" && sel.sheet === "PJ-06" && cell === PJ06_CELLS.volume;
}

function parseLeadingNumber(s) {
  if (typeof s !== "string") return NaN;
  const m = s.match(/-?\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : NaN;
}

function computeDurationHours(startStr, finishStr) {
  if (!startStr || !finishStr) return null;
  const [sh, sm] = startStr.split(":").map(Number);
  const [fh, fm] = finishStr.split(":").map(Number);
  if ([sh, sm, fh, fm].some(n => !Number.isFinite(n))) return null;
  let hours = (fh + fm / 60) - (sh + sm / 60);
  if (hours <= 0) hours += 24; // crosses midnight
  return hours > 0 ? hours : null;
}

function formatDuration(hours) {
  const rounded = Math.round(hours * 100) / 100;
  return `${rounded} hour${rounded === 1 ? "" : "s"}`;
}

function buildPIText(diameterRaw, lengthRaw) {
  const dMm = parseLeadingNumber(diameterRaw);
  const lM = parseLeadingNumber(lengthRaw);
  const value = (Number.isFinite(dMm) && Number.isFinite(lM))
    ? (0.5 * (dMm / 1000) * lM).toFixed(2)
    : "…............";
  return `PI = 0.5liters x D x L / 1 hour = ${value} (lit/h)`;
}

function buildTIText(volumeRaw, hours) {
  const v = parseLeadingNumber(volumeRaw);
  const value = (Number.isFinite(v) && hours) ? (v / hours).toFixed(2) : "…..";
  return `TI = V/T = ${value} (lit/h)`;
}

/* Wires live recompute so the PI/TI preview shown in the app always
   matches what will be written to the exported file - the user sees the
   real sentence with the real number, never a value they have to compute
   or transcribe themselves. */
function wirePJ06Computations(container) {
  const byCell = ref => container.querySelector(`[data-cell="${ref}"]`);
  // Delegate on the whole container instead of attaching listeners to each
  // dependency input individually: always re-reads the current DOM state
  // of every dependency on ANY input/change bubbling up from the form, so
  // there's no fixed set of listeners that can miss an event from a native
  // <input type="time"> picker or fall out of sync with a re-rendered
  // field (confirmed a real "PI/TI stop updating" report was reproducible
  // with per-input listeners on some browsers - 2026-08-08).
  const recomputeAll = () => {
    const val = ref => byCell(ref)?.value ?? "";
    const piInput = byCell(PJ06_CELLS.pi);
    if (piInput) piInput.value = buildPIText(val(PJ06_CELLS.diameter), val(PJ06_CELLS.length));
    const tiInput = byCell(PJ06_CELLS.ti);
    if (tiInput) {
      const hours = computeDurationHours(val(PJ06_CELLS.timeStart), val(PJ06_CELLS.timeFinish));
      tiInput.value = buildTIText(val(PJ06_CELLS.volume), hours);
    }
  };

  container.addEventListener("input", recomputeAll);
  container.addEventListener("change", recomputeAll);
  recomputeAll();
}

/* ---------- PJ-02 Design / Actual / Discrepancy triples ----------
   "II. Inspection of Alignment" (X, Y) and "III. Inspection of Elevation
   for sewer bottom" (H1, Slope) each have a Discrepancy cell that must
   always equal Actual - Design (confirmed by XuBi: "Δ = Giá trị thực tế
   − Giá trị thiết kế") - never something the user computes by hand and
   types in, same reasoning as PJ-06's PI/TI. X and Y's Design/Actual
   cells (D54/J54/D55/J55) ALSO each hold a fixed "X = "/"Y = " label
   baked into the same cell as the value, identical in shape to this same
   sheet's D45-D49 equipment fields - those already have append:true in
   field_map.json, but D54/J54/D55/J55 were missed during field-map
   generation. Compensated for here (forcing dataset.append) rather than
   hand-editing the generated data file. H1/Slope's cells have no fixed
   label at all, so their Discrepancy is a bare number. */
const PJ02_TRIPLES = [
  { design: "D54", actual: "J54", discrepancy: "P54" },
  { design: "D55", actual: "J55", discrepancy: "P55" },
  { design: "D58", actual: "J58", discrepancy: "P58" },
  { design: "D59", actual: "J59", discrepancy: "P59" },
];
const PJ02_APPEND_CELLS = new Set(["D54", "J54", "D55", "J55"]);
const PJ02_DISCREPANCY_CELLS = new Set(PJ02_TRIPLES.map(t => t.discrepancy));

function isPJ02Sheet(sel) {
  return sel.fileKey === "Final5-10_ITP" && sel.sheet === "PJ-02";
}
function pj02NeedsForcedAppend(sel, cell) {
  return isPJ02Sheet(sel) && PJ02_APPEND_CELLS.has(cell);
}
function isPJ02DiscrepancyCell(sel, cell) {
  return isPJ02Sheet(sel) && PJ02_DISCREPANCY_CELLS.has(cell);
}

function roundClean(n, decimals = 3) {
  const factor = 10 ** decimals;
  return Math.round(n * factor) / factor;
}

function buildDiscrepancyValue(templateValue, designRaw, actualRaw) {
  const prefix = templateValue ? templateValue.trim() : "";
  const d = parseLeadingNumber(designRaw);
  const a = parseLeadingNumber(actualRaw);
  if (!Number.isFinite(d) || !Number.isFinite(a)) {
    return prefix ? `${prefix} …..` : "";
  }
  const delta = roundClean(a - d);
  return prefix ? `${prefix} ${delta}` : String(delta);
}

function wirePJ02Computations(container, fields) {
  const fieldByCell = {};
  fields.forEach(f => { fieldByCell[f.cell] = f; });
  const byCell = ref => container.querySelector(`[data-cell="${ref}"]`);

  const recomputeAll = () => {
    PJ02_TRIPLES.forEach(({ design, actual, discrepancy }) => {
      const discInput = byCell(discrepancy);
      if (!discInput) return;
      const tv = fieldByCell[discrepancy]?.template_value;
      discInput.value = buildDiscrepancyValue(tv, byCell(design)?.value, byCell(actual)?.value);
    });
  };

  container.addEventListener("input", recomputeAll);
  container.addEventListener("change", recomputeAll);
  recomputeAll();
}

/* ---------- Repeating-row table headers (PJ-03a, PJ-04, PJ-05) ----------
   These 3 sheets each hold a per-point/per-joint table that was expanded
   from a handful of template rows to 20-45 (project history: PJ-04 went
   5->20 points, PJ-05 went 7->30 injection points). The field-map
   generator captured a sample value for every expanded row but never a
   header/context, because the real column headers live in ONE shared
   header row above the whole block, not per-cell - so 322/333 of PJ-04's
   fields, 182/193 of PJ-05's, and 93/103 of PJ-03a's fell back to a bare
   "Ô <cell>" label with no indication of which physical column of the
   printed table they belong to (confirmed unusable against a real
   screenshot, 2026-08-08). Column headers and the exact repeat/grouping
   arithmetic below were extracted directly from each real template's
   header row and verified formulas (PJ-03a's 3 side-by-side column-groups
   restart their STT via `=<prevGroupLastCell>+1`, confirmed to start at
   1/16/31; PJ-04's Coordinates block cycles X/Y/Z every 3 rows starting
   row 25; PJ-05's Grout Injection block is a flat 1-row-per-point table
   starting row 25). */
function colLetters(cellRef) { return cellRef.match(/^[A-Z]+/)[0]; }
function cellRowNumber(cellRef) { return parseInt(cellRef.match(/\d+/)[0], 10); }

const REPEATING_ROW_TABLES = {
  "PJ-05": {
    startRow: 25,
    groupSize: 1,
    groupLabel: (n) => `Điểm bơm vữa #${n}`,
    columns: {
      D: "Pipe No./ Tên Cống",
      G: "Chainage of sewer pipe unit/ Lý trình của đốt cống",
      L: "Start/ Bắt đầu",
      P: "Finish/ Kết thúc",
      T: "Pressure/ Áp lực bơm (kg/cm2)",
      X: "Volume/ Thể tích bơm (m3)",
    },
  },
  "PJ-04": {
    startRow: 25,
    groupSize: 3,
    axisLabels: ["X", "Y", "Z"],
    groupLabel: (n, axis) => `Điểm đo #${n} — trục ${axis}`,
    columns: {
      L: "Design/ Thiết kế",
      P: "Actual/ Thực tế",
      T: "Tolerance/ Sai số (mm)",
      X: "Accept/ Đạt",
      Z: "Reject/ Loại",
    },
  },
};

const PJ03A_GROUPS = [
  { joint: "C", width: "F", sttBase: 1 },
  { joint: "L", width: "O", sttBase: 16 },
  { joint: "U", width: "X", sttBase: 31 },
];
const PJ03A_START_ROW = 29;
const PJ03A_COLUMNS = { joint: "Joint/ Mối nối", width: "Cushion ring width/ Bề rộng vòng đệm (m)" };

/* Returns {group, column} to synthesize a missing f.context, or null if
   this cell isn't part of a known repeating table. `group` becomes the
   fieldset legend (so every column of the same point/joint is grouped
   together, same as fields that already have a real context); `column`
   becomes the per-field hint shown next to its cell reference. */
function syntheticFieldContext(sel, f) {
  if (sel.fileKey !== "Final5-10_ITP") return null;
  const col = colLetters(f.cell);
  const row = cellRowNumber(f.cell);

  if (sel.sheet === "PJ-03a") {
    const group = PJ03A_GROUPS.find(g => g.joint === col || g.width === col);
    if (!group || row < PJ03A_START_ROW) return null;
    const stt = group.sttBase + (row - PJ03A_START_ROW);
    return { group: `Mối nối #${stt}`, column: col === group.joint ? PJ03A_COLUMNS.joint : PJ03A_COLUMNS.width };
  }

  const cfg = REPEATING_ROW_TABLES[sel.sheet];
  if (!cfg || row < cfg.startRow) return null;
  const column = cfg.columns[col];
  if (!column) return null;
  const pointNumber = Math.floor((row - cfg.startRow) / cfg.groupSize) + 1;
  if (cfg.groupSize > 1) {
    const axis = cfg.axisLabels[(row - cfg.startRow) % cfg.groupSize];
    return { group: cfg.groupLabel(pointNumber, axis), column };
  }
  return { group: cfg.groupLabel(pointNumber), column };
}

function renderXlsxFields(container, fields, fileKey, sheet) {
  const sel = { fileKey, sheet };
  if (sheetIsManualEntryOnly(sel)) {
    const note = document.createElement("p");
    note.className = "hint";
    note.style.color = "#b36b00";
    note.textContent = "Lưu ý: toàn bộ ô của form này (thông số/model/số kiểm định thiết bị) luôn là ô nhập tay thủ công, không gợi ý dropdown, chiếu thẳng theo từng ô đã xác định.";
    container.appendChild(note);
  }

  // Group consecutive fields that share the same row-context under one
  // fieldset (e.g. all 4 spec columns of "Jacking Machine" together) -
  // much easier to scan than fixed chunks of 25 unrelated cells.
  const groups = [];
  let current = null;
  for (const f of fields) {
    const synthetic = f.context ? null : syntheticFieldContext(sel, f);
    const label = f.context || synthetic?.group;
    const key = label || `__cell_${f.cell}`;
    if (current && current.key === key) {
      current.fields.push(f);
    } else {
      current = { key, label, fields: [f] };
      groups.push(current);
    }
  }

  groups.forEach(g => {
    const fs = document.createElement("fieldset");
    const legend = document.createElement("legend");
    legend.textContent = g.label
      ? truncate(g.label.replace(/\n/g, " / "), 70)
      : `Ô ${g.fields[0].cell}`;
    fs.appendChild(legend);
    g.fields.forEach(f => fs.appendChild(buildXlsxFieldRow(f, sel)));
    container.appendChild(fs);
  });

  if (fileKey === "Final5-10_ITP" && sheet === "PJ-06") {
    wirePJ06Computations(container);
  } else if (fileKey === "Final5-10_ITP" && sheet === "PJ-02") {
    wirePJ02Computations(container, fields);
  }
}

function buildXlsxFieldRow(f, sel) {
  const row = document.createElement("div");
  row.className = "field-row";
  const label = document.createElement("label");
  const isComputed = isPJ06ComputedCell(sel, f.cell);
  const isVolumeField = isPJ06VolumeCell(sel, f.cell);
  const isPJ02Discrepancy = isPJ02DiscrepancyCell(sel, f.cell);
  const isPJ02Append = pj02NeedsForcedAppend(sel, f.cell);
  const isNameField = isNameContextField(f);
  const isPositionField = isPositionContextField(f);
  // Real context missing from field_map (PJ-03a/PJ-04/PJ-05's expanded
  // repeating-row tables) falls back to the column header reconstructed
  // from the real template - see syntheticFieldContext for why.
  const synthetic = f.context ? null : syntheticFieldContext(sel, f);
  // Group legend already shows the row context, so here just show the
  // cell reference plus a short sample-value hint (still useful to
  // distinguish which of several columns in the same group this is,
  // e.g. Model vs Manufacturer vs Certificate No.). Computed cells (PI/TI,
  // PJ-02's Discrepancy) get a hint explaining WHY there's no sample
  // value/free typing here, instead of showing a stale example number as
  // if it were editable.
  label.innerHTML = isComputed
    ? `<span class="cellref">${f.cell}</span> <span class="ctx">(tự tính, không nhập tay)</span>`
    : isPJ02Discrepancy
    ? `<span class="cellref">${f.cell}</span> <span class="ctx">(tự tính = Thực tế − Thiết kế)</span>`
    : isVolumeField
    ? `<span class="cellref">${f.cell}</span> <span class="ctx">(chỉ nhập số, "liter" tự thêm vào)</span>`
    : synthetic
    ? `<span class="cellref">${f.cell}</span> <span class="ctx">${synthetic.column}${f.sample_value ? ` (vd: ${truncate(f.sample_value, 20)})` : ""}</span>`
    : f.sample_value
    ? `<span class="cellref">${f.cell}</span> <span class="ctx">(vd: ${truncate(f.sample_value, 30)})</span>`
    : `<span class="cellref">${f.cell}</span>`;
  row.appendChild(label);

  // Name/Position signature-block fields are legitimate personnel picks
  // even on sheets otherwise forced to manual entry (e.g. PJ-01 has its
  // own Name/Position fields despite MANUAL_ENTRY_SHEETS applying to the
  // rest of that sheet's equipment-spec cells).
  const forceManual = sheetIsManualEntryOnly(sel) && !isNameField && !isPositionField;
  const forceText = mustForceText(f.cell, sel, f.sample_value);

  let input;
  let otherInput = null;
  const sample = f.sample_value;
  if (isComputed) {
    input = document.createElement("input");
    input.type = "text";
    input.readOnly = true;
    input.classList.add("computed-field");
    input.value = f.cell === PJ06_CELLS.pi ? buildPIText("", "") : buildTIText("", null);
  } else if (isPJ02Discrepancy) {
    input = document.createElement("input");
    input.type = "text";
    input.readOnly = true;
    input.classList.add("computed-field");
    input.dataset.deltaPrefix = f.template_value ? f.template_value.trim() : "";
    input.value = buildDiscrepancyValue(f.template_value, "", "");
  } else if (isVolumeField) {
    input = document.createElement("input");
    input.type = "number";
    input.step = "any";
    input.placeholder = "vd: 30";
    input.dataset.unitSuffix = (f.template_value || "liter").trim();
  } else if (!forceManual && isNameField) {
    ({ select: input, otherInput } = buildSelectWithOther(allPersonnelNames()));
    input.dataset.uppercase = "true";
  } else if (!forceManual && isPositionField) {
    ({ select: input, otherInput } = buildSelectWithOther(allPositionOptions()));
  } else if (!forceManual && f.type === "accept_reject") {
    input = buildAcceptRejectSelect(sel.sheet);
  } else if (!forceManual && isLocationLike(sample)) {
    ({ select: input, otherInput } = buildSelectWithOther(locations));
  } else if (!forceManual && isPipeLike(sample)) {
    ({ select: input, otherInput } = buildSelectWithOther(pipeNames));
  } else if (!forceManual && isPersonLike(sample)) {
    ({ select: input, otherInput } = buildSelectWithOther(allPersonnelNames()));
    input.dataset.uppercase = "true";
  } else if (!forceManual && f.type === "date") {
    input = document.createElement("input");
    input.type = "date";
  } else if (!forceManual && f.type === "time") {
    input = document.createElement("input");
    input.type = "time";
  } else if (!forceManual && f.type === "number") {
    input = document.createElement("input");
    input.type = "number";
    input.step = "any";
  } else {
    input = document.createElement("input");
    input.type = "text";
  }
  input.dataset.cell = f.cell;
  input.dataset.type = forceText ? "text-forced" : f.type;
  if (f.append || isPJ02Append) {
    input.dataset.append = "true";
    input.placeholder = "Nhập giá trị, sẽ tự nối vào cuối dòng nhãn";
  }
  row.appendChild(input);
  if (otherInput) row.appendChild(otherInput);
  return row;
}




function buildSelect(options, def) {
  const sel = document.createElement("select");
  const blank = document.createElement("option");
  blank.value = "";
  blank.textContent = "-- để trống --";
  sel.appendChild(blank);
  options.forEach(o => {
    const opt = document.createElement("option");
    opt.value = o;
    opt.textContent = o;
    sel.appendChild(opt);
  });
  const other = document.createElement("option");
  other.value = "__other__";
  other.textContent = "Khác (nhập tay)...";
  sel.appendChild(other);
  return sel;
}

/* Pairs a dropdown with an inline text field for its "Khác (nhập tay)..."
   option, so a manual value can be typed directly in the app instead of
   being silently skipped on export (forcing the user to open the exported
   Excel/Word file and type it in there afterward). The text field is the
   very next sibling of the select in the DOM - export reads it via
   select.nextElementSibling rather than a shared data-cell/data-tag, so a
   single querySelectorAll("[data-cell]"/"[data-tag]") pass never double-
   counts the pair. */
function buildSelectWithOther(options) {
  const select = buildSelect(options, "");
  const otherInput = document.createElement("input");
  otherInput.type = "text";
  otherInput.className = "other-input hidden";
  otherInput.placeholder = "Nhập giá trị...";
  select.addEventListener("change", () => {
    if (select.value === "__other__") {
      otherInput.classList.remove("hidden");
      otherInput.focus();
    } else {
      otherInput.classList.add("hidden");
      otherInput.value = "";
    }
  });
  return { select, otherInput };
}

/* Resolves an input's effective value, substituting the paired manual-entry
   text field whenever a select is currently set to "Khác (nhập tay)...". */
function effectiveValue(inp) {
  if (inp.tagName === "SELECT" && inp.value === "__other__") {
    const other = inp.nextElementSibling;
    return (other && other.classList.contains("other-input")) ? other.value.trim() : "";
  }
  return inp.value;
}

function truncate(s, n) { return s.length > n ? s.slice(0, n) + "…" : s; }

/* ---------- docx form rendering ---------- */
function renderDocxFields(container, fields) {
  const fs = document.createElement("fieldset");
  const legend = document.createElement("legend");
  legend.textContent = "Thông tin biên bản";
  fs.appendChild(legend);

  fields.forEach(f => {
    const row = document.createElement("div");
    row.className = "field-row";
    const label = document.createElement("label");
    label.textContent = f.label;
    row.appendChild(label);

    let input;
    let otherInput = null;
    if (f.type === "dropdown-location") {
      ({ select: input, otherInput } = buildSelectWithOther(locations));
    } else if (f.type === "dropdown-personnel") {
      const names = (personnel[f.group] || []).map(p => p.name);
      ({ select: input, otherInput } = buildSelectWithOther(names));
      input.addEventListener("change", () => autoFillRole(f.tag, input.value));
    } else if (f.type === "dropdown-material") {
      const names = materials.map(m => m.name_en + " / " + m.name_vi);
      ({ select: input, otherInput } = buildSelectWithOther(names));
      input.addEventListener("change", () => autoFillMaterialSource(f.tag, input.value));
    } else if (f.type === "text-linked" || f.type === "text-linked-material") {
      input = document.createElement("input");
      input.type = "text";
      input.dataset.linked = "true";
    } else if (f.type === "day") {
      input = document.createElement("input"); input.type = "number"; input.min = 1; input.max = 31; input.placeholder = "DD";
    } else if (f.type === "month") {
      input = document.createElement("input"); input.type = "number"; input.min = 1; input.max = 12; input.placeholder = "MM";
    } else if (f.type === "year") {
      input = document.createElement("input"); input.type = "number"; input.min = 2020; input.max = 2100; input.placeholder = "YYYY";
    } else if (f.type === "number") {
      input = document.createElement("input"); input.type = "number"; input.step = "any";
    } else {
      input = document.createElement("input"); input.type = "text";
    }
    input.dataset.tag = f.tag;
    row.appendChild(input);
    if (otherInput) row.appendChild(otherInput);
    fs.appendChild(row);
  });
  container.appendChild(fs);
}

function autoFillRole(nameTag, selectedName) {
  const roleTag = nameTag.replace("ten_", "chuc_vu_");
  const roleInput = document.querySelector(`[data-tag="${roleTag}"]`);
  if (!roleInput) return;
  const all = [].concat(...Object.values(personnel).map(g => g));
  const person = all.find(p => p.name === selectedName);
  roleInput.value = person ? person.role_vi : "";
}

function autoFillMaterialSource(nameTag, selectedLabel) {
  const m = materials.find(mm => (mm.name_en + " / " + mm.name_vi) === selectedLabel);
  const enInput = document.querySelector(`[data-tag="${nameTag.replace("_ten", "_nguon_goc_en")}"]`);
  const viInput = document.querySelector(`[data-tag="${nameTag.replace("_ten", "_nguon_goc_vi")}"]`);
  if (enInput) enInput.value = m ? m.source_en : "";
  if (viInput) viInput.value = m ? m.source_vi : "";
}

/* ---------- Export logic ---------- */
async function doExport() {
  const statusHolder = document.getElementById("formArea");
  let status = document.getElementById("formStatus");
  if (!status) {
    status = document.createElement("div");
    status.id = "formStatus";
    status.className = "status";
    statusHolder.appendChild(status);
  }
  status.textContent = "Đang xuất file...";
  status.className = "status";

  try {
    if (currentSelection.kind === "xlsx") {
      await exportXlsx();
    } else {
      await exportDocx();
    }
    status.textContent = "Xuất file thành công.";
    status.className = "status ok";
  } catch (e) {
    console.error(e);
    let detail = e.message;
    // docxtemplater throws a MultiError whose real information is in
    // e.properties.errors (an array of individual TemplateError objects).
    // The top-level e.message alone ("Multi error") is useless - surface
    // every sub-error's explanation and exact tag location instead.
    if (e.properties && Array.isArray(e.properties.errors)) {
      console.error("docxtemplater sub-errors:", e.properties.errors);
      const lines = e.properties.errors.map((err, i) => {
        const p = err.properties || {};
        const loc = p.xtag ? ` (tag: "${p.xtag}")` : "";
        const explanation = p.explanation || err.message || String(err);
        return `${i + 1}. ${explanation}${loc}`;
      });
      detail = lines.join(" | ");
    }
    status.textContent = "Lỗi khi xuất file: " + detail;
    status.className = "status err";
  }
}

/* ---------- XLSX export via raw XML surgery (PizZip) ----------
   Deliberately NOT using SheetJS read+rewrite: the free/community build of
   SheetJS does not reliably round-trip print areas, page setup, merged
   cells and embedded images. Instead we patch only the <c> elements for
   cells the user actually filled in, directly inside the original
   worksheet XML, leaving every other byte of the template untouched -
   the same technique used throughout this project when editing these
   files with Python/openpyxl. */

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildCellXml(cellRef, styleAttr, value, isNumber) {
  const s = styleAttr ? ` s="${styleAttr}"` : "";
  if (isNumber) {
    return `<c r="${cellRef}"${s}><v>${value}</v></c>`;
  }
  // inline string: does not require touching sharedStrings.xml at all
  return `<c r="${cellRef}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function colLettersToNumber(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function patchCellInSheetXml(xml, cellRef, value, isNumber) {
  // IMPORTANT: check self-closing FIRST. A self-closing cell like
  // <c r="F14" s="23"/> was previously mis-matched by the "with content"
  // regex below (its "attrs" group greedily swallowed the trailing "/",
  // then searched forward for the NEXT "</c>" anywhere in the XML -
  // silently deleting every cell in between). Self-closing cells must
  // never reach that regex at all.
  const selfClosing = new RegExp(`<c r="${cellRef}"([^>]*)/>`);
  // Negative lookbehind ensures this only matches a true opening tag
  // ("...">) and never a self-closing one ("..."/>), even as a fallback.
  const withContent = new RegExp(`<c r="${cellRef}"([^>]*?)(?<!/)>((?:(?!</c>)[\\s\\S])*?)</c>`);

  let styleAttr = null;
  let m = xml.match(selfClosing) || xml.match(withContent);
  if (m) {
    const sMatch = m[1].match(/s="(\d+)"/);
    if (sMatch) styleAttr = sMatch[1];
  }

  const newCell = buildCellXml(cellRef, styleAttr, value, isNumber);

  if (xml.match(selfClosing)) {
    return xml.replace(selfClosing, newCell);
  }
  if (xml.match(withContent)) {
    return xml.replace(withContent, newCell);
  }

  // Cell doesn't exist yet in the XML (template only materializes non-empty
  // cells). OOXML REQUIRES <c> children of a <row> to stay in strict column
  // order, so we must insert the new cell in the correct position relative
  // to whichever sibling cells already exist - never blindly at the start.
  const rowNum = cellRef.match(/\d+/)[0];
  const targetCol = colLettersToNumber(cellRef.match(/^[A-Z]+/)[0]);

  const rowOpenMatch = xml.match(new RegExp(`<row r="${rowNum}"[^>]*>`));
  const rowSelfMatch = xml.match(new RegExp(`<row r="${rowNum}"[^>]*/>`));

  if (rowSelfMatch) {
    // empty row element (no cells at all yet) - safe to just add the one cell
    const tag = rowSelfMatch[0];
    const openTag = tag.slice(0, -2) + ">"; // "/>" -> ">"
    return xml.replace(tag, `${openTag}${newCell}</row>`);
  }

  if (rowOpenMatch) {
    const rowStart = rowOpenMatch.index + rowOpenMatch[0].length;
    const rowEndIdx = xml.indexOf("</row>", rowStart);
    if (rowEndIdx === -1) {
      console.warn("Malformed row (no closing tag) for", cellRef, "- value not written");
      return xml;
    }
    const rowInner = xml.slice(rowStart, rowEndIdx);
    const existingCells = [...rowInner.matchAll(/<c r="([A-Z]+)\d+"[^>]*(?:\/>|>.*?<\/c>)/gs)];

    let insertAt = rowStart; // default: before everything (row has no cells yet)
    for (const cellMatch of existingCells) {
      const cCol = colLettersToNumber(cellMatch[1]);
      if (cCol < targetCol) {
        insertAt = rowStart + cellMatch.index + cellMatch[0].length;
      } else {
        break;
      }
    }
    return xml.slice(0, insertAt) + newCell + xml.slice(insertAt);
  }

  console.warn("Could not locate row for", cellRef, "- value not written");
  return xml;
}

/* ---------- Strip workbook down to a single sheet on export ----------
   Mirrors the same safe technique used earlier in this project with
   Python/openpyxl (see clean_xlsx_lib.py): remove every other <sheet>
   entry from workbook.xml, remap the kept sheet's Print_Area/Print_Titles
   defined names to localSheetId=0, drop the now-unused worksheet parts,
   their rels, and any drawings/media that are ONLY referenced by removed
   sheets - all while leaving the target sheet's own XML completely
   untouched (already patched separately by patchCellInSheetXml). */
function stripToSingleSheet(zip, targetSheetFile) {
  const wbXml = zip.file("xl/workbook.xml").asText();
  const relsXml = zip.file("xl/_rels/workbook.xml.rels").asText();
  const ctXml = zip.file("[Content_Types].xml").asText();
  const appXml = zip.file("docProps/app.xml") ? zip.file("docProps/app.xml").asText() : null;

  const sheetTags = [...wbXml.matchAll(/<sheet [^>]*\/>/g)].map(m => m[0]);
  const relEntries = [...relsXml.matchAll(/<Relationship [^>]*\/>/g)].map(m => m[0]);
  const relMap = {}; // rId -> target
  relEntries.forEach(r => {
    const id = r.match(/Id="(rId\d+)"/)[1];
    const target = r.match(/Target="([^"]+)"/)[1];
    relMap[id] = target;
  });

  let keepTag = null, keepIndex = -1;
  sheetTags.forEach((tag, i) => {
    const rid = tag.match(/r:id="(rId\d+)"/)[1];
    const target = relMap[rid];
    if (target && target.endsWith("/" + targetSheetFile)) {
      keepTag = tag;
      keepIndex = i;
    }
  });
  if (!keepTag) {
    console.warn("stripToSingleSheet: could not identify target sheet, skipping strip");
    return;
  }
  const keepRid = keepTag.match(/r:id="(rId\d+)"/)[1];

  // 1. workbook.xml: keep only the target <sheet>, remap defined names
  let wb2 = wbXml.replace(/<sheets>.*?<\/sheets>/s, `<sheets>${keepTag}</sheets>`);

  // The template's <workbookView activeTab="N"> points at whichever tab was
  // last active when the ORIGINAL 9-sheet file was saved (e.g. "8" = the
  // 9th/last tab). Once every sheet but one is removed, that index almost
  // always points past the end of the now-1-entry <sheets> list. Windows
  // Excel and LibreOffice silently tolerate this and re-clamp it, but real
  // Excel for Mac does not - it's a confirmed cause of "can't open this
  // file" on Mac (root-caused 2026-08-06 against an exported PJ-02 file).
  // Simplest safe fix: drop the attribute; it defaults to the first (only)
  // tab, which is exactly what we want since exactly one sheet remains.
  wb2 = wb2.replace(/\s*activeTab="\d+"/, "");

  const keptDefinedNames = [];
  const dnRegex = /<definedName name="([^"]+)" localSheetId="(\d+)"([^>]*)>([^<]*)<\/definedName>/g;
  let dnMatch;
  while ((dnMatch = dnRegex.exec(wb2)) !== null) {
    const [, name, idx, attrs, val] = dnMatch;
    if (parseInt(idx, 10) === keepIndex) {
      keptDefinedNames.push({ name, attrs, val });
    }
  }
  wb2 = wb2.replace(/<definedNames>.*?<\/definedNames>/s, "");
  // externalReferences intentionally left untouched - see note above at
  // the workbook.xml.rels step.
  if (keptDefinedNames.length) {
    const dnXml = "<definedNames>" + keptDefinedNames
      .map(d => `<definedName name="${d.name}" localSheetId="0"${d.attrs}>${d.val}</definedName>`)
      .join("") + "</definedNames>";
    // CT_Workbook requires strict child-element order per the OOXML schema:
    // ... sheets, functionGroups?, externalReferences?, definedNames?, calcPr? ...
    // definedNames must come AFTER externalReferences, never before. Since
    // externalReferences is deliberately kept intact (see note below),
    // inserting definedNames right after </sheets> puts it BEFORE
    // externalReferences whenever the latter is present - a schema-order
    // violation that real Excel (Win + Mac) rejects outright as corrupt,
    // even though LibreOffice and openpyxl silently tolerate it. Insert
    // after </externalReferences> when present; fall back to right after
    // </sheets> only for templates that have no external references at all.
    if (wb2.includes("</externalReferences>")) {
      wb2 = wb2.replace("</externalReferences>", "</externalReferences>" + dnXml);
    } else {
      wb2 = wb2.replace("</sheets>", "</sheets>" + dnXml);
    }
  }

  // docProps/app.xml: regenerate HeadingPairs/TitlesOfParts so they describe
  // the single kept sheet instead of the original template's full sheet
  // list. Like the activeTab fix above, this is metadata Excel itself
  // maintains on save and never validates against the actual <sheets> count
  // when it's simply reading a file - except real Excel for Mac, which does
  // check it and refuses to open the file when the two disagree. Excel only
  // lists NON-hidden defined names here (built-in Print_Area/Print_Titles),
  // never the hidden legacy helper names, so filter accordingly.
  if (appXml) {
    const keepSheetName = keepTag.match(/name="([^"]+)"/)[1];
    const visibleNames = keptDefinedNames.filter(d => !/hidden="1"/.test(d.attrs));
    const titleParts = [
      keepSheetName,
      ...visibleNames.map(d => `'${keepSheetName}'!${d.name.replace(/^_xlnm\./, "")}`),
    ];
    const app2 = appXml.replace(
      /<HeadingPairs>[\s\S]*?<\/TitlesOfParts>/,
      `<HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>${visibleNames.length}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="${titleParts.length}" baseType="lpstr">${titleParts.map(t => `<vt:lpstr>${xmlEscape(t)}</vt:lpstr>`).join("")}</vt:vector></TitlesOfParts>`
    );
    zip.file("docProps/app.xml", app2);
  }

  // 2. workbook.xml.rels: keep only rels needed (target sheet + shared parts)
  //    IMPORTANT: externalLink relationships are intentionally LEFT ALONE.
  //    Some hidden legacy defined names in these templates reference
  //    external workbooks by numeric index (e.g. "[9]SheetName!#REF!"),
  //    which corresponds to the Nth <externalReference> entry in
  //    workbook.xml. If we stripped externalReferences/externalLinkN.xml
  //    but left those defined names behind, the index becomes dangling -
  //    real Excel (unlike LibreOffice, which is lenient) flags this as
  //    corruption and silently deletes records to "repair" it. Since the
  //    original template already opens fine in Excel with this structure
  //    intact, the safest fix is to simply never touch it.
  const droppedTargets = [];
  const keptRels = [];
  relEntries.forEach(r => {
    const id = r.match(/Id="(rId\d+)"/)[1];
    const type = r.match(/Type="([^"]+)"/)[1];
    const target = r.match(/Target="([^"]+)"/)[1];
    if (type.endsWith("/worksheet")) {
      if (id === keepRid) keptRels.push(r);
      else droppedTargets.push(target);
    } else if (target.includes("calcChain")) {
      droppedTargets.push(target);
    } else {
      keptRels.push(r); // includes externalLink rels - kept untouched
    }
  });
  const rels2 = relsXml.replace(/(<Relationships[^>]*>).*(<\/Relationships>)/s, `$1${keptRels.join("")}$2`);

  // 3. Content_Types: drop overrides for removed parts
  let ct2 = ctXml;
  droppedTargets.forEach(t => {
    const partName = "/xl/" + t;
    ct2 = ct2.replace(new RegExp(`<Override PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`), "");
  });
  ct2 = ct2.replace(/<Override PartName="\/xl\/calcChain\.xml"[^>]*\/>/, "");

  zip.file("xl/workbook.xml", wb2);
  zip.file("xl/_rels/workbook.xml.rels", rels2);
  zip.file("[Content_Types].xml", ct2);

  // 4. remove the dropped worksheet files, their _rels, and calcChain.xml
  //    (drawings/media exclusively owned by removed sheets are left in the
  //    zip as harmless orphans - safer than risking removal of something
  //    still shared, and they add negligible size)
  droppedTargets.forEach(t => {
    const fname = "xl/" + t;
    if (zip.file(fname)) zip.remove(fname);
    const base = t.split("/").pop();
    const relFname = "xl/worksheets/_rels/" + base + ".rels";
    if (zip.file(relFname)) zip.remove(relFname);
  });
  if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");
}

function parseSharedStrings(sstXml) {
  if (!sstXml) return [];
  // Each <si> can contain a plain <t> or multiple <r><t>...</t></r> runs -
  // concatenate all <t> text within each <si> to get its full string.
  const items = [...sstXml.matchAll(/<si[^>]*>([\s\S]*?)<\/si>/g)];
  return items.map(m => {
    const texts = [...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(t => t[1]);
    return texts.join("");
  });
}

function getCurrentCellText(xml, sharedStrings, cellRef) {
  const cellMatch =
    xml.match(new RegExp(`<c r="${cellRef}"[^>]*t="s"[^>]*>\\s*<v>(\\d+)</v>\\s*</c>`)) ||
    xml.match(new RegExp(`<c r="${cellRef}"[^>]*t="s"[^>]*/>`));
  if (cellMatch && cellMatch[1] !== undefined) {
    const idx = parseInt(cellMatch[1], 10);
    return sharedStrings[idx] || "";
  }
  const inlineMatch = xml.match(new RegExp(`<c r="${cellRef}"[^>]*t="inlineStr"[^>]*>[\\s\\S]*?<t[^>]*>([^<]*)</t>`));
  if (inlineMatch) return inlineMatch[1];
  const plainMatch = xml.match(new RegExp(`<c r="${cellRef}"[^>]*>\\s*<v>([^<]*)</v>\\s*</c>`));
  if (plainMatch) return plainMatch[1];
  return "";
}

async function exportXlsx() {
  const { fileKey, sheet } = currentSelection;
  const templateFile = XLSX_TEMPLATE_DIR + XLSX_FILES[fileKey].file;
  const resp = await fetch(templateFile);
  const buf = await resp.arrayBuffer();

  const zip = new PizZip(buf);
  const sheetManifest = await loadJSON(DATA_DIR + "sheet_manifest.json");
  const sheetFile = sheetManifest[fileKey][sheet];
  const sheetPath = `xl/worksheets/${sheetFile}`;

  let xml = zip.file(sheetPath).asText();
  const sharedStrings = parseSharedStrings(
    zip.file("xl/sharedStrings.xml") ? zip.file("xl/sharedStrings.xml").asText() : null
  );

  // Recompute PJ-06's PI/TI right here from the current D/L/V/time inputs,
  // rather than trusting whatever the readonly preview field already
  // holds - guarantees the exported sentence is correct even if some edge
  // case (e.g. a value pasted in without firing an input event) left the
  // live preview stale.
  if (fileKey === "Final5-10_ITP" && sheet === "PJ-06") {
    const cellVal = ref => document.querySelector(`#dataForm [data-cell="${ref}"]`)?.value || "";
    const piInput = document.querySelector(`#dataForm [data-cell="${PJ06_CELLS.pi}"]`);
    if (piInput) piInput.value = buildPIText(cellVal(PJ06_CELLS.diameter), cellVal(PJ06_CELLS.length));
    const tiInput = document.querySelector(`#dataForm [data-cell="${PJ06_CELLS.ti}"]`);
    const durationHours = computeDurationHours(cellVal(PJ06_CELLS.timeStart), cellVal(PJ06_CELLS.timeFinish));
    if (tiInput) tiInput.value = buildTIText(cellVal(PJ06_CELLS.volume), durationHours);
  }

  // Same reasoning as PJ-06 above: recompute PJ-02's Discrepancy cells
  // (Actual - Design) from the current form state right before export,
  // rather than trusting the live preview.
  if (fileKey === "Final5-10_ITP" && sheet === "PJ-02") {
    const cellVal = ref => document.querySelector(`#dataForm [data-cell="${ref}"]`)?.value || "";
    PJ02_TRIPLES.forEach(({ design, actual, discrepancy }) => {
      const discInput = document.querySelector(`#dataForm [data-cell="${discrepancy}"]`);
      if (!discInput) return;
      discInput.value = buildDiscrepancyValue(discInput.dataset.deltaPrefix, cellVal(design), cellVal(actual));
    });
  }

  const inputs = document.querySelectorAll("#dataForm [data-cell]");
  let count = 0;
  inputs.forEach(inp => {
    let val = effectiveValue(inp);
    if (val === "") return; // leave template's blank cell untouched
    const cellRef = inp.dataset.cell;
    const isNumber = inp.dataset.type === "number";

    // P30-style fields: the raw template's default text is just a fixed
    // unit word ("liter") with blank space to fill in a number - always
    // reconstruct "<number> <unit>" rather than trusting the user typed
    // the unit themselves (the input is type=number, so it can't contain
    // "liter" anyway).
    if (inp.dataset.unitSuffix) {
      val = `${val} ${inp.dataset.unitSuffix}`;
    }
    // Name signature-block fields: every real historical example is
    // ALL CAPS - keep the exported cell consistent with that convention
    // even though the dropdown itself shows normal-case names for
    // readability.
    if (inp.dataset.uppercase === "true") {
      val = val.toLocaleUpperCase("vi-VN");
    }

    if (inp.dataset.append === "true") {
      // No dedicated value cell exists (label spans the whole row) - append
      // the user's value directly onto the existing label text instead of
      // overwriting a neighbouring cell (which would land on the checkbox
      // columns and visually "jump" to the wrong place).
      const currentLabel = getCurrentCellText(xml, sharedStrings, cellRef);
      const sep = /[:=]\s*$/.test(currentLabel) ? " " : ": ";
      const combined = currentLabel.replace(/\s+$/, "") + sep + val;
      xml = patchCellInSheetXml(xml, cellRef, combined, false);
      count++;
      return;
    }

    xml = patchCellInSheetXml(xml, cellRef, isNumber ? parseFloat(val) : val, isNumber);
    count++;
  });

  // Duration (X29) has no field-map entry - the raw template hardcodes a
  // stale "1 hour" there. Patch it directly from the same start/finish
  // times used for TI so the two numbers can never disagree.
  if (fileKey === "Final5-10_ITP" && sheet === "PJ-06") {
    const cellVal = ref => document.querySelector(`#dataForm [data-cell="${ref}"]`)?.value || "";
    const durationHours = computeDurationHours(cellVal(PJ06_CELLS.timeStart), cellVal(PJ06_CELLS.timeFinish));
    if (durationHours) {
      xml = patchCellInSheetXml(xml, PJ06_CELLS.duration, formatDuration(durationHours), false);
      count++;
    }
  }

  zip.file(sheetPath, xml);

  // Export ONLY the sheet the user actually filled in, not the whole
  // multi-sheet workbook - avoids confusion and shrinks the file.
  const totalSheetsInFile = Object.keys(sheetManifest[fileKey]).length;
  if (totalSheetsInFile > 1) {
    stripToSingleSheet(zip, sheetFile);
  }

  // Cells patched above are written as raw <c> values, never through Excel's
  // own edit path, so Excel has no reason to think any formula's precedent
  // changed. Without fullCalcOnLoad, cells like PJ-05's Total column
  // (SUM(X25:AA54)) keep showing the template's stale cached <v> (e.g. "0")
  // until the user manually forces a recalc. Force one on every open instead.
  forceFullRecalcOnLoad(zip);

  const blob = zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const outName = sanitizeFilename(`${todayStr()}-${SHEET_LABELS[sheet] || sheet}`) + ".xlsx";
  saveBlob(blob, outName);
  console.log(`Patched ${count} cell(s) in ${sheetPath}`);
}

/* Forces Excel/LibreOffice to fully recompute every formula the moment the
   exported file is opened, instead of trusting the template's cached <v>
   values - necessary because this app never edits cells through a real
   spreadsheet engine, so no calc chain is ever invalidated on its own. */
function forceFullRecalcOnLoad(zip) {
  const wbPath = "xl/workbook.xml";
  const wbFile = zip.file(wbPath);
  if (!wbFile) return;
  let wbXml = wbFile.asText();
  if (/<calcPr\b[^>]*\/>/.test(wbXml)) {
    wbXml = wbXml.replace(/<calcPr\b([^>]*)\/>/, (m, attrs) => {
      const cleaned = attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, "");
      return `<calcPr${cleaned} fullCalcOnLoad="1"/>`;
    });
  } else {
    wbXml = wbXml.replace("</workbook>", `<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
  }
  zip.file(wbPath, wbXml);
}

/* Strips characters illegal in Windows/macOS filenames while preserving
   Vietnamese diacritics and spacing, so exported names stay human-readable
   (e.g. "20260807-PJ-05 - Grout Injection.xlsx") instead of collapsing into
   underscore runs. */
function sanitizeFilename(s) {
  return String(s)
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}


async function exportDocx() {
  const { formKey } = currentSelection;
  const info = docxForms[formKey];
  const resp = await fetch(DOCX_TEMPLATE_DIR + info.template);
  const buf = await resp.arrayBuffer();
  const zip = new PizZip(buf);
  const doc = new window.docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => "",
    delimiters: { start: "{{", end: "}}" },
  });

  const data = {};
  document.querySelectorAll("#dataForm [data-tag]").forEach(inp => {
    data[inp.dataset.tag] = effectiveValue(inp) || "";
  });
  doc.render(data);

  const out = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const outName = sanitizeFilename(`${todayStr()}-${info.title}`) + ".docx";
  saveBlob(out, outName);
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function todayStr() {
  const d = new Date();
  const pad = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
}

init();
