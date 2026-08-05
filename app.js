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
  if (typeof v !== "string") return false;
  const all = [].concat(...Object.values(personnel).map(g => g.map(p => p.name)));
  return all.includes(v.trim());
}
function allPersonnelNames() {
  return [].concat(...Object.values(personnel).map(g => g.map(p => p.name))).filter(Boolean);
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
  // group into chunks of 25 for readability using <fieldset>
  const groupSize = 25;
  for (let i = 0; i < fields.length; i += groupSize) {
    const chunk = fields.slice(i, i + groupSize);
    const fs = document.createElement("fieldset");
    const legend = document.createElement("legend");
    legend.textContent = `Ô ${chunk[0].cell} – ${chunk[chunk.length - 1].cell}`;
    fs.appendChild(legend);
    chunk.forEach(f => fs.appendChild(buildXlsxFieldRow(f, sel)));
    container.appendChild(fs);
  }
}

function buildXlsxFieldRow(f, sel) {
  const row = document.createElement("div");
  row.className = "field-row";
  const label = document.createElement("label");
  label.textContent = f.sample_value ? `${f.cell} (vd: ${truncate(f.sample_value, 28)})` : f.cell;
  row.appendChild(label);

  const forceManual = sheetIsManualEntryOnly(sel);
  const forceText = mustForceText(f.cell, sel, f.sample_value);

  let input;
  const sample = f.sample_value;
  if (!forceManual && isLocationLike(sample)) {
    input = buildSelect(locations, "");
  } else if (!forceManual && isPipeLike(sample)) {
    input = buildSelect(pipeNames, "");
  } else if (!forceManual && isPersonLike(sample)) {
    input = buildSelect(allPersonnelNames(), "");
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
  row.appendChild(input);
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
    if (f.type === "dropdown-location") {
      input = buildSelect(locations, "");
    } else if (f.type === "dropdown-personnel") {
      const names = (personnel[f.group] || []).map(p => p.name);
      input = buildSelect(names, "");
      input.addEventListener("change", () => autoFillRole(f.tag, input.value));
    } else if (f.type === "dropdown-material") {
      const names = materials.map(m => m.name_en + " / " + m.name_vi);
      input = buildSelect(names, "");
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
    status.textContent = "Lỗi khi xuất file: " + e.message;
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

function patchCellInSheetXml(xml, cellRef, value, isNumber) {
  // capture existing style id (s="NN") if the cell already exists, so we
  // never lose the template's formatting/number-format for that cell.
  const selfClosing = new RegExp(`<c r="${cellRef}"([^>]*)/>`);
  const withContent = new RegExp(`<c r="${cellRef}"([^>]*)>.*?</c>`, "s");

  let styleAttr = null;
  let m = xml.match(selfClosing) || xml.match(withContent);
  if (m) {
    const attrs = m[1];
    const sMatch = attrs.match(/s="(\d+)"/);
    if (sMatch) styleAttr = sMatch[1];
  }

  const newCell = buildCellXml(cellRef, styleAttr, value, isNumber);

  if (xml.match(withContent)) {
    return xml.replace(withContent, newCell);
  }
  if (xml.match(selfClosing)) {
    return xml.replace(selfClosing, newCell);
  }

  // Cell doesn't exist yet in the XML at all (rare - template had a fully
  // empty row/cell not materialized). Insert it into the correct <row>,
  // in column order, or create the row if missing.
  const rowNum = cellRef.match(/\d+/)[0];
  const rowRegexOpen = new RegExp(`(<row r="${rowNum}"[^>]*>)`);
  const rowRegexSelf = new RegExp(`<row r="${rowNum}"([^>]*)/>`);
  if (xml.match(rowRegexOpen)) {
    return xml.replace(rowRegexOpen, `$1${newCell}`);
  }
  if (xml.match(rowRegexSelf)) {
    return xml.replace(rowRegexSelf, (full, attrs) => `<row r="${rowNum}"${attrs}>${newCell}</row>`);
  }
  // Last resort: cannot safely locate the row - skip silently rather than
  // risk corrupting the file. (Should not happen for cells listed in
  // field_map.json, which were themselves read from this same template.)
  console.warn("Could not locate row for", cellRef, "- value not written");
  return xml;
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

  const inputs = document.querySelectorAll("#dataForm [data-cell]");
  let count = 0;
  inputs.forEach(inp => {
    const val = inp.value;
    if (val === "__other__" || val === "") return; // leave template's blank cell untouched
    const cellRef = inp.dataset.cell;
    const isNumber = inp.dataset.type === "number";
    xml = patchCellInSheetXml(xml, cellRef, isNumber ? parseFloat(val) : val, isNumber);
    count++;
  });

  zip.file(sheetPath, xml);

  const blob = zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const outName = `${todayStr()}-${fileKey}-${sheet}.xlsx`.replace(/[^\w\-.]+/g, "_");
  saveBlob(blob, outName);
  console.log(`Patched ${count} cell(s) in ${sheetPath}`);
}


async function exportDocx() {
  const { formKey } = currentSelection;
  const info = docxForms[formKey];
  const resp = await fetch(DOCX_TEMPLATE_DIR + info.template);
  const buf = await resp.arrayBuffer();
  const zip = new PizZip(buf);
  const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => "" });

  const data = {};
  document.querySelectorAll("#dataForm [data-tag]").forEach(inp => {
    const val = inp.value === "__other__" ? "" : inp.value;
    data[inp.dataset.tag] = val || "";
  });
  doc.render(data);

  const out = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const outName = `${todayStr()}-${formKey}.docx`;
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
