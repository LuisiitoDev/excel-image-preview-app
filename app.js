/* ── Element references ──────────────────────────── */
const fileInput    = document.getElementById("fileInput");
const excelPreview = document.getElementById("excelPreview");
const imagePreview = document.getElementById("imagePreview");
const emptyState   = document.getElementById("emptyState");
const tableWrapper = document.getElementById("tableWrapper");
const sheetSelect  = document.getElementById("sheetSelect");
const previewImg   = document.getElementById("previewImg");
const meta         = document.getElementById("meta");
const errorBox     = document.getElementById("errorBox");
const fileLabelText = document.querySelector(".file-label-text");

let workbookCache = null;
let objectUrlCache = null;

/* ── Main event listener ─────────────────────────── */
fileInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  resetViews();
  updateFileLabel(file.name);
  showMeta(file);

  if (file.type.startsWith("image/")) {
    previewImage(file);
    return;
  }

  const isXlsx =
    file.name.toLowerCase().endsWith(".xlsx") ||
    file.type.includes("spreadsheetml");

  if (isXlsx) {
    await previewExcel(file);
    return;
  }

  showError(
    `Formato no soportado: <strong>${escapeHtml(file.name)}</strong>. ` +
    "Por favor sube un archivo <code>.xlsx</code> o una imagen."
  );
});

/* ── Reset ───────────────────────────────────────── */
function resetViews() {
  excelPreview.classList.add("hidden");
  imagePreview.classList.add("hidden");
  emptyState.classList.add("hidden");
  errorBox.classList.add("hidden");
  meta.classList.add("hidden");
  tableWrapper.innerHTML = "";
  sheetSelect.innerHTML = "";
  errorBox.innerHTML = "";

  if (objectUrlCache) {
    URL.revokeObjectURL(objectUrlCache);
    objectUrlCache = null;
  }

  workbookCache = null;
}

/* ── File label ──────────────────────────────────── */
function updateFileLabel(name) {
  if (fileLabelText) fileLabelText.textContent = name;
}

/* ── Metadata ────────────────────────────────────── */
function showMeta(file) {
  const mb = (file.size / (1024 * 1024)).toFixed(2);
  const type = file.type || "desconocido";
  meta.innerHTML =
    `Archivo: <strong>${escapeHtml(file.name)}</strong> &nbsp;·&nbsp; ` +
    `Tamaño: ${mb}&nbsp;MB &nbsp;·&nbsp; Tipo: ${escapeHtml(type)}`;
  meta.classList.remove("hidden");
}

/* ── Error ───────────────────────────────────────── */
function showError(html) {
  errorBox.innerHTML = html;
  errorBox.classList.remove("hidden");
}

/* ── Image preview ───────────────────────────────── */
function previewImage(file) {
  objectUrlCache = URL.createObjectURL(file);
  previewImg.src = objectUrlCache;
  previewImg.alt = `Vista previa de ${escapeHtml(file.name)}`;
  imagePreview.classList.remove("hidden");
}

/* ── Excel preview ───────────────────────────────── */
async function previewExcel(file) {
  let wb;
  try {
    const data = await file.arrayBuffer();
    wb = XLSX.read(data, { type: "array" });
  } catch {
    showError("No se pudo leer el archivo .xlsx. Verifica que sea un libro de Excel válido.");
    return;
  }

  workbookCache = wb;

  wb.SheetNames.forEach((sheetName, i) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    if (i === 0) option.selected = true;
    sheetSelect.appendChild(option);
  });

  renderSheet(wb, wb.SheetNames[0]);
  excelPreview.classList.remove("hidden");

  sheetSelect.onchange = () => {
    renderSheet(workbookCache, sheetSelect.value);
  };
}

/* ── Sheet renderer ──────────────────────────────── */
function renderSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  if (!rows.length) {
    tableWrapper.innerHTML =
      `<p style="padding:1rem;color:var(--warn);">` +
      `La hoja "${escapeHtml(sheetName)}" está vacía.</p>`;
    return;
  }

  const maxCols = rows.reduce((m, row) => Math.max(m, row.length), 0);
  const header  = rows[0] || [];
  const body    = rows.slice(1);

  const fragment = document.createDocumentFragment();
  const table    = document.createElement("table");
  const thead    = table.createTHead();
  const headerRow = thead.insertRow();

  for (let c = 0; c < maxCols; c++) {
    const th = document.createElement("th");
    const colName = header[c] !== "" ? header[c] : `Columna ${c + 1}`;
    th.textContent = String(colName);
    headerRow.appendChild(th);
  }

  const tbody = table.createTBody();
  body.forEach((row) => {
    const tr = tbody.insertRow();
    for (let c = 0; c < maxCols; c++) {
      const td = tr.insertCell();
      td.textContent = String(row[c] ?? "");
    }
  });

  tableWrapper.innerHTML = "";
  fragment.appendChild(table);
  tableWrapper.appendChild(fragment);
}

/* ── Helpers ─────────────────────────────────────── */
function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
