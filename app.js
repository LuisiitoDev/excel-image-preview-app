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
    `Formato no soportado: "${file.name}". Por favor sube un archivo .xlsx o una imagen.`
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
  const strong = document.createElement("strong");

  meta.textContent = "";
  strong.textContent = file.name;

  meta.appendChild(document.createTextNode("Archivo: "));
  meta.appendChild(strong);
  meta.appendChild(
    document.createTextNode(` \u00A0·\u00A0 Tamaño: ${mb}\u00A0MB \u00A0·\u00A0 Tipo: ${type}`)
  );
  meta.classList.remove("hidden");
}

/* ── Error ───────────────────────────────────────── */
function showError(message) {
  errorBox.textContent = message;
  errorBox.classList.remove("hidden");
}

/* ── Image preview ───────────────────────────────── */
function previewImage(file) {
  objectUrlCache = URL.createObjectURL(file);
  previewImg.src = objectUrlCache;
  previewImg.alt = `Vista previa de ${file.name}`;
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
    const notice = document.createElement("p");
    notice.className = "empty-sheet-notice";
    notice.textContent = `La hoja "${sheetName}" está vacía.`;
    tableWrapper.innerHTML = "";
    tableWrapper.appendChild(notice);
    return;
  }

  const maxCols = rows.reduce((m, row) => Math.max(m, row.length), 0);
  const header  = rows[0] || [];
  const body    = rows.slice(1);
  const visibleCols = [];

  for (let c = 0; c < maxCols; c++) {
    const columnHasValue = rows.some((row) => {
      const cell = row[c];
      return cell !== null && cell !== undefined && String(cell).trim() !== "";
    });

    if (columnHasValue) visibleCols.push(c);
  }

  if (!visibleCols.length) {
    const notice = document.createElement("p");
    notice.className = "empty-sheet-notice";
    notice.textContent = `La hoja "${sheetName}" no contiene columnas con valores.`;
    tableWrapper.innerHTML = "";
    tableWrapper.appendChild(notice);
    return;
  }

  const fragment = document.createDocumentFragment();
  const table    = document.createElement("table");
  const thead    = table.createTHead();
  const headerRow = thead.insertRow();

  for (const c of visibleCols) {
    const th = document.createElement("th");
    const colName = header[c] !== "" ? header[c] : `Columna ${c + 1}`;
    th.textContent = String(colName);
    headerRow.appendChild(th);
  }

  const tbody = table.createTBody();
  body.forEach((row) => {
    const tr = tbody.insertRow();
    for (const c of visibleCols) {
      const td = tr.insertCell();
      td.textContent = String(row[c] ?? "");
    }
  });

  tableWrapper.innerHTML = "";
  fragment.appendChild(table);
  tableWrapper.appendChild(fragment);
}
