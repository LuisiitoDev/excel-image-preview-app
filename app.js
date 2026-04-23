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

function hasCellValue(cell) {
  return cell !== null && cell !== undefined && String(cell).trim() !== "";
}

function getCellText(cell) {
  if (cell === undefined) return "";
  if (cell.w !== undefined) return cell.w;
  if (cell.v !== undefined) return cell.v;
  return "";
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
    wb = XLSX.read(data, { type: "array", cellStyles: true });
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
  const rowMeta = ws["!rows"] || [];
  const colMeta = ws["!cols"] || [];
  const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;

  const showEmpty = (msg) => {
    const notice = document.createElement("p");
    notice.className = "empty-sheet-notice";
    notice.textContent = msg;
    tableWrapper.innerHTML = "";
    tableWrapper.appendChild(notice);
  };

  if (!range) {
    showEmpty(`La hoja "${sheetName}" está vacía.`);
    return;
  }

  // Collect visible row indices (skip rows marked hidden)
  const visibleRowIdxs = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    if (rowMeta[r]?.hidden) continue;
    visibleRowIdxs.push(r);
  }

  // Collect visible column indices (skip columns marked hidden AND empty columns)
  const visibleColIdxs = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    if (colMeta[c]?.hidden) continue;
    const hasValue = visibleRowIdxs.some((r) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return hasCellValue(getCellText(cell));
    });
    if (hasValue) visibleColIdxs.push(c);
  }

  if (!visibleRowIdxs.length || !visibleColIdxs.length) {
    showEmpty(`La hoja "${sheetName}" está vacía.`);
    return;
  }

  // Build data rows using only the visible indices
  const visibleRows = visibleRowIdxs.map((r) =>
    visibleColIdxs.map((c) => getCellText(ws[XLSX.utils.encode_cell({ r, c })]))
  );

  const header = visibleRows[0];
  const body   = visibleRows.slice(1);

  const fragment  = document.createDocumentFragment();
  const table     = document.createElement("table");
  const thead     = table.createTHead();
  const headerRow = thead.insertRow();

  header.forEach((cell, visibleIndex) => {
    const th = document.createElement("th");
    th.textContent = String(hasCellValue(cell) ? cell : `Columna ${visibleIndex + 1}`);
    headerRow.appendChild(th);
  });

  const tbody = table.createTBody();
  body.forEach((row) => {
    const tr = tbody.insertRow();
    row.forEach((cell) => {
      const td = tr.insertCell();
      td.textContent = String(cell ?? "");
    });
  });

  tableWrapper.innerHTML = "";
  fragment.appendChild(table);
  tableWrapper.appendChild(fragment);
}
