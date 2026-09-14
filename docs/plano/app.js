/* global ExcelJS, parseCompetenciaDocx, buildPlanoRows */
"use strict";

const HEADERS = [
  "Perfil",
  "Competencia",
  "Tipo de Contenido",
  "Contenido",
  "Observaciones",
  "Nombre Módulo",
  "Tipo de modulo",
  "pendiente",
];

async function generatePlano(allRows) {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Plano");
  const headerRow = ws.addRow(HEADERS);
  headerRow.font = { bold: true };
  allRows.forEach((row) => ws.addRow(row));
  ws.columns.forEach((col) => {
    col.width = 40;
  });
  return workbook.xlsx.writeBuffer();
}

/* ---------- UI wiring ---------- */

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const competenciaInput = document.getElementById("competencia-input");
const generateBtn = document.getElementById("generate-btn");
const statusEl = document.getElementById("status");
const errorEl = document.getElementById("error");
const resultsEl = document.getElementById("results");
const resultsHeadingEl = document.getElementById("results-heading");
const resultsListEl = document.getElementById("results-list");
const downloadBtn = document.getElementById("download-btn");
const dropZone = document.getElementById("competencia-drop-zone");
const fileNameEl = document.getElementById("selected-competencia-name");

let selectedFiles = [];
let outputBuffer = null;

function triggerBlobDownload(bufferOrArray, filename, mime) {
  const blob = new Blob([bufferOrArray], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function setSelectedFiles(fileList) {
  selectedFiles = Array.from(fileList).filter((f) => f.name.toLowerCase().endsWith(".docx"));
  if (selectedFiles.length === 0) {
    fileNameEl.textContent = "Ningún archivo seleccionado";
  } else if (selectedFiles.length === 1) {
    fileNameEl.textContent = selectedFiles[0].name;
  } else {
    fileNameEl.textContent = `${selectedFiles.length} archivos seleccionados`;
  }
  generateBtn.disabled = selectedFiles.length === 0;
  resultsEl.hidden = true;
  errorEl.hidden = true;
  outputBuffer = null;
}

competenciaInput.addEventListener("change", () => setSelectedFiles(competenciaInput.files));

["dragover", "dragleave", "drop"].forEach((evt) => {
  dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.toggle("is-dragover", evt === "dragover");
  });
});
dropZone.addEventListener("drop", (e) => {
  if (e.dataTransfer.files.length) setSelectedFiles(e.dataTransfer.files);
});

function competenciaMeta(data, rows) {
  return `${data.nombre || "(sin nombre)"} · ${data.perfiles || "(sin perfil)"} · ${rows.length} filas`;
}

function renderResults(entries) {
  resultsListEl.innerHTML = "";
  const okCount = entries.filter((r) => r.status === "ok").length;
  resultsHeadingEl.textContent =
    entries.length === 1 ? "Resultado" : `Resultados (${okCount}/${entries.length} procesados)`;

  for (const entry of entries) {
    const li = document.createElement("li");
    li.className = "result-row";

    const header = document.createElement("div");
    header.className = "result-header";

    const nameEl = document.createElement("span");
    nameEl.className = "result-name";
    nameEl.textContent = entry.file.name;
    header.appendChild(nameEl);

    const badge = document.createElement("span");
    badge.className = `badge ${entry.status === "ok" ? "badge-ok" : "badge-error"}`;
    badge.textContent = entry.status === "ok" ? "OK" : "ERROR";
    header.appendChild(badge);

    li.appendChild(header);

    if (entry.status === "ok") {
      const meta = document.createElement("div");
      meta.className = "result-meta";
      meta.textContent = competenciaMeta(entry.data, entry.rows);
      li.appendChild(meta);
    } else {
      const err = document.createElement("div");
      err.className = "result-error";
      err.textContent = entry.message;
      li.appendChild(err);
    }

    resultsListEl.appendChild(li);
  }

  resultsEl.hidden = false;
  downloadBtn.hidden = okCount === 0;
}

generateBtn.addEventListener("click", async () => {
  generateBtn.disabled = true;
  statusEl.hidden = false;
  statusEl.textContent = "Generando…";
  errorEl.hidden = true;
  resultsEl.hidden = true;
  outputBuffer = null;

  try {
    const entries = [];
    for (const file of selectedFiles) {
      try {
        const buffer = await file.arrayBuffer();
        const data = await parseCompetenciaDocx(buffer);
        if (data.activities.length === 0) {
          throw new Error(
            "No se encontraron Actividades Clave en el documento. Verifica que el Word tenga la sección 'Actividades Clave y Criterios de Desempeño'."
          );
        }
        const rows = buildPlanoRows(data);
        entries.push({ file, status: "ok", data, rows });
      } catch (err) {
        entries.push({ file, status: "error", message: err.message || String(err) });
      }
    }

    const allRows = entries.filter((e) => e.status === "ok").flatMap((e) => e.rows);

    if (allRows.length > 0) {
      outputBuffer = await generatePlano(allRows);
    }

    renderResults(entries);
  } catch (err) {
    errorEl.hidden = false;
    errorEl.textContent = err.message || String(err);
  } finally {
    statusEl.hidden = true;
    generateBtn.disabled = false;
  }
});

downloadBtn.addEventListener("click", () => {
  if (!outputBuffer) return;
  triggerBlobDownload(outputBuffer, "Plano.xlsx", XLSX_MIME);
});
