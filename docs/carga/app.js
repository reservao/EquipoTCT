/* global JSZip, ExcelJS, parseCompetenciaDocx, isPlaceholderCodigo */
"use strict";

/* ---------- Carga Competencia row building ---------- */

// Only FUNCIONAL competencias are supported so far — same Word format SOT
// already parses (Actividad Clave / Criterio de Desempeño). CONDUCTUAL
// competencias (behavioral scale, no activities/criterios) use a different
// Word structure this module doesn't have a sample of yet.
const TIPO_COMPETENCIA = "FUNCIONAL";

// Anything not literally present in the Word (e.g. the competencia's own
// código, which every sample document ships as a "Por definir" placeholder)
// is left blank rather than guessed, per explicit instruction.
function buildCompetenciaRows(data) {
  const codigo = isPlaceholderCodigo(data.codigo) ? "" : data.codigo;
  const competencia = data.nombre || "";
  const rows = [];

  data.activities.forEach((activity, actIdx) => {
    const actividadClave = `${actIdx + 1}.\t${activity.title}`;
    activity.criterios.forEach((criterioText, critIdx) => {
      const criterio = `${actIdx + 1}.${critIdx + 1}.\t${criterioText}`;
      // Each criterio becomes two rows in the target sheet: one for the
      // "cumple" indicator and one for "no cumple", matching the reference
      // planilla's own layout exactly.
      rows.push([TIPO_COMPETENCIA, codigo, competencia, "", actividadClave, criterio, "Indicador si", "SI"]);
      rows.push([TIPO_COMPETENCIA, codigo, competencia, "", actividadClave, criterio, "Indicador no", "NO"]);
    });
  });

  return rows;
}

async function generatePlanilla(templateArrayBuffer, allRows) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);

  const ws = workbook.getWorksheet("CARGA COMPETENCIA");
  if (!ws) throw new Error("La plantilla no tiene la hoja 'CARGA COMPETENCIA'.");

  allRows.forEach((row, idx) => {
    const excelRow = ws.getRow(idx + 2);
    row.forEach((value, colIdx) => {
      excelRow.getCell(colIdx + 1).value = value || null;
    });
  });

  return workbook.xlsx.writeBuffer();
}

/* ---------- UI wiring ---------- */

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const TEMPLATE_URL = "assets/planilla-carga-template.xlsx";

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

function renderResults(entries) {
  resultsListEl.innerHTML = "";
  const okCount = entries.filter((r) => r.status === "ok").length;
  resultsHeadingEl.textContent =
    entries.length === 1 ? "Resultado" : `Resultados (${okCount}/${entries.length} procesados)`;

  for (const entry of entries) {
    const li = document.createElement("li");
    li.className = "result-row";
    const badgeClass = entry.status === "ok" ? "badge-ok" : "badge-error";
    const badgeText = entry.status === "ok" ? "OK" : "ERROR";
    const meta =
      entry.status === "ok"
        ? `${entry.data.nombre || "(sin nombre)"} · ${entry.data.activities.length} actividades clave`
        : "";
    li.innerHTML = `
      <div class="result-header">
        <span class="result-name">${entry.file.name}</span>
        <span class="badge ${badgeClass}">${badgeText}</span>
      </div>
      ${meta ? `<div class="result-meta">${meta}</div>` : ""}
      ${entry.status === "error" ? `<div class="result-error">${entry.message}</div>` : ""}
    `;
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
    const templateArrayBuffer = await fetch(TEMPLATE_URL).then((r) => r.arrayBuffer());

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
        entries.push({ file, status: "ok", data });
      } catch (err) {
        entries.push({ file, status: "error", message: err.message || String(err) });
      }
    }

    const allRows = entries.filter((e) => e.status === "ok").flatMap((e) => buildCompetenciaRows(e.data));

    if (allRows.length > 0) {
      outputBuffer = await generatePlanilla(templateArrayBuffer, allRows);
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
  triggerBlobDownload(outputBuffer, "Planilla de carga.xlsx", XLSX_MIME);
});
