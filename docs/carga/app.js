/* global JSZip, ExcelJS, parseCompetenciaDocx, isPlaceholderCodigo, parsePruebaDocx */
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

/* ---------- Carga Instrumento Prueba row building ---------- */

// The documents always use fixed multiple-choice questions with a single
// correct alternative, so these are constant literal values, same as the
// reference planilla's own example rows.
const TIPO_PREGUNTA = "Alternativas";
const TIPO_RESPUESTA = "Alternativa Default";
const MAX_ALTERNATIVAS = 6;

// A row is emitted for every question so the literal fields (nombre,
// pregunta, tipo) don't need to be typed by hand -- but when the correct
// answer couldn't be determined (or a question uses a combination format
// with more than 6 options), the RESPUESTA_* columns are left blank rather
// than guessed, per the same "blank if not literal" rule as competencias.
// needsReview is carried alongside so the row can be flagged visually in
// the generated sheet instead of the user having to hunt for blank cells.
function buildPruebaRows(data) {
  const nombre = data.nombre || "";
  return data.questions.map((q, idx) => {
    const pregunta = `${idx + 1}.\t${q.pregunta}`;
    const values = [
      "", // CODIGO PRUEBA/ENCUESTA -- not present in any sample document
      nombre,
      TIPO_PREGUNTA,
      pregunta,
      "", // IMAGEN_PREGUNTA -- image handling not supported yet
      TIPO_RESPUESTA,
    ];
    for (let i = 0; i < MAX_ALTERNATIVAS; i++) {
      values.push(q.needsReview ? "" : q.alternativas[i] || "");
    }
    values.push(q.needsReview ? "" : q.correctaIndex + 1);
    return { values, needsReview: q.needsReview };
  });
}

// Dark red fill + white text, applied to every cell of a row that needs
// manual review -- makes those rows impossible to miss when skimming the
// sheet, instead of only showing up as blank cells someone has to notice.
//
// ExcelJS shares one style object across every cell that hasn't been
// individually styled yet (all the template's blank data-entry cells point
// to the same object) -- setting `cell.fill = ...` directly mutates that
// shared object, silently painting every other untouched cell that happens
// to share it too. Replacing `cell.style` wholesale (spreading the existing
// style into a new object) gives the cell its own independent style entry
// instead of mutating the shared one.
function markRowForReview(excelRow, columnCount) {
  for (let col = 1; col <= columnCount; col++) {
    const cell = excelRow.getCell(col);
    cell.style = {
      ...cell.style,
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF7A1212" } },
      font: { ...cell.font, color: { argb: "FFFFFFFF" } },
    };
  }
}

async function generatePlanilla(templateArrayBuffer, competenciaRows, pruebaRows) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);

  if (competenciaRows.length > 0) {
    const ws = workbook.getWorksheet("CARGA COMPETENCIA");
    if (!ws) throw new Error("La plantilla no tiene la hoja 'CARGA COMPETENCIA'.");
    competenciaRows.forEach((row, idx) => {
      const excelRow = ws.getRow(idx + 2);
      row.forEach((value, colIdx) => {
        excelRow.getCell(colIdx + 1).value = value || null;
      });
    });
  }

  if (pruebaRows.length > 0) {
    const ws = workbook.getWorksheet("CARGA INSTRUMENTO PRUEBA");
    if (!ws) throw new Error("La plantilla no tiene la hoja 'CARGA INSTRUMENTO PRUEBA'.");
    pruebaRows.forEach(({ values, needsReview }, idx) => {
      const excelRow = ws.getRow(idx + 2);
      values.forEach((value, colIdx) => {
        excelRow.getCell(colIdx + 1).value = value || null;
      });
      if (needsReview) markRowForReview(excelRow, values.length);
    });
  }

  return workbook.xlsx.writeBuffer();
}

/* ---------- UI wiring ---------- */

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const TEMPLATE_URL = "assets/planilla-carga-template.xlsx";

const competenciaInput = document.getElementById("competencia-input");
const pruebaInput = document.getElementById("prueba-input");
const generateBtn = document.getElementById("generate-btn");
const statusEl = document.getElementById("status");
const errorEl = document.getElementById("error");
const resultsEl = document.getElementById("results");
const resultsHeadingEl = document.getElementById("results-heading");
const resultsListEl = document.getElementById("results-list");
const downloadBtn = document.getElementById("download-btn");
const competenciaDropZone = document.getElementById("competencia-drop-zone");
const pruebaDropZone = document.getElementById("prueba-drop-zone");
const competenciaFileNameEl = document.getElementById("selected-competencia-name");
const pruebaFileNameEl = document.getElementById("selected-prueba-name");

let selectedCompetenciaFiles = [];
let selectedPruebaFiles = [];
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

function describeSelection(files) {
  if (files.length === 0) return "Ningún archivo seleccionado";
  if (files.length === 1) return files[0].name;
  return `${files.length} archivos seleccionados`;
}

function setSelectedFiles(fileList, kind) {
  const files = Array.from(fileList).filter((f) => f.name.toLowerCase().endsWith(".docx"));
  if (kind === "competencia") {
    selectedCompetenciaFiles = files;
    competenciaFileNameEl.textContent = describeSelection(files);
  } else {
    selectedPruebaFiles = files;
    pruebaFileNameEl.textContent = describeSelection(files);
  }
  generateBtn.disabled = selectedCompetenciaFiles.length === 0 && selectedPruebaFiles.length === 0;
  resultsEl.hidden = true;
  errorEl.hidden = true;
  outputBuffer = null;
}

competenciaInput.addEventListener("change", () => setSelectedFiles(competenciaInput.files, "competencia"));
pruebaInput.addEventListener("change", () => setSelectedFiles(pruebaInput.files, "prueba"));

[
  [competenciaDropZone, "competencia"],
  [pruebaDropZone, "prueba"],
].forEach(([zone, kind]) => {
  ["dragover", "dragleave", "drop"].forEach((evt) => {
    zone.addEventListener(evt, (e) => {
      e.preventDefault();
      zone.classList.toggle("is-dragover", evt === "dragover");
    });
  });
  zone.addEventListener("drop", (e) => {
    if (e.dataTransfer.files.length) setSelectedFiles(e.dataTransfer.files, kind);
  });
});

function competenciaMeta(data) {
  return `${data.nombre || "(sin nombre)"} · ${data.activities.length} actividades clave`;
}

function pruebaMeta(data) {
  const total = data.questions.length;
  const needsReview = data.questions.filter((q) => q.needsReview).length;
  const withImage = data.questions.filter((q) => q.hasImage).length;
  const parts = [`${data.nombre || "(sin nombre)"}`, `${total} preguntas`];
  if (needsReview > 0) parts.push(`${needsReview} requieren revisión manual`);
  if (withImage > 0) parts.push(`${withImage} con imagen (no incluida)`);
  return parts.join(" · ");
}

// Built with createElement/textContent rather than innerHTML template
// strings -- entry.file.name is the uploaded file's own name, which the
// person providing the file fully controls (e.g. a filename ending in
// ".docx" can still contain "<img src=x onerror=...>" before that). Setting
// it via textContent renders it as inert text; interpolating it into an
// innerHTML string would execute it.
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
      meta.textContent = entry.kind === "competencia" ? competenciaMeta(entry.data) : pruebaMeta(entry.data);
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
    const templateArrayBuffer = await fetch(TEMPLATE_URL).then((r) => r.arrayBuffer());

    const entries = [];
    for (const file of selectedCompetenciaFiles) {
      try {
        const buffer = await file.arrayBuffer();
        const data = await parseCompetenciaDocx(buffer);
        if (data.activities.length === 0) {
          throw new Error(
            "No se encontraron Actividades Clave en el documento. Verifica que el Word tenga la sección 'Actividades Clave y Criterios de Desempeño'."
          );
        }
        entries.push({ file, kind: "competencia", status: "ok", data });
      } catch (err) {
        entries.push({ file, kind: "competencia", status: "error", message: err.message || String(err) });
      }
    }

    for (const file of selectedPruebaFiles) {
      try {
        const buffer = await file.arrayBuffer();
        const data = await parsePruebaDocx(buffer);
        if (data.questions.length === 0) {
          throw new Error("No se encontraron preguntas en el documento. Verifica que el Word tenga el formato de TCO esperado.");
        }
        entries.push({ file, kind: "prueba", status: "ok", data });
      } catch (err) {
        entries.push({ file, kind: "prueba", status: "error", message: err.message || String(err) });
      }
    }

    const competenciaRows = entries
      .filter((e) => e.status === "ok" && e.kind === "competencia")
      .flatMap((e) => buildCompetenciaRows(e.data));
    const pruebaRows = entries
      .filter((e) => e.status === "ok" && e.kind === "prueba")
      .flatMap((e) => buildPruebaRows(e.data));

    if (competenciaRows.length > 0 || pruebaRows.length > 0) {
      outputBuffer = await generatePlanilla(templateArrayBuffer, competenciaRows, pruebaRows);
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
