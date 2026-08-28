/* global JSZip, ExcelJS */
"use strict";

/* ---------- docx parsing (ported from src/lib/docx-parser.ts) ---------- */

const HEADER_LABELS = {
  codigo: "codigo",
  nombre: "nombre",
  version: "version",
  gerencia: "gerencia",
  superintendencia: "superintendencia",
  area: "area",
  "perfil(es)": "perfiles",
  "fecha de elaboracion": "fecha",
};

function normalizeText(text) {
  return text.replace(/ /g, " ").trim();
}

function stripAccentsLower(text) {
  return normalizeText(text)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

function parseParagraphs(documentXml) {
  const paragraphMatches = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  return paragraphMatches.map((p) => {
    const styleMatch = p.match(/<w:pStyle w:val="([^"]+)"/);
    const ilvlMatch = p.match(/<w:ilvl w:val="([^"]+)"/);
    const textMatches = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    const text = textMatches
      .map((t) => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, ""))
      .join("");
    return {
      text,
      style: styleMatch ? styleMatch[1] : null,
      ilvl: ilvlMatch ? ilvlMatch[1] : null,
    };
  });
}

async function parseCompetenciaDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) {
    throw new Error("El archivo no parece ser un .docx válido (falta word/document.xml).");
  }
  const documentXml = await documentEntry.async("text");
  const paragraphs = parseParagraphs(documentXml);

  const data = {
    codigo: "",
    nombre: "",
    version: "",
    gerencia: "",
    superintendencia: "",
    area: "",
    perfiles: "",
    fecha: "",
    activities: [],
  };

  for (let i = 0; i < paragraphs.length - 1; i++) {
    const labelKey = stripAccentsLower(paragraphs[i].text);
    const field = HEADER_LABELS[labelKey];
    if (!field) continue;
    for (let j = i + 1; j < paragraphs.length; j++) {
      const value = normalizeText(paragraphs[j].text);
      if (value) {
        data[field] = value;
        break;
      }
      if (j - i > 3) break;
    }
  }

  const startIdx = paragraphs.findIndex(
    (p) => p.style === "Ttulo1" && stripAccentsLower(p.text).startsWith("actividades clave")
  );
  if (startIdx !== -1) {
    let endIdx = paragraphs.findIndex((p, idx) => idx > startIdx && p.style === "Ttulo1");
    if (endIdx === -1) endIdx = paragraphs.length;

    let current = null;
    for (let i = startIdx + 1; i < endIdx; i++) {
      const p = paragraphs[i];
      const text = normalizeText(p.text);
      if (!text) continue;
      if (p.ilvl === "0") {
        current = { title: text, criterios: [] };
        data.activities.push(current);
      } else if (p.ilvl === "1" && current) {
        current.criterios.push(text);
      }
    }
  }

  return data;
}

function isPlaceholderCodigo(codigo) {
  const c = stripAccentsLower(codigo);
  return c === "" || c.includes("por definir");
}

/* ---------- SOT generation (ported from src/lib/sot-generator.ts) ---------- */

const BASE_CRITERIO_ROW = 16;
const BASE_CRITERIO_COUNT = 6;
const MERGE_COLUMN_PAIRS = [
  ["C", "E"],
  ["I", "M"],
];

function unmergeRow(ws, row) {
  for (const [a, b] of MERGE_COLUMN_PAIRS) {
    ws.unMergeCells(`${a}${row}:${b}${row}`);
  }
}

function mergeRow(ws, row) {
  for (const [a, b] of MERGE_COLUMN_PAIRS) {
    ws.mergeCells(`${a}${row}:${b}${row}`);
  }
}

// Note: ExcelJS's spliceRows() is a no-op when the deleted range extends
// through the sheet's last row (a library quirk), which is exactly our case
// when shrinking the criterios block since it always sits at the bottom of
// the sheet. So excess rows are cleared and hidden instead of removed.
function setCriterioRowCount(ws, desiredCount) {
  const diff = desiredCount - BASE_CRITERIO_COUNT;
  if (diff === 0) return;

  if (diff < 0) {
    const removeFrom = BASE_CRITERIO_ROW + desiredCount;
    const removeCount = -diff;
    for (let r = removeFrom; r < removeFrom + removeCount; r++) {
      unmergeRow(ws, r);
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
      });
      row.height = 0;
      row.hidden = true;
    }
  } else {
    const lastRow = BASE_CRITERIO_ROW + BASE_CRITERIO_COUNT - 1;
    ws.duplicateRow(lastRow, diff, true);
    for (let k = 0; k < diff; k++) {
      mergeRow(ws, lastRow + 1 + k);
    }
  }
}

function fillActivitySheet(ws, data, activityIndex) {
  const activity = data.activities[activityIndex];

  ws.getCell("D5").value = data.perfiles.toUpperCase();
  ws.getCell("D7").value = data.nombre.toUpperCase();
  if (!isPlaceholderCodigo(data.codigo)) {
    ws.getCell("L7").value = data.codigo;
  }
  ws.getCell("D9").value = `${activityIndex + 1}.\t${activity.title}`;

  setCriterioRowCount(ws, activity.criterios.length);

  activity.criterios.forEach((criterio, j) => {
    const row = BASE_CRITERIO_ROW + j;
    ws.getCell(`B${row}`).value = `${activityIndex + 1}.${j + 1}`;
    ws.getCell(`C${row}`).value = criterio;
    ws.getCell(`F${row}`).value = "Encargado Técnico";
  });
}

async function generateSotWorkbook(data, templateArrayBuffer) {
  if (data.activities.length === 0) {
    throw new Error(
      "No se encontraron Actividades Clave en el documento. Verifica que el Word tenga la sección 'Actividades Clave y Criterios de Desempeño'."
    );
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);

  const hojaDatos = workbook.getWorksheet("1. Hoja de datos");
  if (!hojaDatos) throw new Error("La plantilla SOT no tiene la hoja '1. Hoja de datos'.");
  hojaDatos.getCell("D7").value = data.perfiles.toUpperCase();
  hojaDatos.getCell("D9").value = `${data.gerencia} ${data.superintendencia} ${data.area}`
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

  const templateAct = workbook.getWorksheet("1. ACT 1");
  if (!templateAct) throw new Error("La plantilla SOT no tiene la hoja '1. ACT 1'.");
  const pristineActModel = JSON.parse(JSON.stringify(templateAct.model));

  const prebuiltActNames = ["1. ACT 1", "1. ACT 2", "1. ACT 3"];
  const activityCount = data.activities.length;

  for (let i = activityCount; i < prebuiltActNames.length; i++) {
    const ws = workbook.getWorksheet(prebuiltActNames[i]);
    if (ws) workbook.removeWorksheet(ws.id);
  }

  for (let i = prebuiltActNames.length; i < activityCount; i++) {
    const name = `1. ACT ${i + 1}`;
    const clone = workbook.addWorksheet(name);
    clone.model = Object.assign({}, pristineActModel, { name });
    // ExcelJS's worksheet model getter emits merges under `merges`, but its
    // setter reads them back under `mergeCells` — an asymmetry in the
    // library itself — so the assignment above silently drops them and they
    // must be re-applied explicitly.
    const merges = pristineActModel.merges || [];
    for (const range of merges) {
      clone.mergeCells(range);
    }
  }

  for (let i = 0; i < activityCount; i++) {
    const ws = workbook.getWorksheet(`1. ACT ${i + 1}`);
    if (!ws) continue;
    fillActivitySheet(ws, data, i);
  }

  return workbook.xlsx.writeBuffer();
}

function buildOutputFileName(data) {
  const base = data.nombre || "SOT";
  const slug = base
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80);
  return `SOT_${slug || "competencia"}.xlsx`;
}

/* ---------- UI wiring ---------- */

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const fileInput = document.getElementById("file-input");
const generateBtn = document.getElementById("generate-btn");
const statusEl = document.getElementById("status");
const errorEl = document.getElementById("error");
const resultsEl = document.getElementById("results");
const resultsHeadingEl = document.getElementById("results-heading");
const resultsListEl = document.getElementById("results-list");
const downloadAllBtn = document.getElementById("download-all-btn");
const dropZone = document.getElementById("drop-zone");
const fileNameEl = document.getElementById("selected-file-name");

let selectedFiles = [];
let results = [];

function triggerBlobDownload(bufferOrArray, filename, mime) {
  const blob = new Blob([bufferOrArray], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function setSelectedFiles(fileList) {
  selectedFiles = fileList ? Array.from(fileList) : [];
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
}

fileInput.addEventListener("change", (e) => {
  setSelectedFiles(e.target.files);
});

["dragover", "dragenter"].forEach((evt) => {
  dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.add("is-dragover");
  });
});
["dragleave", "dragend", "drop"].forEach((evt) => {
  dropZone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropZone.classList.remove("is-dragover");
  });
});
dropZone.addEventListener("drop", (e) => {
  setSelectedFiles(e.dataTransfer.files);
});

function renderResults() {
  resultsListEl.replaceChildren();
  let successCount = 0;

  for (const r of results) {
    const li = document.createElement("li");
    li.className = "result-row";

    const header = document.createElement("div");
    header.className = "result-header";

    const nameEl = document.createElement("span");
    nameEl.className = "result-name";
    nameEl.textContent = r.file.name;

    const badge = document.createElement("span");
    badge.className = `badge ${r.status === "done" ? "badge-ok" : "badge-error"}`;
    badge.textContent = r.status === "done" ? "Generado" : "Error";

    header.appendChild(nameEl);
    header.appendChild(badge);
    li.appendChild(header);

    if (r.status === "done") {
      successCount++;
      const meta = document.createElement("div");
      meta.className = "result-meta";
      const activityCount = r.data.activities.length;
      meta.textContent = `${r.data.nombre || "—"} · ${activityCount} actividad${activityCount === 1 ? "" : "es"} clave`;
      li.appendChild(meta);

      const dlBtn = document.createElement("button");
      dlBtn.className = "download-small";
      dlBtn.textContent = `Descargar ${r.outputFileName}`;
      dlBtn.addEventListener("click", () => triggerBlobDownload(r.buffer, r.outputFileName, XLSX_MIME));
      li.appendChild(dlBtn);
    } else {
      const errEl = document.createElement("div");
      errEl.className = "result-error";
      errEl.textContent = r.error;
      li.appendChild(errEl);
    }

    resultsListEl.appendChild(li);
  }

  resultsHeadingEl.textContent =
    results.length === 1 ? "Resultado" : `Resultados (${successCount}/${results.length} generados)`;
  downloadAllBtn.hidden = successCount === 0;
  resultsEl.hidden = false;
}

async function handleGenerate() {
  if (selectedFiles.length === 0) return;
  errorEl.hidden = true;
  resultsEl.hidden = true;
  generateBtn.disabled = true;
  results = [];

  try {
    const templateResponse = await fetch("assets/sot-template.xlsx");
    if (!templateResponse.ok) {
      throw new Error("No se pudo cargar la plantilla del SOT.");
    }
    const templateBuffer = await templateResponse.arrayBuffer();

    for (let i = 0; i < selectedFiles.length; i++) {
      const file = selectedFiles[i];
      statusEl.textContent =
        selectedFiles.length > 1 ? `Generando ${i + 1}/${selectedFiles.length}: ${file.name}` : "Generando…";
      statusEl.hidden = false;

      try {
        const docxBuffer = await file.arrayBuffer();
        const data = await parseCompetenciaDocx(docxBuffer);
        // Each generation reads (and internally mutates the in-memory model of)
        // the template, so every file gets its own untouched copy of the bytes.
        const workbookBuffer = await generateSotWorkbook(data, templateBuffer.slice(0));
        const outputFileName = buildOutputFileName(data);
        results.push({ file, status: "done", data, buffer: workbookBuffer, outputFileName });
      } catch (err) {
        results.push({
          file,
          status: "error",
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    renderResults();
  } catch (err) {
    errorEl.textContent = err instanceof Error ? err.message : String(err);
    errorEl.hidden = false;
  } finally {
    statusEl.hidden = true;
    generateBtn.disabled = false;
  }
}

async function handleDownloadAll() {
  const successes = results.filter((r) => r.status === "done");
  if (successes.length === 0) return;

  if (successes.length === 1) {
    const r = successes[0];
    triggerBlobDownload(r.buffer, r.outputFileName, XLSX_MIME);
    return;
  }

  const zip = new JSZip();
  const usedNames = new Set();
  for (const r of successes) {
    let name = r.outputFileName;
    if (usedNames.has(name)) {
      const base = name.replace(/\.xlsx$/, "");
      let n = 2;
      while (usedNames.has(`${base}_${n}.xlsx`)) n++;
      name = `${base}_${n}.xlsx`;
    }
    usedNames.add(name);
    zip.file(name, r.buffer);
  }
  const zipBuffer = await zip.generateAsync({ type: "arraybuffer" });
  triggerBlobDownload(zipBuffer, "SOTs.zip", "application/zip");
}

generateBtn.addEventListener("click", handleGenerate);
downloadAllBtn.addEventListener("click", handleDownloadAll);
