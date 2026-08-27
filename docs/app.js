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

const fileInput = document.getElementById("file-input");
const generateBtn = document.getElementById("generate-btn");
const statusEl = document.getElementById("status");
const errorEl = document.getElementById("error");
const summaryEl = document.getElementById("summary");
const summaryBodyEl = document.getElementById("summary-body");
const downloadBtn = document.getElementById("download-btn");
const dropZone = document.getElementById("drop-zone");
const fileNameEl = document.getElementById("selected-file-name");

let selectedFile = null;
let generatedBuffer = null;
let generatedFileName = null;

function setSelectedFile(file) {
  selectedFile = file || null;
  fileNameEl.textContent = selectedFile ? selectedFile.name : "Ningún archivo seleccionado";
  generateBtn.disabled = !selectedFile;
  summaryEl.hidden = true;
  errorEl.hidden = true;
}

fileInput.addEventListener("change", (e) => {
  setSelectedFile(e.target.files && e.target.files[0]);
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
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) setSelectedFile(file);
});

function renderSummary(data) {
  const rows = [
    ["Competencia", data.nombre || "—"],
    ["Perfil(es)", data.perfiles || "—"],
    [
      "Gerencia / Superintendencia / Área",
      [data.gerencia, data.superintendencia, data.area].filter(Boolean).join(" / ") || "—",
    ],
    ["Código", isPlaceholderCodigo(data.codigo) ? "—" : data.codigo],
  ];

  const dl = document.createElement("dl");
  dl.className = "summary-grid";
  for (const [label, value] of rows) {
    const dt = document.createElement("dt");
    dt.textContent = label;
    const dd = document.createElement("dd");
    dd.textContent = value;
    dl.appendChild(dt);
    dl.appendChild(dd);
  }

  const heading = document.createElement("h3");
  heading.className = "activities-heading";
  heading.textContent = `Actividades Clave (${data.activities.length})`;

  const ul = document.createElement("ul");
  ul.className = "activities-list";
  data.activities.forEach((a, i) => {
    const li = document.createElement("li");
    const strong = document.createElement("strong");
    strong.textContent = `${i + 1}. ${a.title}`;
    const span = document.createElement("span");
    span.textContent = ` (${a.criterios.length} criterio${a.criterios.length === 1 ? "" : "s"} de desempeño)`;
    li.appendChild(strong);
    li.appendChild(span);
    ul.appendChild(li);
  });

  summaryBodyEl.replaceChildren(dl, heading, ul);
  summaryEl.hidden = false;
}

async function handleGenerate() {
  if (!selectedFile) return;
  errorEl.hidden = true;
  summaryEl.hidden = true;
  generateBtn.disabled = true;
  statusEl.textContent = "Generando…";
  statusEl.hidden = false;

  try {
    const [docxBuffer, templateResponse] = await Promise.all([
      selectedFile.arrayBuffer(),
      fetch("assets/sot-template.xlsx"),
    ]);
    if (!templateResponse.ok) {
      throw new Error("No se pudo cargar la plantilla del SOT.");
    }
    const templateBuffer = await templateResponse.arrayBuffer();

    const data = await parseCompetenciaDocx(docxBuffer);
    const workbookBuffer = await generateSotWorkbook(data, templateBuffer);

    generatedBuffer = workbookBuffer;
    generatedFileName = buildOutputFileName(data);

    renderSummary(data);
    downloadBtn.textContent = `Descargar ${generatedFileName}`;
    downloadBtn.hidden = false;
  } catch (err) {
    errorEl.textContent = err instanceof Error ? err.message : String(err);
    errorEl.hidden = false;
    downloadBtn.hidden = true;
  } finally {
    statusEl.hidden = true;
    generateBtn.disabled = false;
  }
}

function handleDownload() {
  if (!generatedBuffer || !generatedFileName) return;
  try {
    const blob = new Blob([generatedBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = generatedFileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (err) {
    console.error("download failed", err);
    errorEl.textContent = "No se pudo descargar el archivo: " + (err instanceof Error ? err.message : String(err));
    errorEl.hidden = false;
  }
}

generateBtn.addEventListener("click", handleGenerate);
downloadBtn.addEventListener("click", handleDownload);
