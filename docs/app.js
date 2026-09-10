/* global JSZip, ExcelJS, parseCompetenciaDocx, isPlaceholderCodigo */
"use strict";

/* docx parsing (parseCompetenciaDocx, isPlaceholderCodigo, HEADER_LABELS,
   etc.) now lives in competencia-parser.js, shared with the "Carga
   competencias y pruebas" module — see index.html for the script load
   order (competencia-parser.js must load before this file). */

/* ---------- SOT generation (ported from src/lib/sot-generator.ts) ---------- */

const BASE_CRITERIO_ROW = 16;
const MERGE_COLUMN_PAIRS = [
  ["C", "E"],
  ["I", "M"],
];

function unmergeRow(ws, row) {
  for (const [a, b] of MERGE_COLUMN_PAIRS) {
    ws.unMergeCells(`${a}${row}:${b}${row}`);
  }
}

function columnLetterToNumber(letter) {
  return letter.charCodeAt(0) - 64;
}

// ExcelJS shares one style object across every cell that happens to have the
// same formatting — even across different sheets — and mutating a cell's
// `.border` (rather than replacing the whole `.style`) edits that shared
// object in place, silently changing every other cell that looked the same.
// Deep-cloning before storing/assigning avoids aliasing into that shared
// object.
function cloneStyle(value) {
  return JSON.parse(JSON.stringify(value));
}

function mergeRow(ws, row) {
  for (const [a, b] of MERGE_COLUMN_PAIRS) {
    const fromCol = columnLetterToNumber(a);
    const toCol = columnLetterToNumber(b);
    // ExcelJS's mergeCells() overwrites every cell in the range with the
    // anchor cell's style, which would erase the distinct per-column
    // borders these cells carry (e.g. the box's right-edge border) — so
    // capture a clone of each one first (see cloneStyle) and reapply after
    // merging, via a full `.style=` replacement so the restored value gets
    // its own independent style slot instead of re-sharing the anchor's.
    const styles = [];
    for (let col = fromCol; col <= toCol; col++) {
      styles.push(cloneStyle(ws.getRow(row).getCell(col).style));
    }
    // Force-unmerge first (ignoring the error if it wasn't merged) rather
    // than skipping rows duplicateRow() already left merged: that shifted
    // merge state satisfies mergeCells()'s own "already merged" check and
    // the in-memory model, but silently fails to survive writeBuffer() for
    // some rows — merging fresh here is what actually makes it into the file.
    try {
      ws.unMergeCells(`${a}${row}:${b}${row}`);
    } catch (err) {
      // wasn't merged; nothing to undo
    }
    ws.mergeCells(`${a}${row}:${b}${row}`);
    styles.forEach((style, idx) => {
      ws.getRow(row).getCell(fromCol + idx).style = style;
    });
  }
}

// Note: ExcelJS's spliceRows() is a no-op when the deleted range extends
// through the sheet's last row (a library quirk), which is exactly our case
// when shrinking the criterios block since it always sits at the bottom of
// the sheet. So excess rows are cleared and hidden instead of removed.
//
// The template's own row count for this block isn't consistent across
// sheets either — "1. ACT 3" ships with only 5 criterio rows while "1. ACT
// 1"/"1. ACT 2" have 6 — so the current count is read from the sheet itself
// rather than assumed fixed, otherwise a 6th criterio on that sheet lands in
// a brand new, unstyled row.
function setCriterioRowCount(ws, desiredCount) {
  const currentCount = ws.rowCount - BASE_CRITERIO_ROW + 1;
  const diff = desiredCount - currentCount;
  if (diff === 0) return;

  if (diff < 0) {
    const removeFrom = BASE_CRITERIO_ROW + desiredCount;
    const removeCount = -diff;

    // The template only closes the criterio box's bottom border on its
    // final row — middle rows rely on the row below for the dividing line —
    // so that border must move onto the new last visible row before the
    // rest get hidden, otherwise the box is left open at the bottom. Done
    // via a full `.style=` replacement (see cloneStyle) rather than
    // `.border=`, which would edit the shared style object in place and
    // silently repaint every other cell using that same style.
    const lastTemplateRow = BASE_CRITERIO_ROW + currentCount - 1;
    const newLastRow = removeFrom - 1;
    ws.getRow(lastTemplateRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const target = ws.getRow(newLastRow).getCell(colNumber);
      const newStyle = cloneStyle(target.style);
      newStyle.border = cloneStyle(cell.border);
      target.style = newStyle;
    });

    for (let r = removeFrom; r < removeFrom + removeCount; r++) {
      unmergeRow(ws, r);
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
      });
      row.hidden = true;
    }
  } else {
    // Duplicating the first (plain middle-style) row and inserting right
    // after it pushes every row below — including the closing row, with its
    // bottom border — further down intact, instead of repeating that
    // closing border on every newly added row.
    ws.duplicateRow(BASE_CRITERIO_ROW, diff, true);
    const newLastRow = BASE_CRITERIO_ROW + desiredCount - 1;
    for (let r = BASE_CRITERIO_ROW; r <= newLastRow; r++) {
      mergeRow(ws, r);
    }
  }
}

const CRITERIO_TEXT_COLUMNS = ["C", "D", "E"];
const POINTS_PER_LINE = 15; // Calibri 11's natural single-line row height
const PIXELS_PER_WIDTH_UNIT = 7; // Excel column-width-unit → pixel, for Calibri 11 (MDW≈7px)
const AVG_CHAR_WIDTH_PX = 6; // rough average glyph width for Calibri 11

// The criterio cell wraps its text (word wrap is on in the template), but
// Excel doesn't auto-fit row height for merged cells — a well-known Excel
// limitation, not something a saved file can opt out of — so the height has
// to be estimated here from the actual text length and merged column width,
// instead of just keeping whatever fixed height that row happened to have
// in the template (which was sized for a completely different document's
// text).
function estimateCriterioRowHeight(ws, text) {
  const widthUnits = CRITERIO_TEXT_COLUMNS.reduce((sum, col) => sum + (ws.getColumn(col).width || 0), 0);
  const widthPx = widthUnits * PIXELS_PER_WIDTH_UNIT;
  const charsPerLine = Math.max(10, Math.floor(widthPx / AVG_CHAR_WIDTH_PX));
  const lines = Math.max(1, Math.ceil(text.length / charsPerLine));
  return lines * POINTS_PER_LINE + 6;
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
    ws.getRow(row).height = estimateCriterioRowHeight(ws, criterio);
  });
}

// Gerencia / Superintendencia / Área often repeat the same value (e.g. a
// transversal competencia has all three set to "Transversal"), so dedupe
// before joining rather than concatenating the same word 2-3 times.
function dedupeJoin(values) {
  const seen = new Set();
  const parts = [];
  for (const raw of values) {
    const value = (raw || "").trim();
    if (!value) continue;
    const key = value.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    parts.push(value);
  }
  return parts.join(" ").toUpperCase();
}

// Some competencias use Gerencia/Superintendencia/Área; others (a different
// faena's template) use Sector/Proceso(s)/Subproceso(s) instead and never
// fill the first set at all. Fall back to the second set when the first is
// entirely empty, so the ÁREA field isn't left blank either way.
function buildAreaField(data) {
  const primary = dedupeJoin([data.gerencia, data.superintendencia, data.area]);
  if (primary) return primary;
  return dedupeJoin([data.sector, data.proceso, data.subproceso]);
}

// Each sheet in the template carries two logos: Circular HR fixed on the
// top-right (never touched) and the client company's own logo on the
// top-left (this is the one this tool replaces). Anchors below are copied
// from the reference SOT's own drawingN.xml for that left slot. "Hoja de
// datos" keeps its own anchor; every ACT sheet reuses the original
// "1. ACT 1" anchor, since dynamically added ACT sheets (beyond the 3
// pre-built ones) have no anchor of their own to copy.
const LOGO_ANCHOR_HOJA_DATOS = {
  tl: { nativeCol: 1, nativeColOff: 57151, nativeRow: 1, nativeRowOff: 19051 },
  br: { nativeCol: 4, nativeColOff: 57150, nativeRow: 5, nativeRowOff: 7649 },
  editAs: "oneCell",
};
const LOGO_ANCHOR_ACT = {
  tl: { nativeCol: 1, nativeColOff: 38100, nativeRow: 1, nativeRowOff: 19051 },
  br: { nativeCol: 3, nativeColOff: 847725, nativeRow: 3, nativeRowOff: 334689 },
  editAs: "oneCell",
};
// Cloned ACT sheets (4th activity onward) start out blank, so the fixed
// Circular HR logo has to be copied over explicitly too — this is "1. ACT
// 1"'s own anchor for it, reused for the same reason as LOGO_ANCHOR_ACT above.
const CIRCULAR_LOGO_ANCHOR_ACT = {
  tl: { nativeCol: 11, nativeColOff: 173832, nativeRow: 2, nativeRowOff: 19050 },
  br: { nativeCol: 12, nativeColOff: 90488, nativeRow: 3, nativeRowOff: 161925 },
  editAs: "oneCell",
};

// Always clears whatever company logo the template shipped with (so a SOT
// generated without an uploaded logo ends up with that slot empty, not the
// template's default), then adds the new one if provided. Only removes
// images anchored at this exact top-left slot, so the fixed Circular HR
// logo on the right is never touched.
function replaceCompanyLogo(ws, imageId, anchor) {
  if (!ws) return;
  ws._media = ws._media.filter(
    (m) => !(m.range && m.range.tl && m.range.tl.nativeCol === anchor.tl.nativeCol)
  );
  if (imageId !== null) {
    ws.addImage(imageId, anchor);
  }
}

async function generateSotWorkbook(data, templateArrayBuffer, logo) {
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
  hojaDatos.getCell("D9").value = buildAreaField(data);

  const templateAct = workbook.getWorksheet("1. ACT 1");
  if (!templateAct) throw new Error("La plantilla SOT no tiene la hoja '1. ACT 1'.");

  const prebuiltActNames = ["1. ACT 1", "1. ACT 2", "1. ACT 3"];
  const activityCount = data.activities.length;

  for (let i = activityCount; i < prebuiltActNames.length; i++) {
    const ws = workbook.getWorksheet(prebuiltActNames[i]);
    if (ws) workbook.removeWorksheet(ws.id);
  }

  // Extra ACT sheets (beyond the 3 pre-built ones) are cloned by copying
  // each cell's value/style straight from "1. ACT 1", cell by cell. An
  // earlier version cloned through the worksheet's `.model` getter/setter,
  // but that round-trip silently drops some per-cell border styles (a
  // library quirk), leaving cloned sheets with broken-looking table borders
  // — most visibly around the Circular HR logo in the header.
  const templateMerges = templateAct.model.merges || [];

  // The clone loop below copies cells and merges but not images, so the
  // fixed Circular HR logo (unlike the company logo, never re-uploaded per
  // generation) needs its bytes captured once here and re-added per clone.
  const circularLogoMedia = templateAct._media.find(
    (m) => m.range.tl.nativeCol !== LOGO_ANCHOR_ACT.tl.nativeCol
  );
  const circularLogoImage = circularLogoMedia ? workbook.getImage(circularLogoMedia.imageId) : null;
  const circularLogoImageId = circularLogoImage
    ? workbook.addImage({ buffer: circularLogoImage.buffer, extension: circularLogoImage.extension })
    : null;

  for (let i = prebuiltActNames.length; i < activityCount; i++) {
    const clone = workbook.addWorksheet(`1. ACT ${i + 1}`);
    // A brand new worksheet defaults to portrait/100%-scale/no-print-area
    // instead of "1. ACT 1"'s landscape/65%-scale page setup and its 80%
    // page-break-preview view, which is what made cloned sheets look
    // stretched/mis-scaled compared to the pre-built ones.
    clone.pageSetup = cloneStyle(templateAct.pageSetup);
    clone.views = cloneStyle(templateAct.views);
    templateAct.columns.forEach((col, idx) => {
      if (col.width != null) clone.getColumn(idx + 1).width = col.width;
    });
    // Merging first matters: ExcelJS's mergeCells() overwrites every cell in
    // the range with the anchor cell's style, so merging after copying
    // styles would erase the per-cell border differences the template
    // relies on for its box outlines.
    for (const range of templateMerges) {
      clone.mergeCells(range);
    }
    templateAct.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = clone.getRow(rowNumber);
      newRow.height = row.height;
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        newRow.getCell(colNumber).style = cell.style;
      });
    });
    if (circularLogoImageId !== null) {
      clone.addImage(circularLogoImageId, CIRCULAR_LOGO_ANCHOR_ACT);
    }
  }

  for (let i = 0; i < activityCount; i++) {
    const ws = workbook.getWorksheet(`1. ACT ${i + 1}`);
    if (!ws) continue;
    fillActivitySheet(ws, data, i);
  }

  const imageId = logo
    ? workbook.addImage({ buffer: logo.buffer, extension: logo.extension })
    : null;
  replaceCompanyLogo(hojaDatos, imageId, LOGO_ANCHOR_HOJA_DATOS);
  for (let i = 0; i < activityCount; i++) {
    replaceCompanyLogo(workbook.getWorksheet(`1. ACT ${i + 1}`), imageId, LOGO_ANCHOR_ACT);
  }

  // The template's own bookViews carries a stale activeTab left over from
  // whoever last edited it in Excel, so every generated SOT opened on
  // whichever sheet they happened to have selected. Force it back to the
  // first sheet so the file always opens on "1. Hoja de datos".
  workbook.views = [{ activeTab: 0, firstSheet: 0 }];

  return workbook.xlsx.writeBuffer();
}

function buildOutputFileName(data) {
  const base = data.nombre || "";
  const clean = base
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
  return `SOT ${clean || "competencia"}.xlsx`;
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
const logoInput = document.getElementById("logo-input");
const logoNameEl = document.getElementById("selected-logo-name");

let selectedFiles = [];
let selectedLogoFile = null;
let results = [];

const LOGO_EXTENSIONS = { "image/png": "png", "image/jpeg": "jpeg" };

logoInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  selectedLogoFile = file || null;
  logoNameEl.textContent = selectedLogoFile
    ? selectedLogoFile.name
    : "Sin logo — quedará vacío junto al de Circular HR";
});

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

    let logo = null;
    if (selectedLogoFile) {
      const extension = LOGO_EXTENSIONS[selectedLogoFile.type];
      if (!extension) {
        throw new Error("El logo debe ser un archivo PNG o JPG.");
      }
      logo = { buffer: await selectedLogoFile.arrayBuffer(), extension };
    }

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
        // The logo buffer, however, is only ever read (added once per workbook
        // via workbook.addImage), so it's safe to share across iterations.
        const workbookBuffer = await generateSotWorkbook(data, templateBuffer.slice(0), logo);
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
