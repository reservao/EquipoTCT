import ExcelJS from "exceljs";
import path from "path";
import { CompetenciaData, isPlaceholderCodigo } from "./docx-parser";

const TEMPLATE_PATH = path.join(process.cwd(), "templates", "sot-template.xlsx");

// In the template, every "ACT" sheet has its criterios table starting at row 16
// (6 rows, 16-21), with a C:E merge and an I:M merge on each row.
const BASE_CRITERIO_ROW = 16;
const BASE_CRITERIO_COUNT = 6;
const MERGE_COLUMN_PAIRS: [string, string][] = [
  ["C", "E"],
  ["I", "M"],
];

function unmergeRow(ws: ExcelJS.Worksheet, row: number) {
  for (const [a, b] of MERGE_COLUMN_PAIRS) {
    ws.unMergeCells(`${a}${row}:${b}${row}`);
  }
}

function mergeRow(ws: ExcelJS.Worksheet, row: number) {
  for (const [a, b] of MERGE_COLUMN_PAIRS) {
    ws.mergeCells(`${a}${row}:${b}${row}`);
  }
}

// Note: ExcelJS's spliceRows() is a no-op when the deleted range extends
// through the sheet's last row (a library quirk), which is exactly our case
// when shrinking the criterios block since it always sits at the bottom of
// the sheet. So excess rows are cleared and hidden instead of removed.
function setCriterioRowCount(ws: ExcelJS.Worksheet, desiredCount: number) {
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

function fillActivitySheet(ws: ExcelJS.Worksheet, data: CompetenciaData, activityIndex: number) {
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

// Gerencia / Superintendencia / Área often repeat the same value (e.g. a
// transversal competencia has all three set to "Transversal"), so dedupe
// before joining rather than concatenating the same word 2-3 times.
function buildAreaField(data: CompetenciaData): string {
  const seen = new Set<string>();
  const parts: string[] = [];
  for (const raw of [data.gerencia, data.superintendencia, data.area]) {
    const value = raw.trim();
    if (!value) continue;
    const key = value.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    parts.push(value);
  }
  return parts.join(" ").toUpperCase();
}

export async function generateSotWorkbook(data: CompetenciaData): Promise<Buffer> {
  if (data.activities.length === 0) {
    throw new Error(
      "No se encontraron Actividades Clave en el documento. Verifica que el Word tenga la sección 'Actividades Clave y Criterios de Desempeño'."
    );
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_PATH);

  const hojaDatos = workbook.getWorksheet("1. Hoja de datos");
  if (!hojaDatos) throw new Error("La plantilla SOT no tiene la hoja '1. Hoja de datos'.");
  hojaDatos.getCell("D7").value = data.perfiles.toUpperCase();
  hojaDatos.getCell("D9").value = buildAreaField(data);

  const templateAct = workbook.getWorksheet("1. ACT 1");
  if (!templateAct) throw new Error("La plantilla SOT no tiene la hoja '1. ACT 1'.");
  // Snapshot the pristine template model before any sheet gets its values overwritten,
  // so it can still be used as a style source for activities beyond the 3 pre-built sheets.
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
    clone.model = { ...pristineActModel, name };
    // ExcelJS's worksheet model getter emits merges under `merges`, but its
    // setter reads them back under `mergeCells` — an asymmetry in the
    // library itself — so the assignment above silently drops them and they
    // must be re-applied explicitly.
    const merges: string[] = pristineActModel.merges ?? [];
    for (const range of merges) {
      clone.mergeCells(range);
    }
  }

  for (let i = 0; i < activityCount; i++) {
    const ws = workbook.getWorksheet(`1. ACT ${i + 1}`);
    if (!ws) continue;
    fillActivitySheet(ws, data, i);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export function buildOutputFileName(data: CompetenciaData): string {
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
