// Rebuilds docs/carga/assets/planilla-carga-template.xlsx from a raw
// reference planilla (the .xlsm/.xlsx export with real example data and
// macros), for when that reference file needs to be updated later.
//
//   node scripts/prepare-carga-template.mjs <raw-planilla.xlsm> <out.xlsx>
//
// What this strips and why:
//   - VBA macros: ExcelJS can't preserve them on save regardless, and the
//     ones in the reference file are just manual-entry formatting helpers
//     (uppercasing, phone format), not business logic.
//   - Example data (row 2+) in every sheet except CARGA COMPETENCIA and
//     CARGA INSTRUMENTO PRUEBA: those two get generated fresh by the app;
//     formulas already extended down as a fill-in-later pattern (e.g. an
//     email-from-ID formula) are left alone since they're part of the
//     template's own scaffolding, not sample content.
//   - Excel "Table" objects on every sheet: ExcelJS's reader mis-parses a
//     table's headerRowCount when the source XML omits it (defaults it to 0
//     instead of the OOXML-spec implicit 1), then faithfully writes that
//     wrong value back out. Excel then flags the file as needing repair on
//     open and can drop a sheet's data entirely. The tables are only
//     filter-dropdown/banding convenience, not something the app depends
//     on, so removing them avoids the whole bug class.
import { createRequire } from "module";
import { readFileSync, writeFileSync } from "fs";

// Reuses the same vendored ExcelJS build the static app itself loads in the
// browser, rather than adding exceljs as an npm dependency just for this
// one-off script.
const ExcelJS = createRequire(import.meta.url)("../docs/vendor/exceljs.min.js");

const [, , inputPath, outputPath] = process.argv;
if (!inputPath || !outputPath) {
  console.error("Usage: node scripts/prepare-carga-template.mjs <raw-planilla.xlsm> <out.xlsx>");
  process.exit(1);
}

const FILL_SHEETS = new Set(["CARGA COMPETENCIA", "CARGA INSTRUMENTO PRUEBA"]);

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(readFileSync(inputPath));

workbook.eachSheet((ws) => {
  for (const name of Object.keys(ws.tables || {})) {
    ws.removeTable(name);
  }

  if (FILL_SHEETS.has(ws.name)) {
    for (let r = ws.rowCount; r >= 2; r--) {
      ws.getRow(r).eachCell({ includeEmpty: true }, (cell) => {
        cell.value = null;
      });
    }
  } else {
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      row.eachCell({ includeEmpty: false }, (cell) => {
        const isFormula = typeof cell.value === "object" && cell.value !== null && "formula" in cell.value;
        if (!isFormula) cell.value = null;
      });
    });
  }
});

const outBuffer = await workbook.xlsx.writeBuffer();
writeFileSync(outputPath, outBuffer);
console.log(`Wrote ${outputPath} (${outBuffer.byteLength} bytes)`);
