// Rebuilds docs/carga/assets/planilla-carga-template.xlsx from a raw
// reference planilla (the .xlsm/.xlsx export with real example data,
// macros, protection, etc.), for when that reference file needs to be
// updated later.
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
//   - Sheet protection, autoFilters, hidden rows/columns, and cell
//     notes/comments on every sheet: this is a plain data-import template
//     with no reason to lock, filter, hide, or annotate anything.
//   - The background fill on every data row (rows 2+) in every sheet: the
//     reference file highlights some lookup/classification columns with a
//     tinted fill all the way down, which just looks like stray "painted"
//     cells once the example data is gone. Header row (row 1) styling is
//     left alone.
//   - Custom row heights on CARGA INSTRUMENTO PRUEBA (rows 2+): the
//     reference file has wildly inconsistent per-row heights left over
//     from its own example content; reset to the sheet default so the
//     empty template looks uniform.
import { createRequire } from "module";
import { readFileSync, writeFileSync } from "fs";

// Reuses the same vendored builds the static app itself loads in the
// browser, rather than adding npm dependencies just for this one-off script.
const require = createRequire(import.meta.url);
const ExcelJS = require("../docs/vendor/exceljs.min.js");
const JSZip = require("../docs/vendor/jszip.min.js");

// ExcelJS's own note removal (cell.note = undefined) only empties a note's
// text, it doesn't remove the underlying legacy-comment VML shape and
// relationship — the cell still shows an (empty) note indicator in Excel.
// Stripping the comment/VML parts directly from the written zip is the only
// way to guarantee they're actually gone.
async function stripCommentsFromBuffer(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const commentPaths = Object.keys(zip.files).filter((f) => /^xl\/comments\d*\.xml$/.test(f));
  const vmlPaths = Object.keys(zip.files).filter((f) => /^xl\/drawings\/vmlDrawing\d*\.vml$/.test(f));
  if (commentPaths.length === 0 && vmlPaths.length === 0) return buffer;

  for (const path of [...commentPaths, ...vmlPaths]) zip.remove(path);

  const contentTypesPath = "[Content_Types].xml";
  let contentTypes = await zip.file(contentTypesPath).async("text");
  contentTypes = contentTypes.replace(/<Override PartName="\/(?:xl\/comments\d*\.xml)"[^/]*\/>/g, "");
  zip.file(contentTypesPath, contentTypes);

  const sheetRelsPaths = Object.keys(zip.files).filter((f) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(f));
  for (const relsPath of sheetRelsPaths) {
    let rels = await zip.file(relsPath).async("text");
    if (!/comments|vmlDrawing/.test(rels)) continue;
    rels = rels.replace(/<Relationship[^>]*Target="\.\.\/(?:comments\d*\.xml|drawings\/vmlDrawing\d*\.vml)"[^/]*\/>/g, "");
    zip.file(relsPath, rels);

    const sheetPath = relsPath.replace("_rels/", "").replace(".rels", "");
    let sheetXml = await zip.file(sheetPath).async("text");
    sheetXml = sheetXml.replace(/<legacyDrawing[^/]*\/>/g, "");
    zip.file(sheetPath, sheetXml);
  }

  // JSZip defaults to no compression on repack, which would bloat the file
  // several-fold since Excel XML compresses very well.
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

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
  ws.unprotect();
  ws.autoFilter = undefined;

  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    row.hidden = false;
    // includeEmpty:false here on purpose — a cell ExcelJS never
    // materialized can't carry a note or fill to begin with, and forcing
    // every column of every row into existence (some sheets have 10,000+
    // template rows) balloons the file from ~180KB to well over 1MB.
    row.eachCell({ includeEmpty: false }, (cell) => {
      // Only touch .note on cells that actually have one — assigning
      // undefined to a cell that never had a note makes ExcelJS's writer
      // fabricate an empty legacy VML comment shape for it (verbose XML,
      // one shape per cell), ballooning the file instead of cleaning it.
      if (cell.note) cell.note = undefined;
      if (rowNumber > 1 && cell.fill && cell.fill.type) {
        cell.fill = { type: "pattern", pattern: "none" };
      }
    });
    if (ws.name === "CARGA INSTRUMENTO PRUEBA" && rowNumber > 1) {
      row.height = undefined;
    }
  });
  ws.columns?.forEach((col) => {
    col.hidden = false;
  });

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

const rawBuffer = await workbook.xlsx.writeBuffer();
const outBuffer = await stripCommentsFromBuffer(rawBuffer);
writeFileSync(outputPath, outBuffer);
console.log(`Wrote ${outputPath} (${outBuffer.byteLength ?? outBuffer.length} bytes)`);
