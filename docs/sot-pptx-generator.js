/* global PptxGenJS, isPlaceholderCodigo */
"use strict";

/* Generates the PPT version of the SOT, mirroring the same visual design as
   the Excel SOT (generated in app.js) from the same parsed Competencia data
   (see competencia-parser.js). Ported from a Node-only prototype that was
   reverse-engineered against real reference decks (colors, bar layout,
   table structure, pagination behavior) — see git history for that process.
   Colors are theme-parameterized (SOT_PPTX_DEFAULT_THEME) and the company
   logo is passed in as a data URI, both driven by the customization panel
   in index.html/app.js. */

const SOT_PPTX_DEFAULT_THEME = {
  navy: "0C3F6E",
  teal: "00B2AB",
  headerBlueDark: "00649E",
  headerBlueLight: "388ABC",
  headerNavySub: "091F83",
  tan: "C3A68E",
};
const SOT_PPTX_WHITE = "FFFFFF";
const SOT_PPTX_BLACK = "000000";
const SOT_PPTX_BORDER = "BFBFBF";

const SOT_PPTX_PAGE_W = 10;
const SOT_PPTX_PAGE_H = 7.5;
const SOT_PPTX_MARGIN_X = 0.157;
const SOT_PPTX_TABLE_X = 0.15;
const SOT_PPTX_TABLE_TOP = 1.61;
const SOT_PPTX_BARS_RIGHT_X = 7.113;
const SOT_PPTX_BARS_W = 6.93;
const SOT_PPTX_BARS_RIGHT_W = 2.807;
// Widened from the reference's own proportions (which reserved ~2in on the
// right for the Conocimientos/Herramientas panel) now that panel lives on
// its own slide -- the freed width goes mostly to CRITERIOS, the column
// whose wrapped line count drives how many criterios fit per slide.
const SOT_PPTX_COL_WIDTHS = [1.1, 5.3, 0.4, 0.4, 2.5];
const SOT_PPTX_TABLE_W = SOT_PPTX_COL_WIDTHS.reduce((a, b) => a + b, 0);

const SOT_PPTX_FONT = "Calibri";

// Rough per-criterio row-height estimate, mirroring the same wrapped-text
// idea as the Excel generator's estimateCriterioRowHeight -- used only to
// decide how many whole activities fit on one slide, not to size actual
// pptx rows (pptxgenjs/PowerPoint auto-sizes those from the text itself).
// Figures matched empirically against the reference deck's own row heights
// (near-zero top/bottom cell padding, ~10pt line height).
const SOT_PPTX_CHARS_PER_INCH_AT_9PT = 16.5;
const SOT_PPTX_LINE_HEIGHT_IN = 0.18;
const SOT_PPTX_MIN_ROW_HEIGHT_IN = 0.3;

function sotPptxEstimateRowHeight(text) {
  const charsPerLine = Math.max(10, Math.floor((SOT_PPTX_COL_WIDTHS[1] - 0.16) * SOT_PPTX_CHARS_PER_INCH_AT_9PT));
  const lines = Math.max(1, Math.ceil(text.length / charsPerLine));
  return Math.max(SOT_PPTX_MIN_ROW_HEIGHT_IN, lines * SOT_PPTX_LINE_HEIGHT_IN + 0.02);
}

function sotPptxEstimateActivityHeight(activity) {
  return activity.criterios.reduce((sum, c) => sum + sotPptxEstimateRowHeight(c), 0);
}

// Packs whole activities onto slides -- never splits one activity's own
// criterios across two slides (confirmed against both reference decks: one
// slide held 2 shorter activities together, another held a single long one
// alone, but no activity was ever cut mid-way). If a single activity alone
// still doesn't fit the estimate, it still gets its own slide (nothing else
// to do with it).
function sotPptxPaginateActivities(activities, capacity) {
  const slides = [];
  let current = [];
  let currentHeight = 0;
  for (const activity of activities) {
    const h = sotPptxEstimateActivityHeight(activity);
    if (current.length > 0 && currentHeight + h > capacity) {
      slides.push(current);
      current = [];
      currentHeight = 0;
    }
    current.push(activity);
    currentHeight += h;
  }
  if (current.length > 0) slides.push(current);
  return slides;
}

function sotPptxAddHeaderBars(slide, data, logoDataUri, theme) {
  // Top navy band + logo box (logo box is the same navy, just drawn
  // separately so a real logo image can sit on top of it).
  slide.addShape("rect", {
    x: 0, y: 0.08, w: SOT_PPTX_PAGE_W, h: 0.64, fill: { color: theme.navy }, line: { type: "none" },
  });
  if (logoDataUri) {
    slide.addImage({
      data: logoDataUri, x: SOT_PPTX_MARGIN_X + 0.05, y: 0.13, w: 1.1, h: 0.54,
      sizing: { type: "contain", w: 1.1, h: 0.54 },
    });
  }
  slide.addText(
    [
      { text: "SOT", options: { bold: true, breakLine: true } },
      { text: data.nombre.toUpperCase(), options: {} },
    ],
    { x: 1.5, y: 0.1, w: 8.3, h: 0.55, fontFace: SOT_PPTX_FONT, fontSize: 13, color: SOT_PPTX_WHITE, valign: "top", margin: 0 }
  );

  // Competencia name + Perfil teal bars (left), Código/Vigencia teal bars (right).
  slide.addShape("roundRect", {
    x: SOT_PPTX_MARGIN_X, y: 0.771, w: SOT_PPTX_BARS_W, h: 0.364, rectRadius: 0.04,
    fill: { color: theme.teal }, line: { type: "none" },
  });
  slide.addText(data.nombre.toUpperCase(), {
    x: SOT_PPTX_MARGIN_X + 0.08, y: 0.771, w: SOT_PPTX_BARS_W - 0.16, h: 0.364, fontFace: SOT_PPTX_FONT, fontSize: 13, bold: true,
    color: SOT_PPTX_WHITE, valign: "middle", margin: 0,
  });
  slide.addShape("roundRect", {
    x: SOT_PPTX_MARGIN_X, y: 1.173, w: SOT_PPTX_BARS_W, h: 0.364, rectRadius: 0.04,
    fill: { color: theme.teal }, line: { type: "none" },
  });
  slide.addText(`PERFIL: ${(data.perfiles || "").toUpperCase()}`, {
    x: SOT_PPTX_MARGIN_X + 0.08, y: 1.173, w: SOT_PPTX_BARS_W - 0.16, h: 0.364, fontFace: SOT_PPTX_FONT, fontSize: 15, bold: true,
    color: SOT_PPTX_WHITE, valign: "middle", margin: 0,
  });

  slide.addShape("roundRect", {
    x: SOT_PPTX_BARS_RIGHT_X, y: 0.771, w: SOT_PPTX_BARS_RIGHT_W, h: 0.364, rectRadius: 0.04,
    fill: { color: theme.teal }, line: { type: "none" },
  });
  slide.addText(`CÓDIGO: ${isPlaceholderCodigo(data.codigo) ? "" : data.codigo || ""}`, {
    x: SOT_PPTX_BARS_RIGHT_X + 0.08, y: 0.771, w: SOT_PPTX_BARS_RIGHT_W - 0.16, h: 0.364, fontFace: SOT_PPTX_FONT, fontSize: 12, bold: true,
    color: SOT_PPTX_WHITE, valign: "middle", margin: 0,
  });
  slide.addShape("roundRect", {
    x: SOT_PPTX_BARS_RIGHT_X, y: 1.173, w: SOT_PPTX_BARS_RIGHT_W, h: 0.364, rectRadius: 0.04,
    fill: { color: theme.teal }, line: { type: "none" },
  });
  slide.addText("VIGENCIA: ", {
    x: SOT_PPTX_BARS_RIGHT_X + 0.08, y: 1.173, w: SOT_PPTX_BARS_RIGHT_W - 0.16, h: 0.364, fontFace: SOT_PPTX_FONT, fontSize: 12, bold: true,
    color: SOT_PPTX_WHITE, valign: "middle", margin: 0,
  });
}

function sotPptxBuildTableRows(activities, theme) {
  const headerFont = { fontFace: SOT_PPTX_FONT, fontSize: 10, bold: true, color: SOT_PPTX_WHITE, valign: "middle", align: "center", margin: [2, 2, 2, 2] };
  const subHeaderFont = { ...headerFont, fontSize: 9 };
  const rows = [];
  rows.push([
    { text: "ACTIVIDAD\nCLAVE", options: { ...headerFont, fill: { color: theme.headerBlueDark }, rowspan: 2 } },
    { text: "CRITERIOS DE DESEMPEÑO", options: { ...headerFont, fill: { color: theme.headerBlueLight }, rowspan: 2 } },
    { text: "CUMPLE", options: { ...headerFont, fill: { color: theme.headerBlueDark }, colspan: 2 } },
    { text: "OBSERVACIONES", options: { ...headerFont, fill: { color: theme.headerBlueLight }, rowspan: 2 } },
  ]);
  rows.push([
    { text: "SÍ", options: { ...subHeaderFont, fill: { color: theme.headerNavySub } } },
    { text: "NO", options: { ...subHeaderFont, fill: { color: theme.headerNavySub } } },
  ]);

  const bodyFont = { fontFace: SOT_PPTX_FONT, fontSize: 9, color: SOT_PPTX_BLACK, valign: "middle", margin: [0, 0.08, 0, 0.08] };
  activities.forEach((activity) => {
    activity.criterios.forEach((criterio, critIdx) => {
      const row = [];
      if (critIdx === 0) {
        row.push({
          text: `${activity.number}.${activity.title}`,
          options: { ...bodyFont, align: "left", rowspan: activity.criterios.length },
        });
      }
      row.push({ text: `${activity.number}.${critIdx + 1}.${criterio}`, options: { ...bodyFont, align: "left" } });
      row.push({ text: "", options: { ...bodyFont } });
      row.push({ text: "", options: { ...bodyFont } });
      row.push({ text: "", options: { ...bodyFont } });
      rows.push(row);
    });
  });
  return rows;
}

function sotPptxAddActivitySlide(pres, data, activityGroup, logoDataUri, theme) {
  const slide = pres.addSlide();
  sotPptxAddHeaderBars(slide, data, logoDataUri, theme);
  slide.addTable(sotPptxBuildTableRows(activityGroup, theme), {
    x: SOT_PPTX_TABLE_X,
    y: SOT_PPTX_TABLE_TOP,
    w: SOT_PPTX_TABLE_W,
    colW: SOT_PPTX_COL_WIDTHS,
    border: { type: "solid", color: SOT_PPTX_BORDER, pt: 0.5 },
    autoPage: false,
  });
}

// Conocimientos / Equipos y Herramientas live on their own slide, separate
// from the activities table -- a long list here doesn't risk cramping or
// corrupting the activity rows, and gets the full slide width (two columns)
// instead of a narrow side panel.
function sotPptxAddKnowledgeSlides(pres, data, logoDataUri, theme) {
  const conocimientos = [...data.conocimientosBasicos, ...data.conocimientosTecnicos];
  const herramientas = data.herramientas;
  if (conocimientos.length === 0 && herramientas.length === 0) return;

  const colW = (SOT_PPTX_TABLE_W - 0.3) / 2;
  const colGap = 0.3;
  const top = SOT_PPTX_TABLE_TOP;
  const bottom = SOT_PPTX_PAGE_H - 0.2;
  const available = bottom - top;

  const estimateLines = (items, fontSize, widthIn) => {
    const usable = widthIn - 0.2;
    const charsPerLine = Math.max(8, Math.floor(usable * (16 - fontSize * 0.4)));
    return items.reduce((sum, t) => sum + Math.max(1, Math.ceil((t.length + 2) / charsPerLine)), 0);
  };
  const lineHeightIn = (fs) => (fs / 72) * 1.3;
  const headerHeightIn = 0.32;

  let fontSize = 11;
  let conLines = estimateLines(conocimientos, fontSize, colW);
  let herLines = estimateLines(herramientas, fontSize, colW);
  while (
    fontSize > 7 &&
    Math.max(conLines, herLines) * lineHeightIn(fontSize) + headerHeightIn > available
  ) {
    fontSize -= 0.5;
    conLines = estimateLines(conocimientos, fontSize, colW);
    herLines = estimateLines(herramientas, fontSize, colW);
  }

  const slide = pres.addSlide();
  sotPptxAddHeaderBars(slide, data, logoDataUri, theme);
  const headerFont = { fontFace: SOT_PPTX_FONT, fontSize: 12, bold: true, color: SOT_PPTX_BLACK };
  const bodyFont = { fontFace: SOT_PPTX_FONT, fontSize, color: SOT_PPTX_BLACK };

  slide.addTable(
    [
      [{ text: "CONOCIMIENTOS", options: { ...headerFont, fill: { color: theme.tan }, align: "left" } }],
      [
        {
          text: conocimientos.map((c) => ({ text: c, options: { bullet: true, breakLine: true } })),
          options: { ...bodyFont, fill: { color: theme.tan }, valign: "top" },
        },
      ],
    ],
    { x: SOT_PPTX_TABLE_X, y: top, w: colW, colW: [colW], autoPage: false, border: { type: "none" } }
  );
  slide.addTable(
    [
      [{ text: "EQUIPOS Y HERRAMIENTAS", options: { ...headerFont, fill: { color: theme.tan }, align: "left" } }],
      [
        {
          text: herramientas.map((h) => ({ text: h, options: { bullet: true, breakLine: true } })),
          options: { ...bodyFont, fill: { color: theme.tan }, valign: "top" },
        },
      ],
    ],
    { x: SOT_PPTX_TABLE_X + colW + colGap, y: top, w: colW, colW: [colW], autoPage: false, border: { type: "none" } }
  );
}

async function generateSotPptx(data, { logoDataUri = null, theme = {} } = {}) {
  if (data.activities.length === 0) {
    throw new Error(
      "No se encontraron Actividades Clave en el documento. Verifica que el Word tenga la sección 'Actividades Clave y Criterios de Desempeño'."
    );
  }

  const mergedTheme = { ...SOT_PPTX_DEFAULT_THEME, ...theme };
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "SOT", width: SOT_PPTX_PAGE_W, height: SOT_PPTX_PAGE_H });
  pres.layout = "SOT";

  const availableH = SOT_PPTX_PAGE_H - SOT_PPTX_TABLE_TOP - 0.1;
  const numberedActivities = data.activities.map((activity, i) => ({ ...activity, number: i + 1 }));
  const slidesOfActivities = sotPptxPaginateActivities(numberedActivities, availableH);

  slidesOfActivities.forEach((activityGroup) => {
    sotPptxAddActivitySlide(pres, data, activityGroup, logoDataUri, mergedTheme);
  });
  sotPptxAddKnowledgeSlides(pres, data, logoDataUri, mergedTheme);

  return pres.write({ outputType: "arraybuffer" });
}
