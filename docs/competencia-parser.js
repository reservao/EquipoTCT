/* global JSZip */
"use strict";

/* ---------- docx parsing (ported from src/lib/docx-parser.ts) ---------- */
/* Shared between the SOT module and the "Carga competencias y pruebas"
   module — both read the same Competencia Word format. */

// Some Competencia Words use "Gerencia(s)" / "Subgerencia(s)" / "Área(s)"
// instead of "Gerencia" / "Superintendencia" / "Área" — same fields,
// different label wording depending on who authored the document. Others
// (a different faena's template) skip that taxonomy entirely and use
// "Sector" / "Proceso(s)" / "Subproceso(s)" instead — see buildAreaField.
const HEADER_LABELS = {
  codigo: "codigo",
  nombre: "nombre",
  version: "version",
  gerencia: "gerencia",
  "gerencia(s)": "gerencia",
  superintendencia: "superintendencia",
  subgerencia: "superintendencia",
  "subgerencia(s)": "superintendencia",
  subgerencias: "superintendencia",
  "subgerencias(s)": "superintendencia",
  area: "area",
  "area(s)": "area",
  sector: "sector",
  "proceso(s)": "proceso",
  proceso: "proceso",
  "subproceso(s)": "subproceso",
  "subproceso (s)": "subproceso",
  subproceso: "subproceso",
  "perfil(es)": "perfiles",
  perfil: "perfiles",
  nivel: "perfiles",
  "fecha de elaboracion": "fecha",
};

function normalizeText(text) {
  return text.replace(/ /g, " ").trim();
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

// "Conocimientos" is rendered as a 2-column table (not a bulleted list like
// Actividades/Herramientas): the first column merges "Básicos"/"Técnicos"
// down across several rows via vMerge, the second column holds one
// knowledge item's text per row. Table cells are still just <w:p> elements
// under the hood, so this walks the raw <w:tbl> markup directly rather than
// the flat paragraph list used for everything else.
function parseConocimientosTable(documentXml) {
  const result = { basicos: [], tecnicos: [] };
  const tables = documentXml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) || [];

  const cellText = (cellXml) => {
    const textMatches = cellXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
    return normalizeText(textMatches.map((t) => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join(""));
  };

  for (const table of tables) {
    const rows = table.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || [];
    if (rows.length < 2) continue;
    const headerCells = rows[0].match(/<w:tc>[\s\S]*?<\/w:tc>/g) || [];
    if (headerCells.length < 2) continue;
    const header0 = stripAccentsLower(cellText(headerCells[0]));
    const header1 = stripAccentsLower(cellText(headerCells[1]));
    if (!header0.includes("tipo de conocimiento") || !header1.includes("ambito")) continue;

    let currentCategory = null;
    for (let i = 1; i < rows.length; i++) {
      const cells = rows[i].match(/<w:tc>[\s\S]*?<\/w:tc>/g) || [];
      if (cells.length < 2) continue;
      const categoryText = cellText(cells[0]);
      if (categoryText) currentCategory = stripAccentsLower(categoryText);
      const item = cellText(cells[1]);
      if (!item) continue;
      if (currentCategory && currentCategory.startsWith("basic")) {
        result.basicos.push(item);
      } else if (currentCategory && currentCategory.startsWith("tecnic")) {
        result.tecnicos.push(item);
      }
    }
    break;
  }

  return result;
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
    sector: "",
    proceso: "",
    subproceso: "",
    perfiles: "",
    fecha: "",
    activities: [],
    conocimientosBasicos: [],
    conocimientosTecnicos: [],
    herramientas: [],
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

  const conocimientos = parseConocimientosTable(documentXml);
  data.conocimientosBasicos = conocimientos.basicos;
  data.conocimientosTecnicos = conocimientos.tecnicos;

  const herramientasIdx = paragraphs.findIndex(
    (p) => p.style === "Ttulo1" && stripAccentsLower(p.text).startsWith("herramientas")
  );
  if (herramientasIdx !== -1) {
    let endIdx = paragraphs.findIndex((p, idx) => idx > herramientasIdx && p.style === "Ttulo1");
    if (endIdx === -1) endIdx = paragraphs.length;
    for (let i = herramientasIdx + 1; i < endIdx; i++) {
      const text = normalizeText(paragraphs[i].text);
      if (text) data.herramientas.push(text);
    }
  }

  return data;
}

function isPlaceholderCodigo(codigo) {
  const c = stripAccentsLower(codigo);
  return c === "" || c.includes("por definir");
}
