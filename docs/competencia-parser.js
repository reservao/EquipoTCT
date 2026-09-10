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
