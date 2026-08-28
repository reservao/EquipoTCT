import JSZip from "jszip";

export interface SotActivity {
  title: string;
  criterios: string[];
}

export interface CompetenciaData {
  codigo: string;
  nombre: string;
  version: string;
  gerencia: string;
  superintendencia: string;
  area: string;
  perfiles: string;
  fecha: string;
  activities: SotActivity[];
}

interface ParsedParagraph {
  text: string;
  style: string | null;
  ilvl: string | null;
}

type StringField = Exclude<keyof CompetenciaData, "activities">;

// Keys must already be stripped of accents/case, since they are matched
// against the output of stripAccentsLower(). Some Competencia Words use
// "Gerencia(s)" / "Subgerencia(s)" / "Área(s)" instead of "Gerencia" /
// "Superintendencia" / "Área" — same fields, different label wording
// depending on who authored the document.
const HEADER_LABELS: Record<string, StringField> = {
  "codigo": "codigo",
  "nombre": "nombre",
  "version": "version",
  "gerencia": "gerencia",
  "gerencia(s)": "gerencia",
  "superintendencia": "superintendencia",
  "subgerencia": "superintendencia",
  "subgerencia(s)": "superintendencia",
  "subgerencias": "superintendencia",
  "subgerencias(s)": "superintendencia",
  "area": "area",
  "area(s)": "area",
  "perfil(es)": "perfiles",
  "perfil": "perfiles",
  "fecha de elaboracion": "fecha",
};

function normalize(text: string): string {
  return text.replace(/ /g, " ").trim();
}

function stripAccentsLower(text: string): string {
  return normalize(text)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "");
}

function parseParagraphs(documentXml: string): ParsedParagraph[] {
  const paragraphMatches = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
  return paragraphMatches.map((p) => {
    const styleMatch = p.match(/<w:pStyle w:val="([^"]+)"/);
    const ilvlMatch = p.match(/<w:ilvl w:val="([^"]+)"/);
    const textMatches = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) ?? [];
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

export async function parseCompetenciaDocx(buffer: Buffer): Promise<CompetenciaData> {
  const zip = await JSZip.loadAsync(buffer);
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) {
    throw new Error("El archivo no parece ser un .docx válido (falta word/document.xml).");
  }
  const documentXml = await documentEntry.async("text");
  const paragraphs = parseParagraphs(documentXml);

  const data: CompetenciaData = {
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

  // 1. Header fields: a label paragraph followed by its value paragraph.
  for (let i = 0; i < paragraphs.length - 1; i++) {
    const labelKey = stripAccentsLower(paragraphs[i].text);
    const field = HEADER_LABELS[labelKey];
    if (!field) continue;
    // find the next non-empty paragraph as the value
    for (let j = i + 1; j < paragraphs.length; j++) {
      const value = normalize(paragraphs[j].text);
      if (value) {
        data[field] = value;
        break;
      }
      if (j - i > 3) break; // give up, too far
    }
  }

  // 2. Actividades Clave y Criterios de Desempeño
  const startIdx = paragraphs.findIndex(
    (p) => p.style === "Ttulo1" && stripAccentsLower(p.text).startsWith("actividades clave")
  );
  if (startIdx !== -1) {
    let endIdx = paragraphs.findIndex(
      (p, idx) => idx > startIdx && p.style === "Ttulo1"
    );
    if (endIdx === -1) endIdx = paragraphs.length;

    let current: SotActivity | null = null;
    for (let i = startIdx + 1; i < endIdx; i++) {
      const p = paragraphs[i];
      const text = normalize(p.text);
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

export function isPlaceholderCodigo(codigo: string): boolean {
  const c = stripAccentsLower(codigo);
  return c === "" || c.includes("por definir");
}
