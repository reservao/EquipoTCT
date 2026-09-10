/* global JSZip, stripAccentsLower, normalizeText */
"use strict";

/* ---------- docx parsing for "pruebas" (TCO / exámenes teóricos) ---------- */
/* Different documents mark the correct alternative differently -- some
   authors color the text (we've seen at least two red shades), others use
   bold -- with no single convention shared across the team. Rather than
   hardcode specific colors, each alternative's run-level formatting
   (color, bold, italic, underline, font, size) is compared against its
   siblings within the same question: whichever alternative's formatting
   differs from the rest is the one marked as correct. If no alternative
   stands out, or more than one does, the question is left unresolved
   (blank) rather than guessed. */

// Word's default run color when none is explicitly set; some authors also
// apply literal black instead of leaving it unset. Neither is a deliberate
// highlight, so both are treated as "no color".
function normalizeColor(color) {
  if (!color) return null;
  const upper = color.toUpperCase();
  return upper === "000000" || upper === "AUTO" ? null : upper;
}

function getRunFormatting(runXml) {
  const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  const rPr = rPrMatch ? rPrMatch[1] : "";
  const boldMatch = rPr.match(/<w:b\b([^>]*)\/?>/);
  const bold = !!boldMatch && !/w:val="(?:0|false|none)"/i.test(boldMatch[1] || "");
  const italicMatch = rPr.match(/<w:i\b([^>]*)\/?>/);
  const italic = !!italicMatch && !/w:val="(?:0|false|none)"/i.test(italicMatch[1] || "");
  const underlineMatch = rPr.match(/<w:u\b([^>]*)\/?>/);
  const underline = !!underlineMatch && !/w:val="none"/i.test(underlineMatch[1] || "");
  const colorMatch = rPr.match(/<w:color[^>]*w:val="([0-9A-Fa-f]{6})"/);
  const color = normalizeColor(colorMatch ? colorMatch[1] : null);
  const fontMatch = rPr.match(/<w:rFonts[^>]*w:ascii="([^"]+)"/);
  const font = fontMatch ? fontMatch[1] : null;
  const sizeMatch = rPr.match(/<w:sz\b[^>]*w:val="(\d+)"/);
  const size = sizeMatch ? sizeMatch[1] : null;
  return { bold, italic, underline, color, font, size };
}

function parseParagraphsWithRuns(documentXml) {
  const paragraphMatches = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  return paragraphMatches.map((p) => {
    const runMatches = p.match(/<w:r\b[\s\S]*?<\/w:r>/g) || [];
    const runs = runMatches.map((r) => {
      const textMatches = r.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
      const text = textMatches.map((t) => t.replace(/<w:t[^>]*>/, "").replace(/<\/w:t>/, "")).join("");
      const hasDrawing = /<w:drawing\b/.test(r) || /<w:pict\b/.test(r);
      return { text, hasDrawing, ...getRunFormatting(r) };
    });
    const text = runs.map((r) => r.text).join("");
    const hasImage = runs.some((r) => r.hasDrawing);
    return { text, hasImage, runs };
  });
}

// A paragraph's "signature" is the formatting of its dominant run (the one
// with the most visible text) -- short embedded artifacts (a stray space
// with different formatting, a trailing period) shouldn't outweigh the
// paragraph's actual content.
function paragraphSignature(paragraph) {
  const withText = paragraph.runs.filter((r) => r.text.trim());
  if (withText.length === 0) return null;
  withText.sort((a, b) => b.text.length - a.text.length);
  const dom = withText[0];
  return JSON.stringify([dom.color, dom.bold, dom.italic, dom.underline, dom.font, dom.size]);
}

function isAllBold(paragraph) {
  const withText = paragraph.runs.filter((r) => r.text.trim());
  return withText.length > 0 && withText.every((r) => r.bold);
}

// Reviewer annotations ("Ok", "OK") occasionally left in the alternatives
// list -- not real answer content, and would otherwise show up as a bogus
// extra RESPUESTA_N column.
function isAnnotationOnly(text) {
  return /^ok$/i.test(text.trim());
}

// Some questions replace plain alternatives with a "which combination is
// correct" structure: a handful of unlabeled reference statements followed
// by answer choices phrased as roman-numeral combinations (e.g. "Alternativa
// I y III", "Solo II", "I, II y IV"). There's no reliable way to tell where
// the reference statements end and the real answer choices begin from
// formatting alone, so these are flagged for manual review instead of
// risking misaligned columns.
function looksLikeComboAnswer(text) {
  const t = text.trim();
  return /^alternativas?\b/i.test(t) || /^(solo\s+)?[ivx]+(\s*[,y]\s+[ivx]+)*$/i.test(t);
}

async function parsePruebaDocx(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentEntry = zip.file("word/document.xml");
  if (!documentEntry) {
    throw new Error("El archivo no parece ser un .docx válido (falta word/document.xml).");
  }
  const documentXml = await documentEntry.async("text");
  const paragraphs = parseParagraphsWithRuns(documentXml);

  // The evaluated competencia's name is the last fully-bold header line
  // right before the "NOMBRE COMPLETO:" signature field -- literally
  // present in every sample document, unlike a test código (never found in
  // any of them, so that field is left blank).
  const nombreCompletoIdx = paragraphs.findIndex((p) => stripAccentsLower(p.text).includes("nombre completo"));
  let nombre = "";
  if (nombreCompletoIdx > 0) {
    for (let i = nombreCompletoIdx - 1; i >= 0; i--) {
      const text = normalizeText(paragraphs[i].text);
      if (!text) continue;
      if (isAllBold(paragraphs[i])) nombre = text;
      break;
    }
  }

  const instrIdx = paragraphs.findIndex((p) => stripAccentsLower(p.text).includes("instrucciones"));
  const searchStart = instrIdx === -1 ? 0 : instrIdx + 1;

  const questionIdxs = [];
  for (let i = searchStart; i < paragraphs.length; i++) {
    if (normalizeText(paragraphs[i].text) && isAllBold(paragraphs[i])) questionIdxs.push(i);
  }

  const questions = [];
  for (let qi = 0; qi < questionIdxs.length; qi++) {
    const qIdx = questionIdxs[qi];
    const nextQIdx = qi + 1 < questionIdxs.length ? questionIdxs[qi + 1] : paragraphs.length;
    const pregunta = normalizeText(paragraphs[qIdx].text);

    const hasImage = paragraphs.slice(qIdx + 1, nextQIdx).some((p) => p.hasImage);

    const altParagraphs = [];
    for (let i = qIdx + 1; i < nextQIdx; i++) {
      const text = normalizeText(paragraphs[i].text);
      if (text && !isAnnotationOnly(text)) altParagraphs.push({ text, sig: paragraphSignature(paragraphs[i]) });
    }

    const comboStyle = altParagraphs.filter((a) => looksLikeComboAnswer(a.text)).length >= 2;

    let alternativas = [];
    let correctaIndex = null;
    let needsReview = false;

    if (altParagraphs.length === 0) {
      needsReview = true;
    } else if (comboStyle) {
      // Only the trailing run of combo-style choices are real alternatives;
      // the rest are unlabeled reference statements this parser can't
      // safely fold into the question text without risking wrong numbering.
      needsReview = true;
    } else if (altParagraphs.length > 6) {
      needsReview = true;
    } else {
      alternativas = altParagraphs.map((a) => a.text);
      const counts = new Map();
      altParagraphs.forEach((a) => counts.set(a.sig, (counts.get(a.sig) || 0) + 1));
      const sortedCounts = [...counts.values()].sort((a, b) => b - a);
      const tie = sortedCounts.length > 1 && sortedCounts[0] === sortedCounts[1];
      let majoritySig = null;
      let majorityCount = -1;
      counts.forEach((count, sig) => {
        if (count > majorityCount) {
          majorityCount = count;
          majoritySig = sig;
        }
      });
      const outlierIdxs = tie ? [] : altParagraphs.reduce((acc, a, idx) => {
        if (a.sig !== majoritySig) acc.push(idx);
        return acc;
      }, []);
      if (outlierIdxs.length === 1) {
        correctaIndex = outlierIdxs[0];
      } else {
        needsReview = true;
      }
    }

    questions.push({ pregunta, alternativas, correctaIndex, hasImage, needsReview });
  }

  return { nombre, questions };
}
