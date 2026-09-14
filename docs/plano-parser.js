"use strict";

/* ---------- "UCL a Plano" row building ---------- */
/* Turns the same Competencia Word docx already read by SOT/Carga into the
   flat "Plano" row format: one row per Actividad Clave, Criterio,
   Conocimiento (básico/técnico) or Herramienta, tagged with the profile and
   competencia name it belongs to. */

// A single blank space, not an empty string -- every row in the reference
// example file ("Ej. Archivo plano.xlsx") uses a literal space in this
// column, never truly empty.
const BLANK_OBSERVACION = " ";

// "Nombre Módulo" / "Tipo de módulo" categorize each Conocimiento/Herramienta
// against an internal training-module catalog that isn't derivable from the
// Word document itself (the reference file's own formulas for these show
// #N/A wherever no match exists) -- left blank rather than guessed.
// "pendiente" is always empty in the reference file too; its purpose isn't
// documented anywhere, so it's left blank as well.
function buildPlanoRows(data) {
  const perfil = data.perfiles || "";
  const competencia = data.nombre || "";
  const rows = [];

  const pushRow = (tipo, contenido) => {
    rows.push([perfil, competencia, tipo, contenido, BLANK_OBSERVACION, "", "", ""]);
  };

  data.activities.forEach((activity, actIdx) => {
    pushRow("Actividad Clave", `${actIdx + 1}. ${activity.title}`);
    activity.criterios.forEach((criterioText, critIdx) => {
      pushRow("Criterios", `${actIdx + 1}.${critIdx + 1}. ${criterioText}`);
    });
  });

  data.conocimientosBasicos.forEach((text) => pushRow("Conocimiento básico", text));
  data.conocimientosTecnicos.forEach((text) => pushRow("Conocimiento técnico", text));
  data.herramientas.forEach((text) => pushRow("Herramientas", text));

  return rows;
}
