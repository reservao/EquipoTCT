# Equipo TCT — Plataforma

Plataforma del equipo de Talentos y Capacidades del Trabajo (TCT) de Circular HR.

## Módulo SOT — Set de Observación en Terreno

Convierte el Word de una Competencia (UCL) en el Excel del SOT correspondiente. Sube uno o varios documentos `.docx`, opcionalmente el logo de la empresa, y descarga los Excel generados (individualmente o todos juntos en un `.zip`).

Todo el procesamiento ocurre en el navegador — ningún archivo se sube a un servidor.

- **App:** [reservao.github.io/EquipoTCT](https://reservao.github.io/EquipoTCT/)
- **Código:** `docs/` (HTML/CSS/JS estático, sin build) + `docs/vendor/` (jszip, exceljs vendorizados) + `docs/assets/sot-template.xlsx` (plantilla base del SOT)

## Módulo Carga Competencias y Pruebas

Sube uno o varios Word de Competencias (UCL) y/o de Pruebas (TCO) y genera la planilla de carga (`CARGA COMPETENCIA` + `CARGA INSTRUMENTO PRUEBA`, entre otras hojas) lista para importar. Deja en blanco cualquier campo que no venga literalmente en el documento (p. ej. el código de la competencia o de la prueba) en vez de adivinarlo.

- **Competencias:** por ahora solo procesa competencias **FUNCIONALES** — el mismo formato de Word que usa SOT. Las competencias **CONDUCTUALES** quedan pendientes.
- **Pruebas (TCO):** cada pregunta tiene su respuesta correcta marcada en el Word con algún formato distinto al resto de las alternativas (color, negrita, etc.), sin una convención única entre quienes redactan las pruebas. En vez de buscar colores específicos, se compara el formato de cada alternativa contra sus hermanas dentro de la misma pregunta — la que se distingue del resto se marca como correcta. Las preguntas de combinaciones (enunciados I, II, III… seguidos de opciones tipo "Alternativas I y III") también se reconocen: el párrafo en blanco que separa los enunciados de las opciones reales es la señal para saber dónde empiezan las alternativas de verdad. Solo si ninguna alternativa se distingue con claridad esa pregunta queda sin respuesta marcada para completarla manualmente en el Excel. Las preguntas con una imagen incrustada se detectan pero la imagen no se incluye en la planilla (columna `IMAGEN_PREGUNTA` en blanco).

- **App:** [reservao.github.io/EquipoTCT/carga/](https://reservao.github.io/EquipoTCT/carga/)
- **Código:** `docs/carga/` + `docs/competencia-parser.js` (parser de Word de competencias, compartido con SOT) + `docs/prueba-parser.js` (parser de Word de pruebas/TCO) + `docs/carga/assets/planilla-carga-template.xlsx` (plantilla base)

## Pruebas

Los scripts en `scripts/` usan Playwright para probar la app estática end-to-end:

```bash
npm install
npx playwright install chromium   # si no está ya instalado
node scripts/e2e-test.mjs <competencia.docx>[,<otra.docx>,...] <output-dir> [base-url]
node scripts/e2e-logo-test.mjs <competencia.docx> <logo.png> <output-dir> [base-url]
node scripts/e2e-carga-test.mjs <competencia.docx>[,<otra.docx>,...] <output-dir> [base-url] [prueba.docx,...]
```

Por defecto apuntan a `http://localhost:8899/index.html` (o `/carga/index.html`); sirve `docs/` con cualquier servidor estático (ej. `python3 -m http.server 8899 -d docs`) antes de correrlos.
