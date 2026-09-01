# Equipo TCT — Plataforma

Plataforma del equipo de Talentos y Capacidades del Trabajo (TCT) de Circular HR.

## Módulo SOT — Set de Observación en Terreno

Convierte el Word de una Competencia (UCL) en el Excel del SOT correspondiente. Sube uno o varios documentos `.docx`, opcionalmente el logo de la empresa, y descarga los Excel generados (individualmente o todos juntos en un `.zip`).

Todo el procesamiento ocurre en el navegador — ningún archivo se sube a un servidor.

- **App:** [reservao.github.io/EquipoTCT](https://reservao.github.io/EquipoTCT/)
- **Código:** `docs/` (HTML/CSS/JS estático, sin build) + `docs/vendor/` (jszip, exceljs vendorizados) + `docs/assets/sot-template.xlsx` (plantilla base del SOT)

## Pruebas

Los scripts en `scripts/` usan Playwright para probar la app estática end-to-end:

```bash
npm install
npx playwright install chromium   # si no está ya instalado
node scripts/e2e-test.mjs <competencia.docx>[,<otra.docx>,...] <output-dir> [base-url]
node scripts/e2e-logo-test.mjs <competencia.docx> <logo.png> <output-dir> [base-url]
```

Por defecto apuntan a `http://localhost:8899/index.html`; sirve `docs/` con cualquier servidor estático (ej. `python3 -m http.server 8899 -d docs`) antes de correrlos.
