// Manual regression check for the SOT PPT feature (docs/index.html +
// docs/app.js + docs/sot-pptx-generator.js): serve docs/ locally, then:
//   node scripts/e2e-sot-ppt-test.mjs <competencia.docx> <output-dir> [base-url]
import { chromium } from "playwright";
import path from "path";

const DOCX_PATH = process.argv[2];
const OUT_DIR = process.argv[3];
const BASE_URL = process.argv[4] ?? "http://localhost:8899/index.html";

if (!DOCX_PATH || !OUT_DIR) {
  console.error("Usage: node scripts/e2e-sot-ppt-test.mjs <path-to-competencia.docx> <output-dir> [base-url]");
  process.exit(1);
}

const executablePath =
  process.env.PLAYWRIGHT_CHROMIUM_PATH ?? "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";
const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium" });
const page = await browser.newPage();
page.on("console", (msg) => {
  if (msg.type() === "error") console.log("[console]", msg.type(), msg.text());
});
page.on("pageerror", (err) => console.log("[pageerror]", err.message));

await page.goto(BASE_URL, { waitUntil: "load" });

console.log("--- checking theme panel defaults ---");
console.log("theme-navy value:", await page.locator("#theme-navy").inputValue());

console.log("--- changing a theme color (navy -> red) to test live preview ---");
await page.evaluate(() => {
  const el = document.getElementById("theme-navy");
  el.value = "#ff0000";
  el.dispatchEvent(new Event("input"));
});
const bandBg = await page.locator("#pv-band").evaluate((el) => getComputedStyle(el).backgroundColor);
console.log("preview band background after change:", bandBg);

await page.setInputFiles("#file-input", [DOCX_PATH]);
console.log("selected file text:", await page.locator("#selected-file-name").textContent());

await page.click("#generate-btn");
await page.waitForSelector("#results:not([hidden])", { timeout: 30000 });
console.log("=== RESULTS HEADING ===", await page.locator("#results-heading").textContent());

const rows = await page.locator("#results-list .result-row").all();
console.log("result rows:", rows.length);
for (const row of rows) {
  console.log(" row:", (await row.textContent()).replace(/\s+/g, " ").trim());
}

const errorVisible = await page.locator("#error:not([hidden])").count();
console.log("top-level error visible:", errorVisible > 0);
if (errorVisible > 0) {
  console.log("error text:", await page.locator("#error").textContent());
}

console.log("--- downloading PPT via per-row button ---");
const pptBtn = page.locator("#results-list .result-row button.download-small", { hasText: "Descargar PPT" }).first();
const [pptDownload] = await Promise.all([page.waitForEvent("download", { timeout: 30000 }), pptBtn.click()]);
const pptOutPath = path.join(OUT_DIR, pptDownload.suggestedFilename());
await pptDownload.saveAs(pptOutPath);
console.log("PPT downloaded filename:", pptDownload.suggestedFilename());
console.log("PPT saved to:", pptOutPath);

const pptRowErrorVisible = await page
  .locator("#results-list .result-row .result-error:not([hidden])")
  .count();
console.log("per-row PPT error visible:", pptRowErrorVisible > 0);
if (pptRowErrorVisible > 0) {
  console.log("per-row PPT error text:", await page.locator("#results-list .result-row .result-error").first().textContent());
}

console.log("--- testing config export ---");
const [configDownload] = await Promise.all([
  page.waitForEvent("download", { timeout: 10000 }),
  page.click("#config-download-btn"),
]);
const configOutPath = path.join(OUT_DIR, configDownload.suggestedFilename());
await configDownload.saveAs(configOutPath);
console.log("config downloaded to:", configOutPath);

console.log("--- testing config import (reset color then re-import) ---");
await page.evaluate(() => {
  const el = document.getElementById("theme-navy");
  el.value = "#000000";
  el.dispatchEvent(new Event("input"));
});
await page.setInputFiles("#config-input", [configOutPath]);
await page.waitForTimeout(500);
console.log("theme-navy after import:", await page.locator("#theme-navy").inputValue());
console.log("config name text:", await page.locator("#selected-config-name").textContent());

await browser.close();
