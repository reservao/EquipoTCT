// Manual regression check for docs/carga/ (the "Carga competencias y
// pruebas" module): serve docs/ locally, then:
//   node scripts/e2e-carga-test.mjs <competencia.docx>[,<otra.docx>,...] <output-dir> [base-url]
import { chromium } from "playwright";
import path from "path";

const DOCX_PATHS = (process.argv[2] ?? "").split(",").filter(Boolean);
const OUT_DIR = process.argv[3];
const BASE_URL = process.argv[4] ?? "http://localhost:8899/carga/index.html";

if (DOCX_PATHS.length === 0 || !OUT_DIR) {
  console.error(
    "Usage: node scripts/e2e-carga-test.mjs <path-to-competencia.docx>[,<path2.docx>,...] <output-dir> [base-url]"
  );
  process.exit(1);
}

const executablePath =
  process.env.PLAYWRIGHT_CHROMIUM_PATH ?? "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";
const browser = await chromium.launch({ executablePath });
const page = await browser.newPage();
page.on("console", (msg) => {
  if (msg.type() === "error") console.log("[console]", msg.type(), msg.text());
});
page.on("pageerror", (err) => console.log("[pageerror]", err.message));

await page.goto(BASE_URL, { waitUntil: "load" });

await page.setInputFiles("#competencia-input", DOCX_PATHS);
console.log("selected file text:", await page.locator("#selected-competencia-name").textContent());

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

const downloadBtnVisible = await page.locator("#download-btn:not([hidden])").count();
console.log("download-btn visible:", downloadBtnVisible > 0);

if (downloadBtnVisible > 0) {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#download-btn")]);
  const outPath = path.join(OUT_DIR, download.suggestedFilename());
  await download.saveAs(outPath);
  console.log("downloaded filename:", download.suggestedFilename());
  console.log("saved to:", outPath);
}

await browser.close();
