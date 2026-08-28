// Manual regression check for docs/ (the static, client-side SOT tool):
// serve docs/ locally (e.g. `python3 -m http.server 8899` from docs/), then:
//   node scripts/e2e-test.mjs <path-to-competencia.docx>[,<path2.docx>,...] <output-dir> [base-url]
import { chromium } from "playwright";
import path from "path";

const DOCX_PATHS = (process.argv[2] ?? "").split(",").filter(Boolean);
const OUT_DIR = process.argv[3];
const BASE_URL = process.argv[4] ?? "http://localhost:8899/index.html";

if (DOCX_PATHS.length === 0 || !OUT_DIR) {
  console.error(
    "Usage: node scripts/e2e-test.mjs <path-to-competencia.docx>[,<path2.docx>,...] <output-dir> [base-url]"
  );
  process.exit(1);
}

const executablePath =
  process.env.PLAYWRIGHT_CHROMIUM_PATH ?? "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";
const browser = await chromium.launch({ executablePath });
const page = await browser.newPage();

page.on("console", (msg) => console.log("[console]", msg.type(), msg.text()));
page.on("pageerror", (err) => console.log("[pageerror]", err.message));

await page.goto(BASE_URL);

const fileInput = page.locator("#file-input");
await fileInput.setInputFiles(DOCX_PATHS);

await page.locator("#selected-file-name").waitFor({ state: "visible" });
console.log("selected file text:", await page.locator("#selected-file-name").textContent());

await page.locator("#generate-btn").click();

await page.locator("#results").waitFor({ state: "visible", timeout: 20000 });
console.log("=== RESULTS HEADING ===", await page.locator("#results-heading").textContent());

const rows = page.locator("#results-list .result-row");
const rowCount = await rows.count();
console.log("result rows:", rowCount);
for (let i = 0; i < rowCount; i++) {
  console.log(` row ${i}:`, (await rows.nth(i).innerText()).replace(/\n/g, " | "));
}

const errorVisible = await page.locator("#error").isVisible();
console.log("top-level error visible:", errorVisible);
if (errorVisible) {
  console.log("error text:", await page.locator("#error").textContent());
}

console.log("download-all-btn visible:", await page.locator("#download-all-btn").isVisible());
console.log("download-all-btn text:", await page.locator("#download-all-btn").textContent());

const downloadPromise = page.waitForEvent("download", { timeout: 15000 }).catch((e) => {
  console.log("no download event:", e.message);
  return null;
});
await page.locator("#download-all-btn").click();
const download = await downloadPromise;
if (!download) {
  console.log("bailing, no download captured");
  await browser.close();
  process.exit(0);
}
const suggested = download.suggestedFilename();
console.log("downloaded filename:", suggested);
const outPath = path.join(OUT_DIR, suggested);
await download.saveAs(outPath);
console.log("saved to:", outPath);

await browser.close();
