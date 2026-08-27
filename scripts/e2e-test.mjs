// Manual regression check for docs/ (the static, client-side SOT tool):
// serve docs/ locally (e.g. `python3 -m http.server 8899` from docs/), then:
//   node scripts/e2e-test.mjs <path-to-competencia.docx> <output-dir> [base-url]
import { chromium } from "playwright";
import path from "path";

const DOCX_PATH = process.argv[2];
const OUT_DIR = process.argv[3];
const BASE_URL = process.argv[4] ?? "http://localhost:8899/index.html";

if (!DOCX_PATH || !OUT_DIR) {
  console.error("Usage: node scripts/e2e-test.mjs <path-to-competencia.docx> <output-dir> [base-url]");
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
await fileInput.setInputFiles(DOCX_PATH);

await page.locator("#selected-file-name").waitFor({ state: "visible" });
console.log("selected file text:", await page.locator("#selected-file-name").textContent());

await page.locator("#generate-btn").click();

await page.locator("#summary").waitFor({ state: "visible", timeout: 15000 });
const summaryText = await page.locator("#summary").innerText();
console.log("=== SUMMARY ===\n" + summaryText);

const errorVisible = await page.locator("#error").isVisible();
console.log("error visible:", errorVisible);
if (errorVisible) {
  console.log("error text:", await page.locator("#error").textContent());
}

console.log("download-btn visible:", await page.locator("#download-btn").isVisible());
console.log("download-btn text:", await page.locator("#download-btn").textContent());

const clickResult = await page.evaluate(() => {
  try {
    const btn = document.getElementById("download-btn");
    btn.click();
    return "clicked ok";
  } catch (e) {
    return "threw: " + (e && e.message);
  }
});
console.log("direct evaluate click result:", clickResult);

const downloadPromise = page.waitForEvent("download", { timeout: 10000 }).catch((e) => {
  console.log("no download event:", e.message);
  return null;
});
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
