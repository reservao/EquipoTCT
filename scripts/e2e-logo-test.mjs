// Verifies the company-logo replacement: without an uploaded logo the left
// slot should end up empty (Circular HR on the right untouched); with one,
// the left slot should carry the uploaded image and Circular HR stays put.
//   node scripts/e2e-logo-test.mjs <docx> <logo-png> <output-dir> [base-url]
import { chromium } from "playwright";
import path from "path";

const DOCX_PATH = process.argv[2];
const LOGO_PATH = process.argv[3];
const OUT_DIR = process.argv[4];
const BASE_URL = process.argv[5] ?? "http://localhost:8899/index.html";

if (!DOCX_PATH || !LOGO_PATH || !OUT_DIR) {
  console.error("Usage: node scripts/e2e-logo-test.mjs <docx> <logo-png> <output-dir> [base-url]");
  process.exit(1);
}

const executablePath =
  process.env.PLAYWRIGHT_CHROMIUM_PATH ?? "/opt/pw-browsers/chromium-1194/chrome-linux/chrome";

async function run(withLogo) {
  const browser = await chromium.launch({ executablePath });
  const page = await browser.newPage();
  page.on("pageerror", (err) => console.log("[pageerror]", err.message));

  await page.goto(BASE_URL);
  await page.locator("#file-input").setInputFiles(DOCX_PATH);
  if (withLogo) {
    await page.locator("#logo-input").setInputFiles(LOGO_PATH);
  }
  await page.locator("#generate-btn").click();
  await page.locator("#results").waitFor({ state: "visible", timeout: 30000 });

  const downloadPromise = page.waitForEvent("download", { timeout: 15000 });
  await page.locator("#download-all-btn").click();
  const download = await downloadPromise;
  const outPath = path.join(OUT_DIR, withLogo ? "with-logo.xlsx" : "no-logo.xlsx");
  await download.saveAs(outPath);
  console.log((withLogo ? "WITH logo" : "NO logo") + " saved to:", outPath);

  await browser.close();
}

await run(false);
await run(true);
