import fs from "fs";
import { parseCompetenciaDocx } from "../src/lib/docx-parser";
import { generateSotWorkbook, buildOutputFileName } from "../src/lib/sot-generator";

async function main() {
  const inputPath = process.argv[2];
  const outputDir = process.argv[3] ?? "/tmp";
  const buffer = fs.readFileSync(inputPath);
  const data = await parseCompetenciaDocx(buffer);
  console.log(JSON.stringify(data, null, 2));

  const workbookBuffer = await generateSotWorkbook(data);
  const fileName = buildOutputFileName(data);
  const outPath = `${outputDir}/${fileName}`;
  fs.writeFileSync(outPath, workbookBuffer);
  console.log("Written:", outPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
