import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const workbookPath = "Report/MSc Major Project Lit. Review.xlsx";
const outputDir = ".codex_work/literature_render";
await fs.mkdir(outputDir, { recursive: true });

const input = await FileBlob.load(workbookPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheets = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 6000,
});
console.log("SHEETS");
console.log(sheets.ndjson);

const actualRanges = {
  "Full Systems": "A1:Q7",
  "FER": "A1:H8",
  "CWD": "A1:H7",
};

const requestedSheet = process.argv[2];
for (const sheet of workbook.worksheets.items) {
  if (requestedSheet && sheet.name !== requestedSheet) continue;
  const address = actualRanges[sheet.name];
  console.log(`\nSHEET ${sheet.name} USED ${address}`);
  const data = await workbook.inspect({
    kind: "table",
    sheetId: sheet.name,
    range: address,
    include: "values,formulas",
    tableMaxRows: 20,
    tableMaxCols: 20,
    tableMaxCellChars: 1200,
    maxChars: 120000,
  });
  console.log(data.ndjson);

  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  const safeName = sheet.name.replace(/[^a-z0-9_-]+/gi, "_");
  await fs.writeFile(
    `${outputDir}/${safeName}.png`,
    new Uint8Array(await preview.arrayBuffer()),
  );
}
