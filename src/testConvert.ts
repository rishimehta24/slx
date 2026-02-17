/**
 * Basic tests for converter behavior (WORKING-format alignment).
 * Run with: npm run build && node dist/testConvert.js
 */
import { convert, ConversionError } from "./convertPdfToXls";
import path from "path";
import fs from "fs";

function assert(cond: boolean, msg: string) {
  if (!cond) throw new Error(msg);
}

async function main() {
  console.log("Testing ConversionError shape...");
  const err = new ConversionError(
    "Missing incident date/time",
    "test.pdf",
    "incident_analysis_report",
    [9, 10],
    ["Incident Date/Time"],
    true
  );
  assert(err.name === "ConversionError", "ConversionError.name");
  assert(err.fileName === "test.pdf", "fileName");
  assert(err.sheetName === "incident_analysis_report", "sheetName");
  assert(Array.isArray(err.rowNumbers) && err.rowNumbers?.length === 2, "rowNumbers");
  assert(err.looksLikeMergedColumns === true, "looksLikeMergedColumns");
  console.log("  ConversionError OK");

  const samplePdf = process.env.TEST_PDF || path.join(__dirname, "..", "Anson-HTML.pdf");
  if (fs.existsSync(samplePdf)) {
    console.log("Running convert on", samplePdf, "...");
    const out = path.join(__dirname, "..", "test-output.xls");
    await convert(samplePdf, out);
    assert(fs.existsSync(out), "output file created");
    console.log("  Convert OK, output:", out);
  } else {
    console.log("Skip convert test (no TEST_PDF or Anson-HTML.pdf at project root)");
  }

  console.log("All tests passed.");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
