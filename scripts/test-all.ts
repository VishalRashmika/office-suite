import { runPptxTestSuite } from "./test-pptx-suite";
import { runTestSuite as runPdfTestSuite } from "./test-pdf-suite";
import { runPdfTextSelectionTestSuite } from "./test-pdf-text-selection";

async function main() {
  console.log("==================================================");
  console.log("🚀 RUNNING FULL OFFICE SUITE TEST RUNNER");
  console.log("==================================================");

  // 1. Run PPTX Suite
  await runPptxTestSuite();

  // 2. Run PDF Suite
  await runPdfTestSuite();

  // 3. Run PDF Text Layer & Selection Suite
  await runPdfTextSelectionTestSuite();

  console.log("==================================================");
  console.log("⭐ ALL TEST SUITES PASSED! OFFICE SUITE READY! ⭐");
  console.log("==================================================");
}

main().catch((err) => {
  console.error("Test execution failed:", err);
  process.exit(1);
});
