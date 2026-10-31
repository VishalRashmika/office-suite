import { runPptxTestSuite } from "./test-pptx-suite";
import { runTestSuite as runPdfTestSuite } from "./test-pdf-suite";

async function main() {
  console.log("==================================================");
  console.log("🚀 RUNNING FULL OFFICE SUITE TEST RUNNER");
  console.log("==================================================");

  // 1. Run PPTX Suite
  await runPptxTestSuite();

  // 2. Run PDF Suite
  await runPdfTestSuite();

  console.log("==================================================");
  console.log("⭐ ALL TEST SUITES PASSED! OFFICE SUITE READY! ⭐");
  console.log("==================================================");
}

main().catch((err) => {
  console.error("Test execution failed:", err);
  process.exit(1);
});
