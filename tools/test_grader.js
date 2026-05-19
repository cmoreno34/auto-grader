// =============================================================================
// Node-side smoke test for the grader.
// Loads every test_submissions/*.xlsx, grades it, and asserts the expected score.
// Run with:   node tools/test_grader.js
// =============================================================================

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const Grader = require("../grader.js");

const ROOT = path.resolve(__dirname, "..");
const RUBRICS_DIR = path.join(ROOT, "rubrics");
const SUBS_DIR = path.join(ROOT, "test_submissions");

function loadRubrics() {
  const manifest = JSON.parse(fs.readFileSync(path.join(RUBRICS_DIR, "manifest.json"), "utf-8"));
  return manifest.rubrics.map(id =>
    JSON.parse(fs.readFileSync(path.join(RUBRICS_DIR, `${id}.json`), "utf-8"))
  );
}

function gradeFile(file, rubrics) {
  const wb = XLSX.readFile(file);
  const rubric = Grader.detectExercise(wb, rubrics);
  if (!rubric) return { error: "no rubric matched", sheets: wb.SheetNames };
  return Grader.gradeWorkbook(wb, rubric);
}

function summarize(graded) {
  const breakdown = graded.results.map(r => `${r.id}=${r.student_letter}(${r.status})`).join("  ");
  return `${graded.total}/${graded.maxTotal} pts — ${breakdown}`;
}

function main() {
  const rubrics = loadRubrics();
  const files = fs.readdirSync(SUBS_DIR).filter(f => f.endsWith(".xlsx")).sort();

  let passed = 0, failed = 0;
  const lines = [];

  for (const f of files) {
    const full = path.join(SUBS_DIR, f);
    const g = gradeFile(full, rubrics);
    if (g.error) {
      lines.push(`FAIL ${f}: ${g.error} (sheets: ${g.sheets.join(", ")})`);
      failed++;
      continue;
    }

    // Expected outcome from filename suffix.
    let expected;
    if (f.endsWith("_perfect.xlsx")) expected = "perfect";
    else if (f.endsWith("_wrong.xlsx")) expected = "wrong";
    else if (f.endsWith("_blank.xlsx")) expected = "blank";
    else expected = "?";

    // Ignore manual / conceptual questions in the assertion — they can never be auto-graded.
    const autoGraded = g.results.filter(r => r.status !== "manual");
    const allCorrect = autoGraded.every(r => r.status === "correct");
    const noneCorrect = autoGraded.every(r => r.status !== "correct" && r.status !== "partial");

    // For "blank" and "wrong" we just require that no auto-gradable question scored points.
    // (Some templates ship with a few non-answer cells pre-populated, which
    // the grader will faithfully reflect in its decision — that's expected.)
    let ok = false;
    if (expected === "perfect") ok = allCorrect;
    else if (expected === "blank") ok = noneCorrect;
    else if (expected === "wrong") ok = noneCorrect;

    const flag = ok ? "PASS" : "FAIL";
    if (ok) passed++; else failed++;
    lines.push(`${flag} ${f.padEnd(28)} [exp ${expected.padEnd(7)}] ${summarize(g)}`);
  }

  console.log(lines.join("\n"));
  console.log("");
  console.log(`Summary: ${passed} passed, ${failed} failed, of ${files.length} files.`);
  process.exit(failed === 0 ? 0 : 1);
}

main();
