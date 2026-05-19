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
    else if (f.endsWith("_attempted.xlsx")) expected = "attempted";
    else expected = "?";

    // Ignore manual / conceptual questions in the assertion — they can never be auto-graded.
    const autoGraded = g.results.filter(r => r.status !== "manual");
    const allCorrect  = autoGraded.every(r => r.status === "correct");
    const allWrong    = autoGraded.every(r => r.status === "wrong");
    const noneScored  = autoGraded.every(r => r.grade === 0);
    // 'attempted' submissions write garbage values; not every type supports the
    // attempted state (argmax / freq_pair always land on a discrete letter).
    // The acceptance bar: at least one question got the attempted/partial credit
    // AND no question was graded correct (we wrote garbage values, not correct ones).
    const someAttempted = autoGraded.some(r => r.status === "attempted" || r.status === "partial");
    const noneCorrectStrict = autoGraded.every(r => r.status !== "correct");

    let ok = false;
    if (expected === "perfect")        ok = allCorrect;
    else if (expected === "blank")     ok = noneScored;
    else if (expected === "wrong")     ok = allWrong;
    else if (expected === "attempted") ok = someAttempted && noneCorrectStrict;

    const flag = ok ? "PASS" : "FAIL";
    if (ok) passed++; else failed++;
    lines.push(`${flag} ${f.padEnd(28)} [exp ${expected.padEnd(7)}] ${summarize(g)}`);
  }

  console.log(lines.join("\n"));
  console.log("");
  console.log(`Summary: ${passed} passed, ${failed} failed, of ${files.length} files.`);

  // ---- Multi-exercise detection check ----
  const fakeWb = {
    SheetNames: [
      // schelling sheets
      "Schelling_Template", "Dynamics_Template",
      // kmeans sheets
      "Student - WCSS k=3", "Student - Silhouette",
      // logit sheets
      "PART A&B · Template", "LOGIT · Template"
    ]
  };
  const matched = Grader.detectAllExercises(fakeWb, rubrics);
  const matchedIds = matched.map(r => r.exercise_id).sort();
  const expected = ["kmeans", "logit", "schelling"].sort();
  const multiOk = JSON.stringify(matchedIds) === JSON.stringify(expected);
  console.log(`\nMulti-detect: ${multiOk ? "PASS" : "FAIL"}  got [${matchedIds.join(", ")}]  expected [${expected.join(", ")}]`);
  if (!multiOk) failed++;

  process.exit(failed === 0 ? 0 : 1);
}

main();
