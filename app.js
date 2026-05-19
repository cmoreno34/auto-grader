// =============================================================================
// Auto Grader — UI / file handling.  Grading logic lives in grader.js.
// =============================================================================

const RUBRICS = [];
let ALL_RESULTS = [];

async function loadRubrics() {
  const resp = await fetch("rubrics/manifest.json");
  const manifest = await resp.json();
  for (const id of manifest.rubrics) {
    const r = await fetch(`rubrics/${id}.json`);
    RUBRICS.push(await r.json());
  }
  document.getElementById("supported-list").textContent =
    RUBRICS.map(r => r.display_name).join(" · ");
}

async function processXlsx(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellFormula: false });
  const rubric = Grader.detectExercise(wb, RUBRICS);
  if (!rubric) {
    return {
      filename: file.name,
      error: `No matching exercise. Sheets found: [${wb.SheetNames.join(", ")}]`
    };
  }
  const graded = Grader.gradeWorkbook(wb, rubric);
  return { filename: file.name, ...graded };
}

async function processZip(file) {
  const out = [];
  const zip = await JSZip.loadAsync(file);
  const entries = Object.values(zip.files).filter(
    e => !e.dir && (e.name.toLowerCase().endsWith(".xlsx") || e.name.toLowerCase().endsWith(".xlsm"))
  );
  for (const e of entries) {
    try {
      const blob = await e.async("blob");
      const fakeFile = new File([blob], e.name);
      const result = await processXlsx(fakeFile);
      out.push(result);
    } catch (err) {
      out.push({ filename: e.name, error: String(err) });
    }
  }
  return out;
}

async function processFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".zip")) return await processZip(file);
  if (name.endsWith(".xlsx") || name.endsWith(".xlsm")) return [await processXlsx(file)];
  return [{ filename: file.name, error: "Unsupported file type." }];
}

// ---------- Rendering --------------------------------------------------------

function statusBadge(status) {
  const map = {
    correct:   { cls: "status-correct", txt: "✓ correct" },
    partial:   { cls: "status-partial", txt: "≈ close (50%)" },
    attempted: { cls: "status-partial", txt: "~ attempted (50%)" },
    wrong:     { cls: "status-wrong",   txt: "✗ wrong" },
    blank:     { cls: "status-blank",   txt: "— blank" },
    unknown:   { cls: "status-unknown", txt: "? unknown" },
    manual:    { cls: "status-manual",  txt: "manual" }
  };
  const m = map[status] || map.unknown;
  return `<span class="${m.cls}">${m.txt}</span>`;
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderResult(r) {
  if (r.error) {
    return `<div class="result-card">
      <h2><span class="filename">${escapeHtml(r.filename)}</span><span class="score-pill">— error —</span></h2>
      <p class="status error">${escapeHtml(r.error)}</p>
    </div>`;
  }
  const rows = r.results.map(row => `
    <tr>
      <td class="center">${escapeHtml(row.id)}</td>
      <td>${escapeHtml(row.question)}</td>
      <td class="cell-mono">${escapeHtml(row.correct_label)}</td>
      <td class="cell-mono">${escapeHtml(row.student_label || (row.student_value !== null && row.student_value !== undefined ? String(row.student_value) : ""))}</td>
      <td class="cell-mono">${escapeHtml(typeof row.student_value === "number" ? row.student_value : "")}</td>
      <td class="center">${statusBadge(row.status)}</td>
      <td class="right cell-mono">${row.grade}/${row.max}</td>
    </tr>`).join("");

  return `<div class="result-card">
    <h2>
      <span><span class="exercise-badge">${escapeHtml(r.rubric.display_name)}</span> <span class="filename">${escapeHtml(r.filename)}</span></span>
      <span class="score-pill">${r.total}/${r.maxTotal} points</span>
    </h2>
    <table>
      <thead>
        <tr>
          <th class="center">#</th>
          <th>Question</th>
          <th>Correct answer</th>
          <th>Student answer (mapped)</th>
          <th>Raw value in Excel</th>
          <th class="center">Status</th>
          <th class="right">Grade</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>`;
}

function render(results) {
  const root = document.getElementById("results");
  root.innerHTML = results.map(renderResult).join("");
  document.getElementById("export-bar").classList.toggle("hidden", results.length === 0);
}

function exportCsv() {
  const rows = [
    ["File", "Exercise", "Question", "Correct letter", "Correct label",
     "Student letter", "Student value in Excel", "Status", "Grade", "Max"]
  ];
  for (const r of ALL_RESULTS) {
    if (r.error) {
      rows.push([r.filename, "ERROR", "", "", "", "", r.error, "", "", ""]);
      continue;
    }
    for (const row of r.results) {
      rows.push([
        r.filename, r.rubric.display_name, row.id,
        row.correct_letter, row.correct_label,
        row.student_letter, String(row.student_value ?? ""),
        row.status, row.grade, row.max
      ]);
    }
    rows.push([r.filename, r.rubric.display_name, "TOTAL", "", "", "", "", "", r.total, r.maxTotal]);
  }
  const csv = rows.map(r => r.map(cell => {
    const s = String(cell ?? "");
    return /[",\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
  }).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `grading_results_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

async function handleFiles(fileList) {
  const status = document.getElementById("status");
  status.classList.remove("hidden", "error", "ok");
  status.textContent = `Processing ${fileList.length} file(s)…`;
  ALL_RESULTS = [];
  try {
    for (const f of fileList) {
      const r = await processFile(f);
      ALL_RESULTS.push(...r);
    }
    render(ALL_RESULTS);
    status.classList.add("ok");
    status.textContent = `Done. ${ALL_RESULTS.length} submission(s) graded.`;
  } catch (e) {
    console.error(e);
    status.classList.add("error");
    status.textContent = `Error: ${e.message || e}`;
  }
}

function setupUI() {
  const drop = document.getElementById("drop-zone");
  const input = document.getElementById("file-input");

  drop.addEventListener("click", () => input.click());
  drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("dragover"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("dragover"));
  drop.addEventListener("drop", e => {
    e.preventDefault();
    drop.classList.remove("dragover");
    handleFiles([...e.dataTransfer.files]);
  });
  input.addEventListener("change", e => handleFiles([...e.target.files]));

  document.getElementById("export-csv").addEventListener("click", exportCsv);
  document.getElementById("clear-btn").addEventListener("click", () => {
    ALL_RESULTS = [];
    document.getElementById("results").innerHTML = "";
    document.getElementById("status").classList.add("hidden");
    document.getElementById("export-bar").classList.add("hidden");
    input.value = "";
  });
}

(async function init() {
  setupUI();
  try {
    await loadRubrics();
  } catch (e) {
    const status = document.getElementById("status");
    status.classList.remove("hidden");
    status.classList.add("error");
    status.textContent = `Failed to load rubrics: ${e.message || e}`;
  }
})();
