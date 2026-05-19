// =============================================================================
// Pure grading logic — works in both browser and Node.
// In the browser, exposes `window.Grader`. In Node, `module.exports`.
// =============================================================================

(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.Grader = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {

  function colToNum(col) {
    let n = 0;
    for (let i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }
  function numToCol(n) {
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function readCell(wb, sheetName, cellRef) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) return null;
    const cell = sheet[cellRef];
    if (!cell) return null;
    return cell.v !== undefined ? cell.v : null;
  }

  function isBlank(v) {
    if (v === null || v === undefined) return true;
    if (typeof v !== "string") return false;
    const t = v.trim();
    if (t === "") return true;
    // Template placeholders the student should overwrite, e.g. "(compute)", "(from STEP B)",
    // "(your answer)", "(compare)".  Anything wrapped in parens is treated as unfilled.
    if (/^\(.*\)$/.test(t)) return true;
    return false;
  }

  function normalizeString(s, mode) {
    if (s == null) return "";
    let out = String(s);
    if (mode === "upper_strip") return out.toUpperCase().replace(/\s+/g, "");
    return out.trim();
  }

  function levenshtein(a, b) {
    const m = a.length, n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;
    const dp = Array(n + 1).fill(0).map((_, i) => i);
    for (let i = 1; i <= m; i++) {
      let prev = dp[0]; dp[0] = i;
      for (let j = 1; j <= n; j++) {
        const tmp = dp[j];
        dp[j] = a[i - 1] === b[j - 1] ? prev : Math.min(prev, dp[j - 1], dp[j]) + 1;
        prev = tmp;
      }
    }
    return dp[n];
  }

  function getOptionLabel(q, letter) {
    if (!q.options) return "";
    const o = q.options.find(x => x.letter === letter);
    if (!o) return "";
    if (o.value !== undefined) return `${letter}) ${o.value}`;
    if (o.keyword) return `${letter}) ${o.keyword}`;
    if (o.values) return `${letter}) [${o.values.join(", ")}]`;
    if (o.description) return `${letter}) ${o.description}`;
    return letter;
  }

  // ---------- Graders by question type ----------

  function gradeNumericChoice(wb, q) {
    const v = readCell(wb, q.ref_sheet, q.ref_cell);
    if (isBlank(v)) return { letter: "blank", studentValue: null, partial: false };
    const num = typeof v === "number" ? v : parseFloat(v);
    if (isNaN(num)) return { letter: "unknown", studentValue: v, partial: false };

    const tolAbs = q.tolerance_abs ?? 0;
    const tolFuzzyPct = q.tolerance_fuzzy_pct ?? 0;

    let exactMatch = null, exactDiff = Infinity;
    let fuzzyMatch = null, fuzzyPct = Infinity;

    for (const opt of q.options) {
      const diff = Math.abs(num - opt.value);
      if (diff <= tolAbs && diff < exactDiff) {
        exactMatch = opt;
        exactDiff = diff;
      } else if (tolFuzzyPct > 0) {
        const denom = Math.abs(opt.value) > 1e-9 ? Math.abs(opt.value) : 1;
        const pct = (diff / denom) * 100;
        if (pct <= tolFuzzyPct && pct < fuzzyPct) {
          fuzzyMatch = opt;
          fuzzyPct = pct;
        }
      }
    }
    if (exactMatch) return { letter: exactMatch.letter, studentValue: num, partial: false };
    if (fuzzyMatch) return { letter: fuzzyMatch.letter, studentValue: num, partial: true };
    return { letter: "no_match", studentValue: num, partial: false };
  }

  function gradeStringChoice(wb, q) {
    const v = readCell(wb, q.ref_sheet, q.ref_cell);
    if (isBlank(v)) return { letter: "blank", studentValue: null, partial: false };
    const mode = q.normalize || "upper_strip";
    const sn = normalizeString(v, mode);
    for (const opt of q.options) {
      if (normalizeString(opt.value, mode) === sn) {
        return { letter: opt.letter, studentValue: v, partial: false };
      }
    }
    let best = null, bestDist = Infinity;
    for (const opt of q.options) {
      const d = levenshtein(normalizeString(opt.value, mode), sn);
      if (d < bestDist) { bestDist = d; best = opt; }
    }
    if (best && bestDist <= Math.max(2, Math.floor(sn.length * 0.2))) {
      return { letter: best.letter, studentValue: v, partial: true };
    }
    return { letter: "no_match", studentValue: v, partial: false };
  }

  function gradeTextChoice(wb, q) {
    const v = readCell(wb, q.ref_sheet, q.ref_cell);
    if (isBlank(v)) return { letter: "blank", studentValue: null, partial: false };
    const sv = String(v).toUpperCase().trim();
    const matches = q.options.filter(opt => sv.includes(opt.keyword.toUpperCase()));
    if (matches.length === 0) return { letter: "no_match", studentValue: v, partial: false };
    const correctMatch = matches.find(opt => opt.letter === q.correct);
    if (correctMatch) {
      // Even if several options share the keyword (e.g. STAY for A/C or MOVE for B/D),
      // we credit the correct letter at full credit — the cell value is consistent with it.
      return { letter: q.correct, studentValue: v, partial: false };
    }
    return { letter: matches[0].letter, studentValue: v, partial: false };
  }

  function gradeArgmaxChoice(wb, q) {
    const values = q.compare_cells.map(c => ({
      ...c,
      val: readCell(wb, q.ref_sheet, c.cell)
    }));
    if (values.every(v => isBlank(v.val))) {
      return { letter: "blank", studentValue: null, partial: false };
    }
    if (values.some(v => isBlank(v.val))) {
      return { letter: "blank", studentValue: values.map(v => v.val), partial: false };
    }
    const nums = values.map(v => ({ ...v, num: typeof v.val === "number" ? v.val : parseFloat(v.val) }));
    if (nums.some(v => isNaN(v.num))) {
      return { letter: "unknown", studentValue: nums.map(v => v.val), partial: false };
    }
    const maxNum = Math.max(...nums.map(v => v.num));
    const eqThresh = (q.equal_threshold_pct || 5) / 100;
    const winners = nums.filter(v => maxNum > 0 ? Math.abs(v.num - maxNum) / maxNum < eqThresh : v.num === maxNum);
    let letter;
    if (winners.length > 1 && q.equal_letter) letter = q.equal_letter;
    else letter = winners[0].letter_if_max;
    const studentValue = nums.map(v => `${v.label}=${v.num.toFixed(2)}`).join(", ");
    return { letter, studentValue, partial: false };
  }

  function gradeRangeChoice(wb, q) {
    const v = readCell(wb, q.ref_sheet, q.ref_cell);
    if (isBlank(v)) return { letter: "blank", studentValue: null, partial: false };
    const num = typeof v === "number" ? v : parseFloat(v);
    if (isNaN(num)) return { letter: "unknown", studentValue: v, partial: false };
    for (const r of q.ranges) {
      if (num >= r.min && num <= r.max) {
        return { letter: r.letter, studentValue: num, partial: false };
      }
    }
    return { letter: "no_match", studentValue: num, partial: false };
  }

  function gradeGridRowContains(wb, q) {
    const m = q.grid_range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) return { letter: "unknown", studentValue: null, partial: false };
    const startCol = colToNum(m[1]), endCol = colToNum(m[3]);
    const startRow = parseInt(m[2]), endRow = parseInt(m[4]);
    const required = new Set(q.required_letters.map(l => l.toUpperCase()));
    const rows = [];
    for (let r = startRow; r <= endRow; r++) {
      const letters = new Set();
      for (let c = startCol; c <= endCol; c++) {
        const ref = numToCol(c) + r;
        const v = readCell(wb, q.ref_sheet, ref);
        if (!isBlank(v)) {
          String(v).toUpperCase().split(/[/\s,;]+/).forEach(L => {
            if (L) letters.add(L);
          });
        }
      }
      rows.push(letters);
    }
    if (rows.every(s => s.size === 0)) {
      return { letter: "blank", studentValue: null, partial: false };
    }
    let exactRowIdx = -1, partialRowIdx = -1, partialMatches = 0;
    for (let i = 0; i < rows.length; i++) {
      const allReq = [...required].every(L => rows[i].has(L));
      if (allReq) { exactRowIdx = i + 1; break; }
      const matchCount = [...required].filter(L => rows[i].has(L)).length;
      if (matchCount > partialMatches) {
        partialMatches = matchCount;
        partialRowIdx = i + 1;
      }
    }
    const display = rows.map((s, i) => `Row ${i + 1}: ${[...s].join("")}`).join(" | ");
    if (exactRowIdx > 0 && q.row_to_letter[String(exactRowIdx)]) {
      return { letter: q.row_to_letter[String(exactRowIdx)], studentValue: display, partial: false };
    }
    if (partialRowIdx > 0 && partialMatches >= required.size - 1 && q.row_to_letter[String(partialRowIdx)]) {
      return { letter: q.row_to_letter[String(partialRowIdx)], studentValue: display, partial: true };
    }
    return { letter: "no_match", studentValue: display, partial: false };
  }

  function gradeStringTupleChoice(wb, q) {
    const values = q.ref_cells.map(c => readCell(wb, q.ref_sheet, c));
    if (values.every(isBlank)) {
      return { letter: "blank", studentValue: null, partial: false };
    }
    const mode = q.normalize || "upper_strip";
    const norm = values.map(v => normalizeString(v, mode));
    for (const opt of q.options) {
      const optNorm = opt.values.map(s => normalizeString(s, mode));
      if (norm.length === optNorm.length && norm.every((s, i) => s === optNorm[i])) {
        return { letter: opt.letter, studentValue: values.join(" | "), partial: false };
      }
    }
    let best = null, bestMatches = 0;
    for (const opt of q.options) {
      const optNorm = opt.values.map(s => normalizeString(s, mode));
      const matches = norm.filter((s, i) => s === optNorm[i]).length;
      if (matches > bestMatches) { bestMatches = matches; best = opt; }
    }
    if (best && bestMatches > 0 && bestMatches < norm.length) {
      return { letter: best.letter, studentValue: values.join(" | "), partial: true };
    }
    return { letter: "no_match", studentValue: values.join(" | "), partial: false };
  }

  function gradeFreqPairChoice(wb, q) {
    const [posCell, negCell] = q.ref_cells_pair;
    const posRaw = readCell(wb, q.ref_sheet, posCell);
    const negRaw = readCell(wb, q.ref_sheet, negCell);
    if (isBlank(posRaw) && isBlank(negRaw)) {
      return { letter: "blank", studentValue: null, partial: false };
    }
    const pos = typeof posRaw === "number" ? posRaw : parseFloat(posRaw);
    const neg = typeof negRaw === "number" ? negRaw : parseFloat(negRaw);
    const posOk = !isNaN(pos) && pos > 0;
    const negOk = !isNaN(neg) && neg > 0;
    let key;
    if (posOk && !negOk) key = "pos>0 and neg==0";
    else if (!posOk && negOk) key = "pos==0 and neg>0";
    else if (posOk && negOk) key = "pos>0 and neg>0";
    else key = "pos==0 and neg==0";
    const opt = q.options.find(o => o.logic && o.logic.includes(key));
    return {
      letter: opt ? opt.letter : "no_match",
      studentValue: `pos=${posRaw}, neg=${negRaw}`,
      partial: false
    };
  }

  // ---------- Dispatcher ----------

  function gradeQuestion(wb, q, pointsEach) {
    let answer;
    try {
      switch (q.type) {
        case "numeric_choice":      answer = gradeNumericChoice(wb, q); break;
        case "string_choice":       answer = gradeStringChoice(wb, q); break;
        case "text_choice":         answer = gradeTextChoice(wb, q); break;
        case "argmax_choice":       answer = gradeArgmaxChoice(wb, q); break;
        case "range_choice":        answer = gradeRangeChoice(wb, q); break;
        case "grid_row_contains":   answer = gradeGridRowContains(wb, q); break;
        case "string_tuple_choice": answer = gradeStringTupleChoice(wb, q); break;
        case "freq_pair_choice":    answer = gradeFreqPairChoice(wb, q); break;
        default:                    answer = { letter: "manual", studentValue: null, partial: false };
      }
    } catch (e) {
      console.error("Grading error", q.id, e);
      answer = { letter: "unknown", studentValue: String(e), partial: false };
    }

    let grade = 0;
    let status = "wrong";
    if (answer.letter === "blank") {
      status = "blank";
    } else if (answer.letter === "manual") {
      status = "manual";
    } else if (answer.letter === "unknown") {
      // Cell had something we couldn't parse (e.g. garbage text in a numeric field).
      // Treated as an attempt: 50% credit.
      grade = pointsEach / 2;
      status = "attempted";
    } else if (answer.letter === q.correct) {
      grade = answer.partial ? pointsEach / 2 : pointsEach;
      status = answer.partial ? "partial" : "correct";
    } else if (answer.letter === "no_match") {
      // Student computed *something* (a real value), but it doesn't match any option.
      // We can't tell which letter they'd pick in Canvas — but they clearly attempted,
      // so we credit 50%.  Per teacher's instruction.
      grade = pointsEach / 2;
      status = "attempted";
    } else {
      // Student's value matches a specific wrong option — they'd pick that letter in Canvas.
      grade = 0;
      status = "wrong";
    }

    return {
      id: q.id,
      question: q.question,
      correct_letter: q.correct,
      correct_label: getOptionLabel(q, q.correct),
      student_letter: answer.letter,
      student_label: ["A", "B", "C", "D"].includes(answer.letter) ? getOptionLabel(q, answer.letter) : "",
      student_value: answer.studentValue,
      grade,
      max: pointsEach,
      status,
      ref: q.ref_sheet ? `${q.ref_sheet}!${q.ref_cell || (q.ref_cells || q.ref_cells_pair || [q.grid_range || ""]).join(",")}` : ""
    };
  }

  function detectExercise(wb, rubrics) {
    for (const r of rubrics) {
      if (r.detect_by_sheet.every(name => wb.SheetNames.includes(name))) {
        return r;
      }
    }
    return null;
  }

  // Returns every rubric whose required sheets are all present in this workbook.
  // A single xlsx may contain several exercises (e.g. a student copied multiple
  // exercise sheets into one workbook) — return one rubric per match.
  function detectAllExercises(wb, rubrics) {
    return rubrics.filter(r =>
      r.detect_by_sheet.every(name => wb.SheetNames.includes(name))
    );
  }

  function gradeWorkbook(wb, rubric) {
    const results = [];
    let total = 0, maxTotal = 0;
    for (const q of rubric.mc_questions) {
      const r = gradeQuestion(wb, q, rubric.mc_points_each);
      results.push(r);
      total += r.grade;
      maxTotal += r.max;
    }
    return { rubric, results, total, maxTotal };
  }

  return {
    gradeWorkbook,
    detectExercise,
    detectAllExercises,
    gradeQuestion,
    getOptionLabel,
    // Exposed for tests:
    _internals: {
      readCell, isBlank, normalizeString, levenshtein, colToNum, numToCol,
      gradeNumericChoice, gradeStringChoice, gradeTextChoice, gradeArgmaxChoice,
      gradeRangeChoice, gradeGridRowContains, gradeStringTupleChoice, gradeFreqPairChoice
    }
  };
});
