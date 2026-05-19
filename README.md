# Auto Grader — Big Data Final Exam

Browser-only grader for the practice Excel exercises. The teacher (or each student) drops one or several `.xlsx` files — or a `.zip` containing many — and the page produces a graded table per submission. Everything runs locally in the browser; nothing is uploaded and there is no API cost.

## Supported exercises

Detection is automatic from the sheet names of the uploaded file:

| Exercise              | Detect-by sheets                                                 | MC questions  |
|-----------------------|------------------------------------------------------------------|---------------|
| Cryptography (Crito)  | `1. Playfair`, `2. Rail Fence`                                   | 4 (8 pts)     |
| k-Means & Silhouette  | `Student - WCSS k=3`, `Student - Silhouette`                     | 4 (8 pts)     |
| Logit & Sentiment     | `PART A&B · Template`, `LOGIT · Template`                        | 5 (10 pts)    |
| Schelling Segregation | `Schelling_Template`, `Dynamics_Template`                        | 4 (8 pts)     |
| Translation           | `EX5_STEP_A_TPL`, `EX5_STEP_B_TPL`, `EX5_STEP_C_TPL`, `EX5_STEP_D_TPL` | 5 auto + 1 manual (10 + 2 pts) |

Each exercise is graded only on its multiple-choice questions (2 points each, as per the spec). The grader **infers** which letter the student would have picked in Canvas, based on the value they computed in the corresponding Excel cell.

## Local preview

Because the page fetches the JSON rubrics over HTTP, opening `index.html` directly from the filesystem will not work in some browsers. The easiest way to preview locally:

```powershell
cd auto_grader
python -m http.server 8000
# open http://localhost:8000
```

## Deploy to GitHub Pages

1. Create a new repo (e.g. `cmoreno34/auto-grader`).
2. From inside `auto_grader/`:
   ```powershell
   git init
   git add index.html app.js grader.js styles.css rubrics README.md
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/cmoreno34/auto-grader.git
   git push -u origin main
   ```
3. On GitHub: **Settings → Pages → Source = Deploy from a branch → main → / (root) → Save**.
4. Your page is at `https://cmoreno34.github.io/auto-grader/` after ~1 minute.

Note: do **not** commit `node_modules/`, `test_submissions/`, or `sample_batch.zip` if you don't want them online (a `.gitignore` is recommended).

## Project layout

```
auto_grader/
├── index.html           # UI (drop area, results, export)
├── app.js               # UI wiring (drag/drop, zip, CSV export)
├── grader.js            # Pure grading logic — also runnable from Node
├── styles.css
├── rubrics/
│   ├── manifest.json    # lists which rubrics exist
│   ├── crito.json
│   ├── kmeans.json
│   ├── logit.json
│   └── schelling.json
├── tools/
│   ├── make_test_submissions.py   # generates synthetic perfect/wrong/blank xlsx
│   └── test_grader.js             # Node smoke test against those xlsx
└── test_submissions/    # generated; safe to delete
```

## How a rubric works

Each `rubrics/<exercise>.json` describes the multiple-choice questions for one exercise. Each question has a `type` that drives how the grader picks a letter from the student's Excel:

| Type                  | What it does                                                                                          |
|-----------------------|--------------------------------------------------------------------------------------------------------|
| `numeric_choice`      | Reads one cell, finds the closest option value within tolerance                                       |
| `string_choice`       | Reads one cell, exact (normalised) string match against options; near-misses get 50% credit          |
| `text_choice`         | Reads one cell, keyword match (handles A/B/C/D options that share a keyword like STAY/MOVE)          |
| `argmax_choice`       | Reads several cells, picks the option mapped to the cell with the largest value                       |
| `range_choice`        | Reads one cell, picks the option whose `[min, max]` range contains the value                          |
| `grid_row_contains`   | Reads a 2D range, finds the row that contains a given set of letters                                  |
| `string_tuple_choice` | Reads several cells as a tuple, compares against tuple options                                        |
| `freq_pair_choice`    | Reads a `(pos, neg)` frequency pair, classifies into 4 cases                                          |

If a cell is empty the question is reported as `blank` (no points, no penalty). Near-correct numeric values can earn 50% credit via `tolerance_fuzzy_pct`.

## Testing

```powershell
cd auto_grader
python tools/make_test_submissions.py    # writes test_submissions/*.xlsx
node tools/test_grader.js                # 12 PASS, 0 FAIL
```

`make_test_submissions.py` generates three variants per exercise:
- `*_perfect.xlsx` — answer cells set to the correct value → expected 100%
- `*_wrong.xlsx`   — answer cells set to a deliberately wrong option → expected 0%
- `*_blank.xlsx`   — the bare template, untouched → expected 0% (blanks)

## Adding a new exercise

1. Inspect the new template/solution to identify each MC's reference cell.
2. Create `rubrics/<new_id>.json` following the schema (see existing rubrics).
3. Add `<new_id>` to `rubrics/manifest.json`.
4. (Optional) Re-run `tools/make_test_submissions.py` and `node tools/test_grader.js` to confirm.

That's it — no code changes needed.

## What the system does (and doesn't)

The students take the multiple-choice quiz in **Canvas**, not in the Excel. What the grader does:

> Given the student's filled-in Excel, **infer which letter they would have picked in Canvas** based on the value they computed in the corresponding cell, and tell you whether that inference matches the correct answer.

That gives the teacher a fast sanity check: if a student's Excel work agrees with the right MC option, their Canvas answer should also be right (and vice versa). The teacher still consults Canvas for the actual recorded answer.

`Q6` in the Translation exercise is conceptual (no Excel cell carries it), so it is reported as `manual` and never auto-graded.
