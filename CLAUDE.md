# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Run the app
streamlit run grading_app.py

# Install dependencies
pip install -r requirements.txt

# Quick test (no UI — runs module-level __main__ blocks)
python docx_parser.py
python scoring.py
```

## Code Architecture

### Data Flow

```
.docx file ──→ docx_parser.py ──→ StudentSubmission (dataclass)
                                       │
    User scores via Streamlit UI ──────┤
                                       │
                                       ↓
                               scoring.py ──→ .xlsx export
```

### Module Breakdown

**`grading_app.py`** (674 lines) — Streamlit UI:
- Sidebar: file uploads (reference docx + student docx batch), student selector with prev/next navigation, progress bar, export button
- Main area: tabbed per-major-question grading view, radio-button scoring for 程序 (0–3) and 结果 (0–2) per sub-question
- Auto-save to `output/autosave.json` on every score change, signature-guarded (only restores if same reference doc)
- Incomplete grading check before export

**`docx_parser.py`** (466 lines) — .docx parsing:
- State-machine parser (SEEK_MAJOR → IN_MAJOR → IN_SUB → CAPTURE_IMAGE) walks paragraphs
- Extracts student info from cover page via regex (学号/姓名/专业班级), with 3 fallback strategies
- Extracts inline/floating images from Word drawing ML (`wp:inline`, `wp:anchor`, `w:drawing`)
- Fills missing sub-questions with placeholders (ensures 4×5 structure)
- `question_signature()` generates a hash for variant detection

**`scoring.py`** (200 lines) — Scoring & export:
- `SCORE_CONFIG` dict governs max values (3+2 per sub, 5 per sub, 25 per major, 100 total)
- `export_to_xlsx()` builds a flat-column spreadsheet: student info → 程序 scores → 结果 scores → per-major totals → grand total
- Styled headers, frozen panes, highlighted total columns

### Key Design Decisions

- **Score key format**: `"{major}-{sub}-{section}"` (e.g. `"1-3-程序"`) — used across all modules as the shared key for the `scores` dict
- **Image extraction**: relies on Python's `python-docx` library and raw XML namespace traversal for Word drawing elements
- **Auto-save**: JSON to `output/autosave.json`, keyed by reference document signature to prevent cross-contamination
- **The app expects exactly 4 major questions with 5 sub-questions each**; any missing sub-questions are auto-filled with placeholders

### Scoring Rules

| Level | Max Score |
|-------|-----------|
| 程序 (program) per sub-question | 3.0 |
| 结果 (result) per sub-question | 2.0 |
| Sub-question total | 5.0 (capped) |
| Per major question (5 sub-qs) | 25.0 |
| Grand total (4 major qs) | 100.0 |
