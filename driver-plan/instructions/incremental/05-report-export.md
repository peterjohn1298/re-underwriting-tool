# Milestone 5: Report Export

**Files:** `services/excel_generator.py`, `services/word_generator.py`, `services/pdf_generator.py`
**Depends on:** Milestone 1 (pipeline output dict)

## Goal
One-click Excel, Word, PDF report generation from pipeline output.

## Acceptance Criteria
- [ ] Excel: pro forma, sensitivity, market data, ML, MC, rent prediction sheets
- [ ] Word: narrative memo with all analysis sections
- [ ] PDF: one-page summary with key metrics and verdict
- [ ] All three wrapped in try/except — failures store None path gracefully
- [ ] Download buttons check os.path.exists() before rendering
- [ ] Files named with 12-char hex UUID job_id

## Key Notes
- Generators called at end of pipeline in streamlit_app.py
- OUTPUT_DIR from config.py — ensure directory exists before writing
- Streamlit Cloud may not support all PDF backends (WeasyPrint needs system libs)
