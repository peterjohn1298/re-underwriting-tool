# Validation — Section 5: Report Export

## Deployment Evidence
- Files: services/excel_generator.py, services/word_generator.py, services/pdf_generator.py
- Download buttons on Results page → Downloads tab

## Verified Behaviors
- [x] Excel generated at OUTPUT_DIR/underwriting_{job_id}.xlsx after pipeline completes
- [x] Word generated at OUTPUT_DIR/investment_memo_{job_id}.docx after pipeline completes
- [x] PDF generated at OUTPUT_DIR/investment_report_{job_id}.pdf after pipeline completes
- [x] Download buttons check os.path.exists(path) before rendering
- [x] "not available" caption shown gracefully if generation failed
- [x] Files named with job_id (12-char hex UUID) for uniqueness
- [x] All three generators receive: pro_forma, market_data, ml_valuation, monte_carlo,
      lease_analysis, rent_prediction, sensitivity, backtest, ai_memo, unit_mix

## Screenshot Instructions (Manual)
Screenshots to capture:
- Downloads tab with all three download buttons active
- Downloaded Excel file opened showing pro forma sheet
