# HK.Finance — Financial Modeler

## What this prototype does
- Upload `.xlsx`, `.xls`, or `.csv`
- Detects a tabular sheet and year columns
- Maps common financial labels (Revenue, EBITDA, PAT, CFO, Capex, Debt, Cash, Equity, Assets, Shares, Price)
- Produces:
  - A — DCF Valuation
  - B — Contingent Valuation (Black-Scholes)
  - C — Financial Performance Dashboard
- Base / Best / Worst scenarios
- Dynamic browser recalculation
- Interactive charts
- Formula-linked Excel export
- PDF report export

## Expected input structure
The most robust source file has:
- one label column such as `Metric`, `Particulars`, or `Account`
- year columns such as `2021`, `2022`, `2023`, `2024`, `2025`
- rows for common financial statement metrics.

The mapping engine is alias-based, so future company files can use common synonyms.

## Production XLSM architecture
A real `.xlsm` file needs an embedded VBA project (`vbaProject.bin`). Browser-only spreadsheet libraries can generate formulas and workbook sheets, but should not be used to fabricate a VBA binary.

Recommended production architecture:
1. Keep a reviewed master `.xlsm` template with:
   - VBA buttons
   - refresh/reset/export macros
   - Goal Seek / Solver automation
   - dashboard controls
2. Server receives the generated model data/assumptions.
3. Server injects/updates the workbook sheets while preserving the VBA project.
4. Return the finished `.xlsm` and PDF.

## Suggested VBA macro set
- `RefreshDashboard`
- `ResetInputs`
- `SwitchScenario`
- `RunGoalSeek`
- `RunSolver`
- `ExportPDF`

## Banking / NBFC / insurance handling
The app flags these sectors because FCFF DCF is often unsuitable. Production logic should identify sector and switch to:
- Banks: Dividend Discount / Excess Return / Residual Income
- Insurers: Embedded Value / Appraisal Value / Residual Income
- NBFCs: Excess Return / Equity DCF with regulatory capital considerations

## Deployment
This prototype is a static site. For a production build:
- Frontend: React/Next.js
- Backend: FastAPI or Node.js
- Excel engine: LibreOffice headless + approved `.xlsm` template or a commercial Excel-compatible engine
- Storage: object storage
- Authentication: optional
- Audit log: recommended for finance workflows
