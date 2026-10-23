# iba

BUSM1042 Assignment 3 (group report), Q2: how booking lead time affects the likelihood of cancellation.

## `output/`

| File | What it is |
|---|---|
| `BUSM1042_A3_Q2_LeadTime_v3.xlsx` | **Current workbook.** Raw data, cleaned data, formula-driven summary tables, self-checks and Figure 2 (native Excel clustered column chart). Formulas calculate on open. |
| `Figure2_v3_preview.png`, `Q2_Summary_v3_preview.pdf` | LibreOffice render previews of v3, for checking only (not for the report). |
| `verification_report_v3.md` | Every v3 check: pandas vs openpyxl loop vs Excel formulas vs expected values, chart XML assertions and the v1 → v3 change log. |
| `build_workbook.py`, `verify.py` | Scripts that rebuild and re-verify the current (v3) workbook. |
| `BUSM1042_A3_Q2_LeadTime.xlsx`, `Figure2_preview.png`, `Q2_Summary_preview.pdf`, `verification_report.md` | v1 (line-chart Figure 2), kept unchanged for reference. |

## Rebuild

The source dataset is not stored in this repository. Put a copy outside `output/` and run:

```bash
python3 -I output/build_workbook.py <IBA-A3-Hotel_bookings_dataset-T32026.xlsx> build/BUSM1042_A3_Q2_LeadTime_v3.xlsx
python3 -I output/verify.py recalc build/BUSM1042_A3_Q2_LeadTime_v3.xlsx recalc <path to xlsx skill>/scripts/recalc.py
python3 -I output/verify.py check  <dataset.xlsx> build/BUSM1042_A3_Q2_LeadTime_v3.xlsx recalc/BUSM1042_A3_Q2_LeadTime_v3.xlsx output/verification_report_v3.md
python3 -I output/verify.py render recalc/BUSM1042_A3_Q2_LeadTime_v3.xlsx output
```

## Figure 2 into Word

Open the workbook in desktop Excel, let it calculate, then Save. Go to sheet `Figure2` → click the chart → Copy → in Word, Paste Special → Picture (Enhanced Metafile). Put the caption above it and the note and source lines below it.
