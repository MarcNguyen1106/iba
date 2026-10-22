# iba

BUSM1042 Assignment 3 (group report), Q2: how booking lead time affects the likelihood of cancellation.

## `output/`

| File | What it is |
|---|---|
| `BUSM1042_A3_Q2_LeadTime.xlsx` | The workbook: raw data, cleaned data, formula-driven summary tables, self-checks and Figure 2 (native Excel chart). Formulas calculate on open. |
| `Figure2_preview.png`, `Q2_Summary_preview.pdf` | LibreOffice render previews, for checking only (not for the report). |
| `verification_report.md` | Every check: pandas vs openpyxl loop vs Excel formulas vs expected values. |
| `build_workbook.py`, `verify.py` | Scripts that rebuild and re-verify the workbook. |

## Rebuild

The source dataset is not stored in this repository. Put a copy outside `output/` and run:

```bash
python3 -I output/build_workbook.py <IBA-A3-Hotel_bookings_dataset-T32026.xlsx> build/BUSM1042_A3_Q2_LeadTime.xlsx
python3 -I output/verify.py recalc build/BUSM1042_A3_Q2_LeadTime.xlsx recalc <path to xlsx skill>/scripts/recalc.py
python3 -I output/verify.py check  <dataset.xlsx> build/BUSM1042_A3_Q2_LeadTime.xlsx recalc/BUSM1042_A3_Q2_LeadTime.xlsx output/verification_report.md
python3 -I output/verify.py render recalc/BUSM1042_A3_Q2_LeadTime.xlsx output
```

## Figure 2 into Word

Open the workbook in Excel → sheet `Figure2` → click the chart → Copy → in Word, Paste Special → Picture (Enhanced Metafile). Put the caption above it and the note and source lines below it.
