# iba

## Boba's Tulip & Tea House 🌷🍵🎀

A peaceful, pink, Hello-Kitty-styled coffee & flower shop game in a single HTML file.
Open `index.html` in any browser to play. It needs no install or build, and it works offline
(only the cute fonts load from Google Fonts).

### How to play
- **Garden**: tap a pot to plant a pink or white tulip bulb, water it, tap to cheer it on, and pick the bloom.
- **Drink Bar**:
  - **Matcha Latte**: hold to pour the matcha, whisk in a W/M zigzag, hold to pour the milk, then seal and add the straw.
  - **Trà Xoài Kem Cheese Trân Châu Trắng**: hold to pour the mango tea, whip the kem cheese, scoop in the white pearls, then seal and add the straw.
- **Customers** of every age (kids, teens, grandparents…) drift in now and then, at most three at a time.
  They ask for flowers, drinks or both. There are no timers and no money, only kind words, which are saved
  in the 📖 Kindness Journal along with a sticker book.
- **Close the shop** any time for me-time: sip your own drink, kanpai with Kurimanju-senpai (“ハーッ!!!”),
  arrange a vase, listen to the rain, or take a nap.

Progress is saved in your browser (localStorage).

### Art
Every drawing is hand-made inline SVG, drawn by studying reference images (no images are bundled):
- **Tulips**: real *Tulipa gesneriana* photos. The cup has a broad front petal, side petals with soft
  points and a back petal peeking up. White tulips are tinted green at the base, and the leaves are broad
  and blue-green. Growth goes bulb → sprout → leaves → bud → opening → bloom.
- **Kurimanju-senpai**: anime stills and the official nendoroid. That gives the bun-shaped head, the brown
  chestnut cap with a curved edge, ear bumps, hatched blush, the "ω" mouth and the "ハーッ!!!" face.
- **Hello Kitty plush**: Sanrio character art. That gives the wide head, the bow on her left ear, oval eyes,
  the yellow nose, whiskers and no mouth.
- **Drinks**: photos of iced matcha lattes (the matcha floats on the denser milk) and Vietnamese trà xoài
  with see-through agar pearls and thick kem cheese.

## BUSM1042 Assignment 3

BUSM1042 Assignment 3 (group report), Q2: how booking lead time affects the likelihood of cancellation.

### `output/`: group dataset for Q1–Q5 (PROMPT A)

| File | What it is |
|---|---|
| `BUSM1042_A3_Group_Clean_Data.xlsx` | Shared, audited dataset: Raw_Data (exact source copy), Column_Profile, Clean_Data (`tblClean`, 39 derived and flag columns), Data_Quality_Log (DQ01–DQ20), Sensitivity, OneVar_Summary (21 one-variable tables) and Checks. Formulas calculate on open. |
| `BUSM1042_A3_Data_Quality_Notes.docx` / `.pdf` | The data-quality notes for the report and the lecturer; every number comes from the workbook's calculated values. |
| `verification_report_group.md` | Three-way verification (pandas, plain row loop, Excel formulas), Raw_Data integrity, Word number check, deviations. |
| `build_group_workbook.py`, `build_quality_notes.py`, `verify_group.py` | Scripts that rebuild and re-verify both files. |

Rebuild (source dataset kept outside the repo):

```bash
python3 -I output/build_group_workbook.py <dataset.xlsx> build_g/BUSM1042_A3_Group_Clean_Data.xlsx
python3 -I output/verify_group.py recalc build_g/BUSM1042_A3_Group_Clean_Data.xlsx recalc_g
python3 -I output/build_quality_notes.py recalc_g/BUSM1042_A3_Group_Clean_Data.xlsx build_g/BUSM1042_A3_Data_Quality_Notes.docx "<viewed date>"
python3 -I output/verify_group.py render build_g/BUSM1042_A3_Data_Quality_Notes.docx build_g/notes_render
python3 -I output/verify_group.py check <dataset.xlsx> build_g/BUSM1042_A3_Group_Clean_Data.xlsx recalc_g/BUSM1042_A3_Group_Clean_Data.xlsx build_g/BUSM1042_A3_Data_Quality_Notes.docx output/verification_report_group.md
```

### `output/`: Q2 lead-time workbook and Figure 2

| File | What it is |
|---|---|
| `BUSM1042_A3_Q2_LeadTime_v3.xlsx` | **Current workbook.** Raw data, cleaned data, formula-driven summary tables, self-checks and Figure 2 (native Excel clustered column chart). Formulas calculate on open. |
| `Figure2_v3_preview.png`, `Q2_Summary_v3_preview.pdf` | LibreOffice render previews of v3, for checking only (not for the report). |
| `verification_report_v3.md` | Every v3 check: pandas vs openpyxl loop vs Excel formulas vs expected values, chart XML assertions and the v1 → v3 change log. |
| `build_workbook.py`, `verify.py` | Scripts that rebuild and re-verify the current (v3) workbook. |
| `BUSM1042_A3_Q2_LeadTime.xlsx`, `Figure2_preview.png`, `Q2_Summary_preview.pdf`, `verification_report.md` | v1 (line-chart Figure 2), kept unchanged for reference. |

### Rebuild

The source dataset is not stored in this repository. Put a copy outside `output/` and run:

```bash
python3 -I output/build_workbook.py <IBA-A3-Hotel_bookings_dataset-T32026.xlsx> build/BUSM1042_A3_Q2_LeadTime_v3.xlsx
python3 -I output/verify.py recalc build/BUSM1042_A3_Q2_LeadTime_v3.xlsx recalc <path to xlsx skill>/scripts/recalc.py
python3 -I output/verify.py check  <dataset.xlsx> build/BUSM1042_A3_Q2_LeadTime_v3.xlsx recalc/BUSM1042_A3_Q2_LeadTime_v3.xlsx output/verification_report_v3.md
python3 -I output/verify.py render recalc/BUSM1042_A3_Q2_LeadTime_v3.xlsx output
```

### Figure 2 into Word

Open the workbook in desktop Excel, let it calculate, then Save. Go to sheet `Figure2` → click the chart → Copy → in Word, Paste Special → Picture (Enhanced Metafile). Put the caption above it and the note and source lines below it.
