# MKTG3007 figure redesign

Scripts that redraw Figure 1 (entry routes in Australian coffee) and Figure C1
(staged entry roadmap and the international distribution life cycle) and put
them back into the report.

```
pip install matplotlib numpy pillow
python make_figures.py                 # writes output/figure1.{png,svg}, output/figureC1.{png,svg}
python update_docx.py report.docx output/figure1.png output/figureC1.png report_newfigures.docx
```

- Each figure is drawn on a canvas at its exact print size (14.2 cm and 15.9 cm wide) in a
  millimetre coordinate system, so font sizes, marker diameters and line weights are what
  prints. Word shows the PNGs at 1:1, 400 dpi.
- `make_figures.py` prints a layout check (text against text, marks, curves and the canvas
  edge) and the computed crossover month for Figure C1.
- `update_docx.py` replaces `word/media/image1.png` and `image2.png` under the same names,
  sets only the two image heights in `document.xml` to the PNG aspect ratio, and asserts
  that nothing else in `document.xml` changed.
