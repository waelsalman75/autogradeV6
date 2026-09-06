# Excel Auto Grader — Final v6 Batch

Codespaces-ready Excel auto-grader with **single-student or multi-student batch grading** through the same batch upload page.

## Uploads
1. Rubric `.xlsx`
2. Teacher Model Answer `.xlsx`
3. Original Student Template / Starter workbook `.xlsx`
4. One or more Student Completed workbooks `.xlsx`

## v6 grading rules included
- **Column width / row height:** ±0.2 tolerance = full marks; outside tolerance = 0.
- **Font partial credit:** correct/model font = full marks; wrong font that is different from the original starter-template default = 50%; unchanged starter-template default = 0.
- **Copied formula partial credit:** exact/equivalent copied formula = full marks; meaningful structural or calculated-value match = 50%; missing/unrelated formula = 0.
- **Merge partial credit:** correct target merge-and-centre = full marks; a new merge on other non-default cells = 50%; no relevant new merge = 0.
- **Chart partial credit:** correct chart family/source with only a 2D/3D mismatch = 50%.
- **Two-pass verification:** disagreements become `Needs review` rather than silent zeroes.
- **Filename normalization:** browser duplicate suffixes like `(1)` or `(2)` are ignored during filename checking.

## Batch outputs
- Batch Summary XLSX
- Batch Summary CSV
- ZIP of all individual grading reports
- Individual XLSX grading report for each student

## Run in GitHub Codespaces
```bash
python -m pip install -r requirements.txt
python app.py
```
Then open **port 8765** from the Codespaces **Ports** tab.

Or run:
```bash
bash run_codespace.sh
```

## Included rubric
`PE2_Ver_A_rubric_v6.xlsx` and `rubric_template.xlsx` both contain the latest v6 rubric.

## Notes
The starter/template workbook is required in this v6 package because template-aware rules need to distinguish an unchanged default format from an incorrect-but-attempted formatting change, and pre-existing merges from student-created merges.
