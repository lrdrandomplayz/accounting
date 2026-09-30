# Senior 1 Accounting Notes

Free study notes, Excel workbooks and model questions for Senior 1 (high school) accounting. Free to use, print and share. Not for sale.

## Topics

| # | Topic | Notes | Excel workbook | PDF |
|---|---|---|---|---|
| 8 | Irrecoverable Debts and Allowance for Receivables | [notes](notes/Chapter_08_Irrecoverable_Debts.md) | [xlsx](excel/Ch08_Irrecoverable_Debts.xlsx) | [pdf](pdf/Chapter_08_Irrecoverable_Debts.pdf) |
| 10 | Tangible Non-current Assets and Depreciation | [notes](notes/Chapter_10_Depreciation.md) | [xlsx](excel/Ch10_Depreciation.xlsx) | [pdf](pdf/Chapter_10_Depreciation.pdf) |
| 11 | Accruals and Prepayments | [notes](notes/Chapter_11_Accruals_Prepayments.md) | [xlsx](excel/Ch11_Accruals_Prepayments.xlsx) | [pdf](pdf/Chapter_11_Accruals_Prepayments.pdf) |
| 12 | Fundamental Accounting Principles and Concepts | [notes](notes/Chapter_12_Accounting_Principles.md) | [xlsx](excel/Ch12_Accounting_Concepts.xlsx) | [pdf](pdf/Chapter_12_Accounting_Principles.pdf) |
| FS | Statement of Profit or Loss and Statement of Financial Position | [notes](notes/Financial_Statements.md) | [xlsx](excel/Financial_Statements.xlsx) | [pdf](pdf/Financial_Statements.pdf) |

Model questions: [Model_Questions.md](model-questions/Model_Questions.md) ([pdf](pdf/Model_Questions.pdf))

Every topic has multiple choice questions, short answers and **5 structured questions** graded from Level 1 (easy) to Level 5 (challenge). Start at Level 1 and work up.

Answer key with full workings: [Answer_Key.md](model-questions/Answer_Key.md) ([pdf](pdf/Answer_Key.pdf))

## How to study each topic

1. Read the notes. Each chapter has learning goals, key words, worked examples, common mistakes and a short summary.
2. Open the Excel workbook. Change the blue numbers in the worked example sheets and watch the answers update.
3. Try the **Practice** sheet in the workbook. Type answers in the yellow cells. The Check column marks each one "Correct" or "Try again".
4. Answer the matching section of the model questions on paper.
5. Mark your work with the answer key.

## What is in each Excel workbook

| Workbook | Sheets |
|---|---|
| Ch08_Irrecoverable_Debts | Allowance Calculator, Allowance Ledger (T-account), Practice |
| Ch10_Depreciation | Straight Line, Reducing Balance, Compare Methods, Revaluation, Part Year, Disposal, Practice |
| Ch11_Accruals_Prepayments | Expense Calculator, Income Calculator, Time Apportion, Ledger Account, Practice |
| Ch12_Accounting_Concepts | Concepts table, NRV Calculator, Quiz with drop-down lists |
| Financial_Statements | Trial Balance, Adjustments, SPL, SFP (Amina Stores) and a Practice question (Ben's Bikes) |

Colour key in every workbook: blue text is an input you can change, black text is a formula, yellow cells are for your answers. The Answers sheet is hidden; right-click a sheet tab and choose Unhide to see it.

## For teachers and maintainers

* Rebuild the workbooks: `python3 tools/build_excel.py` (needs `openpyxl`). Open and save each file once in Excel or LibreOffice so the formulas calculate.
* Rebuild the PDFs: `python3 tools/build_pdf.py` (needs the `markdown` package and Chromium).
* Terms follow the Cambridge IGCSE / O Level style: revenue, trade receivables, irrecoverable debts, non-current assets, Statement of Profit or Loss, Statement of Financial Position. "Income statement" is used for the ledger entry name.
