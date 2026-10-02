# Senior 1 Accounting Notes

Free study notes, Excel workbooks and model questions for Senior 1 (high school) accounting. Free to use, print and share. Not for sale. All amounts are in Malaysian Ringgit (RM).

## Topics

| # | Topic | Notes | Excel workbook | PDF |
|---|---|---|---|---|
| 8 | Irrecoverable Debts and Allowance for Receivables | [notes](notes/Chapter_08_Irrecoverable_Debts.md) | [xlsx](excel/Ch08_Irrecoverable_Debts.xlsx) | [pdf](pdf/Chapter_08_Irrecoverable_Debts.pdf) |
| 10 | Tangible Non-current Assets and Depreciation | [notes](notes/Chapter_10_Depreciation.md) | [xlsx](excel/Ch10_Depreciation.xlsx) | [pdf](pdf/Chapter_10_Depreciation.pdf) |
| 11 | Accruals and Prepayments | [notes](notes/Chapter_11_Accruals_Prepayments.md) | [xlsx](excel/Ch11_Accruals_Prepayments.xlsx) | [pdf](pdf/Chapter_11_Accruals_Prepayments.pdf) |
| 12 | Fundamental Accounting Principles and Concepts | [notes](notes/Chapter_12_Accounting_Principles.md) | [xlsx](excel/Ch12_Accounting_Concepts.xlsx) | [pdf](pdf/Chapter_12_Accounting_Principles.pdf) |
| FS | Statement of Profit or Loss and Statement of Financial Position | [notes](notes/Financial_Statements.md) | [xlsx](excel/Financial_Statements.xlsx) | [pdf](pdf/Financial_Statements.pdf) |
| Revision | One-page revision sheets for every topic | [notes](notes/Revision_Sheets.md) | | [pdf](pdf/Revision_Sheets.pdf) |
| Glossary | Key terms in English and Chinese | [notes](notes/Glossary.md) | | [pdf](pdf/Glossary.pdf) |
| PA | Prinsip Akaun (Malay with English): Persamaan Perakaunan, Lejar dan Akaun Kawalan, SPM-style questions, Malay/English/Chinese glossary | [nota](prinsip-akaun/README.md) | [xlsx](prinsip-akaun/excel) | [pdf](prinsip-akaun/pdf) |

Model questions: [Model_Questions.md](model-questions/Model_Questions.md) ([pdf](pdf/Model_Questions.pdf))

Every topic has multiple choice questions, short answers and **5 structured questions** graded from Level 1 (easy) to Level 5 (challenge). Start at Level 1 and work up.

Answer key with full workings: [Answer_Key.md](model-questions/Answer_Key.md) ([pdf](pdf/Answer_Key.pdf))

Year-end journal practice: each chapter with journals (8, 10, 11 and the financial statements) has two extra self-marking journal questions (J8.1, J8.2 and so on) in its model questions section and workbook.

Mock test: [Mock_Test.md](model-questions/Mock_Test.md) ([pdf](pdf/Mock_Test.pdf)), a 2-hour, 100-mark paper covering all five topics. Mark scheme: [Mock_Test_Answers.md](model-questions/Mock_Test_Answers.md) ([pdf](pdf/Mock_Test_Answers.pdf)). Self-marking version: [Mock_Test.xlsx](excel/Mock_Test.xlsx), with a Scores tab that adds up your marks.

## How to study each topic

1. Read the notes. Each chapter has learning goals, key words, worked examples, common mistakes and a short summary.
2. Open the Excel workbook. On the orange **Worked** tabs, change any blue figure and watch every answer update.
3. Try the blue **Q** tabs in order, Level 1 to Level 5. Type answers in the yellow cells. The column beside each one shows ✓ (correct) or ✗ (try again), and your score is at the top.
4. Check written answers and full solutions on the green **Ans** tabs.
5. For extra practice, open the purple **P** tab. Type any set number from 1 to 999 to get a fresh set of figures for the same question. The answers are on the matching **P ... Ans** tab. These figures do not match the printed answer key.
6. Answer the same questions on paper from the model questions, then mark them with the answer key.

## Formats used

| Item | Columns |
|---|---|
| Journal | Date, Particulars, Debit, Credit. The first row shows the year, with RM under Debit and Credit |
| Ledger account | Dr side and Cr side, each with Date, Particulars, Folio, Amount. The first row under the headings shows the year, with RM in the Amount column; entries then show only the month and day. A new year row starts when the year changes. Folio: GJ = general journal, CB = cash book |
| Statement of Profit or Loss | Add/Less marker, Particulars, RM, RM, RM |
| Statement of Financial Position | Add/Less marker, Particulars, RM, RM, RM, with Cost, Accumulated Depreciation and Carrying Amount under the RM headings |

## What is in each Excel workbook

| Workbook | Worked tabs (orange) | Question tabs (blue), with answer tabs (green) | Practice tab (purple) |
|---|---|---|---|
| Ch08_Irrecoverable_Debts | Worked 8.1-8.2, Worked 8.3 | Q8.8 to Q8.12, QJ8.1, QJ8.2 | P8.11 |
| Ch10_Depreciation | Straight Line, Reducing Balance, Other Methods, Disposal | Q10.9 to Q10.13, QJ10.1, QJ10.2 | P10.12 |
| Ch11_Accruals_Prepayments | Expense Examples, Income Examples | Q11.8 to Q11.12, QJ11.1, QJ11.2 | P11.11 |
| Ch12_Accounting_Concepts | Concepts | Quiz, Q12.15 to Q12.19 | P12.17 |
| Financial_Statements | Worked Example (Amina Stores) | QFS.6 to QFS.10, QJFS.1, QJFS.2 | PFS.8 |
| Mock_Test | Scores | Mock A, Mock B1 to Mock B4, Mock C | |

Colour key in every workbook: blue figures are data you can change, black figures are formulas, yellow cells are for your answers, green cells show the correct answers.

## For teachers and maintainers

Each structured question is written once in `tools/build_excel.py`. The same definition produces the question tab, the answer tab, and the Markdown in the model questions, answer key and notes, so the figures always agree.

1. `python3 tools/build_excel.py` writes the workbooks and `tools/manifest.json` (needs `openpyxl`).
2. Open and save each workbook once in Excel or LibreOffice so the formulas calculate.
3. `python3 tools/build_docs.py` fills the generated sections of the Markdown files (between `BEGIN GENERATED` and `END GENERATED` markers). It uses LibreOffice to recalculate if step 2 was skipped.
4. `python3 tools/build_pdf.py` rebuilds the PDFs (needs the `markdown` package and Chromium).

Terms follow the Cambridge IGCSE / O Level style: revenue, trade receivables, irrecoverable debts, non-current assets, carrying amount, Statement of Profit or Loss, Statement of Financial Position. "Income statement" is used for the ledger entry name.
