# Chapter 11: Accruals and Prepayments

> Excel file for this chapter: `excel/Ch11_Accruals_Prepayments.xlsx`. The worked examples below come from its orange tabs; questions 11.8 to 11.12 are self-marking on the blue tabs.
>
> All amounts are in Malaysian Ringgit (RM). In the ledger accounts below, a line with no amount in this example is left out.

## 1. Learning goals

1. Explain the accruals (matching) concept.
2. Identify accrued expenses, prepaid expenses, accrued income and income received in advance.
3. Calculate the correct expense or income for the Statement of Profit or Loss.
4. Prepare ledger accounts with balances carried down.
5. Show accruals and prepayments in the Statement of Financial Position.

## 2. The big idea

The Statement of Profit or Loss shows income **earned** and expenses **used up** in the year. It does **not** show only what was paid or received in cash.

> **Accruals (matching) concept:** match the income of a period with the expenses of the same period, whether or not cash has been paid or received.

So at the year end you adjust for anything paid **too little** or **too much**.

## 3. The four adjustments

| Adjustment | What it means | Example | SPL effect | SFP heading |
|---|---|---|---|---|
| Accrued expense (expense owing) | Used this year, **not yet paid** | December electricity bill arrives in January | **Add** to expense | Current liability (other payables) |
| Prepaid expense (paid in advance) | Paid this year, **for next year** | Insurance paid for 12 months from 1 April | **Subtract** from expense | Current asset (other receivables) |
| Accrued income (income due) | Earned this year, **not yet received** | Tenant has not paid December rent | **Add** to income | Current asset (other receivables) |
| Income received in advance | Received this year, **for next year** | Tenant pays January rent in December | **Subtract** from income | Current liability (other payables) |

**Memory trick:** if the business **owes** money it is a **liability**. If the business is **owed** something (money or a service) it is an **asset**.

## 4. The formulas

**Expense for the year**

> Expense = Amount paid - Opening accrual + Opening prepayment + Closing accrual - Closing prepayment

**Income for the year**

> Income = Amount received - Opening accrued income + Opening income in advance + Closing accrued income - Closing income in advance

Most questions have only one or two of these figures. Put 0 for the rest.

## 5. Accrued expense

### Worked example 11.1

Rent is RM1,000 per month. The year ends 31 December 2025. During the year the business paid RM11,000. December rent is still owing.

Expense for year = 12 x 1,000 = **RM12,000**. Accrual = 12,000 - 11,000 = **RM1,000**.

<!-- BEGIN GENERATED: ch11.ex1.ledger -->
**Dr** &emsp;&emsp; **Rent Account** &emsp;&emsp; **Cr**

| Date | Particulars | Folio | Amount | Date | Particulars | Folio | Amount |
|---|---|---|---:|---|---|---|---:|
| **2025** | | | **RM** | **2025** | | | **RM** |
|  | Bank (paid during the year) | CB1 | 11,000 | Dec 31 | Income statement | GJ1 | 12,000 |
| Dec 31 | Balance c/d (owing) |  | 1,000 |  | | |  |
| | | | **12,000** | | | | **12,000** |
|  | | |  | **2026** | | |  |
|  | | |  | Jan 1 | Balance b/d (owing) |  | 1,000 |
<!-- END GENERATED: ch11.ex1.ledger -->

The **credit** balance brought down is a **liability**: the business owes RM1,000 rent.

## 6. Prepaid expense

### Worked example 11.2

On 1 April 2025 the business paid RM3,600 insurance for 12 months. The year ends 31 December 2025.

* Monthly cost = 3,600 / 12 = 300.
* Months used this year (April to December) = 9. Expense = 9 x 300 = **RM2,700**.
* Months for next year (January to March) = 3. Prepaid = 3 x 300 = **RM900**.

<!-- BEGIN GENERATED: ch11.ex2.ledger -->
**Dr** &emsp;&emsp; **Insurance Account** &emsp;&emsp; **Cr**

| Date | Particulars | Folio | Amount | Date | Particulars | Folio | Amount |
|---|---|---|---:|---|---|---|---:|
| **2025** | | | **RM** | **2025** | | | **RM** |
|  | Bank (paid during the year) | CB1 | 3,600 | Dec 31 | Income statement | GJ1 | 2,700 |
|  | | |  | Dec 31 | Balance c/d (prepaid) |  | 900 |
| | | | **3,600** | | | | **3,600** |
| **2026** | | |  |  | | |  |
| Jan 1 | Balance b/d (prepaid) |  | 900 |  | | |  |
<!-- END GENERATED: ch11.ex2.ledger -->

The **debit** balance brought down is an **asset**: the business has already paid for 3 months of cover.

**Tip:** unused stationery or fuel at the year end is treated the same way. Deduct it from the expense and show it as a current asset.

## 7. Opening and closing balances together

### Worked example 11.3

Electricity: owing at the start of the year RM200. Paid during the year RM2,500. Owing at the end of the year RM300.

Expense = 2,500 - 200 + 300 = **RM2,600**

<!-- BEGIN GENERATED: ch11.ex3.ledger -->
**Dr** &emsp;&emsp; **Electricity Account** &emsp;&emsp; **Cr**

| Date | Particulars | Folio | Amount | Date | Particulars | Folio | Amount |
|---|---|---|---:|---|---|---|---:|
| **2025** | | | **RM** | **2025** | | | **RM** |
|  | Bank (paid during the year) | CB1 | 2,500 | Jan 1 | Balance b/d (owing) |  | 200 |
| Dec 31 | Balance c/d (owing) |  | 300 | Dec 31 | Income statement | GJ1 | 2,600 |
| | | | **2,800** | | | | **2,800** |
|  | | |  | **2026** | | |  |
|  | | |  | Jan 1 | Balance b/d (owing) |  | 300 |
<!-- END GENERATED: ch11.ex3.ledger -->

Why subtract the opening accrual? The RM200 belonged to **last** year's expense but was paid **this** year.

## 8. Income adjustments

### Worked example 11.4: income received in advance

A business rents out a room for RM500 per month. During 2025 it received RM6,500, which includes January 2026 rent.

Income for 2025 = 12 x 500 = **RM6,000**. Received in advance = **RM500** (current liability).

<!-- BEGIN GENERATED: ch11.ex4.ledger -->
**Dr** &emsp;&emsp; **Rent Received Account** &emsp;&emsp; **Cr**

| Date | Particulars | Folio | Amount | Date | Particulars | Folio | Amount |
|---|---|---|---:|---|---|---|---:|
| **2025** | | | **RM** | **2025** | | | **RM** |
| Dec 31 | Income statement | GJ1 | 6,000 |  | Bank (received during the year) | CB1 | 6,500 |
| Dec 31 | Balance c/d (in advance) |  | 500 |  | | |  |
| | | | **6,500** | | | | **6,500** |
|  | | |  | **2026** | | |  |
|  | | |  | Jan 1 | Balance b/d (in advance) |  | 500 |
<!-- END GENERATED: ch11.ex4.ledger -->

### Worked example 11.5: accrued income

Commission received during the year was RM1,800. A further RM200 is still due at the year end.

Income = 1,800 + 200 = **RM2,000**. The RM200 is a current asset (other receivables).

## 9. In the financial statements

**Statement of Profit or Loss (extract) for the year ended 31 December 2025**

| | Particulars | RM | RM | RM |
|---|---|---:|---:|---:|
| Add: | **Other income** | | | |
| | Rent received (6,500 - 500) | | 6,000 | |
| | Commission received (1,800 + 200) | | 2,000 | 8,000 |
| Less: | **Expenses** | | | |
| | Rent (11,000 + 1,000) | | 12,000 | |
| | Insurance (3,600 - 900) | | 2,700 | |
| | Electricity (2,500 - 200 + 300) | | 2,600 | (17,300) |

**Statement of Financial Position (extract) at 31 December 2025**

| | Particulars | RM | RM | RM |
|---|---|---:|---:|---:|
| | | **Cost** | **Accumulated Depreciation** | **Carrying Amount** |
| | **Current assets** | | | |
| | Other receivables (900 insurance prepaid + 200 commission due) | | 1,100 | |
| Less: | **Current liabilities** | | | |
| | Other payables (1,000 rent + 300 electricity + 500 rent received in advance) | | 1,800 | |

## 10. Effect on profit if you forget

| Forgotten item | Effect on profit |
|---|---|
| Accrued expense | Profit **overstated** |
| Prepaid expense | Profit **understated** |
| Accrued income | Profit **understated** |
| Income received in advance | Profit **overstated** |

## 11. Common mistakes

* Putting an accrued expense under current assets. It is a **liability**.
* Counting months wrongly. Write out the months if unsure.
* Using the cash paid figure in the SPL without adjusting.
* Placing the balance c/d on the wrong side. Accrued expense: c/d on the debit side, b/d on the credit side. Prepaid expense: c/d on the credit side, b/d on the debit side.

## 12. Quick summary

* SPL shows what belongs to the year, not what was paid.
* Accrued expense: add to expense, current liability.
* Prepaid expense: subtract from expense, current asset.
* Accrued income: add to income, current asset.
* Income in advance: subtract from income, current liability.

## 13. Check yourself

1. Wages paid RM24,000. Wages owing at the year end RM800. What is the wages expense?
2. Rates paid RM1,500 cover 15 months. Year end is after 12 months of the period. What is the prepayment?
3. Where does rent received in advance appear in the Statement of Financial Position?

Answers: see `model-questions/Answer_Key.md`, Chapter 11 section.
