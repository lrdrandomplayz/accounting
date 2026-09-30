# Chapter 11: Accruals and Prepayments

> Excel file for this chapter: `excel/Ch11_Accruals_Prepayments.xlsx`

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

Rent is $1,000 per month. The year ends 31 December 2025. During the year the business paid $11,000. December rent is still owing.

Expense for year = 12 x 1,000 = **$12,000**. Accrual = 12,000 - 11,000 = **$1,000**.

**Rent account**

| Dr | | $ | Cr | | $ |
|---|---|---|---|---|---|
| 2025 | Bank (paid during year) | 11,000 | 2025 Dec 31 | Income statement | 12,000 |
| Dec 31 | Balance c/d (accrued) | 1,000 | | | |
| | | **12,000** | | | **12,000** |
| | | | 2026 Jan 1 | Balance b/d | 1,000 |

The **credit** balance brought down is a **liability**: the business owes $1,000 rent.

## 6. Prepaid expense

### Worked example 11.2

On 1 April 2025 the business paid $3,600 insurance for 12 months. The year ends 31 December 2025.

* Monthly cost = 3,600 / 12 = 300.
* Months used this year (April to December) = 9. Expense = 9 x 300 = **$2,700**.
* Months for next year (January to March) = 3. Prepaid = 3 x 300 = **$900**.

**Insurance account**

| Dr | | $ | Cr | | $ |
|---|---|---|---|---|---|
| 2025 Apr 1 | Bank | 3,600 | 2025 Dec 31 | Income statement | 2,700 |
| | | | Dec 31 | Balance c/d (prepaid) | 900 |
| | | **3,600** | | | **3,600** |
| 2026 Jan 1 | Balance b/d | 900 | | | |

The **debit** balance brought down is an **asset**: the business has already paid for 3 months of cover.

**Tip:** unused stationery or fuel at the year end is treated the same way. Deduct it from the expense and show it as a current asset.

## 7. Opening and closing balances together

### Worked example 11.3

Electricity: owing at the start of the year $200. Paid during the year $2,500. Owing at the end of the year $300.

Expense = 2,500 - 200 + 300 = **$2,600**

**Electricity account**

| Dr | | $ | Cr | | $ |
|---|---|---|---|---|---|
| | Bank | 2,500 | Jan 1 | Balance b/d (owing) | 200 |
| Dec 31 | Balance c/d (owing) | 300 | Dec 31 | Income statement | 2,600 |
| | | **2,800** | | | **2,800** |
| | | | Jan 1 | Balance b/d | 300 |

Why subtract the opening accrual? The $200 belonged to **last** year's expense but was paid **this** year.

## 8. Income adjustments

### Worked example 11.4: income received in advance

A business rents out a room for $500 per month. During 2025 it received $6,500, which includes January 2026 rent.

Income for 2025 = 12 x 500 = **$6,000**. Received in advance = **$500** (current liability).

**Rent received account**

| Dr | | $ | Cr | | $ |
|---|---|---|---|---|---|
| Dec 31 | Income statement | 6,000 | | Bank | 6,500 |
| Dec 31 | Balance c/d (in advance) | 500 | | | |
| | | **6,500** | | | **6,500** |
| | | | Jan 1 | Balance b/d | 500 |

### Worked example 11.5: accrued income

Commission received during the year was $1,800. A further $200 is still due at the year end.

Income = 1,800 + 200 = **$2,000**. The $200 is a current asset (other receivables).

## 9. In the financial statements

**Statement of Profit or Loss (extract)**

| | $ |
|---|---|
| Add other income: Rent received (6,500 - 500) | 6,000 |
| Commission received (1,800 + 200) | 2,000 |
| Less expenses: Rent (11,000 + 1,000) | 12,000 |
| Insurance (3,600 - 900) | 2,700 |
| Electricity (2,500 - 200 + 300) | 2,600 |

**Statement of Financial Position (extract)**

| Current assets | $ |
|---|---|
| Other receivables (900 insurance prepaid + 200 commission due) | 1,100 |

| Current liabilities | $ |
|---|---|
| Other payables (1,000 rent + 300 electricity + 500 rent received in advance) | 1,800 |

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
* Placing the balance c/d on the wrong side. Accrual: debit side c/d, credit side b/d. Prepayment: credit side c/d, debit side b/d.

## 12. Quick summary

* SPL shows what belongs to the year, not what was paid.
* Accrued expense: add to expense, current liability.
* Prepaid expense: subtract from expense, current asset.
* Accrued income: add to income, current asset.
* Income in advance: subtract from income, current liability.

## 13. Check yourself

1. Wages paid $24,000. Wages owing at the year end $800. What is the wages expense?
2. Rates paid $1,500 cover 15 months. Year end is after 12 months of the period. What is the prepayment?
3. Where does rent received in advance appear in the Statement of Financial Position?

Answers: see `model-questions/Answer_Key.md`, Chapter 11 section.
