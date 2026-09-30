"""Build the Excel workbooks for the Senior 1 accounting notes.

Run from the repository root:  python3 tools/build_excel.py
Writes excel/*.xlsx and tools/manifest.json (used by tools/build_docs.py).
Open and save each workbook once (Excel or LibreOffice) so the formulas calculate.
"""
import json
from pathlib import Path

from openpyxl import Workbook

from engine import (D, F_ANS, F_DATA, F_INPUT, INT, NUM, NUM2, PCT, T, TXT, V, Page, sofp_rows, spl_rows)

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "excel"

CONCEPTS = ["Business entity", "Duality", "Money measurement", "Historical cost", "Going concern",
            "Realisation", "Accruals", "Consistency", "Prudence", "Materiality"]
EXP_INC = ["Expense", "Income"]
CAP_REV = ["Capital expenditure", "Revenue expenditure"]
PLACES = ["Current assets (other receivables)", "Current liabilities (other payables)",
          "Non-current assets", "Non-current liabilities"]
QUALITIES = ["Relevance", "Reliability", "Comparability", "Understandability"]

TAB_README, TAB_WORKED, TAB_Q, TAB_A = "1F3864", "ED7D31", "2F5597", "70AD47"


def q(t):
    return f'"{t}"'


# =====================================================================================
# Chapter 8
# =====================================================================================
def q8_8(p):
    p.text("Sara sells goods on credit. Her financial year ends on 31 December 2025. Two credit customers cannot pay:")
    p.bullets(["31 March 2025: Peter, who owes RM350, is declared bankrupt.",
               "30 September 2025: Jane, who owes RM200, leaves the country. Sara decides the debt will never be paid."])
    p.data([("peter", "Debt owed by Peter (RM)", 350), ("jane", "Debt owed by Jane (RM)", 200)])
    p.part("a) Prepare the journal entries to write off both debts. (4)")
    p.journal([("d", "2025 Mar 31", "Irrecoverable debts", "{peter}"), ("c", "", "Peter", "{peter}"),
               ("n", "Being debt written off: customer declared bankrupt"),
               ("d", "Sep 30", "Irrecoverable debts", "{jane}"), ("c", "", "Jane", "{jane}"),
               ("n", "Being debt written off: customer left the country")])
    p.part("b) Prepare the Irrecoverable debts account for the year, showing the transfer to the income statement. (3)")
    p.ledger("Irrecoverable Debts Account", [
        (("2025 Mar 31", "Peter", "GJ1", "{peter}"), ("2025 Dec 31", "Income statement", "GJ2", "{peter}+{jane}")),
        (("Sep 30", "Jane", "GJ1", "{jane}"), None),
        "TOTAL"])
    p.part("c) State the effect of the write-offs on profit and on trade receivables. (2)")
    p.calc([("dp", "Decrease in profit for the year (RM)", "{peter}+{jane}"),
            ("dtr", "Decrease in trade receivables (RM)", "{peter}+{jane}")])


def q8_9(p):
    p.text("At 31 December 2025, trade receivables are RM24,000. This includes RM800 owed by Ali, who has "
           "disappeared. The business writes off Ali's debt. It then creates an allowance for irrecoverable debts "
           "of 5% of trade receivables for the first time.")
    p.data([("tr", "Trade receivables before writing off Ali's debt (RM)", 24000),
            ("ali", "Debt owed by Ali (RM)", 800), ("rate", "Allowance rate", 0.05, PCT)])
    p.part("a) Calculate trade receivables after writing off Ali's debt. (1)")
    p.calc([("tra", "Trade receivables after the write-off (RM)", "{tr}-{ali}")])
    p.part("b) Calculate the allowance for irrecoverable debts. (2)")
    p.calc([("al", "Allowance for irrecoverable debts (RM)", "ROUND({tra}*{rate},2)")])
    p.part("c) Show the Statement of Profit or Loss extract for the year. (2)")
    p.statement(["Statement of Profit or Loss (extract) for the year ended 31 December 2025"], [
        ("Less:", "Expenses", {}, "h"),
        ("", "Irrecoverable debts", {2: "{ali}"}),
        ("", "Allowance for irrecoverable debts (created)", {2: V("{al}", line="under"), 3: "{ali}+{al}"})])
    p.part("d) Show the Statement of Financial Position extract at 31 December 2025. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Trade receivables", {1: "{tra}"}),
        ("Less:", "Allowance for irrecoverable debts", {1: V("-{al}", line="under"), 2: "{tra}-{al}"})], "sofp")
    p.part("e) Name the accounting concept which explains why the allowance is created. (1)")
    p.choice([("c", "Concept", q("Prudence"), CONCEPTS)])


def q8_10(p):
    p.text("Lee's year ends on 30 June 2026. The allowance for irrecoverable debts at 1 July 2025 was RM900. "
           "During the year:")
    p.bullets(["A debt of RM400 owed by Omar was written off.",
               "Lucy paid RM250 by cheque. Her debt was written off in 2024."])
    p.text("Trade receivables at 30 June 2026 were RM26,000 after writing off Omar's debt. The allowance is to be "
           "4% of trade receivables.")
    p.data([("al_open", "Allowance at 1 July 2025 (RM)", 900), ("omar", "Omar's debt written off (RM)", 400),
            ("lucy", "Received from Lucy (RM)", 250), ("tr", "Trade receivables at 30 June 2026 (RM)", 26000),
            ("rate", "Allowance rate", 0.04, PCT)])
    p.part("a) Prepare the journal entries to record the money received from Lucy. (4)")
    p.journal([("d", "", "Lucy", "{lucy}"), ("c", "", "Irrecoverable debts recovered", "{lucy}"),
               ("n", "Being debt written off in 2024 reinstated"),
               ("d", "", "Bank", "{lucy}"), ("c", "", "Lucy", "{lucy}"),
               ("n", "Being cheque received from Lucy")])
    p.part("b) Calculate the new allowance and the change in the allowance. (2)")
    p.calc([("al", "New allowance at 30 June 2026 (RM)", "ROUND({tr}*{rate},2)"),
            ("inc", "Increase in the allowance (RM)", "{al}-{al_open}")])
    p.part("c) Prepare the Allowance for irrecoverable debts account for the year. (4)")
    p.ledger("Allowance for Irrecoverable Debts Account", [
        (("2026 Jun 30", "Balance c/d", "", "{al}"), ("2025 Jul 1", "Balance b/d", "", "{al_open}")),
        (None, ("2026 Jun 30", "Income statement", "GJ1", "{inc}")),
        "TOTAL",
        (None, ("2026 Jul 1", "Balance b/d", "", "{al}"))])
    p.part("d) Show the Statement of Profit or Loss extract for the three items above. (3)")
    p.statement(["Statement of Profit or Loss (extract) for the year ended 30 June 2026"], [
        ("Add:", "Other income", {}, "h"),
        ("", "Irrecoverable debts recovered", {3: "{lucy}"}),
        ("Less:", "Expenses", {}, "h"),
        ("", "Irrecoverable debts", {2: "{omar}"}),
        ("", "Increase in allowance for irrecoverable debts",
         {2: V("{inc}", line="under"), 3: V("-({omar}+{inc})", line="under")}),
        ("", "Net effect on profit", {3: V("{lucy}-({omar}+{inc})", line="total")}, "b")])
    p.part("e) Show the Statement of Financial Position extract at 30 June 2026. (2)")
    p.statement(["Statement of Financial Position (extract) at 30 June 2026"], [
        ("", "Current assets", {}, "h"),
        ("", "Trade receivables", {1: "{tr}"}),
        ("Less:", "Allowance for irrecoverable debts", {1: V("-{al}", line="under"), 2: "{tr}-{al}"})], "sofp")


def q8_11(p):
    p.text("Grace keeps an allowance for irrecoverable debts of 3% of trade receivables. Her year end is "
           "31 December. There was no allowance before 2023.")
    p.data([(None, "2023"), ("b23", "Trade receivables before write-offs (RM)", 40000),
            ("w23", "Debts to be written off (RM)", 0),
            (None, "2024"), ("b24", "Trade receivables before write-offs (RM)", 50000),
            ("w24", "Debts to be written off (RM)", 0),
            (None, "2025"), ("b25", "Trade receivables before write-offs (RM)", 36000),
            ("w25", "Debts to be written off (RM)", 1000),
            (None, "All years"), ("rate", "Allowance rate", 0.03, PCT)], md=True)
    p.part("a) Calculate the allowance required at the end of each year. (3)")
    p.calc([("al23", "Allowance at 31 December 2023 (RM)", "ROUND(({b23}-{w23})*{rate},2)"),
            ("al24", "Allowance at 31 December 2024 (RM)", "ROUND(({b24}-{w24})*{rate},2)"),
            ("al25", "Allowance at 31 December 2025 (RM)", "ROUND(({b25}-{w25})*{rate},2)")])
    p.part("b) State the amount charged or credited to the Statement of Profit or Loss each year, and whether it is "
           "an expense or income. (3)")
    p.calc([("m23", "Amount to the SPL in 2023 (RM)", "{al23}"),
            ("m24", "Amount to the SPL in 2024 (RM)", "ABS({al24}-{al23})"),
            ("m25", "Amount to the SPL in 2025 (RM)", "ABS({al25}-{al24})")])
    p.choice([("t23", "2023: expense or income?", 'IF({al23}>=0,"Expense","Income")', EXP_INC),
              ("t24", "2024: expense or income?", 'IF({al24}>={al23},"Expense","Income")', EXP_INC),
              ("t25", "2025: expense or income?", 'IF({al25}>={al24},"Expense","Income")', EXP_INC)])
    p.part("c) Prepare the Allowance for irrecoverable debts account for the three years. (6)")
    p.ledger("Allowance for Irrecoverable Debts Account", [
        (("2023 Dec 31", "Balance c/d", "", "{al23}"), ("2023 Dec 31", "Income statement", "GJ1", "{al23}")),
        "TOTAL",
        (None, ("2024 Jan 1", "Balance b/d", "", "{al23}")),
        (("2024 Dec 31", "Balance c/d", "", "{al24}"), ("2024 Dec 31", "Income statement", "GJ2", "{al24}-{al23}")),
        "TOTAL",
        (("2025 Dec 31", "Income statement", "GJ3", "{al24}-{al25}"), ("2025 Jan 1", "Balance b/d", "", "{al24}")),
        (("2025 Dec 31", "Balance c/d", "", "{al25}"), None),
        "TOTAL",
        (None, ("2026 Jan 1", "Balance b/d", "", "{al25}"))])
    p.part("d) Show the Statement of Financial Position extract for trade receivables at 31 December 2025. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Trade receivables", {1: "{b25}-{w25}"}),
        ("Less:", "Allowance for irrecoverable debts",
         {1: V("-{al25}", line="under"), 2: "{b25}-{w25}-{al25}"})], "sofp")


def q8_12(p):
    p.text("Hamid's draft profit for the year ended 31 December 2025 is RM18,000. The following items have "
           "not yet been dealt with:")
    p.bullets(["Trade receivables at 31 December 2025 are RM41,500 in the books.",
               "Kim owes RM1,500 and has been declared bankrupt. The debt must be written off.",
               "During the year RM300 was received from a customer whose debt was written off in 2023. "
               "The bookkeeper debited bank and credited trade receivables by mistake.",
               "The allowance for irrecoverable debts at 1 January 2025 was RM2,500. It is to be 5% of trade "
               "receivables at 31 December 2025."])
    p.data([("draft", "Draft profit (RM)", 18000), ("trb", "Trade receivables in the books (RM)", 41500),
            ("kim", "Kim's debt (RM)", 1500),
            ("rec", "Debt recovered, wrongly credited to trade receivables (RM)", 300),
            ("al_open", "Allowance at 1 January 2025 (RM)", 2500), ("rate", "Allowance rate", 0.05, PCT)])
    p.part("a) Calculate the correct trade receivables at 31 December 2025. (3)")
    p.calc([("tr", "Correct trade receivables (RM)", "{trb}-{kim}+{rec}")])
    p.part("b) Calculate the allowance required and the change in the allowance. State whether the change is an "
           "expense or income. (3)")
    p.calc([("al", "Allowance required (RM)", "ROUND({tr}*{rate},2)"),
            ("dec", "Change in the allowance (RM)", "ABS({al}-{al_open})")])
    p.choice([("ei", "Is the change an expense or income?", 'IF({al}>{al_open},"Expense","Income")', EXP_INC)])
    p.part("c) Calculate the corrected profit for the year. (4)")
    p.statement(["Hamid: calculation of corrected profit for the year ended 31 December 2025"], [
        ("", "Draft profit", {3: "{draft}"}),
        ("Less:", "Irrecoverable debt written off (Kim)", {3: "-{kim}"}),
        ("Add:", "Irrecoverable debts recovered", {3: "{rec}"}),
        ("Add:", "Decrease in allowance for irrecoverable debts", {3: V("{al_open}-{al}", line="under")}),
        ("", "Corrected profit", {3: V("{draft}-{kim}+{rec}+{al_open}-{al}", line="total")}, "b")])
    p.part("d) Prepare the Allowance for irrecoverable debts account for the year. (3)")
    p.ledger("Allowance for Irrecoverable Debts Account", [
        (("2025 Dec 31", "Income statement", "GJ1", "{al_open}-{al}"), ("2025 Jan 1", "Balance b/d", "", "{al_open}")),
        (("Dec 31", "Balance c/d", "", "{al}"), None),
        "TOTAL",
        (None, ("2026 Jan 1", "Balance b/d", "", "{al}"))])
    p.part("e) Show the Statement of Financial Position extract for trade receivables. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Trade receivables", {1: "{tr}"}),
        ("Less:", "Allowance for irrecoverable debts", {1: V("-{al}", line="under"), 2: "{tr}-{al}"})], "sofp")
    p.part('f) Hamid says: "I will make the allowance 20% next year, just to be safe." Evaluate this idea. (3)')
    p.written("For: a bigger allowance is more cautious, so assets and profit are less likely to be overstated.\n"
              "Against: the allowance must be a reasonable estimate based on past experience. 20% is far above "
              "Hamid's real losses, so profit and trade receivables would be understated and the statements would "
              "not be reliable. A sudden change also breaks the consistency concept.\n"
              "Conclusion: keep a realistic rate (about 5%) and change it only when experience of bad debts changes. "
              "Prudence means do not overstate; it does not mean understate on purpose.", lines=5)


def worked_8_1(p):
    p.text("Worked examples 8.1 and 8.2 from the notes.")
    p.data([("kofi", "Example 8.1: debt owed by Kofi (RM)", 600),
            ("lina", "Example 8.2: debt recovered from Lina (RM)", 300)], title="Inputs")
    p.part("Example 8.1: Kofi is declared bankrupt on 30 June 2025")
    p.tag("ch08.ex1.journal").journal([("d", "2025 Jun 30", "Irrecoverable debts", "{kofi}"), ("c", "", "Kofi", "{kofi}"),
                                       ("n", "Being debt written off: customer declared bankrupt")])
    p.gap()
    p.tag("ch08.ex1.kofi").ledger("Kofi Account", [
        (("2025 Jun 1", "Balance b/d", "", "{kofi}"), ("2025 Jun 30", "Irrecoverable debts", "GJ1", "{kofi}"))])
    p.gap()
    p.tag("ch08.ex1.id").ledger("Irrecoverable Debts Account", [
        (("2025 Jun 30", "Kofi", "GJ1", "{kofi}"), ("2025 Dec 31", "Income statement", "GJ2", "{kofi}"))])
    p.part("Example 8.2: Lina pays a debt that was written off last year")
    p.tag("ch08.ex2.journal").journal([("d", "", "Lina", "{lina}"), ("c", "", "Irrecoverable debts recovered", "{lina}"),
                                       ("n", "Being debt written off last year reinstated"),
                                       ("d", "", "Bank", "{lina}"), ("c", "", "Lina", "{lina}"),
                                       ("n", "Being cheque received from Lina")])


def worked_8_3(p):
    p.text("Worked example 8.3: an allowance of 5% of trade receivables over three years.")
    yrs = ("1", "2", "3")
    p.tag("ch08.ex3.table").table(["Item", "2023", "2024", "2025"], [
        ("Trade receivables before write-offs", [D(20500, "b1"), D(26800, "b2"), D(18600, "b3")]),
        ("Less: irrecoverable debts written off", [D(500, "w1"), D(800, "w2"), D(600, "w3")]),
        ("Trade receivables after write-offs", [V(f"{{b{y}}}-{{w{y}}}", key=f"t{y}", line="top") for y in yrs]),
        ("Allowance rate", [D(0.05, "r1", PCT), D(0.05, "r2", PCT), D(0.05, "r3", PCT)]),
        ("Step 1: allowance required at year end", [V(f"ROUND({{t{y}}}*{{r{y}}},2)", key=f"a{y}") for y in yrs]),
        ("Less: allowance brought forward", [D(0, "o1"), V("{a1}", key="o2"), V("{a2}", key="o3")]),
        ("Step 2: change (+ increase, - decrease)", [V(f"{{a{y}}}-{{o{y}}}", key=f"c{y}") for y in yrs]),
        ("Step 3: treatment in the SPL", [V(f'IF({{c{y}}}>0,"Expense "&TEXT({{c{y}}},"#,##0"),'
                                            f'IF({{c{y}}}<0,"Income "&TEXT(-{{c{y}}},"#,##0"),"No entry"))', fmt=TXT)
                                          for y in yrs]),
        ("Net trade receivables in the SOFP", [V(f"{{t{y}}}-{{a{y}}}", line="total") for y in yrs]),
    ], 16, 5, checks=False)
    p.gap()
    rows = []
    for y, yr in (("1", "2023"), ("2", "2024"), ("3", "2025")):
        rows += [((f"{yr} Dec 31", "Income statement (decrease)", "GJ1", f"IF({{c{y}}}<0,-{{c{y}}},0)"),
                  (f"{yr} Jan 1", "Balance b/d", "", f"{{o{y}}}")),
                 (("Dec 31", "Balance c/d", "", f"{{a{y}}}"),
                  ("Dec 31", "Income statement (increase)", "GJ1", f"IF({{c{y}}}>0,{{c{y}}},0)")),
                 "TOTAL"]
    rows.append((None, ("2026 Jan 1", "Balance b/d", "", "{a3}")))
    p.tag("ch08.ex3.ledger").ledger("Allowance for Irrecoverable Debts Account", rows)
    p.note("A dash means no amount in that year. Only the change in the allowance goes to the income statement.")
    p.gap()
    p.tag("ch08.ex3.sofp").statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Trade receivables", {1: "{t3}"}),
        ("Less:", "Allowance for irrecoverable debts", {1: V("-{a3}", line="under"), 2: "{t3}-{a3}"})], "sofp")


# =====================================================================================
# Chapter 10
# =====================================================================================
def q10_9(p):
    p.text("On 1 January 2024 a business bought a delivery van for RM18,000. It expects to use the van for 5 years "
           "and then sell it for RM3,000. The straight line method is used. The year end is 31 December.")
    p.data([("cost", "Cost of van (RM)", 18000), ("res", "Residual value (RM)", 3000),
            ("life", "Useful life (years)", 5, INT)])
    p.part("a) Calculate the annual depreciation. (2)")
    p.calc([("dep", "Annual depreciation (RM)", "({cost}-{res})/{life}")])
    p.part("b) Complete the table. (4)")
    p.table(["Year", "Depreciation (RM)", "Accumulated depreciation (RM)", "Carrying amount at end of year (RM)"], [
        ("2024", [V("{dep}"), V("{dep}"), V("{cost}-{dep}")]),
        ("2025", [V("{dep}"), V("2*{dep}"), V("{cost}-2*{dep}")])], 7, 7)
    p.part("c) Show the Statement of Financial Position extract for the van at 31 December 2025. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Non-current assets", {}, "h"),
        ("", "Delivery van", {1: "{cost}", 2: "2*{dep}", 3: "{cost}-2*{dep}"})], "sofp")
    p.part("d) State one cause of depreciation for a delivery van. (1)")
    p.choice([("cause", "Cause of depreciation", q("Wear and tear"),
               ["Wear and tear", "Passage of time", "Obsolescence", "Depletion", "Rising prices"],
               ["wear and tear", "passage of time", "obsolescence"])])


def q10_10(p):
    p.text("A business bought a new machine. It paid the amounts below.")
    p.data([("price", "Purchase price of machine (RM)", 22000), ("deliv", "Delivery to the factory (RM)", 600),
            ("inst", "Installation (RM)", 900), ("ins", "Insurance for the first year (RM)", 400),
            ("rep", "Repair after six months of use (RM)", 250)], md=True)
    p.part("a) Calculate the cost of the machine to be shown as a non-current asset. (2)")
    p.calc([("cost", "Cost of the machine (RM)", "{price}+{deliv}+{inst}")])
    p.part("b) Classify each item as capital expenditure or revenue expenditure. (5)")
    p.choice([("c1", "Purchase price", q("Capital expenditure"), CAP_REV),
              ("c2", "Delivery", q("Capital expenditure"), CAP_REV),
              ("c3", "Installation", q("Capital expenditure"), CAP_REV),
              ("c4", "Insurance for the first year", q("Revenue expenditure"), CAP_REV),
              ("c5", "Repair after six months", q("Revenue expenditure"), CAP_REV)])
    p.part("c) The installation cost was recorded as an expense by mistake. State the effect on profit for the year "
           "and on non-current assets. Ignore depreciation. (2)")
    p.calc([("pu", "Profit for the year is understated by (RM)", "{inst}"),
            ("nu", "Non-current assets are understated by (RM)", "{inst}")])


def q10_11(p):
    p.text("The year end is 31 December.")
    p.bullets(["A machine was bought on 1 April 2025 for RM12,000. Depreciation is 10% per year on cost, straight "
               "line, calculated on a monthly basis.",
               "Loose tools were valued at RM2,000 on 1 January 2025. Tools costing RM800 were bought during 2025. "
               "On 31 December 2025 the loose tools were valued at RM2,100."])
    p.data([("cost", "Cost of machine (RM)", 12000), ("rate", "Depreciation rate per year", 0.1, PCT),
            ("m", "Months owned in 2025", 9, INT), ("lo", "Loose tools at 1 January 2025 (RM)", 2000),
            ("lb", "Loose tools bought in 2025 (RM)", 800), ("lc", "Loose tools at 31 December 2025 (RM)", 2100)])
    p.part("a) Calculate depreciation on the machine for 2025 and for 2026. (3)")
    p.calc([("d25", "Depreciation for 2025 (RM)", "{cost}*{rate}*{m}/12"),
            ("d26", "Depreciation for 2026 (RM)", "{cost}*{rate}")])
    p.part("b) Calculate the carrying amount of the machine at 31 December 2026. (1)")
    p.calc([("ca", "Carrying amount at 31 December 2026 (RM)", "{cost}-{d25}-{d26}")])
    p.part("c) Prepare the journal entry to record the machine depreciation for 2025. (2)")
    p.journal([("d", "2025 Dec 31", "Income statement (depreciation)", "{d25}"),
               ("c", "", "Provision for depreciation of machinery", "{d25}"),
               ("n", "Being depreciation for 9 months at 10% per year on cost")])
    p.part("d) Calculate the depreciation of loose tools for 2025. (2)")
    p.calc([("lt", "Depreciation of loose tools (RM)", "{lo}+{lb}-{lc}")])
    p.part("e) Explain why the revaluation method is used for loose tools. (2)")
    p.written("Loose tools are many small, low-cost items. Keeping a record and a depreciation calculation for "
              "each tool would take too long and cost more than it is worth. So the tools are counted and valued "
              "at each year end instead.")


def q10_12(p):
    p.text("On 1 January 2023 Musa bought a machine for RM40,000. It is depreciated at 25% per year, reducing "
           "balance. The year end is 31 December. On 1 January 2026 the machine was sold for RM15,000 by cheque.")
    p.data([("cost", "Cost of machine (RM)", 40000), ("rate", "Reducing balance rate", 0.25, PCT),
            ("sale", "Sale proceeds on 1 January 2026 (RM)", 15000)])
    p.part("a) Calculate the depreciation for 2023, 2024 and 2025. (3)")
    p.table(["Year", "Carrying amount at start (RM)", "Depreciation (RM)", "Accumulated depreciation (RM)",
             "Carrying amount at end (RM)"], [
        ("2023", [V("{cost}", key="s23"), V("ROUND({s23}*{rate},2)", key="d23"), V("{d23}", key="a23"),
                  V("{s23}-{d23}", key="e23")]),
        ("2024", [V("{e23}", key="s24"), V("ROUND({s24}*{rate},2)", key="d24"), V("{a23}+{d24}", key="a24"),
                  V("{s24}-{d24}", key="e24")]),
        ("2025", [V("{e24}", key="s25"), V("ROUND({s25}*{rate},2)", key="d25"), V("{a24}+{d25}", key="a25"),
                  V("{s25}-{d25}", key="e25")]),
    ], 3, 6)
    p.part("b) Show the carrying amount of the machine at 31 December 2025. (1)")
    p.calc([("ca", "Carrying amount at 31 December 2025 (RM)", "{e25}")])
    p.part("c) Prepare the Machinery disposal account. (5)")
    p.ledger("Machinery Disposal Account", [
        (("2026 Jan 1", "Machinery", "GJ1", "{cost}"), ("2026 Jan 1", "Provision for depreciation", "GJ1", "{a25}")),
        (None, ("Jan 1", "Bank", "CB1", "{sale}")),
        (None, ("Dec 31", "Income statement (loss on disposal)", "GJ2", "{e25}-{sale}")),
        "TOTAL"])
    p.part("d) State where the result of the disposal appears in the financial statements. (1)")
    p.choice([("w", "The loss on disposal is shown as", q("Expense in the SPL"),
               ["Expense in the SPL", "Other income in the SPL", "Current liability", "Deducted from capital"])])


def q10_13(p):
    p.text("Kwame's year ends on 31 December. He depreciates vans at 20% per year on cost, straight line. A full "
           "year's depreciation is charged in the year of purchase and none in the year of sale.")
    p.bullets(["1 January 2023: bought Van A for RM25,000.", "1 July 2024: bought Van B for RM30,000.",
               "1 October 2025: sold Van A for RM12,000, paid by cheque."])
    p.data([("ca", "Cost of Van A (RM)", 25000), ("cb", "Cost of Van B (RM)", 30000),
            ("rate", "Depreciation rate on cost", 0.2, PCT), ("sale", "Sale proceeds of Van A (RM)", 12000)])
    acc_a = "2*{ca}*{rate}"
    p.part("a) Calculate the depreciation charge for 2023, 2024 and 2025. (3)")
    p.calc([("d23", "Depreciation for 2023 (RM)", "{ca}*{rate}"),
            ("d24", "Depreciation for 2024 (RM)", "({ca}+{cb})*{rate}"),
            ("d25", "Depreciation for 2025 (RM)", "{cb}*{rate}")])
    p.part("b) Prepare the Provision for depreciation of vans account for 2023, 2024 and 2025. (6)")
    p.ledger("Provision for Depreciation of Vans Account", [
        (("2023 Dec 31", "Balance c/d", "", "{d23}"), ("2023 Dec 31", "Income statement", "GJ1", "{d23}")),
        "TOTAL",
        (None, ("2024 Jan 1", "Balance b/d", "", "{d23}")),
        (("2024 Dec 31", "Balance c/d", "", "{d23}+{d24}"), ("2024 Dec 31", "Income statement", "GJ2", "{d24}")),
        "TOTAL",
        (("2025 Oct 1", "Disposal (Van A)", "GJ3", acc_a), ("2025 Jan 1", "Balance b/d", "", "{d23}+{d24}")),
        (("Dec 31", "Balance c/d", "", f"{{d23}}+{{d24}}+{{d25}}-{acc_a}"), ("2025 Dec 31", "Income statement", "GJ4", "{d25}")),
        "TOTAL",
        (None, ("2026 Jan 1", "Balance b/d", "", f"{{d23}}+{{d24}}+{{d25}}-{acc_a}"))])
    p.part("c) Prepare the Disposal of van account. (4)")
    p.ledger("Disposal of Van Account", [
        (("2025 Oct 1", "Vans (Van A)", "GJ3", "{ca}"), ("2025 Oct 1", "Provision for depreciation", "GJ3", acc_a)),
        (None, ("Oct 1", "Bank", "CB1", "{sale}")),
        (None, ("Dec 31", "Income statement (loss on disposal)", "GJ4", f"{{ca}}-{acc_a}-{{sale}}")),
        "TOTAL"])
    p.part("d) Show the Statement of Financial Position extract for vans at 31 December 2025. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Non-current assets", {}, "h"),
        ("", "Vans (Van B)", {1: "{cb}", 2: "2*{cb}*{rate}", 3: "{cb}-2*{cb}*{rate}"})], "sofp")
    p.part("e) Kwame wants to change to the reducing balance method in 2026 so his profit looks higher. Advise him, "
           "naming the concept involved. (3)")
    p.choice([("cn", "Concept involved", q("Consistency"), CONCEPTS)])
    p.written("The consistency concept says Kwame should use the same depreciation method every year.\n"
              "He may change only if the new method gives a fairer view of how the vans lose value, and he must "
              "explain the change.\n"
              "Changing just to make profit look higher would mislead users and stop them comparing this year with "
              "earlier years. He should not change.", lines=4)


def worked_10_sl(p):
    p.text("Worked examples 10.1 and 10.4: straight line method. "
           "Annual depreciation = (Cost - Residual value) / Useful life.")
    p.data([("cost", "Cost of machine (RM)", 20000), ("res", "Residual value (RM)", 2000),
            ("life", "Useful life (years, up to 8)", 4, INT)], title="Inputs")
    p.calc([("dep", "Annual depreciation (RM)", "({cost}-{res})/{life}")])
    p.gap()
    rows = []
    for n in range(1, 9):
        blank = f'IF({n}>{{life}},"",'
        acc = "{d1}" if n == 1 else f"{{a{n-1}}}+{{d{n}}}"
        rows.append((str(n), [V(blank + "{dep})", key=f"d{n}"), V(blank + acc + ")", key=f"a{n}"),
                              V(blank + f"{{cost}}-{{a{n}}})", key=f"e{n}")]))
    p.tag("ch10.ex1.table").table(["Year", "Depreciation for year (RM)", "Accumulated depreciation (RM)",
                                   "Carrying amount at end (RM)"], rows, 7, 8, checks=False)
    p.note("The charge is the SAME every year. Rows after the last year of useful life stay blank.")
    p.part("Example 10.4: the ledger accounts for the first two years")
    p.tag("ch10.ex4.machinery").ledger("Machinery Account", [
        (("Year 1 Jan 1", "Bank", "CB1", "{cost}"), ("Year 1 Dec 31", "Balance c/d", "", "{cost}")),
        "TOTAL",
        (("Year 2 Jan 1", "Balance b/d", "", "{cost}"), None)])
    p.gap()
    p.tag("ch10.ex4.provision").ledger("Provision for Depreciation of Machinery Account", [
        (("Year 1 Dec 31", "Balance c/d", "", "{dep}"), ("Year 1 Dec 31", "Income statement", "GJ1", "{dep}")),
        "TOTAL",
        (("Year 2 Dec 31", "Balance c/d", "", "2*{dep}"), ("Year 2 Jan 1", "Balance b/d", "", "{dep}")),
        (None, ("Year 2 Dec 31", "Income statement", "GJ2", "{dep}")),
        "TOTAL",
        (None, ("Year 3 Jan 1", "Balance b/d", "", "2*{dep}"))])
    p.gap()
    p.tag("ch10.ex4.sofp").statement(["Statement of Financial Position (extract) at the end of Year 2"], [
        ("", "Non-current assets", {}, "h"),
        ("", "Machinery", {1: "{cost}", 2: "2*{dep}", 3: "{cost}-2*{dep}"})], "sofp")


def worked_10_rb(p):
    p.text("Worked example 10.2: reducing balance method. Depreciation = carrying amount at start of year x rate.")
    p.data([("cost", "Cost of equipment (RM)", 10000), ("rate", "Depreciation rate", 0.2, PCT)], title="Inputs")
    rows = []
    for n in range(1, 5):
        start = "{cost}" if n == 1 else f"{{e{n-1}}}"
        rows.append((str(n), [V(start, key=f"s{n}"), V(f"ROUND({{s{n}}}*{{rate}},2)", key=f"d{n}"),
                              V("{d1}" if n == 1 else f"{{a{n-1}}}+{{d{n}}}", key=f"a{n}"),
                              V(f"{{s{n}}}-{{d{n}}}", key=f"e{n}")]))
    p.tag("ch10.ex2.table").table(["Year", "Carrying amount at start (RM)", "Depreciation (RM)",
                                   "Accumulated depreciation (RM)", "Carrying amount at end (RM)"],
                                  rows, 3, 7, checks=False)
    p.note("The charge FALLS every year because the carrying amount falls. Common mistake: using cost after Year 1.")


def worked_10_other(p):
    p.text("Straight line vs reducing balance on the same asset, the revaluation method and part-year depreciation.")
    p.data([("cost", "Cost of asset (RM)", 16000), ("rate", "Rate per year", 0.25, PCT)], title="Compare methods: inputs")
    rows = []
    for n in range(1, 5):
        prior = "".join(f"-{{rb{k}}}" for k in range(1, n))
        rows.append((str(n), [V("{cost}*{rate}", key=f"sl{n}"),
                              V(f"{{cost}}-{n}*{{cost}}*{{rate}}"),
                              V(f"ROUND(({{cost}}{prior})*{{rate}},2)", key=f"rb{n}"),
                              V("{cost}" + "".join(f"-{{rb{k}}}" for k in range(1, n + 1)))]))
    p.table(["Year", "Straight line depreciation (RM)", "Straight line carrying amount (RM)",
             "Reducing balance depreciation (RM)", "Reducing balance carrying amount (RM)"], rows, 3, 7, checks=False)
    p.note("Straight line applies the rate to COST. Reducing balance applies the rate to the carrying amount.")
    p.part("Worked example 10.3: revaluation method (loose tools)")
    p.data([("lo", "Loose tools at start of year (RM)", 1200), ("lb", "Tools bought during the year (RM)", 500),
            ("lc", "Loose tools valued at end of year (RM)", 1100)], title="Inputs")
    p.calc([("ldep", "Depreciation = opening + purchases - closing (RM)", "{lo}+{lb}-{lc}")])
    p.part("Section 9: asset bought part way through the year (monthly basis)")
    p.data([("pc", "Cost of machine (RM)", 12000), ("pr", "Rate per year on cost", 0.1, PCT),
            ("pm", "Months owned in the first year", 9, INT)], title="Inputs")
    p.calc([("full", "Full year depreciation (RM)", "{pc}*{pr}"),
            ("part", "Depreciation for the months owned (RM)", "{pc}*{pr}*{pm}/12")])


def worked_10_disposal(p):
    p.text("Worked example 10.5: disposal of a van. Profit or loss on disposal = sale proceeds - carrying amount.")
    p.data([("cost", "Cost of van (RM)", 30000), ("acc", "Accumulated depreciation to date of sale (RM)", 18000),
            ("sale", "Sale proceeds (RM)", 10000)], title="Inputs")
    p.calc([("ca", "Carrying amount at date of sale (RM)", "{cost}-{acc}"),
            ("pl", "Profit (+) or loss (-) on disposal (RM)", "{sale}-{ca}"),
            ("txt", "Treatment", 'IF({pl}>0,"Profit: other income in the SPL",IF({pl}<0,"Loss: expense in the SPL",'
                                 '"No profit or loss"))', TXT)])
    p.gap()
    p.tag("ch10.ex5.ledger").ledger("Disposal of Van Account", [
        (("Date of sale", "Van (cost)", "GJ1", "{cost}"), ("Date of sale", "Provision for depreciation", "GJ1", "{acc}")),
        (("Year end", "Income statement (profit on disposal)", "GJ2", "MAX({pl},0)"), ("Date of sale", "Bank", "CB1", "{sale}")),
        (None, ("Year end", "Income statement (loss on disposal)", "GJ2", "MAX(-{pl},0)")),
        "TOTAL"])
    p.note("A dash means no amount. A profit balances the account on the debit side; a loss on the credit side.")


# =====================================================================================
# Chapter 11
# =====================================================================================
def q11_8(p):
    p.text("A business's year ends on 31 December 2025.")
    p.bullets(["Rent is RM800 per month. During the year the business paid RM8,800 rent.",
               "Insurance of RM1,800 was paid during the year. Of this, RM300 is for 2026."])
    p.data([("rm", "Rent per month (RM)", 800), ("rp", "Rent paid during the year (RM)", 8800),
            ("ip", "Insurance paid during the year (RM)", 1800), ("ipre", "Insurance for 2026 (RM)", 300)])
    p.part("a) Calculate the rent expense for the year and the amount of rent owing. (2)")
    p.calc([("re", "Rent expense for the year (RM)", "12*{rm}"),
            ("ro", "Rent owing at 31 December 2025 (RM)", "12*{rm}-{rp}")])
    p.part("b) Calculate the insurance expense for the year. (1)")
    p.calc([("ie", "Insurance expense for the year (RM)", "{ip}-{ipre}")])
    p.part("c) State where the rent owing and the insurance prepaid appear in the Statement of Financial Position. (2)")
    p.choice([("w1", "Rent owing", q("Current liabilities (other payables)"), PLACES),
              ("w2", "Insurance prepaid", q("Current assets (other receivables)"), PLACES)])
    p.part("d) Name the concept which requires these adjustments. (1)")
    p.choice([("cn", "Concept", q("Accruals"), CONCEPTS)])


def q11_9(p):
    p.text("A business's year ends on 31 December 2025.")
    p.bullets(["Electricity paid: March RM600, June RM550, September RM500. The bill for October to December, "
               "RM580, was paid in January 2026.",
               "On 1 July 2025 the business paid RM2,400 insurance for 12 months."])
    p.data([("e1", "Electricity paid in March (RM)", 600), ("e2", "Electricity paid in June (RM)", 550),
            ("e3", "Electricity paid in September (RM)", 500), ("eo", "Electricity owing at year end (RM)", 580),
            ("ins", "Insurance paid on 1 July 2025 for 12 months (RM)", 2400),
            ("mu", "Months of insurance used in 2025", 6, INT)])
    ins_pre = "{ins}-{ins}/12*{mu}"
    p.part("a) Prepare the Electricity account for the year. Show the balance carried down and brought down. (4)")
    p.ledger("Electricity Account", [
        (("2025 Mar", "Bank", "CB1", "{e1}"), ("2025 Dec 31", "Income statement", "GJ1", "{e1}+{e2}+{e3}+{eo}")),
        (("Jun", "Bank", "CB1", "{e2}"), None),
        (("Sep", "Bank", "CB1", "{e3}"), None),
        (("Dec 31", "Balance c/d", "", "{eo}"), None),
        "TOTAL",
        (None, ("2026 Jan 1", "Balance b/d", "", "{eo}"))])
    p.part("b) Prepare the Insurance account for the year. Show the balance carried down and brought down. (4)")
    p.ledger("Insurance Account", [
        (("2025 Jul 1", "Bank", "CB1", "{ins}"), ("2025 Dec 31", "Income statement", "GJ1", "{ins}/12*{mu}")),
        (None, ("Dec 31", "Balance c/d", "", ins_pre)),
        "TOTAL",
        (("2026 Jan 1", "Balance b/d", "", ins_pre), None)])
    p.part("c) Show the amounts in the Statement of Financial Position at 31 December 2025. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Other receivables (insurance prepaid)", {2: ins_pre}),
        ("Less:", "Current liabilities", {}, "h"),
        ("", "Other payables (electricity owing)", {2: "{eo}"})], "sofp")


def q11_10(p):
    p.text("Sam lets part of his premises for RM600 per month. His year ends on 31 December 2025.")
    p.bullets(["Rent received during 2025 was RM7,800. This includes the rent for January 2026.",
               "Commission received during 2025 was RM2,100. A further RM350 of commission for 2025 is still due."])
    p.data([("rm", "Rent per month (RM)", 600), ("rr", "Rent received during 2025 (RM)", 7800),
            ("cr", "Commission received during 2025 (RM)", 2100), ("cd", "Commission still due (RM)", 350)])
    adv = "{rr}-12*{rm}"
    p.part("a) Calculate the rent income for 2025 and the amount received in advance. (2)")
    p.calc([("ri", "Rent income for 2025 (RM)", "12*{rm}"), ("ra", "Rent received in advance (RM)", adv)])
    p.part("b) Prepare the Rent receivable account for the year. (4)")
    p.ledger("Rent Receivable Account", [
        (("2025 Dec 31", "Income statement", "GJ1", "12*{rm}"), ("2025", "Bank (received during the year)", "CB1", "{rr}")),
        (("Dec 31", "Balance c/d", "", adv), None),
        "TOTAL",
        (None, ("2026 Jan 1", "Balance b/d", "", adv))])
    p.part("c) Calculate the commission income for 2025. (1)")
    p.calc([("ci", "Commission income for 2025 (RM)", "{cr}+{cd}")])
    p.part("d) Show the other receivables and other payables in the Statement of Financial Position. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Other receivables (commission due)", {2: "{cd}"}),
        ("Less:", "Current liabilities", {}, "h"),
        ("", "Other payables (rent received in advance)", {2: adv})], "sofp")
    p.part("e) State the effect on profit if Sam forgot to adjust for the rent received in advance. (1)")
    p.calc([("ov", "Profit would be overstated by (RM)", adv)])


def q11_11(p):
    p.text("Zara's year ends on 31 March 2026. The following information is available.")
    p.data([(None, "Wages"), ("wo", "Owing at 1 April 2025 (RM)", 300), ("wp", "Paid during the year (RM)", 18000),
            ("wc", "Owing at 31 March 2026 (RM)", 450),
            (None, "Rates"), ("ro", "Prepaid at 1 April 2025 (RM)", 200), ("rp", "Paid during the year (RM)", 1600),
            ("rc", "Prepaid at 31 March 2026 (RM)", 250),
            (None, "Commission receivable"), ("cr", "Received during the year (RM)", 3000),
            ("cd", "Still due at 31 March 2026 (RM)", 250)], md=True)
    p.part("a) Calculate the amount for the Statement of Profit or Loss for wages, rates and commission receivable. (3)")
    p.calc([("we", "Wages expense (RM)", "{wp}-{wo}+{wc}"), ("re", "Rates expense (RM)", "{rp}+{ro}-{rc}"),
            ("ce", "Commission receivable income (RM)", "{cr}+{cd}")])
    p.part("b) Prepare the Wages account for the year, showing the balance carried down. (5)")
    p.ledger("Wages Account", [
        (("2025", "Bank (paid during the year)", "CB1", "{wp}"), ("2025 Apr 1", "Balance b/d", "", "{wo}")),
        (("2026 Mar 31", "Balance c/d", "", "{wc}"), ("2026 Mar 31", "Income statement", "GJ1", "{wp}-{wo}+{wc}")),
        "TOTAL",
        (None, ("2026 Apr 1", "Balance b/d", "", "{wc}"))])
    p.part("c) Calculate the total of other receivables and other payables in the Statement of Financial Position "
           "at 31 March 2026. (2)")
    p.calc([("orc", "Other receivables: rates prepaid + commission due (RM)", "{rc}+{cd}"),
            ("opy", "Other payables: wages owing (RM)", "{wc}")])


def q11_12(p):
    p.text("Tariq calculated a draft profit of RM32,000 for the year ended 31 December 2025. He used only the cash "
           "paid and received. No adjustments have been made.")
    p.bullets(["Insurance: RM750 was prepaid at 1 January 2025 (covering January to March 2025). On 1 April 2025 he "
               "paid RM3,600 for the year to 31 March 2026.",
               "Rent: RM1,000 per month until 30 June 2025, then RM1,200 per month from 1 July 2025. He paid "
               "RM13,000 during the year. Nothing was owing or prepaid at 1 January 2025.",
               "Wages: RM21,000 was paid during 2025. This includes RM400 owing at 1 January 2025. RM650 is owing "
               "at 31 December 2025.",
               "Rent received: RM4,500 was received in 2025. This includes RM300 owing from 2024 and RM500 for "
               "January 2026.",
               "Stationery: unused stationery at 31 December 2025 cost RM120."])
    p.data([("draft", "Draft profit (RM)", 32000), ("io", "Insurance prepaid at 1 January 2025 (RM)", 750),
            ("ip", "Insurance paid on 1 April 2025 for 12 months (RM)", 3600),
            ("im", "Months of that policy used in 2025", 9, INT),
            ("r1", "Rent per month, January to June (RM)", 1000), ("r2", "Rent per month, July to December (RM)", 1200),
            ("rp", "Rent paid during 2025 (RM)", 13000), ("wp", "Wages paid during 2025 (RM)", 21000),
            ("wo", "Wages owing at 1 January 2025 (RM)", 400), ("wc", "Wages owing at 31 December 2025 (RM)", 650),
            ("rr", "Rent received in 2025 (RM)", 4500), ("rro", "Of which: owing from 2024 (RM)", 300),
            ("rra", "Of which: for January 2026 (RM)", 500), ("st", "Unused stationery (RM)", 120)])
    pre = "{ip}*(12-{im})/12"
    p.part("a) Calculate the correct amount for the Statement of Profit or Loss for insurance, rent, wages and rent "
           "received. (5)")
    p.calc([("ins", "Insurance expense (RM)", "{io}+{ip}*{im}/12"), ("rent", "Rent expense (RM)", "6*{r1}+6*{r2}"),
            ("wag", "Wages expense (RM)", "{wp}-{wo}+{wc}"), ("rin", "Rent received: income for 2025 (RM)", "{rr}-{rro}-{rra}")])
    p.part("b) Calculate the corrected profit for the year. Show each adjustment. (5)")
    p.statement(["Tariq: calculation of corrected profit for the year ended 31 December 2025"], [
        ("", "Draft profit", {3: "{draft}"}),
        ("Add:", "Insurance: cash paid less expense", {3: "{ip}-{ins}"}),
        ("Less:", "Rent: expense less cash paid", {3: "-({rent}-{rp})"}),
        ("Less:", "Wages: expense less cash paid", {3: "-({wag}-{wp})"}),
        ("Less:", "Rent received: cash less income", {3: "-({rr}-{rin})"}),
        ("Add:", "Unused stationery", {3: V("{st}", line="under")}),
        ("", "Corrected profit",
         {3: V("{draft}+({ip}-{ins})-({rent}-{rp})-({wag}-{wp})-({rr}-{rin})+{st}", line="total")}, "b")])
    p.part("c) Calculate other receivables and other payables for the Statement of Financial Position. (4)")
    p.calc([("orc", "Other receivables: insurance prepaid + unused stationery (RM)", f"{pre}+{{st}}"),
            ("opy", "Other payables: rent owing + wages owing + rent received in advance (RM)", "({rent}-{rp})+{wc}+{rra}")])
    p.part("d) Prepare the Insurance account for the year. (4)")
    p.ledger("Insurance Account", [
        (("2025 Jan 1", "Balance b/d", "", "{io}"), ("2025 Dec 31", "Income statement", "GJ1", "{ins}")),
        (("Apr 1", "Bank", "CB1", "{ip}"), ("Dec 31", "Balance c/d", "", pre)),
        "TOTAL",
        (("2026 Jan 1", "Balance b/d", "", pre), None)])


def expense_ledger(p, tag, name, k):
    p.tag(tag).ledger(name, [
        (("2025 Jan 1", "Balance b/d (prepaid)", "", f"{{{k}_op}}"), ("2025 Jan 1", "Balance b/d (owing)", "", f"{{{k}_oa}}")),
        (("2025", "Bank (paid during the year)", "CB1", f"{{{k}_paid}}"), ("Dec 31", "Income statement", "GJ1", f"{{{k}_exp}}")),
        (("Dec 31", "Balance c/d (owing)", "", f"{{{k}_ca}}"), ("Dec 31", "Balance c/d (prepaid)", "", f"{{{k}_cp}}")),
        "TOTAL",
        (("2026 Jan 1", "Balance b/d (prepaid)", "", f"{{{k}_cp}}"), ("2026 Jan 1", "Balance b/d (owing)", "", f"{{{k}_ca}}"))])


def worked_11_exp(p):
    p.text("Worked examples 11.1 to 11.3. Expense = paid - opening accrual + opening prepayment + closing accrual "
           "- closing prepayment.")
    data = (("rent", "Rent", 11000, 0, 0, 1000, 0), ("ins", "Insurance", 3600, 0, 0, 0, 900),
            ("elec", "Electricity", 2500, 200, 0, 300, 0))
    rows = []
    for k, label, paid, oa, op, ca, cp in data:
        rows.append((label, [D(paid, f"{k}_paid"), D(oa, f"{k}_oa"), D(op, f"{k}_op"), D(ca, f"{k}_ca"),
                             D(cp, f"{k}_cp"),
                             V(f"{{{k}_paid}}-{{{k}_oa}}+{{{k}_op}}+{{{k}_ca}}-{{{k}_cp}}", key=f"{k}_exp"),
                             V(f"{{{k}_ca}}"), V(f"{{{k}_cp}}")]))
    p.table(["Expense", "Paid in year", "Opening accrual (owing)", "Opening prepayment", "Closing accrual (owing)",
             "Closing prepayment", "Expense for the SPL", "SOFP: other payables", "SOFP: other receivables"],
            rows, 7, 3, checks=False)
    p.note("All amounts in RM. Accrued expense: ADD, current liability. Prepaid expense: SUBTRACT, current asset.")
    p.part("Example 11.1: Rent account (RM1,000 per month; RM11,000 paid)")
    expense_ledger(p, "ch11.ex1.ledger", "Rent Account", "rent")
    p.part("Example 11.2: Insurance account (RM3,600 paid on 1 April for 12 months)")
    expense_ledger(p, "ch11.ex2.ledger", "Insurance Account", "ins")
    p.part("Example 11.3: Electricity account (opening and closing balances)")
    expense_ledger(p, "ch11.ex3.ledger", "Electricity Account", "elec")
    p.note("A dash means there is no amount for that line in this example.")


def worked_11_inc(p):
    p.text("Worked examples 11.4 and 11.5. Income = received - opening accrued income + opening income in advance "
           "+ closing accrued income - closing income in advance.")
    data = (("rr", "Rent received", 6500, 0, 0, 0, 500), ("cm", "Commission received", 1800, 0, 0, 200, 0))
    rows = []
    for k, label, rec, oa, oi, ca, ci in data:
        rows.append((label, [D(rec, f"{k}_rec"), D(oa, f"{k}_oa"), D(oi, f"{k}_oi"), D(ca, f"{k}_ca"),
                             D(ci, f"{k}_ci"),
                             V(f"{{{k}_rec}}-{{{k}_oa}}+{{{k}_oi}}+{{{k}_ca}}-{{{k}_ci}}", key=f"{k}_inc"),
                             V(f"{{{k}_ca}}"), V(f"{{{k}_ci}}")]))
    p.table(["Income", "Received in year", "Opening accrued income (due)", "Opening income in advance",
             "Closing accrued income (due)", "Closing income in advance", "Income for the SPL",
             "SOFP: other receivables", "SOFP: other payables"], rows, 7, 3, checks=False)
    p.note("All amounts in RM. Accrued income: ADD, current asset. Income in advance: SUBTRACT, current liability.")
    p.part("Example 11.4: Rent received account")
    k = "rr"
    p.tag("ch11.ex4.ledger").ledger("Rent Received Account", [
        (("2025 Jan 1", "Balance b/d (due)", "", f"{{{k}_oa}}"), ("2025 Jan 1", "Balance b/d (in advance)", "", f"{{{k}_oi}}")),
        (("Dec 31", "Income statement", "GJ1", f"{{{k}_inc}}"), ("2025", "Bank (received during the year)", "CB1", f"{{{k}_rec}}")),
        (("Dec 31", "Balance c/d (in advance)", "", f"{{{k}_ci}}"), ("Dec 31", "Balance c/d (due)", "", f"{{{k}_ca}}")),
        "TOTAL",
        (("2026 Jan 1", "Balance b/d (due)", "", f"{{{k}_ca}}"), ("2026 Jan 1", "Balance b/d (in advance)", "", f"{{{k}_ci}}"))])
    p.part("Splitting a payment by months (example 11.2)")
    p.data([("ap", "Amount paid (RM)", 3600), ("mc", "Months the payment covers", 12, INT),
            ("mt", "Months falling in this financial year", 9, INT)], title="Inputs")
    p.calc([("pm", "Cost per month (RM)", "{ap}/{mc}", NUM2), ("te", "Expense for this year (RM)", "{ap}/{mc}*{mt}", NUM2),
            ("tp", "Prepayment carried to next year (RM)", "{ap}-{ap}/{mc}*{mt}", NUM2)])


# =====================================================================================
# Chapter 12
# =====================================================================================
def q12_15(p):
    p.text("Name the concept described by each statement. (1 mark each)")
    p.part("a) to e) Choose the concept for each statement. (5)")
    p.choice([("a", "a) The business is treated as separate from its owner.", q("Business entity"), CONCEPTS),
              ("b", "b) Assets are recorded at the price paid for them.", q("Historical cost"), CONCEPTS),
              ("c", "c) Every transaction affects two accounts.", q("Duality"), CONCEPTS),
              ("d", "d) Only items with a money value are recorded.", q("Money measurement"), CONCEPTS),
              ("e", "e) The same methods are used from one year to the next.", q("Consistency"), CONCEPTS)])
    p.part("f) Name one quality that makes accounting information useful. (1)")
    p.choice([("f", "One quality of useful information",
               q("Any one: Relevance, Reliability, Comparability, Understandability"),
               QUALITIES, [x.lower() for x in QUALITIES])])


def q12_16(p):
    p.text("For each situation, name the concept which applies and explain how it applies. (2 marks each)")
    parts = [
        ("a", "a) The owner paid for a family holiday with a business cheque. It was recorded as drawings.",
         "Business entity", "The holiday is a private expense of the owner. The business is separate from the owner, "
         "so the payment is drawings, not a business expense."),
        ("b", "b) Goods were sold on credit on 29 December 2025 and paid for on 5 January 2026. The sale was recorded "
              "in 2025.", "Realisation", "Revenue is recorded when the goods pass to the customer (29 December 2025), "
         "not when the cash is received."),
        ("c", "c) A waste bin costing RM8 was recorded as an expense, although it will last for years.",
         "Materiality", "RM8 is too small to affect any user's decision. Treating it as a non-current asset and "
         "depreciating it is not worth the effort."),
        ("d", "d) Premises bought for RM80,000 are now worth RM120,000. They are still shown at RM80,000.",
         "Historical cost", "Assets are recorded at the price paid. Cost is objective and backed by documents; the "
         "RM120,000 is only an estimate."),
    ]
    for key, label, concept, model in parts:
        p.part(label)
        p.choice([(key, "Concept", q(concept), CONCEPTS)])
        p.written(model, lines=2)


def q12_17(p):
    p.text("At the year end a trader has four items of inventory.")
    items = (("A", 1200, 1800, 100), ("B", 900, 950, 200), ("C", 2000, 2500, 300), ("D (damaged)", 600, 250, 0))
    p.table(["Item", "Cost (RM)", "Expected selling price (RM)", "Costs to sell or repair (RM)"],
            [(name, [D(c, f"c{i}"), D(s, f"s{i}"), D(k, f"k{i}")]) for i, (name, c, s, k) in enumerate(items)],
            7, 8, checks=False)
    p.part("a) Calculate the value of each item and the total value of closing inventory. (4)")
    rows = [(name, [V(f"{{s{i}}}-{{k{i}}}", key=f"n{i}"), V(f"MIN({{c{i}}},{{n{i}}})", key=f"v{i}")])
            for i, (name, *_rest) in enumerate(items)]
    rows.append(("Total closing inventory", [None, V("{v0}+{v1}+{v2}+{v3}", key="tv", line="total")]))
    p.table(["Item", "Net realisable value (RM)", "Value used for inventory (RM)"], rows, 7, 11,
            total_rows=("Total closing inventory",))
    p.part("b) The trader had valued all inventory at cost. State the effect of your correct valuation on gross "
           "profit. (2)")
    p.calc([("dg", "Gross profit falls by (RM)", "({c0}+{c1}+{c2}+{c3})-{tv}")])
    p.part("c) Name the concept used and explain why it applies. (2)")
    p.choice([("cn", "Concept", q("Prudence"), CONCEPTS)])
    p.written("Prudence: inventory is valued at the lower of cost and net realisable value. The expected losses on "
              "items B and D are recorded now, so assets and profit are not overstated.", lines=2)


def q12_18(p):
    p.text("Mira made the decisions below. For each one, name the concept she has not followed, explain why, and "
           "state the correct treatment. (3 marks each)")
    parts = [
        ("a", "a) She changed her depreciation method from straight line to reducing balance this year. She plans to "
              "change back next year to make her profit look steady.", "Consistency", None,
         "Why: switching methods to smooth profit makes the years impossible to compare and misleads users.\n"
         "Correct treatment: choose the method that suits the asset and use it every year. Change only for a good "
         "reason and explain the change."),
        ("b", "b) She recorded the repayment of her personal car loan as a business expense.", "Business entity", None,
         "Why: the car loan is the owner's private debt, not the business's.\n"
         "Correct treatment: record the payment as drawings."),
        ("c", "c) She included in 2025 revenue an order received on 28 December 2025 for goods to be delivered in "
              "February 2026.", "Realisation", None,
         "Why: the goods have not passed to the customer, so the revenue has not been earned.\n"
         "Correct treatment: record the sale in February 2026 when the goods are delivered."),
        ("d", "d) She did not record December electricity of RM400 because the bill had not been paid.", "Accruals", None,
         "Why: the electricity was used in 2025, so it is a 2025 expense even though it is unpaid.\n"
         "Correct treatment: add RM400 to the electricity expense and show RM400 in other payables (current "
         "liabilities)."),
        ("e", "e) She increased the value of premises from cost of RM100,000 to a market value of RM150,000 and added "
              "the RM50,000 to profit.", "Historical cost", ["historical cost", "prudence", "realisation"],
         "Why: premises must stay at cost. The RM50,000 gain has not been earned by a sale (prudence and "
         "realisation also apply).\nCorrect treatment: keep premises at RM100,000 and do not add RM50,000 to profit."),
    ]
    for key, label, concept, accept, model in parts:
        p.part(label)
        row = (key, "Concept not followed", q(concept), CONCEPTS) + ((accept,) if accept else ())
        p.choice([row])
        p.written(model, lines=2)


def q12_19(p):
    p.text("A business is thinking of closing down. The following information is available.")
    rows = [("Premises (carrying amount)", 80000, 95000, "pr"), ("Equipment (carrying amount)", 20000, 6000, "eq"),
            ("Inventory (at cost)", 15000, 9000, "inv"), ("Trade receivables", 12000, 10500, "tr"),
            ("Allowance for irrecoverable debts", 600, None, "al"), ("Bank", 3000, 3000, "bk"),
            ("Trade payables", 8000, 8000, "tp"), ("Loan", 25000, 25000, "ln")]
    p.table(["Item", "Value in the books (RM)", "Value if the business closes (RM)"],
            [(label, [D(b, f"{k}_b"), D(c, f"{k}_c") if c is not None else T("-")]) for label, b, c, k in rows],
            13, 9, checks=False)
    p.text("If the business closes, RM10,500 is the amount of trade receivables expected to be collected.")
    p.part("a) Calculate net assets on the going concern basis (book values). (3)")
    p.calc([("ga", "Total assets, going concern basis (RM)", "{pr_b}+{eq_b}+{inv_b}+({tr_b}-{al_b})+{bk_b}"),
            ("gn", "Net assets, going concern basis (RM)", "{ga}-{tp_b}-{ln_b}")])
    p.part("b) Calculate net assets if the business closes. (3)")
    p.calc([("ca", "Total assets if the business closes (RM)", "{pr_c}+{eq_c}+{inv_c}+{tr_c}+{bk_c}"),
            ("cn", "Net assets if the business closes (RM)", "{ca}-{tp_c}-{ln_c}")])
    p.part("c) Explain why the two figures are different. Name the concept involved. (3)")
    p.choice([("gc", "Concept", q("Going concern"), CONCEPTS)])
    p.written("Going concern assumes the business will continue, so assets are shown at cost less depreciation "
              "because they will be used, not sold. If the business closes, assets must be valued at what they "
              "would raise now. Equipment and inventory sold quickly fetch much less, so net assets fall by RM5,900.",
              lines=3)
    p.part("d) While the business continues trading, premises must stay at RM80,000 even though they are worth "
           "RM95,000. Name two concepts that support this. (2)")
    acc = ["historical cost", "prudence", "realisation"]
    p.choice([("d1", "First concept", q("Historical cost"), CONCEPTS, acc),
              ("d2", "Second concept", q("Prudence"), CONCEPTS, acc)])
    p.part("e) Name two users of the financial statements who would be interested in these figures, and give a "
           "reason for each. (2)")
    p.written("Any two, with a reason:\nOwner: to decide whether to close or keep trading.\nBank or lender: to see "
              "whether the RM25,000 loan will be repaid.\nSuppliers (trade payables): to see whether they will be "
              "paid.\nEmployees: to judge whether their jobs are safe.", lines=4)


QUIZ = [
    ("12.1 The owner's private car is not in the business accounts.", "Business entity"),
    ("12.2 Land is recorded at the price paid, not its current value.", "Historical cost"),
    ("12.3 The same depreciation method is used every year.", "Consistency"),
    ("12.4 A RM5 calculator is recorded as an expense.", "Materiality"),
    ("12.5 Revenue is recorded when goods are delivered.", "Realisation"),
    ("12.6 Staff skill is not recorded in the accounts.", "Money measurement"),
    ("12.7 Every transaction has a debit and a credit.", "Duality"),
    ("12.8 Assets are shown at cost less depreciation, not closing-down prices.", "Going concern"),
    ("12.9 December electricity paid in January is December's expense.", "Accruals"),
    ("12.10 An allowance for irrecoverable debts is created.", "Prudence"),
    ("Extra: Inventory is valued at the lower of cost and net realisable value.", "Prudence"),
    ("Extra: Goods taken by the owner are recorded as drawings.", "Business entity"),
    ("Extra: A credit sale on 30 December, paid on 10 January, is this year's revenue.", "Realisation"),
    ("Extra: Insurance prepaid is deducted from this year's expense.", "Accruals"),
    ("Extra: Buying equipment by cheque increases equipment and decreases bank.", "Duality"),
]


def quiz(p):
    p.text("Section A of the Chapter 12 model questions, plus five extra situations. Pick the concept for each.")
    p.choice([(f"z{i}", s, q(a), CONCEPTS) for i, (s, a) in enumerate(QUIZ)])


def worked_12(p):
    p.text("The ten accounting concepts and the inventory valuation rule.")
    rows = [
        ("Business entity", "The business is separate from its owner.",
         "Owner's private house is not recorded; goods taken are drawings."),
        ("Duality", "Every transaction has two equal effects: a debit and a credit.",
         "Buy equipment by cheque: equipment up, bank down."),
        ("Money measurement", "Only record items with a money value.", "Staff skill and customer loyalty are not recorded."),
        ("Historical cost", "Record assets at their original cost.", "Land bought for RM50,000 stays at RM50,000."),
        ("Going concern", "Assume the business will continue trading.",
         "Assets at cost less depreciation, not closing-down value."),
        ("Realisation", "Record revenue when it is earned (goods pass to the customer).",
         "A credit sale on 28 December is this year's revenue."),
        ("Accruals", "Match income and expenses to the period they belong to.", "Accrued electricity; depreciation."),
        ("Consistency", "Use the same methods every year.", "Keep the same depreciation method."),
        ("Prudence", "Do not overstate assets or profit.",
         "Allowance for irrecoverable debts; inventory at lower of cost and NRV."),
        ("Materiality", "Small items can be treated simply.", "A RM5 stapler is an expense, not a non-current asset."),
    ]
    p.table(["Concept", "Meaning", "Example"], [(c, [T(m), T(e)]) for c, m, e in rows], 7, 12, checks=False)
    p.part("Inventory: lower of cost and net realisable value (prudence)")
    p.data([("cost", "Cost of goods (RM)", 800), ("sp", "Expected selling price (RM)", 700),
            ("cs", "Costs to sell or repair (RM)", 50)], title="Inputs")
    p.calc([("nrv", "Net realisable value (RM)", "{sp}-{cs}"),
            ("val", "Value for inventory: lower of cost and NRV (RM)", "MIN({cost},{nrv})")])


# =====================================================================================
# Financial statements
# =====================================================================================
AMINA = [("rev", "Revenue", None, 110000), ("pur", "Purchases", 62000, None), ("sr", "Sales returns", 1500, None),
         ("pr", "Purchases returns", None, 1200), ("ci", "Carriage inwards", 800, None),
         ("co", "Carriage outwards", 600, None), ("oi", "Inventory at 1 January 2025", 8000, None),
         ("wag", "Wages", 14000, None), ("rent", "Rent", 5500, None), ("ins", "Insurance", 1800, None),
         ("gen", "General expenses", 2100, None), ("da", "Discount allowed", 400, None),
         ("dr", "Discount received", None, 500), ("id", "Irrecoverable debts", 700, None),
         ("prem", "Premises at cost", 60000, None), ("eq", "Equipment at cost", 20000, None),
         ("eqd", "Provision for depreciation: equipment", None, 6000), ("tr", "Trade receivables", 12000, None),
         ("afd", "Allowance for irrecoverable debts", None, 400), ("tp", "Trade payables", None, 7500),
         ("bank", "Bank", 4300, None), ("cash", "Cash", 200, None), ("loan", "Loan (repayable 2030)", None, 10000),
         ("cap", "Capital", None, 67300), ("draw", "Drawings", 9000, None)]


def worked_fs(p):
    p.text("Amina Stores: the full worked example from the notes. Change any blue figure and the statements update.")
    p.tag("fs.ex.tb").tb(AMINA, "Amina Stores: Trial Balance at 31 December 2025")
    p.data([("close", "1. Closing inventory (RM)", 9500), ("rent_ow", "2. Rent owing (RM)", 500),
            ("ins_pre", "3. Insurance prepaid (RM)", 300), ("al_rate", "4. Allowance: % of trade receivables", 0.05, PCT),
            ("eq_rate", "5. Equipment depreciation: % of cost per year", 0.1, PCT),
            ("int_rate", "6. Loan interest per year (unpaid)", 0.05, PCT)],
           title="Additional information at 31 December 2025")
    allow = "ROUND({tr}*{al_rate},2)"
    p.tag("fs.ex.spl").statement(["Amina Stores", "Statement of Profit or Loss for the year ended 31 December 2025"], spl_rows(
        "{rev}", "{oi}", "{pur}", "{close}", sr="{sr}", pr="{pr}", ci="{ci}",
        incomes=[("Discount received", "{dr}")],
        expenses=[("Wages", "{wag}"), ("Rent", "{rent}+{rent_ow}"), ("Insurance", "{ins}-{ins_pre}"),
                  ("General expenses", "{gen}"), ("Carriage outwards", "{co}"), ("Discount allowed", "{da}"),
                  ("Irrecoverable debts", "{id}"),
                  ("Increase in allowance for irrecoverable debts", f"MAX({allow}-{{afd}},0)"),
                  ("Depreciation: equipment", "{eq}*{eq_rate}"), ("Loan interest", "{loan}*{int_rate}")]))
    p.gap()
    p.tag("fs.ex.sofp").statement(["Amina Stores", "Statement of Financial Position at 31 December 2025"], sofp_rows(
        nca=[("Premises", "{prem}", "0"), ("Equipment", "{eq}", "{eqd}+{eq}*{eq_rate}")],
        ca=[("Inventory", "{close}"), ("TR", "{tr}", allow), ("Other receivables (insurance prepaid)", "{ins_pre}"),
            ("Bank", "{bank}"), ("Cash", "{cash}")],
        cl=[("Trade payables", "{tp}"), ("Other payables (rent + loan interest owing)", "{rent_ow}+{loan}*{int_rate}")],
        ncl=[("Loan (repayable 2030)", "{loan}")],
        opening="{cap}", profit="{s_profit}", drawings="{draw}"), "sofp")
    p.gap()
    p.calc([("chk", "Check: net assets minus capital at end of year (must be 0)", "{f_na}-{f_cc}")])


def qfs_6(p):
    p.text("Lina's Shop. The following figures are for the year ended 31 December 2025.")
    p.data([("rev", "Revenue (RM)", 60000), ("sr", "Sales returns (RM)", 1000), ("pur", "Purchases (RM)", 38000),
            ("pr", "Purchases returns (RM)", 800), ("ci", "Carriage inwards (RM)", 500),
            ("oi", "Inventory at 1 January 2025 (RM)", 4500), ("cl", "Inventory at 31 December 2025 (RM)", 5200)],
           md=True)
    p.part("a) Prepare the Statement of Profit or Loss for the year, as far as gross profit. (6)")
    p.statement(["Lina's Shop", "Statement of Profit or Loss (extract) for the year ended 31 December 2025"],
                spl_rows("{rev}", "{oi}", "{pur}", "{cl}", sr="{sr}", pr="{pr}", ci="{ci}", to_gp=True))
    p.part("b) Explain why carriage inwards is included in cost of sales. (1)")
    p.written("Carriage inwards is the cost of bringing purchased goods into the business. It is part of the cost of "
              "buying the goods, so it belongs in cost of sales.", lines=2)


def qfs_7(p):
    p.text("Omar. Balances at 31 December 2025, after all adjustments.")
    p.data([("prem", "Premises at cost (RM)", 50000), ("mv", "Motor vehicle at cost (RM)", 15000),
            ("mvd", "Provision for depreciation: motor vehicle (RM)", 5000), ("inv", "Inventory (RM)", 6000),
            ("tr", "Trade receivables (RM)", 7500), ("bank", "Bank (RM)", 2300), ("cash", "Cash (RM)", 200),
            ("tp", "Trade payables (RM)", 4800), ("loan", "Loan, repayable 2029 (RM)", 10000),
            ("cap", "Capital at 1 January 2025 (RM)", 55000), ("np", "Profit for the year (RM)", 14200),
            ("draw", "Drawings (RM)", 8000)], md=True)
    p.part("a) Prepare the Statement of Financial Position at 31 December 2025. (8)")
    p.statement(["Omar", "Statement of Financial Position at 31 December 2025"], sofp_rows(
        nca=[("Premises", "{prem}", "0"), ("Motor vehicle", "{mv}", "{mvd}")],
        ca=[("Inventory", "{inv}"), ("Trade receivables", "{tr}"), ("Bank", "{bank}"), ("Cash", "{cash}")],
        cl=[("Trade payables", "{tp}")], ncl=[("Loan (repayable 2029)", "{loan}")],
        opening="{cap}", profit="{np}", drawings="{draw}"), "sofp")
    p.part("b) State the working capital. (1)")
    p.calc([("wc", "Working capital (RM)", "{inv}+{tr}+{bank}+{cash}-{tp}")])
    p.part("c) Explain the difference between a current liability and a non-current liability. (2)")
    p.written("A current liability must be paid within 12 months, for example trade payables. A non-current "
              "liability is repayable after more than 12 months, for example a loan repayable in 2029.", lines=2)


NADIA = [("rev", "Revenue", None, 52000), ("pur", "Purchases", 30000, None),
         ("oi", "Inventory at 1 January 2025", 3000, None), ("wag", "Wages", 8000, None), ("rent", "Rent", 4400, None),
         ("ins", "Insurance", 1200, None), ("gen", "General expenses", 900, None), ("fx", "Fixtures at cost", 10000, None),
         ("fxd", "Provision for depreciation: fixtures", None, 2000), ("tr", "Trade receivables", 5000, None),
         ("tp", "Trade payables", None, 3100), ("bank", "Bank", 2600, None), ("cap", "Capital", None, 14000),
         ("draw", "Drawings", 6000, None)]


def qfs_8(p):
    p.text("Nadia's Crafts. Trial balance at 31 December 2025 and additional information.")
    p.tb(NADIA, "Nadia's Crafts: Trial Balance at 31 December 2025")
    p.data([("close", "1. Closing inventory (RM)", 3500), ("rent_ow", "2. Rent owing (RM)", 400),
            ("ins_pre", "3. Insurance prepaid (RM)", 200),
            ("fx_rate", "4. Fixtures depreciation: % of cost per year", 0.1, PCT)],
           title="Additional information at 31 December 2025", md=True)
    p.part("a) Prepare the Statement of Profit or Loss for the year ended 31 December 2025. (10)")
    p.statement(["Nadia's Crafts", "Statement of Profit or Loss for the year ended 31 December 2025"], spl_rows(
        "{rev}", "{oi}", "{pur}", "{close}",
        expenses=[("Wages", "{wag}"), ("Rent", "{rent}+{rent_ow}"), ("Insurance", "{ins}-{ins_pre}"),
                  ("General expenses", "{gen}"), ("Depreciation: fixtures", "{fx}*{fx_rate}")]))
    p.part("b) Prepare the Statement of Financial Position at 31 December 2025. (10)")
    p.statement(["Nadia's Crafts", "Statement of Financial Position at 31 December 2025"], sofp_rows(
        nca=[("Fixtures", "{fx}", "{fxd}+{fx}*{fx_rate}")],
        ca=[("Inventory", "{close}"), ("Trade receivables", "{tr}"),
            ("Other receivables (insurance prepaid)", "{ins_pre}"), ("Bank", "{bank}")],
        cl=[("Trade payables", "{tp}"), ("Other payables (rent owing)", "{rent_ow}")],
        opening="{cap}", profit="{s_profit}", drawings="{draw}"), "sofp")


BEN = [("rev", "Revenue", None, 82000), ("pur", "Purchases", 45000, None), ("sr", "Sales returns", 800, None),
       ("pr", "Purchases returns", None, 900), ("ci", "Carriage inwards", 400, None),
       ("oi", "Inventory at 1 July 2025", 6200, None), ("wag", "Wages", 12600, None), ("elec", "Electricity", 2300, None),
       ("ins", "Insurance", 1600, None), ("adv", "Advertising", 1200, None), ("id", "Irrecoverable debts", 350, None),
       ("rr", "Rent received", None, 2400), ("mv", "Motor vehicles at cost", 24000, None),
       ("mvd", "Provision for depreciation: motor vehicles", None, 6000), ("fx", "Fixtures at cost", 8000, None),
       ("fxd", "Provision for depreciation: fixtures", None, 2000), ("tr", "Trade receivables", 9000, None),
       ("afd", "Allowance for irrecoverable debts", None, 500), ("tp", "Trade payables", None, 5300),
       ("bank", "Bank", 3150, None), ("cap", "Capital", None, 23000), ("draw", "Drawings", 7500, None)]


def qfs_9(p):
    p.text("Ben's Bikes. Trial balance at 30 June 2026 and additional information.")
    p.tb(BEN, "Ben's Bikes: Trial Balance at 30 June 2026")
    p.data([("close", "1. Closing inventory (RM)", 7100), ("elec_ow", "2. Electricity owing (RM)", 250),
            ("ins_pre", "3. Insurance prepaid (RM)", 400),
            ("rr_adv", "4. Rent received in advance for July 2026 (RM)", 200),
            ("al_rate", "5. Allowance: % of trade receivables", 0.04, PCT),
            ("mv_rate", "6. Motor vehicles depreciation: reducing balance", 0.2, PCT),
            ("fx_rate", "7. Fixtures depreciation: % of cost (straight line)", 0.1, PCT)],
           title="Additional information at 30 June 2026", md=True)
    allow = "ROUND({tr}*{al_rate},2)"
    dmv = "ROUND(({mv}-{mvd})*{mv_rate},2)"
    p.part("a) Prepare the Statement of Profit or Loss for the year ended 30 June 2026. (16)")
    p.statement(["Ben's Bikes", "Statement of Profit or Loss for the year ended 30 June 2026"], spl_rows(
        "{rev}", "{oi}", "{pur}", "{close}", sr="{sr}", pr="{pr}", ci="{ci}",
        incomes=[("Rent received", "{rr}-{rr_adv}"),
                 ("Decrease in allowance for irrecoverable debts", f"{{afd}}-{allow}")],
        expenses=[("Wages", "{wag}"), ("Electricity", "{elec}+{elec_ow}"), ("Insurance", "{ins}-{ins_pre}"),
                  ("Advertising", "{adv}"), ("Irrecoverable debts", "{id}"), ("Depreciation: motor vehicles", dmv),
                  ("Depreciation: fixtures", "{fx}*{fx_rate}")]))
    p.part("b) Prepare the Statement of Financial Position at 30 June 2026. (14)")
    p.statement(["Ben's Bikes", "Statement of Financial Position at 30 June 2026"], sofp_rows(
        nca=[("Motor vehicles", "{mv}", "{mvd}+" + dmv), ("Fixtures", "{fx}", "{fxd}+{fx}*{fx_rate}")],
        ca=[("Inventory", "{close}"), ("TR", "{tr}", allow), ("Other receivables (insurance prepaid)", "{ins_pre}"),
            ("Bank", "{bank}")],
        cl=[("Trade payables", "{tp}"),
            ("Other payables (electricity owing + rent in advance)", "{elec_ow}+{rr_adv}")],
        opening="{cap}", profit="{s_profit}", drawings="{draw}"), "sofp")


KOFI = [("rev", "Revenue", None, 98000), ("pur", "Purchases", 54000, None), ("sr", "Sales returns", 1200, None),
        ("pr", "Purchases returns", None, 1000), ("ci", "Carriage inwards", 700, None),
        ("co", "Carriage outwards", 900, None), ("oi", "Inventory at 1 April 2025", 7400, None),
        ("wag", "Wages", 15500, None), ("rr", "Rent and rates", 6300, None), ("elec", "Electricity", 2100, None),
        ("da", "Discount allowed", 350, None), ("dr", "Discount received", None, 650),
        ("cm", "Commission received", None, 1200), ("id", "Irrecoverable debts", 450, None),
        ("prem", "Premises at cost", 40000, None), ("eq", "Equipment at cost", 30000, None),
        ("eqd", "Provision for depreciation: equipment", None, 12000), ("disp", "Disposal account", None, 1500),
        ("tr", "Trade receivables", 14000, None), ("afd", "Allowance for irrecoverable debts", None, 600),
        ("tp", "Trade payables", None, 8200), ("od", "Bank overdraft", None, 1800), ("cash", "Cash", 150, None),
        ("loan", "Loan (repayable 2029)", None, 8000), ("cap", "Capital", None, 50100),
        ("draw", "Drawings", 10000, None)]


def qfs_10(p):
    p.text("Kofi Traders. Trial balance at 31 March 2026 and additional information.")
    p.tb(KOFI, "Kofi Traders: Trial Balance at 31 March 2026")
    p.data([("close", "1. Closing inventory (RM)", 8100), ("elec_ow", "2. Electricity owing (RM)", 300),
            ("rr_pre", "3. Rent and rates prepaid (RM)", 500), ("cm_due", "4. Commission receivable due (RM)", 250),
            ("id_x", "5. Further irrecoverable debt to write off (RM)", 400),
            ("al_rate", "6. Allowance: % of remaining trade receivables", 0.05, PCT),
            ("d_cost", "7. Equipment sold on 31 March 2026: cost (RM)", 5000),
            ("d_acc", "7. Its accumulated depreciation at 1 April 2025 (RM)", 3000),
            ("d_sale", "7. Sale proceeds, already debited to bank and credited to disposal (RM)", 1500),
            ("eq_rate", "8. Remaining equipment: reducing balance rate", 0.2, PCT),
            ("int_rate", "9. Loan interest per year (unpaid)", 0.06, PCT)],
           title="Additional information at 31 March 2026", md=True)
    p.text("10. No depreciation is charged in the year of sale. Premises are not depreciated.")
    trn = "({tr}-{id_x})"
    allow = f"ROUND({trn}*{{al_rate}},2)"
    dep = "ROUND(({eq}-{d_cost}-({eqd}-{d_acc}))*{eq_rate},2)"
    loss = "{d_cost}-{d_acc}-{d_sale}"
    p.part("a) Prepare the Disposal account. (4)")
    p.ledger("Disposal Account", [
        (("2026 Mar 31", "Equipment", "GJ1", "{d_cost}"), ("2026 Mar 31", "Bank", "CB1", "{d_sale}")),
        (None, ("Mar 31", "Provision for depreciation", "GJ1", "{d_acc}")),
        (None, ("Mar 31", "Income statement (loss on disposal)", "GJ2", loss)),
        "TOTAL"])
    p.part("b) Prepare the Statement of Profit or Loss for the year ended 31 March 2026. (18)")
    p.statement(["Kofi Traders", "Statement of Profit or Loss for the year ended 31 March 2026"], spl_rows(
        "{rev}", "{oi}", "{pur}", "{close}", sr="{sr}", pr="{pr}", ci="{ci}",
        incomes=[("Discount received", "{dr}"), ("Commission received", "{cm}+{cm_due}")],
        expenses=[("Wages", "{wag}"), ("Rent and rates", "{rr}-{rr_pre}"), ("Electricity", "{elec}+{elec_ow}"),
                  ("Carriage outwards", "{co}"), ("Discount allowed", "{da}"), ("Irrecoverable debts", "{id}+{id_x}"),
                  ("Increase in allowance for irrecoverable debts", f"{allow}-{{afd}}"),
                  ("Depreciation: equipment", dep), ("Loss on disposal", loss), ("Loan interest", "{loan}*{int_rate}")]))
    p.part("c) Prepare the Statement of Financial Position at 31 March 2026. (16)")
    p.statement(["Kofi Traders", "Statement of Financial Position at 31 March 2026"], sofp_rows(
        nca=[("Premises", "{prem}", "0"), ("Equipment", "{eq}-{d_cost}", f"{{eqd}}-{{d_acc}}+{dep}")],
        ca=[("Inventory", "{close}"), ("TR", trn, allow),
            ("Other receivables (prepaid + commission due)", "{rr_pre}+{cm_due}"), ("Cash", "{cash}")],
        cl=[("Trade payables", "{tp}"), ("Other payables (electricity + loan interest)", "{elec_ow}+{loan}*{int_rate}"),
            ("Bank overdraft", "{od}")],
        ncl=[("Loan (repayable 2029)", "{loan}")],
        opening="{cap}", profit="{s_profit}", drawings="{draw}"), "sofp")


# =====================================================================================
# Year-end journal practice (Section D)
# =====================================================================================
def j8_1(p):
    p.text("Mei Ling's year ends on 31 December 2025. Prepare journal entries for the following. Narratives are required.")
    p.bullets(["30 April 2025: Ahmad, who owes RM480, is declared bankrupt. The debt is written off.",
               "15 July 2025: a cheque for RM350 is received from Faridah. Her debt was written off in 2024.",
               "31 December 2025: the allowance for irrecoverable debts is to be 5% of trade receivables of RM23,000. "
               "The allowance brought forward is RM900."])
    p.data([("ahmad", "Ahmad's debt (RM)", 480), ("far", "Received from Faridah (RM)", 350),
            ("tr", "Trade receivables at 31 December 2025 (RM)", 23000), ("rate", "Allowance rate", 0.05, PCT),
            ("al_open", "Allowance brought forward (RM)", 900)])
    p.part("a) Calculate the change in the allowance for irrecoverable debts. (2)")
    p.calc([("al", "Allowance required at 31 December 2025 (RM)", "ROUND({tr}*{rate},2)"),
            ("inc", "Increase in the allowance (RM)", "{al}-{al_open}")])
    p.part("b) Prepare the journal entries. (8)")
    p.journal([("d", "2025 Apr 30", "Irrecoverable debts", "{ahmad}"), ("c", "", "Ahmad", "{ahmad}"),
               ("n", "Being debt written off: customer declared bankrupt"),
               ("d", "Jul 15", "Faridah", "{far}"), ("c", "", "Irrecoverable debts recovered", "{far}"),
               ("n", "Being debt written off in 2024 reinstated"),
               ("d", "Jul 15", "Bank", "{far}"), ("c", "", "Faridah", "{far}"),
               ("n", "Being cheque received from Faridah"),
               ("d", "Dec 31", "Income statement", "{inc}"), ("c", "", "Allowance for irrecoverable debts", "{inc}"),
               ("n", "Being increase in the allowance for irrecoverable debts")])


def j8_2(p):
    p.text("Chong's year ends on 30 June 2026. All entries are made on 30 June 2026.")
    p.bullets(["Trade receivables are RM32,000 before writing off a debt of RM1,500 owed by Chen, who has disappeared.",
               "The allowance for irrecoverable debts brought forward is RM2,000. It is to be 5% of the remaining trade receivables.",
               "Irrecoverable debts of RM600 were written off earlier in the year. Together with Chen's debt, they are "
               "transferred to the income statement.",
               "Irrecoverable debts recovered during the year total RM300. This is transferred to the income statement."])
    p.data([("trb", "Trade receivables before writing off Chen (RM)", 32000), ("chen", "Chen's debt (RM)", 1500),
            ("al_open", "Allowance brought forward (RM)", 2000), ("rate", "Allowance rate", 0.05, PCT),
            ("id_early", "Irrecoverable debts written off earlier (RM)", 600), ("rec", "Debts recovered (RM)", 300)])
    p.part("a) Calculate the new allowance and the change in the allowance. (3)")
    p.calc([("al", "Allowance required (RM)", "ROUND(({trb}-{chen})*{rate},2)"),
            ("dec", "Decrease in the allowance (RM)", "{al_open}-{al}")])
    p.part("b) Prepare the journal entries. (8)")
    p.journal([("d", "2026 Jun 30", "Irrecoverable debts", "{chen}"), ("c", "", "Chen", "{chen}"),
               ("n", "Being debt written off: customer cannot be found"),
               ("d", "Jun 30", "Allowance for irrecoverable debts", "{dec}"), ("c", "", "Income statement", "{dec}"),
               ("n", "Being decrease in the allowance for irrecoverable debts"),
               ("d", "Jun 30", "Income statement", "{id_early}+{chen}"), ("c", "", "Irrecoverable debts", "{id_early}+{chen}"),
               ("n", "Being irrecoverable debts for the year transferred"),
               ("d", "Jun 30", "Irrecoverable debts recovered", "{rec}"), ("c", "", "Income statement", "{rec}"),
               ("n", "Being debts recovered transferred")])


def j10_1(p):
    p.text("Rahman's year ends on 31 December 2025.")
    p.bullets(["1 January 2025: bought a machine for RM36,000 by cheque. It will last 5 years with a residual value "
               "of RM6,000. Straight line method.",
               "Motor vans cost RM50,000. Accumulated depreciation at 1 January 2025 was RM18,000. Vans are "
               "depreciated at 20% per year, reducing balance."])
    p.data([("mc", "Cost of machine (RM)", 36000), ("mr", "Residual value of machine (RM)", 6000),
            ("ml", "Useful life of machine (years)", 5, INT), ("vc", "Cost of motor vans (RM)", 50000),
            ("va", "Accumulated depreciation on vans at 1 January 2025 (RM)", 18000),
            ("vr", "Reducing balance rate for vans", 0.2, PCT)])
    p.part("a) Calculate the depreciation for 2025 on the machine and on the motor vans. (3)")
    p.calc([("dm", "Depreciation of machine (RM)", "({mc}-{mr})/{ml}"),
            ("dv", "Depreciation of motor vans (RM)", "ROUND(({vc}-{va})*{vr},2)")])
    p.part("b) Prepare the journal entries for the purchase and for the depreciation. (6)")
    p.journal([("d", "2025 Jan 1", "Machinery", "{mc}"), ("c", "", "Bank", "{mc}"),
               ("n", "Being purchase of a machine by cheque"),
               ("d", "Dec 31", "Income statement", "{dm}"), ("c", "", "Provision for depreciation of machinery", "{dm}"),
               ("n", "Being depreciation of machinery for the year"),
               ("d", "Dec 31", "Income statement", "{dv}"), ("c", "", "Provision for depreciation of motor vans", "{dv}"),
               ("n", "Being depreciation of motor vans for the year")])


def j10_2(p):
    p.text("On 1 October 2025 Rahman sold equipment for RM7,500, paid into the bank. The equipment cost RM24,000. "
           "Accumulated depreciation to the date of sale was RM15,000. The year ends on 31 December 2025.")
    p.data([("cost", "Cost of equipment (RM)", 24000), ("acc", "Accumulated depreciation to date of sale (RM)", 15000),
            ("sale", "Sale proceeds (RM)", 7500)])
    p.part("a) Calculate the carrying amount and the loss on disposal. (2)")
    p.calc([("ca", "Carrying amount at the date of sale (RM)", "{cost}-{acc}"),
            ("loss", "Loss on disposal (RM)", "{cost}-{acc}-{sale}")])
    p.part("b) Prepare the journal entries to record the disposal. (8)")
    p.journal([("d", "2025 Oct 1", "Disposal", "{cost}"), ("c", "", "Equipment", "{cost}"),
               ("n", "Being cost of equipment sold transferred"),
               ("d", "Oct 1", "Provision for depreciation of equipment", "{acc}"), ("c", "", "Disposal", "{acc}"),
               ("n", "Being accumulated depreciation on equipment sold transferred"),
               ("d", "Oct 1", "Bank", "{sale}"), ("c", "", "Disposal", "{sale}"),
               ("n", "Being proceeds from the sale of equipment"),
               ("d", "Dec 31", "Income statement", "{cost}-{acc}-{sale}"), ("c", "", "Disposal", "{cost}-{acc}-{sale}"),
               ("n", "Being loss on disposal transferred")])


def j11_1(p):
    p.text("Siva's year ends on 31 December 2025. Prepare the year-end adjustment journals.")
    p.bullets(["Electricity of RM420 for December is owing.",
               "On 1 October 2025 insurance of RM2,400 was paid for 12 months.",
               "Wages of RM650 are owing."])
    p.data([("eo", "Electricity owing (RM)", 420), ("ip", "Insurance paid on 1 October 2025 (RM)", 2400),
            ("im", "Months of insurance used in 2025", 3, INT), ("wo", "Wages owing (RM)", 650)])
    p.part("a) Calculate the insurance expense for 2025 and the amount prepaid. (2)")
    p.calc([("ie", "Insurance expense for 2025 (RM)", "{ip}/12*{im}"),
            ("ipre", "Insurance prepaid at 31 December 2025 (RM)", "{ip}-{ip}/12*{im}")])
    p.part("b) Prepare the journal entries. (6)")
    p.journal([("d", "2025 Dec 31", "Electricity", "{eo}"), ("c", "", "Accrued expenses (other payables)", "{eo}"),
               ("n", "Being electricity owing at the year end"),
               ("d", "Dec 31", "Prepaid expenses (other receivables)", "{ip}-{ip}/12*{im}"),
               ("c", "", "Insurance", "{ip}-{ip}/12*{im}"),
               ("n", "Being insurance paid in advance for 2026"),
               ("d", "Dec 31", "Wages", "{wo}"), ("c", "", "Accrued expenses (other payables)", "{wo}"),
               ("n", "Being wages owing at the year end")])


def j11_2(p):
    p.text("Siva's year ends on 31 December 2025. Prepare the year-end adjustment journals.")
    p.bullets(["Siva lets a room for RM500 per month. Rent received in 2025 was RM6,500, which includes January 2026.",
               "Commission receivable of RM380 for 2025 has not been received.",
               "Unused stationery at 31 December 2025 cost RM150."])
    p.data([("rm", "Rent per month (RM)", 500), ("rr", "Rent received in 2025 (RM)", 6500),
            ("cd", "Commission receivable due (RM)", 380), ("st", "Unused stationery (RM)", 150)])
    p.part("a) Calculate the rent income for 2025 and the rent received in advance. (2)")
    p.calc([("ri", "Rent income for 2025 (RM)", "12*{rm}"), ("ra", "Rent received in advance (RM)", "{rr}-12*{rm}")])
    p.part("b) Prepare the journal entries. (6)")
    p.journal([("d", "2025 Dec 31", "Rent receivable", "{rr}-12*{rm}"),
               ("c", "", "Income received in advance (other payables)", "{rr}-12*{rm}"),
               ("n", "Being rent received in advance for January 2026"),
               ("d", "Dec 31", "Accrued income (other receivables)", "{cd}"), ("c", "", "Commission receivable", "{cd}"),
               ("n", "Being commission earned but not yet received"),
               ("d", "Dec 31", "Inventory of stationery (other receivables)", "{st}"), ("c", "", "Stationery", "{st}"),
               ("n", "Being unused stationery carried forward")])


def jfs_1(p):
    p.text("Amina Stores (see the worked example). Prepare the journal entries for the year-end adjustments at "
           "31 December 2025.")
    p.data([("close", "1. Closing inventory (RM)", 9500), ("rent_ow", "2. Rent owing (RM)", 500),
            ("ins_pre", "3. Insurance prepaid (RM)", 300), ("tr", "4. Trade receivables (RM)", 12000),
            ("afd", "4. Allowance brought forward (RM)", 400), ("al_rate", "4. New allowance: % of trade receivables", 0.05, PCT),
            ("eq", "5. Equipment at cost (RM)", 20000), ("eq_rate", "5. Depreciation: % of cost per year", 0.1, PCT),
            ("loan", "6. Loan (RM)", 10000), ("int_rate", "6. Loan interest per year, unpaid", 0.05, PCT)], md=True)
    inc = "ROUND({tr}*{al_rate},2)-{afd}"
    p.part("a) Prepare the six journal entries. (12)")
    p.journal([("d", "2025 Dec 31", "Inventory", "{close}"), ("c", "", "Income statement", "{close}"),
               ("n", "Being closing inventory recorded"),
               ("d", "Dec 31", "Rent", "{rent_ow}"), ("c", "", "Accrued expenses (other payables)", "{rent_ow}"),
               ("n", "Being rent owing"),
               ("d", "Dec 31", "Prepaid expenses (other receivables)", "{ins_pre}"), ("c", "", "Insurance", "{ins_pre}"),
               ("n", "Being insurance paid in advance"),
               ("d", "Dec 31", "Income statement", inc), ("c", "", "Allowance for irrecoverable debts", inc),
               ("n", "Being increase in the allowance to 5% of trade receivables"),
               ("d", "Dec 31", "Income statement", "{eq}*{eq_rate}"), ("c", "", "Provision for depreciation of equipment", "{eq}*{eq_rate}"),
               ("n", "Being depreciation of equipment for the year"),
               ("d", "Dec 31", "Loan interest", "{loan}*{int_rate}"), ("c", "", "Accrued expenses (other payables)", "{loan}*{int_rate}"),
               ("n", "Being loan interest owing")])


def jfs_2(p):
    p.text("Nadia's Crafts. After all year-end adjustments at 31 December 2025, the balances below remain. Prepare the "
           "closing journal entries that transfer them to the income statement, then transfer the profit and "
           "drawings to capital.")
    p.data([("rev", "Revenue (RM)", 52000), ("oi", "Inventory at 1 January 2025 (RM)", 3000), ("pur", "Purchases (RM)", 30000),
            ("close", "Inventory at 31 December 2025 (RM)", 3500), ("wag", "Wages (RM)", 8000), ("rent", "Rent (RM)", 4800),
            ("ins", "Insurance (RM)", 1000), ("gen", "General expenses (RM)", 900),
            ("dep", "Depreciation: fixtures (RM)", 1000), ("draw", "Drawings (RM)", 6000)], md=True)
    exp = "{wag}+{rent}+{ins}+{gen}+{dep}"
    profit = f"{{rev}}-({{oi}}+{{pur}})+{{close}}-({exp})"
    p.part("a) Calculate the profit for the year. (2)")
    p.calc([("np", "Profit for the year (RM)", profit)])
    p.part("b) Prepare the closing journal entries. (10)")
    p.journal([("d", "2025 Dec 31", "Revenue", "{rev}"), ("c", "", "Income statement", "{rev}"),
               ("n", "Being revenue transferred"),
               ("d", "Dec 31", "Income statement", "{oi}+{pur}"), ("c", "", "Inventory (opening)", "{oi}"),
               ("c", "", "Purchases", "{pur}"), ("n", "Being opening inventory and purchases transferred"),
               ("d", "Dec 31", "Inventory (closing)", "{close}"), ("c", "", "Income statement", "{close}"),
               ("n", "Being closing inventory recorded"),
               ("d", "Dec 31", "Income statement", exp), ("c", "", "Wages", "{wag}"), ("c", "", "Rent", "{rent}"),
               ("c", "", "Insurance", "{ins}"), ("c", "", "General expenses", "{gen}"),
               ("c", "", "Depreciation: fixtures", "{dep}"), ("n", "Being expenses transferred"),
               ("d", "Dec 31", "Income statement", profit), ("c", "", "Capital", profit),
               ("n", "Being profit for the year transferred to capital"),
               ("d", "Dec 31", "Capital", "{draw}"), ("c", "", "Drawings", "{draw}"),
               ("n", "Being drawings transferred to capital")])


# =====================================================================================
# Mock test
# =====================================================================================
MOCK_MCQ = [
    ("Which double entry writes off an irrecoverable debt?",
     ["Debit irrecoverable debts, credit trade receivable", "Debit trade receivable, credit irrecoverable debts",
      "Debit bank, credit irrecoverable debts", "Debit allowance for irrecoverable debts, credit bank"], "A"),
    ("Trade receivables are RM40,000. The allowance is to be 5%. The allowance brought forward is RM2,500. "
     "What is the effect on the Statement of Profit or Loss?",
     ["Expense of RM2,000", "Income of RM500", "Expense of RM500", "Income of RM2,000"], "B"),
    ("A machine costs RM30,000, has a residual value of RM3,000 and a useful life of 6 years. What is the annual "
     "straight line depreciation?", ["RM5,000", "RM3,000", "RM4,500", "RM5,500"], "C"),
    ("Equipment costs RM25,000 and is depreciated at 20% per year, reducing balance. What is the depreciation "
     "charge in Year 2?", ["RM5,000", "RM3,200", "RM20,000", "RM4,000"], "D"),
    ("A vehicle cost RM20,000. Accumulated depreciation is RM12,000. It is sold for RM9,500. What is the result?",
     ["Profit of RM1,500", "Loss of RM1,500", "Profit of RM9,500", "Loss of RM8,000"], "A"),
    ("Rent paid during the year was RM7,200. At the year end RM600 was owing. What is the rent expense?",
     ["RM6,600", "RM7,200", "RM7,800", "RM8,400"], "C"),
    ("How is income received in advance shown in the Statement of Financial Position?",
     ["Current asset", "Current liability", "Other income", "Non-current liability"], "B"),
    ("A RM10 stapler is recorded as an expense, not a non-current asset. Which concept applies?",
     ["Prudence", "Consistency", "Going concern", "Materiality"], "D"),
    ("Where is carriage inwards shown?",
     ["Expenses in the Statement of Profit or Loss", "Cost of sales", "Current liabilities", "Other income"], "B"),
    ("Current assets are RM25,000, current liabilities RM9,000 and a long-term loan RM10,000. What is the working capital?",
     ["RM6,000", "RM34,000", "RM16,000", "RM26,000"], "C"),
]


def mock_a(p):
    p.text("Section A: multiple choice. 1 mark each. Choose A, B, C or D.")
    p.mcq([(f"m{i}", stem, opts, ans) for i, (stem, opts, ans) in enumerate(MOCK_MCQ, 1)])


def mock_b1(p):
    p.text("Lim's year ends on 31 December 2025. Trade receivables are RM45,600 before writing off a debt of RM1,600 "
           "owed by Tan, who is bankrupt. The allowance for irrecoverable debts brought forward is RM1,500. It is to "
           "be 4% of the remaining trade receivables.")
    p.data([("trb", "Trade receivables before writing off Tan (RM)", 45600), ("tan", "Tan's debt (RM)", 1600),
            ("al_open", "Allowance brought forward (RM)", 1500), ("rate", "Allowance rate", 0.04, PCT)])
    p.part("a) Calculate trade receivables after the write-off, the new allowance and the change in the allowance. (3)")
    p.calc([("tr", "Trade receivables after the write-off (RM)", "{trb}-{tan}"),
            ("al", "New allowance (RM)", "ROUND({tr}*{rate},2)"), ("inc", "Increase in the allowance (RM)", "{al}-{al_open}")])
    p.part("b) Prepare the journal entries to write off Tan's debt and to record the change in the allowance. (4)")
    p.journal([("d", "2025 Dec 31", "Irrecoverable debts", "{tan}"), ("c", "", "Tan", "{tan}"),
               ("n", "Being debt written off: customer bankrupt"),
               ("d", "Dec 31", "Income statement", "{inc}"), ("c", "", "Allowance for irrecoverable debts", "{inc}"),
               ("n", "Being increase in the allowance")])
    p.part("c) Show the Statement of Financial Position extract for trade receivables. (3)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Trade receivables", {1: "{tr}"}),
        ("Less:", "Allowance for irrecoverable debts", {1: V("-{al}", line="under"), 2: "{tr}-{al}"})], "sofp")


def mock_b2(p):
    p.text("On 1 January 2024 a business bought a van for RM48,000. It is depreciated at 25% per year, reducing "
           "balance. The year ends on 31 December. On 1 January 2026 the van was sold for RM25,000 by cheque.")
    p.data([("cost", "Cost of van (RM)", 48000), ("rate", "Reducing balance rate", 0.25, PCT),
            ("sale", "Sale proceeds (RM)", 25000)])
    p.part("a) Calculate the depreciation for 2024 and 2025 and the carrying amount at 31 December 2025. (3)")
    p.calc([("d24", "Depreciation for 2024 (RM)", "ROUND({cost}*{rate},2)"),
            ("d25", "Depreciation for 2025 (RM)", "ROUND(({cost}-{d24})*{rate},2)"),
            ("ca", "Carrying amount at 31 December 2025 (RM)", "{cost}-{d24}-{d25}")])
    p.part("b) Prepare the Disposal of van account. (5)")
    p.ledger("Disposal of Van Account", [
        (("2026 Jan 1", "Van", "GJ1", "{cost}"), ("2026 Jan 1", "Provision for depreciation", "GJ1", "{d24}+{d25}")),
        (None, ("Jan 1", "Bank", "CB1", "{sale}")),
        (None, ("Dec 31", "Income statement (loss on disposal)", "GJ2", "{ca}-{sale}")),
        "TOTAL"])
    p.part("c) State where the loss appears, and name the concept that requires depreciation. (2)")
    p.choice([("w", "The loss on disposal is shown as", '"Expense in the SPL"',
               ["Expense in the SPL", "Other income in the SPL", "Current liability", "Deducted from capital"]),
              ("c", "Concept that requires depreciation", '"Accruals"', CONCEPTS, ["accruals", "prudence"])])


def mock_b3(p):
    p.text("Wong's year ends on 31 December 2025.")
    p.bullets(["Rent is RM900 per month. RM9,900 was paid during the year.",
               "Insurance of RM400 was prepaid at 1 January 2025. On 1 April 2025 RM3,000 was paid for 12 months."])
    p.data([("rm", "Rent per month (RM)", 900), ("rp", "Rent paid in 2025 (RM)", 9900),
            ("io", "Insurance prepaid at 1 January 2025 (RM)", 400), ("ip", "Insurance paid on 1 April 2025 (RM)", 3000),
            ("im", "Months of that policy used in 2025", 9, INT)])
    pre = "{ip}*(12-{im})/12"
    p.part("a) Calculate the rent expense, the rent owing, the insurance expense and the insurance prepaid. (4)")
    p.calc([("re", "Rent expense (RM)", "12*{rm}"), ("ro", "Rent owing (RM)", "12*{rm}-{rp}"),
            ("ie", "Insurance expense (RM)", "{io}+{ip}*{im}/12"), ("ipr", "Insurance prepaid (RM)", pre)])
    p.part("b) Prepare the Insurance account for the year. (4)")
    p.ledger("Insurance Account", [
        (("2025 Jan 1", "Balance b/d", "", "{io}"), ("2025 Dec 31", "Income statement", "GJ1", "{io}+{ip}*{im}/12")),
        (("Apr 1", "Bank", "CB1", "{ip}"), ("Dec 31", "Balance c/d", "", pre)),
        "TOTAL",
        (("2026 Jan 1", "Balance b/d", "", pre), None)])
    p.part("c) Show where the rent owing and the insurance prepaid appear. (2)")
    p.statement(["Statement of Financial Position (extract) at 31 December 2025"], [
        ("", "Current assets", {}, "h"),
        ("", "Other receivables (insurance prepaid)", {2: pre}),
        ("Less:", "Current liabilities", {}, "h"),
        ("", "Other payables (rent owing)", {2: "12*{rm}-{rp}"})], "sofp")


def mock_b4(p):
    p.part("a) Name the concept that applies to each situation. (4)")
    p.choice([("a", "Goods taken by the owner are recorded as drawings.", '"Business entity"', CONCEPTS),
              ("b", "A credit sale on 30 December is revenue of that year.", '"Realisation"', CONCEPTS),
              ("c", "The same inventory valuation method is used every year.", '"Consistency"', CONCEPTS),
              ("d", "Staff loyalty is not recorded in the accounts.", '"Money measurement"', CONCEPTS)])
    p.part("b) Goods cost RM3,000. They can be sold for RM2,600 after repairs of RM400. Calculate the net realisable "
           "value and the value for inventory, and name the concept. (3)")
    p.data([("cost", "Cost (RM)", 3000), ("sp", "Selling price (RM)", 2600), ("rep", "Repairs needed (RM)", 400)])
    p.calc([("nrv", "Net realisable value (RM)", "{sp}-{rep}"), ("val", "Value for inventory (RM)", "MIN({cost},{nrv})")])
    p.choice([("cn", "Concept", '"Prudence"', CONCEPTS)])
    p.part("c) Explain the going concern concept and how asset values would change if a business were closing. (3)")
    p.written("Going concern assumes the business will continue trading for the foreseeable future, so assets are shown "
              "at cost less depreciation. If the business were closing, assets would be shown at what they would raise "
              "if sold now, which is often much lower.")


SITI = [("rev", "Revenue", None, 125000), ("pur", "Purchases", 68000, None), ("sr", "Sales returns", 2000, None),
        ("pr", "Purchases returns", None, 1500), ("ci", "Carriage inwards", 1200, None),
        ("co", "Carriage outwards", 800, None), ("oi", "Inventory at 1 January 2025", 9000, None),
        ("wag", "Wages", 18000, None), ("rent", "Rent", 7700, None), ("ins", "Insurance", 2400, None),
        ("elec", "Electricity", 1900, None), ("da", "Discount allowed", 500, None), ("dr", "Discount received", None, 700),
        ("rr", "Rent received", None, 3000), ("id", "Irrecoverable debts", 600, None),
        ("prem", "Premises at cost", 80000, None), ("fx", "Fixtures at cost", 15000, None),
        ("fxd", "Provision for depreciation: fixtures", None, 3000), ("mv", "Motor vehicle at cost", 32000, None),
        ("mvd", "Provision for depreciation: motor vehicle", None, 12000), ("tr", "Trade receivables", 16200, None),
        ("afd", "Allowance for irrecoverable debts", None, 700), ("tp", "Trade payables", None, 9800),
        ("od", "Bank overdraft", None, 2100), ("cash", "Cash", 300, None), ("loan", "Loan (repayable 2030)", None, 20000),
        ("cap", "Capital", None, 89800), ("draw", "Drawings", 12000, None)]


def mock_c(p):
    p.text("Siti Enterprise. Trial balance at 31 December 2025 and additional information.")
    p.tb(SITI, "Siti Enterprise: Trial Balance at 31 December 2025")
    p.data([("close", "1. Closing inventory (RM)", 10500), ("rent_ow", "2. Rent owing (RM)", 700),
            ("ins_pre", "3. Insurance prepaid (RM)", 600), ("rr_adv", "4. Rent received in advance (RM)", 250),
            ("id_x", "5. Further irrecoverable debt to write off (RM)", 200),
            ("al_rate", "6. Allowance: % of remaining trade receivables", 0.05, PCT),
            ("fx_rate", "7. Fixtures: % of cost per year (straight line)", 0.1, PCT),
            ("mv_rate", "7. Motor vehicle: reducing balance rate", 0.25, PCT),
            ("int_rate", "8. Loan interest per year (unpaid)", 0.06, PCT)],
           title="Additional information at 31 December 2025", md=True)
    p.text("Premises are not depreciated.")
    trn = "({tr}-{id_x})"
    allow = f"ROUND({trn}*{{al_rate}},2)"
    dmv = "ROUND(({mv}-{mvd})*{mv_rate},2)"
    p.part("a) Prepare the Statement of Profit or Loss for the year ended 31 December 2025. (25)")
    p.statement(["Siti Enterprise", "Statement of Profit or Loss for the year ended 31 December 2025"], spl_rows(
        "{rev}", "{oi}", "{pur}", "{close}", sr="{sr}", pr="{pr}", ci="{ci}",
        incomes=[("Discount received", "{dr}"), ("Rent received", "{rr}-{rr_adv}")],
        expenses=[("Wages", "{wag}"), ("Rent", "{rent}+{rent_ow}"), ("Insurance", "{ins}-{ins_pre}"),
                  ("Electricity", "{elec}"), ("Carriage outwards", "{co}"), ("Discount allowed", "{da}"),
                  ("Irrecoverable debts", "{id}+{id_x}"),
                  ("Increase in allowance for irrecoverable debts", f"{allow}-{{afd}}"),
                  ("Depreciation: fixtures", "{fx}*{fx_rate}"), ("Depreciation: motor vehicle", dmv),
                  ("Loan interest", "{loan}*{int_rate}")]))
    p.part("b) Prepare the Statement of Financial Position at 31 December 2025. (25)")
    p.statement(["Siti Enterprise", "Statement of Financial Position at 31 December 2025"], sofp_rows(
        nca=[("Premises", "{prem}", "0"), ("Fixtures", "{fx}", "{fxd}+{fx}*{fx_rate}"),
             ("Motor vehicle", "{mv}", "{mvd}+" + dmv)],
        ca=[("Inventory", "{close}"), ("TR", trn, allow), ("Other receivables (insurance prepaid)", "{ins_pre}"),
            ("Cash", "{cash}")],
        cl=[("Trade payables", "{tp}"),
            ("Other payables (rent + rent received in advance + interest)", "{rent_ow}+{rr_adv}+{loan}*{int_rate}"),
            ("Bank overdraft", "{od}")],
        ncl=[("Loan (repayable 2030)", "{loan}")],
        opening="{cap}", profit="{s_profit}", drawings="{draw}"), "sofp")


# =====================================================================================
# Practice (new numbers): figures vary with the set number. Rules are (min, max, step) or a list to pick from.
# =====================================================================================
PRACTICE = {
    "ch08": [("8.11", q8_11, {"b23": (30000, 60000, 1000), "b24": (30000, 60000, 1000), "b25": (30000, 60000, 1000),
                              "w25": (500, 2500, 100), "rate": [0.02, 0.03, 0.04, 0.05]})],
    "ch10": [("10.12", q10_12, {"cost": (24000, 64000, 8000), "rate": [0.2, 0.25], "sale": (6000, 24000, 500)})],
    "ch11": [("11.11", q11_11, {"wo": (100, 600, 50), "wp": (15000, 25000, 500), "wc": (100, 800, 50),
                                "ro": (100, 400, 50), "rp": (1200, 2400, 100), "rc": (100, 400, 50),
                                "cr": (2000, 5000, 100), "cd": (100, 600, 50)})],
    "ch12": [("12.17", q12_17, {"c0": (800, 2000, 100), "s0": (1500, 3000, 100), "k0": (50, 300, 50),
                                "c1": (600, 1500, 100), "s1": (500, 1400, 50), "k1": (100, 300, 50),
                                "c2": (1500, 3000, 100), "s2": (1800, 3600, 100), "k2": (100, 400, 50),
                                "c3": (400, 1000, 50), "s3": (100, 500, 50)})],
    "fs": [("FS.8", qfs_8, {"rev": (45000, 55000, 1000), "pur": (20000, 26000, 500), "oi": (2000, 4000, 100),
                            "wag": (5000, 7000, 100), "rent": (3000, 4000, 100), "ins": (900, 1500, 100),
                            "gen": (500, 1200, 50), "fx": (22000, 30000, 1000), "fxd": (1000, 3000, 500),
                            "tr": (3000, 7000, 100), "tp": (2000, 4000, 100), "bank": (4000, 7000, 100),
                            "draw": (4000, 8000, 500), "cap": "BALANCE", "close": (2500, 4500, 100),
                            "rent_ow": (200, 600, 50), "ins_pre": (100, 300, 50), "fx_rate": [0.1]})],
}


# =====================================================================================
# Workbook assembly
# =====================================================================================
LEVELS = {1: "Level 1: Easy", 2: "Level 2: Easy to medium", 3: "Level 3: Medium", 4: "Level 4: Hard",
          5: "Level 5: Challenge"}


def Q(qid, lvl, fn, section="structured", label=None, sheet=None):
    return {"qid": qid, "level": lvl, "fn": fn, "section": section, "label": label or LEVELS.get(lvl, ""),
            "sheet": sheet or f"Q{qid}"}


def J(qid, lvl, fn):
    return Q(qid, lvl, fn, "journals", f"Journal practice, {LEVELS[lvl]}")


BOOKS = [
    ("Ch08_Irrecoverable_Debts.xlsx", "ch08", "Chapter 8: Irrecoverable Debts and Allowance for Receivables",
     [("Worked 8.1-8.2", worked_8_1), ("Worked 8.3", worked_8_3)],
     [Q("8.8", 1, q8_8), Q("8.9", 2, q8_9), Q("8.10", 3, q8_10), Q("8.11", 4, q8_11), Q("8.12", 5, q8_12),
      J("J8.1", 2, j8_1), J("J8.2", 3, j8_2)]),
    ("Ch10_Depreciation.xlsx", "ch10", "Chapter 10: Tangible Non-current Assets and Depreciation",
     [("Straight Line", worked_10_sl), ("Reducing Balance", worked_10_rb), ("Other Methods", worked_10_other),
      ("Disposal", worked_10_disposal)],
     [Q("10.9", 1, q10_9), Q("10.10", 2, q10_10), Q("10.11", 3, q10_11), Q("10.12", 4, q10_12), Q("10.13", 5, q10_13),
      J("J10.1", 2, j10_1), J("J10.2", 3, j10_2)]),
    ("Ch11_Accruals_Prepayments.xlsx", "ch11", "Chapter 11: Accruals and Prepayments",
     [("Expense Examples", worked_11_exp), ("Income Examples", worked_11_inc)],
     [Q("11.8", 1, q11_8), Q("11.9", 2, q11_9), Q("11.10", 3, q11_10), Q("11.11", 4, q11_11), Q("11.12", 5, q11_12),
      J("J11.1", 2, j11_1), J("J11.2", 3, j11_2)]),
    ("Ch12_Accounting_Concepts.xlsx", "ch12", "Chapter 12: Fundamental Accounting Principles and Concepts",
     [("Concepts", worked_12)],
     [Q("quiz", 0, quiz, "quiz", "Quiz: which concept applies?", "Quiz"),
      Q("12.15", 1, q12_15), Q("12.16", 2, q12_16), Q("12.17", 3, q12_17), Q("12.18", 4, q12_18), Q("12.19", 5, q12_19)]),
    ("Financial_Statements.xlsx", "fs", "Statement of Profit or Loss and Statement of Financial Position",
     [("Worked Example", worked_fs)],
     [Q("FS.6", 1, qfs_6), Q("FS.7", 2, qfs_7), Q("FS.8", 3, qfs_8), Q("FS.9", 4, qfs_9), Q("FS.10", 5, qfs_10),
      J("JFS.1", 2, jfs_1), J("JFS.2", 3, jfs_2)]),
    ("Mock_Test.xlsx", "mock", "Mock Test: Senior 1 Accounting (2 hours, 100 marks)",
     [],
     [Q("A", 0, mock_a, "A", "Section A: multiple choice, 10 marks", "Mock A"),
      Q("B1", 0, mock_b1, "B", "Section B, Chapter 8: 10 marks", "Mock B1"),
      Q("B2", 0, mock_b2, "B", "Section B, Chapter 10: 10 marks", "Mock B2"),
      Q("B3", 0, mock_b3, "B", "Section B, Chapter 11: 10 marks", "Mock B3"),
      Q("B4", 0, mock_b4, "B", "Section B, Chapter 12: 10 marks", "Mock B4"),
      Q("C", 0, mock_c, "C", "Section C, financial statements: 50 marks", "Mock C")]),
]


def heading_for(q):
    if q["section"] == "structured":
        return f"Question {q['qid']} ({q['label']})"
    if q["section"] == "journals":
        return f"Question {q['qid']} ({q['label']})"
    return q["label"]


def readme(wb, chap, topic, worked, questions, practice):
    p = Page(wb, "Read Me", "w", tab=TAB_README)
    p.title(topic, "Senior 1 Accounting. Free to use and share. Not for sale.")
    p.part("How to use this workbook")
    lines = ["1. Read the notes first. Then open the orange Worked tabs: change any blue figure and watch the "
             "answers update.",
             "2. Try the blue question tabs in order. Type your answers in the yellow cells.",
             "3. The small column next to each yellow cell shows \u2713 (correct) or \u2717 (try again). Your score "
             "is at the top of each question.",
             "4. Particulars and dates are given so the sheet can mark your figures. Try writing each account, journal "
             "or statement on paper first.",
             "5. Written answers cannot be marked by Excel. Compare them with the model answers on the green tabs.",
             "6. The green Ans tabs show the full solutions in exam format."]
    if practice:
        lines.append("7. Purple P tabs are practice with new numbers: type any set number from 1 to 999 to get fresh "
                     "figures. Their answers are on the matching P ... Ans tab, not in the printed answer key.")
    if chap == "mock":
        lines = ["1. Sit the paper in 2 hours without notes. Section A: 10 marks. Section B: 40 marks. Section C: 50 marks.",
                 "2. Type answers in the yellow cells on the blue Mock tabs. The Scores tab adds up your correct cells.",
                 "3. The green Ans tabs are the mark scheme. Written answers are marked by comparing with the model answer.",
                 "4. The printed paper is model-questions/Mock_Test.md (or pdf/Mock_Test.pdf)."]
    for line in lines:
        p.text(line, record=False)
    p.part("Colour key")
    r = p.r
    p.box(r, 0, 5, 1000, color="0000FF", fill_=F_DATA, fmt=NUM, align="right")
    p.box(r, 6, 25, "Blue figure: data from the question. Teachers may change it to make a new version.")
    p.box(r + 1, 0, 5, 1000, fmt=NUM, align="right")
    p.box(r + 1, 6, 25, "Black figure: a formula. Do not type over it.")
    p.box(r + 2, 0, 5, None, fill_=F_INPUT)
    p.box(r + 2, 6, 25, "Yellow cell: type your answer here (numbers only; enter 0 for nil).")
    p.box(r + 3, 0, 5, 1000, bold=True, fill_=F_ANS, fmt=NUM, align="right")
    p.box(r + 3, 6, 25, "Green cell: the correct answer (Ans tabs).")
    p.r += 4
    p.part("Formats used")
    for line in ["Journal: Date | Particulars | Debit | Credit. The first row shows the year, with RM under Debit and Credit.",
                 "Ledger account: Dr side and Cr side, each with Date | Particulars | Folio | Amount. The first row under "
                 "the headings shows the year, with RM in the Amount column; entries then show the month and day. "
                 "Folio: GJ = general journal, CB = cash book.",
                 "Statement of Profit or Loss: marker (Add / Less) | Particulars | RM | RM | RM.",
                 "Statement of Financial Position: the same, with Cost | Accumulated Depreciation | Carrying Amount "
                 "under the RM headings."]:
        p.text(line, record=False)
    p.part("Sheets in this workbook")
    rows = [(n, [T("Worked example (orange tab)")]) for n, _ in worked]
    for q in questions:
        rows.append((q["sheet"], [T(f"{heading_for(q)}. Answers: {q['sheet']} Ans (green tab).")]))
    for qid, _fn, _v in practice:
        rows.append((f"P{qid}", [T(f"Practice with new numbers for Question {qid} (purple tab). Answers: P{qid} Ans.")]))
    if chap == "mock":
        rows.insert(0, ("Scores", [T("Your marks for every section of the mock test.")]))
    p.table(["Sheet", "What it is"], rows, 7, 24, checks=False)
    p.finish()


def scores_sheet(wb, pages):
    p = Page(wb, "Scores", "w", tab=TAB_README)
    p.title("Mock test: your scores", "Marks come from the \u2713 cells on each blue Mock tab.")
    rows = []
    for pq, q in pages:
        rng = f"'{pq.name}'!B1:AH400"
        rows.append((q["label"], [V(f'COUNTIF({rng},"\u2713")', key=f"s_{q['qid']}", fmt=INT), T(str(pq.inputs))]))
    keys = "+".join("{s_" + q["qid"] + "}" for _pq, q in pages)
    total = sum(pq.inputs for pq, _q in pages)
    rows.append(("Total", [V(keys, line="total", fmt=INT), T(str(total))]))
    p.table(["Section", "Answer cells correct", "Answer cells in total"], rows, 13, 9, checks=False,
            total_rows=("Total",))
    p.note("Each answer cell is worth roughly one mark in Sections A and B. Section C has more cells than marks, "
           "so use the Ans tabs as the mark scheme for the final mark out of 100.")
    p.finish()
    return p


TAB_P = "7030A0"


def build():
    OUT.mkdir(exist_ok=True)
    manifest = {}
    for fname, chap, topic, worked, questions in BOOKS:
        practice = PRACTICE.get(chap, [])
        wb = Workbook()
        wb.remove(wb.active)
        readme(wb, chap, topic, worked, questions, practice)
        sheets = []
        for name, fn in worked:
            p = Page(wb, name, "w", tab=TAB_WORKED)
            p.title(name)
            p.status()
            fn(p)
            p.finish()
            sheets.append({"sheet": name, "mode": "w", "blocks": p.blocks})
        items = [(q["sheet"], heading_for(q), q, False, None) for q in questions]
        items += [(f"P{qid}", f"Practice with new numbers: Question {qid}", {"qid": qid, "fn": fn, "section": "practice",
                   "level": 0, "label": "practice", "sheet": f"P{qid}"}, True, vary) for qid, fn, vary in practice]
        pages = []
        for name, heading, q, prac, vary in items:
            ans = f"{name} Ans"
            pq = Page(wb, name, "q", twin=ans, tab=TAB_P if prac else TAB_Q, practice=prac, vary=vary)
            pq.title(heading)
            pq.status()
            q["fn"](pq)
            pages.append((pq, name, ans, heading, q, prac, vary))
        for pq, name, ans, heading, q, prac, vary in pages:
            pa = Page(wb, ans, "a", twin=name, tab=TAB_A, practice=prac, vary=vary)
            pa.title(heading + ": Answers")
            pa.status()
            q["fn"](pa)
            assert pa.r == pq.r, f"layout mismatch in {name}"
            pq.finish()
            pa.finish()
            mode = "p" if prac else "q"
            meta = {"qid": q["qid"], "level": q["level"], "section": q["section"], "label": q["label"]}
            sheets.append({"sheet": name, "mode": mode, **meta, "blocks": pq.blocks})
            sheets.append({"sheet": ans, "mode": "pa" if prac else "a", **meta, "blocks": pa.blocks})
        if chap == "mock":
            scores_sheet(wb, [(pq, q) for pq, _n, _a, _h, q, _p, _v in pages])
            wb.move_sheet("Scores", offset=-(len(wb.sheetnames) - 2))
        wb.active = 0
        wb.save(OUT / fname)
        manifest[fname] = {"chapter": chap, "sheets": sheets}
        print("wrote", fname)
    (ROOT / "tools" / "manifest.json").write_text(json.dumps(manifest, indent=1))


if __name__ == "__main__":
    build()
