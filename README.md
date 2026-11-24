# conference-expense-report

A [Claude](https://claude.ai) skill that turns a folder of travel or conference receipts (photos, PDFs, screenshots) into a reimbursement-ready Excel workbook.

It works with receipts in any language and any currency.

## What you get

- **Summary sheet:** every expense in timeline order. Pre-trip costs such as registration, visa and insurance come first.
- **Columns:** `# | Date | Category | Item | Currency | Amount (original) | USD | Receipt | Comments`
- **Receipt links:** each row links to its own sheet with the receipt image, and each receipt sheet links back to the summary.
- **Exchange rates:** a small, editable rate table. Change a rate and every converted amount updates.
- **Totals:** a total, subtotals by category with a check sum, and amounts spent per currency.
- **Comments:** factual notes an accountant needs, such as shared bills, cash payments, missing card slips and fee breakdowns.

## Install

**Option A: skill file**

1. Download [`conference-expense-report.skill`](conference-expense-report.skill).
2. Open it in the Claude desktop app.
3. Click **Save skill**.

**Option B: copy the text**

1. Open [`conference-expense-report/SKILL.md`](conference-expense-report/SKILL.md) and copy everything.
2. Paste it into a Claude chat and say "save this as a skill."

## Use

1. In Claude (Cowork), select the folder with your receipts. Subfolders are fine.
2. Ask in plain words, for example: *"These are receipts from my conference in Lisbon, Mar 10–15. Make the reimbursement Excel."*
3. Mention anything the receipts can't show, such as *"the dinner on 3/12 covered two people"* or *"use my department's exchange rate."*

You can also type `/conference-expense-report` to run it by name.

## Notes

- Claude reads every receipt. It never guesses amounts, and it lists anything that looks missing.
- The default reimbursement currency is USD. Ask for another one if your institution uses it.
- To open the result, you need Excel or any other spreadsheet app.
