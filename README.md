# Payout Reconciliation Report

![tests](https://github.com/Adnan040404/payout-reconciliation-excel-report/actions/workflows/tests.yml/badge.svg)

I wanted a report that answers one question quickly: which invoices got paid,
which were paid short, and which were never paid at all. This workbook does that
with ordinary Excel formulas, so it keeps working when you paste in new data.

The data is made up. I generated it with a script, so nothing here comes from a
real client or company.

![Summary sheet](screenshots/summary.png)

## The problem

When money arrives through marketplaces, retailers or payment processors, the
invoices and the payments come from different systems and rarely line up
one-to-one. The usual trouble:

| Status | What happened |
|---|---|
| Unpaid | The invoice went out and no payment came back |
| Short pay | The payment was smaller than the invoice |
| Overpaid | The payment was bigger than the invoice |
| Duplicate | The same invoice was paid two or more times |
| Unapplied | A payment arrived that doesn't match any invoice |

Checking this by eye across a few hundred rows is slow, and mistakes are easy to
miss.

## What's in the workbook

- **Summary** has the count and dollar amount for each status, a few headline
  numbers (money owed, excess received, unapplied payments) and a chart.
- **Reconciliation** lists every invoice next to what was actually paid, the
  difference, and a colour-coded status. Filter the Status column to see which
  invoices to chase.
- **Payments** lists every payment and marks it Applied or Unapplied.

![Reconciliation sheet](screenshots/reconciliation.png)

## How the matching works

Each invoice is matched to payments on account plus PO number. Payments for the
same PO are added together and compared with the invoice amount. If the two are
within the tolerance (one cent by default, editable in `Summary!B4`) it's Paid.
Otherwise it's Short Pay or Overpaid. If there were two or more payments that
add up to exactly 2x, 3x or 4x the invoice, it's flagged as a Duplicate instead
of an overpayment.

Everything is a formula (`SUMIFS`, `COUNTIFS`, `IF`), so pasting a new month into
the Payments sheet and the invoice columns updates the whole workbook.

## Limits

- The duplicate rule is deliberately simple. On real data I would also look at
  payment dates and remittance references before calling something a duplicate.
- Formulas are fine for a few thousand rows. Beyond that I'd move the matching
  into SQL or Python. I did that in
  [dropship-reconciliation-engine](https://github.com/Adnan040404/dropship-reconciliation-engine).

## Using your own data

The workbook is built from two CSV files:

| File | Columns |
|---|---|
| Invoices | `invoice_id, account_code, po_number, invoice_amount, invoice_date` |
| Payments | `payment_id, account_code, po_number, payment_amount, payment_date, payment_ref` |

```bash
pip install -r requirements.txt
python build_report.py --invoices my_invoices.csv --payments my_payments.csv --out my_report.xlsx
```

With no arguments it rebuilds the sample workbook from the files in `data/`. If a column
is missing or an amount isn't a number, it stops and says which file, which column, and
which row, rather than producing a report with quiet errors. Excel calculates the
formulas when you open the file.

```bash
python -m pytest tests -q      # 7 tests
```

## Contact

Muhammad Adnan, [LinkedIn](https://linkedin.com/in/muhammad-adnan-740336293),
adnandanish0404@gmail.com
