# AR Statement Automation Template

An Excel-based automation tool that streamlines the preparation and distribution of monthly Accounts Receivable statements. Built using Power Query, VBA, and conditional formatting to eliminate repetitive manual work and reduce statement preparation time significantly.

## What It Does

- Automatically cleans and formats raw AR statement data on refresh
- Removes unnecessary columns, keeping only customer-relevant information
- Applies accounting/currency formatting to all dollar amounts
- Color codes rows by status — red for credits, yellow for past due, bold for due this month
- Separates Claim Investigations into a dedicated Claims tab automatically
- Calculates a dynamic Monthly Summary and Payment Schedule that updates based on the current date
- Exports a clean, formatted .xlsx file with a single button click — no manual copy/paste required

## Tools & Techniques Used

- **Oracle ERP (AR Module)** — source system for raw statement exports processed by this template
- **Power Query** — automated data transformation, column filtering, and sorting on refresh
- **VBA (Macros)** — one-click data clearing and automated export with dynamic filename generation
- **Conditional Formatting** — dynamic row highlighting based on payment status
- **SUMPRODUCT & SUMIF Formulas** — dynamic summary calculations that update automatically based on the current month and year using TODAY()
- **Excel Table References** — formula references that never shift regardless of data size

## File Structure

| Tab | Purpose |
|-----|---------|
| Instructions | Full user guide, formula reference, and contents list |
| Raw Data | Paste raw exported statement data here |
| Clean Statement | Formatted output with summary boxes and export buttons |
| Claims | Claim Investigation rows separated automatically |

## How It Works

1. Open the template — a fresh copy opens automatically
2. Click Clear Raw Data to wipe previous account data
3. Paste the raw statement export into the Raw Data tab starting at A1
4. Click Data → Refresh All to trigger Power Query cleanup
5. Review the Clean Statement tab
6. Click Export Clean Statement — enter the account number and the file saves automatically as Account Number - Month Year Aging.xlsx

## Summary Box Calculations

The template includes two dynamic summary boxes that recalculate based on the current calendar month.

Monthly Summary includes Due thru EOM for positive invoices due this month that are not yet past due, Past Due for invoices with a positive Days Late value, Total Credits for the sum of all credit memos, and Net Balance as the combined total.

Payment Schedule breaks down what is due by the 10th, 20th, and end of month to help customers plan payment runs.

## Skills Demonstrated

- Excel automation without third-party tools
- Power Query data transformation and dynamic sorting
- VBA macro development for workflow automation
- Financial data presentation and formatting
- Process improvement and workflow documentation
- End-to-end solution design from problem identification to deployment
