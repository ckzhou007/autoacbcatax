# autoacb.py

A small, dependency-light command-line tool that turns a spreadsheet of buy/sell
transactions into a CRA-ready Adjusted Cost Base (ACB) and capital gains
calculation — including the superficial loss rule, return of capital, and
stock splits — output as a new Excel workbook with **Schedule 3** and
**T1135** tabs ready for transcription onto your tax return.

> **Disclaimer:** This tool is provided for informational purposes to help
> organize your own records. It is not tax or legal advice, and it isn't a
> substitute for professional advice from an accountant or the CRA's own
> guidance. Always review the output before filing, and verify edge cases
> against your own situation.

## Features

- **ACB tracking per ticker**, pooled correctly across multiple buys/sells
- **Superficial loss rule** for both full and *partial* dispositions, using
  the `min(S, P, B) / S × Total Loss` formula, including correct
  handling when a stock split falls inside the 61-day window. 
- **Return of Capital (ROC)** — reduces ACB under s.53(2)(h), with automatic
  deemed-gain handling if ACB would go negative
- **Stock splits and reverse splits / consolidations** — share count and
  ACB/share adjust automatically without changing total ACB
- **Multi-currency support** with automatic CAD conversion using exchange
  rates straight from the Bank of Canada
- **Multi-year in one pass** — every year present in your data is processed
  and given its own Schedule 3 and T1135 tab automatically
- **CRA Schedule 3** output, formatted for direct transcription
- **T1135 (Foreign Income Verification)** monthly tracking per ticker, with
  live formulas that recalculate if you edit the reportable flag
- Single dependency (`openpyxl`), single command to run

### ⚠️ The "Registered Account" Trap (TFSA / RRSP / FHSA)

The Canada Revenue Agency’s superficial loss rules look at **all accounts owned by you and your spouse**—including tax-sheltered accounts. 

If you sell a security at a loss in your taxable account, and repurchase that same ticker inside a **TFSA, RRSP, FHSA, or RESP** within the 61-day window:

1. Your capital loss is **denied**.
2. **The tax shield is permanently destroyed.** Because registered accounts do not track ACB, the denied loss cannot be added to the cost base of the repurchased shares. It simply vanishes.

*Note: autoacb only calculates the taxable data you paste into it. It cannot look inside your brokerage's TFSA or RRSP accounts to warn you if you have accidentally triggered a cross-account superficial loss.*

## 🛠️ Installation & Requirements (For Beginners)

This script requires **Python** (a free programming language) and a helper library called **openpyxl** to read and write Excel files. You do not need to know how to code to use it.

### Step 1: Install Python
Download and run the official installer for your operating system:

* **Windows**: Download the installer from [Python.org](https://www.python.org/downloads/). **CRITICAL STEP:** When the installer window opens, you must check the box at the very bottom that says **"Add python.exe to PATH"** before clicking Install. If you skip this, the steps below will fail.
* **Mac**: Download the macOS installer from [Python.org](https://www.python.org/downloads/). Open the downloaded package and click through the standard installation prompts.
* **Linux**: Python 3 comes pre-installed on virtually all Linux distributions. You can verify this by opening a terminal and typing `python3 --version`.

### Step 2: Install the Excel Helper (openpyxl)
Once Python is installed, you need to open your computer's command line to install the Excel plugin:

1. Open your terminal
   
2. Copy and paste the following exact line into that window and press **Enter**:

 -- **Windows & Mac**: pip install openpyxl

 -- **Linux**: sudo apt install python3-openpyxl

## Usage

1. Download `autoacb.py` to your working directory.
2. Create an Excel workbook containing two sheets: `Transactions` and `ExchangeRate`.
3. Run the script:

   ```bash
   python3(python if using windows) autoacb.py sourcefile.xlsx
   ```

This writes a new file, `sourcefile_taxReview.xlsx`, in the same folder as
your input file — your original workbook is never modified.

> **Tip:** if you re-run the script, it overwrites the output file without
> asking. On Windows, having that file open in Excel will cause the script to
> fail loudly (file in use). On Linux/Mac, most spreadsheet apps don't lock
> the file at the OS level, so it can get overwritten *silently* while still
> open elsewhere — close it first to be safe.

## Input format

Your input `.xlsx` needs two sheets:

### `Transactions`

| Date       | Symbol | settlementDate | Type | Quantity | Price | Fee | Currency |
|------------|--------|-----------------|------|----------|-------|-----|----------|
| 2024-01-03 | AAPL   | 2024-01-05      | BUY  | 10       | 150   | 5   | USD      |
| 2024-06-10 | AAPL   | 2024-06-12      | SELL | 4        | 180   | 5   | USD      |

- If both `Date` and `settlementDate` are present, **settlementDate is used**
  for ACB purposes (not the trade date).
- `Fee`/`Commission` is optional — missing values are treated as $0.
- `Quantity` can be positive or negative; `Type` is authoritative for
  direction, and the sign is normalized automatically.
- **`Type`** accepts: `BUY`, `SELL`, `ROC`, `SPLIT` — plus common broker
  synonyms (`DIY_BUY`, `ReturnOfCapital`, `Consolidation`, `Reverse Split`,
  etc.). Anything else (dividends, transfers, interest, ...) is skipped and
  reported in a one-line console summary — it never appears in the output.
- **`ROC`** (e.g. a T3 slip Box 42): leave `Quantity` blank (defaults to 1)
  and put the *total amount received* in `Price`. If T3 provides a CAD amount 
  but the security is not in CAD, convert it to the security's currency first
  and enter that in Price, because the FX rate is applied to the Price to get the CAD amount. 
- **`SPLIT`** (also covers reverse splits/consolidations): `Quantity` = new
  shares, `Price` = old shares — e.g. a 5-for-1 split is `Quantity=5,
  Price=1`; a 10-for-1 reverse split is `Quantity=1, Price=10`.

### `ExchangeRate`

Paste this straight from a [Bank of Canada Valet CSV download](https://www.bankofcanada.ca/valet/observations/FXUSDCAD,FXEURCAD/csv?start_date=2024-01-01) —
no reformatting needed(row 1 is the header):

| date       | FXUSDCAD | FXEURCAD |
|------------|----------|----------|
| 2024-01-05 | 1.3445   | 1.4521   |
| 2024-06-12 | 1.3689   | 1.4702   |

- Any `FX<CCY>CAD` column is auto-detected, so add as many currencies as you
  need.
- CAD itself needs no rows — it's always rate 1.
- For each transaction, the most recent rate **on or before** the
  transaction date is used.

## Output

The output workbook contains your original sheets plus:

| Sheet                  | Contents                                                                 |
|-------------------------|---------------------------------------------------------------------------|
| `Summary`               | Full chronological ACB ledger per ticker — every transaction, running ACB, running share balance, and superficial-loss detail, reviewable line by line. |
| `Schedule 3 - <year>`   | One per year with reportable activity, columns matching CRA Schedule 3 (publicly traded shares section), ready to transcribe. |
| `T1135 - <year>`        | One per year, tickers as rows and months as columns, for the $100,000 specified foreign property test. Non-CAD tickers default to reportable; edit the "T1135 Reportable?" column directly — the monthly totals are live formulas that update automatically. |

## How it works (brief)

1. Reads and validates both sheets, normalizing transaction and currency.
2. Builds a chronological ledger per ticker, applying ACB pooling rules.
3. On each disposition, checks the 61-day window (30 days before/after) for
   a superficial loss, pro-rating it for partial dispositions using share
   counts rescaled for any stock splits inside the window.
4. Applies ROC reductions and split/consolidation adjustments in date order.
5. Writes the Summary ledger, then aggregates reportable dispositions and
   month-end foreign holdings into the Schedule 3 and T1135 tabs.

## Sources

- [AdjustedCostBase.ca - Applying the Superficial Loss Rule for a Partial Disposition of Shares](https://www.adjustedcostbase.ca/blog/applying-the-superficial-loss-rule-for-a-partial-disposition-of-shares/)
- [AdjustedCostBase.ca — Calculating ACB with Foreign Currency Transactions](https://www.adjustedcostbase.ca/blog/calculating-adjusted-cost-base-with-foreign-currency-transactions/)
- [Bank of Canada - Daily exchange rates: Lookup tool](https://www.bankofcanada.ca/rates/exchange/daily-exchange-rates-lookup/)


## Contributing

Issues and discussions are welcome.

## License

Released under the [MIT License](LICENSE).

## ⚠️ Disclaimer

*autoacb is an open-source automation tool written for personal portfolio organization. It does not constitute certified accounting software, nor does it provide legal or tax advice. The user assumes sole responsibility for verifying the output against their official brokerage statements before submitting figures to the Canada Revenue Agency.*
