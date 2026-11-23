# Validus Case Study – Fund IRR, NAV and FX Hedging

A Python program that reads a fund's cashflows from Excel and then:

- validates them with Pydantic, correcting the data where needed;
- calculates fund-level and currency-level IRRs;
- builds a NAV schedule for each currency;
- proposes FX forward trades for a 3-month rolling hedge of 100% of NAV.

## Setup

Requires Python 3.10+.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running the program

The input file is not in the repo. Place it at `data/CST_Case_Study_data__1_.xlsx`, then run:

```bash
python main.py data/CST_Case_Study_data__1_.xlsx
```

Since the brief asks for a program rather than a one-off script, the brief's parameters are optional arguments whose defaults reproduce the requested outputs:

| Option | Default | |
|---|---|---|
| `--fund` | `Fund I` | Fund to analyse |
| `--hedge-ratio` | `1.0` | Share of NAV to hedge, 0 to 1 |
| `--tenor-months` | `3` | Months per rolling forward |
| `--output-format` | `excel` | `excel` (one workbook) or `csv` (three files) |
| `--output-dir` | `output` | Where results are saved |
| `--log-level` | `INFO` | `DEBUG` for row-level detail |

The input must be an `.xlsx` or `.xlsm` workbook. Anything else, such as a missing file, a CSV, an old `.xls` file or missing columns, stops the run with a clear error.

## Output

The results are printed to the console, and also saved to `output/`:

- **Excel** (default): `fund_i_results.xlsx`, with sheets `IRR`, `NAV` and `FX Hedges`.
- **CSV**: `fund_i_irr.csv`, `fund_i_nav.csv` and `fund_i_fx_hedges.csv`.

For Fund I:

| | IRR |
|---|---|
| Fund (EUR) | 10.0476% |
| EUR | 7.8870% |
| GBP | 9.9507% |
| USD | 12.0515% |

There are 40 hedge trades: 20 in EURGBP and 20 in EURUSD, each selling the foreign currency against EUR.

## Tests

```bash
python -m unittest discover -s tests -t .
```

The tests use `unittest` and cover validation and cleaning, IRR, NAV, hedging and the command line. End-to-end tests run the whole program and compare its output files with known-good md5 hashes. The tests that use the real input file are skipped if it isn't in `data/`.

## Data corrections

Each correction is logged when the program runs.

| ID | Problem | Fix |
|---|---|---|
| 17 | Currency `GPB` | → `GBP` |
| 21, 51 | Date stored as day-first text with extra spaces / a backtick | Parsed as `DD/MM/YYYY` |

The 6 `Principal Repayment` rows were a cashflow type missing from the model, so they were fixed in the code rather than the data.

## Key assumptions

- **IRR:** XIRR with an actual/365 day count, the same as Excel.
- **NAV:**
    - NAV(t) is the PV of the remaining cashflows at the currency's IRR, including any cashflow on date t. This is what makes both of the brief's checks hold: NAV(0) = 0, and NAV on the final date equals the exit value.
    - The positions have no coupon on the exit date; this is treated as intended.
- **Fund IRR:** uses the EUR amounts in the file. The FX rates there are fixed, so it is a blend of the currency IRRs and doesn't include FX moves.
- **Hedges:**
    - Only GBP and USD are hedged, by selling them forward against EUR, the fund's base currency.
    - Each notional is the NAV just *after* that day's cashflow; otherwise the first hedge would be 0.
- **Dates:**
    - Rolls are every 3 months from the first cashflow, keeping month-ends.
    - Dates that fall on weekends or holidays are moved with the modified-following rule, using both currencies' holiday calendars.
- **Not modelled:** spot settlement lag and forward pricing.
