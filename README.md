# Validus Case Study – Fund IRR, NAV and FX Hedging

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

The input file is not committed to the repo. Place it at
`data/CST_Case_Study_data__1_.xlsx` before running.

```bash
python main.py data/CST_Case_Study_data__1_.xlsx --fund "Fund I"
```

## Test

```bash
python -m unittest discover -s tests -t .
```

## Design

TODO

## Data corrections

TODO

## Assumptions

- **No coupon on the exit date.** Every position pays 19 quarterly coupons, the
  last on 2030-06-30, and only the principal on 2030-09-30. This is treated as
  the intended deal terms, not a missing row. It is why each currency IRR sits
  slightly below its coupon rate (GBP: 10% coupon, 9.95% IRR) and why NAV stays
  above par and falls to the principal at exit.
- **Hedge sizing.** Each forward's notional is the hedge ratio times the
  currency's NAV just *after* that day's cashflow, i.e. the value still
  exposed over the coming period. The brief's NAV includes the day's cashflow,
  which is 0 on the investment date and would leave the first quarter
  unhedged. Notionals are rounded to 2 decimal places.
- **What is hedged.** Only non-base currencies (GBP, USD). The fund is long
  them, so each forward sells the foreign currency for EUR. `direction`
  refers to the notional currency; pairs use market quoting order (EURGBP,
  EURUSD).
- **Roll dates.** Every 3 months from the first cashflow, keeping month-ends
  as month-ends (30 Sep → 31 Dec → 31 Mar), so rolls match the quarter-end
  cashflow dates. Each forward delivers on the next roll date; the last one
  delivers on the final cashflow date and is not rolled.
- **Business days.** Trade and delivery dates are moved off weekends and
  holidays with the modified-following convention, using both currencies'
  settlement calendars from the `holidays` package: TARGET2 for EUR, England
  bank holidays for GBP, the Federal Reserve for USD. Month-end rolls stay in
  their month, e.g. Sat 31 Mar 2029 moves back past Good Friday to Thu 29 Mar.
  A moved trade is still sized on the quarter-end roll date, so it doesn't
  hedge a coupon that is paid a day or two later. Future one-off holidays are
  only known once the package adds them.
- **Not modelled.** Spot settlement lag and forward points/rates. The output
  is the list of trades to place, not their pricing.
