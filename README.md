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
