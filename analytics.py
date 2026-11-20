"""Calculations: IRR, NAV schedules and FX hedges. Data in, numbers out, no I/O."""

import logging
from datetime import date
from decimal import Decimal

from models import Cashflow, FxForward, NavPoint

logger = logging.getLogger(__name__)


# --- Fund data --------------------------------------------------------------

def base_currency(cashflows: list[Cashflow]) -> str:
    """The fund's reporting currency, read from the data. Every cashflow must agree."""
    currencies = {cf.base_currency for cf in cashflows}
    if len(currencies) != 1:
        raise ValueError(f"Expected one base currency, found {sorted(currencies)}")
    currency = currencies.pop()
    logger.info("Base currency: %s", currency)
    return currency


# --- IRR ---------------------------------------------------------------------
# Day count is actual/365, the same convention as Excel's XIRR, so results can
# be checked against a spreadsheet. The IRR search works in floats (it is a
# numerical search, not money); NAVs are computed and reported in Decimal.

DAYS_IN_YEAR = 365


def year_fraction(start: date, end: date) -> float:
    """Time between two dates in years, actual/365."""
    return (end - start).days / DAYS_IN_YEAR


def npv(rate: float, dates: list[date], amounts: list[Decimal]) -> float:
    """Net present value of dated cashflows at an annual rate, as at the earliest date."""
    if rate <= -1:
        raise ValueError(f"Rate must be greater than -100%, got {rate}")
    t0 = min(dates)
    return sum(
        float(amount) / (1 + rate) ** year_fraction(t0, d)
        for d, amount in zip(dates, amounts)
    )


def xirr(
    dates: list[date],
    amounts: list[Decimal],
    low: float = -0.99,
    high: float = 10.0,
    tolerance: float = 1e-12,
    max_iterations: int = 200,
) -> float:
    """The annual rate at which npv(...) == 0, found by bisection.

    Bisection is slower than Newton's method but cannot diverge: as long as the
    NPV changes sign between `low` and `high`, it always converges. With one
    outflow followed by inflows (as here) there is exactly one root.
    """
    if len(dates) != len(amounts) or not dates:
        raise ValueError("dates and amounts must be non-empty and the same length")
    if not (any(a > 0 for a in amounts) and any(a < 0 for a in amounts)):
        raise ValueError("IRR needs at least one positive and one negative cashflow")

    npv_low = npv(low, dates, amounts)
    if npv_low * npv(high, dates, amounts) > 0:
        raise ValueError(f"No IRR found between {low:.0%} and {high:.0%}")

    for iteration in range(1, max_iterations + 1):
        mid = (low + high) / 2
        npv_mid = npv(mid, dates, amounts)
        if npv_mid == 0 or (high - low) / 2 < tolerance:
            logger.debug("XIRR converged to %.10f after %d iterations", mid, iteration)
            return mid
        if npv_low * npv_mid < 0:
            high = mid
        else:
            low, npv_low = mid, npv_mid
    raise ValueError(f"XIRR did not converge in {max_iterations} iterations")


def _irr_of(cashflows: list[Cashflow], amount_field: str) -> float:
    return xirr([cf.date for cf in cashflows], [getattr(cf, amount_field) for cf in cashflows])


def currency_irrs(cashflows: list[Cashflow]) -> dict[str, float]:
    """IRR per local currency, using local-currency amounts."""
    irrs = {}
    for currency in sorted({cf.local_currency for cf in cashflows}):
        in_currency = [cf for cf in cashflows if cf.local_currency == currency]
        irrs[currency] = _irr_of(in_currency, "amount_local")
        logger.info("IRR %s (local): %.4f%%", currency, irrs[currency] * 100)
    return irrs


def fund_irr(cashflows: list[Cashflow]) -> float:
    """IRR for the whole fund, using base-currency amounts."""
    irr = _irr_of(cashflows, "amount_base")
    logger.info("IRR fund (base): %.4f%%", irr * 100)
    return irr


# --- NAV ---------------------------------------------------------------------
# Convention: NAV(t) includes the cashflow paid on date t, i.e. it is the value
# just before that day's cashflow. This is the reading under which both of the
# brief's sanity checks hold: at the first date the sum is the full NPV, which
# is zero at the IRR; at the final date only the exit cashflow remains.
#
#     NAV(t) = sum over s >= t of  CF_s / (1 + r) ^ ((s - t) / 365)

NAV_TOLERANCE = Decimal("0.01")  # one cent / penny


def nav_at(
    valuation_date: date, dates: list[date], amounts: list[Decimal], rate: float
) -> Decimal:
    """NAV on one date: PV of the cashflows on or after it, discounted at `rate`."""
    one_plus_rate = 1 + Decimal(str(rate))
    nav = Decimal(0)
    for d, amount in zip(dates, amounts):
        if d >= valuation_date:
            years = Decimal((d - valuation_date).days) / DAYS_IN_YEAR
            nav += amount / one_plus_rate ** years
    return nav


def nav_schedule(cashflows: list[Cashflow], currency: str, rate: float) -> list[NavPoint]:
    """NAV on each cashflow date for one currency, in local currency."""
    in_currency = [cf for cf in cashflows if cf.local_currency == currency]
    if not in_currency:
        raise ValueError(f"No cashflows in {currency}")
    dates = [cf.date for cf in in_currency]
    amounts = [cf.amount_local for cf in in_currency]

    schedule = [
        NavPoint(date=d, currency=currency, nav=nav_at(d, dates, amounts, rate))
        for d in sorted(set(dates))
    ]
    for point in schedule:
        logger.debug("NAV %s %s: %s", currency, point.date, f"{point.nav:,.2f}")
    return schedule


def nav_schedules(
    cashflows: list[Cashflow], irrs: dict[str, float]
) -> dict[str, list[NavPoint]]:
    """NAV schedule for every currency, each discounted at its own IRR."""
    schedules = {}
    for currency, rate in irrs.items():
        schedule = nav_schedule(cashflows, currency, rate)
        if check_schedule(schedule, cashflows):
            logger.info("NAV %s: %d dates, passes sanity checks", currency, len(schedule))
        else:
            logger.warning("NAV %s: FAILS sanity checks", currency)
        schedules[currency] = schedule
    return schedules


def check_schedule(
    schedule: list[NavPoint],
    cashflows: list[Cashflow],
    tolerance: Decimal = NAV_TOLERANCE,
) -> bool:
    """True if NAV(first date) == 0 and NAV(final date) == the exit cashflow."""
    if not schedule:
        return False
    currency = schedule[0].currency
    final_date = schedule[-1].date
    exit_value = sum(
        (cf.amount_local for cf in cashflows
         if cf.local_currency == currency and cf.date == final_date),
        Decimal(0),
    )
    start_ok = abs(schedule[0].nav) <= tolerance
    end_ok = abs(schedule[-1].nav - exit_value) <= tolerance
    logger.debug(
        "NAV %s checks: NAV(0)=%s (ok=%s), NAV(final)=%s vs exit %s (ok=%s)",
        currency, schedule[0].nav, start_ok, schedule[-1].nav, exit_value, end_ok,
    )
    return start_ok and end_ok


# --- FX hedging --------------------------------------------------------------

def roll_dates(start: date, end: date, months: int) -> list[date] | None:
    """Dates on which a new hedge is put on."""
    # TODO
    return None


def add_months(d: date, months: int) -> date | None:
    """Shift a date by whole months (watch out for month-end dates)."""
    # TODO
    return None


def nav_on(schedule: list[NavPoint], d: date) -> Decimal | None:
    """Look up (or work out) the NAV exposure to hedge on a given date."""
    # TODO
    return None


def propose_hedges(
    schedules: dict[str, list[NavPoint]],
    base_currency: str,
    hedge_ratio: float = 1.0,
    tenor_months: int = 3,
) -> list[FxForward] | None:
    """Rolling FX forward trades covering hedge_ratio of each currency's NAV."""
    # TODO
    return None
