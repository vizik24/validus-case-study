"""Calculations: IRR, NAV schedules and FX hedges. No file or console I/O."""

import calendar
import logging
import math
from collections.abc import Collection, Iterable
from datetime import date, timedelta
from decimal import Decimal

import holidays

from models import Cashflow, Direction, FxForward, NavPoint

logger = logging.getLogger(__name__)


# --- Fund data --------------------------------------------------------------

def base_currency(cashflows: list[Cashflow]) -> str:
    """Find the fund's reporting currency, which every cashflow must share.

    :param cashflows: One fund's cashflows.
    :returns: The base currency, e.g. ``"EUR"``.
    :raises ValueError: If there is not exactly one.
    """
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
    """Years from ``start`` to ``end``, actual/365."""
    return (end - start).days / DAYS_IN_YEAR


def npv(rate: float, dates: list[date], amounts: list[Decimal]) -> float:
    """NPV of dated cashflows at an annual rate, as at the earliest date.

    :param rate: Annual rate as a fraction (``0.1`` is 10%).
    :param dates: Date of each cashflow.
    :param amounts: Amount of each cashflow.
    :returns: The NPV.
    :raises ValueError: If ``rate`` is -100% or less.
    """
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
    tolerance: float = 1e-14,
) -> float:
    """Find the IRR (as Excel's ``XIRR``, actual/365) by bisection.

    Bisection can't diverge. No iteration limit is needed because
    ``tolerance`` must be at least the float spacing at the edge of the
    range, the closest the two ends can ever get.

    :param dates: Date of each cashflow.
    :param amounts: Amount of each cashflow.
    :param low: Lowest rate searched.
    :param high: Highest rate searched.
    :param tolerance: Stop when the IRR is known to within this much.
    :returns: The IRR as a fraction (``0.1`` is 10%).
    :raises ValueError: If the inputs are empty, unequal or all one sign;
        no IRR lies in the range; or ``tolerance`` is out of range.
    """
    if len(dates) != len(amounts) or not dates:
        raise ValueError("dates and amounts must be non-empty and the same length")
    if not (any(a > 0 for a in amounts) and any(a < 0 for a in amounts)):
        raise ValueError("IRR needs at least one positive and one negative cashflow")
    smallest = math.ulp(max(abs(low), abs(high)))
    if not smallest <= tolerance < high - low:  # also False for NaN
        raise ValueError(
            f"Tolerance must be from {smallest:.1e} to below {high - low}, got {tolerance}"
        )

    npv_low = npv(low, dates, amounts)
    if npv_low * npv(high, dates, amounts) > 0:
        raise ValueError(f"No IRR found between {low:.0%} and {high:.0%}")

    iteration = 0
    while True:
        iteration += 1
        mid = (low + high) / 2
        npv_mid = npv(mid, dates, amounts)
        if npv_mid == 0 or (high - low) / 2 < tolerance:
            logger.debug("XIRR converged to %.16f after %d iterations", mid, iteration)
            return mid
        if npv_low * npv_mid < 0:
            high = mid
        else:
            low, npv_low = mid, npv_mid


def currency_irrs(cashflows: list[Cashflow]) -> dict[str, float]:
    """IRR per currency on local amounts, i.e. before any FX effect.

    :param cashflows: One fund's cashflows.
    :returns: IRR per local currency.
    """
    irrs = {}
    for currency in sorted({cf.local_currency for cf in cashflows}):
        in_currency = [cf for cf in cashflows if cf.local_currency == currency]
        dates = [cf.date for cf in in_currency]
        amounts = [cf.amount_local for cf in in_currency]
        irrs[currency] = xirr(dates, amounts)
        logger.info("IRR %s (local): %.4f%%", currency, irrs[currency] * 100)
    return irrs


def fund_irr(cashflows: list[Cashflow]) -> float:
    """IRR of all cashflows pooled in the base currency (money-weighted).

    The file's base amounts use fixed FX rates, so this includes no FX moves.

    :param cashflows: One fund's cashflows.
    :returns: The fund IRR.
    """
    dates = [cf.date for cf in cashflows]
    amounts = [cf.amount_base for cf in cashflows]
    irr = xirr(dates, amounts)
    logger.info("IRR fund (base): %.4f%%", irr * 100)
    return irr


# --- NAV ---------------------------------------------------------------------

NAV_TOLERANCE = Decimal("0.01")  # one cent / penny


def nav_at(
    valuation_date: date,
    dates: list[date],
    amounts: list[Decimal],
    rate: float,
    include_same_day: bool = True,
) -> Decimal:
    """NAV on one date: the PV of the cashflows still to come, in Decimal.

    NAV(t) = sum over s >= t of CF_s / (1 + r) ^ ((s - t) / 365). Including
    the cashflow on t is what makes NAV(0) = 0 and NAV(final) = exit value.

    :param valuation_date: Date to value on.
    :param dates: Date of each cashflow.
    :param amounts: Amount of each cashflow.
    :param rate: Discount rate, normally the IRR.
    :param include_same_day: Count a cashflow on ``valuation_date`` as still
        to come (the default); if false, value just after it.
    :returns: The NAV.
    """
    one_plus_rate = 1 + Decimal(str(rate))
    nav = Decimal(0)
    for cashflow_date, amount in zip(dates, amounts):
        if cashflow_date > valuation_date or (include_same_day and cashflow_date == valuation_date):
            years = Decimal((cashflow_date - valuation_date).days) / DAYS_IN_YEAR
            nav += amount / one_plus_rate ** years
    return nav


def nav_schedule_by_currency(cashflows: list[Cashflow], currency: str, rate: float) -> list[NavPoint]:
    """NAV on each cashflow date for one currency, in local currency.

    :param cashflows: One fund's cashflows; other currencies are ignored.
    :param currency: Currency to value.
    :param rate: Discount rate, normally that currency's IRR.
    :returns: One NavPoint per date, in date order.
    :raises ValueError: If there are no cashflows in ``currency``.
    """
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
    """NAV schedule per currency at its own IRR; a failed check is a warning.

    :param cashflows: One fund's cashflows.
    :param irrs: IRR per currency.
    :returns: NAV schedule keyed by currency code.
    """
    schedules = {}
    for currency, rate in irrs.items():
        schedule = nav_schedule_by_currency(cashflows, currency, rate)
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
    """Check NAV(first date) = 0 and NAV(final date) = the exit cashflows.

    :param schedule: One currency's NAV schedule, in date order.
    :param cashflows: The cashflows it was built from.
    :param tolerance: Largest difference allowed.
    :returns: True if both checks pass.
    """
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
# The brief's strategy: rolling FX forwards (3-month, 100% of NAV by default)
# selling each non-base currency for the base currency.

PAIR_PRIORITY = ["EUR", "GBP", "USD"]
NOTIONAL_STEP = Decimal("0.01")


def add_months(d: date, months: int) -> date:
    """Shift a date by whole months, keeping month-ends (30 Sep + 3 = 31 Dec).

    Other days are clamped to the month's length (30 Jan + 1 = 28 Feb).
    """
    month_index = d.year * 12 + (d.month - 1) + months
    year, month = divmod(month_index, 12)
    month += 1
    last_day = calendar.monthrange(year, month)[1]
    if d.day == calendar.monthrange(d.year, d.month)[1]:
        return date(year, month, last_day)
    return date(year, month, min(d.day, last_day))


def roll_dates(start: date, end: date, months: int) -> list[date]:
    """Unadjusted roll dates: ``start``, ``start + months``, ... before ``end``.

    Each is counted from ``start``, so a short month can't shift later dates.

    :raises ValueError: If ``months`` is not positive.
    """
    if months <= 0:
        raise ValueError(f"Tenor must be a positive number of months, got {months}")
    dates = []
    step = 0
    while (d := add_months(start, step * months)) < end:
        dates.append(d)
        step += 1
    return dates


# Settlement holiday calendar for each currency (weekends are always closed):
# EUR: TARGET2, the ECB's payment system. GBP: England & Wales bank holidays
# (London). USD: Federal Reserve, which settles USD payments.
HOLIDAY_CALENDARS = {
    "EUR": lambda years: holidays.financial_holidays("XECB", years=years),
    "GBP": lambda years: holidays.country_holidays("GB", subdiv="ENG", years=years),
    "USD": lambda years: holidays.financial_holidays("XFED", years=years),
}


def settlement_holidays(currencies: Iterable[str], years: Iterable[int]) -> set[date]:
    """Holidays (not weekends) on which any of the currencies can't settle.

    :raises ValueError: If a currency has no calendar.
    """
    years = list(years)
    closed: set[date] = set()
    for currency in currencies:
        if currency not in HOLIDAY_CALENDARS:
            raise ValueError(f"No holiday calendar for {currency}")
        closed.update(HOLIDAY_CALENDARS[currency](years))
    return closed


def is_business_day(d: date, closed: Collection[date]) -> bool:
    """True if ``d`` is a weekday and not in ``closed``."""
    return d.weekday() < 5 and d not in closed


def adjust_business_day(d: date, closed: Collection[date]) -> date:
    """Modified following: the next business day, unless that is next month,
    then the previous one (Sat 30 Sep 2028 -> Fri 29 Sep, not Mon 2 Oct).
    """
    following = d
    while not is_business_day(following, closed):
        following += timedelta(days=1)
    if following.month == d.month:
        return following
    preceding = d
    while not is_business_day(preceding, closed):
        preceding -= timedelta(days=1)
    return preceding


def exposure_on(d: date, dates: list[date], amounts: list[Decimal], rate: float) -> Decimal:
    """NAV to hedge from ``d``: the value just *after* that day's cashflow.

    Including the day's cashflow would give 0 on the investment date and
    leave the first period unhedged.
    """
    return nav_at(d, dates, amounts, rate, include_same_day=False)


def currency_pair(a: str, b: str) -> str:
    """Market-convention pair name, e.g. ``("GBP", "EUR")`` -> ``"EURGBP"``."""
    first, second = sorted([a, b], key=PAIR_PRIORITY.index)
    return first + second


def propose_hedges(
    cashflows: list[Cashflow],
    irrs: dict[str, float],
    base_currency: str,
    hedge_ratio: float = 1.0,
    tenor_months: int = 3,
) -> list[FxForward]:
    """Propose rolling FX forwards hedging each non-base currency's NAV.

    ``direction`` refers to the notional currency: "Sell GBP 100m" in EURGBP
    means the fund delivers GBP and receives EUR.

    :param cashflows: One fund's cashflows.
    :param irrs: IRR per currency, used to value the exposure.
    :param base_currency: The fund's reporting currency, which is not hedged.
    :param hedge_ratio: Fraction of NAV to hedge.
    :param tenor_months: Length of each forward in months.
    :returns: The trades, by currency and date.
    :raises ValueError: If ``hedge_ratio`` is negative or ``tenor_months``
        is not positive.
    """
    if hedge_ratio < 0:
        raise ValueError(f"Hedge ratio must not be negative, got {hedge_ratio}")
    ratio = Decimal(str(hedge_ratio))

    trades: list[FxForward] = []
    for currency in sorted(irrs):
        if currency == base_currency:
            logger.info("Hedge %s: base currency, no FX risk", currency)
            continue
        in_currency = [cf for cf in cashflows if cf.local_currency == currency]
        dates = [cf.date for cf in in_currency]
        amounts = [cf.amount_local for cf in in_currency]
        final_date = max(dates)
        pair = currency_pair(currency, base_currency)
        # +1 year so a final date in late December can roll into January.
        closed = settlement_holidays(
            {currency, base_currency}, range(min(dates).year, final_date.year + 2)
        )

        # Unadjusted period boundaries: each roll date, then the final date.
        # Forward i trades on boundary i and delivers on boundary i + 1.
        boundaries = roll_dates(min(dates), final_date, tenor_months) + [final_date]
        adjusted = []
        for d in boundaries:
            moved = adjust_business_day(d, closed)
            if moved != d:
                logger.debug(
                    "%s: %s (%s) is not a business day, moved to %s (%s)",
                    pair, d, d.strftime("%a"), moved, moved.strftime("%a"),
                )
            adjusted.append(moved)

        for roll_date, trade_date, delivery_date in zip(boundaries, adjusted, adjusted[1:]):
            # Sized on the unadjusted roll date: the period's exposure starts
            # after the quarter-end cashflow even if the trade is done a day
            # or two earlier, so a coupon paid at quarter-end is not hedged
            # for the whole period.
            exposure = exposure_on(roll_date, dates, amounts, irrs[currency])
            notional = (abs(exposure) * ratio).quantize(NOTIONAL_STEP)
            if notional == 0:
                logger.debug("Hedge %s %s: no exposure, no trade", currency, trade_date)
                continue
            trade = FxForward(
                currency_pair=pair,
                trade_date=trade_date,
                delivery_date=delivery_date,
                # Long the foreign asset -> sell it forward; a net liability -> buy.
                direction=Direction.SELL if exposure > 0 else Direction.BUY,
                notional_currency=currency,
                notional_amount=notional,
            )
            logger.debug(
                "Hedge %s %s -> %s: %s %s %s",
                pair, trade.trade_date, trade.delivery_date,
                trade.direction.value, currency, f"{notional:,.2f}",
            )
            trades.append(trade)
        logger.info(
            "Hedge %s: %d forward(s) in %s, %d-month tenor, %s%% of NAV",
            currency, sum(t.notional_currency == currency for t in trades),
            pair, tenor_months, ratio * 100,
        )
    return trades
