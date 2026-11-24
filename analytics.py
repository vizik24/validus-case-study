"""Calculations: IRR, NAV schedules and FX hedges.

Every function here takes validated :class:`~models.Cashflow` objects and
returns numbers or models. None of them reads or writes files, so they can be
tested without any input workbook.
"""

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
    """Find the fund's reporting currency from its cashflows.

    :param cashflows: One fund's cashflows.
    :returns: The base currency they all share, e.g. ``"EUR"``.
    :raises ValueError: If the cashflows name more than one base currency,
        or there are no cashflows.
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
    """Time between two dates in years, on an actual/365 day count.

    :param start: Earlier date.
    :param end: Later date.
    :returns: ``(end - start).days / 365``; negative if ``end`` is earlier.
    """
    return (end - start).days / DAYS_IN_YEAR


def npv(rate: float, dates: list[date], amounts: list[Decimal]) -> float:
    """Net present value of dated cashflows, as at the earliest date.

    .. math:: NPV = \\sum_i CF_i / (1 + r)^{(d_i - d_0) / 365}

    :param rate: Annual rate as a fraction (``0.1`` is 10%).
    :param dates: Date of each cashflow, in any order.
    :param amounts: Amount of each cashflow, matching ``dates``.
    :returns: The NPV, as a float.
    :raises ValueError: If ``rate`` is -100% or less, where the discount
        factor is undefined.
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
    """Find the annual rate at which the NPV of dated cashflows is 0.

    Matches Excel's ``XIRR`` (actual dates, actual/365). Bisection halves
    ``low`` to ``high``, keeping the half where the NPV changes sign; unlike
    Newton's method it cannot diverge, and one outflow followed by inflows
    has exactly one IRR. No iteration limit is needed: ``tolerance`` must be
    at least the float spacing (``math.ulp``) at the edge of the range, the
    closest the two ends can ever get, so it is always reached.

    :param dates: Date of each cashflow.
    :param amounts: Amount of each cashflow, matching ``dates``.
    :param low: Lowest rate searched.
    :param high: Highest rate searched.
    :param tolerance: Stop when the IRR is known to within this much; the
        default ``1e-14`` keeps NAVs accurate to well under a cent.
    :returns: The IRR as a fraction (``0.1`` is 10%).
    :raises ValueError: If the inputs are empty, unequal in length or all
        one sign; no IRR lies in the range; or ``tolerance`` is not finite,
        below the float spacing, or not smaller than the range.
    """
    if len(dates) != len(amounts) or not dates:
        raise ValueError("dates and amounts must be non-empty and the same length")
    if not (any(a > 0 for a in amounts) and any(a < 0 for a in amounts)):
        raise ValueError("IRR needs at least one positive and one negative cashflow")
    smallest = math.ulp(max(abs(low), abs(high)))
    if not math.isfinite(tolerance):
        raise ValueError(f"Tolerance must be a finite number, got {tolerance}")
    if tolerance < smallest:
        raise ValueError(
            f"Tolerance must be at least {smallest:.2e}, the float precision "
            f"between {low} and {high}; got {tolerance:.2e}"
        )
    if tolerance >= high - low:
        raise ValueError(f"Tolerance must be smaller than the range {high - low}, got {tolerance}")

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
    """Calculate the IRR of each currency's positions in that currency.

    Uses the local-currency amounts, so each IRR is the return of the
    positions themselves, before any FX effect.

    :param cashflows: One fund's cashflows.
    :returns: IRR per local currency, as fractions, in currency order.
    :raises ValueError: If a currency's IRR cannot be found (see
        :func:`xirr`).
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
    """Calculate the IRR of the whole fund in its base currency.

    Uses the base-currency amounts from the file. Those were converted at
    fixed FX rates, so this IRR blends the currency IRRs and does not
    include any FX moves.

    :param cashflows: One fund's cashflows.
    :returns: The fund IRR as a fraction.
    :raises ValueError: If the IRR cannot be found (see :func:`xirr`).
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
    """Calculate the NAV on one date: the PV of the cashflows still to come.

    Calculated in Decimal; ``rate`` is converted via ``str`` so the float's
    binary noise does not leak in.

    :param valuation_date: Date to value on.
    :param dates: Date of each cashflow.
    :param amounts: Amount of each cashflow, matching ``dates``.
    :param rate: Annual discount rate as a fraction, normally the IRR.
    :param include_same_day: If true (the default, and the brief's
        convention), a cashflow on ``valuation_date`` counts as still to
        come, giving the value just *before* that day's cashflow. If false,
        it counts as already paid, giving the value just *after*.
    :returns: The NAV, in the same currency as ``amounts``.
    """
    one_plus_rate = 1 + Decimal(str(rate))
    nav = Decimal(0)
    for cashflow_date, amount in zip(dates, amounts):
        if cashflow_date > valuation_date or (include_same_day and cashflow_date == valuation_date):
            years = Decimal((cashflow_date - valuation_date).days) / DAYS_IN_YEAR
            nav += amount / one_plus_rate ** years
    return nav


def nav_schedule_by_currency(cashflows: list[Cashflow], currency: str, rate: float) -> list[NavPoint]:
    """Build the NAV schedule for one currency. Filter the rows keeping only the
    cashflows in the given currency. Apply the nav_at function to each date
    a cashflow is present for.

    :param cashflows: One fund's cashflows; other currencies are ignored.
    :param currency: Local currency to value.
    :param rate: Discount rate, normally that currency's IRR.
    :returns: One :class:`~models.NavPoint` per distinct cashflow date, in
        date order, with the NAV in local currency.
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
    """Build the NAV schedule for every currency, each at its own IRR.

    Each schedule is checked with :func:`check_schedule`. A failure is
    logged as a warning rather than raised, so the results can still be
    inspected.

    :param cashflows: One fund's cashflows.
    :param irrs: IRR per currency, from :func:`currency_irrs`.
    :returns: NAV schedule per currency. Dictionary where each key is currency code.
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
    """Check a NAV schedule against the brief's two sanity checks.

    The checks are NAV(first date) = 0, and NAV(final date) = the exit
    value, i.e. the cashflows on the final date.

    :param schedule: One currency's NAV schedule, in date order.
    :param cashflows: The cashflows the schedule was built from.
    :param tolerance: Largest difference allowed, in local currency.
    :returns: True if both checks pass; False if either fails or the
        schedule is empty.
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
# Strategy from the brief: 3-month rolling FX forwards covering 100% of NAV.
#
# - Only non-base currencies carry FX risk. The fund reports in EUR and is long
#   its GBP and USD positions, so it sells those currencies forward for EUR.
# - On each roll date a forward is traded with delivery `tenor_months` later.
#   At delivery it settles and the next one is traded, sized to the new NAV.
# - Notional = hedge_ratio x the NAV just AFTER that day's cashflow: the value
#   still exposed over the coming period. (NAV including the day's cashflow is
#   0 on the investment date, which would leave the first period unhedged.)
# - Roll dates keep month-ends as month-ends, so they line up with the
#   quarter-end cashflow dates. The last forward delivers on the final
#   cashflow date.
# - Trade and delivery dates are then moved to good business days with the
#   modified-following convention, using the holidays of both currencies in
#   the pair. Each forward still delivers on the day the next one is traded,
#   and is still sized on the unadjusted roll date.

# Market quoting order: the currency earlier in this list is the base of the
# pair (EURGBP, EURUSD, GBPUSD). Unlisted currencies come after, alphabetically.
PAIR_PRIORITY = ["EUR", "GBP", "AUD", "NZD", "USD", "CAD", "CHF", "JPY"]
NOTIONAL_STEP = Decimal("0.01")


def add_months(d: date, months: int) -> date:
    """Shift a date by whole months.

    A month-end stays a month-end: 30 Sep + 3 months is 31 Dec. Any other
    day is clamped to the length of the target month: 30 Jan + 1 month is
    28 Feb (29 in a leap year).

    :param d: Starting date.
    :param months: Months to add; may be negative.
    :returns: The shifted date.
    """
    month_index = d.year * 12 + (d.month - 1) + months
    year, month = divmod(month_index, 12)
    month += 1
    last_day = calendar.monthrange(year, month)[1]
    if d.day == calendar.monthrange(d.year, d.month)[1]:
        return date(year, month, last_day)
    return date(year, month, min(d.day, last_day))


def roll_dates(start: date, end: date, months: int) -> list[date]:
    """List the dates on which a new hedge is traded.

    The dates are ``start``, ``start + months``, ``start + 2 * months`` and
    so on, up to but not including ``end``. Each is counted from ``start``
    rather than from the previous roll, so a short month can't pull later
    dates off their day. Dates are unadjusted; see
    :func:`adjust_business_day`.

    :param start: First roll date, normally the first cashflow date.
    :param end: Final date; no roll is made on or after it.
    :param months: Tenor of each forward, in months.
    :returns: The roll dates, in order.
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
    """Collect the holidays on which any of the currencies cannot settle.

    Uses the calendars in :data:`HOLIDAY_CALENDARS`. Weekends are not
    included; :func:`is_business_day` handles those.

    :param currencies: Currencies whose calendars to combine, e.g. both
        sides of a pair.
    :param years: Calendar years to cover.
    :returns: Every holiday in any of the calendars.
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
    """Check whether a date is a weekday and not a holiday.

    :param d: Date to check.
    :param closed: Holidays, from :func:`settlement_holidays`.
    :returns: True if trades can settle on ``d``.
    """
    return d.weekday() < 5 and d not in closed


def adjust_business_day(d: date, closed: Collection[date]) -> date:
    """Move a date to a business day by the modified-following rule.

    The date moves to the next business day, unless that falls in the next
    month, in which case it moves to the previous business day instead. This
    keeps month-end rolls in their own month: Sat 30 Sep 2028 becomes
    Fri 29 Sep, not Mon 2 Oct.

    :param d: Unadjusted date.
    :param closed: Holidays, from :func:`settlement_holidays`.
    :returns: ``d`` if it is already a business day, otherwise the adjusted
        date.
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
    """Calculate the NAV to hedge from date ``d``.

    This is the NAV just *after* that day's cashflow: the PV of the
    cashflows strictly after ``d``, which is the value still exposed over
    the coming period. (The NAV including the day's cashflow is 0 on the
    investment date, which would leave the first period unhedged.)

    :param d: Roll date.
    :param dates: Date of each cashflow.
    :param amounts: Amount of each cashflow, matching ``dates``.
    :param rate: Discount rate, normally the currency's IRR.
    :returns: The exposure, in the same currency as ``amounts``.
    """
    return nav_at(d, dates, amounts, rate, include_same_day=False)


def currency_pair(a: str, b: str) -> str:
    """Name the pair of two currencies in market convention.

    The currency earlier in :data:`PAIR_PRIORITY` comes first, so
    ``("GBP", "EUR")`` gives ``"EURGBP"``. Currencies not in the list come
    after those that are, alphabetically.

    :param a: One currency.
    :param b: The other currency.
    :returns: The six-letter pair, e.g. ``"EURUSD"``.
    """
    def rank(code: str) -> tuple[int, str]:
        return (PAIR_PRIORITY.index(code) if code in PAIR_PRIORITY else len(PAIR_PRIORITY), code)
    first, second = sorted([a, b], key=rank)
    return first + second


def propose_hedges(
    cashflows: list[Cashflow],
    irrs: dict[str, float],
    base_currency: str,
    hedge_ratio: float = 1.0,
    tenor_months: int = 3,
) -> list[FxForward]:
    """Propose rolling FX forwards hedging each foreign currency's NAV.

    For every currency other than the base currency, a forward is traded on
    each roll date and delivers on the next, the last one delivering on the
    final cashflow date. Dates are moved to business days on both
    currencies' calendars. Each notional is ``hedge_ratio`` times the
    exposure on the *unadjusted* roll date (see :func:`exposure_on`), so a
    coupon paid at quarter-end is not hedged when a trade is moved a day or
    two earlier.

    ``direction`` refers to the notional currency: "Sell GBP 100m" in
    EURGBP means the fund delivers 100m GBP and receives EUR at the forward
    rate. Periods with no exposure get no trade.

    :param cashflows: One fund's cashflows.
    :param irrs: IRR per currency, used to value the exposure.
    :param base_currency: The fund's reporting currency, which is not
        hedged.
    :param hedge_ratio: Fraction of NAV to hedge, e.g. ``1.0`` for 100%.
    :param tenor_months: Length of each forward, in months.
    :returns: The trades, grouped by currency and in date order.
    :raises ValueError: If ``hedge_ratio`` is negative, ``tenor_months`` is
        not positive, or a currency has no holiday calendar.
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
