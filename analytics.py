"""Calculations: IRR, NAV schedules and FX hedges. Data in, numbers out, no I/O."""

import calendar
import logging
from collections.abc import Collection, Iterable
from datetime import date, timedelta
from decimal import Decimal

import holidays

from models import Cashflow, Direction, FxForward, NavPoint

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


def currency_irrs(cashflows: list[Cashflow]) -> dict[str, float]:
    """IRR per local currency, using local-currency amounts."""
    irrs = {}
    for currency in sorted({cf.local_currency for cf in cashflows}):
        in_currency = [cf for cf in cashflows if cf.local_currency == currency]
        dates = [cf.date for cf in in_currency]
        amounts = [cf.amount_local for cf in in_currency]
        irrs[currency] = xirr(dates, amounts)
        logger.info("IRR %s (local): %.4f%%", currency, irrs[currency] * 100)
    return irrs


def fund_irr(cashflows: list[Cashflow]) -> float:
    """IRR for the whole fund, using base-currency amounts."""
    dates = [cf.date for cf in cashflows]
    amounts = [cf.amount_base for cf in cashflows]
    irr = xirr(dates, amounts)
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
    valuation_date: date,
    dates: list[date],
    amounts: list[Decimal],
    rate: float,
    include_same_day: bool = True,
) -> Decimal:
    """NAV on one date: PV of the remaining cashflows, discounted at `rate`.

    By default a cashflow on the valuation date itself counts as remaining
    (the brief's convention). With include_same_day=False it is treated as
    already paid, giving the value just after that day's cashflow.
    """
    one_plus_rate = 1 + Decimal(str(rate))
    nav = Decimal(0)
    for d, amount in zip(dates, amounts):
        if d > valuation_date or (include_same_day and d == valuation_date):
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

    A month-end date stays a month-end (30 Sep + 3 = 31 Dec); any other day is
    clamped to the length of the target month (31 Jan + 1 = 28/29 Feb).
    """
    month_index = d.year * 12 + (d.month - 1) + months
    year, month = divmod(month_index, 12)
    month += 1
    last_day = calendar.monthrange(year, month)[1]
    if d.day == calendar.monthrange(d.year, d.month)[1]:
        return date(year, month, last_day)
    return date(year, month, min(d.day, last_day))


def roll_dates(start: date, end: date, months: int) -> list[date]:
    """Dates on which a new hedge is put on: start, start + months, ... before end.

    Each date is counted from `start` rather than from the previous roll, so a
    short month can't pull later dates off their day.
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
    """Dates on which any of the currencies cannot settle (excluding weekends)."""
    years = list(years)
    closed: set[date] = set()
    for currency in currencies:
        if currency not in HOLIDAY_CALENDARS:
            raise ValueError(f"No holiday calendar for {currency}")
        closed.update(HOLIDAY_CALENDARS[currency](years))
    return closed


def is_business_day(d: date, closed: Collection[date]) -> bool:
    return d.weekday() < 5 and d not in closed


def adjust_business_day(d: date, closed: Collection[date]) -> date:
    """Modified following: the next business day, unless that falls in the next
    month, in which case the previous business day. Keeps month-end rolls in
    their own month (Sat 30 Sep -> Fri 29 Sep, not Mon 2 Oct).
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
    """The NAV to hedge on date d: PV of the cashflows strictly after d."""
    return nav_at(d, dates, amounts, rate, include_same_day=False)


def currency_pair(a: str, b: str) -> str:
    """Market-convention pair name for two currencies, e.g. ("GBP", "EUR") -> "EURGBP"."""
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
    """Rolling FX forward trades covering hedge_ratio of each foreign currency's NAV.

    `direction` refers to the notional currency: Sell GBP 100m in EURGBP means
    the fund delivers 100m GBP and receives EUR at the forward rate.
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
