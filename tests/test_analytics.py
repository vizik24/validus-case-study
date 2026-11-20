import datetime as dt
import unittest
from decimal import Decimal
from pathlib import Path

import analytics
import excel_io
from models import Cashflow, Direction, FxForward, NavPoint

D0 = dt.date(2025, 9, 30)
DATA_FILE = Path(__file__).resolve().parent.parent / "data" / "CST_Case_Study_data__1_.xlsx"


def make_cashflow(
    base_currency: str = "EUR",
    *,
    date: dt.date = D0,
    currency: str = "GBP",
    amount: str | int = 1,
    fx: str = "1",
    cashflow_type: str = "Interest",
) -> Cashflow:
    amount = Decimal(amount)
    return Cashflow(
        id=1, fund_name="Fund I", date=date, cashflow_type=cashflow_type,
        local_currency=currency, amount_local=amount, amount_base=amount * Decimal(fx),
        base_currency=base_currency,
    )


def days(n: int) -> dt.date:
    return D0 + dt.timedelta(days=n)


def annual_bond(currency: str = "GBP", coupon: int = 10, years: int = 3, fx: str = "1"):
    """Invest 100, receive `coupon` every 365 days, get 100 back with the last coupon.

    Exactly 365 days apart, so under actual/365 the IRR is exactly coupon%.
    """
    flows = [make_cashflow(date=D0, currency=currency, amount=-100, fx=fx, cashflow_type="Investment")]
    for year in range(1, years + 1):
        flows.append(make_cashflow(date=days(365 * year), currency=currency, amount=coupon, fx=fx))
    flows.append(make_cashflow(
        date=days(365 * years), currency=currency, amount=100, fx=fx,
        cashflow_type="Principal Repayment",
    ))
    return flows


def dates_and_amounts(flows):
    return [cf.date for cf in flows], [cf.amount_local for cf in flows]


class TestBaseCurrency(unittest.TestCase):
    def test_single_base_currency(self):
        self.assertEqual(analytics.base_currency([make_cashflow(), make_cashflow()]), "EUR")

    def test_mixed_base_currencies_fail(self):
        with self.assertRaises(ValueError):
            analytics.base_currency([make_cashflow("EUR"), make_cashflow("USD")])

    def test_no_cashflows_fail(self):
        with self.assertRaises(ValueError):
            analytics.base_currency([])


class TestYearFraction(unittest.TestCase):
    def test_365_days_is_one_year(self):
        self.assertEqual(analytics.year_fraction(D0, days(365)), 1.0)

    def test_leap_year_is_actual_over_365(self):
        self.assertAlmostEqual(
            analytics.year_fraction(dt.date(2028, 1, 1), dt.date(2029, 1, 1)), 366 / 365
        )


class TestNpv(unittest.TestCase):
    def test_zero_rate_is_plain_sum(self):
        self.assertAlmostEqual(
            analytics.npv(0.0, [D0, days(100)], [Decimal(-100), Decimal(30)]), -70.0
        )

    def test_discounts_one_year(self):
        self.assertAlmostEqual(
            analytics.npv(0.10, [D0, days(365)], [Decimal(-100), Decimal(110)]), 0.0
        )

    def test_rate_at_or_below_minus_100_percent_fails(self):
        with self.assertRaises(ValueError):
            analytics.npv(-1.0, [D0], [Decimal(1)])


class TestXirr(unittest.TestCase):
    def test_one_year(self):
        self.assertAlmostEqual(analytics.xirr([D0, days(365)], [Decimal(-100), Decimal(110)]), 0.10)

    def test_two_years_compounds(self):
        self.assertAlmostEqual(analytics.xirr([D0, days(730)], [Decimal(-100), Decimal(121)]), 0.10)

    def test_negative_return(self):
        self.assertAlmostEqual(analytics.xirr([D0, days(365)], [Decimal(-100), Decimal(90)]), -0.10)

    def test_annual_bond_at_par_returns_its_coupon(self):
        self.assertAlmostEqual(analytics.xirr(*dates_and_amounts(annual_bond(coupon=10))), 0.10)

    def test_order_of_cashflows_does_not_matter(self):
        dates, amounts = dates_and_amounts(annual_bond())
        self.assertAlmostEqual(analytics.xirr(dates[::-1], amounts[::-1]), 0.10)

    def test_irregular_dates_give_zero_npv(self):
        dates = [D0, days(45), days(200), days(600)]
        amounts = [Decimal(-1000), Decimal(50), Decimal(80), Decimal(1000)]
        rate = analytics.xirr(dates, amounts)
        self.assertAlmostEqual(analytics.npv(rate, dates, amounts), 0.0, places=6)

    def test_all_same_sign_fails(self):
        with self.assertRaises(ValueError):
            analytics.xirr([D0, days(365)], [Decimal(100), Decimal(110)])

    def test_mismatched_lengths_fail(self):
        with self.assertRaises(ValueError):
            analytics.xirr([D0, days(365)], [Decimal(-100)])


class TestCurrencyAndFundIrr(unittest.TestCase):
    def test_each_currency_uses_its_own_cashflows(self):
        flows = annual_bond("GBP", coupon=10) + annual_bond("USD", coupon=5)
        irrs = analytics.currency_irrs(flows)
        self.assertEqual(sorted(irrs), ["GBP", "USD"])
        self.assertAlmostEqual(irrs["GBP"], 0.10)
        self.assertAlmostEqual(irrs["USD"], 0.05)

    def test_fund_irr_uses_base_amounts_and_blends_currencies(self):
        flows = annual_bond("GBP", coupon=10, fx="1.2") + annual_bond("USD", coupon=5, fx="0.8")
        irr = analytics.fund_irr(flows)
        self.assertGreater(irr, 0.05)
        self.assertLess(irr, 0.10)

    def test_constant_fx_single_currency_fund_matches_local_irr(self):
        flows = annual_bond("GBP", coupon=10, fx="1.15")
        self.assertAlmostEqual(analytics.fund_irr(flows), analytics.currency_irrs(flows)["GBP"])


class TestNavAt(unittest.TestCase):
    def test_discounts_a_future_cashflow(self):
        nav = analytics.nav_at(D0, [days(365)], [Decimal(110)], 0.10)
        self.assertAlmostEqual(nav, Decimal(100), places=8)

    def test_includes_the_cashflow_on_the_valuation_date(self):
        self.assertEqual(analytics.nav_at(D0, [D0], [Decimal(50)], 0.10), Decimal(50))

    def test_ignores_past_cashflows(self):
        nav = analytics.nav_at(days(365), [D0, days(365)], [Decimal(-100), Decimal(110)], 0.10)
        self.assertEqual(nav, Decimal(110))

    def test_returns_decimal(self):
        self.assertIsInstance(analytics.nav_at(D0, [days(10)], [Decimal(1)], 0.05), Decimal)


class TestNavSchedule(unittest.TestCase):
    def setUp(self):
        self.flows = annual_bond("GBP", coupon=10, years=3)
        self.schedule = analytics.nav_schedule(self.flows, "GBP", 0.10)

    def test_one_point_per_cashflow_date_in_order(self):
        self.assertEqual([p.date for p in self.schedule], [D0, days(365), days(730), days(1095)])

    def test_nav_zero_at_start(self):
        self.assertAlmostEqual(self.schedule[0].nav, Decimal(0), places=8)

    def test_nav_on_coupon_dates_is_par_plus_coupon(self):
        # Value just before the coupon is paid: 10 now plus 100 of remaining bond.
        for point in self.schedule[1:]:
            self.assertAlmostEqual(point.nav, Decimal(110), places=8)

    def test_only_uses_the_requested_currency(self):
        flows = self.flows + annual_bond("USD", coupon=5)
        self.assertEqual(analytics.nav_schedule(flows, "GBP", 0.10), self.schedule)

    def test_unknown_currency_fails(self):
        with self.assertRaises(ValueError):
            analytics.nav_schedule(self.flows, "JPY", 0.10)

    def test_nav_schedules_uses_each_currencys_irr(self):
        flows = annual_bond("GBP", coupon=10) + annual_bond("USD", coupon=5)
        schedules = analytics.nav_schedules(flows, analytics.currency_irrs(flows))
        for currency, schedule in schedules.items():
            self.assertTrue(analytics.check_schedule(schedule, flows), currency)


class TestCheckSchedule(unittest.TestCase):
    def setUp(self):
        self.flows = annual_bond()
        self.schedule = analytics.nav_schedule(self.flows, "GBP", 0.10)

    def test_valid_schedule_passes(self):
        self.assertTrue(analytics.check_schedule(self.schedule, self.flows))

    def test_wrong_rate_fails_nav_zero_check(self):
        schedule = analytics.nav_schedule(self.flows, "GBP", 0.08)
        self.assertFalse(analytics.check_schedule(schedule, self.flows))

    def test_wrong_exit_value_fails(self):
        last = self.schedule[-1]
        tampered = self.schedule[:-1] + [NavPoint(date=last.date, currency="GBP", nav=last.nav + 1)]
        self.assertFalse(analytics.check_schedule(tampered, self.flows))

    def test_empty_schedule_fails(self):
        self.assertFalse(analytics.check_schedule([], self.flows))


class TestAddMonths(unittest.TestCase):
    def test_mid_month_keeps_day(self):
        self.assertEqual(analytics.add_months(dt.date(2025, 11, 15), 3), dt.date(2026, 2, 15))

    def test_month_end_stays_month_end(self):
        self.assertEqual(analytics.add_months(dt.date(2025, 9, 30), 3), dt.date(2025, 12, 31))
        self.assertEqual(analytics.add_months(dt.date(2025, 12, 31), 3), dt.date(2026, 3, 31))
        self.assertEqual(analytics.add_months(dt.date(2026, 3, 31), 3), dt.date(2026, 6, 30))

    def test_day_clamped_to_short_month(self):
        self.assertEqual(analytics.add_months(dt.date(2026, 1, 30), 1), dt.date(2026, 2, 28))

    def test_leap_year_february(self):
        self.assertEqual(analytics.add_months(dt.date(2028, 1, 31), 1), dt.date(2028, 2, 29))
        self.assertEqual(analytics.add_months(dt.date(2027, 11, 30), 3), dt.date(2028, 2, 29))

    def test_zero_and_negative_months(self):
        self.assertEqual(analytics.add_months(D0, 0), D0)
        self.assertEqual(analytics.add_months(dt.date(2026, 3, 31), -1), dt.date(2026, 2, 28))


class TestRollDates(unittest.TestCase):
    def test_quarterly_rolls_over_five_years(self):
        rolls = analytics.roll_dates(D0, dt.date(2030, 9, 30), 3)
        self.assertEqual(len(rolls), 20)
        self.assertEqual(rolls[0], D0)
        self.assertEqual(rolls[1], dt.date(2025, 12, 31))
        self.assertEqual(rolls[-1], dt.date(2030, 6, 30))

    def test_end_date_is_not_a_roll(self):
        self.assertNotIn(dt.date(2026, 3, 31), analytics.roll_dates(D0, dt.date(2026, 3, 31), 3))

    def test_counted_from_start_so_day_does_not_drift(self):
        # 31 Jan -> 28 Feb -> must come back to 31 Mar, not 28 Mar.
        rolls = analytics.roll_dates(dt.date(2026, 1, 31), dt.date(2026, 4, 30), 1)
        self.assertEqual(rolls, [dt.date(2026, 1, 31), dt.date(2026, 2, 28), dt.date(2026, 3, 31)])

    def test_start_on_or_after_end_gives_no_rolls(self):
        self.assertEqual(analytics.roll_dates(D0, D0, 3), [])

    def test_non_positive_tenor_fails(self):
        with self.assertRaises(ValueError):
            analytics.roll_dates(D0, days(365), 0)


class TestBusinessDays(unittest.TestCase):
    YEARS = range(2028, 2031)

    def closed(self, *currencies):
        return analytics.settlement_holidays(currencies, self.YEARS)

    def test_weekends_are_never_business_days(self):
        self.assertFalse(analytics.is_business_day(dt.date(2028, 9, 30), set()))  # Saturday
        self.assertFalse(analytics.is_business_day(dt.date(2028, 10, 1), set()))  # Sunday
        self.assertTrue(analytics.is_business_day(dt.date(2028, 9, 29), set()))   # Friday

    def test_currency_calendars(self):
        good_friday, juneteenth = dt.date(2029, 3, 30), dt.date(2029, 6, 19)
        self.assertFalse(analytics.is_business_day(good_friday, self.closed("EUR")))
        self.assertFalse(analytics.is_business_day(good_friday, self.closed("GBP")))
        self.assertTrue(analytics.is_business_day(good_friday, self.closed("USD")))
        self.assertFalse(analytics.is_business_day(juneteenth, self.closed("USD")))
        self.assertTrue(analytics.is_business_day(juneteenth, self.closed("EUR", "GBP")))

    def test_unknown_currency_fails(self):
        with self.assertRaises(ValueError):
            analytics.settlement_holidays(["JPY"], self.YEARS)

    def test_business_day_is_unchanged(self):
        self.assertEqual(analytics.adjust_business_day(D0, set()), D0)

    def test_mid_month_weekend_rolls_forward(self):
        self.assertEqual(analytics.adjust_business_day(dt.date(2028, 9, 16), set()), dt.date(2028, 9, 18))

    def test_holiday_rolls_forward(self):
        self.assertEqual(
            analytics.adjust_business_day(dt.date(2029, 1, 1), self.closed("EUR")), dt.date(2029, 1, 2)
        )

    def test_month_end_weekend_rolls_back(self):
        self.assertEqual(analytics.adjust_business_day(dt.date(2028, 9, 30), set()), dt.date(2028, 9, 29))
        # Following would be 2 Jan (1 Jan is a holiday), crossing the year.
        self.assertEqual(
            analytics.adjust_business_day(dt.date(2028, 12, 31), self.closed("EUR")), dt.date(2028, 12, 29)
        )

    def test_rolls_back_past_good_friday_when_both_sides_closed(self):
        # Sat 31 Mar 2029: Easter Monday ahead, Good Friday behind.
        sat = dt.date(2029, 3, 31)
        self.assertEqual(analytics.adjust_business_day(sat, self.closed("EUR", "GBP")), dt.date(2029, 3, 29))
        self.assertEqual(analytics.adjust_business_day(sat, self.closed("USD")), dt.date(2029, 3, 30))


class TestExposureAndPair(unittest.TestCase):
    def test_exposure_excludes_same_day_cashflow(self):
        dates, amounts = dates_and_amounts(annual_bond(coupon=10))
        # On the investment date: NAV including the -100 is 0, but 100 is exposed.
        self.assertAlmostEqual(analytics.nav_at(D0, dates, amounts, 0.10), Decimal(0), places=8)
        self.assertAlmostEqual(analytics.exposure_on(D0, dates, amounts, 0.10), Decimal(100), places=8)
        # On a coupon date: 110 including the coupon, 100 left after it.
        self.assertAlmostEqual(
            analytics.exposure_on(days(365), dates, amounts, 0.10), Decimal(100), places=8
        )

    def test_currency_pair_uses_market_convention(self):
        self.assertEqual(analytics.currency_pair("GBP", "EUR"), "EURGBP")
        self.assertEqual(analytics.currency_pair("USD", "EUR"), "EURUSD")
        self.assertEqual(analytics.currency_pair("USD", "GBP"), "GBPUSD")


class TestProposeHedges(unittest.TestCase):
    def setUp(self):
        self.flows = annual_bond("GBP", coupon=10, years=2) + annual_bond("EUR", coupon=5, years=2)
        self.irrs = analytics.currency_irrs(self.flows)

    def hedges(self, flows=None, irrs=None, **kwargs):
        kwargs.setdefault("tenor_months", 12)
        return analytics.propose_hedges(flows or self.flows, irrs or self.irrs, "EUR", **kwargs)

    def test_base_currency_is_not_hedged(self):
        self.assertEqual({t.notional_currency for t in self.hedges()}, {"GBP"})

    def test_sells_foreign_currency_sized_to_nav_after_cashflow(self):
        trades = self.hedges()
        self.assertEqual(len(trades), 2)
        for trade in trades:
            self.assertEqual(trade.currency_pair, "EURGBP")
            self.assertEqual(trade.direction, Direction.SELL)
            self.assertEqual(trade.notional_amount, Decimal("100.00"))  # par bond at its IRR

    def test_forwards_roll_from_one_delivery_to_the_next_trade(self):
        trades = self.hedges()
        self.assertEqual(trades[0].trade_date, D0)
        self.assertEqual(trades[0].delivery_date, trades[1].trade_date)
        self.assertEqual(trades[-1].delivery_date, days(730))

    def test_last_delivery_is_capped_at_final_cashflow(self):
        # 2028 is a leap year, so 3 x 365 days ends on 29 Sep, a day before
        # the 12-month roll would deliver.
        flows = annual_bond("GBP", coupon=10, years=3)
        trades = self.hedges(flows=flows, irrs=analytics.currency_irrs(flows))
        self.assertEqual(trades[-1].delivery_date, dt.date(2028, 9, 29))

    def test_hedge_ratio_scales_notional(self):
        for trade in self.hedges(hedge_ratio=0.5):
            self.assertEqual(trade.notional_amount, Decimal("50.00"))

    def test_zero_hedge_ratio_gives_no_trades(self):
        self.assertEqual(self.hedges(hedge_ratio=0), [])

    def test_negative_hedge_ratio_fails(self):
        with self.assertRaises(ValueError):
            self.hedges(hedge_ratio=-0.1)

    def test_net_liability_is_bought_forward(self):
        # Borrow 100 GBP, pay 10 a year, repay 100: the fund is short GBP.
        flows = [make_cashflow(date=cf.date, currency="GBP", amount=-cf.amount_local,
                               cashflow_type=cf.cashflow_type.value)
                 for cf in annual_bond("GBP", coupon=10, years=2)]
        trades = self.hedges(flows=flows, irrs=analytics.currency_irrs(flows))
        self.assertTrue(trades)
        self.assertTrue(all(t.direction == Direction.BUY for t in trades))

    def test_notional_rounded_to_cents(self):
        for trade in self.hedges(hedge_ratio=1 / 3):
            self.assertEqual(trade.notional_amount, trade.notional_amount.quantize(Decimal("0.01")))


class TestFxForwardModel(unittest.TestCase):
    def make(self, **overrides):
        fields = dict(
            currency_pair="EURGBP", trade_date=D0, delivery_date=days(92),
            direction="Sell", notional_currency="GBP", notional_amount=Decimal("1"),
        )
        fields.update(overrides)
        return FxForward(**fields)

    def test_valid_trade(self):
        self.assertEqual(self.make().direction, Direction.SELL)

    def test_delivery_must_be_after_trade(self):
        with self.assertRaises(ValueError):
            self.make(delivery_date=D0)

    def test_notional_must_be_positive(self):
        with self.assertRaises(ValueError):
            self.make(notional_amount=Decimal("0"))

    def test_notional_currency_must_be_in_pair(self):
        with self.assertRaises(ValueError):
            self.make(notional_currency="USD")


@unittest.skipUnless(DATA_FILE.exists(), "input file not present (data/ is git-ignored)")
class TestRealData(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cashflows, _ = excel_io.load_cashflows(DATA_FILE, "Fund I")
        cls.irrs = analytics.currency_irrs(cls.cashflows)

    def test_every_nav_schedule_passes_sanity_checks(self):
        for currency, schedule in analytics.nav_schedules(self.cashflows, self.irrs).items():
            self.assertTrue(analytics.check_schedule(schedule, self.cashflows), currency)

    def test_fund_irr_lies_between_currency_irrs(self):
        # FX rates in the file are constant, so the fund IRR is a blend.
        irr = analytics.fund_irr(self.cashflows)
        self.assertGreater(irr, min(self.irrs.values()))
        self.assertLess(irr, max(self.irrs.values()))

    def test_hedges_cover_foreign_currencies_quarterly(self):
        trades = analytics.propose_hedges(self.cashflows, self.irrs, "EUR")
        self.assertEqual(len(trades), 40)  # 20 quarters x GBP, USD
        for currency, invested in [("GBP", Decimal("100000000")), ("USD", Decimal("200000000"))]:
            legs = [t for t in trades if t.notional_currency == currency]
            self.assertEqual(len(legs), 20)
            self.assertTrue(all(t.direction == Direction.SELL for t in legs))
            # First hedge = the amount just invested (PV of the rest at the IRR).
            self.assertAlmostEqual(legs[0].notional_amount, invested, delta=Decimal("0.01"))
            for this, following in zip(legs, legs[1:]):
                self.assertEqual(this.delivery_date, following.trade_date)
            self.assertEqual(legs[-1].delivery_date, dt.date(2030, 9, 30))

    def test_hedge_dates_are_business_days_for_both_currencies(self):
        trades = analytics.propose_hedges(self.cashflows, self.irrs, "EUR")
        for trade in trades:
            closed = analytics.settlement_holidays({trade.notional_currency, "EUR"}, range(2025, 2032))
            self.assertTrue(analytics.is_business_day(trade.trade_date, closed), trade)
            self.assertTrue(analytics.is_business_day(trade.delivery_date, closed), trade)
        gbp_trade_dates = {t.trade_date for t in trades if t.notional_currency == "GBP"}
        self.assertIn(dt.date(2029, 3, 29), gbp_trade_dates)      # Sat 31 Mar, Good Friday 30th
        self.assertNotIn(dt.date(2028, 9, 30), gbp_trade_dates)   # Saturday

    def test_moved_trades_are_sized_on_the_quarter_end(self):
        # Traded Fri 29 Dec 2028, but sized after the 31 Dec coupon is paid.
        gbp = [cf for cf in self.cashflows if cf.local_currency == "GBP"]
        expected = analytics.exposure_on(
            dt.date(2028, 12, 31), [cf.date for cf in gbp], [cf.amount_local for cf in gbp], self.irrs["GBP"]
        ).quantize(Decimal("0.01"))
        trades = analytics.propose_hedges(self.cashflows, self.irrs, "EUR")
        trade = next(t for t in trades if t.notional_currency == "GBP" and t.trade_date == dt.date(2028, 12, 29))
        self.assertEqual(trade.notional_amount, expected)


if __name__ == "__main__":
    unittest.main()
