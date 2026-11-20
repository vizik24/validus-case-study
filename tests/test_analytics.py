import datetime as dt
import unittest
from decimal import Decimal
from pathlib import Path

import analytics
import excel_io
from models import Cashflow, NavPoint

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


if __name__ == "__main__":
    unittest.main()
