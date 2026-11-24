"""Unit tests for the Cashflow model and the cleaning step in excel_io."""

import datetime as dt
import unittest
from decimal import Decimal

import excel_io
from models import Cashflow


def make_row(**overrides) -> dict:
    """One valid raw row, keyed by Excel header, with ``overrides`` applied."""
    row = {
        "ID": 1,
        "Fund Name": "Fund I",
        "Date": dt.datetime(2025, 9, 30),
        "Cashflow Type": "Investment",
        "Local Currency": "GBP",
        "Cashflow Amount Local": -100,
        "Cashflow Amount Base": -114.57,
        "Base Currency": "EUR",
    }
    row.update(overrides)
    return row


class TestCashflowModel(unittest.TestCase):
    def test_principal_repayment_is_allowed(self):
        Cashflow.model_validate(make_row(**{"Cashflow Type": "Principal Repayment"}))

    def test_unknown_currency_fails(self):
        _, errors = excel_io.validate_rows([make_row(**{"Local Currency": "GPB"})])
        self.assertEqual(errors[0][0]["loc"], ("Local Currency",))

    def test_amounts_are_exact_decimals(self):
        cf = Cashflow.model_validate(make_row(**{"Cashflow Amount Base": 2864344.6379468376}))
        self.assertIsInstance(cf.amount_local, Decimal)
        self.assertEqual(cf.amount_base, Decimal("2864344.6379468376"))

    def test_missing_amount_fails(self):
        _, errors = excel_io.validate_rows([make_row(**{"Cashflow Amount Local": None})])
        self.assertIn(0, errors)


class TestCleaners(unittest.TestCase):
    def test_clean_date_handles_source_junk(self):
        self.assertEqual(excel_io.clean_date(" 30/09/2030  00:00:00"), dt.date(2030, 9, 30))
        self.assertEqual(excel_io.clean_date("30/09/2027  00:00:00`"), dt.date(2027, 9, 30))

    def test_clean_date_is_day_first(self):
        self.assertEqual(excel_io.clean_date("01/02/2027"), dt.date(2027, 2, 1))

    def test_clean_currency(self):
        self.assertEqual(excel_io.clean_currency(" gpb "), "GBP")
        self.assertEqual(excel_io.clean_currency("XYZ"), "XYZ")


class TestCleanAndRevalidate(unittest.TestCase):
    def test_only_failing_fields_are_cleaned(self):
        rows = [make_row(), make_row(ID=2, Date=" 30/09/2030  00:00:00", **{"Local Currency": "GPB"})]
        _, errors = excel_io.validate_rows(rows)
        self.assertEqual(list(errors), [1])

        with self.assertLogs("excel_io", "INFO") as logs:
            cleaned = excel_io.clean_rows(rows, errors)
        self.assertEqual(sum("Corrected" in line for line in logs.output), 2)  # audit trail
        self.assertEqual(rows[1]["Local Currency"], "GPB")  # input not mutated

        cashflows, errors = excel_io.validate_rows(cleaned)
        self.assertEqual(errors, {})
        self.assertEqual(cashflows[1].local_currency, "GBP")

    def test_unfixable_row_still_fails(self):
        rows = [make_row(Date="garbage")]
        _, errors = excel_io.validate_rows(rows)
        cleaned = excel_io.clean_rows(rows, errors)
        _, errors = excel_io.validate_rows(cleaned)
        self.assertEqual(cleaned, rows)  # nothing could be fixed
        self.assertIn(0, errors)


if __name__ == "__main__":
    unittest.main()
