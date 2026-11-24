"""End-to-end tests: run the whole program and compare its outputs by md5.

CSV files are hashed byte for byte. A workbook's bytes change on every save
(it records the save time), so its hash is taken over the sheet names and cell
values instead. If an output changes on purpose, check it by eye, then copy
the new hashes from the failure message into the EXPECTED_* dicts.
"""

import datetime as dt
import hashlib
import shutil
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path

import openpyxl
import pandas as pd

import excel_io
import main

DATA_FILE = Path(__file__).resolve().parent.parent / "data" / "CST_Case_Study_data__1_.xlsx"

EXPECTED_SYNTHETIC_CSV = {
    "fund_i_irr.csv": "e98a29ad20b1721421ee927306baf02e",
    "fund_i_nav.csv": "bd6beeefe85c1dce074467914f16ef63",
    "fund_i_fx_hedges.csv": "08da7c885047d6e8c8061b829ca82492",
}
EXPECTED_SYNTHETIC_XLSX = {"fund_i_results.xlsx": "fe9c9670f0336ae7e92b2a503acd8c9e"}
EXPECTED_REAL_CSV = {
    "fund_i_irr.csv": "61cd453a52dcc1f703782e768f7f917d",
    "fund_i_nav.csv": "b06a27551419194d170b9fdb2a1a8ca9",
    "fund_i_fx_hedges.csv": "699550ffd417d6a4733768afac6fa63b",
}
EXPECTED_REAL_XLSX = {"fund_i_results.xlsx": "ecfe425d4f2f39c704b45c07720fdaa1"}


def md5(path: Path) -> str:
    """md5 of a CSV's bytes, or of a workbook's sheet names and cell values."""
    if path.suffix == ".xlsx":
        sheets = pd.read_excel(path, sheet_name=None)
        data = "".join(f"{name}\n{df.to_csv(index=False)}" for name, df in sheets.items()).encode()
    else:
        data = path.read_bytes()
    return hashlib.md5(data).hexdigest()


def output_hashes(output_dir: Path) -> dict[str, str]:
    """File name -> md5 for every file the program wrote."""
    return {path.name: md5(path) for path in sorted(output_dir.iterdir())}


def run_main(*args: str) -> tuple[int, str, str]:
    """Run the program as the command line would; return (exit code, stdout, stderr)."""
    stdout, stderr = StringIO(), StringIO()
    code = 0
    with redirect_stdout(stdout), redirect_stderr(stderr):
        try:
            main.main([*map(str, args)])
        except SystemExit as exc:
            code = exc.code
    return code, stdout.getvalue(), stderr.getvalue()


HEADERS = excel_io.REQUIRED_COLUMNS


def write_workbook(path: Path, rows: list[list], headers: list[str] = HEADERS) -> Path:
    """Write a test input workbook and return its path."""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(headers)
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path


def synthetic_rows() -> list[list]:
    """A small two-fund input with the same kinds of mess as the real one.

    Fund I: a GBP loan (100 at 10%) and a EUR loan (50 at 8%), quarterly for a
    year. Fund II: a USD loan to be filtered out. Row 3 has the GPB typo and
    row 4 a text date with junk around it.
    """
    quarter_ends = [dt.datetime(2025, 9, 30), dt.datetime(2025, 12, 31), dt.datetime(2026, 3, 31),
                    dt.datetime(2026, 6, 30), dt.datetime(2026, 9, 30)]
    rows = []

    def loan(fund, currency, fx, principal, coupon):
        rows.append([fund, quarter_ends[0], "Investment", currency, -principal, -principal * fx])
        for date in quarter_ends[1:]:
            rows.append([fund, date, "Interest", currency, coupon, coupon * fx])
        rows.append([fund, quarter_ends[-1], "Principal Repayment", currency, principal, principal * fx])

    loan("Fund I", "GBP", 1.15, 100, 2.5)
    loan("Fund I", "EUR", 1, 50, 1)
    loan("Fund II", "USD", 0.85, 200, 6)
    rows[1][3] = "GPB"                       # currency typo
    rows[2][1] = " 31/03/2026  00:00:00`"    # text date with junk
    return [[i, *row, "EUR"] for i, row in enumerate(rows, start=1)]


class EndToEndTestCase(unittest.TestCase):
    """Base class: gives each test a temporary folder, removed afterwards."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)

    def run_ok(self, input_path: Path, *extra: str) -> tuple[Path, str]:
        """Run the program, assert it succeeded, return (output folder, stdout)."""
        output_dir = self.tmp / "out"
        code, stdout, stderr = run_main(
            input_path, "--output-dir", output_dir, "--log-level", "ERROR", *extra
        )
        self.assertEqual(code, 0, stderr)
        return output_dir, stdout

    def assert_hashes(self, output_dir: Path, expected: dict[str, str]) -> None:
        """Assert the outputs match; on failure, show the new hashes."""
        actual = output_hashes(output_dir)
        self.assertEqual(actual, expected, f"\nNew hashes, if the change is intended:\n{actual}")


class TestSyntheticEndToEnd(EndToEndTestCase):
    """Runs everywhere: builds its own input file, so needs no data/ folder."""

    def setUp(self):
        super().setUp()
        self.input = write_workbook(self.tmp / "cashflows.xlsx", synthetic_rows())

    def test_csv_output_matches_known_good_hashes(self):
        output_dir, _ = self.run_ok(self.input, "--output-format", "csv")
        self.assert_hashes(output_dir, EXPECTED_SYNTHETIC_CSV)

    def test_excel_output_matches_known_good_hash(self):
        output_dir, _ = self.run_ok(self.input)
        self.assert_hashes(output_dir, EXPECTED_SYNTHETIC_XLSX)

    def test_summary_is_printed(self):
        _, stdout = self.run_ok(self.input, "--output-format", "csv")
        for heading in ["Fund I: IRR", "Fund I: NAV schedule", "Fund I: FX forward hedges"]:
            self.assertIn(heading, stdout)


@unittest.skipUnless(DATA_FILE.exists(), "input file not present")
class TestRealDataEndToEnd(EndToEndTestCase):
    def test_csv_output_matches_known_good_hashes(self):
        output_dir, _ = self.run_ok(DATA_FILE, "--output-format", "csv")
        self.assert_hashes(output_dir, EXPECTED_REAL_CSV)

    def test_excel_output_matches_known_good_hash(self):
        output_dir, _ = self.run_ok(DATA_FILE)
        self.assert_hashes(output_dir, EXPECTED_REAL_XLSX)


class TestInputFileErrors(EndToEndTestCase):
    """Bad input must stop with exit code 1 and a clear message, no traceback."""

    def assert_fails(self, input_path: Path, message: str) -> None:
        """Assert that a run fails cleanly with ``message`` and writes nothing."""
        code, _, stderr = run_main(input_path, "--output-dir", self.tmp / "out")
        self.assertEqual(code, 1)
        self.assertIn(message, stderr)
        self.assertNotIn("Traceback", stderr)
        self.assertFalse((self.tmp / "out").exists(), "nothing should be written")

    def test_missing_file(self):
        self.assert_fails(self.tmp / "missing.xlsx", "Input file not found")

    def test_csv_file(self):
        path = self.tmp / "cashflows.csv"
        path.write_text("ID,Fund Name\n1,Fund I\n")
        self.assert_fails(path, "must be an Excel workbook")

    def test_renamed_csv_with_xlsx_extension(self):
        path = self.tmp / "cashflows.xlsx"
        path.write_text("ID,Fund Name\n1,Fund I\n")
        self.assert_fails(path, "not a valid Excel workbook")

    def test_missing_column(self):
        headers = [h for h in HEADERS if h != "Cashflow Amount Base"]
        rows = [row[:6] + row[7:] for row in synthetic_rows()]
        self.assert_fails(write_workbook(self.tmp / "cashflows.xlsx", rows, headers),
                          "Missing column(s) ['Cashflow Amount Base']")

if __name__ == "__main__":
    unittest.main()
