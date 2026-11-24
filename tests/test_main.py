"""Unit tests for the command-line options."""

import unittest
from contextlib import redirect_stderr
from io import StringIO

import main


class TestParseArgs(unittest.TestCase):
    def parse(self, *extra):
        """Parse a command line with a dummy input file plus ``extra`` options."""
        return main.parse_args(["input.xlsx", *extra])

    def test_defaults_match_the_brief(self):
        args = self.parse()
        self.assertEqual(args.fund, "Fund I")
        self.assertEqual(args.hedge_ratio, 1.0)
        self.assertEqual(args.tenor_months, 3)
        self.assertEqual(args.output_format, "excel")

    def test_invalid_hedge_options_are_rejected(self):
        for extra in [("--hedge-ratio", "-0.1"), ("--hedge-ratio", "1.5"), ("--hedge-ratio", "half"),
                      ("--tenor-months", "0"), ("--tenor-months", "2.5")]:
            with self.subTest(extra=extra), redirect_stderr(StringIO()), self.assertRaises(SystemExit):
                self.parse(*extra)


if __name__ == "__main__":
    unittest.main()
