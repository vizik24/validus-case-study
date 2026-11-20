import unittest
from contextlib import redirect_stderr
from io import StringIO

import main


class TestParseArgs(unittest.TestCase):
    def parse(self, *extra):
        return main.parse_args(["input.xlsx", *extra])

    def assert_rejected(self, *extra):
        with redirect_stderr(StringIO()), self.assertRaises(SystemExit):
            self.parse(*extra)

    def test_defaults_match_the_brief(self):
        args = self.parse()
        self.assertEqual(args.fund, "Fund I")
        self.assertEqual(args.hedge_ratio, 1.0)
        self.assertEqual(args.tenor_months, 3)
        self.assertEqual(args.log_level, "INFO")

    def test_hedge_strategy_can_be_changed(self):
        args = self.parse("--hedge-ratio", "0.5", "--tenor-months", "6")
        self.assertEqual(args.hedge_ratio, 0.5)
        self.assertEqual(args.tenor_months, 6)

    def test_hedge_ratio_bounds(self):
        self.assertEqual(self.parse("--hedge-ratio", "0").hedge_ratio, 0.0)
        self.assert_rejected("--hedge-ratio", "-0.1")
        self.assert_rejected("--hedge-ratio", "1.5")
        self.assert_rejected("--hedge-ratio", "half")

    def test_tenor_must_be_positive_whole_months(self):
        self.assert_rejected("--tenor-months", "0")
        self.assert_rejected("--tenor-months", "2.5")

    def test_log_level_is_case_insensitive(self):
        self.assertEqual(self.parse("--log-level", "debug").log_level, "DEBUG")


if __name__ == "__main__":
    unittest.main()
