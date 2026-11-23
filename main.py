"""Entry point: parses command-line arguments and runs each step in order."""

import argparse
import logging
import sys
from pathlib import Path

import analytics
import excel_io

logger = logging.getLogger(__name__)

LOG_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


def hedge_ratio(text: str) -> float:
    """argparse type: a fraction of NAV between 0 and 1."""
    try:
        value = float(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"not a number: {text!r}")
    if not 0 <= value <= 1:
        raise argparse.ArgumentTypeError(f"must be between 0 and 1 (e.g. 0.5 for 50%), got {value}")
    return value


def positive_int(text: str) -> int:
    """argparse type: a whole number of at least 1."""
    try:
        value = int(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"not a whole number: {text!r}")
    if value < 1:
        raise argparse.ArgumentTypeError(f"must be at least 1, got {value}")
    return value


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fund IRR, NAV and FX hedge calculator")
    parser.add_argument("input", type=Path, help="Path to the cashflow Excel file")
    parser.add_argument("--fund", default="Fund I", help="Fund to analyse")
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    parser.add_argument(
        "--hedge-ratio",
        type=hedge_ratio,
        default=1.0,
        help="Share of NAV exposure to hedge, 0 to 1 (default: 1.0 = 100%%)",
    )
    parser.add_argument(
        "--tenor-months",
        type=positive_int,
        default=3,
        help="Length of each rolling FX forward in months (default: 3)",
    )
    parser.add_argument(
        "--log-level",
        type=str.upper,
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="How much detail to log (default: INFO; DEBUG shows every row-level step)",
    )
    return parser.parse_args(argv)


def configure_logging(level: str) -> None:
    """Send log lines to stderr, so they never mix with printed results."""
    logging.basicConfig(level=level, format=LOG_FORMAT, datefmt="%H:%M:%S", stream=sys.stderr)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    configure_logging(args.log_level)
    logger.info(
        "Run started: input=%s fund=%r hedge_ratio=%s tenor_months=%s output_dir=%s",
        args.input, args.fund, args.hedge_ratio, args.tenor_months, args.output_dir,
    )

    try:
        cashflows, _corrections = excel_io.load_cashflows(args.input, args.fund)
        base_currency = analytics.base_currency(cashflows)
    except (excel_io.InputFileError, excel_io.DataValidationError, ValueError) as exc:
        logger.error("%s", exc)
        sys.exit(1)

    fund_irr = analytics.fund_irr(cashflows)
    currency_irrs = analytics.currency_irrs(cashflows)

    schedules = analytics.nav_schedules(cashflows, currency_irrs)

    trades = analytics.propose_hedges(
        cashflows, currency_irrs, base_currency, args.hedge_ratio, args.tenor_months
    )

    excel_io.print_summary(fund_irr, currency_irrs, schedules, trades)
    excel_io.write_outputs(args.output_dir, fund_irr, currency_irrs, schedules, trades)
    logger.info("Run finished")


if __name__ == "__main__":
    main()
