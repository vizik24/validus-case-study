"""Command-line entry point.

Parses the arguments, then runs each step in order: load the cashflows,
calculate the IRRs, NAV schedules and hedges, print a summary and save the
results. Run ``python main.py --help`` for the options.
"""

import argparse
import logging
import sys
from pathlib import Path

import analytics
import excel_io
from models import Results

logger = logging.getLogger(__name__)

LOG_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


def hedge_ratio(text: str) -> float:
    """Parse ``--hedge-ratio``: a fraction of NAV between 0 and 1.

    :param text: The argument as typed, e.g. ``"0.5"``.
    :returns: The ratio as a float.
    :raises argparse.ArgumentTypeError: If it is not a number or is outside
        0 to 1. argparse turns this into a usage error.
    """
    try:
        value = float(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"not a number: {text!r}")
    if not 0 <= value <= 1:
        raise argparse.ArgumentTypeError(f"must be between 0 and 1 (e.g. 0.5 for 50%), got {value}")
    return value


def positive_int(text: str) -> int:
    """Parse ``--tenor-months``: a whole number of at least 1.

    :param text: The argument as typed, e.g. ``"3"``.
    :returns: The number as an int.
    :raises argparse.ArgumentTypeError: If it is not a whole number or is
        less than 1. argparse turns this into a usage error.
    """
    try:
        value = int(text)
    except ValueError:
        raise argparse.ArgumentTypeError(f"not a whole number: {text!r}")
    if value < 1:
        raise argparse.ArgumentTypeError(f"must be at least 1, got {value}")
    return value


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse the command-line arguments.

    The defaults reproduce the brief: Fund I, a 100% hedge ratio and
    3-month forwards, saved as an Excel workbook in ``output/``.

    :param argv: Arguments to parse, without the program name. ``None``
        means ``sys.argv[1:]``.
    :returns: The parsed options.
    """
    parser = argparse.ArgumentParser(description="Fund IRR, NAV and FX hedge calculator")
    parser.add_argument("input", type=Path, help="Path to the cashflow Excel file (.xlsx or .xlsm)")
    parser.add_argument("--fund", default="Fund I", help="Fund to analyse (default: Fund I)")
    parser.add_argument(
        "--output-dir", type=Path, default=Path("output"),
        help="Folder for the result files (default: output)",
    )
    parser.add_argument(
        "--output-format",
        type=str.lower,
        choices=excel_io.OUTPUT_FORMATS,
        default="excel",
        help="excel: one workbook with IRR, NAV and FX Hedges sheets (default); "
             "csv: one CSV file per table",
    )
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
    """Send log lines to stderr, so they never mix with printed results.

    :param level: Level name, e.g. ``"INFO"`` or ``"DEBUG"``.
    """
    # force=True replaces any earlier set-up, so each call to main() (e.g. in
    # tests) uses its own level and the current stderr.
    logging.basicConfig(
        level=level, format=LOG_FORMAT, datefmt="%H:%M:%S", stream=sys.stderr, force=True
    )


def main(argv: list[str] | None = None) -> None:
    """Run the whole program.

    Errors the user can fix (a bad input file, invalid data, an unknown
    fund, an output file that can't be written) are logged as one clear
    message and end the run with exit code 1, without a traceback.

    :param argv: Arguments, as for :func:`parse_args`.
    :raises SystemExit: With code 1 on any of the errors above, or code 2
        from argparse on a bad option.
    """
    args = parse_args(argv)
    configure_logging(args.log_level)
    logger.info(
        "Run started: input=%s fund=%r hedge_ratio=%s tenor_months=%s output=%s (%s)",
        args.input, args.fund, args.hedge_ratio, args.tenor_months,
        args.output_dir, args.output_format,
    )

    try:
        cashflows = excel_io.load_cashflows(args.input, args.fund)
        base_currency = analytics.base_currency(cashflows)
    except (excel_io.InputFileError, excel_io.DataValidationError, ValueError) as exc:
        logger.error("%s", exc)
        sys.exit(1)

    currency_irrs = analytics.currency_irrs(cashflows)
    results = Results(
        fund_name=args.fund,
        base_currency=base_currency,
        fund_irr=analytics.fund_irr(cashflows),
        currency_irrs=currency_irrs,
        nav_schedules=analytics.nav_schedules(cashflows, currency_irrs),
        hedges=analytics.propose_hedges(
            cashflows, currency_irrs, base_currency, args.hedge_ratio, args.tenor_months
        ),
    )

    excel_io.print_summary(results)
    try:
        excel_io.write_outputs(results, args.output_dir, args.output_format)
    except OSError as exc:  # e.g. the workbook is open in Excel, or no write permission
        logger.error("Could not write results to %s: %s", args.output_dir, exc)
        sys.exit(1)
    logger.info("Run finished")


if __name__ == "__main__":
    main()
