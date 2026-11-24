"""Input and output: reading and validating the workbook, printing and saving results."""

import datetime as dt
import logging
import re
import zipfile
from decimal import Decimal
from pathlib import Path

import pandas as pd
from pydantic import ValidationError

from models import Cashflow, Results

logger = logging.getLogger(__name__)


# --- Input -------------------------------------------------------------------

class InputFileError(Exception):
    """Raised when the input path is not a readable Excel workbook."""


class DataValidationError(Exception):
    """Raised when one or more input rows cannot be turned into a Cashflow."""


EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
# .xlsx/.xlsm files are zip archives; every zip file starts with these bytes.
ZIP_SIGNATURE = b"PK\x03\x04"
REQUIRED_COLUMNS = [field.alias for field in Cashflow.model_fields.values()]


def check_excel_file(path: Path) -> None:
    """Fail early, with a clear message, if ``path`` is not an Excel workbook.

    :param path: The input file.
    :raises InputFileError: If it is missing, not a file, an old ``.xls``,
        not ``.xlsx``/``.xlsm``, or not a zip archive (e.g. a renamed CSV).
    """
    if not path.exists():
        raise InputFileError(f"Input file not found: {path}")
    if not path.is_file():
        raise InputFileError(f"Input path is not a file: {path}")
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise InputFileError(
            f"{path.name} is an old-style .xls workbook, which is not supported. "
            "Open it in Excel and save it as .xlsx."
        )
    if suffix not in EXCEL_SUFFIXES:
        raise InputFileError(
            f"Input must be an Excel workbook (.xlsx or .xlsm), got {path.name!r}"
        )
    with path.open("rb") as f:
        if f.read(len(ZIP_SIGNATURE)) != ZIP_SIGNATURE:
            raise InputFileError(
                f"{path.name} has an Excel extension but is not a valid Excel workbook "
                "(it may be a renamed CSV, password-protected or corrupted)"
            )


def read_excel_rows(path: Path) -> list[dict]:
    """Read the first sheet as raw rows, leaving all value checks to Pydantic.

    :param path: The input workbook.
    :returns: One dict per row, keyed by column header; empty cells are ``None``.
    :raises InputFileError: If the file is not a readable Excel workbook.
    :raises DataValidationError: If a required column is missing or there
        are no data rows.
    """
    check_excel_file(path)
    logger.info("Reading %s", path)
    # dtype=object stops pandas guessing column types.
    try:
        df = pd.read_excel(path, sheet_name=0, dtype=object)
    except (zipfile.BadZipFile, KeyError, ValueError, OSError) as exc:
        raise InputFileError(f"Could not read {path.name} as an Excel workbook: {exc}") from exc
    logger.info("Read %d rows x %d columns", *df.shape)

    # Checked once here rather than as one Pydantic error per row.
    missing_columns = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_columns:
        raise DataValidationError(
            f"Missing column(s) {missing_columns} in {path.name}. "
            f"Found: {[str(c) for c in df.columns]}"
        )
    if df.empty:
        raise DataValidationError(f"{path.name} has a header row but no data rows")

    # Empty cells come through as NaN, which Pydantic would accept as a
    # number. Turn them into None so missing values fail validation.
    df = df.where(df.notna(), None)
    return df.to_dict(orient="records")


RowErrors = dict[int, list[dict]]  # row index -> Pydantic error dicts


def validate_rows(rows: list[dict]) -> tuple[list[Cashflow], RowErrors]:
    """Try to turn every raw row into a Cashflow, without raising.

    :param rows: Raw rows from :func:`read_excel_rows`.
    :returns: ``(cashflows, errors)``: the rows that passed, and Pydantic's
        errors for the rest, keyed by row index.
    """
    cashflows: list[Cashflow] = []
    errors: RowErrors = {}
    for i, row in enumerate(rows):
        try:
            cashflows.append(Cashflow.model_validate(row))
        except ValidationError as exc:
            errors[i] = exc.errors()
    logger.info(
        "Validated %d rows: %d valid, %d invalid (%d error(s))",
        len(rows), len(cashflows), len(errors), sum(len(e) for e in errors.values()),
    )
    return cashflows, errors


def format_errors(rows: list[dict], errors: RowErrors) -> list[str]:
    """One readable line per error, with its Excel row number and ID.

    :param rows: The rows that were validated.
    :param errors: Errors from :func:`validate_rows`.
    :returns: The messages.
    """
    lines = []
    for i, row_errors in errors.items():
        for err in row_errors:
            field = ".".join(str(part) for part in err["loc"])
            lines.append(
                f"Excel row {i + 2} (ID {rows[i].get('ID')}): "  # row 1 is the header
                f"{field}: {err['msg']} (got {err['input']!r})"
            )
    return lines


def check_duplicate_ids(cashflows: list[Cashflow]) -> list[str]:
    """Find repeated IDs; this needs every row, so it can't live in the model.

    :param cashflows: Validated cashflows.
    :returns: One message per repeated ID.
    """
    seen: set[int] = set()
    problems = []
    for cf in cashflows:
        if cf.id in seen:
            problems.append(f"Duplicate ID {cf.id}")
        seen.add(cf.id)
    return problems


# --- Cleaning ----------------------------------------------------------------
# Each cleaner takes the raw cell value and returns a fixed value, or the value
# unchanged if it doesn't know how to fix it. The second validation pass then
# reports anything that is still wrong, so a cleaner never hides a problem.

DATE_FORMATS = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y")  # day-first, as in the source file
CURRENCY_TYPOS = {"GPB": "GBP"}


def clean_date(value):
    """Parse a day-first text date with stray spaces, quotes or a backtick.

    :param value: Raw cell value, e.g. ``" 30/09/2027  00:00:00`"``.
    :returns: The date, or ``value`` unchanged if it can't be parsed.
    """
    if not isinstance(value, str):
        return value
    text = " ".join(value.strip().strip("`'").split())  # collapse runs of spaces
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return value


def clean_currency(value):
    """Strip and upper-case a currency code, then fix known typos.

    :param value: Raw cell value, e.g. ``" gpb "``.
    :returns: The cleaned code, or ``value`` unchanged if it isn't a string.
    """
    if not isinstance(value, str):
        return value
    code = value.strip().upper()
    return CURRENCY_TYPOS.get(code, code)


# Excel header -> cleaner. Only fields listed here can be fixed automatically.
CLEANERS = {
    "Date": clean_date,
    "Local Currency": clean_currency,
    "Base Currency": clean_currency,
}


def clean_rows(rows: list[dict], errors: RowErrors) -> list[dict]:
    """Apply the matching cleaner to each field that failed validation.

    Every change is logged at INFO with the row's ID, as an audit trail.

    :param rows: The validated rows; not modified.
    :param errors: Errors from :func:`validate_rows`.
    :returns: A copy of ``rows`` with the fixable fields repaired.
    """
    cleaned = [dict(row) for row in rows]
    corrections = 0
    for i, row_errors in errors.items():
        for field in sorted({err["loc"][0] for err in row_errors}):
            cleaner = CLEANERS.get(field)
            if cleaner is None:
                logger.debug("ID %s: no cleaner for field %r", rows[i].get("ID"), field)
                continue
            old = cleaned[i][field]
            new = cleaner(old)
            if new != old:
                cleaned[i][field] = new
                corrections += 1
                logger.info("Corrected ID %s: %s %r -> %s", rows[i].get("ID"), field, old, _show(new))
            else:
                logger.debug("ID %s: %s %r could not be cleaned", rows[i].get("ID"), field, old)
    logger.info("Applied %d correction(s)", corrections)
    return cleaned


def _show(value) -> str:
    """Format a value for a log line: ISO for dates, ``repr`` otherwise."""
    return value.isoformat() if isinstance(value, dt.date) else repr(value)


def filter_fund(cashflows: list[Cashflow], fund_name: str) -> list[Cashflow]:
    """Keep only one fund's cashflows.

    :param cashflows: Cashflows for every fund in the file.
    :param fund_name: Exact fund name, e.g. ``"Fund I"``.
    :returns: The matching cashflows.
    :raises ValueError: If there are none; the message lists the funds found.
    """
    matches = [cf for cf in cashflows if cf.fund_name == fund_name]
    if not matches:
        available = sorted({cf.fund_name for cf in cashflows})
        raise ValueError(f"No cashflows found for {fund_name!r}. Available funds: {available}")
    logger.info("Selected %d of %d cashflows for %r", len(matches), len(cashflows), fund_name)
    return matches


def load_cashflows(path: Path, fund_name: str) -> list[Cashflow]:
    """Read, validate, clean and revalidate the whole file, then filter to one fund.

    :param path: The input workbook.
    :param fund_name: Fund to keep.
    :returns: The fund's validated cashflows.
    :raises InputFileError: If the file is not a readable Excel workbook.
    :raises DataValidationError: If any row is still invalid after cleaning
        or an ID is repeated. Rows are never dropped silently, since a
        missing cashflow would change the IRR.
    :raises ValueError: If the fund is not in the file.
    """
    rows = read_excel_rows(path)
    cashflows, errors = validate_rows(rows)

    if errors:
        for line in format_errors(rows, errors):
            logger.debug("Validation error: %s", line)
        logger.info("Cleaning the fields that failed, then validating again")
        rows = clean_rows(rows, errors)
        cashflows, errors = validate_rows(rows)

    # All remaining problems: rows still invalid after cleaning, plus duplicate IDs.
    problems = format_errors(rows, errors) + check_duplicate_ids(cashflows)
    if problems:
        raise DataValidationError(
            f"{len(problems)} problem(s) in the input data:\n" + "\n".join(problems)
        )
    return filter_fund(cashflows, fund_name)


# --- Output ------------------------------------------------------------------
# The console, the Excel workbook and the CSV files all show the same three
# tables, built once as DataFrames. Values are rounded here, for output only:
# IRRs to 10 decimal places, money to 2.

OUTPUT_FORMATS = ("excel", "csv")
IRR_DECIMALS = 10
MONEY_STEP = Decimal("0.01")


def round_money(value: Decimal) -> Decimal:
    """Round to 2 decimal places, showing a tiny negative as 0.00, not -0.00."""
    rounded = value.quantize(MONEY_STEP)
    return abs(rounded) if rounded == 0 else rounded


def irr_table(results: Results) -> pd.DataFrame:
    """One row for the fund (base amounts), then one per currency (local amounts)."""
    rows = [("Fund", results.base_currency, results.fund_irr)]
    rows += [("Currency", c, irr) for c, irr in sorted(results.currency_irrs.items())]
    return pd.DataFrame(
        [(level, currency, round(irr, IRR_DECIMALS)) for level, currency, irr in rows],
        columns=["Level", "Currency", "IRR"],
    )


def nav_table(results: Results) -> pd.DataFrame:
    """One row per currency and date, with the NAV in local currency."""
    return pd.DataFrame(
        [
            (currency, point.date, round_money(point.nav))
            for currency, schedule in sorted(results.nav_schedules.items())
            for point in schedule
        ],
        columns=["Currency", "Date", "NAV"],
    )


def hedge_table(results: Results) -> pd.DataFrame:
    """One row per proposed forward."""
    return pd.DataFrame(
        [
            (t.currency_pair, t.trade_date, t.delivery_date,
             t.direction.value, t.notional_currency, t.notional_amount)
            for t in results.hedges
        ],
        columns=["Currency Pair", "Trade Date", "Delivery Date",
                 "Direction", "Notional Currency", "Notional Amount"],
    )


# Sheet / file name -> table builder. The order is the order of the sheets.
TABLES = {"IRR": irr_table, "NAV": nav_table, "FX Hedges": hedge_table}


def print_summary(results: Results) -> None:
    """Print the three tables to stdout; logs go to stderr, so they never mix.

    :param results: The run's results.
    """
    money = "{:,.2f}".format
    irr = irr_table(results)
    irr["IRR"] = irr["IRR"].map("{:.4%}".format)
    # Wide on screen: one row per date, one column per currency.
    nav = nav_table(results).pivot(index="Date", columns="Currency", values="NAV").map(money)
    hedges = hedge_table(results)
    hedges["Notional Amount"] = hedges["Notional Amount"].map(money)

    for title, table, index in [
        ("IRR", irr, False),
        ("NAV schedule (local currency)", nav, True),
        (f"FX forward hedges ({len(hedges)} trades)", hedges, False),
    ]:
        print(f"\n=== {results.fund_name}: {title} ===\n")
        print(table.to_string(index=index))
    print()


def output_stem(name: str) -> str:
    """Make a file-name prefix, e.g. ``"Fund I"`` -> ``"fund_i"``."""
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") or "fund"


def write_outputs(results: Results, output_dir: Path, output_format: str = "excel") -> list[Path]:
    """Save the three tables, replacing any existing files of the same name.

    ``"excel"`` writes ``<fund>_results.xlsx`` with sheets IRR, NAV and FX
    Hedges; ``"csv"`` writes ``<fund>_irr.csv``, ``<fund>_nav.csv`` and
    ``<fund>_fx_hedges.csv``.

    :param results: The run's results.
    :param output_dir: Folder to write to; created if missing.
    :param output_format: ``"excel"`` or ``"csv"``.
    :returns: The paths written.
    :raises ValueError: If ``output_format`` is not recognised.
    :raises OSError: If a file can't be written, e.g. it is open in Excel.
    """
    if output_format not in OUTPUT_FORMATS:
        raise ValueError(f"Output format must be one of {OUTPUT_FORMATS}, got {output_format!r}")
    output_dir.mkdir(parents=True, exist_ok=True)
    tables = {name: build(results) for name, build in TABLES.items()}
    stem = output_stem(results.fund_name)

    if output_format == "excel":
        path = output_dir / f"{stem}_results.xlsx"
        with pd.ExcelWriter(path) as writer:
            for name, table in tables.items():
                table.to_excel(writer, sheet_name=name, index=False)
        paths = [path]
    else:
        paths = []
        for name, table in tables.items():
            path = output_dir / f"{stem}_{output_stem(name)}.csv"
            # Fixed "\n" line endings, so the bytes are the same on any OS.
            table.to_csv(path, index=False, lineterminator="\n")
            paths.append(path)
    for path in paths:
        logger.info("Wrote %s", path)
    return paths
