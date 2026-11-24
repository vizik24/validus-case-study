"""Input and output: everything that touches files or the console.

Input
    :func:`load_cashflows` reads the Excel sheet, validates each row as a
    :class:`~models.Cashflow`, repairs the fields that fail, validates again
    and keeps one fund's rows.

Output
    :func:`print_summary` prints the results and :func:`write_outputs` saves
    them, as one Excel workbook or three CSV files. Both show the same three
    tables, built by :func:`irr_table`, :func:`nav_table` and
    :func:`hedge_table`.
"""

import csv
import datetime as dt
import logging
import re
import zipfile
from collections import Counter
from decimal import Decimal
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
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

    Checks the extension and the file's first bytes, so a CSV renamed to
    ``.xlsx`` is caught here rather than as a confusing error from deep in
    pandas.

    :param path: The input file given on the command line.
    :raises InputFileError: If the file is missing, is a directory, is an old
        ``.xls`` workbook, has a non-Excel extension, or does not start with
        the zip signature that every ``.xlsx`` file has.
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
    """Check the file, then read its first sheet as raw rows.

    Nothing is converted or checked beyond the column headers: every cell is
    read as-is (``dtype=object``) and empty cells become ``None``, so all
    value checks are left to Pydantic.

    :param path: The input workbook.
    :returns: One dict per data row, keyed by column header.
    :raises InputFileError: If the file is not a readable Excel workbook.
    :raises DataValidationError: If a required column is missing or the
        sheet has no data rows.
    """
    check_excel_file(path)
    logger.info("Reading %s", path)
    # dtype=object stops pandas guessing column types, so all checking is
    # left to Pydantic.
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
    logger.debug("Columns: %s", list(df.columns))
    missing = df.isna().sum()
    for column, count in missing[missing > 0].items():
        logger.debug("Column %r has %d empty cell(s)", column, count)

    # Empty cells come through as NaN, which Pydantic would accept as a
    # number. Turn them into None so missing values fail validation.
    df = df.where(df.notna(), None)

    return df.to_dict(orient="records")


RowErrors = dict[int, list[dict]]  # row index -> Pydantic error dicts


def validate_rows(rows: list[dict]) -> tuple[list[Cashflow], RowErrors]:
    """Try to turn every raw row into a :class:`~models.Cashflow`.

    Nothing is raised here, so the caller can decide whether to clean and
    retry or to stop.

    :param rows: Raw rows, as returned by :func:`read_excel_rows`.
    :returns: A pair ``(cashflows, errors)``: the rows that passed, and
        Pydantic's error dicts for the rows that failed, keyed by their
        index in ``rows``.
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
    """Turn validation errors into readable messages.

    :param rows: The rows that were validated.
    :param errors: Errors from :func:`validate_rows`, keyed by row index.
    :returns: One line per error, giving the Excel row number (header is
        row 1), the row's ID, the field, the reason and the bad value.
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
    """Find IDs used by more than one cashflow.

    This check needs every row at once, so it can't live in the model,
    which sees one row at a time.

    :param cashflows: Validated cashflows.
    :returns: One message per repeated ID; empty if all IDs are unique.
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

    Example: ``" 30/09/2027  00:00:00`"`` becomes ``date(2027, 9, 30)``.

    :param value: The raw cell value.
    :returns: The parsed date, or ``value`` unchanged if it is not a string
        or matches none of :data:`DATE_FORMATS`.
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
    """Normalise case and whitespace, then fix known typos.

    Example: ``" gpb "`` becomes ``"GBP"``. An unknown code such as
    ``"XYZ"`` is returned upper-cased, and validation then rejects it.

    :param value: The raw cell value.
    :returns: The cleaned code, or ``value`` unchanged if it is not a string.
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

    Only failing fields are touched, and only those with an entry in
    :data:`CLEANERS`. Every value that changes is logged at INFO with the
    row's ID, which is the audit trail of corrections. This function does
    not decide whether a row is now valid; the next call to
    :func:`validate_rows` does.

    :param rows: The raw rows that were validated. Not modified.
    :param errors: Errors from :func:`validate_rows`, keyed by row index.
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
    """Format a value for a log line.

    :param value: Any value.
    :returns: ISO format for dates, ``repr`` for everything else.
    """
    return value.isoformat() if isinstance(value, dt.date) else repr(value)


def filter_fund(cashflows: list[Cashflow], fund_name: str) -> list[Cashflow]:
    """Keep only the cashflows belonging to one fund.

    :param cashflows: Validated cashflows for every fund in the file.
    :param fund_name: Exact fund name, e.g. ``"Fund I"``.
    :returns: The matching cashflows, in file order.
    :raises ValueError: If no cashflow has that fund name. The message
        lists the funds that are in the file.
    """
    matches = [cf for cf in cashflows if cf.fund_name == fund_name]
    if not matches:
        available = sorted({cf.fund_name for cf in cashflows})
        raise ValueError(f"No cashflows found for {fund_name!r}. Available funds: {available}")
    logger.info("Selected %d of %d cashflows for %r", len(matches), len(cashflows), fund_name)
    _log_fund_profile(matches)
    return matches


def _log_fund_profile(cashflows: list[Cashflow]) -> None:
    """Log an overview of the selected fund at DEBUG level.

    Logs the date range, a count of cashflows by currency and type, and the
    net local cashflow per currency. Does nothing unless DEBUG is enabled.

    :param cashflows: One fund's cashflows; must not be empty.
    """
    if not logger.isEnabledFor(logging.DEBUG):
        return
    dates = [cf.date for cf in cashflows]
    logger.debug("Date range: %s to %s", min(dates), max(dates))
    by_type = Counter((cf.local_currency, cf.cashflow_type.value) for cf in cashflows)
    for (currency, cashflow_type), count in sorted(by_type.items()):
        logger.debug("  %s %-20s x%d", currency, cashflow_type, count)
    for currency in sorted({cf.local_currency for cf in cashflows}):
        total = sum(cf.amount_local for cf in cashflows if cf.local_currency == currency)
        logger.debug("  %s net local cashflow: %s", currency, f"{total:,.2f}")


def load_cashflows(path: Path, fund_name: str) -> list[Cashflow]:
    """Load one fund's cashflows from the input workbook.

    The steps are: read, validate, clean the failing fields and validate
    again (only if anything failed), check for duplicate IDs, then filter to
    the fund. The whole file is validated, not just the chosen fund, so a
    bad row anywhere is reported.

    :param path: The input workbook.
    :param fund_name: Fund to keep, e.g. ``"Fund I"``.
    :returns: The fund's validated cashflows.
    :raises InputFileError: If the file is not a readable Excel workbook.
    :raises DataValidationError: If columns are missing, or any row is still
        invalid after cleaning, or an ID is repeated. Rows are never dropped
        silently, because a missing cashflow would change the IRR without
        anyone noticing.
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
# tables, built once by the *_table functions below. Values are rounded here,
# for output only: IRRs to 10 decimal places, money to 2.

OUTPUT_FORMATS = ("excel", "csv")
IRR_DECIMALS = 10
MONEY_STEP = Decimal("0.01")

Table = tuple[list[str], list[list]]  # (headers, rows)


def round_money(value: Decimal) -> Decimal:
    """Round a money amount to 2 decimal places for output.

    A tiny negative, such as a NAV(0) of ``-0.0001``, would round to
    ``-0.00``; it is returned as ``0.00`` instead.

    :param value: The exact amount.
    :returns: The rounded amount.
    """
    rounded = value.quantize(MONEY_STEP)
    return abs(rounded) if rounded == 0 else rounded


def irr_table(results: Results) -> Table:
    """Build the IRR table.

    :param results: The run's results.
    :returns: Headers ``Level, Currency, IRR`` and one row for the fund (on
        base-currency amounts) followed by one row per currency (on local
        amounts). IRRs are fractions rounded to :data:`IRR_DECIMALS` places.
    """
    headers = ["Level", "Currency", "IRR"]
    rows = [["Fund", results.base_currency, round(results.fund_irr, IRR_DECIMALS)]]
    for currency, irr in sorted(results.currency_irrs.items()):
        rows.append(["Currency", currency, round(irr, IRR_DECIMALS)])
    return headers, rows


def nav_table(results: Results) -> Table:
    """Build the NAV table in long format.

    :param results: The run's results.
    :returns: Headers ``Currency, Date, NAV`` and one row per currency and
        date, with the NAV in local currency rounded to 2 decimal places.
    """
    headers = ["Currency", "Date", "NAV"]
    rows = [
        [currency, point.date, round_money(point.nav)]
        for currency, schedule in sorted(results.nav_schedules.items())
        for point in schedule
    ]
    return headers, rows


def hedge_table(results: Results) -> Table:
    """Build the FX hedge table.

    :param results: The run's results.
    :returns: One row per proposed forward, with the pair, trade and
        delivery dates, direction, notional currency and notional amount.
    """
    headers = [
        "Currency Pair", "Trade Date", "Delivery Date",
        "Direction", "Notional Currency", "Notional Amount",
    ]
    rows = [
        [t.currency_pair, t.trade_date, t.delivery_date,
         t.direction.value, t.notional_currency, t.notional_amount]
        for t in results.hedges
    ]
    return headers, rows


# Sheet / file name -> table builder. The order is the order of the sheets.
TABLES = {"IRR": irr_table, "NAV": nav_table, "FX Hedges": hedge_table}


# --- Console -----------------------------------------------------------------

def _format_table(headers: list[str], rows: list[list[str]]) -> str:
    """Lay out a plain-text table for the console.

    Each column is as wide as its longest cell. Numbers (including
    percentages and comma-grouped amounts) are right-aligned; other text is
    left-aligned.

    :param headers: Column headers.
    :param rows: Rows of already-formatted cell text.
    :returns: The table as one string: header, a dashed rule, then the rows.
    """
    widths = [max(len(str(x)) for x in column) for column in zip(headers, *rows)]

    def is_number(text: str) -> bool:
        return bool(re.fullmatch(r"-?[\d,]+(\.\d+)?%?", text))

    def line(cells):
        return "  ".join(
            cell.rjust(w) if is_number(cell) else cell.ljust(w)
            for cell, w in zip(cells, widths)
        ).rstrip()

    return "\n".join([line(headers), line(["-" * w for w in widths]), *map(line, rows)])


def _money(value: Decimal) -> str:
    """Format an amount for the console, e.g. ``1,234,567.89``.

    :param value: The amount.
    :returns: The amount with thousands separators and 2 decimal places.
    """
    return f"{value:,.2f}"


def print_summary(results: Results) -> None:
    """Print the IRR, NAV and hedge tables to stdout.

    Logs go to stderr, so the printed results never mix with log lines. The
    NAV table is printed wide, with one column per currency, which is easier
    to read on screen than the long format saved to file.

    :param results: The run's results.
    """
    print(f"\n=== {results.fund_name}: IRR ===\n")
    irr_rows = [
        ["Fund (base amounts)", currency, f"{irr:.4%}"] if level == "Fund"
        else ["Currency (local amounts)", currency, f"{irr:.4%}"]
        for level, currency, irr in irr_table(results)[1]
    ]
    print(_format_table(["Level", "Currency", "IRR"], irr_rows))

    # Wide layout for the console: one row per date, one column per currency.
    currencies = sorted(results.nav_schedules)
    navs = {
        (p.currency, p.date): round_money(p.nav)
        for schedule in results.nav_schedules.values() for p in schedule
    }
    dates = sorted({d for _, d in navs})
    print(f"\n=== {results.fund_name}: NAV schedule (local currency) ===\n")
    print(_format_table(
        ["Date", *currencies],
        [[d.isoformat(), *(_money(navs[c, d]) if (c, d) in navs else "" for c in currencies)]
         for d in dates],
    ))

    print(f"\n=== {results.fund_name}: FX forward hedges ({len(results.hedges)} trades) ===\n")
    headers, rows = hedge_table(results)
    print(_format_table(headers, [
        [pair, trade.isoformat(), delivery.isoformat(), direction, currency, _money(notional)]
        for pair, trade, delivery, direction, currency, notional in rows
    ]))
    print()


# --- Files -------------------------------------------------------------------

def output_stem(fund_name: str) -> str:
    """Make a safe file-name prefix from a fund name.

    Example: ``"Fund I"`` becomes ``"fund_i"``.

    :param fund_name: The fund (or table) name.
    :returns: Lower-case letters and digits joined by underscores, or
        ``"fund"`` if nothing is left.
    """
    return re.sub(r"[^a-z0-9]+", "_", fund_name.lower()).strip("_") or "fund"


def write_outputs(results: Results, output_dir: Path, output_format: str = "excel") -> list[Path]:
    """Save the IRR, NAV and hedge tables to files.

    ``"excel"`` writes one workbook, ``<fund>_results.xlsx``, with sheets
    ``IRR``, ``NAV`` and ``FX Hedges``. ``"csv"`` writes three files,
    ``<fund>_irr.csv``, ``<fund>_nav.csv`` and ``<fund>_fx_hedges.csv``.
    Existing files with the same names are replaced.

    :param results: The run's results.
    :param output_dir: Folder to write to; created if it does not exist.
    :param output_format: ``"excel"`` or ``"csv"``.
    :returns: The paths written.
    :raises ValueError: If ``output_format`` is not one of
        :data:`OUTPUT_FORMATS`.
    :raises OSError: If a file cannot be written, e.g. because the workbook
        is open in Excel.
    """
    if output_format not in OUTPUT_FORMATS:
        raise ValueError(f"Output format must be one of {OUTPUT_FORMATS}, got {output_format!r}")
    output_dir.mkdir(parents=True, exist_ok=True)
    tables = {name: build(results) for name, build in TABLES.items()}
    stem = output_stem(results.fund_name)

    if output_format == "excel":
        paths = [_write_workbook(output_dir / f"{stem}_results.xlsx", tables)]
    else:
        paths = [
            _write_csv(output_dir / f"{stem}_{output_stem(name)}.csv", *table)
            for name, table in tables.items()
        ]
    for path in paths:
        logger.info("Wrote %s", path)
    return paths


def _csv_value(value) -> str:
    """Format one cell for a CSV file.

    :param value: A date, Decimal or any other value.
    :returns: ISO format for dates, plain digits for Decimals (never
        scientific notation), ``str`` for everything else.
    """
    if isinstance(value, dt.date):
        return value.isoformat()
    if isinstance(value, Decimal):
        return format(value, "f")  # plain digits, never scientific notation
    return str(value)


def _write_csv(path: Path, headers: list[str], rows: list[list]) -> Path:
    """Write one table as a UTF-8 CSV file.

    Line endings are always ``\\n``, so the same results give the same bytes
    on any operating system, which the md5 tests rely on.

    :param path: File to write.
    :param headers: Column headers.
    :param rows: Table rows.
    :returns: ``path``.
    """
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, lineterminator="\n")
        writer.writerow(headers)
        writer.writerows([[_csv_value(v) for v in row] for row in rows])
    return path


# Excel number formats by column header. Values stay numbers and dates in the
# workbook; only their display is formatted.
EXCEL_FORMATS = {
    "IRR": "0.0000%",
    "Date": "yyyy-mm-dd",
    "Trade Date": "yyyy-mm-dd",
    "Delivery Date": "yyyy-mm-dd",
    "NAV": "#,##0.00",
    "Notional Amount": "#,##0.00",
}


def _write_workbook(path: Path, tables: dict[str, Table]) -> Path:
    """Write the tables to one Excel workbook, one sheet per table.

    Each sheet has a bold, frozen header row and columns sized to their
    contents. Numbers and dates are stored as real values; only their
    display is set, from :data:`EXCEL_FORMATS`.

    :param path: Workbook to write.
    :param tables: Sheet name mapped to ``(headers, rows)``, in sheet order.
    :returns: ``path``.
    """
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for name, (headers, rows) in tables.items():
        sheet = workbook.create_sheet(name)
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        sheet.freeze_panes = "A2"
        for column, header in enumerate(headers, start=1):
            number_format = EXCEL_FORMATS.get(header)
            if number_format:
                for (cell,) in sheet.iter_rows(min_row=2, min_col=column, max_col=column):
                    cell.number_format = number_format
            width = max(len(header), *(len(_csv_value(r[column - 1])) for r in rows)) if rows else len(header)
            sheet.column_dimensions[get_column_letter(column)].width = width + 4
    workbook.save(path)
    return path
