"""Input and output: reading/validating the Excel file, printing and writing results."""

import datetime as dt
import logging
from collections import Counter
from pathlib import Path

import pandas as pd
from pydantic import ValidationError

from models import Cashflow, FxForward, NavPoint

logger = logging.getLogger(__name__)


# --- Input -------------------------------------------------------------------

class DataValidationError(Exception):
    """Raised when one or more input rows cannot be turned into a Cashflow."""


def read_excel_rows(path: Path) -> list[dict]:
    """Read the sheet and return one raw dict per row."""
    # dtype=object stops pandas guessing column types, so all cleaning left 
    # to pydantic.
    logger.info("Reading %s", path)
    df = pd.read_excel(path, sheet_name=0, dtype=object)
    logger.info("Read %d rows x %d columns", *df.shape)
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
    """Try to turn every raw row into a Cashflow.

    Returns the rows that passed plus the errors for the ones that failed,
    keyed by row index. Nothing is raised here, so the caller can decide
    whether to clean and retry or to stop.
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
    """One readable line per error, with its Excel row number and ID."""
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
    """Needs every row at once, so it can't live in the model."""
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
    """Parse day-first text dates with stray spaces or a trailing backtick."""
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
    """Normalise case and whitespace, then fix known typos."""
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


def clean_rows(rows: list[dict], errors: RowErrors) -> tuple[list[dict], list[str]]:
    """Apply the matching cleaner to each field that failed validation.

    Returns a new list of rows (the input is not changed) and a log of every
    value that was changed, for the audit trail.
    """
    cleaned = [dict(row) for row in rows]
    corrections: list[str] = []
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
                line = f"ID {rows[i].get('ID')}: {field} {old!r} -> {_show(new)}"
                corrections.append(line)
                logger.info("Corrected %s", line)
            else:
                logger.debug("ID %s: %s %r could not be cleaned", rows[i].get("ID"), field, old)
    logger.info("Applied %d correction(s)", len(corrections))
    return cleaned, corrections


def _show(value) -> str:
    """Readable form for log lines: ISO dates, repr for everything else."""
    return value.isoformat() if isinstance(value, dt.date) else repr(value)


def filter_fund(cashflows: list[Cashflow], fund_name: str) -> list[Cashflow]:
    """Keep only the cashflows belonging to one fund."""
    matches = [cf for cf in cashflows if cf.fund_name == fund_name]
    if not matches:
        available = sorted({cf.fund_name for cf in cashflows})
        raise ValueError(f"No cashflows found for {fund_name!r}. Available funds: {available}")
    logger.info("Selected %d of %d cashflows for %r", len(matches), len(cashflows), fund_name)
    _log_fund_profile(matches)
    return matches


def _log_fund_profile(cashflows: list[Cashflow]) -> None:
    """Debug-level overview of what the selected fund contains."""
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


def load_cashflows(path: Path, fund_name: str) -> tuple[list[Cashflow], list[str]]:
    """Read -> validate -> (clean failing fields -> validate again) -> filter.

    Returns the fund's cashflows and the log of corrections made. Raises
    DataValidationError if any row is still invalid after cleaning: silently
    dropping a cashflow would change the IRR without anyone noticing.
    """
    rows = read_excel_rows(path)
    cashflows, errors = validate_rows(rows)

    corrections: list[str] = []
    if errors:
        for line in format_errors(rows, errors):
            logger.debug("Validation error: %s", line)
        logger.info("Cleaning the fields that failed, then validating again")
        rows, corrections = clean_rows(rows, errors)
        cashflows, errors = validate_rows(rows)

    problems = format_errors(rows, errors) + check_duplicate_ids(cashflows)
    if problems:
        raise DataValidationError(
            f"{len(problems)} problem(s) in the input data:\n" + "\n".join(problems)
        )
    return filter_fund(cashflows, fund_name), corrections


# --- Output ------------------------------------------------------------------

def print_summary(
    fund_irr: float | None,
    currency_irrs: dict[str, float] | None,
    schedules: dict[str, list[NavPoint]] | None,
    trades: list[FxForward] | None,
) -> None:
    """Print a readable summary to the console."""
    # TODO
    return None


def write_outputs(
    output_dir: Path,
    fund_irr: float | None,
    currency_irrs: dict[str, float] | None,
    schedules: dict[str, list[NavPoint]] | None,
    trades: list[FxForward] | None,
) -> None:
    """Write results to files (CSV / Excel) in output_dir."""
    # TODO
    return None
