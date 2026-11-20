"""Pydantic models: the shape of valid data flowing through the program."""

import datetime as dt
from decimal import Decimal
from enum import Enum

from pydantic import BaseModel, ConfigDict, Field, field_validator


ALLOWED_CURRENCIES = {"EUR", "GBP", "USD"}


class CashflowType(str, Enum):
    INVESTMENT = "Investment"
    INTEREST = "Interest"
    PRINCIPAL_REPAYMENT = "Principal Repayment"


class Cashflow(BaseModel):
    """One row of the input sheet. Aliases map the Excel headers to field names.

    The model only checks data, it never fixes it. Fixes live in excel_io.py's
    cleaning step, which runs only on fields that fail here.
    """

    model_config = ConfigDict(populate_by_name=True)

    id: int = Field(alias="ID")
    fund_name: str = Field(alias="Fund Name")
    date: dt.date = Field(alias="Date")
    cashflow_type: CashflowType = Field(alias="Cashflow Type")
    local_currency: str = Field(alias="Local Currency")
    # Money is Decimal, not float. Pydantic converts floats via their shortest
    # string form, so 0.1 becomes Decimal("0.1") rather than its binary
    # approximation, and NaN/inf are rejected.
    amount_local: Decimal = Field(alias="Cashflow Amount Local")
    amount_base: Decimal = Field(alias="Cashflow Amount Base")
    base_currency: str = Field(alias="Base Currency")

    @field_validator("local_currency", "base_currency")
    @classmethod
    def check_currency(cls, v: str) -> str:
        if v not in ALLOWED_CURRENCIES:
            raise ValueError(f"currency must be one of {sorted(ALLOWED_CURRENCIES)}")
        return v


class NavPoint(BaseModel):
    """One point on a NAV schedule."""

    date: dt.date
    currency: str
    nav: Decimal


class Direction(str, Enum):
    BUY = "Buy"
    SELL = "Sell"


class FxForward(BaseModel):
    """A proposed FX forward hedge trade."""

    currency_pair: str  # e.g. "EURGBP"
    trade_date: dt.date
    delivery_date: dt.date
    direction: Direction
    notional_currency: str
    notional_amount: Decimal
