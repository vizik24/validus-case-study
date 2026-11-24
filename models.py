"""Pydantic models for the input rows and the results. They check data, never fix it."""

import datetime as dt
from decimal import Decimal
from enum import Enum

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


ALLOWED_CURRENCIES = {"EUR", "GBP", "USD"}


class CashflowType(str, Enum):
    """Values of the ``Cashflow Type`` column."""

    INVESTMENT = "Investment"
    INTEREST = "Interest"
    PRINCIPAL_REPAYMENT = "Principal Repayment"


class Cashflow(BaseModel):
    """One input row. Aliases are the Excel headers, so a raw row validates directly.

    :ivar amount_local: Amount in the position's currency; negative is paid out.
    :ivar amount_base: The same amount in the fund's base (reporting) currency.
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
        """Reject codes outside :data:`ALLOWED_CURRENCIES`; cleaning fixes typos later."""
        if v not in ALLOWED_CURRENCIES:
            raise ValueError(f"currency must be one of {sorted(ALLOWED_CURRENCIES)}")
        return v


class NavPoint(BaseModel):
    """One point on a currency's NAV schedule, in local currency."""

    date: dt.date
    currency: str
    nav: Decimal


class Direction(str, Enum):
    """Whether a forward buys or sells its notional currency."""

    BUY = "Buy"
    SELL = "Sell"


class FxForward(BaseModel):
    """A proposed FX forward hedge trade.

    :ivar currency_pair: Market convention, e.g. ``"EURGBP"``.
    :ivar direction: Buy or sell, referring to the notional currency.
    """

    currency_pair: str
    trade_date: dt.date
    delivery_date: dt.date
    direction: Direction
    notional_currency: str
    notional_amount: Decimal = Field(gt=0)

    @model_validator(mode="after")
    def check_trade(self) -> "FxForward":
        """Delivery must follow the trade, and the notional must be in the pair."""
        if self.delivery_date <= self.trade_date:
            raise ValueError("delivery_date must be after trade_date")
        if self.notional_currency not in self.currency_pair:
            raise ValueError("notional_currency must be one side of currency_pair")
        return self


class Results(BaseModel):
    """Everything one run produces for a fund: what gets printed and saved."""

    fund_name: str
    base_currency: str
    fund_irr: float
    currency_irrs: dict[str, float]
    nav_schedules: dict[str, list[NavPoint]]
    hedges: list[FxForward]
