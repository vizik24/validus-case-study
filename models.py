"""Pydantic models describing valid data as it moves through the program.

The input side is :class:`Cashflow`, one validated row of the Excel sheet.
The output side is :class:`NavPoint`, :class:`FxForward` and
:class:`Results`, which together hold everything a run prints and saves.

The models only *check* data; they never change it. Repairs to bad input
live in :mod:`excel_io`.
"""

import datetime as dt
from decimal import Decimal
from enum import Enum

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


ALLOWED_CURRENCIES = {"EUR", "GBP", "USD"}


class CashflowType(str, Enum):
    """The kinds of cashflow that appear in the input sheet.

    The values are the exact strings used in the ``Cashflow Type`` column.
    """

    INVESTMENT = "Investment"
    INTEREST = "Interest"
    PRINCIPAL_REPAYMENT = "Principal Repayment"


class Cashflow(BaseModel):
    """One row of the input sheet.

    Each field has an alias equal to its Excel column header, so a raw row
    dict from pandas can be passed straight to :meth:`model_validate`.
    ``populate_by_name`` also allows the Python field names, which keeps
    tests short.

    The model only checks data; it never fixes it. Fixes live in the cleaning
    step of :mod:`excel_io`, which runs only on fields that fail here.

    :ivar id: Row identifier from the ``ID`` column; must be unique.
    :ivar fund_name: Fund the cashflow belongs to, e.g. ``"Fund I"``.
    :ivar date: Date the cashflow is paid or received.
    :ivar cashflow_type: Investment, interest or principal repayment.
    :ivar local_currency: Currency of the position, one of
        :data:`ALLOWED_CURRENCIES`.
    :ivar amount_local: Amount in the local currency. Negative is money paid
        out by the fund (an investment); positive is money received.
    :ivar amount_base: The same amount converted to the fund's base currency.
    :ivar base_currency: The fund's reporting currency, one of
        :data:`ALLOWED_CURRENCIES`.
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
        """Reject any currency code outside :data:`ALLOWED_CURRENCIES`.

        The check is exact: ``" gbp "`` and ``"GPB"`` both fail here and are
        left for :func:`excel_io.clean_currency` to repair.

        :param v: The currency code from the row.
        :returns: The code, unchanged.
        :raises ValueError: If the code is not an allowed currency.
        """
        if v not in ALLOWED_CURRENCIES:
            raise ValueError(f"currency must be one of {sorted(ALLOWED_CURRENCIES)}")
        return v


class NavPoint(BaseModel):
    """One point on a currency's NAV schedule.

    :ivar date: Valuation date.
    :ivar currency: Local currency the NAV is expressed in.
    :ivar nav: Net asset value on that date, in the local currency.
    """

    date: dt.date
    currency: str
    nav: Decimal


class Direction(str, Enum):
    """Whether a forward buys or sells its notional currency."""

    BUY = "Buy"
    SELL = "Sell"


class FxForward(BaseModel):
    """A proposed FX forward hedge trade.

    :ivar currency_pair: Pair in market convention, e.g. ``"EURGBP"``.
    :ivar trade_date: Date the forward is traded.
    :ivar delivery_date: Date the forward settles; must be after
        ``trade_date``.
    :ivar direction: Buy or sell, referring to the notional currency.
    :ivar notional_currency: Currency the amount is quoted in; must be one
        side of ``currency_pair``.
    :ivar notional_amount: Amount of the notional currency to deliver or
        receive; must be greater than zero.
    """

    currency_pair: str
    trade_date: dt.date
    delivery_date: dt.date
    direction: Direction
    notional_currency: str
    notional_amount: Decimal = Field(gt=0)

    @model_validator(mode="after")
    def check_trade(self) -> "FxForward":
        """Check the rules that involve more than one field.

        :returns: The trade, unchanged.
        :raises ValueError: If delivery is not after the trade date, or the
            notional currency is not part of the pair.
        """
        if self.delivery_date <= self.trade_date:
            raise ValueError("delivery_date must be after trade_date")
        if self.notional_currency not in self.currency_pair:
            raise ValueError("notional_currency must be one side of currency_pair")
        return self


class Results(BaseModel):
    """Everything one run produces for a fund: what gets printed and saved.

    :ivar fund_name: Fund the results are for.
    :ivar base_currency: The fund's reporting currency.
    :ivar fund_irr: Fund-level IRR on base-currency amounts, as a fraction
        (``0.1`` is 10%).
    :ivar currency_irrs: IRR per local currency, on local-currency amounts.
    :ivar nav_schedules: NAV schedule per local currency.
    :ivar hedges: Proposed FX forwards, in trade order within each currency.
    """

    fund_name: str
    base_currency: str
    fund_irr: float
    currency_irrs: dict[str, float]
    nav_schedules: dict[str, list[NavPoint]]
    hedges: list[FxForward]
