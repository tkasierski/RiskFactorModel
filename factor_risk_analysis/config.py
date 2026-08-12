"""Default model configuration."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

ReturnMethod = Literal["total_return", "pct_change"]


@dataclass(frozen=True)
class FactorDefinition:
    """Public market equivalent used as a factor."""

    name: str
    ticker: str
    method: ReturnMethod
    description: str


DEFAULT_FACTORS: tuple[FactorDefinition, ...] = (
    FactorDefinition(
        name="Equity",
        ticker="SPY",
        method="total_return",
        description="Equity risk; SPDR S&P 500 ETF Trust monthly total return.",
    ),
    FactorDefinition(
        name="Currency",
        ticker="DX-Y.NYB",
        method="pct_change",
        description="Currency risk; U.S. Dollar Index monthly percentage change.",
    ),
    FactorDefinition(
        name="Credit",
        ticker="HYG",
        method="total_return",
        description=(
            "High-yield credit-market risk proxy; iShares iBoxx $ High Yield Corporate Bond ETF monthly total return. "
            "A positive beta means the fund tends to participate with HYG, so HYG weakness is expected to hurt the fund, "
            "all else equal. This is not a pure credit-spread beta because HYG also reflects rates/duration, carry and liquidity."
        ),
    ),
    FactorDefinition(
        name="Volatility",
        ticker="^VIX",
        method="pct_change",
        description="Volatility risk; CBOE Volatility Index monthly percentage change.",
    ),
)

DEFAULT_CRISIS_TICKER = "^VIX"
DEFAULT_CRISIS_THRESHOLD = 30.0
DEFAULT_ROLLING_WINDOW = 36

MIN_OBSERVATIONS = 24
LIMITED_HISTORY_OBSERVATIONS = 36
PREFERRED_OBSERVATIONS = 60

SHEET_SUMMARY = "Summary"
SHEET_FUND = "Fund Returns"
SHEET_FACTORS = "Factor Data"
SHEET_REGRESSION = "Regression"
SHEET_CONDITIONAL = "Conditional Regression"
SHEET_RISK = "Risk Statistics"
SHEET_DIAGNOSTICS = "Diagnostics"
SHEET_ROLLING = "Rolling Analysis"
SHEET_DEFINITIONS = "Definitions"

FACTOR_NAMES = tuple(factor.name for factor in DEFAULT_FACTORS)
