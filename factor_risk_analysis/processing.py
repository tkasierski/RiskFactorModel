"""Alignment, validation, and helper metrics for the app preview."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import pandas as pd

from .config import LIMITED_HISTORY_OBSERVATIONS, MIN_OBSERVATIONS, PREFERRED_OBSERVATIONS
from .data import DataError


@dataclass(frozen=True)
class AlignmentResult:
    data: pd.DataFrame
    warnings: tuple[str, ...]
    dropped_rows: int


def align_analysis_data(
    fund_returns: pd.Series,
    factor_frame: pd.DataFrame,
    benchmark_returns: pd.Series | None = None,
) -> AlignmentResult:
    """Inner-join all model series to a common month-end index."""

    fund = fund_returns.rename("Fund")
    pieces: list[pd.Series | pd.DataFrame] = [fund, factor_frame]
    if benchmark_returns is not None:
        pieces.append(benchmark_returns.rename("Benchmark"))

    combined_before_drop = pd.concat(pieces, axis=1).sort_index()
    aligned = combined_before_drop.dropna(how="any")

    if "Benchmark" in aligned.columns:
        ordered_cols = ["Fund", "Benchmark"] + [col for col in aligned.columns if col not in {"Fund", "Benchmark"}]
        aligned = aligned[ordered_cols]

    dropped_rows = int(len(combined_before_drop) - len(aligned))
    warnings: list[str] = []
    n = len(aligned)

    if n < MIN_OBSERVATIONS:
        raise DataError(
            f"Only {n} complete monthly observations are available after alignment. "
            f"At least {MIN_OBSERVATIONS} are required."
        )
    if n < LIMITED_HISTORY_OBSERVATIONS:
        warnings.append(
            f"Only {n} complete monthly observations are available. Regression will run, "
            "but the sample is short."
        )
    elif n < PREFERRED_OBSERVATIONS:
        warnings.append(
            f"{n} complete monthly observations are available. This is usable but below the "
            f"preferred {PREFERRED_OBSERVATIONS}+ month history."
        )

    if dropped_rows > 0:
        warnings.append(f"{dropped_rows} month(s) were dropped because one or more required series were missing.")

    rolling_annual_obs = max(n - 11, 0)
    if rolling_annual_obs < 30:
        warnings.append(
            f"Annual historical VaR/CVaR will use only {rolling_annual_obs} rolling 12-month observations. "
            "Treat the tail statistics as limited-sample estimates."
        )

    return AlignmentResult(data=aligned, warnings=tuple(warnings), dropped_rows=dropped_rows)


def quick_summary(data: pd.DataFrame) -> pd.DataFrame:
    """Small summary table shown in Streamlit before downloading the workbook."""

    rows = []
    for col in [c for c in ["Fund", "Benchmark"] if c in data.columns]:
        r = data[col].dropna()
        rows.append(
            {
                "Series": col,
                "Observations": len(r),
                "First month": r.index.min().date() if len(r) else None,
                "Last month": r.index.max().date() if len(r) else None,
                "Monthly avg return": r.mean(),
                "Monthly risk": r.std(ddof=1),
                "Annualized avg return": r.mean() * 12,
                "Annualized risk": r.std(ddof=1) * (12 ** 0.5),
            }
        )
    return pd.DataFrame(rows)
