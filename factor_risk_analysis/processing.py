"""Alignment, validation, and helper metrics for the app preview."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .config import LIMITED_HISTORY_OBSERVATIONS, MIN_OBSERVATIONS, PREFERRED_OBSERVATIONS
from .data import DataError


@dataclass(frozen=True)
class AlignmentResult:
    data: pd.DataFrame
    warnings: tuple[str, ...]
    dropped_rows: int
    dropped_details: pd.DataFrame
    crisis_months: int


def align_analysis_data(
    fund_returns: pd.Series,
    factor_frame: pd.DataFrame,
    benchmark_returns: pd.Series | None = None,
    crisis_threshold: float = 30.0,
) -> AlignmentResult:
    """Align all model series to a common month-end index and report exclusions."""

    fund = fund_returns.rename("Fund")
    pieces: list[pd.Series | pd.DataFrame] = [fund, factor_frame]
    if benchmark_returns is not None:
        pieces.append(benchmark_returns.rename("Benchmark"))

    combined_before_drop = pd.concat(pieces, axis=1).sort_index()
    missing_mask = combined_before_drop.isna().any(axis=1)
    dropped = combined_before_drop.loc[missing_mask]
    aligned = combined_before_drop.loc[~missing_mask].copy()

    if "Benchmark" in aligned.columns:
        ordered_cols = ["Fund", "Benchmark"] + [col for col in aligned.columns if col not in {"Fund", "Benchmark"}]
        aligned = aligned[ordered_cols]

    dropped_details = pd.DataFrame(columns=["Date", "Missing series"])
    if not dropped.empty:
        rows = []
        for idx, row in dropped.iterrows():
            missing = [str(col) for col in dropped.columns if pd.isna(row[col])]
            rows.append({"Date": pd.Timestamp(idx), "Missing series": ", ".join(missing)})
        dropped_details = pd.DataFrame(rows)

    dropped_rows = int(len(dropped))
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

    crisis_months = 0
    if "VIX Level" in aligned.columns:
        crisis_months = int((aligned["VIX Level"] > crisis_threshold).sum())
        if crisis_months < 20:
            warnings.append(
                f"Conditional factor estimates are based on only {crisis_months} crisis observation(s) "
                f"with month-end VIX above {crisis_threshold:g}; incremental crisis betas may be statistically unstable."
            )

    rolling_annual_obs = max(n - 11, 0)
    if rolling_annual_obs < 30:
        warnings.append(
            f"Annual historical VaR/CVaR will use only {rolling_annual_obs} rolling 12-month observations. "
            "Treat the tail statistics as limited-sample estimates."
        )

    return AlignmentResult(
        data=aligned,
        warnings=tuple(warnings),
        dropped_rows=dropped_rows,
        dropped_details=dropped_details,
        crisis_months=crisis_months,
    )


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
