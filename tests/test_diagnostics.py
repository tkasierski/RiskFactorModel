from __future__ import annotations

import pandas as pd

from factor_risk_analysis.data import latest_completed_month_end
from factor_risk_analysis.processing import align_analysis_data


def test_latest_completed_month_end_excludes_partial_current_month():
    cutoff = latest_completed_month_end("2026-08-11", today="2026-08-11")
    assert cutoff == pd.Timestamp("2026-07-31")


def test_latest_completed_month_end_respects_historical_full_month():
    cutoff = latest_completed_month_end("2025-12-31", today="2026-08-11")
    assert cutoff == pd.Timestamp("2025-12-31")


def test_alignment_reports_dropped_months_and_crisis_warning():
    idx = pd.date_range("2020-01-31", periods=30, freq="ME")
    fund = pd.Series(0.01, index=idx, name="Fund")
    factors = pd.DataFrame(
        {
            "Equity": 0.01,
            "Currency": 0.002,
            "Credit": 0.004,
            "Volatility": 0.01,
            "VIX Level": 20.0,
        },
        index=idx,
    )
    factors.loc[idx[5], "Credit"] = float("nan")
    factors.loc[idx[10], "VIX Level"] = 35.0

    result = align_analysis_data(fund, factors, crisis_threshold=30.0)

    assert result.dropped_rows == 1
    assert result.dropped_details.iloc[0]["Date"] == idx[5]
    assert "Credit" in result.dropped_details.iloc[0]["Missing series"]
    assert result.crisis_months == 1
    assert any("crisis observation" in warning for warning in result.warnings)
