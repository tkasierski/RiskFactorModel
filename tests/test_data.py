from __future__ import annotations

import pandas as pd

from factor_risk_analysis.data import autodetect_return_scale, normalize_to_month_end, uploaded_returns_to_series


def test_autodetect_percent_strings():
    values, method = autodetect_return_scale(pd.Series(["1.0%", "-2.5%", "0.0%"]))
    assert method == "percent-string"
    assert values.tolist() == [0.01, -0.025, 0.0]


def test_autodetect_percentage_points():
    values, method = autodetect_return_scale(pd.Series([1.0, -2.5, 0.0]))
    assert method == "percentage-points"
    assert values.round(6).tolist() == [0.01, -0.025, 0.0]


def test_autodetect_decimal_returns():
    values, method = autodetect_return_scale(pd.Series([0.01, -0.025, 0.0]))
    assert method == "decimal"
    assert values.tolist() == [0.01, -0.025, 0.0]


def test_normalize_to_month_end_uses_last_observation_in_duplicate_month():
    series = pd.Series(
        [0.01, 0.02, 0.03],
        index=pd.to_datetime(["2024-01-15", "2024-01-31", "2024-02-29"]),
    )
    normalized = normalize_to_month_end(series)
    assert normalized.index.tolist() == [pd.Timestamp("2024-01-31"), pd.Timestamp("2024-02-29")]
    assert normalized.iloc[0] == 0.02


def test_uploaded_returns_to_series_date_filter():
    df = pd.DataFrame(
        {
            "Date": ["2024-01-31", "2024-02-29", "2024-03-31"],
            "Return": ["1.0%", "2.0%", "3.0%"],
        }
    )
    series, method = uploaded_returns_to_series(df, "Date", "Return", "2024-02-01", "2024-03-31")
    assert method == "percent-string"
    assert series.index.tolist() == [pd.Timestamp("2024-02-29"), pd.Timestamp("2024-03-31")]
    assert series.tolist() == [0.02, 0.03]
