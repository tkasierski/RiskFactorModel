"""Data ingestion and Yahoo Finance retrieval utilities."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Iterable

import numpy as np
import pandas as pd

from .config import FactorDefinition


class DataError(ValueError):
    """Raised when an input data source cannot be used."""


@dataclass(frozen=True)
class MarketSeriesMeta:
    ticker: str
    method: str
    first_observation: pd.Timestamp | None
    last_observation: pd.Timestamp | None
    source: str = "Yahoo Finance via yfinance"


def _as_timestamp(value: date | datetime | str | pd.Timestamp) -> pd.Timestamp:
    ts = pd.Timestamp(value)
    if pd.isna(ts):
        raise DataError(f"Invalid date: {value!r}")
    return ts.normalize()


def latest_completed_month_end(
    end: date | datetime | str | pd.Timestamp,
    today: date | datetime | str | pd.Timestamp | None = None,
) -> pd.Timestamp:
    """Return the latest fully completed calendar month allowed by ``end``.

    Monthly analysis should never use a partial current month. If the requested
    end date itself falls before that month's calendar month-end, the prior
    month is the latest complete month for that requested range.
    """

    requested_end = _as_timestamp(end)
    requested_period = requested_end.to_period("M")
    requested_month_end = requested_period.to_timestamp("M")
    if requested_end < requested_month_end:
        requested_complete = (requested_period - 1).to_timestamp("M")
    else:
        requested_complete = requested_month_end

    today_ts = _as_timestamp(today if today is not None else pd.Timestamp.today())
    last_completed_current = (today_ts.to_period("M") - 1).to_timestamp("M")
    return min(requested_complete, last_completed_current)


def _download_start_for_returns(start: date | datetime | str | pd.Timestamp) -> pd.Timestamp:
    """Return a pre-start date sufficient for first monthly return calculation."""

    return _as_timestamp(start) - pd.DateOffset(days=70)


def _download_end_exclusive(end: date | datetime | str | pd.Timestamp) -> pd.Timestamp:
    """Yahoo's end date is exclusive. Add a small buffer to include month-end data."""

    return _as_timestamp(end) + pd.DateOffset(days=7)


def _flatten_yfinance_columns(df: pd.DataFrame, ticker: str) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        # yfinance may return either (Price, Ticker) or (Ticker, Price).
        if ticker in df.columns.get_level_values(0):
            df = df.xs(ticker, axis=1, level=0, drop_level=True)
        elif ticker in df.columns.get_level_values(-1):
            df = df.xs(ticker, axis=1, level=-1, drop_level=True)
        else:
            df.columns = [col[-1] if isinstance(col, tuple) else col for col in df.columns]
    return df


def download_yahoo_history(
    ticker: str,
    start: date | datetime | str | pd.Timestamp,
    end: date | datetime | str | pd.Timestamp,
) -> pd.DataFrame:
    """Download daily Yahoo Finance data for one ticker.

    The import is local so tests and workbook code can run in environments where
    yfinance is not installed yet.
    """

    try:
        import yfinance as yf  # type: ignore
    except ImportError as exc:  # pragma: no cover - exercised only without dependency
        raise DataError("yfinance is required to download public market data.") from exc

    start_ts = _download_start_for_returns(start)
    end_ts = _download_end_exclusive(end)

    data = yf.download(
        ticker,
        start=start_ts.strftime("%Y-%m-%d"),
        end=end_ts.strftime("%Y-%m-%d"),
        auto_adjust=False,
        actions=True,
        progress=False,
        threads=False,
    )

    if data is None or data.empty:
        raise DataError(f"No Yahoo Finance data returned for ticker {ticker!r}.")

    data = _flatten_yfinance_columns(data, ticker)
    data.index = pd.to_datetime(data.index).tz_localize(None)
    return data.sort_index()


def _price_column_for_method(history: pd.DataFrame, method: str) -> str:
    """Choose the price column used to calculate returns."""

    if method == "total_return" and "Adj Close" in history.columns:
        return "Adj Close"
    if "Adj Close" in history.columns:
        return "Adj Close"
    if "Close" in history.columns:
        return "Close"
    raise DataError("Downloaded market data is missing both Adj Close and Close columns.")


def month_end_series_from_daily(history: pd.DataFrame, value_col: str) -> pd.Series:
    if value_col not in history.columns:
        raise DataError(f"Column {value_col!r} not found in downloaded data.")
    values = history[value_col].dropna()
    if values.empty:
        raise DataError(f"Column {value_col!r} contains no usable data.")
    monthly = values.resample("ME").last().dropna()
    monthly.index = monthly.index.to_period("M").to_timestamp("M")
    return monthly


def monthly_returns_from_history(
    history: pd.DataFrame,
    method: str,
    start: date | datetime | str | pd.Timestamp,
    end: date | datetime | str | pd.Timestamp,
) -> pd.Series:
    """Calculate returns using fully completed calendar months only."""

    price_col = _price_column_for_method(history, method)
    monthly_values = month_end_series_from_daily(history, price_col)
    returns = monthly_values.pct_change().dropna()

    start_month = _as_timestamp(start).to_period("M").to_timestamp("M")
    end_month = latest_completed_month_end(end)
    returns = returns.loc[(returns.index >= start_month) & (returns.index <= end_month)]
    return returns.astype(float)


def download_monthly_returns(
    ticker: str,
    method: str,
    start: date | datetime | str | pd.Timestamp,
    end: date | datetime | str | pd.Timestamp,
) -> tuple[pd.Series, MarketSeriesMeta]:
    history = download_yahoo_history(ticker, start, end)
    returns = monthly_returns_from_history(history, method=method, start=start, end=end)
    meta = MarketSeriesMeta(
        ticker=ticker,
        method=method,
        first_observation=returns.index.min() if not returns.empty else None,
        last_observation=returns.index.max() if not returns.empty else None,
    )
    return returns, meta


def download_month_end_level(
    ticker: str,
    start: date | datetime | str | pd.Timestamp,
    end: date | datetime | str | pd.Timestamp,
) -> tuple[pd.Series, MarketSeriesMeta]:
    history = download_yahoo_history(ticker, start, end)
    value_col = "Close" if "Close" in history.columns else _price_column_for_method(history, "pct_change")
    levels = month_end_series_from_daily(history, value_col)
    start_month = _as_timestamp(start).to_period("M").to_timestamp("M")
    end_month = latest_completed_month_end(end)
    levels = levels.loc[(levels.index >= start_month) & (levels.index <= end_month)]
    meta = MarketSeriesMeta(
        ticker=ticker,
        method="month_end_level",
        first_observation=levels.index.min() if not levels.empty else None,
        last_observation=levels.index.max() if not levels.empty else None,
    )
    return levels.astype(float), meta


def build_factor_frame(
    factors: Iterable[FactorDefinition],
    start: date | datetime | str | pd.Timestamp,
    end: date | datetime | str | pd.Timestamp,
    crisis_ticker: str,
) -> tuple[pd.DataFrame, list[MarketSeriesMeta]]:
    """Download and combine factor returns and the VIX level used for crisis classification."""

    pieces: list[pd.Series] = []
    metadata: list[MarketSeriesMeta] = []
    for factor in factors:
        series, meta = download_monthly_returns(factor.ticker, factor.method, start, end)
        series.name = factor.name
        pieces.append(series)
        metadata.append(meta)

    vix_level, vix_meta = download_month_end_level(crisis_ticker, start, end)
    vix_level.name = "VIX Level"
    pieces.append(vix_level)
    metadata.append(vix_meta)

    frame = pd.concat(pieces, axis=1).sort_index()
    return frame, metadata


def normalize_to_month_end(series: pd.Series) -> pd.Series:
    """Normalize a dated return series to unique month-end index values."""

    s = series.copy()
    s.index = pd.to_datetime(s.index, errors="coerce")
    s = s[~s.index.isna()]
    s.index = s.index.to_period("M").to_timestamp("M")
    s = s.sort_index()
    # If more than one value is provided for a month, use the last non-null value.
    s = s.groupby(level=0).last()
    return s.dropna().astype(float)


def autodetect_return_scale(values: pd.Series) -> tuple[pd.Series, str]:
    """Parse and scale uploaded return values.

    Returns are normalized to decimal units. Examples:
    - "2.5%" -> 0.025
    - 0.025 -> 0.025
    - 2.5 -> 0.025
    """

    raw = values.copy()
    contains_percent_symbol = raw.astype(str).str.contains("%", regex=False, na=False).any()

    if contains_percent_symbol:
        parsed = (
            raw.astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
            .replace({"": np.nan, "nan": np.nan, "None": np.nan})
        )
        numeric = pd.to_numeric(parsed, errors="coerce") / 100.0
        return numeric, "percent-string"

    numeric = pd.to_numeric(raw, errors="coerce")
    non_null = numeric.dropna()
    if non_null.empty:
        raise DataError("Uploaded return column does not contain numeric returns.")

    abs_q95 = non_null.abs().quantile(0.95)
    max_abs = non_null.abs().max()

    # Monthly returns shown as 2.5 for 2.5% are common. Decimal monthly returns
    # rarely exceed +/-100%, so values above one are treated as percentage points.
    if abs_q95 > 1.0 or max_abs > 2.0:
        return numeric / 100.0, "percentage-points"

    return numeric, "decimal"


def read_spreadsheet_preview(file_obj: BinaryIO | BytesIO, suffix: str) -> dict[str, pd.DataFrame]:
    """Read uploaded tabular data into pandas for Streamlit preview/column selection."""

    suffix = suffix.lower()
    file_obj.seek(0)
    if suffix == ".csv":
        return {"CSV": pd.read_csv(file_obj)}
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        xls = pd.ExcelFile(file_obj)
        return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
    raise DataError("Unsupported upload type. Use .xlsx, .xlsm, .xls, or .csv.")


def guess_date_column(df: pd.DataFrame) -> str | None:
    best_col: str | None = None
    best_count = -1
    for col in df.columns:
        parsed = pd.to_datetime(df[col], errors="coerce")
        count = int(parsed.notna().sum())
        if count > best_count:
            best_col = str(col)
            best_count = count
    return best_col


def guess_return_column(df: pd.DataFrame, date_col: str | None = None) -> str | None:
    candidates = [col for col in df.columns if str(col) != str(date_col)]
    if not candidates:
        return None

    # Prefer column names that look like return data.
    for col in candidates:
        lowered = str(col).lower()
        if "return" in lowered or "ror" in lowered or "performance" in lowered:
            return str(col)

    best_col: str | None = None
    best_count = -1
    for col in candidates:
        parsed = pd.to_numeric(
            df[col].astype(str).str.replace("%", "", regex=False).str.replace(",", "", regex=False),
            errors="coerce",
        )
        count = int(parsed.notna().sum())
        if count > best_count:
            best_col = str(col)
            best_count = count
    return best_col


def uploaded_returns_to_series(
    df: pd.DataFrame,
    date_col: str,
    return_col: str,
    start: date | datetime | str | pd.Timestamp,
    end: date | datetime | str | pd.Timestamp,
) -> tuple[pd.Series, str]:
    """Create a normalized monthly return series from selected upload columns."""

    if date_col not in df.columns:
        raise DataError(f"Date column {date_col!r} not found.")
    if return_col not in df.columns:
        raise DataError(f"Return column {return_col!r} not found.")

    dates = pd.to_datetime(df[date_col], errors="coerce")
    scaled_returns, scale_method = autodetect_return_scale(df[return_col])
    series = pd.Series(scaled_returns.to_numpy(), index=dates, name="Fund")
    series = normalize_to_month_end(series)

    start_month = _as_timestamp(start).to_period("M").to_timestamp("M")
    end_month = latest_completed_month_end(end)
    series = series.loc[(series.index >= start_month) & (series.index <= end_month)]
    if series.empty:
        raise DataError("No uploaded return observations fall inside the selected date range after excluding incomplete calendar months.")
    return series, scale_method


def file_suffix(uploaded_file_name: str) -> str:
    return Path(uploaded_file_name).suffix.lower()
