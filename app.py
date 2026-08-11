from __future__ import annotations

from datetime import date
from io import BytesIO

import pandas as pd
import streamlit as st

from factor_risk_analysis.audit import add_dropped_month_audit
from factor_risk_analysis.config import (
    DEFAULT_CRISIS_THRESHOLD,
    DEFAULT_CRISIS_TICKER,
    DEFAULT_FACTORS,
    DEFAULT_ROLLING_WINDOW,
    FactorDefinition,
)
from factor_risk_analysis.data import (
    DataError,
    build_factor_frame,
    download_monthly_returns,
    file_suffix,
    guess_date_column,
    guess_return_column,
    latest_completed_month_end,
    read_spreadsheet_preview,
    uploaded_returns_to_series,
)
from factor_risk_analysis.processing import align_analysis_data, quick_summary
from factor_risk_analysis.workbook import WorkbookSettings, create_workbook

st.set_page_config(page_title="Factor Risk Analysis", layout="wide")


@st.cache_data(show_spinner=False)
def cached_public_returns(ticker: str, method: str, start: date, end: date):
    return download_monthly_returns(ticker, method, start, end)


@st.cache_data(show_spinner=False)
def cached_factor_frame(factors_payload: tuple[tuple[str, str, str, str], ...], start: date, end: date, crisis_ticker: str):
    factors = tuple(FactorDefinition(name=name, ticker=ticker, method=method, description=description) for name, ticker, method, description in factors_payload)
    return build_factor_frame(factors, start, end, crisis_ticker)


def build_factor_definitions_from_sidebar() -> tuple[FactorDefinition, ...]:
    with st.sidebar.expander("Advanced factor settings", expanded=False):
        st.caption("Defaults reflect the agreed model specification. Change these only if you want to test alternative proxies.")
        edited = []
        for factor in DEFAULT_FACTORS:
            ticker = st.text_input(f"{factor.name} ticker", value=factor.ticker, key=f"factor_{factor.name}")
            edited.append(
                FactorDefinition(
                    name=factor.name,
                    ticker=ticker.strip(),
                    method=factor.method,
                    description=factor.description,
                )
            )
    return tuple(edited)


def render_upload_controls(start_date: date, end_date: date):
    uploaded = st.file_uploader("Upload fund return file", type=["xlsx", "xlsm", "xls", "csv"])
    if uploaded is None:
        return None

    try:
        sheets = read_spreadsheet_preview(BytesIO(uploaded.getvalue()), file_suffix(uploaded.name))
    except Exception as exc:
        st.error(f"Could not read uploaded file: {exc}")
        return None

    sheet_name = next(iter(sheets.keys()))
    if len(sheets) > 1:
        sheet_name = st.selectbox("Sheet", list(sheets.keys()))

    df = sheets[sheet_name]
    if df.empty:
        st.error("The selected sheet is empty.")
        return None

    guessed_date = guess_date_column(df)
    guessed_return = guess_return_column(df, guessed_date)

    columns = [str(c) for c in df.columns]
    col1, col2 = st.columns(2)
    with col1:
        date_col = st.selectbox("Date column", columns, index=columns.index(guessed_date) if guessed_date in columns else 0)
    with col2:
        fallback_index = columns.index(guessed_return) if guessed_return in columns else min(1, len(columns) - 1)
        return_col = st.selectbox("Monthly return column", columns, index=fallback_index)

    st.dataframe(df.head(10), use_container_width=True)

    try:
        fund_returns, scale_method = uploaded_returns_to_series(df, date_col, return_col, start_date, end_date)
    except Exception as exc:
        st.error(f"Could not parse uploaded returns: {exc}")
        return None

    return fund_returns, scale_method, f"Uploaded file: {uploaded.name}; sheet: {sheet_name}; columns: {date_col}, {return_col}"


def main() -> None:
    st.title("Factor Risk Analysis")
    st.write(
        "Generate an auditable Microsoft Excel workbook using a standard and conditional four-factor regression. "
        "The conditional model treats months with month-end VIX above the selected threshold as crisis months."
    )

    st.sidebar.header("Analysis inputs")
    mode = st.sidebar.radio("Fund input type", ["Upload monthly returns", "Public ticker"], index=0)

    today = date.today()
    default_start = date(today.year - 7, 1, 1)
    default_end = today
    start_date = st.sidebar.date_input("Start date", value=default_start)
    end_date = st.sidebar.date_input("End date", value=default_end)
    if start_date >= end_date:
        st.error("Start date must be before end date.")
        return

    completed_month = latest_completed_month_end(end_date)
    st.sidebar.caption(f"Latest completed month included: {completed_month.strftime('%Y-%m-%d')}")

    risk_free_pct = st.sidebar.number_input("Annual risk-free rate (%)", min_value=-10.0, max_value=50.0, value=0.0, step=0.25, format="%.2f")
    risk_free_rate = risk_free_pct / 100.0

    rolling_window = st.sidebar.selectbox("Rolling regression window", [12, 24, 36, 60], index=[12, 24, 36, 60].index(DEFAULT_ROLLING_WINDOW))
    crisis_threshold = st.sidebar.number_input("Crisis VIX threshold", min_value=0.0, max_value=100.0, value=float(DEFAULT_CRISIS_THRESHOLD), step=1.0)

    factors = build_factor_definitions_from_sidebar()
    factors_payload = tuple((factor.name, factor.ticker, factor.method, factor.description) for factor in factors)

    benchmark_ticker = st.sidebar.text_input("Optional benchmark ticker", value="").strip().upper()

    fund_returns = None
    fund_name = "Fund"
    fund_source = ""
    return_scale_method = None

    if mode == "Upload monthly returns":
        st.subheader("Uploaded private fund returns")
        fund_name = st.text_input("Fund name", value="Uploaded Fund")
        parsed = render_upload_controls(start_date, end_date)
        if parsed is not None:
            fund_returns, return_scale_method, fund_source = parsed
    else:
        st.subheader("Public fund / security ticker")
        ticker = st.text_input("Fund/security ticker", value="").strip().upper()
        fund_name = ticker or "Public Fund"
        if ticker:
            with st.spinner(f"Downloading {ticker} monthly total returns..."):
                try:
                    fund_returns, fund_meta = cached_public_returns(ticker, "total_return", start_date, end_date)
                    fund_source = f"Yahoo Finance via yfinance; ticker: {ticker}; method: total_return"
                except DataError as exc:
                    st.error(str(exc))
                    return
                except Exception as exc:
                    st.error(f"Unexpected error downloading {ticker}: {exc}")
                    return

    if fund_returns is None:
        st.info("Provide a fund input above to continue.")
        return

    generate = st.button("Generate Excel workbook", type="primary")
    if not generate:
        st.stop()

    with st.spinner("Downloading factors and building workbook..."):
        try:
            factor_frame, factor_metadata = cached_factor_frame(factors_payload, start_date, end_date, DEFAULT_CRISIS_TICKER)

            benchmark_returns = None
            if benchmark_ticker:
                benchmark_returns, _ = cached_public_returns(benchmark_ticker, "total_return", start_date, end_date)

            alignment = align_analysis_data(
                fund_returns,
                factor_frame,
                benchmark_returns,
                crisis_threshold=float(crisis_threshold),
            )
            aligned = alignment.data

            settings = WorkbookSettings(
                fund_name=fund_name,
                input_mode=mode,
                fund_source=fund_source,
                requested_start=str(start_date),
                requested_end=str(end_date),
                risk_free_rate_annual=risk_free_rate,
                rolling_window=int(rolling_window),
                crisis_threshold=float(crisis_threshold),
                factors=factors,
                benchmark_name=benchmark_ticker if benchmark_ticker else None,
                return_scale_method=return_scale_method,
                data_warnings=alignment.warnings,
                dropped_rows=alignment.dropped_rows,
                factor_metadata=tuple(meta.__dict__ for meta in factor_metadata),
            )

            workbook_bytes = create_workbook(aligned, settings)
            workbook_bytes = add_dropped_month_audit(workbook_bytes, alignment.dropped_details)
        except DataError as exc:
            st.error(str(exc))
            return
        except Exception as exc:
            st.exception(exc)
            return

    st.success("Workbook generated.")

    summary = quick_summary(aligned)
    if not summary.empty:
        pct_cols = ["Monthly avg return", "Monthly risk", "Annualized avg return", "Annualized risk"]
        display = summary.copy()
        for col in pct_cols:
            display[col] = display[col].map(lambda x: f"{x:.2%}" if pd.notna(x) else "")
        st.dataframe(display, use_container_width=True)

    if alignment.warnings:
        for warning in alignment.warnings:
            st.warning(warning)

    if not alignment.dropped_details.empty:
        with st.expander("Dropped month details"):
            dropped_display = alignment.dropped_details.copy()
            dropped_display["Date"] = pd.to_datetime(dropped_display["Date"]).dt.strftime("%Y-%m-%d")
            st.dataframe(dropped_display, use_container_width=True)

    safe_name = (fund_name or "factor_risk_analysis").replace("/", "_").replace("\\", "_").replace(" ", "_")
    st.download_button(
        label="Download Excel workbook",
        data=workbook_bytes,
        file_name=f"{safe_name}_factor_risk_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
