# Factor Risk Analysis

Streamlit application for generating an auditable Microsoft Excel workbook that breaks down hedge fund or public fund returns using a conditional factor risk model.

The default framework uses four public market equivalents:

| Exposure | Default ticker | Return treatment |
|---|---:|---|
| Equity | `SPY` | Monthly total return from Yahoo-adjusted close |
| Currency | `DX-Y.NYB` | Monthly percentage change |
| Credit | `HYG` | Monthly total return from Yahoo-adjusted close |
| Volatility | `^VIX` | Monthly percentage change |

A crisis month is defined as a month in which month-end VIX is above 30 by default. The workbook includes both a standard four-factor regression and a conditional regression with factor-by-crisis interaction terms.

## Features

- Analyze either:
  - an uploaded private fund return spreadsheet with one date column and one monthly return column; or
  - a public ticker downloaded from Yahoo Finance.
- Optional benchmark ticker.
- User-selected annual risk-free rate.
- Month-end alignment of fund, benchmark, and factor data.
- Minimum history controls:
  - fewer than 24 monthly observations: blocked;
  - 24 to 35 observations: warning;
  - 36 to 59 observations: limited-history warning;
  - 60 or more observations: preferred.
- Excel workbook with live formulas, including `LINEST` regression formulas.
- Regression statistics:
  - coefficients;
  - standard errors;
  - t-statistics;
  - p-values;
  - R-squared;
  - adjusted R-squared;
  - observations;
  - F-statistic;
  - model p-value;
  - correlation matrix.
- Risk and performance statistics:
  - monthly and annualized arithmetic return;
  - monthly and annualized standard deviation;
  - Sharpe ratio;
  - Sortino ratio;
  - monthly and annualized downside deviation below zero;
  - monthly-resolution maximum drawdown;
  - annual 95% historical VaR based on rolling 12-month compounded returns;
  - annual 95% historical CVaR based on rolling 12-month compounded returns.
- Benchmark metrics:
  - return;
  - volatility;
  - Sharpe;
  - maximum drawdown;
  - correlation;
  - beta;
  - upside capture;
  - downside capture.
- Charts:
  - cumulative return;
  - drawdown;
  - actual vs. model-predicted monthly returns;
  - factor exposures;
  - rolling beta/exposure chart;
  - rolling R-squared chart.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

## Run the app

```bash
streamlit run app.py
```

## Uploaded return file format

The uploaded file should contain one date column and one monthly return column. Percent formatting is auto-detected. For example, both `2.5%` and `0.025` are treated as 2.5% monthly return; values like `2.5` are interpreted as percentage points and converted to `0.025`.

## Excel workbook structure

The generated workbook includes the following sheets:

- `Summary`
- `Fund Returns`
- `Factor Data`
- `Regression`
- `Conditional Regression`
- `Risk Statistics`
- `Diagnostics`
- `Rolling Analysis`
- `Definitions`

The regression sheets use Excel `LINEST` array formulas. The workbook is set to recalculate on open in Microsoft 365.

## Methodology notes

- Public security returns are calculated using Yahoo Finance adjusted close, which is intended to reflect dividends/distributions and splits. Raw price return is not used for total-return securities.
- VIX is used two ways:
  - monthly percentage change as the volatility factor;
  - month-end level for the crisis dummy.
- Annual arithmetic return is calculated as monthly arithmetic mean multiplied by 12.
- Annualized risk and downside deviation use square-root-of-12 scaling.
- Annual historical VaR/CVaR are calculated from rolling 12-month compounded return observations, not from parametric scaling of monthly VaR.
