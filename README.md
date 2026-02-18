# Cryptocurrency Portfolio Backtesting & Strategy Research (Python)

This project is a **research + backtesting pipeline** for building and evaluating systematic cryptocurrency portfolios using real market data. It pulls the **top-N cryptocurrencies by market cap** (CoinGecko), downloads **5 years of daily prices** (Yahoo Finance via `yfinance`), constructs **trend-following** and **time-series momentum** signals, turns them into **portfolio weights**, and runs a full **performance evaluation** including benchmarks, transaction costs, out-of-sample testing, parameter sensitivity, and macro co-movement analysis.

## What’s inside

### Step 1 — Universe + Data Download + Excel Export
- Fetches the **top N coins by market cap** using CoinGecko `coins/markets`.
- Maps CoinGecko symbols to **Yahoo Finance tickers** (e.g., `BTC-USD`), with a small exceptions map for edge cases.
- Downloads **daily auto-adjusted Close prices** for:
  - selected cryptos
  - benchmarks: **S&P 500 (`^GSPC`)** and **NASDAQ 100 (`^NDX`)**
- Saves everything into a single workbook: `crypto_data_5y.xlsx` with sheets:
  - `crypto_prices`, `benchmarks`, `crypto_meta` :contentReference[oaicite:1]{index=1}

### Step 2 — Cleaning + Returns + EDA
- Filters assets by **coverage** (default: keep coins with at least **80%** valid data).
- Fills remaining gaps using forward-fill/back-fill.
- Computes:
  - arithmetic returns (`pct_change`)
  - log returns (`log(Pt/Pt-1)`)
- Provides quick diagnostics + a correlation heatmap (example: top 15 cryptos). :contentReference[oaicite:2]{index=2}

### Step 3 — Trend-Following Signals
Two signal families are implemented:
- **Moving Average Crossover (MA)**: MA(short) vs MA(long), signal in `{−1, 0, +1}` (default: 50/200).
- **Time-Series Momentum (MOM)**: lookback return `Pt / Pt-L - 1`, signal in `{−1, 0, +1}` (default: 252 trading days). :contentReference[oaicite:3]{index=3}

The notebook also includes:
- % time each asset is “long”
- signal distribution histograms
- single-asset visual checks (price + scaled signal). :contentReference[oaicite:4]{index=4}

### Step 4 — Portfolio Construction (Weights)
Creates portfolios from signals using:
- **Long-only** weights (only `+1` positions included)
- **Long/short** weights (uses both `+1` and `−1`, normalized to `sum(|w|)=1`)
- Rebalancing frequency options: **daily / weekly / monthly**
- Weighting schemes:
  - **equal-weight**
  - **inverse-volatility-weight** using rolling volatility :contentReference[oaicite:5]{index=5}

### Step 5 — Backtesting & Performance
- Computes portfolio daily returns from aligned returns + weights.
- Builds equity curves via compounding.
- Compares multiple references:
  - **BTC buy & hold** (if present)
  - **equal-weight crypto universe**
  - **S&P 500** and **NASDAQ 100**
- Builds “hedged” versions (strategy return minus S&P 500 return).
- Computes key KPIs:
  - **CAGR**
  - **Annualized volatility**
  - **Sharpe ratio** (simple mean/vol, no RF subtraction in the current metric function)
  - **Maximum drawdown** :contentReference[oaicite:6]{index=6}

### Step 6 — Transaction Costs
- Estimates turnover: `0.5 * sum_i |w_i,t − w_i,t−1|`
- Applies proportional trading cost (default fee rate: **0.10%** per unit turnover).
- Compares gross vs net equity curves and metrics.
- Exports final weight panels to CSV:
  - `MA_weights_full.csv`, `MOM_weights_full.csv`, plus vol-weighted variants :contentReference[oaicite:7]{index=7}

### Step 7 — Out-of-Sample Evaluation
- Splits the history by a ratio rule (default: **60% train / 40% test**).
- Reports in-sample and out-of-sample metrics and equity curves for MA and MOM. :contentReference[oaicite:8]{index=8}

### Step 8 — Sensitivity Analysis
- MA grid (examples): (20,100), (50,200), (100,300) × rebalance {D,W,M}
- Momentum grid: lookbacks {90,180,252,360} × rebalance {D,W,M}
- Stores performance results in summary tables. :contentReference[oaicite:9]{index=9}

### Step 9 — Co-movement with Macro & Commodities
Downloads macro proxies (Yahoo Finance):
- Gold futures (`GC=F`), Oil futures (`CL=F`), VIX (`^VIX`)
Then:
- computes static correlations
- plots rolling correlations (60-day)
- runs OLS regressions of strategy returns on macro/market factors using `statsmodels`. :contentReference[oaicite:10]{index=10}

### Step 10 — Sentiment Filter (Fear & Greed)
- Downloads the **Crypto Fear & Greed Index** (alternative.me API)
- Normalizes to `[-1, 1]` and builds a basic filter (default: sentiment ≥ 0).
- Applies it to MA long signals (gating longs when sentiment is weak).
- Compares performance of **original MA** vs **sentiment-filtered MA**. :contentReference[oaicite:11]{index=11}

---

## Requirements

- Python 3.9+ recommended
- Core packages:
  - `pandas`, `numpy`
  - `matplotlib`, `seaborn`
  - `yfinance`
  - `requests`
  - `statsmodels`

Install:
```bash
pip install pandas numpy matplotlib seaborn yfinance requests statsmodels
