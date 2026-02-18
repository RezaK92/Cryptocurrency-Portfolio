# ============================
# STEP 1: DATA DOWNLOAD + EXCEL
# ============================

from dataclasses import dataclass
from typing import List, Optional, Tuple
from datetime import datetime, timedelta

import logging
import requests
import pandas as pd
import yfinance as yf

logging.getLogger("yfinance").setLevel(logging.CRITICAL)

# ---------- 1. ASSET CONFIGURATION ----------

@dataclass
class AssetConfig:
    """
    Configuration for each asset (crypto or benchmark).
    """
    name: str
    ticker: str          # For yfinance (e.g. 'BTC-USD', '^GSPC')
    asset_class: str     # 'crypto', 'index', etc.
    coingecko_id: Optional[str] = None  # only used for cryptos


COINGECKO_MARKETS_URL = "https://api.coingecko.com/api/v3/coins/markets"

# mapping for odd symbols if needed (kept but rarely used)
TICKER_EXCEPTIONS = {
    "BSC-USD": "BSC-USD",
    "FIGR_HELOC": "FIGR",
    "SUSDS": "SUSD-USD",
}


# Purpose:
#   Pull the current top-N cryptocurrencies (by market cap) from CoinGecko and translate them into a
#   Yahoo Finance-compatible universe.
# Inputs:
#   - n: number of coins to request (must be > 0)
#   - vs_currency: pricing currency used by CoinGecko (default: "usd")
# Outputs:
#   - assets: List[AssetConfig] with Yahoo tickers (e.g., "BTC-USD")
#   - meta_df: DataFrame with rank/symbol/name/yahoo_ticker for reporting and auditing
# Notes:
#   Applies TICKER_EXCEPTIONS when CoinGecko symbols do not map cleanly to Yahoo tickers.
def get_top_n_crypto_assets_from_coingecko(
    n: int = 50,
    vs_currency: str = "usd"
) -> Tuple[List[AssetConfig], pd.DataFrame]:
    """
    Query CoinGecko for the top N coins by market cap and build:
      - a list of AssetConfig objects for yfinance
      - a meta DataFrame with rank, symbol, name, yahoo_ticker
    """
    if n <= 0:
        raise ValueError("n must be positive")

    per_page = min(n, 250)

    params = {
        "vs_currency": vs_currency,
        "order": "market_cap_desc",
        "per_page": per_page,
        "page": 1,
        "sparkline": "false",
    }

    print(f"Fetching top {n} cryptos from CoinGecko...")
    response = requests.get(COINGECKO_MARKETS_URL, params=params, timeout=30)
    response.raise_for_status()
    data = response.json()
    data = data[:n]

    assets: List[AssetConfig] = []
    meta_records = []

    for coin in data:
        symbol = (coin.get("symbol") or "").upper()
        name = coin.get("name", "")
        cg_id = coin.get("id", "")
        rank = coin.get("market_cap_rank", None)

        if symbol in TICKER_EXCEPTIONS:
            yahoo_ticker = TICKER_EXCEPTIONS[symbol]
        else:
            yahoo_ticker = f"{symbol}-USD"

        assets.append(
            AssetConfig(
                name=name,
                ticker=yahoo_ticker,
                asset_class="crypto",
                coingecko_id=cg_id,
            )
        )

        meta_records.append(
            {
                "rank": rank,
                "symbol": symbol,
                "name": name,
                "yahoo_ticker": yahoo_ticker,
            }
        )

    meta_df = pd.DataFrame(meta_records).sort_values("rank").reset_index(drop=True)
    return assets, meta_df


BENCHMARK_ASSETS: List[AssetConfig] = [
    AssetConfig("S&P 500", "^GSPC", "index"),
    AssetConfig("NASDAQ 100", "^NDX", "index"),
]

TODAY = datetime.today().date()
START_DATE = TODAY - timedelta(days=5 * 365)
END_DATE = TODAY
PRICE_INTERVAL = "1d"


# Purpose:
#   Download daily (or chosen interval) adjusted close price history from Yahoo Finance via yfinance.
# Inputs:
#   - assets: list of AssetConfig objects (tickers are taken from AssetConfig.ticker)
#   - start/end/interval: standard yfinance download parameters
# Output:
#   - prices: DataFrame indexed by date with one column per ticker (Close prices, auto-adjusted)
# Notes:
#   Handles yfinance's different return shapes for single-ticker vs multi-ticker downloads.
def download_price_history_yahoo(
    assets: List[AssetConfig],
    start: datetime.date = START_DATE,
    end: datetime.date = END_DATE,
    interval: str = PRICE_INTERVAL,
) -> pd.DataFrame:
    """
    Download adjusted close prices for a list of assets from Yahoo Finance.
    """
    tickers = [a.ticker for a in assets]

    print(f"Downloading data for {len(tickers)} assets from Yahoo Finance...")
    data = yf.download(
        tickers=tickers,
        start=start,
        end=end,
        interval=interval,
        auto_adjust=True,
        progress=True,
        group_by="ticker",
    )

    if len(tickers) == 1:
        prices = data["Close"].to_frame()
        prices.columns = tickers
    else:
        try:
            prices = data.xs("Close", level=1, axis=1)
        except KeyError:
            if "Close" in data:
                prices = data["Close"]
            else:
                prices = pd.DataFrame()
                for t in tickers:
                    if t in data.columns.levels[0]:
                        prices[t] = data[t]["Close"]

    prices = prices.sort_index().dropna(how="all")
    return prices


# Purpose:
#   Persist Step 1 outputs into a single Excel workbook for reproducibility and separation of concerns.
# Inputs:
#   - crypto_prices: price panel for cryptos (columns = tickers)
#   - benchmark_prices: price panel for benchmark indices
#   - crypto_meta: metadata table from CoinGecko (rank/symbol/name/yahoo ticker)
# Output:
#   - Writes an Excel file with three sheets: crypto_prices, benchmarks, crypto_meta
def save_data_to_excel(
    crypto_prices: pd.DataFrame,
    benchmark_prices: pd.DataFrame,
    crypto_meta: pd.DataFrame,
    filename: str = "crypto_data_5y.xlsx",
) -> None:
    """
    Save cryptos, benchmarks, and metadata into a single Excel file.
    """
    print(f"Sorting data alphabetically and saving to {filename}...")

    crypto_prices = crypto_prices.sort_index(axis=1)
    benchmark_prices = benchmark_prices.sort_index(axis=1)
    crypto_meta = crypto_meta.sort_values(by="symbol")

    with pd.ExcelWriter(filename) as writer:
        crypto_prices.to_excel(writer, sheet_name="crypto_prices")
        benchmark_prices.to_excel(writer, sheet_name="benchmarks")
        crypto_meta.to_excel(writer, sheet_name="crypto_meta", index=False)
    print("Done.")


if __name__ == "__main__":
    try:
        crypto_assets, crypto_meta = get_top_n_crypto_assets_from_coingecko(n=50)
        print("\n--- Top 5 Cryptos (Preview) ---")
        print(crypto_meta[["rank", "name", "yahoo_ticker"]].head())

        print("\n--- Downloading Crypto Prices ---")
        crypto_prices = download_price_history_yahoo(crypto_assets)

        print("\n--- Downloading Benchmark Prices ---")
        benchmark_prices = download_price_history_yahoo(BENCHMARK_ASSETS)

        print(f"\nCrypto prices shape: {crypto_prices.shape}")
        print(f"Benchmark prices shape: {benchmark_prices.shape}")

        save_data_to_excel(
            crypto_prices=crypto_prices,
            benchmark_prices=benchmark_prices,
            crypto_meta=crypto_meta,
            filename="crypto_data_5y.xlsx",
        )

    except Exception as e:
        print(f"\n[ERROR] An error occurred in STEP 1: {e}")


# ================================
# STEP 2: CLEANING + RETURNS + EDA
# ================================

import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns


# Purpose:
#   Load the Excel workbook produced in Step 1 and restore indexes/dtypes needed for time-series work.
# Output:
#   - crypto_prices, benchmark_prices: DataFrames with a DatetimeIndex
#   - crypto_meta: metadata DataFrame
def load_step1_data(filename="crypto_data_5y.xlsx"):
    crypto_prices = pd.read_excel(filename, sheet_name="crypto_prices", index_col=0)
    benchmark_prices = pd.read_excel(filename, sheet_name="benchmarks", index_col=0)
    crypto_meta = pd.read_excel(filename, sheet_name="crypto_meta")

    crypto_prices.index = pd.to_datetime(crypto_prices.index)
    benchmark_prices.index = pd.to_datetime(benchmark_prices.index)
    return crypto_prices, benchmark_prices, crypto_meta


crypto_prices, benchmark_prices, crypto_meta = load_step1_data()


# Purpose:
#   Basic cleaning for price panels:
#   - Drop assets with too many missing observations (coverage filter)
#   - Fill remaining gaps using forward-fill and back-fill
# Inputs:
#   - df: price DataFrame (dates x assets)
#   - min_valid_ratio: minimum fraction of non-missing observations required to keep an asset
# Outputs:
#   - df_clean: cleaned DataFrame
#   - keep_cols: Index of retained asset columns
def clean_price_data(df, min_valid_ratio=0.80):
    valid_ratio = df.notna().mean()
    keep_cols = valid_ratio[valid_ratio >= min_valid_ratio].index
    df_clean = df[keep_cols].copy()
    df_clean = df_clean.ffill().bfill()
    return df_clean, keep_cols


crypto_prices_clean, kept_cryptos = clean_price_data(crypto_prices, min_valid_ratio=0.80)
benchmark_prices_clean, kept_bench = clean_price_data(benchmark_prices, min_valid_ratio=0.80)

print(f"Kept {len(kept_cryptos)} out of {crypto_prices.shape[1]} cryptos after cleaning.")
print(f"Kept benchmarks: {list(kept_bench)}")


# Purpose:
#   Compute both arithmetic (pct_change) and log returns from a price DataFrame.
# Output:
#   - arithmetic: simple returns (r_t = P_t/P_{t-1} - 1)
#   - logret: log returns (log(P_t/P_{t-1}))
def compute_returns(df):
    arithmetic = df.pct_change().dropna()
    logret = np.log(df / df.shift(1)).dropna()
    return arithmetic, logret


crypto_ret, crypto_logret = compute_returns(crypto_prices_clean)
bench_ret, bench_logret = compute_returns(benchmark_prices_clean)


# Purpose:
#   Generate a compact diagnostic summary for a panel (dates, size, missingness).
# Output:
#   - Single-row DataFrame with start/end dates, number of observations, number of assets, and total NA count.
def data_summary(df):
    desc = pd.DataFrame(
        {
            "start_date": df.index.min(),
            "end_date": df.index.max(),
            "n_obs": df.shape[0],
            "n_assets": df.shape[1],
            "missing_values": df.isna().sum().sum(),
        },
        index=[0],
    )
    return desc


print("\nCRYPTO DATA SUMMARY:")
print(data_summary(crypto_prices_clean))

print("\nBENCHMARK DATA SUMMARY:")
print(data_summary(benchmark_prices_clean))


# Purpose:
#   Visual EDA: correlation heatmap of a return panel (useful for dependence structure sanity checks).
# Notes:
#   For large universes, call this on a subset of columns to keep the plot interpretable.
def plot_correlation_matrix(returns, title="Correlation Matrix"):
    corr = returns.corr()
    plt.figure(figsize=(12, 10))
    sns.heatmap(corr, cmap="coolwarm", annot=False)
    plt.title(title)
    plt.show()


plot_correlation_matrix(crypto_ret.iloc[:, :15], "Crypto Correlation Matrix (Top 15)")


# ============================
# STEP 3: TREND-FOLLOWING SIGNALS
# ============================

# Purpose:
#   Build moving-average crossover trend signals for each asset:
#   - +1 when short MA > long MA
#   - -1 when short MA < long MA
#   -  0 otherwise / insufficient history
# Outputs:
#   - MA_short, MA_long: rolling mean series
#   - signal: DataFrame of {-1, 0, +1}
def compute_MA_signals(prices, short=50, long=200):
    """
    Moving average crossover signals in {-1, 0, +1}.
    """
    MA_short = prices.rolling(short).mean()
    MA_long = prices.rolling(long).mean()

    signal = pd.DataFrame(index=prices.index, columns=prices.columns, dtype=float)
    signal[MA_short > MA_long] = 1.0
    signal[MA_short < MA_long] = -1.0
    signal = signal.fillna(0.0).astype(float)
    return MA_short, MA_long, signal


MA50, MA200, MA_signal = compute_MA_signals(crypto_prices_clean, short=50, long=200)


# Purpose:
#   Build time-series momentum signals based on a lookback return:
#   - momentum = P_t / P_{t-lookback} - 1
#   - +1 if momentum > 0, -1 if momentum < 0, else 0
# Outputs:
#   - momentum: momentum values
#   - signal: DataFrame of {-1, 0, +1}
def compute_momentum_signal(prices, lookback=252):
    """
    Time-series momentum: price_t / price_{t-lookback} - 1.
    Signal in {-1, 0, +1}.
    """
    momentum = prices / prices.shift(lookback) - 1.0
    signal = pd.DataFrame(index=prices.index, columns=prices.columns, dtype=float)

    signal[momentum > 0] = 1.0
    signal[momentum < 0] = -1.0
    signal = signal.fillna(0.0).astype(float)
    return momentum, signal


momentum, MOM_signal = compute_momentum_signal(crypto_prices_clean, lookback=252)


# Purpose:
#   Summarize how often each asset is in a long regime (signal == +1), expressed as a percentage of time.
def summarize_signals(signal_df):
    """
    % time in long position (signal = +1).
    """
    return (signal_df == 1.0).mean() * 100


MA_long_percentage = summarize_signals(MA_signal)
MOM_long_percentage = summarize_signals(MOM_signal)

print("\n=== % Time in Long Position (MA50/200) ===")
print(MA_long_percentage.round(2))

print("\n=== % Time in Long Position (Momentum 12M) ===")
print(MOM_long_percentage.round(2))


# Purpose:
#   Plot the cross-sectional distribution of "% time long" across assets for a given signal matrix.
def plot_signal_distribution(signal_df, title):
    sns.histplot((signal_df == 1).mean() * 100, bins=10)
    plt.title(title)
    plt.xlabel("% Time in Long Position")
    plt.ylabel("Count of Cryptos")
    plt.show()


plot_signal_distribution(MA_signal, "Distribution of MA Trend Signals Across Assets")
plot_signal_distribution(MOM_signal, "Distribution of Momentum Signals Across Assets")


# Purpose:
#   Plot a single asset's price history together with its (scaled) signal path for visual inspection.
# Notes:
#   Signal is scaled for display only; it is not used for any computation.
def plot_single_signal(asset, prices, signal):
    plt.figure(figsize=(12, 5))
    plt.plot(prices[asset], label=f"{asset} Price", alpha=0.7)
    plt.plot(signal[asset] * prices[asset].max() * 0.1,
             label="Signal (scaled)", color="red", alpha=0.6)
    plt.title(f"Trend Signal for {asset}")
    plt.legend()
    plt.show()


example_asset = crypto_prices_clean.columns[0]
plot_single_signal(example_asset, crypto_prices_clean, MA_signal)
plot_single_signal(example_asset, crypto_prices_clean, MOM_signal)


# ================================
# STEP 4: PORTFOLIO WEIGHTS
# ================================

# Purpose:
#   Produce a sequence of rebalancing dates from a DatetimeIndex.
# Inputs:
#   - freq: 'D' (daily), 'W' (weekly; last trading day of week), 'M' (month-end)
# Output:
#   - DatetimeIndex of rebalance points used by the portfolio construction functions.
def get_rebalance_dates(index, freq="M"):
    """
    Get rebalancing dates from a DatetimeIndex.
    freq: 'D', 'W', 'M'
    """
    if freq == "D":
        return index

    s = index.to_series()
    if freq == "W":
        return s.resample("W-FRI").last().dropna().index
    elif freq == "M":
        return s.resample("ME").last().dropna().index
    else:
        raise ValueError("freq must be 'D', 'W', or 'M'")


# Purpose:
#   Rolling volatility estimator: rolling standard deviation of returns over a lookback window.
def rolling_volatility(returns, window=30):
    return returns.rolling(window).std()


# Purpose:
#   Construct long-only portfolio weights from signals and apply them between rebalance dates.
# Process:
#   1) Convert signals to long-only indicator (signal > 0)
#   2) On each rebalance date, compute weights using:
#      - 'equal': equal weight across active long positions
#      - 'volatility': inverse-vol weighting using rolling volatility
#   3) Forward-fill weights until the next rebalance date
# Output:
#   - weights: DataFrame (dates x assets) with weights that sum to 1 when positions exist, else 0
def build_portfolio_weights(
    returns: pd.DataFrame,
    signals: pd.DataFrame,
    freq: str = "M",
    weight_scheme: str = "equal",
    vol_window: int = 30,
) -> pd.DataFrame:
    """
    Long-only weights from signals in {0,1} or {-1,0,+1} (we use long side only).
    """
    signals = signals.loc[returns.index].fillna(0.0)

    # Convert any -1 to 0 for long-only versions
    long_signals = (signals > 0).astype(float)

    rebal_dates = get_rebalance_dates(returns.index, freq=freq)
    weights = pd.DataFrame(0.0, index=returns.index, columns=returns.columns)

    if weight_scheme == "volatility":
        vol = rolling_volatility(returns, window=vol_window)
    else:
        vol = None

    for i, date in enumerate(rebal_dates):
        if date not in returns.index:
            continue

        sig = long_signals.loc[date].copy()
        active = sig > 0
        if active.sum() == 0:
            w = pd.Series(0.0, index=returns.columns)
        else:
            if weight_scheme == "equal":
                raw = sig[active].astype(float)
                total = raw.sum()
                w = pd.Series(0.0, index=returns.columns)
                w[active] = raw / total
            elif weight_scheme == "volatility":
                current_vol = vol.loc[date]
                current_vol = current_vol.replace(0, np.nan)

                raw = pd.Series(0.0, index=returns.columns)
                raw[active] = 1.0 / current_vol[active]
                raw = raw.replace([np.inf, -np.inf], np.nan).fillna(0.0)

                total = raw.sum()
                if total == 0:
                    w = pd.Series(0.0, index=returns.columns)
                else:
                    w = raw / total
            else:
                raise ValueError("weight_scheme must be 'equal' or 'volatility'")

        start_date = date
        if i < len(rebal_dates) - 1:
            end_date = rebal_dates[i + 1]
        else:
            end_date = returns.index[-1] + pd.Timedelta(days=1)

        mask = (returns.index >= start_date) & (returns.index < end_date)
        weights.loc[mask] = w.values

    return weights


# Purpose:
#   Construct long/short weights from {-1,0,+1} signals and apply them between rebalance dates.
# Normalization:
#   On each rebalance date, normalize so gross exposure equals 1: sum(|w_i|) = 1.
# Output:
#   - weights: DataFrame (dates x assets) with both positive and negative weights
def build_long_short_portfolio_weights(
    returns: pd.DataFrame,
    signals: pd.DataFrame,
    freq: str = "M",
    weight_scheme: str = "equal",
    vol_window: int = 30,
) -> pd.DataFrame:
    """
    LONG/SHORT weights from signals in {-1,0,+1}.
    Normalized so sum(|w|) = 1 on rebal days.
    """
    signals = signals.loc[returns.index].fillna(0.0)
    rebal_dates = get_rebalance_dates(returns.index, freq=freq)
    weights = pd.DataFrame(0.0, index=returns.index, columns=returns.columns)

    if weight_scheme == "volatility":
        vol = rolling_volatility(returns, window=vol_window)
    else:
        vol = None

    for i, date in enumerate(rebal_dates):
        if date not in returns.index:
            continue

        sig = signals.loc[date].copy()
        active = sig != 0
        if active.sum() == 0:
            w = pd.Series(0.0, index=returns.columns)
        else:
            if weight_scheme == "equal":
                raw = sig[active].astype(float)
                gross = np.abs(raw).sum()
                w = pd.Series(0.0, index=returns.columns)
                w[active] = raw / gross
            elif weight_scheme == "volatility":
                current_vol = vol.loc[date]
                current_vol = current_vol.replace(0, np.nan)

                raw = pd.Series(0.0, index=returns.columns)
                raw[active] = np.sign(sig[active]) * (1.0 / current_vol[active])
                raw = raw.replace([np.inf, -np.inf], np.nan).fillna(0.0)

                gross = np.abs(raw).sum()
                if gross == 0:
                    w = pd.Series(0.0, index=returns.columns)
                else:
                    w = raw / gross
            else:
                raise ValueError("weight_scheme must be 'equal' or 'volatility'")

        start_date = date
        if i < len(rebal_dates) - 1:
            end_date = rebal_dates[i + 1]
        else:
            end_date = returns.index[-1] + pd.Timedelta(days=1)

        mask = (returns.index >= start_date) & (returns.index < end_date)
        weights.loc[mask] = w.values

    return weights


rebalance_freq = "M"
weight_scheme = "equal"

MA_weights = build_portfolio_weights(
    returns=crypto_ret,
    signals=MA_signal,
    freq=rebalance_freq,
    weight_scheme=weight_scheme,
)

MOM_weights = build_portfolio_weights(
    returns=crypto_ret,
    signals=MOM_signal,
    freq=rebalance_freq,
    weight_scheme=weight_scheme,
)

MA_weights_vol = build_portfolio_weights(
    returns=crypto_ret,
    signals=MA_signal,
    freq=rebalance_freq,
    weight_scheme="volatility",
)

MOM_weights_vol = build_portfolio_weights(
    returns=crypto_ret,
    signals=MOM_signal,
    freq=rebalance_freq,
    weight_scheme="volatility",
)

MA_LS_weights = build_long_short_portfolio_weights(
    returns=crypto_ret,
    signals=MA_signal,
    freq=rebalance_freq,
    weight_scheme="equal",
)

MOM_LS_weights = build_long_short_portfolio_weights(
    returns=crypto_ret,
    signals=MOM_signal,
    freq=rebalance_freq,
    weight_scheme="equal",
)

MA_LS_weights_vol = build_long_short_portfolio_weights(
    returns=crypto_ret,
    signals=MA_signal,
    freq=rebalance_freq,
    weight_scheme="volatility",
)

MOM_LS_weights_vol = build_long_short_portfolio_weights(
    returns=crypto_ret,
    signals=MOM_signal,
    freq=rebalance_freq,
    weight_scheme="volatility",
)

ma_num_assets = (MA_weights > 0).sum(axis=1)
mom_num_assets = (MOM_weights > 0).sum(axis=1)

print("MA_weights_vol shape:", MA_weights_vol.shape)
print("MOM_weights_vol shape:", MOM_weights_vol.shape)

plt.figure(figsize=(12, 5))
plt.plot(ma_num_assets, label="MA strategy")
plt.plot(mom_num_assets, label="Momentum strategy")
plt.title("Number of Assets Held Each Day (Long-only)")
plt.xlabel("Date")
plt.ylabel("Count of Assets in Portfolio")
plt.legend()
plt.grid(True)
plt.show()

print("MA_weights shape:", MA_weights.shape)
print("MOM_weights shape:", MOM_weights.shape)

print("\nCheck daily sum of MA weights (first 10 days):")
print(MA_weights.sum(axis=1).head(10))


# ================================
# STEP 5: BACKTESTING & PERFORMANCE
# ================================

# Purpose:
#   Compute a portfolio return series given an asset return panel and a weight panel.
# Notes:
#   Aligns both index (dates) and columns (assets) to avoid silent misalignment errors.
def compute_portfolio_returns(returns: pd.DataFrame, weights: pd.DataFrame) -> pd.Series:
    returns, weights = returns.align(weights, join="inner", axis=0)
    returns, weights = returns.align(weights, join="inner", axis=1)
    port_ret = (weights * returns).sum(axis=1)
    port_ret.name = "portfolio_return"
    return port_ret


MA_port_ret = compute_portfolio_returns(crypto_ret, MA_weights)
MOM_port_ret = compute_portfolio_returns(crypto_ret, MOM_weights)

MA_LS_port_ret = (MA_LS_weights * crypto_ret).sum(axis=1)
MOM_LS_port_ret = (MOM_LS_weights * crypto_ret).sum(axis=1)

MA_LS_port_ret_vol = (MA_LS_weights_vol * crypto_ret).sum(axis=1)
MOM_LS_port_ret_vol = (MOM_LS_weights_vol * crypto_ret).sum(axis=1)

MA_port_ret_vol = compute_portfolio_returns(crypto_ret, MA_weights_vol)
MOM_port_ret_vol = compute_portfolio_returns(crypto_ret, MOM_weights_vol)

print("MA_port_ret shape:", MA_port_ret.shape)
print("MOM_port_ret shape:", MOM_port_ret.shape)
print("MA_port_ret_vol shape:", MA_port_ret_vol.shape)
print("MOM_port_ret_vol shape:", MOM_port_ret_vol.shape)


# Purpose:
#   Convert a return series into a cumulative equity curve via compounded growth.
# Output:
#   - Series representing portfolio value starting from initial_capital.
def equity_curve(returns: pd.Series, initial_capital: float = 1.0) -> pd.Series:
    return initial_capital * (1 + returns.dropna()).cumprod()


# --- Equity curves for all variants ---

MA_equity = equity_curve(MA_port_ret)
MOM_equity = equity_curve(MOM_port_ret)

MA_equity_vol = equity_curve(MA_port_ret_vol)
MOM_equity_vol = equity_curve(MOM_port_ret_vol)

MA_LS_equity = equity_curve(MA_LS_port_ret)
MOM_LS_equity = equity_curve(MOM_LS_port_ret)

MA_LS_equity_vol = equity_curve(MA_LS_port_ret_vol)
MOM_LS_equity_vol = equity_curve(MOM_LS_port_ret_vol)


# --- Benchmarks & BTC, equal-weight universe ---

btc_ticker_candidates = [c for c in crypto_ret.columns if c in ["BTC-USD", "XBT-USD", "BTCUSD=X"]]
if len(btc_ticker_candidates) > 0:
    btc_ticker = btc_ticker_candidates[0]
    BTC_buyhold_ret = crypto_ret[btc_ticker].dropna()
    BTC_equity = equity_curve(BTC_buyhold_ret)
else:
    BTC_buyhold_ret = None
    BTC_equity = None
    print("Warning: BTC not found in crypto_ret columns.")

equal_weights = pd.DataFrame(
    1.0 / crypto_ret.shape[1],
    index=crypto_ret.index,
    columns=crypto_ret.columns,
)
EW_crypto_ret = compute_portfolio_returns(crypto_ret, equal_weights)
EW_crypto_equity = equity_curve(EW_crypto_ret)

bench_ret_aligned = bench_ret.dropna()

if "^GSPC" in bench_ret_aligned.columns:
    SP500_ret = bench_ret_aligned["^GSPC"]
    SP500_equity = equity_curve(SP500_ret)
else:
    SP500_ret = None
    SP500_equity = None
    print("Warning: ^GSPC not found in benchmark returns.")

if "^NDX" in bench_ret_aligned.columns:
    NDX_ret = bench_ret_aligned["^NDX"]
    NDX_equity = equity_curve(NDX_ret)
else:
    NDX_ret = None
    NDX_equity = None
    print("Warning: ^NDX not found in benchmark returns.")


# --- Hedged versions (long-only MA/MOM vs SP500) ---

if SP500_ret is not None:
    sp500_aligned = SP500_ret.reindex(MA_port_ret.index).fillna(0.0)

    MA_hedged_ret = MA_port_ret - sp500_aligned
    MOM_hedged_ret = MOM_port_ret - sp500_aligned

    MA_hedged_ret.name = "MA_hedged_ret"
    MOM_hedged_ret.name = "MOM_hedged_ret"

    MA_hedged_equity = equity_curve(MA_hedged_ret)
    MOM_hedged_equity = equity_curve(MOM_hedged_ret)
else:
    MA_hedged_ret = pd.Series(dtype=float)
    MOM_hedged_ret = pd.Series(dtype=float)
    MA_hedged_equity = None
    MOM_hedged_equity = None


# --- Build equity_df with all curves (long-only, long/short, vol, benchmarks, hedged) ---

equity_df = pd.DataFrame(index=crypto_ret.index)

equity_df["MA_strategy"] = MA_equity
equity_df["MOM_strategy"] = MOM_equity
equity_df["EW_crypto"] = EW_crypto_equity

equity_df["MA_LS"] = MA_LS_equity.reindex(equity_df.index)
equity_df["MOM_LS"] = MOM_LS_equity.reindex(equity_df.index)
equity_df["MA_LS_vol"] = MA_LS_equity_vol.reindex(equity_df.index)
equity_df["MOM_LS_vol"] = MOM_LS_equity_vol.reindex(equity_df.index)

if BTC_equity is not None:
    equity_df["BTC_buyhold"] = BTC_equity.reindex(equity_df.index)

if SP500_equity is not None:
    equity_df["SP500"] = SP500_equity.reindex(equity_df.index)

if NDX_equity is not None:
    equity_df["NASDAQ100"] = NDX_equity.reindex(equity_df.index)

equity_df["MA_vol"] = MA_equity_vol.reindex(equity_df.index)
equity_df["MOM_vol"] = MOM_equity_vol.reindex(equity_df.index)

if MA_hedged_equity is not None:
    equity_df["MA_hedged"] = MA_hedged_equity.reindex(equity_df.index)
if MOM_hedged_equity is not None:
    equity_df["MOM_hedged"] = MOM_hedged_equity.reindex(equity_df.index)

equity_df = equity_df.dropna(how="all")
print("\nEquity DataFrame columns:", equity_df.columns)


# --- Global normalized plot ---

plt.figure(figsize=(12, 6))
for col in equity_df.columns:
    series = equity_df[col].dropna()
    plt.plot(series.index, series / series.iloc[0], label=col)
plt.legend()
plt.title("Equity Curves (normalized to 1)")
plt.ylabel("Cumulative Value (normalized)")
plt.xlabel("Date")
plt.grid(True)
plt.show()


# --- Dedicated long-only vs long/short plots ---

plt.figure(figsize=(12, 6))
plt.plot(MA_equity, label="MA long-only")
plt.plot(MA_LS_equity, label="MA long/short")
plt.plot(MA_LS_equity_vol, label="MA long/short (vol-weighted)", linestyle="--")
plt.title("MA Strategy: Long-only vs Long/Short")
plt.xlabel("Date")
plt.ylabel("Equity")
plt.legend()
plt.grid(True)
plt.show()

plt.figure(figsize=(12, 6))
plt.plot(MOM_equity, label="MOM long-only")
plt.plot(MOM_LS_equity, label="MOM long/short")
plt.plot(MOM_LS_equity_vol, label="MOM long/short (vol-weighted)", linestyle="--")
plt.title("Momentum Strategy: Long-only vs Long/Short")
plt.xlabel("Date")
plt.ylabel("Equity")
plt.legend()
plt.grid(True)
plt.show()


# --- Performance metrics ---

# Purpose:
#   Compute standard backtest KPIs for a return series:
#   - CAGR, annualized volatility, Sharpe ratio, maximum drawdown
# Notes:
#   freq defaults to 252 (daily trading days); adjust if using a different sampling frequency.
def performance_metrics(returns: pd.Series, freq: int = 252) -> pd.Series:
    returns = returns.dropna()
    if returns.empty:
        return pd.Series(dtype=float)

    cum_growth = (1 + returns).prod()
    n_periods = len(returns)
    years = n_periods / freq
    cagr = cum_growth ** (1 / years) - 1 if years > 0 else np.nan

    vol = returns.std() * np.sqrt(freq)
    sharpe = (returns.mean() * freq) / vol if vol != 0 else np.nan

    cum_curve = (1 + returns).cumprod()
    running_max = cum_curve.cummax()
    drawdown = cum_curve / running_max - 1
    max_dd = drawdown.min()

    return pd.Series(
        {
            "CAGR": cagr,
            "AnnVol": vol,
            "Sharpe": sharpe,
            "MaxDrawdown": max_dd,
        }
    )


metrics_dict = {}
metrics_dict["MA_strategy"] = performance_metrics(MA_port_ret)
metrics_dict["MOM_strategy"] = performance_metrics(MOM_port_ret)
metrics_dict["EW_crypto"] = performance_metrics(EW_crypto_ret)

if BTC_buyhold_ret is not None:
    metrics_dict["BTC_buyhold"] = performance_metrics(BTC_buyhold_ret)
if SP500_ret is not None:
    metrics_dict["SP500"] = performance_metrics(SP500_ret)
if NDX_ret is not None:
    metrics_dict["NASDAQ100"] = performance_metrics(NDX_ret)

metrics_dict["MA_vol"] = performance_metrics(MA_port_ret_vol)
metrics_dict["MOM_vol"] = performance_metrics(MOM_port_ret_vol)

metrics_dict["MA_LS"] = performance_metrics(MA_LS_port_ret)
metrics_dict["MOM_LS"] = performance_metrics(MOM_LS_port_ret)
metrics_dict["MA_LS_vol"] = performance_metrics(MA_LS_port_ret_vol)
metrics_dict["MOM_LS_vol"] = performance_metrics(MOM_LS_port_ret_vol)

if not MA_hedged_ret.empty:
    metrics_dict["MA_hedged"] = performance_metrics(MA_hedged_ret)
if not MOM_hedged_ret.empty:
    metrics_dict["MOM_hedged"] = performance_metrics(MOM_hedged_ret)

perf_table = pd.DataFrame(metrics_dict).T
print("\n=== Performance Summary (no transaction costs yet) ===")
print(perf_table.round(4))


# ================================
# STEP 6: TRANSACTION COSTS
# ================================

# Purpose:
#   Approximate trading turnover from weight changes:
#   - turnover_t = 0.5 * sum_i |w_{i,t} - w_{i,t-1}|
# Interpretation:
#   The 0.5 factor avoids double-counting buy/sell legs in a fully-invested portfolio.
def compute_daily_turnover(weights: pd.DataFrame) -> pd.Series:
    weights = weights.sort_index()
    dw = weights.diff().abs()
    turnover = 0.5 * dw.sum(axis=1)
    turnover.name = "turnover"
    return turnover


MA_turnover = compute_daily_turnover(MA_weights)
MOM_turnover = compute_daily_turnover(MOM_weights)
MA_turnover_vol = compute_daily_turnover(MA_weights_vol)
MOM_turnover_vol = compute_daily_turnover(MOM_weights_vol)

print("MA_turnover (first 10 rows):")
print(MA_turnover.head(10))
print("\nMOM_turnover (first 10 rows):")
print(MOM_turnover.head(10))


# Purpose:
#   Convert gross returns into net returns by subtracting proportional trading costs:
#   - net_t = gross_t - fee_rate * turnover_t
def apply_transaction_costs(
    gross_returns: pd.Series,
    turnover: pd.Series,
    fee_rate: float = 0.001,
) -> pd.Series:
    gross_returns, turnover = gross_returns.align(turnover, join="inner")
    net_returns = gross_returns - fee_rate * turnover
    net_returns.name = f"{gross_returns.name}_net"
    return net_returns


FEE_RATE = 0.001

MA_port_ret_net = apply_transaction_costs(MA_port_ret, MA_turnover, fee_rate=FEE_RATE)
MOM_port_ret_net = apply_transaction_costs(MOM_port_ret, MOM_turnover, fee_rate=FEE_RATE)

MA_port_ret_vol_net = apply_transaction_costs(MA_port_ret_vol, MA_turnover_vol, fee_rate=FEE_RATE)
MOM_port_ret_vol_net = apply_transaction_costs(MOM_port_ret_vol, MOM_turnover_vol, fee_rate=FEE_RATE)

MA_equity_net = equity_curve(MA_port_ret_net)
MOM_equity_net = equity_curve(MOM_port_ret_net)

plt.figure(figsize=(12, 6))
plt.plot(MA_equity, label="MA gross")
plt.plot(MA_equity_net, label="MA net (with costs)")
plt.title("MA Strategy: Gross vs Net Equity")
plt.legend()
plt.grid(True)
plt.show()

plt.figure(figsize=(12, 6))
plt.plot(MOM_equity, label="MOM gross")
plt.plot(MOM_equity_net, label="MOM net (with costs)")
plt.title("Momentum Strategy: Gross vs Net Equity")
plt.legend()
plt.grid(True)
plt.show()

metrics_with_costs = {
    "MA_gross": performance_metrics(MA_port_ret),
    "MA_net": performance_metrics(MA_port_ret_net),
    "MA_vol_gross": performance_metrics(MA_port_ret_vol),
    "MA_vol_net": performance_metrics(MA_port_ret_vol_net),
    "MOM_gross": performance_metrics(MOM_port_ret),
    "MOM_net": performance_metrics(MOM_port_ret_net),
    "MOM_vol_gross": performance_metrics(MOM_port_ret_vol),
    "MOM_vol_net": performance_metrics(MOM_port_ret_vol_net),
}

perf_costs = pd.DataFrame(metrics_with_costs).T
print("\n=== Performance With and Without Transaction Costs ===")
print(perf_costs.round(4))

print("\n=== SHOWING LAST AVAILABLE PORTFOLIO WEIGHTS ===")
print("\n--- MA Equal-Weight (Last Day) ---")
print(MA_weights.iloc[-1])
print("\n--- MOM Equal-Weight (Last Day) ---")
print(MOM_weights.iloc[-1])
print("\n--- MA Volatility-Weighted (Last Day) ---")
print(MA_weights_vol.iloc[-1])
print("\n--- MOM Volatility-Weighted (Last Day) ---")
print(MOM_weights_vol.iloc[-1])

MA_weights.to_csv("MA_weights_full.csv")
MOM_weights.to_csv("MOM_weights_full.csv")
MA_weights_vol.to_csv("MA_vol_weights_full.csv")
MOM_weights_vol.to_csv("MOM_vol_weights_full.csv")

print("\nSaved these files in your working directory:")
print("   MA_weights_full.csv")
print("   MOM_weights_full.csv")
print("   MA_vol_weights_full.csv")
print("   MOM_vol_weights_full.csv")


# ================================
# STEP 7: OUT-OF-SAMPLE EVALUATION
# ================================

# Purpose:
#   Split a time index into train/test segments using a simple ratio rule.
# Output:
#   - train_end: a date label used to cut the DataFrame into in-sample and out-of-sample periods.
def train_test_split_by_ratio(index, ratio=0.60):
    cutoff = int(len(index) * ratio)
    train_end = index[cutoff]
    return train_end


train_end = train_test_split_by_ratio(crypto_prices_clean.index, ratio=0.60)
test_start = crypto_prices_clean.index[crypto_prices_clean.index > train_end][0]

print("Train ends at:", train_end)
print("Test starts at:", test_start)

crypto_ret_train = crypto_ret.loc[:train_end]
MA_weights_train = MA_weights.loc[:train_end]
MOM_weights_train = MOM_weights.loc[:train_end]

crypto_ret_test = crypto_ret.loc[test_start:]
MA_weights_test = MA_weights.loc[test_start:]
MOM_weights_test = MOM_weights.loc[test_start:]

MA_train_ret = compute_portfolio_returns(crypto_ret_train, MA_weights_train)
MOM_train_ret = compute_portfolio_returns(crypto_ret_train, MOM_weights_train)

MA_test_ret = compute_portfolio_returns(crypto_ret_test, MA_weights_test)
MOM_test_ret = compute_portfolio_returns(crypto_ret_test, MOM_weights_test)

perf_train = pd.DataFrame(
    {
        "MA_in_sample": performance_metrics(MA_train_ret),
        "MOM_in_sample": performance_metrics(MOM_train_ret),
    }
)

perf_test = pd.DataFrame(
    {
        "MA_out_sample": performance_metrics(MA_test_ret),
        "MOM_out_sample": performance_metrics(MOM_test_ret),
    }
)

print("\n=== In-Sample Performance ===")
print(perf_train.round(4))

print("\n=== Out-of-Sample Performance ===")
print(perf_test.round(4))

MA_equity_train = equity_curve(MA_train_ret)
MOM_equity_train = equity_curve(MOM_train_ret)
MA_equity_test = equity_curve(MA_test_ret)
MOM_equity_test = equity_curve(MOM_test_ret)

plt.figure(figsize=(12, 6))
plt.plot(MA_equity_train, label="MA (Train)")
plt.plot(MOM_equity_train, label="MOM (Train)")
plt.title("In-Sample Equity Curves")
plt.legend()
plt.grid(True)
plt.show()

plt.figure(figsize=(12, 6))
plt.plot(MA_equity_test, label="MA (Test)")
plt.plot(MOM_equity_test, label="MOM (Test)")
plt.title("Out-of-Sample Equity Curves")
plt.legend()
plt.grid(True)
plt.show()


# ================================
# STEP 8: SENSITIVITY ANALYSIS (MA)
# ================================

ma_params_list = [(20, 100), (50, 200), (100, 300)]
rebalance_freq_list = ["D", "W", "M"]
weight_scheme = "equal"

results = []
for short_ma, long_ma in ma_params_list:
    MA_short_param, MA_long_param, MA_signal_param = compute_MA_signals(
        crypto_prices_clean, short=short_ma, long=long_ma
    )
    for freq in rebalance_freq_list:
        MA_weights_param = build_portfolio_weights(
            returns=crypto_ret,
            signals=MA_signal_param,
            freq=freq,
            weight_scheme=weight_scheme,
        )
        MA_ret_param = compute_portfolio_returns(crypto_ret, MA_weights_param)
        perf = performance_metrics(MA_ret_param)
        results.append(
            {
                "Strategy": "MA",
                "MA_short": short_ma,
                "MA_long": long_ma,
                "RebalanceFreq": freq,
                "CAGR": perf["CAGR"],
                "AnnVol": perf["AnnVol"],
                "Sharpe": perf["Sharpe"],
                "MaxDrawdown": perf["MaxDrawdown"],
            }
        )

ma_sensitivity_df = pd.DataFrame(results)
print("\n=== MA Strategy Sensitivity (no transaction costs) ===")
print(ma_sensitivity_df.round(4))


# ================================
# STEP 8b: SENSITIVITY (MOMENTUM)
# ================================

mom_lookbacks = [90, 180, 252, 360]
rebalance_freq_list = ["D", "W", "M"]

results_mom = []
for lb in mom_lookbacks:
    momentum_param, MOM_signal_param = compute_momentum_signal(
        crypto_prices_clean, lookback=lb
    )
    for freq in rebalance_freq_list:
        MOM_weights_param = build_portfolio_weights(
            returns=crypto_ret,
            signals=MOM_signal_param,
            freq=freq,
            weight_scheme="equal",
        )
        MOM_ret_param = compute_portfolio_returns(crypto_ret, MOM_weights_param)
        perf = performance_metrics(MOM_ret_param)
        results_mom.append(
            {
                "Strategy": "MOM",
                "Lookback": lb,
                "RebalanceFreq": freq,
                "CAGR": perf["CAGR"],
                "AnnVol": perf["AnnVol"],
                "Sharpe": perf["Sharpe"],
                "MaxDrawdown": perf["MaxDrawdown"],
            }
        )

mom_sensitivity_df = pd.DataFrame(results_mom)
print("\n=== Momentum Strategy Sensitivity (no transaction costs) ===")
print(mom_sensitivity_df.round(4))


# ================================
# STEP 9: CO-MOVEMENT WITH MACRO & COMMODITIES
# ================================

import statsmodels.api as sm

macro_tickers = {
    "Gold": "GC=F",
    "Oil": "CL=F",
    "VIX": "^VIX",
}

macro_prices = yf.download(
    list(macro_tickers.values()),
    start=crypto_prices_clean.index.min(),
    end=crypto_prices_clean.index.max(),
    auto_adjust=True,
    progress=False,
)

if isinstance(macro_prices.columns, pd.MultiIndex):
    macro_prices = macro_prices["Close"]

macro_prices.columns = macro_tickers.keys()
macro_returns = macro_prices.pct_change().dropna()

aligned_df = pd.concat(
    [
        MA_port_ret.rename("MA"),
        MOM_port_ret.rename("MOM"),
        bench_ret.rename(columns={"^GSPC": "SP500", "^NDX": "NASDAQ100"}),
        macro_returns,
    ],
    axis=1,
).dropna()

print("\nAligned DataFrame for Co-Movement (first rows):")
print(aligned_df.head())

corr_matrix = aligned_df.corr()
print("\n=== STATIC CORRELATION MATRIX ===")
print(corr_matrix.round(3))


# Purpose:
#   Compute rolling correlation between two aligned Series using a specified window length.
def rolling_corr(series1, series2, window=60):
    return series1.rolling(window).corr(series2)


plt.figure(figsize=(12, 6))
plt.plot(rolling_corr(aligned_df["MA"], aligned_df["SP500"]), label="MA vs SP500")
plt.plot(rolling_corr(aligned_df["MA"], aligned_df["NASDAQ100"]), label="MA vs NASDAQ100")
plt.plot(rolling_corr(aligned_df["MA"], aligned_df["Gold"]), label="MA vs Gold")
plt.plot(rolling_corr(aligned_df["MA"], aligned_df["Oil"]), label="MA vs Oil")
plt.plot(rolling_corr(aligned_df["MA"], aligned_df["VIX"]), label="MA vs VIX")
plt.title("60-Day Rolling Correlation (MA Strategy)")
plt.legend()
plt.grid(True)
plt.show()

plt.figure(figsize=(12, 6))
plt.plot(rolling_corr(aligned_df["MOM"], aligned_df["SP500"]), label="MOM vs SP500")
plt.plot(rolling_corr(aligned_df["MOM"], aligned_df["NASDAQ100"]), label="MOM vs NASDAQ100")
plt.plot(rolling_corr(aligned_df["MOM"], aligned_df["Gold"]), label="MOM vs Gold")
plt.plot(rolling_corr(aligned_df["MOM"], aligned_df["Oil"]), label="MOM vs Oil")
plt.plot(rolling_corr(aligned_df["MOM"], aligned_df["VIX"]), label="MOM vs VIX")
plt.title("60-Day Rolling Correlation (MOM Strategy)")
plt.legend()
plt.grid(True)
plt.show()


# Purpose:
#   Run an OLS regression of a portfolio return series on a set of factor returns, and return the
#   full statsmodels summary (coefficients, t-stats, R^2, etc.).
def regression_summary(port_ret, macro_df):
    df = pd.concat([port_ret, macro_df], axis=1).dropna()
    Y = df.iloc[:, 0]
    X = df.iloc[:, 1:]
    X = sm.add_constant(X)
    model = sm.OLS(Y, X).fit()
    return model.summary()


print("\n=== REGRESSION: MA strategy on macro factors ===")
print(regression_summary(aligned_df["MA"], aligned_df[["SP500", "NASDAQ100", "Gold", "Oil", "VIX"]]))

print("\n=== REGRESSION: MOM strategy on macro factors ===")
print(regression_summary(aligned_df["MOM"], aligned_df[["SP500", "NASDAQ100", "Gold", "Oil", "VIX"]]))


# ================================
# STEP 10: LLM-STYLE SENTIMENT FILTER (FEAR & GREED)
# ================================

# Purpose:
#   Fetch the Crypto Fear & Greed Index (daily) from alternative.me and transform it into a Series.
# Inputs:
#   - start_date/end_date: date bounds for slicing the downloaded history
#   - normalize_to_minus1_1: if True, map [0,100] -> [-1,1] using (value-50)/50
# Output:
#   - sentiment Series indexed by date, ready for alignment with return data.
def download_sentiment_series_from_fear_greed(
    start_date,
    end_date,
    normalize_to_minus1_1: bool = True,
) -> pd.Series:
    """
    Download daily crypto sentiment from the Crypto Fear & Greed Index API.
    Values originally in [0,100]. Optionally map to [-1,1].
    """
    url = "https://api.alternative.me/fng/"
    params = {
        "limit": 0,
        "date_format": "world",  # typically DD-MM-YYYY
    }

    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()["data"]

    df = pd.DataFrame(data)
    # Parse day-first dates (e.g. "30-11-2025")
    df["timestamp"] = pd.to_datetime(df["timestamp"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["timestamp"])
    df = df.sort_values("timestamp").set_index("timestamp")

    df = df.loc[start_date:end_date]
    df["value"] = df["value"].astype(float)

    if normalize_to_minus1_1:
        df["sentiment"] = (df["value"] - 50.0) / 50.0
    else:
        df["sentiment"] = df["value"]

    return df["sentiment"]


sentiment_series = download_sentiment_series_from_fear_greed(
    start_date=crypto_prices_clean.index.min(),
    end_date=crypto_prices_clean.index.max(),
    normalize_to_minus1_1=True,
)

print("Downloaded sentiment series. Head:")
print(sentiment_series.head())

sentiment_aligned = sentiment_series.reindex(crypto_ret.index).ffill().bfill()
sentiment_aligned = sentiment_aligned.clip(-1, 1)

print("\nAligned sentiment (first 10):")
print(sentiment_aligned.head(10))

SENTIMENT_THRESHOLD = 0.0
sentiment_filter = (sentiment_aligned >= SENTIMENT_THRESHOLD).astype(int)
sentiment_filter.name = "sentiment_filter"

print("\nSentiment filter (first 10):")
print(sentiment_filter.head(10))

# Apply filter to MA signals (only +1 side is affected)
MA_signal_LLM = MA_signal.copy()
MA_signal_LLM[MA_signal_LLM > 0] = MA_signal_LLM[MA_signal_LLM > 0].mul(
    sentiment_filter, axis=0
)

orig_long_pct = (MA_signal > 0).mean().mean() * 100
llm_long_pct = (MA_signal_LLM > 0).mean().mean() * 100

print(f"\nAverage % time in long positions (original MA): {orig_long_pct:.2f}%")
print(f"Average % time in long positions (LLM-filtered MA): {llm_long_pct:.2f}%")

MA_LLM_weights = build_portfolio_weights(
    returns=crypto_ret,
    signals=MA_signal_LLM,
    freq=rebalance_freq,
    weight_scheme=weight_scheme,
)

MA_LLM_port_ret = compute_portfolio_returns(crypto_ret, MA_LLM_weights)
MA_LLM_equity = equity_curve(MA_LLM_port_ret)

perf_MA = performance_metrics(MA_port_ret)
perf_MA_LLM = performance_metrics(MA_LLM_port_ret)

comparison = pd.DataFrame(
    {
        "MA_original": perf_MA,
        "MA_with_LLM": perf_MA_LLM,
    }
)

print("\n=== MA Strategy: Original vs LLM-Filtered (Gross Returns) ===")
print(comparison.round(4))

plt.figure(figsize=(12, 6))
plt.plot(MA_equity, label="MA original")
plt.plot(MA_LLM_equity, label="MA with LLM filter")
plt.title("Equity Curves: MA Strategy With and Without Sentiment Filter")
plt.legend()
plt.grid(True)
plt.show()
