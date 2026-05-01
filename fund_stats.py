import pandas as pd
import numpy as np
from scipy import stats

DATE_COL = "date"
RETURN_COL = "return"


# =========================
# Data loading / prep
# =========================
def load_return_series(file_path: str, series_name: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip().lower() for c in df.columns]

    if DATE_COL not in df.columns or RETURN_COL not in df.columns:
        raise ValueError(
            f"{file_path} must contain columns '{DATE_COL}' and '{RETURN_COL}'. "
            f"Found columns: {list(df.columns)}"
        )

    df = df[[DATE_COL, RETURN_COL]].copy()
    df[DATE_COL] = pd.to_datetime(df[DATE_COL])
    df[RETURN_COL] = pd.to_numeric(df[RETURN_COL], errors="coerce")
    df = df.dropna(subset=[DATE_COL, RETURN_COL]).sort_values(DATE_COL).reset_index(drop=True)
    df = df.rename(columns={DATE_COL: "month_end_date", RETURN_COL: series_name})
    return df


def align_series(fund_df: pd.DataFrame, benchmark_df: pd.DataFrame) -> pd.DataFrame:
    merged = pd.merge(fund_df, benchmark_df, on="month_end_date", how="inner")
    merged = merged.sort_values("month_end_date").reset_index(drop=True)
    if merged.empty:
        raise ValueError("No overlapping dates found between fund and benchmark series.")
    return merged


# =========================
# Helper functions
# =========================
def monthly_rf_from_annual(rf_annual: float) -> float:
    return (1 + rf_annual) ** (1 / 12) - 1


def compounded_return(returns: pd.Series, months: int) -> float:
    clean = returns.dropna()
    if len(clean) < months:
        return np.nan
    window = clean.iloc[-months:]
    return (1 + window).prod() - 1


def ytd_return(dates: pd.Series, returns: pd.Series) -> float:
    df = pd.DataFrame({"date": dates, "ret": returns}).dropna().sort_values("date")
    if df.empty:
        return np.nan
    latest_year = df["date"].iloc[-1].year
    ytd_df = df[df["date"].dt.year == latest_year]
    if ytd_df.empty:
        return np.nan
    return (1 + ytd_df["ret"]).prod() - 1


def annualized_return(returns: pd.Series, months: int = None) -> float:
    clean = returns.dropna()
    n = len(clean)
    if n == 0:
        return np.nan
    if months is not None:
        window = clean.iloc[-months:]
        n = months
    else:
        window = clean
    total_return = (1 + window).prod()
    return total_return ** (12 / n) - 1


def cumulative_return(returns: pd.Series) -> float:
    clean = returns.dropna()
    if len(clean) == 0:
        return np.nan
    return (1 + clean).prod() - 1


def annualized_std_dev(returns: pd.Series) -> float:
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan
    return clean.std(ddof=1) * np.sqrt(12)


def sharpe_ratio(returns: pd.Series, rf_annual: float) -> float:
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan
    rf_monthly = monthly_rf_from_annual(rf_annual)
    excess = clean - rf_monthly
    excess_std = excess.std(ddof=1)
    if excess_std == 0 or np.isnan(excess_std):
        return np.nan
    return (excess.mean() / excess_std) * np.sqrt(12)


def sharpe_ratio_ann_shortcut(returns: pd.Series, rf_annual: float) -> float:
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan
    ann_ret = annualized_return(clean)
    ann_vol = annualized_std_dev(clean)
    return (ann_ret - rf_annual) / ann_vol


def sortino_ratio(returns: pd.Series, rf_annual: float) -> float:
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan
    rf_monthly = monthly_rf_from_annual(rf_annual)
    excess = clean - rf_monthly
    downside = np.minimum(excess, 0)
    downside_dev = np.sqrt(np.mean(downside ** 2))
    if downside_dev == 0 or np.isnan(downside_dev):
        return np.nan
    return (excess.mean() / downside_dev) * np.sqrt(12)


def alpha_regression(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    beta_val, alpha_monthly = np.polyfit(benchmark_returns, fund_returns, 1)
    return (1 + alpha_monthly) ** 12 - 1


def jensen_alpha(fund_returns: pd.Series, benchmark_returns: pd.Series, rf_annual: float) -> float:
    df = pd.DataFrame({"fund": fund_returns, "bench": benchmark_returns}).dropna()
    if len(df) < 2:
        return np.nan
    rf_monthly = monthly_rf_from_annual(rf_annual)
    y = df["fund"] - rf_monthly
    x = df["bench"] - rf_monthly
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
    return (1 + intercept) ** 12 - 1


def beta(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    df = pd.DataFrame({"fund": fund_returns, "bench": benchmark_returns}).dropna()
    if len(df) < 2:
        return np.nan
    bench_var = np.var(df["bench"], ddof=1)
    if bench_var == 0 or np.isnan(bench_var):
        return np.nan
    return np.cov(df["fund"], df["bench"], ddof=1)[0, 1] / bench_var


def r_squared(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    df = pd.DataFrame({"fund": fund_returns, "bench": benchmark_returns}).dropna()
    if len(df) < 2:
        return np.nan
    corr = df["fund"].corr(df["bench"])
    if np.isnan(corr):
        return np.nan
    return corr ** 2


def excess_return(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    df = pd.DataFrame({"fund": fund_returns, "bench": benchmark_returns}).dropna()
    if len(df) == 0:
        return np.nan
    relative_growth = ((1 + df["fund"]) / (1 + df["bench"])).prod()
    return relative_growth ** (12 / len(df)) - 1  # type: ignore


# =========================
# Metric tables
# =========================
def build_single_series_metrics(
    dates: pd.Series, returns: pd.Series, series_name: str, rf_annual: float
) -> pd.DataFrame:
    metrics = {
        "3 month return": compounded_return(returns, 3),
        "6 month return": compounded_return(returns, 6),
        "YTD return": ytd_return(dates, returns),
        "1 year return": compounded_return(returns, 12),
        "annualized return": annualized_return(returns),
        "2 year ann return": annualized_return(returns, 24),
        "standard deviation": annualized_std_dev(returns),
        "sharpe": sharpe_ratio(returns, rf_annual),
        "sharpe (ann shortcut)": sharpe_ratio_ann_shortcut(returns, rf_annual),
        "sortino": sortino_ratio(returns, rf_annual),
        "cumulative inception to date": cumulative_return(returns),
    }
    return pd.DataFrame.from_dict(metrics, orient="index", columns=[series_name])


def build_relative_metrics(
    fund_returns: pd.Series, benchmark_returns: pd.Series, rf_annual: float
) -> pd.DataFrame:
    metrics = {
        "alpha": alpha_regression(fund_returns, benchmark_returns),
        "jensen's alpha": jensen_alpha(fund_returns, benchmark_returns, rf_annual),
        "beta": beta(fund_returns, benchmark_returns),
        "R squared": r_squared(fund_returns, benchmark_returns),
        "excess return": excess_return(fund_returns, benchmark_returns),
    }
    return pd.DataFrame.from_dict(metrics, orient="index", columns=["Fund vs Benchmark"])


def build_summary_table(
    fund_df: pd.DataFrame, benchmark_df: pd.DataFrame, rf_annual: float
) -> pd.DataFrame:
    aligned = align_series(fund_df, benchmark_df)

    fund_metrics = build_single_series_metrics(
        aligned["month_end_date"], aligned["Fund Return"], "Fund Return", rf_annual
    )
    benchmark_metrics = build_single_series_metrics(
        aligned["month_end_date"], aligned["Benchmark Return"], "Benchmark Return", rf_annual
    )
    relative_metrics = build_relative_metrics(
        aligned["Fund Return"], aligned["Benchmark Return"], rf_annual
    )

    summary = pd.concat([fund_metrics, benchmark_metrics, relative_metrics], axis=1)
    return summary.reset_index().rename(columns={"index": "Metric"})


# =========================
# Main (CLI fallback)
# =========================
def main():
    import json, os

    config_path = os.path.join(os.path.dirname(__file__), "config.json")
    with open(config_path) as f:
        config = json.load(f)

    fund_name = "Andersen"
    benchmark_name = "MSCI World Index"
    rf_annual = config["default_risk_free_annual_pct"] / 100

    fund_file = config["funds"][fund_name]
    benchmark_file = config["benchmarks"][benchmark_name]
    output_path = config["output_path"]

    fund_df = load_return_series(fund_file, "Fund Return")
    benchmark_df = load_return_series(benchmark_file, "Benchmark Return")
    summary_table = build_summary_table(fund_df, benchmark_df, rf_annual)

    pd.set_option("display.float_format", lambda x: f"{x:,.6f}")
    print(summary_table)

    out_file = os.path.join(output_path, f"return_metrics_summary vs {benchmark_name}.xlsx")
    summary_table.to_excel(out_file, index=False)
    print(f"\nSaved to: {out_file}")


if __name__ == "__main__":
    main()
