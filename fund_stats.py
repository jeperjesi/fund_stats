import pandas as pd
import numpy as np
from scipy import stats


# =========================
# File paths
# =========================
INPUT_FILE_PATH = r"C:\Users\JosephEperjesi\Leap District Advisors\Shared - Documents\General\Andersen\Performance Reporting\Monthly Reporting"
OUTPUT_FILE_PATH = r"C:\Users\JosephEperjesi\Leap District Advisors\Shared - Documents\General\Andersen\Performance Reporting\Monthly Reporting"
FUND_FILE = INPUT_FILE_PATH + r"\Andersen returns history.xlsx"

# BENCHMARK_FILE = INPUT_FILE_PATH + r"\S&P 500 Total Return returns history.xlsx"
# BENCHMARK_SHORT_NAME = "S&P 500 Total Return"

BENCHMARK_FILE = INPUT_FILE_PATH + r"\MSCI World Net Index returns history.xlsx"
BENCHMARK_SHORT_NAME = "MSCI World"

DATE_COL = "date"
RETURN_COL = "return"

RISK_FREE_ANNUAL = 0.04


# =========================
# Data loading / prep
# =========================
def load_return_series(file_path: str, series_name: str) -> pd.DataFrame:
    """
    Load an Excel file containing:
      - month end date
      - monthly return

    Returns a cleaned DataFrame with columns:
      - month_end_date
      - <series_name>
    """
    df = pd.read_excel(file_path)

    # Normalize column names
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
    """
    Inner join on month_end_date so factor statistics use overlapping dates only.
    """
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
    """
    Compound the last N monthly returns.
    """
    clean = returns.dropna()
    if len(clean) < months:
        return np.nan
    window = clean.iloc[-months:]
    return (1 + window).prod() - 1


def ytd_return(dates: pd.Series, returns: pd.Series) -> float:
    """
    YTD based on the latest return provided.
    """
    df = pd.DataFrame({"date": dates, "ret": returns}).dropna().sort_values("date")
    if df.empty:
        return np.nan

    latest_year = df["date"].iloc[-1].year
    ytd_df = df[df["date"].dt.year == latest_year]

    if ytd_df.empty:
        return np.nan

    return (1 + ytd_df["ret"]).prod() - 1


def annualized_return(returns: pd.Series, months: int = None) -> float:
    """
    Annualized geometric return since inception.
    """
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
    """
    Annualized standard deviation of monthly returns.
    Sample std dev (ddof=1).
    """
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan
    return clean.std(ddof=1) * np.sqrt(12)


def sharpe_ratio(returns: pd.Series, rf_annual: float = 0.04) -> float:
    """
    Annualized Sharpe using monthly excess returns.
    """
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan

    rf_monthly = monthly_rf_from_annual(rf_annual)
    excess = clean - rf_monthly

    excess_std = excess.std(ddof=1)
    if excess_std == 0 or np.isnan(excess_std):
        return np.nan

    return (excess.mean() / excess_std) * np.sqrt(12)

def sharpe_ratio_ann_shortcut(returns: pd.Series, rf_annual: float = 0.04) -> float:
    """
    Annualized Sharpe using annualized returns and volatility.
    """
    clean = returns.dropna()
    if len(clean) < 2:
        return np.nan

    ann_ret = annualized_return(clean)
    ann_vol = annualized_std_dev(clean)

    return (ann_ret - rf_annual) / ann_vol


def sortino_ratio(returns: pd.Series, rf_annual: float = 0.04) -> float:
    """
    Annualized Sortino using monthly excess returns
    and downside deviation relative to monthly risk-free rate.
    """
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


def alpha_simple(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    """
    'Alpha' here interpreted as annualized active return:
    annualized fund return minus annualized benchmark return.
    """
    fund_ann = annualized_return(fund_returns)
    bench_ann = annualized_return(benchmark_returns)

    if np.isnan(fund_ann) or np.isnan(bench_ann):
        return np.nan

    return fund_ann - bench_ann

def alpha_regression(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    """
    Jensen's Alpha via OLS regression of fund returns on benchmark returns
    (risk-free rate assumed to be 0%), annualized geometrically.
    """
    # Align the two series and drop any NaNs
    #aligned = pd.concat([fund_returns, benchmark_returns], axis=1).dropna()
    #if len(aligned) < 3:
    #    return np.nan

    #fund = aligned.iloc[:, 0].values
    #bench = aligned.iloc[:, 1].values

    # OLS regression: R_fund = alpha + beta * R_bench
    beta, alpha_monthly = np.polyfit(benchmark_returns, fund_returns, 1)

    # Annualize geometrically
    alpha_annual = (1 + alpha_monthly) ** 12 - 1

    return alpha_annual


def jensen_alpha(fund_returns: pd.Series, benchmark_returns: pd.Series, rf_annual: float = 0.04) -> float:
    """
    Annualized Jensen's alpha from monthly CAPM regression:
    (Rp - Rf) = alpha + beta * (Rb - Rf) + error

    Returns annualized alpha.
    """
    df = pd.DataFrame({
        "fund": fund_returns,
        "bench": benchmark_returns
    }).dropna()

    if len(df) < 2:
        return np.nan

    rf_monthly = monthly_rf_from_annual(rf_annual)
    y = df["fund"] - rf_monthly
    x = df["bench"] - rf_monthly

    slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)

    # Annualize monthly alpha geometrically
    return (1 + intercept) ** 12 - 1


def beta(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    df = pd.DataFrame({
        "fund": fund_returns,
        "bench": benchmark_returns
    }).dropna()

    if len(df) < 2:
        return np.nan

    bench_var = np.var(df["bench"], ddof=1)
    if bench_var == 0 or np.isnan(bench_var):
        return np.nan

    return np.cov(df["fund"], df["bench"], ddof=1)[0, 1] / bench_var


def r_squared(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    df = pd.DataFrame({
        "fund": fund_returns,
        "bench": benchmark_returns
    }).dropna()

    if len(df) < 2:
        return np.nan

    corr = df["fund"].corr(df["bench"])
    if np.isnan(corr):
        return np.nan

    return corr ** 2


def excess_return(fund_returns: pd.Series, benchmark_returns: pd.Series) -> float:
    """
    Annualized geometric excess return of fund relative to benchmark:
    product((1+fund)/(1+benchmark))^(12/n) - 1
    """
    df = pd.DataFrame({
        "fund": fund_returns,
        "bench": benchmark_returns
    }).dropna()

    if len(df) == 0:
        return np.nan

    relative_growth = ((1 + df["fund"]) / (1 + df["bench"])).prod()
    return relative_growth ** (12 / len(df)) - 1 # type: ignore


# =========================
# Metric tables
# =========================
def build_single_series_metrics(dates: pd.Series, returns: pd.Series, series_name: str) -> pd.DataFrame:
    metrics = {
        "3 month return": compounded_return(returns, 3),
        "6 month return": compounded_return(returns, 6),
        "YTD return": ytd_return(dates, returns),
        "1 year return": compounded_return(returns, 12),
        "2 year ann return": annualized_return(returns, 24),
        "standard deviation": annualized_std_dev(returns),
        "sharpe": sharpe_ratio(returns, RISK_FREE_ANNUAL),
        "sharpe (ann shortcut)": sharpe_ratio_ann_shortcut(returns, RISK_FREE_ANNUAL),
        "sortino": sortino_ratio(returns, RISK_FREE_ANNUAL),
        "cumulative inception to date": cumulative_return(returns),
    }

    return pd.DataFrame.from_dict(metrics, orient="index", columns=[series_name])


def build_relative_metrics(fund_returns: pd.Series, benchmark_returns: pd.Series) -> pd.DataFrame:
    metrics = {
        "alpha": alpha_regression(fund_returns, benchmark_returns),
        "jensen's alpha": jensen_alpha(fund_returns, benchmark_returns, RISK_FREE_ANNUAL),
        "beta": beta(fund_returns, benchmark_returns),
        "R squared": r_squared(fund_returns, benchmark_returns),
        "excess return": excess_return(fund_returns, benchmark_returns),
    }

    return pd.DataFrame.from_dict(metrics, orient="index", columns=["Fund vs Benchmark"])


def build_summary_table(fund_df: pd.DataFrame, benchmark_df: pd.DataFrame) -> pd.DataFrame:
    aligned = align_series(fund_df, benchmark_df)

    fund_metrics = build_single_series_metrics(
        aligned["month_end_date"],
        aligned["Fund Return"],
        "Fund Return"
    )

    benchmark_metrics = build_single_series_metrics(
        aligned["month_end_date"],
        aligned["Benchmark Return"],
        "Benchmark Return"
    )

    relative_metrics = build_relative_metrics(
        aligned["Fund Return"],
        aligned["Benchmark Return"]
    )

    summary = pd.concat([fund_metrics, benchmark_metrics, relative_metrics], axis=1)

    # Optional formatting helper columns if you want easier export/reading
    return summary.reset_index().rename(columns={"index": "Metric"})


# =========================
# Main
# =========================
def main():
    fund_df = load_return_series(FUND_FILE, "Fund Return")
    benchmark_df = load_return_series(BENCHMARK_FILE, "Benchmark Return")

    summary_table = build_summary_table(fund_df, benchmark_df)

    pd.set_option("display.float_format", lambda x: f"{x:,.6f}")
    print(summary_table)

    # Save results
    summary_table.to_excel(OUTPUT_FILE_PATH + r"\return_metrics_summary vs " + BENCHMARK_SHORT_NAME + ".xlsx", index=False)
    print("\nSaved to: return_metrics_summary vs " + BENCHMARK_SHORT_NAME + ".xlsx")


if __name__ == "__main__":
    main()