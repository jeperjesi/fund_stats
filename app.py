import io
import json
import os
import tkinter as tk
from tkinter import filedialog

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

from fund_stats import (
    align_series,
    build_single_series_metrics,
    build_summary_table,
    load_return_series,
)

CONFIG_PATH = os.path.join(os.path.dirname(__file__), "config.json")

MONTH_COLS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

RETURN_METRICS = {
    "3 month return",
    "6 month return",
    "YTD return",
    "1 year return",
    "annualized return",
    "2 year ann return",
    "standard deviation",
    "cumulative inception to date",
    "alpha",
    "jensen's alpha",
    "excess return",
}


@st.cache_data
def load_config() -> dict:
    with open(CONFIG_PATH) as f:
        return json.load(f)


def browse_folder(title: str, key: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes("-topmost", 1)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    if path:
        st.session_state[key] = path


def browse_file(title: str, key: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes("-topmost", 1)
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
    )
    root.destroy()
    if path:
        st.session_state[key] = path


def format_metric(metric: str, value) -> str:
    if pd.isna(value):
        return "N/A"
    if metric in RETURN_METRICS:
        return f"{value * 100:.2f}%"
    return f"{value:.4f}"


def fmt_pct(value, decimals: int = 2) -> str:
    if pd.isna(value):
        return ""
    return f"{value * 100:.{decimals}f}%"


def build_vami(dates: pd.Series, returns: pd.Series, start: float = 1000.0) -> pd.Series:
    vami = (1 + returns).cumprod() * start
    return pd.Series(vami.values, index=dates, name="vami")


def build_return_table(dates: pd.Series, returns: pd.Series) -> pd.DataFrame:
    df = pd.DataFrame({"date": dates, "ret": returns}).dropna().sort_values("date")
    df["year"] = df["date"].dt.year
    df["month_name"] = df["date"].dt.strftime("%b")

    pivot = df.pivot(index="year", columns="month_name", values="ret")

    # Ensure all 12 month columns exist in calendar order
    for m in MONTH_COLS:
        if m not in pivot.columns:
            pivot[m] = np.nan
    pivot = pivot[MONTH_COLS]

    # YTD: compounded return for all months present in that year
    def _ytd(row):
        vals = row.dropna()
        return (1 + vals).prod() - 1 if len(vals) else np.nan

    pivot["YTD"] = pivot.apply(_ytd, axis=1)

    # ITD: cumulative from inception through the last month of each year
    df_sorted = df.sort_values("date")
    itd = {}
    for year in pivot.index:
        subset = df_sorted[df_sorted["year"] <= year]["ret"]
        itd[year] = (1 + subset).prod() - 1 if len(subset) else np.nan
    pivot["ITD"] = pd.Series(itd)

    pivot.index.name = "Year"
    return pivot


def format_return_table(pivot: pd.DataFrame) -> pd.DataFrame:
    return pivot.apply(
        lambda col: col.map(lambda v: fmt_pct(v, 2))
    )


def build_excel_multi(summary: pd.DataFrame, sheets: dict) -> bytes:
    """Write summary + one sheet per return table."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name)
    return buf.getvalue()


def derive_benchmark_dir(input_dir: str, benchmark_subpath: str) -> str:
    normalized = input_dir.replace("\\", "/")
    idx = normalized.lower().find("/general/")
    if idx == -1:
        return ""
    prefix = input_dir[:idx]
    return os.path.join(prefix, benchmark_subpath)


def folder_row(
    label: str, config_default: str, state_key: str, browse_key: str, disabled: bool = False
) -> str:
    col1, col2 = st.columns([4, 1])
    with col1:
        typed = st.text_input(
            label,
            value=st.session_state.get(state_key, config_default),
            key=f"{state_key}_input",
            disabled=disabled,
        )
    with col2:
        st.write("")
        st.write("")
        if st.button("Browse", key=browse_key, disabled=disabled):
            browse_folder(f"Select {label.lower()}", state_key)
            st.rerun()
    return st.session_state.get(state_key, typed)


def main():
    st.set_page_config(page_title="Fund Statistics", layout="wide")

    st.markdown(
        """
        <style>
            section[data-testid="stSidebar"] { min-width: 430px; width: 430px; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("Fund Statistics")

    config = load_config()
    fund_names = list(config["funds"].keys()) + ["Custom..."]
    benchmark_names = list(config["benchmarks"].keys()) + ["Custom..."]

    # ── Sidebar ───────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("Parameters")

        st.subheader("Input Folder")
        input_dir = folder_row(
            "Input folder", config["input_path"], "input_path", "input_browse"
        )

        st.subheader("Fund")
        fund_choice = st.selectbox("Fund name", fund_names, key="fund_choice")
        if fund_choice == "Custom...":
            col1, col2 = st.columns([4, 1])
            with col1:
                fund_typed = st.text_input(
                    "Fund file path",
                    value=st.session_state.get("fund_custom_path", ""),
                    key="fund_path_input",
                )
            with col2:
                st.write("")
                st.write("")
                if st.button("Browse", key="fund_browse"):
                    browse_file("Select fund file", "fund_custom_path")
                    st.rerun()
            fund_file = st.session_state.get("fund_custom_path", fund_typed)
        else:
            fund_file = os.path.join(input_dir, config["funds"][fund_choice])

        st.write("Fund returns file should be named '[Fund Name] returns history.xlsx'")
        st.write("File should have a 'date' column with month end date and a 'return' column with monthly returns as decimals (e.g. 0.0123 or 1.23%)")
        st.write("")

        st.subheader("Benchmark")
        include_benchmark = st.checkbox("Include benchmark", value=False)

        # Derive benchmark folder from input path; reset when input_dir changes
        derived_benchmark_dir = derive_benchmark_dir(input_dir, config.get("benchmark_subpath", ""))
        if st.session_state.get("_prev_input_dir") != input_dir:
            st.session_state["_prev_input_dir"] = input_dir
            st.session_state["benchmark_dir"] = derived_benchmark_dir
        if "benchmark_dir" not in st.session_state:
            st.session_state["benchmark_dir"] = derived_benchmark_dir

        col1, col2 = st.columns([4, 1])
        with col1:
            st.text_input(
                "Benchmark folder",
                key="benchmark_dir",
                disabled=not include_benchmark,
            )
        with col2:
            st.write("")
            st.write("")
            if st.button("Browse", key="benchmark_dir_browse", disabled=not include_benchmark):
                browse_folder("Select benchmark folder", "benchmark_dir")
                st.rerun()
        benchmark_dir = st.session_state["benchmark_dir"]

        if include_benchmark and benchmark_dir and not os.path.isdir(benchmark_dir):
            st.warning(f"Benchmark folder not found: {benchmark_dir}")

        benchmark_choice = st.selectbox(
            "Benchmark", benchmark_names, key="benchmark_choice", disabled=not include_benchmark
        )
        if include_benchmark:
            if benchmark_choice == "Custom...":
                col1, col2 = st.columns([4, 1])
                with col1:
                    benchmark_typed = st.text_input(
                        "Benchmark file path",
                        value=st.session_state.get("benchmark_custom_path", ""),
                        key="benchmark_path_input",
                    )
                with col2:
                    st.write("")
                    st.write("")
                    if st.button("Browse", key="benchmark_browse"):
                        browse_file("Select benchmark file", "benchmark_custom_path")
                        st.rerun()
                benchmark_file = st.session_state.get("benchmark_custom_path", benchmark_typed)
                benchmark_label = (
                    os.path.splitext(os.path.basename(benchmark_file))[0]
                    if benchmark_file else "Benchmark"
                )
            else:
                benchmark_file = os.path.join(benchmark_dir, config["benchmarks"][benchmark_choice])
                benchmark_label = benchmark_choice
        else:
            benchmark_file = None
            benchmark_label = None

        st.subheader("Date Range")
        filter_dates = st.checkbox("Filter date range", value=False)
        if filter_dates:
            import datetime
            start_filter = st.date_input(
                "Start date",
                value=datetime.date(2020, 1, 1),
                key="start_filter",
            )
            end_filter = st.date_input(
                "End date",
                value=datetime.date.today(),
                key="end_filter",
            )
        else:
            start_filter = None
            end_filter = None

        st.subheader("Risk-Free Rate")
        rf_pct = st.number_input(
            "Annual risk-free rate (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(config["default_risk_free_annual_pct"]),
            step=0.0001,
            format="%.4f",
        )
        rf_annual = rf_pct / 100

        st.subheader("Output Folder")
        export_to_disk = st.checkbox("Save to output folder after running", value=False)
        output_dir = folder_row(
            "Output folder", config["output_path"], "output_path", "output_browse",
            disabled=not export_to_disk,
        )

        run = st.button("Run", type="primary", use_container_width=True)

    # ── Main area ─────────────────────────────────────────────────────────
    if not run:
        st.info("Configure parameters in the sidebar and click **Run**.")
        return

    if not fund_file:
        st.error("Please select or enter a fund file path.")
        return
    if include_benchmark and not benchmark_file:
        st.error("Please select or enter a benchmark file path.")
        return

    with st.spinner("Loading data and computing metrics…"):
        try:
            fund_df = load_return_series(fund_file, "Fund Return")

            if filter_dates:
                start_ts = pd.Timestamp(start_filter)
                end_ts = pd.Timestamp(end_filter)
                fund_df = fund_df[
                    (fund_df["month_end_date"] >= start_ts) &
                    (fund_df["month_end_date"] <= end_ts)
                ].reset_index(drop=True)

            if include_benchmark:
                benchmark_df = load_return_series(benchmark_file, "Benchmark Return")

                if filter_dates:
                    benchmark_df = benchmark_df[
                        (benchmark_df["month_end_date"] >= start_ts) &
                        (benchmark_df["month_end_date"] <= end_ts)
                    ].reset_index(drop=True)

                aligned = align_series(fund_df, benchmark_df)
                summary = build_summary_table(fund_df, benchmark_df, rf_annual)
                start_date = aligned["month_end_date"].min().strftime("%B %Y")
                end_date = aligned["month_end_date"].max().strftime("%B %Y")
                n_months = len(aligned)

                fund_dates = set(fund_df["month_end_date"])
                bench_dates = set(benchmark_df["month_end_date"])
                aligned_dates = set(aligned["month_end_date"])
                fund_only = sorted(fund_dates - aligned_dates)
                bench_only = sorted(bench_dates - aligned_dates)
            else:
                aligned = fund_df.rename(columns={"Fund Return": "Fund Return"})
                summary = build_single_series_metrics(
                    fund_df["month_end_date"],
                    fund_df["Fund Return"],
                    "Fund Return",
                    rf_annual,
                ).reset_index().rename(columns={"index": "Metric"})
                start_date = fund_df["month_end_date"].min().strftime("%B %Y")
                end_date = fund_df["month_end_date"].max().strftime("%B %Y")
                n_months = len(fund_df)
                fund_only = []
                bench_only = []
        except Exception as e:
            st.error(f"Error: {e}")
            return

    fund_label = (
        fund_choice if fund_choice != "Custom..."
        else os.path.splitext(os.path.basename(fund_file))[0]
    )
    title = f"{fund_label} vs {benchmark_label}" if include_benchmark else fund_label

    st.subheader(title)
    st.info(
        f"**Data range:** {start_date} – {end_date} ({n_months} months)"
        f"   |   **Risk-free rate:** {rf_pct:.4f}%"
    )

    if fund_only or bench_only:
        lines = ["**Some dates were excluded because they appear in only one series (inner join applied):**"]
        if fund_only:
            lines.append(f"- **Fund only** ({len(fund_only)}): {', '.join(d.strftime('%b %Y') for d in fund_only)}")
        if bench_only:
            lines.append(f"- **Benchmark only** ({len(bench_only)}): {', '.join(d.strftime('%b %Y') for d in bench_only)}")
        st.warning("\n".join(lines))

    # ── Tabs ──────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["Summary Statistics", "Monthly Returns", "VAMI"])

    # Build return tables and VAMI here so they're available for save-to-disk
    fund_pivot = build_return_table(fund_df["month_end_date"], fund_df["Fund Return"])
    bench_pivot = None
    if include_benchmark:
        bench_pivot = build_return_table(
            benchmark_df["month_end_date"], benchmark_df["Benchmark Return"]
        )

    if include_benchmark:
        fund_vami = build_vami(aligned["month_end_date"], aligned["Fund Return"])
        bench_vami = build_vami(aligned["month_end_date"], aligned["Benchmark Return"])
        vami_df = pd.DataFrame(
            {fund_label: fund_vami.values, benchmark_label: bench_vami.values},
            index=aligned["month_end_date"],
        )
    else:
        fund_vami = build_vami(fund_df["month_end_date"], fund_df["Fund Return"])
        vami_df = pd.DataFrame(
            {fund_label: fund_vami.values},
            index=fund_df["month_end_date"],
        )
    vami_df.index.name = "Date"

    with tab1:
        display_df = summary.copy()
        for col in display_df.columns:
            if col == "Metric":
                continue
            display_df[col] = display_df.apply(
                lambda row, c=col: format_metric(row["Metric"], row[c]), axis=1
            )
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        summary_filename = (
            f"return_metrics_summary vs {benchmark_label}.xlsx"
            if include_benchmark
            else f"return_metrics_summary {fund_label}.xlsx"
        )
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
        st.download_button(
            label="Download Excel",
            data=buf.getvalue(),
            file_name=summary_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_summary",
        )

    with tab2:
        st.markdown(f"**{fund_label}**")
        st.dataframe(format_return_table(fund_pivot), use_container_width=True)

        if include_benchmark:
            st.markdown(f"**{benchmark_label}**")
            st.dataframe(format_return_table(bench_pivot), use_container_width=True)

        returns_sheets = {f"{fund_label} Returns": fund_pivot}
        if bench_pivot is not None:
            returns_sheets[f"{benchmark_label} Returns"] = bench_pivot

        returns_filename = (
            f"monthly_returns vs {benchmark_label}.xlsx"
            if include_benchmark
            else f"monthly_returns {fund_label}.xlsx"
        )
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sheet_name, df in returns_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name)
        st.download_button(
            label="Download Excel",
            data=buf.getvalue(),
            file_name=returns_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_returns",
        )

    with tab3:
        st.caption("Growth of $1,000 invested at inception")

        vami_reset = vami_df.reset_index()
        vami_reset["date_str"] = vami_reset["Date"].dt.strftime("%b-%y")
        date_order = vami_reset["date_str"].tolist()
        vami_long = vami_reset.drop(columns="Date").melt(
            id_vars="date_str", var_name="Series", value_name="Value"
        )

        base = alt.Chart(vami_long).encode(
            x=alt.X(
                "date_str:O",
                title=None,
                sort=date_order,
                axis=alt.Axis(labelAngle=-45, labelLimit=60),
            ),
            color=alt.Color("Series:N", legend=alt.Legend(title=None)),
        )
        lines = base.mark_line().encode(y=alt.Y("Value:Q", title="Value ($)"))
        points = base.mark_point(opacity=0, size=100).encode(
            y=alt.Y("Value:Q"),
            tooltip=[
                alt.Tooltip("date_str:O", title="Date"),
                alt.Tooltip("Series:N"),
                alt.Tooltip("Value:Q", title="Value ($)", format=",.2f"),
            ],
        )
        chart = (lines + points).properties(height=450, padding={"bottom": 20}).interactive()
        st.altair_chart(chart, use_container_width=True)

        vami_filename = (
            f"vami vs {benchmark_label}.xlsx"
            if include_benchmark
            else f"vami {fund_label}.xlsx"
        )
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            vami_df.to_excel(writer, sheet_name="VAMI")
        st.download_button(
            label="Download Excel",
            data=buf.getvalue(),
            file_name=vami_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_vami",
        )

    # ── Save to disk ──────────────────────────────────────────────────────
    if export_to_disk:
        all_sheets = {f"{fund_label} Returns": fund_pivot}
        if bench_pivot is not None:
            all_sheets[f"{benchmark_label} Returns"] = bench_pivot
        all_sheets["VAMI"] = vami_df

        full_filename = (
            f"return_metrics_summary vs {benchmark_label}.xlsx"
            if include_benchmark
            else f"return_metrics_summary {fund_label}.xlsx"
        )
        if output_dir and os.path.isdir(output_dir):
            out_file = os.path.join(output_dir, full_filename)
            with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
                summary.to_excel(writer, index=False, sheet_name="Summary")
                for sheet_name, df in all_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name)
            st.success(f"Saved to: {out_file}")
        elif output_dir:
            st.warning(f"Output folder not found: {output_dir}")


if __name__ == "__main__":
    main()
