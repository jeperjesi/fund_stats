# Fund Statistics

A Streamlit app for computing and displaying fund performance statistics against a benchmark.

---

## Getting Started

**Prerequisite:** Python 3 must be installed and added to your system PATH.
When installing Python from [python.org](https://www.python.org/downloads/), check **"Add Python to PATH"** on the first screen before clicking Install.

**To launch the app:** double-click `launch.bat`.

On first run it will automatically create a virtual environment and install all dependencies. Subsequent launches skip straight to opening the app in your browser.

---

## Configuration

All configurable defaults live in `config.json`. You should update this file once after cloning to match your local directory structure.

```json
{
  "input_path":        "default folder shown in the Input Folder field",
  "funds": {
    "Fund Name":       "Fund Name returns history.xlsx"
  },
  "benchmarks": {
    "Benchmark Name":  "Benchmark File.xlsx"
  },
  "benchmark_subpath": "General\\Leap\\Benchmark Returns",
  "output_path":       "default folder shown in the Output Folder field",
  "default_risk_free_annual_pct": 4.0
}
```

### Adding a fund or benchmark

Add a new entry to `"funds"` or `"benchmarks"` — the key is the display name shown in the dropdown, and the value is the filename (not the full path).

```json
"funds": {
  "Andersen": "Andersen returns history.xlsx",
  "MyNewFund": "MyNewFund returns history.xlsx"
}
```

### Benchmark folder path

The benchmark folder is derived automatically from the input path. Everything up to `\General\` is taken from the input path, and `General\Leap\Benchmark Returns` (controlled by `benchmark_subpath`) is appended. For example:

| Input path | Derived benchmark folder |
|---|---|
| `C:\Users\Alice\Leap District Advisors\Shared - Documents\General\Andersen\...` | `C:\Users\Alice\Leap District Advisors\Shared - Documents\General\Leap\Benchmark Returns` |
| `C:\Users\Bob\Leap District Advisors\Shared - Documents\General\Andersen\...` | `C:\Users\Bob\Leap District Advisors\Shared - Documents\General\Leap\Benchmark Returns` |

The derived path is shown in the sidebar and can be overridden manually or via Browse if needed.

---

## Input File Format

Fund and benchmark files must be Excel workbooks (`.xlsx`) with exactly two columns:

| Column name | Description |
|---|---|
| `date` | Month-end date |
| `return` | Monthly return as a decimal (e.g. `0.0123` for 1.23%) |

Column names are case-insensitive. Fund files should be named `[Fund Name] returns history.xlsx` and placed in the input folder. Benchmark files should be placed in the benchmark folder.

---

## Using the App

### Sidebar parameters

| Parameter | Description |
|---|---|
| **Input Folder** | Folder containing fund return files. Defaults to the value in `config.json`. |
| **Fund** | Select a fund from the dropdown or choose `Custom...` to browse for any file. |
| **Include benchmark** | Check to enable benchmark comparison. Unchecked by default. |
| **Benchmark folder** | Auto-derived from the input path. Can be overridden. Greyed out when benchmark is disabled. |
| **Benchmark** | Select a benchmark from the dropdown or choose `Custom...`. |
| **Filter date range** | Check to restrict statistics to a specific date window. |
| **Risk-free rate** | Annual risk-free rate as a percentage, used for Sharpe and Sortino calculations. |
| **Save to output folder** | Check to save the Excel file to disk after running. Greyed out by default. |

Click **Run** to compute results.

### Tab 1 — Summary Statistics

Displays a table of performance metrics for the fund (and benchmark if included):

- 3-month, 6-month, YTD, 1-year returns
- Annualized return (since inception or selected date range)
- 2-year annualized return
- Standard deviation (annualized)
- Sharpe ratio, Sharpe ratio (annualized shortcut), Sortino ratio
- Cumulative return (inception to date)
- Alpha (regression), Jensen's alpha, Beta, R², Excess return *(benchmark only)*

If any dates exist in one series but not the other, a warning is shown listing the excluded months.

### Tab 2 — Monthly Returns

A calendar-style grid with years as rows and months as columns, plus:

- **YTD** — compounded return for all available months in that year
- **ITD** — cumulative return from inception through the end of that year

If a benchmark is included, its return table appears below the fund table.

### Excel export

A **Download Excel** button is always available in the browser. It produces a workbook with:

- **Summary** sheet — the statistics table
- **[Fund] Returns** sheet — the monthly return grid
- **[Benchmark] Returns** sheet — the benchmark return grid *(if benchmark included)*

If **Save to output folder** is checked, the same file is also written to the selected output folder on disk.

---

## Project Structure

```
fund_stats/
├── app.py              # Streamlit UI
├── fund_stats.py       # Performance calculation functions
├── config.json         # User-configurable defaults
├── requirements.txt    # Python dependencies
├── launch.bat          # One-click launcher (creates venv on first run)
└── README.md
```

---

## Dependencies

Managed automatically by `launch.bat` via `requirements.txt`:

- [Streamlit](https://streamlit.io/)
- [pandas](https://pandas.pydata.org/)
- [NumPy](https://numpy.org/)
- [SciPy](https://scipy.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
