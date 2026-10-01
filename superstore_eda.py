"""
Exploratory Data Analysis (EDA) on Superstore Sales Data
=========================================================
Libraries : Pandas, Matplotlib, Seaborn
Sections  : 1. Data Loading        3. Visualization
            2. Data Cleaning       4. Outlier Detection

Usage:
    python superstore_eda.py                      # uses the default file path below
    python superstore_eda.py path/to/file.xlsx    # or pass the path as an argument

Every chart is shown on screen (if a display is available) and also saved as a
PNG in the 'eda_plots' folder next to the script.
The Excel file should sit in the same folder as the script.
"""

import atexit
import os
import sys

# Keep the console window open when the script is double-clicked, so output and
# any error message can be read. (Registered first so it also works if an import
# below fails.) Skipped automatically when there is no interactive terminal.
def _pause_before_exit():
    if sys.stdin is not None and sys.stdin.isatty():
        input("\nPress Enter to close this window...")


atexit.register(_pause_before_exit)

import warnings

print("Starting... loading libraries (the first run can take up to a minute).", flush=True)

try:
    import matplotlib.pyplot as plt
    import pandas as pd
    import seaborn as sns
except ImportError as err:
    sys.exit(f"Missing library: {err}\n"
             "Install the requirements with:\n"
             "    pip install pandas matplotlib seaborn openpyxl")

warnings.filterwarnings("ignore")
print("Libraries loaded. Running the analysis...", flush=True)

# --------------------------------------------------------------------------
# Settings
# --------------------------------------------------------------------------
# All paths are relative to the folder that contains this script, so it works
# no matter where it is launched from (double-click, terminal, IDE).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = "sample_-_superstore__2_.xlsx"


def find_data_file():
    """Return the path to the Superstore workbook.
    Uses a path given on the command line if there is one; otherwise looks in the
    script's folder for the default name, then for any .xlsx with 'superstore' in
    its name (so 'sample - superstore (2).xlsx' also works)."""
    if len(sys.argv) > 1:
        return sys.argv[1]
    default = os.path.join(SCRIPT_DIR, DEFAULT_FILE)
    if os.path.exists(default):
        return default
    matches = sorted(f for f in os.listdir(SCRIPT_DIR)
                     if f.lower().endswith(".xlsx") and "superstore" in f.lower()
                     and not f.startswith("~$"))  # ignore Excel temp files
    if matches:
        return os.path.join(SCRIPT_DIR, matches[0])
    return default  # not found: the check below reports a clear error


FILE_PATH = find_data_file()

if not os.path.exists(FILE_PATH):
    sys.exit(f"Data file not found: {FILE_PATH}\n"
             "Put the Superstore .xlsx file in the same folder as this script "
             "(its name must contain 'superstore'), or pass its path as an argument.")
print(f"Using data file: {FILE_PATH}", flush=True)

# True : each chart opens in a window and the script waits until you close it.
# False: charts are only saved to the PNG folder and the script runs straight through.
SHOW_PLOTS = True

PLOT_DIR = os.path.join(SCRIPT_DIR, "eda_plots")
os.makedirs(PLOT_DIR, exist_ok=True)

sns.set_theme(style="whitegrid", palette="deep")
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

BLUE, RED = "#2e86c1", "#c0392b"


def section(title):
    """Print a clear heading in the console output."""
    print("\n" + "=" * 70)
    print(title)
    print("=" * 70)


def save_and_show(filename):
    """Tidy the layout, save the current figure as a PNG, then display it."""
    plt.tight_layout()
    plt.savefig(os.path.join(PLOT_DIR, filename), dpi=120, bbox_inches="tight")
    print(f"[chart saved: {filename}]", flush=True)
    if SHOW_PLOTS:
        print("   -> Chart window opened (check the taskbar if you can't see it). "
              "Close it to continue.", flush=True)
        plt.show()
    plt.close()


# ==========================================================================
# 1. DATA LOADING
# ==========================================================================
section("1. DATA LOADING")

# The workbook has 3 sheets (Orders, People, Returns); the sales data is in 'Orders'
df = pd.read_excel(FILE_PATH, sheet_name="Orders")

print("\nFirst 10 rows:")
print(df.head(10))

print("\nShape (rows, columns):", df.shape)

print("\nColumn names:")
for i, col in enumerate(df.columns, start=1):
    print(f"{i:>2}. {col}")

print("\nBasic info:")
df.info()

print("\nSummary statistics (numeric columns):")
print(df[["Sales", "Quantity", "Discount", "Profit"]].describe().round(2))

# ==========================================================================
# 2. DATA CLEANING
# ==========================================================================
section("2. DATA CLEANING")

# ---- 2.1 Check for missing values ----------------------------------------
print("\nMissing values per column (only columns with gaps):")
missing = df.isnull().sum()
print(missing[missing > 0])

print("\nRows with a missing Postal Code:")
print(df.loc[df["Postal Code"].isnull(), ["City", "State"]].drop_duplicates())

# ---- 2.2 Handle missing values -------------------------------------------
# All missing postal codes belong to one city (Burlington, Vermont). A postal code
# is an identifier, so mean/median/mode make no sense, and dropping rows would
# lose valid sales data. We fill with the city's known ZIP code 05401 (assumption).
mask = (df["City"] == "Burlington") & (df["State"] == "Vermont") & df["Postal Code"].isnull()
df.loc[mask, "Postal Code"] = 5401

# ---- 2.3 Check data types and convert if necessary ------------------------
print("\nData types before conversion:")
print(df.dtypes)

# Postal codes are labels, not numbers: store as 5-digit text (keeps leading zeros)
df["Postal Code"] = df["Postal Code"].astype(int).astype(str).str.zfill(5)

# Make sure the date columns are real datetimes
df["Order Date"] = pd.to_datetime(df["Order Date"])
df["Ship Date"] = pd.to_datetime(df["Ship Date"])

print("\nData types after conversion:")
print(df.dtypes)
print("\nMissing values remaining:", int(df.isnull().sum().sum()))

# ---- 2.4 Consistency check on categorical columns -------------------------
# Inspect category labels for typos / invalid values
print("\nRegion values before cleaning:")
print(df["Region"].value_counts())

# 3 rows have the invalid label "'NorthEast" (stray apostrophe). Check them:
print("\nRows with invalid Region label:")
print(df.loc[df["Region"] == "'NorthEast", ["Order ID", "City", "State", "Region"]])

# They are all in New York, and every other New York row is labelled 'East'
print("\nRegion labels used for State = New York:")
print(df.loc[df["State"] == "New York", "Region"].value_counts())

df["Region"] = df["Region"].replace("'NorthEast", "East")
print("\nRegion values after cleaning:")
print(df["Region"].value_counts())

# ---- 2.5 Check for duplicates ---------------------------------------------
print("\nFully duplicated rows:", df.duplicated().sum())
print("Duplicate Row IDs    :", df["Row ID"].duplicated().sum())
df = df.drop_duplicates()  # nothing is removed here, kept as a safeguard

print("\nFinal cleaned shape:", df.shape)

# ==========================================================================
# 3. VISUALIZATION
# ==========================================================================
section("3. VISUALIZATION")


# ---- 3.1 Bar charts: Category, Sub-Category, Segment, Region, State -------
def bar_sales_profit(by, filename, top_n=None, horizontal=False, figsize=(14, 5)):
    """Draw two bar charts side by side: total Sales and total Profit for one column.
    Negative profit bars are coloured red so loss-making groups stand out."""
    grouped = df.groupby(by)[["Sales", "Profit"]].sum().sort_values("Sales", ascending=False)
    if top_n:
        grouped = grouped.head(top_n)

    fig, axes = plt.subplots(1, 2, figsize=figsize)
    for ax, metric in zip(axes, ["Sales", "Profit"]):
        data = grouped[metric].sort_values(ascending=False)
        colours = [RED if v < 0 else BLUE for v in data] if metric == "Profit" else [BLUE] * len(data)
        if horizontal:
            # reverse so the largest bar appears at the top
            ax.barh(data.index[::-1], data.values[::-1], color=colours[::-1])
            ax.set_xlabel(f"Total {metric} ($)")
        else:
            ax.bar(data.index, data.values, color=colours)
            ax.set_ylabel(f"Total {metric} ($)")
            ax.tick_params(axis="x", rotation=30)
        suffix = f" (top {top_n} by Sales)" if top_n else ""
        ax.set_title(f"Total {metric} by {by}{suffix}")
    save_and_show(filename)

    # Also print the numbers behind the chart, with profit margin
    table = grouped.assign(Margin_pct=(grouped["Profit"] / grouped["Sales"] * 100).round(1)).round(0)
    print(f"\nSales / Profit by {by}:")
    print(table)


bar_sales_profit("Category", "bar_category.png", figsize=(12, 4.5))
bar_sales_profit("Sub-Category", "bar_subcategory.png", horizontal=True, figsize=(15, 7))
bar_sales_profit("Segment", "bar_segment.png", figsize=(12, 4.5))
bar_sales_profit("Region", "bar_region.png", figsize=(12, 4.5))
# 49 states would be unreadable on one chart, so show the top 15 by Sales
bar_sales_profit("State", "bar_state_top15.png", top_n=15, horizontal=True, figsize=(15, 7))

# Extra view for State: the 10 least profitable states
state_profit = df.groupby("State")["Profit"].sum().sort_values()
plt.figure(figsize=(10, 5))
worst10 = state_profit.head(10)
plt.barh(worst10.index[::-1], worst10.values[::-1], color=RED)
plt.title("10 least profitable states (total Profit)")
plt.xlabel("Total Profit ($)")
save_and_show("bar_state_worst10.png")
print(f"\nStates with negative total profit: {(state_profit < 0).sum()} of {len(state_profit)}")

# ---- 3.2 Scatter plots: Sales vs Profit, Discount vs Profit ---------------
# Sales vs Profit (left: all data, right: zoomed in, because most orders are small)
fig, axes = plt.subplots(1, 2, figsize=(15, 5.5))
sns.scatterplot(data=df, x="Sales", y="Profit", hue="Category", alpha=0.5, s=25, ax=axes[0])
axes[0].axhline(0, color="black", lw=0.8)
axes[0].set_title("Sales vs Profit (all orders)")

core = df[df["Sales"] <= df["Sales"].quantile(0.95)]
sns.scatterplot(data=core, x="Sales", y="Profit", hue="Category", alpha=0.5, s=25, ax=axes[1])
axes[1].axhline(0, color="black", lw=0.8)
axes[1].set_title("Sales vs Profit (zoomed: Sales up to 95th percentile)")
save_and_show("scatter_sales_vs_profit.png")

# Discount vs Profit (left: every order, right: average profit at each discount level)
fig, axes = plt.subplots(1, 2, figsize=(15, 5.5))
sns.scatterplot(data=df, x="Discount", y="Profit", alpha=0.35, s=25, color=BLUE, ax=axes[0])
axes[0].axhline(0, color="black", lw=0.8)
axes[0].set_title("Discount vs Profit")

avg_profit = df.groupby("Discount")["Profit"].mean()
axes[1].bar(avg_profit.index.astype(str), avg_profit.values,
            color=[RED if v < 0 else BLUE for v in avg_profit])
axes[1].axhline(0, color="black", lw=0.8)
axes[1].set_title("Average Profit per order line by Discount level")
axes[1].set_xlabel("Discount")
axes[1].set_ylabel("Average Profit ($)")
save_and_show("scatter_discount_vs_profit.png")

print("\nShare of order lines that lose money, by discount level (%):")
print((df.assign(Loss=df["Profit"] < 0).groupby("Discount")["Loss"].mean() * 100).round(1))

# ---- 3.3 Histograms: Sales, Profit, Quantity, Discount --------------------
# Sales and Profit are extremely skewed (a few huge orders), so their histograms
# focus on the 1st-99th percentile. The extremes are examined in Section 4.
fig, axes = plt.subplots(2, 2, figsize=(14, 9))

lo, hi = df["Sales"].quantile([0.01, 0.99])
sns.histplot(df.loc[df["Sales"].between(lo, hi), "Sales"], bins=40, kde=True,
             color=BLUE, ax=axes[0, 0])
axes[0, 0].set_title("Distribution of Sales (1st-99th percentile)")

lo, hi = df["Profit"].quantile([0.01, 0.99])
sns.histplot(df.loc[df["Profit"].between(lo, hi), "Profit"], bins=40, kde=True,
             color="#27ae60", ax=axes[0, 1])
axes[0, 1].axvline(0, color="black", lw=0.8)
axes[0, 1].set_title("Distribution of Profit (1st-99th percentile)")

sns.histplot(df["Quantity"], discrete=True, color="#e67e22", ax=axes[1, 0])
axes[1, 0].set_title("Distribution of Quantity")

sns.histplot(df["Discount"], bins=20, color="#8e44ad", ax=axes[1, 1])
axes[1, 1].set_title("Distribution of Discount")
save_and_show("histograms.png")

print("\nSkewness / median / mean of numeric variables:")
print(df[["Sales", "Profit", "Quantity", "Discount"]].agg(["skew", "median", "mean"]).round(2))

# ---- 3.4 Heatmap: correlation between numeric variables -------------------
corr = df[["Sales", "Profit", "Discount", "Quantity"]].corr()
plt.figure(figsize=(7, 5.5))
sns.heatmap(corr, annot=True, fmt=".2f", cmap="coolwarm", vmin=-1, vmax=1,
            square=True, linewidths=0.5)
plt.title("Correlation between numeric variables")
save_and_show("heatmap_correlation.png")
print("\nCorrelation matrix:")
print(corr.round(2))

# ==========================================================================
# 4. OUTLIER DETECTION
# ==========================================================================
section("4. OUTLIER DETECTION")

# Boxplots: one panel per variable because their scales are very different
fig, axes = plt.subplots(1, 3, figsize=(16, 5.5))
for ax, col, colour in zip(axes, ["Sales", "Profit", "Discount"], [BLUE, "#27ae60", "#8e44ad"]):
    sns.boxplot(y=df[col], color=colour, fliersize=2, ax=ax)
    ax.set_title(f"Boxplot of {col}")
save_and_show("boxplots_outliers.png")

# Count outliers with the IQR rule: a value is an outlier if it lies more than
# 1.5 x IQR below Q1 or above Q3 (the same rule the boxplot whiskers use)
rows = []
for col in ["Sales", "Profit", "Discount"]:
    q1, q3 = df[col].quantile([0.25, 0.75])
    iqr = q3 - q1
    lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
    n_out = int(((df[col] < lower) | (df[col] > upper)).sum())
    rows.append({"Variable": col, "Q1": q1, "Q3": q3, "Lower fence": lower,
                 "Upper fence": upper, "Outliers": n_out,
                 "Pct of rows": n_out / len(df) * 100})
print("\nOutliers by IQR rule:")
print(pd.DataFrame(rows).set_index("Variable").round(2))

# Look at the most extreme records to decide whether they are errors or genuine
cols = ["Order ID", "Sub-Category", "Sales", "Quantity", "Discount", "Profit"]
print("\nTop 5 orders by Sales:")
print(df.nlargest(5, "Sales")[cols])
print("\nTop 5 orders by Profit:")
print(df.nlargest(5, "Profit")[cols])
print("\nBottom 5 orders by Profit (largest losses):")
print(df.nsmallest(5, "Profit")[cols])

# Decision: these are genuine transactions (expensive Copiers, heavily discounted
# Machines), not data errors, so they are kept in the dataset.

# ==========================================================================
# KEY FINDINGS
# ==========================================================================
section("KEY FINDINGS")
print(f"""
- Dataset: {len(df):,} order lines | Total Sales ${df['Sales'].sum():,.0f} | \
Total Profit ${df['Profit'].sum():,.0f} | Margin {df['Profit'].sum() / df['Sales'].sum():.1%}
- Data quality: 11 missing Postal Codes (Burlington, VT) filled; 3 invalid Region
  labels ("'NorthEast") corrected to 'East'; no duplicates.
- Furniture earns a very low margin; Tables, Bookcases and Supplies lose money overall.
- Discount is the only variable clearly hurting Profit (correlation about -0.22);
  orders discounted 30% or more almost always lose money.
- Sales and Profit are heavily right-skewed; their outliers are genuine and were kept.
""")
print(f"Done. Charts saved in: {PLOT_DIR}")
