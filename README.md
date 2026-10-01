# Exploratory Data Analysis on Superstore Sales Data

An end-to-end exploratory data analysis (EDA) of the Superstore sales dataset (9,994 order lines, 2016 to 2019) in Python, using **Pandas**, **Matplotlib** and **Seaborn**.

## What the script does

1. **Data loading**: loads the `Orders` sheet; shows the first 10 rows, shape, column names and basic info.
2. **Data cleaning**: checks missing values, data types and duplicates.
   - 11 missing Postal Codes (all Burlington, VT) filled with 05401.
   - 3 rows with an invalid Region label (`'NorthEast`) corrected to `East`.
   - No duplicate rows found.
3. **Visualization**:
   - Bar charts for Category, Sub-Category, Segment, Region and State
   - Scatter plots: Sales vs Profit and Discount vs Profit
   - Histograms for Sales, Profit, Quantity and Discount
   - Correlation heatmap
4. **Outlier detection**: boxplots and IQR-rule counts for Sales, Profit and Discount.

## Key findings

- Total sales of about $2.30M and total profit of about $286K (12.5% margin).
- Furniture sells almost as much as Technology but earns only a 2.5% margin. Tables, Bookcases and Supplies lose money overall.
- 10 of 49 states have negative total profit, led by Texas.
- Discount is the only variable clearly hurting profit (correlation of about -0.22). Orders discounted 30% or more almost always lose money.
- Sales and Profit are heavily right-skewed. The outliers are genuine transactions (for example, high-value Copiers and heavily discounted Machines), so they were kept.

## Project structure

```
.
├── superstore_eda.py                 # the full EDA script
├── sample - superstore (2).xlsx      # dataset
├── eda_plots/                        # charts saved by the script
└── README.md
```

## How to run

```bash
pip install pandas matplotlib seaborn openpyxl
python superstore_eda.py
```

Keep the Excel file in the same folder as the script. Charts open in windows one at a time; set `SHOW_PLOTS = False` near the top of the script to only save them to `eda_plots/`.
