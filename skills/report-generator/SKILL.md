---
name: report-generator
description: "Generates a professional PDF summary report with charts, figures, and written analysis from a clean CSV or Excel dataset, plus a companion Excel workbook with detailed statistics. Produces bar charts, pie charts, histograms, trend lines, and correlation analysis. Use this skill when the user asks for a data summary report, dataset overview, data profile report, analytics report, or wants a professional document summarizing their data — even if they just say 'summarize this data', 'create a report', or 'analyze this file'."
---

# Report Generator Skill

## Purpose

Transform a dataset into two professional deliverables:
1. A **PDF report** with charts, figures, tables, and written analysis — ready for stakeholders
2. An **Excel workbook** with detailed statistics and raw profile data — for analysts who want to dig deeper

Both outputs are generated every time.

## Dependencies

```python
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # REQUIRED — no display server
import matplotlib.pyplot as plt
import seaborn as sns
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak,
    Table, TableStyle, Image as RLImage, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os
import tempfile
```

## Color Palette

Same palette as the Data Auditor for visual consistency:
```python
COLORS = {
    'primary': '#1a56db',
    'primary_light': '#e8eefb',
    'accent': '#10b981',
    'accent_light': '#d1fae5',
    'warning': '#f59e0b',
    'danger': '#ef4444',
    'purple': '#8b5cf6',
    'gray_dark': '#374151',
    'gray_medium': '#6b7285',
    'gray_light': '#f3f4f6',
    'white': '#ffffff',
}
CHART_PALETTE = ['#1a56db', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#84cc16']

plt.rcParams.update({
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.edgecolor': '#e5e7eb',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'font.family': 'sans-serif',
    'font.size': 10,
    'figure.dpi': 150,
})
```

## When to Use

Activate this skill when the user:
- Uploads a CSV or Excel file and asks for a "report", "summary", or "overview"
- Says "summarize this data", "create a report", or "analyze this file"
- Wants a professional document to share with stakeholders or include in a presentation
- Has already audited/cleaned data and wants the final analysis deliverable
- Asks for "insights", "findings", or "what does this data tell us"

## Report Generation Procedure

### Step 1: Load and Analyze

```python
df = pd.read_csv(filepath)  # or pd.read_excel()
```

- Identify column types: numeric, categorical, datetime, boolean
- Compute summary statistics for all columns
- Detect date columns (try `pd.to_datetime` on string columns)
- Calculate completeness rate: `(non_null_values / total_values) * 100`
- Identify potential ID columns (all unique values) and exclude from analysis
- Detect correlations between numeric columns

### Step 2: Generate All Charts

Save every chart as a temporary PNG. Always use `plt.savefig(path, bbox_inches='tight', dpi=150)` then `plt.close()`.

#### Figure 1: Dataset Overview Bar Chart
A horizontal bar showing key metrics:
```python
fig, ax = plt.subplots(figsize=(8, 3))
metrics = ['Total Records', 'Total Fields', 'Numeric Fields', 'Categorical Fields', 'Completeness %']
values = [total_rows, total_cols, n_numeric, n_categorical, completeness_pct]
colors = [COLORS['primary'], COLORS['primary'], COLORS['accent'], COLORS['purple'], COLORS['warning']]
bars = ax.barh(metrics, values, color=colors, edgecolor='white', height=0.6)
# Add value labels on bars
for bar, val in zip(bars, values):
    ax.text(bar.get_width() + max(values)*0.02, bar.get_y() + bar.get_height()/2,
            f'{val:,.0f}' if val > 1 else f'{val:.1f}%', va='center', fontsize=10, fontweight='bold')
ax.set_title('Dataset Overview', fontsize=14, fontweight='bold')
ax.set_xlim(0, max(values) * 1.2)
plt.savefig(overview_path, bbox_inches='tight', dpi=150)
plt.close()
```

#### Figure 2: Data Composition Pie Chart
Pie chart showing the proportion of numeric vs categorical vs datetime columns:
```python
fig, ax = plt.subplots(figsize=(5, 5))
type_counts = [n_numeric, n_categorical, n_datetime, n_other]
type_labels = ['Numeric', 'Categorical', 'Datetime', 'Other']
# Filter out zero-count types
non_zero = [(l, c) for l, c in zip(type_labels, type_counts) if c > 0]
labels, counts = zip(*non_zero)
colors_pie = [COLORS['primary'], COLORS['accent'], COLORS['warning'], COLORS['purple']][:len(labels)]
wedges, texts, autotexts = ax.pie(counts, labels=labels, colors=colors_pie,
    autopct='%1.0f%%', startangle=90, textprops={'fontsize': 10},
    wedgeprops={'edgecolor': 'white', 'linewidth': 2})
ax.set_title('Column Type Distribution', fontsize=14, fontweight='bold')
plt.savefig(composition_path, bbox_inches='tight', dpi=150)
plt.close()
```

#### Figure 3: Distribution Histograms (Numeric Columns)
For up to 6 numeric columns, create a grid of histograms:
```python
numeric_cols = df.select_dtypes(include=[np.number]).columns
# Exclude likely ID columns (all unique values)
numeric_cols = [c for c in numeric_cols if df[c].nunique() < len(df) * 0.9]
plot_cols = numeric_cols[:6]
if len(plot_cols) > 0:
    n_rows = (len(plot_cols) + 2) // 3
    fig, axes = plt.subplots(nrows=n_rows, ncols=3, figsize=(12, 3.5 * n_rows))
    axes = np.array(axes).flatten()
    for i, col in enumerate(plot_cols):
        ax = axes[i]
        data = df[col].dropna()
        ax.hist(data, bins=min(30, max(10, len(data)//50)), 
                color=CHART_PALETTE[i % len(CHART_PALETTE)], edgecolor='white', alpha=0.85)
        # Add median line
        median_val = data.median()
        ax.axvline(median_val, color=COLORS['danger'], linestyle='--', linewidth=1.5)
        ax.set_title(f'{col}', fontsize=11, fontweight='bold')
        # Add stats annotation
        skew = data.skew()
        skew_label = 'Left-skewed' if skew < -0.5 else 'Right-skewed' if skew > 0.5 else 'Symmetric'
        ax.text(0.97, 0.95, f'Median: {median_val:,.1f}\n{skew_label}',
                transform=ax.transAxes, ha='right', va='top', fontsize=7,
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='#e5e7eb', alpha=0.9))
    for j in range(len(plot_cols), len(axes)):
        axes[j].set_visible(False)
    fig.suptitle('Value Distributions', fontsize=14, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(hist_path, bbox_inches='tight', dpi=150)
    plt.close()
```

#### Figure 4: Top Categories Bar Charts (Categorical Columns)
For up to 4 categorical columns, show the top 8 most frequent values:
```python
cat_cols = df.select_dtypes(include=['object', 'category']).columns
# Exclude likely ID or high-cardinality columns
cat_cols = [c for c in cat_cols if 2 < df[c].nunique() < 50][:4]
if len(cat_cols) > 0:
    fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(12, 8))
    axes = axes.flatten()
    for i, col in enumerate(cat_cols):
        ax = axes[i]
        top_values = df[col].value_counts().head(8)
        bars = ax.barh(top_values.index.astype(str), top_values.values,
                       color=CHART_PALETTE[i % len(CHART_PALETTE)], edgecolor='white')
        ax.set_title(f'{col} — Top Values', fontsize=11, fontweight='bold')
        ax.invert_yaxis()
        # Add count labels
        for bar, val in zip(bars, top_values.values):
            ax.text(bar.get_width() + max(top_values.values)*0.02, 
                    bar.get_y() + bar.get_height()/2, f'{val:,}', va='center', fontsize=8)
    for j in range(len(cat_cols), len(axes)):
        axes[j].set_visible(False)
    fig.suptitle('Top Categories', fontsize=14, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(categories_path, bbox_inches='tight', dpi=150)
    plt.close()
```

#### Figure 5: Correlation Heatmap (If 3+ Numeric Columns)
```python
numeric_for_corr = df[numeric_analysis_cols].dropna()
if len(numeric_analysis_cols) >= 3:
    corr_matrix = numeric_for_corr.corr()
    fig, ax = plt.subplots(figsize=(max(6, len(numeric_analysis_cols)*0.8), 
                                     max(5, len(numeric_analysis_cols)*0.7)))
    mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
    sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r',
                center=0, vmin=-1, vmax=1, ax=ax, linewidths=0.5,
                cbar_kws={'shrink': 0.8, 'label': 'Correlation'},
                annot_kws={'fontsize': 8})
    ax.set_title('Correlation Matrix', fontsize=14, fontweight='bold')
    plt.xticks(rotation=45, ha='right', fontsize=9)
    plt.yticks(fontsize=9)
    plt.savefig(corr_path, bbox_inches='tight', dpi=150)
    plt.close()
```

#### Figure 6: Trend Lines (Only If Date Column Exists)
```python
if date_col is not None:
    trend_cols = numeric_analysis_cols[:3]
    if len(trend_cols) > 0:
        fig, axes = plt.subplots(nrows=len(trend_cols), ncols=1,
                                  figsize=(10, 3.5 * len(trend_cols)), sharex=True)
        if len(trend_cols) == 1:
            axes = [axes]
        sorted_df = df.sort_values(date_col)
        for i, col in enumerate(trend_cols):
            ax = axes[i]
            ax.plot(sorted_df[date_col], sorted_df[col], 
                    color=CHART_PALETTE[i], alpha=0.3, linewidth=0.8)
            # Rolling average
            window = max(7, len(sorted_df) // 20)
            rolling = sorted_df[col].rolling(window=window, center=True).mean()
            ax.plot(sorted_df[date_col], rolling, color=CHART_PALETTE[i], 
                    linewidth=2.5, label=f'{window}-period moving avg')
            ax.set_ylabel(col, fontsize=10)
            ax.legend(fontsize=8, loc='upper left')
            # Shade min/max range
            ax.fill_between(sorted_df[date_col], sorted_df[col].min(), sorted_df[col],
                           alpha=0.05, color=CHART_PALETTE[i])
        axes[-1].set_xlabel('Date')
        fig.suptitle('Trends Over Time', fontsize=14, fontweight='bold', y=1.01)
        plt.tight_layout()
        plt.savefig(trend_path, bbox_inches='tight', dpi=150)
        plt.close()
```

### Step 3: Generate PDF Report

Structure the PDF as follows:

**Page 1: Cover Page**
- Report title: "Data Summary Report"
- Subtitle: Dataset filename
- Date generated
- Overview bar chart (Figure 1)
- Key stats: rows, columns, completeness rate

**Page 2: Executive Summary & Composition**
- 3-4 paragraph plain-language summary answering:
  - What is this dataset about? (infer from column names)
  - How large is it and what time period does it cover?
  - What are the most important findings?
- Data Composition pie chart (Figure 2)
- Key Metrics table

**Page 3: Distribution Analysis**
- Histogram grid (Figure 3)
- Written interpretation for each column:
  - Is the distribution skewed? Normal? Bimodal?
  - What are the key statistics (mean, median, std)?
  - Any notable outliers?

**Page 4: Categorical Analysis**
- Top Categories bar charts (Figure 4)
- Written notes:
  - Does any single value dominate?
  - How many unique values per column?

**Page 5: Correlation Analysis** (only if 3+ numeric columns)
- Correlation heatmap (Figure 5)
- Written notes on strong correlations (|r| > 0.7)
- Warn about potential multicollinearity

**Page 6: Trend Analysis** (only if date column exists)
- Trend line charts (Figure 6)
- Written interpretation:
  - Is the trend increasing, decreasing, or stable?
  - Any seasonal patterns?
  - Any sudden spikes or drops?

**Page 7: Detailed Column Profiles**
- Table with one row per column:
  - Column name, type, non-null count, unique values
  - For numeric: mean, median, std, min, max
  - For categorical: top value, top value frequency

**Page 8: Notable Findings & Methodology**
- 3-5 key insights with specific numbers
- Brief methodology paragraph
- Suggestions for further analysis

**PDF Formatting Rules:**
- Page size: US Letter (8.5" x 11")
- Margins: 0.75 inches
- Title: Helvetica-Bold, 24pt, #1a56db
- Heading 1: Helvetica-Bold, 16pt, #374151
- Heading 2: Helvetica-Bold, 13pt, #374151
- Body: Helvetica, 10pt, #4b5563, justified
- Tables: #1a56db header with white text, alternating row colors
- Page numbers: bottom center
- Header line: thin #1a56db line below "Data Summary Report" on every page

### Step 4: Generate Excel Workbook

Create a companion Excel file with four sheets:

#### Sheet 1: "Overview"
Key metrics in a clean two-column layout:
Dataset Name, Total Rows, Total Columns, Completeness Rate, Numeric Columns, Categorical Columns, Date Range (if applicable), Generated On.

#### Sheet 2: "Column Profiles"
One row per column with: Column Name, Data Type, Non-Null Count, Null Count, Null %, Unique Values, Top Value, Top Value %, Mean (numeric), Median (numeric), Std Dev (numeric), Min (numeric), Max (numeric).

#### Sheet 3: "Correlations"
The full correlation matrix for numeric columns. Apply conditional formatting: strong positive (> 0.7) in green, strong negative (< -0.7) in red.

#### Sheet 4: "Top Categories"
For each categorical column: Column Name, Value, Count, Percentage.
Include top 10 values per column.

**Excel formatting:** Same as Data Auditor (bold headers, alternating rows, auto-width, filters enabled).

### Step 5: Deliver Both Files

- Save PDF as: `report_[dataset_name]_[YYYY-MM-DD].pdf`
- Save Excel as: `report_data_[dataset_name]_[YYYY-MM-DD].xlsx`
- Present both files to the user
- Provide a verbal summary: the 2-3 most interesting findings

## Edge Cases

- **Empty dataset**: Return error, no report
- **Single column**: Generate report but simplify layout (no correlation, no multi-chart pages)
- **All categorical**: Skip histograms, correlation, and trend pages
- **All numeric**: Skip categorical analysis page
- **No date column**: Skip trend analysis page, note in report
- **Very wide (> 50 columns)**: Profile first 20 in detail, summarize rest
- **Very large (> 100k rows)**: Sample 10k rows for charts, use full data for stats where feasible
- **Fewer than 3 numeric columns**: Skip correlation heatmap

## What NOT To Do

- Do NOT make causal claims — describe correlations and patterns only
- Do NOT skip any applicable chart — if the data supports it, include it
- Do NOT use `plt.show()` — always `plt.savefig()` then `plt.close()`
- Do NOT generate only text — PDF with figures + Excel are always required
- Do NOT forget `matplotlib.use('Agg')` at the top
- Do NOT use jargon in the Executive Summary without explanation
- Do NOT include raw data dumps — summarize and visualize instead
