---
name: data-auditor
description: "Audits CSV and Excel files for data quality issues and generates a professional PDF report with charts and figures, plus a detailed Excel workbook. Includes bar charts, pie charts, heatmaps, histograms, and trend analysis. Use this skill whenever the user uploads a dataset and asks for a quality check, data audit, data cleaning assessment, data profiling, or wants to know what's wrong with their data — even if they don't use the word 'audit'. Also trigger when users ask for a 'data health check', 'data quality report', or 'what should I clean in this data'."
---

# Data Auditor Skill

## Purpose

Analyze any uploaded CSV or Excel dataset and produce:
1. A **professional PDF report** with charts, figures, and written analysis
2. A **detailed Excel workbook** with raw audit findings for further exploration

Both outputs are generated every time. The PDF is the primary deliverable for stakeholders; the Excel is the detailed backup for analysts.

## Dependencies

Use these Python libraries (install if needed):
- `pandas` — data loading and analysis
- `matplotlib` — chart generation
- `seaborn` — statistical visualizations and heatmaps
- `numpy` — numerical computations
- `reportlab` — PDF generation with embedded figures
- `openpyxl` — Excel workbook generation

Import setup at the top of every script:
```python
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend — REQUIRED
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak,
    Table, TableStyle, Image as RLImage
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from datetime import datetime
import os
import tempfile
```

## Color Palette

Use this consistent palette across ALL charts and the PDF:
```python
COLORS = {
    'primary': '#1a56db',
    'accent': '#10b981',
    'warning': '#f59e0b',
    'danger': '#ef4444',
    'purple': '#8b5cf6',
    'gray_dark': '#374151',
    'gray_medium': '#6b7285',
    'gray_light': '#f3f4f6',
    'white': '#ffffff',
}

SEVERITY_COLORS = {
    'High': '#ef4444',
    'Medium': '#f59e0b',
    'Low': '#10b981',
}

# For matplotlib charts
CHART_PALETTE = ['#1a56db', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#84cc16']
```

Set matplotlib defaults:
```python
plt.rcParams.update({
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.edgecolor': '#e5e7eb',
    'axes.grid': True,
    'grid.alpha': 0.3,
    'grid.color': '#e5e7eb',
    'font.family': 'sans-serif',
    'font.size': 10,
    'axes.titlesize': 13,
    'axes.titleweight': 'bold',
    'axes.labelsize': 10,
    'figure.dpi': 150,
})
```

## When to Use

Activate this skill when the user:
- Uploads a CSV, TSV, or Excel file and asks about its quality
- Requests a "data audit", "data check", "data profile", or "quality assessment"
- Asks "what's wrong with this data?" or "is this data clean?"
- Wants to know if a dataset is ready for analysis
- Asks for a "data health check" or "data quality report"

## Audit Procedure

### Step 1: Load and Inspect
- Load the uploaded file using pandas
- Record: total rows, total columns, column names, inferred dtypes
- Separate columns into categories: numeric, categorical, datetime, boolean
- Display a brief preview (first 5 rows) for context

### Step 2: Run Quality Checks

Perform ALL of the following checks on every applicable column:

#### Check 1: Missing Values
- Count null, NaN, empty string, and whitespace-only values per column
- Calculate missing percentage: `(missing_count / total_rows) * 100`
- Severity: **High** if > 20%, **Medium** if 5-20%, **Low** if < 5%

#### Check 2: Duplicate Rows
- Count exact duplicate rows across all columns
- Count near-duplicates: rows that match on all columns except one
- Severity: **High** if > 5% duplicates, **Medium** if 1-5%, **Low** if < 1%

#### Check 3: Data Type Validation
- For each column, check if values match the expected type
- Flag columns where > 10% of non-null values don't match the dominant type
- Common issues: numbers stored as strings, dates stored as strings, mixed numeric/text
- Severity: **High** if > 20%, **Medium** if 10-20%, **Low** if < 10%

#### Check 4: Outlier Detection (Numeric Columns Only)
- Calculate IQR for each numeric column
- Flag values below Q1 - 1.5*IQR or above Q3 + 1.5*IQR
- Report count and percentage of outliers per column
- Severity: **High** if > 10%, **Medium** if 3-10%, **Low** if < 3%
- Skip columns with fewer than 10 non-null numeric values

#### Check 5: Date Format Consistency
- Identify columns that contain date-like values
- Check for mixed formats (e.g., MM/DD/YYYY vs DD-MM-YYYY vs YYYY-MM-DD)
- Severity: **High** if > 2 formats, **Medium** if 2 formats, **Low** if consistent

#### Check 6: Text Casing Inconsistencies
- For text columns, categorize values: ALL CAPS, all lower, Title Case, mixed
- Flag columns where no single pattern covers > 80% of values
- Severity: **Medium** if inconsistent, **Low** otherwise

#### Check 7: Whitespace Issues
- Check for leading/trailing spaces in string values
- Check for double spaces within values
- Severity: **Medium** if > 10% affected, **Low** if < 10%

#### Check 8: Column Name Quality
- Check for: spaces, special characters, inconsistent casing
- Check for unnamed columns (e.g., "Unnamed: 0")
- Severity: **Low** for naming issues, **Medium** if unnamed columns exist

### Step 3: Calculate Health Score

Compute an overall score out of 100:

| Category              | Weight | Scoring                                    |
|-----------------------|--------|--------------------------------------------|
| Completeness          | 30%    | 100 - (avg missing % across all columns)   |
| Uniqueness            | 15%    | 100 - (duplicate row %)                    |
| Type Consistency      | 20%    | 100 - (avg type mismatch % across columns) |
| Outlier Reasonability | 15%    | 100 - (avg outlier % across numeric cols)  |
| Format Consistency    | 10%    | 100 if all dates consistent, else penalize |
| Text Quality          | 10%    | 100 - (avg whitespace/casing issue %)      |

Interpretation:
- 90-100: Excellent — minimal cleaning needed
- 70-89: Good — some issues to address
- 50-69: Fair — significant cleaning required
- Below 50: Poor — major quality concerns

### Step 4: Generate Charts

Save all charts as temporary PNG files for embedding in the PDF. Use `plt.savefig()` with `bbox_inches='tight'` and `dpi=150`.

#### Figure 1: Health Score Gauge
Create a donut/ring chart showing the overall health score:
```python
fig, ax = plt.subplots(figsize=(4, 4))
score = health_score  # 0-100
colors_gauge = [COLORS['primary'] if score >= 70 else COLORS['warning'] if score >= 50 else COLORS['danger'], '#e5e7eb']
ax.pie([score, 100 - score], colors=colors_gauge, startangle=90, counterclock=False,
       wedgeprops={'width': 0.3, 'edgecolor': 'white', 'linewidth': 2})
ax.text(0, 0, f'{score}', fontsize=36, fontweight='bold', ha='center', va='center', color=COLORS['gray_dark'])
ax.text(0, -0.15, 'out of 100', fontsize=10, ha='center', va='center', color=COLORS['gray_medium'])
ax.set_title('Data Health Score', fontsize=14, fontweight='bold', pad=20)
plt.savefig(gauge_path, bbox_inches='tight', dpi=150)
plt.close()
```

#### Figure 2: Missing Values Bar Chart
Horizontal bar chart showing missing percentage per column:
```python
fig, ax = plt.subplots(figsize=(8, max(4, len(columns_with_missing) * 0.4)))
# Sort by missing percentage descending
# Color bars by severity: red > 20%, orange 5-20%, green < 5%
# Add percentage labels at the end of each bar
# Add a vertical dashed line at 20% and 5% thresholds
ax.set_xlabel('Missing Values (%)')
ax.set_title('Missing Values by Column', fontsize=14, fontweight='bold')
plt.savefig(missing_path, bbox_inches='tight', dpi=150)
plt.close()
```
Only include columns that have at least some missing values. If no columns have missing values, skip this chart and note "No missing values detected" in the report.

#### Figure 3: Health Score Breakdown Pie Chart
Pie chart showing the weighted contribution of each quality category:
```python
fig, ax = plt.subplots(figsize=(6, 6))
categories = ['Completeness', 'Uniqueness', 'Type Consistency', 'Outlier Reasonability', 'Format Consistency', 'Text Quality']
weights = [30, 15, 20, 15, 10, 10]
scores_per_category = [...]  # Each category's score * weight / 100
colors_pie = CHART_PALETTE[:6]
explode = [0.03] * 6
wedges, texts, autotexts = ax.pie(scores_per_category, labels=categories, colors=colors_pie,
    autopct='%1.0f%%', startangle=90, explode=explode,
    textprops={'fontsize': 9})
ax.set_title('Health Score Breakdown by Category', fontsize=14, fontweight='bold')
plt.savefig(pie_path, bbox_inches='tight', dpi=150)
plt.close()
```

#### Figure 4: Issue Severity Heatmap
A grid heatmap with columns on x-axis, check types on y-axis, cells colored by severity:
```python
fig, ax = plt.subplots(figsize=(max(8, len(columns) * 0.6), 5))
# Create a matrix: rows = check types, columns = dataset columns
# Values: 0 = no issue, 1 = Low, 2 = Medium, 3 = High
# Use seaborn heatmap with custom colormap
cmap = sns.color_palette([COLORS['white'], COLORS['accent'], COLORS['warning'], COLORS['danger']])
# OR use ListedColormap
from matplotlib.colors import ListedColormap
severity_cmap = ListedColormap(['#f3f4f6', '#d1fae5', '#fef3c7', '#fee2e2'])
sns.heatmap(severity_matrix, ax=ax, cmap=severity_cmap, vmin=0, vmax=3,
    xticklabels=column_names, yticklabels=check_names,
    linewidths=1, linecolor='white', cbar=False,
    annot=severity_labels, fmt='s')  # severity_labels = matrix of "H"/"M"/"L"/""
ax.set_title('Issue Severity Heatmap', fontsize=14, fontweight='bold')
plt.xticks(rotation=45, ha='right', fontsize=8)
plt.savefig(heatmap_path, bbox_inches='tight', dpi=150)
plt.close()
```
If the dataset has more than 20 columns, only show the 20 columns with the most issues.

#### Figure 5: Distribution Histograms (Numeric Columns)
For each numeric column (up to 6), create a subplot histogram:
```python
numeric_cols = df.select_dtypes(include=[np.number]).columns[:6]
n_cols = len(numeric_cols)
if n_cols > 0:
    fig, axes = plt.subplots(nrows=2, ncols=3, figsize=(12, 7))
    axes = axes.flatten()
    for i, col in enumerate(numeric_cols):
        ax = axes[i]
        data = df[col].dropna()
        ax.hist(data, bins=30, color=CHART_PALETTE[i % len(CHART_PALETTE)], 
                edgecolor='white', alpha=0.85)
        # Add vertical lines for Q1, median, Q3
        q1, median, q3 = data.quantile([0.25, 0.5, 0.75])
        ax.axvline(median, color=COLORS['danger'], linestyle='--', linewidth=1.5, label=f'Median: {median:.1f}')
        # Mark IQR outlier boundaries
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        ax.axvline(lower, color=COLORS['warning'], linestyle=':', linewidth=1, alpha=0.7)
        ax.axvline(upper, color=COLORS['warning'], linestyle=':', linewidth=1, alpha=0.7)
        ax.set_title(col, fontsize=11, fontweight='bold')
        ax.legend(fontsize=7)
    # Hide unused subplots
    for j in range(n_cols, len(axes)):
        axes[j].set_visible(False)
    fig.suptitle('Value Distributions (Numeric Columns)', fontsize=14, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(hist_path, bbox_inches='tight', dpi=150)
    plt.close()
```
Skip this figure entirely if the dataset has no numeric columns.

#### Figure 6: Trend Lines (Only If Date Column Exists)
If a date/datetime column is detected:
```python
date_col = detected_date_column
numeric_cols_for_trend = df.select_dtypes(include=[np.number]).columns[:3]
if date_col and len(numeric_cols_for_trend) > 0:
    fig, axes = plt.subplots(nrows=len(numeric_cols_for_trend), ncols=1, 
                              figsize=(10, 3.5 * len(numeric_cols_for_trend)), sharex=True)
    if len(numeric_cols_for_trend) == 1:
        axes = [axes]
    for i, col in enumerate(numeric_cols_for_trend):
        ax = axes[i]
        # Sort by date, resample if needed (daily/weekly/monthly based on date range)
        sorted_df = df.sort_values(date_col)
        ax.plot(sorted_df[date_col], sorted_df[col], color=CHART_PALETTE[i], alpha=0.4, linewidth=0.8)
        # Add rolling average
        window = max(7, len(sorted_df) // 20)
        rolling = sorted_df[col].rolling(window=window, center=True).mean()
        ax.plot(sorted_df[date_col], rolling, color=CHART_PALETTE[i], linewidth=2, label=f'{window}-period avg')
        ax.set_ylabel(col, fontsize=10)
        ax.legend(fontsize=8)
    axes[-1].set_xlabel('Date')
    fig.suptitle('Trends Over Time', fontsize=14, fontweight='bold', y=1.01)
    plt.tight_layout()
    plt.savefig(trend_path, bbox_inches='tight', dpi=150)
    plt.close()
```
Skip entirely if no date column is detected. Note "No time-series data detected" in the report.

### Step 5: Generate PDF Report

Use `reportlab` to create a professional multi-page PDF. Structure:

**Page 1: Cover Page**
- Title: "Data Quality Audit Report"
- Subtitle: Dataset filename
- Date generated
- Health Score gauge chart (Figure 1) centered
- Summary stats: total rows, columns, issues found

**Page 2: Executive Summary**
- 3-4 paragraph written summary of findings
- Health Score Breakdown pie chart (Figure 3)
- Key stats table (critical issues, warnings, info items)

**Page 3: Missing Values Analysis**
- Missing Values bar chart (Figure 2)
- Table listing columns with highest missing rates
- Written interpretation

**Page 4: Issue Severity Overview**
- Severity Heatmap (Figure 4)
- Written interpretation of hotspots

**Page 5: Distribution Analysis**
- Histograms grid (Figure 5)
- Written notes on skewness, outliers detected
- Skip this page if no numeric columns

**Page 6: Trend Analysis** (only if date column exists)
- Trend line charts (Figure 6)
- Written notes on observed patterns
- Skip this page if no date column

**Page 7: Detailed Findings Table**
- Full table of all issues found, sorted by severity
- Columns: Column Name, Check Type, Severity, Finding, Affected Rows, Percentage

**Page 8: Recommendations**
- Prioritized list of recommended actions
- Each recommendation includes: what to fix, why, expected impact on health score

**PDF Formatting Rules:**
- Page size: US Letter (8.5 x 11 inches)
- Margins: 0.75 inches
- Title font: Helvetica-Bold, 24pt, color #1a56db
- Heading font: Helvetica-Bold, 16pt, color #374151
- Body font: Helvetica, 10pt, color #4b5563
- Table headers: white text on #1a56db background
- Table rows: alternating white / #f3f4f6
- Page numbers: bottom center
- Header: "Data Quality Audit Report" on every page after cover

### Step 6: Generate Excel Workbook

In addition to the PDF, produce an Excel file with three sheets:

#### Sheet 1: "Summary"
Rows: Dataset Name, Total Rows, Total Columns, Health Score, Health Rating, Critical Issues, Warnings, Info Items, Columns With Missing Data, Duplicate Rows, Date Generated.

#### Sheet 2: "Details"
Columns: Column Name, Check Type, Severity, Finding, Affected Rows, Percentage, Example Values.
Sort by Severity (High first), then Affected Rows (descending).

#### Sheet 3: "Recommendations"
Columns: Priority, Issue, Recommendation, Affected Columns, Estimated Impact.

**Excel Formatting:**
- Bold headers with #1a56db background, white text
- Alternating row colors
- Column auto-width
- Conditional formatting: red fill for High severity, orange for Medium, green for Low
- Freeze top row, enable filters

### Step 7: Deliver Both Files

- Save PDF as: `audit_report_[dataset_name]_[YYYY-MM-DD].pdf`
- Save Excel as: `audit_data_[dataset_name]_[YYYY-MM-DD].xlsx`
- Present both files to the user
- Provide a brief verbal summary: health score, top 3 issues, and #1 recommendation

## Edge Cases

- **Empty file**: Return an error message, generate no report
- **Single row**: Run checks but note "limited statistical validity" on every chart
- **All values missing in a column**: Flag as Critical, recommend dropping
- **No numeric columns**: Skip histograms and outlier detection, note in report
- **No date columns**: Skip trend analysis page entirely
- **Very large files (> 100k rows)**: Sample 10,000 rows, note sampling on cover page
- **Fewer than 3 columns**: Reduce heatmap size, adjust layout
- **Non-CSV/XLSX files**: Inform user this skill only supports CSV and Excel

## What NOT To Do

- Do NOT modify or clean the original dataset — audit only
- Do NOT skip chart generation — every applicable figure must be included
- Do NOT use `plt.show()` — always use `plt.savefig()` then `plt.close()`
- Do NOT generate charts for empty data — skip with a note instead
- Do NOT output only text — both PDF and Excel are always required
- Do NOT forget `matplotlib.use('Agg')` — will fail without display server
