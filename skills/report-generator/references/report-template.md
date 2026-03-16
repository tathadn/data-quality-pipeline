# Report Template Reference

## Visual Identity

All reports from this pipeline share a consistent visual identity:

### Color Palette
- **Primary**: #1a56db — section headings, chart accents, table headers
- **Accent**: #10b981 — positive highlights, secondary charts
- **Warning**: #f59e0b — caution items, medium severity
- **Danger**: #ef4444 — critical items, high severity
- **Purple**: #8b5cf6 — tertiary charts and accents
- **Gray Dark**: #374151 — body text, subheadings
- **Gray Light**: #f3f4f6 — alternating table rows, backgrounds

### Typography
- Title: Helvetica-Bold, 24pt
- Heading 1: Helvetica-Bold, 16pt
- Heading 2: Helvetica-Bold, 13pt
- Body: Helvetica, 10pt
- Caption: Helvetica-Oblique, 8pt
- Table header: Helvetica-Bold, 9pt, white on primary background
- Table body: Helvetica, 9pt

### Chart Style
- White background, light gray grid lines at 30% opacity
- Chart palette: rotate through CHART_PALETTE for multi-series
- Always include a clear title (bold, 14pt) on every chart
- Include axis labels on all charts
- Use `bbox_inches='tight'` on every save to avoid clipping
- Target 150 DPI for crisp rendering without bloated file sizes

## Document Structure

### Audit Report (Data Auditor Skill)
1. Cover Page — health score gauge, dataset stats
2. Executive Summary — written analysis + score breakdown pie
3. Missing Values — bar chart + table
4. Severity Overview — heatmap + interpretation
5. Distributions — histogram grid + notes
6. Trends — trend lines (if applicable)
7. Detailed Findings — full issues table
8. Recommendations — prioritized action list

### Summary Report (Report Generator Skill)
1. Cover Page — overview bar chart, dataset stats
2. Executive Summary — written analysis + composition pie
3. Distributions — histogram grid + stats
4. Categories — top values bar charts
5. Correlations — heatmap (if 3+ numeric cols)
6. Trends — trend lines (if date column)
7. Column Profiles — detailed table
8. Findings & Methodology — insights + next steps

## Tone and Voice

- Professional but accessible
- A non-technical executive should understand the Executive Summary without help
- Use specific numbers: "Revenue increased 23% from Q1 to Q4" not "Revenue went up significantly"
- Avoid hedging unless genuinely uncertain
- Each insight should be 1-2 sentences with concrete data points
- Chart captions should describe what the reader should take away, not just what the chart shows

## Table Formatting

- Bold header row with primary color (#1a56db) background, white text
- Alternating row colors: white / #f3f4f6
- Right-aligned numbers, left-aligned text
- Consistent decimal places: 2 for percentages, 0 for counts, 1 for means
- Thin borders (#e5e7eb)

## Chart Captions

Every chart in the PDF should have a caption below it following this pattern:
- "Figure N: [What the chart shows]. [Key takeaway in one sentence]."
- Example: "Figure 3: Distribution of numeric columns. Revenue is right-skewed with a median of $42.50."

## Written Sections

Each written section (Executive Summary, interpretations) should:
- Start with a one-sentence overview
- Follow with 2-3 supporting observations with specific numbers
- End with an implication or recommendation where appropriate
- Use short paragraphs (3-4 sentences max)
