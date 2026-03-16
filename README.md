# Data Quality Pipeline

A collection of Claude skills for automated data auditing and report generation.

## Skills

- **data-auditor** — Audits CSV/Excel files for data quality issues and produces a PDF report with charts plus an Excel workbook.
- **report-generator** — Generates a professional PDF summary report with charts and analysis from a clean dataset.

## Structure

```
data-quality-pipeline/
├── skills/
│   ├── data-auditor/       # Data quality audit skill
│   └── report-generator/   # Report generation skill
├── examples/
│   ├── raw-data/           # Sample input datasets
│   ├── audit-reports/      # Example audit outputs
│   └── final-reports/      # Example final report outputs
├── scripts/                # Utility scripts
└── docs/                   # Documentation and case studies
```

## Usage

Load a skill into Claude and upload a CSV or Excel file. Claude will invoke the appropriate skill based on your request.

## License

MIT
