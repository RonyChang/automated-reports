# Automated Reports
This repository contains my production-grade Python workflows for automated reporting.
It includes reusable pipelines to extract, validate, and transform data, generate
reports (tables, charts, and PDFs/HTML), and deliver them via email or Slack on
a schedule.

Highlights:
- Python-first: pandas/pyarrow for data, Jinja2/HTML/PDF for rendering.
- Reproducible pipelines: config-driven jobs with clear inputs/outputs.
- Quality gates: data validation checks and unit tests before distribution.
- Scheduling: runs via Airflow/Cron/GitHub Actions with environment-specific configs.
- Observability: structured logs and simple run summaries for quick troubleshooting.
- Delivery: pluggable notifiers (SMTP/Slack) with templated messages.

Typical structure:
- `pipelines/` reusable steps and job DAGs
- `jobs/` entrypoints for specific reports
- `configs/` YAML/JSON settings per report and environment
- `templates/` report templates (HTML/Jinja2)
- `outputs/` generated artifacts (CSV/XLSX/HTML/PDF)
- `tests/` unit and data validation tests
- `scripts/` CLI helpers

Quick start:
1) `python -m venv .venv && source .venv/bin/activate`
2) `pip install -r requirements.txt`
3) `python -m jobs.<report_name> --config configs/<report_name>.yaml`
   (or `make run REPORT=<report_name>`)

CI/CD:
- Lint + tests on PRs
- Scheduled runs for recurring reports
- Artifacts uploaded and notifications sent on success/failure
