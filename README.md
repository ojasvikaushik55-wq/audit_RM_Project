# Research Data Audit Lab

A Streamlit application for a workplace dust exposure / respiratory symptoms / work performance research-methodology project. It is designed to inspect the supplied XLSX and CSV files and show reproducible checks and value provenance in plain language.

## Important scope limit

**This app does not and cannot prove that worker-level observations are real just by checking spreadsheets.** The supplied materials contain explicit terms such as `simulated` and `illustrative`. In particular, the results workbook describes an association as observed in simulated data, and multiple coding guides label variables or scores illustrative. This app surfaces those statements instead of hiding them. For a real-world data claim, the team needs primary collection evidence and a complete formula specification.

The app can establish narrower facts, such as:

- the exact file bytes inspected (SHA-256 fingerprints);
- the number of rows, columns, blank cells, and duplicate worker IDs;
- whether overlapping worker IDs and shared columns match across files;
- whether health symptom totals reproduce from the documented 10 binary items;
- recomputed zone means and Pearson correlations from the integrated table;
- which decimal is stored in which source file, sheet, and cell;
- whether Excel cells contain live formulas or hard-coded values;
- where the workbooks themselves use wording like simulated / illustrative.

## Run locally

Python 3.10 or later is recommended.

```bash
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL printed by Streamlit (usually `http://localhost:8501`).

## Deploy using GitHub + Streamlit Community Cloud

1. Create a GitHub repository and upload the contents of this folder, including `data/`.
2. Confirm `app.py`, `audit_engine.py`, `requirements.txt`, and `data/` are at the repository root.
3. In Streamlit Community Cloud, create an app from that repository and set the main file path to `app.py`.
4. Deploy. No secrets or API keys are needed.

Do not make the repository public if the files contain sensitive or identifiable participant information. The bundled versions currently use worker IDs rather than names, but your institution's research privacy rules still apply.

## Features

- Overview dashboard and zone summaries
- XLSX/CSV source inventory and table explorer
- Selected-cell decimal trace with stored value and source location
- Reproducible health-symptom-score calculation
- Data structure and score-range checks
- Cross-file reconciliation by worker ID and shared columns
- Recomputed zone descriptives and Pearson correlations
- Exact workbook phrases and SHA-256 fingerprints
- Markdown audit report and CSV check export
- Session uploads for revised files (an upload replaces a bundled file with the same filename)

## Formula policy

The engine never guesses a score formula to make a value look legitimate. It recomputes the health symptom score because the codebook defines it as the sum of 10 binary symptoms. For Exposure Index, Safety Practice Score, PPE Compliance, and Performance Impairment Score, the supplied files do not provide all the exact formulas/weights/rounding rules needed for independent reconstruction; the app identifies that gap.

## Repository layout

```text
research_data_audit_app/
├── app.py
├── audit_engine.py
├── requirements.txt
├── README.md
└── data/
    ├── Dataset1 ... Dataset8 workbooks
    ├── master / analysis / final-result workbooks
    └── Dataset4 Orange CSV
```

## Not a fraud detector

Statistical plausibility, a clean ID sequence, exact cross-file matches, or a passing score check are not proof that the underlying answers or measurements were collected. Treat this as an auditability and reproducibility tool, not a certificate of authenticity.
