from pathlib import Path
import json
import io
import pandas as pd
import numpy as np
import streamlit as st

from audit_engine import (
    source_files_from_directory, analyze_files, key_tables, table_label, find_col,
    run_integrity_checks, reconcile_tables, numeric_decimal_summary, trace_selected_cell,
    recompute_zone_summary, correlation_summary, report_markdown,
)

st.set_page_config(page_title="Research Data Audit Lab", page_icon="🔎", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
:root { --ink:#152238; --muted:#667085; --line:#E4E7EC; --soft:#F7F9FC; }
.block-container {max-width: 1500px; padding-top: 1.2rem; padding-bottom: 3rem;}
.hero { padding: 1.35rem 1.55rem; border: 1px solid #dce4ef; border-radius: 18px; background: linear-gradient(120deg,#f7faff 0%,#ffffff 64%,#f2f8f8 100%); margin-bottom: 1rem; }
.hero h1 { margin: 0 0 .25rem 0; color: #16263f; font-size: 2.05rem; letter-spacing:-.03em; }
.hero p { margin: 0; color: #475467; font-size: 1.02rem; }
.small-muted {color:#667085;font-size:.88rem;}
.info-card {border:1px solid #e4e7ec;border-radius:12px;padding:1rem 1.05rem;background:#fff;height:100%;}
.info-card h4 {margin-top:0;margin-bottom:.35rem;color:#1d2939;}
.tag {display:inline-block;padding:.23rem .55rem;border:1px solid #d0d5dd;border-radius:999px;color:#344054;font-size:.78rem;margin-right:.3rem;}
[data-testid="stMetric"] {background:#fff;border:1px solid #e4e7ec;border-radius:12px;padding:12px 14px;}
div[data-testid="stAlert"] {border-radius:12px;}
</style>
<div class="hero">
  <h1>🔎 Research Data Audit Lab</h1>
  <p>A reproducible, plain-language audit of worker-level files, decimals, score formulas, cross-file agreement, and evidence of data provenance.</p>
</div>
""", unsafe_allow_html=True)

DATA_DIR = Path(__file__).parent / "data"

with st.sidebar:
    st.markdown("## Audit controls")
    st.caption("Start with the bundled project files, then add revised copies if needed.")
    use_bundled = st.checkbox("Include bundled datasets", value=True)
    uploads = st.file_uploader("Add / replace datasets", type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                               help="An upload with the same filename replaces the bundled file with that filename for this audit run.")
    st.divider()
    st.markdown("### What this app can establish")
    st.markdown("- Whether records and IDs are internally consistent\n- Whether shared values agree across files\n- Whether a documented score formula reproduces\n- Where a decimal is stored and how it is represented")
    st.markdown("### What it cannot establish")
    st.markdown("- Whether a person actually gave a response\n- Whether a measurement happened in the workplace\n- Who entered or generated a value, without primary records")

file_objects = source_files_from_directory(DATA_DIR) if use_bundled else []
for item in uploads or []:
    file_objects.append({"name": item.name, "bytes": item.getvalue(), "origin": "Uploaded for this session"})

if not file_objects:
    st.warning("No files are loaded. Enable bundled datasets or upload one or more project workbooks/CSVs.")
    st.stop()

with st.spinner("Reading files, extracting source tables, and computing audit checks…"):
    analysis = analyze_files(file_objects)
files = analysis["files"]
tables = analysis["tables"]
file_lookup = {o["name"]: o for o in file_objects}
roles = key_tables(tables)
checks = run_integrity_checks(tables, files)
checks_df = pd.DataFrame(checks)
zone_summary = recompute_zone_summary(tables)
corr_summary = correlation_summary(tables)

# Banner based on explicit workbook language, never based on guesses from statistical patterns.
if analysis["evidence"]:
    st.warning("**Provenance caution:** the supplied files contain wording such as “simulated” and “illustrative.” The app will show those exact locations. Internal consistency cannot convert simulated or illustrative values into verified field observations.")
else:
    st.info("No explicit simulated/illustrative phrase was detected by the built-in text scan. That is not proof that the data are genuine.")

# Headline status metrics
master_table = roles.get("master") or roles.get("analysis")
worker_n = len(master_table["df"]) if master_table else 0
master_df = master_table["df"] if master_table else pd.DataFrame()
unique_workers = int(master_df["worker_id"].nunique()) if "worker_id" in master_df.columns else 0
pass_n = int((checks_df["status"] == "PASS").sum()) if not checks_df.empty else 0
fail_n = int((checks_df["status"] == "FAIL").sum()) if not checks_df.empty else 0

home, explore, decimals, integrity, evidence_tab, guide = st.tabs([
    "Overview", "Explore datasets", "Decimal trace", "Integrity & reconciliation", "Provenance evidence", "Methods & team map"
])

with home:
    st.subheader("Audit overview")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Files inspected", len(files))
    m2.metric("Worker-level tables", len(tables))
    m3.metric("Workers in selected integrated table", worker_n, help="Prefers the master dataset; falls back to the analysis dataset.")
    m4.metric("Checks passing", pass_n, help="A passing check is limited to the exact condition stated; it is not a certificate of authenticity.")
    c1, c2 = st.columns([1.05, 1])
    with c1:
        st.markdown("### What the audit says")
        st.markdown("This app separates three different questions that should not be confused:")
        st.markdown("**1. Does the file hold together?** IDs, missingness, score ranges, duplicate records, and cross-file matching.\n\n**2. Can a decimal be reproduced?** A formula is recomputed only where the workbooks document it and the required inputs are present.\n\n**3. Did this value come from a real worker or workplace measurement?** That requires primary collection records. Spreadsheet checks alone cannot prove it.")
        if fail_n:
            st.error(f"{fail_n} check(s) failed. Open Integrity & reconciliation to see the exact records or fields.")
        else:
            st.info("Passing structural checks mean that the tested structure is consistent—not that the data is authentic.")
    with c2:
        st.markdown("### Exposure-zone overview (recomputed)")
        if not zone_summary.empty:
            chart_cols = [c for c in ["Mean Exposure Index", "Mean Health Score", "Mean Safety Score", "Mean Performance Impairment"] if c in zone_summary.columns]
            if "Exposure Zone" in zone_summary.columns and chart_cols:
                chart = zone_summary.set_index("Exposure Zone")[chart_cols]
                st.bar_chart(chart, height=285)
            st.dataframe(zone_summary.round(3), use_container_width=True, hide_index=True)
        else:
            st.info("No integrated worker-level table could be identified in the files loaded.")
    st.markdown("### First provenance findings")
    if analysis["evidence"]:
        ev_df = pd.DataFrame(analysis["evidence"])
        ev_df = ev_df[["file", "sheet", "cell", "text"]]
        st.dataframe(ev_df, use_container_width=True, hide_index=True)
    else:
        st.caption("No wording hits found. Use the Provenance evidence tab for file fingerprints and limitations.")

with explore:
    st.subheader("Dataset explorer")
    st.write("Every table below is extracted from a source file. The app preserves the original file in the bundle and displays normalized column names for cross-file checks.")
    if not tables:
        st.warning("No worker-level tables found. Upload a workbook with a header containing a Worker ID column.")
    else:
        labels = [table_label(t) for t in tables]
        choice = st.selectbox("Choose a worker-level table", labels, key="explore_table")
        table = tables[labels.index(choice)]
        st.caption(f"Source: **{table['file']}** · Sheet: **{table['sheet']}** · Header is Excel row {table['header_row'] + 1 if table['source_kind'] == 'Excel' else 1}.")
        df = table["df"]
        f1, f2, f3, f4 = st.columns(4)
        f1.metric("Rows", len(df))
        f2.metric("Columns", len(df.columns))
        f3.metric("Missing cells", int(df.isna().sum().sum()))
        f4.metric("Duplicate worker IDs", int(df["worker_id"].duplicated().sum()) if "worker_id" in df else 0)
        search = st.text_input("Filter rows by worker ID or text", placeholder="e.g., W017 or High")
        shown = df.copy()
        if search:
            mask = shown.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
            shown = shown[mask]
        st.dataframe(shown, use_container_width=True, height=430, hide_index=True)
        csv_bytes = shown.to_csv(index=False).encode("utf-8")
        st.download_button("Download this extracted table as CSV", csv_bytes, file_name=f"{table['file'].rsplit('.',1)[0]}_{table['sheet'].replace(' ','_')}.csv", mime="text/csv")
    st.markdown("### Source-file inventory")
    inv = []
    for f in files:
        inv.append({"File": f["name"], "Source": f["origin"], "Size (KB)": round(f["size_bytes"]/1024, 1),
                    "Sheets": len(f["sheets"]), "Worker tables": len(f["tables"]), "Formula cells": f["formula_count"],
                    "SHA-256 (first 16 chars)": f["sha256"][:16], "Read status": "OK" if not f["load_error"] else f["load_error"]})
    st.dataframe(pd.DataFrame(inv), use_container_width=True, hide_index=True)

with decimals:
    st.subheader("Decimal trace: where a value lives and what supports it")
    st.write("Select a numeric field and row. For documented formulas, the app recomputes from the original inputs. For undocumented composites it records the source cell and explains exactly what is missing, instead of inventing a formula.")
    if not tables:
        st.warning("No worker-level tables available.")
    else:
        tlabels = [table_label(t) for t in tables]
        chosen_label = st.selectbox("Source table", tlabels, key="decimal_table")
        t = tables[tlabels.index(chosen_label)]
        ddf = t["df"]
        numeric_cols = [c for c in ddf.columns if pd.to_numeric(ddf[c], errors="coerce").notna().sum() > 0]
        default_idx = next((i for i,c in enumerate(numeric_cols) if any(k in c for k in ["score", "index", "tenure", "exposure", "productivity"])), 0)
        selected_col = st.selectbox("Numeric column", numeric_cols, index=default_idx if numeric_cols else None, key="decimal_column")
        st.caption("A decimal is not evidence of fabrication or authenticity by itself. Means, correlations, averages, weighted scores, and raw entries can all contain decimals for different reasons.")
        summary = numeric_decimal_summary(ddf[selected_col])
        a,b,c,d = st.columns(4)
        a.metric("Numeric values", summary["numeric_count"])
        b.metric("Values with decimals", summary["decimal_count"])
        c.metric("Max decimal places", summary["max_decimal_places"])
        d.metric("Distinct values", int(pd.to_numeric(ddf[selected_col], errors="coerce").nunique()))
        vals = pd.to_numeric(ddf[selected_col], errors="coerce").dropna()
        st.markdown("#### Decimal profile for this column")
        profile = pd.DataFrame({"Statistic": ["Minimum", "Maximum", "Mean of stored values", "Median of stored values", "Decimal-place distribution"],
                                "Value": [summary["min"], summary["max"], float(vals.mean()) if len(vals) else None,
                                          float(vals.median()) if len(vals) else None,
                                          json.dumps(summary["decimal_place_distribution"]) ]})
        st.dataframe(profile, use_container_width=True, hide_index=True)
        if selected_col in ddf.columns and "worker_id" in ddf.columns:
            row_labels = [f"{idx+1:03d} · {ddf.iloc[idx]['worker_id']} · {ddf.iloc[idx][selected_col]}" for idx in range(len(ddf))]
        else:
            row_labels = [f"{idx+1:03d} · {ddf.iloc[idx][selected_col]}" for idx in range(len(ddf))]
        row_choice = st.selectbox("Choose the exact row/value to trace", row_labels, key="decimal_row")
        row_index = row_labels.index(row_choice)
        trace = trace_selected_cell(t, file_lookup, row_index, selected_col)
        st.markdown("#### Trace record")
        x1,x2,x3 = st.columns(3)
        x1.markdown(f"**Source file**\n\n`{trace['file']}`")
        x2.markdown(f"**Sheet / Excel cell**\n\n`{trace['sheet']}!{trace['coordinate']}`")
        x3.markdown(f"**Worker / value**\n\n`{trace['worker_id']}` / `{trace['value']}`")
        st.caption(f"Underlying Excel number format: `{trace['excel_number_format']}` · Stored numeric decimal places: {trace['decimal_places'] if trace['decimal_places'] is not None else 'not numeric'}")
        rule = trace["formula_trace"]
        if rule.get("available") and rule.get("valid"):
            st.markdown("#### Recomputed formula")
            st.code(" + ".join([f"{k}={v}" for k,v in rule["inputs"].items()]) + f" = {rule['sum']:g}", language="text")
            if rule.get("matches"):
                st.success(f"The documented health-score sum reproduces: computed {rule['sum']:g}, stored {rule['observed']}. This confirms score arithmetic for this row, not authenticity of the underlying symptom reports.")
            else:
                st.error(f"The documented score does not match: computed {rule['sum']:g}, stored {rule.get('observed')}.")
        elif rule.get("available") and not rule.get("valid", True):
            st.warning(rule.get("reason", "Inputs unavailable."))
            st.json(rule.get("inputs", {}))
        else:
            st.warning("**Exact formula/source not established from the supplied files.** " + rule.get("reason", "No formula found."))
        with st.expander("Show complete row inputs"):
            st.json({k: (None if pd.isna(v) else (float(v) if isinstance(v, (np.floating,)) else v)) for k,v in trace["row"].items()})
        st.markdown("#### What would establish the decimal's origin?")
        st.markdown("- **Raw entry:** original questionnaire/measurement sheet, respondent or observation ID, date, and the instrument used.\n- **Composite score:** a written formula with components, weights, range, and rounding rule—then recompute from row inputs.\n- **Mean / percentage:** the exact rows used, numerator/denominator, group filter, and rounding convention.\n- **Correlation:** the exact two variables, included rows, missing-data rule, and Pearson/Spearman choice.")

with integrity:
    st.subheader("Integrity checks and cross-file reconciliation")
    st.write("Statuses apply only to the stated test. A PASS on IDs or formulas is not a PASS on real-world data authenticity.")
    status_filter = st.multiselect("Show statuses", ["PASS", "WARN", "FAIL", "INFO"], default=["PASS", "WARN", "FAIL"])
    filtered_checks = checks_df[checks_df["status"].isin(status_filter)] if not checks_df.empty else checks_df
    st.dataframe(filtered_checks, use_container_width=True, hide_index=True, height=360)
    st.markdown("### Reconcile selected source tables")
    if len(tables) >= 2:
        labels = [table_label(t) for t in tables]
        left_label = st.selectbox("Left/source table", labels, index=labels.index(table_label(roles["profile"])) if roles.get("profile") and table_label(roles["profile"]) in labels else 0, key="recon_left")
        right_default = labels.index(table_label(roles["master"])) if roles.get("master") and table_label(roles["master"]) in labels else min(1, len(labels)-1)
        right_label = st.selectbox("Right/target table", labels, index=right_default, key="recon_right")
        result = reconcile_tables(tables[labels.index(left_label)], tables[labels.index(right_label)])
        st.markdown(f"**Result:** {result['status']}")
        r1,r2,r3 = st.columns(3)
        r1.metric("Common worker IDs", result["common_ids"])
        r2.metric("IDs only on left", result["only_left"])
        r3.metric("IDs only on right", result["only_right"])
        if result["field_results"]:
            rdata = pd.DataFrame(result["field_results"])
            st.dataframe(rdata, use_container_width=True, hide_index=True)
            mismatches = rdata[rdata["mismatches"] > 0]
            if len(mismatches):
                st.error("One or more shared fields differ. Select fields above and inspect worker rows before treating the files as reconciled.")
            else:
                st.success("All compared non-missing values in the shared fields match for shared IDs. This only establishes consistency between these two files.")
    else:
        st.info("Load at least two worker-level tables to compare them.")
    st.markdown("### Recomputed correlation matrix")
    if not corr_summary.empty:
        st.dataframe(corr_summary, use_container_width=True)
        st.caption("Recomputed from the selected integrated worker-level data, using pairwise Pearson correlations. Compare with any reported values in analysis/results workbooks; the app does not infer causation or statistical significance from r alone.")
    else:
        st.info("Could not compute correlations: insufficient numeric columns in an integrated table.")

with evidence_tab:
    st.subheader("Provenance evidence and file fingerprints")
    st.markdown("This is the evidence trail the audit can see inside the uploaded workbooks and CSVs. It includes explicit wording and SHA-256 file fingerprints. A fingerprint can show that the bytes later inspected are unchanged; it cannot tell us whether the records were collected in the real world.")
    st.markdown("### Explicit source wording")
    if analysis["evidence"]:
        edf = pd.DataFrame(analysis["evidence"])[["file", "sheet", "cell", "text"]]
        st.dataframe(edf, use_container_width=True, hide_index=True)
    else:
        st.success("No explicit simulation/illustrative wording was detected. The scope limitation still applies.")
    st.markdown("### File hashes and formula counts")
    fdf = pd.DataFrame([{"File": f["name"], "SHA-256": f["sha256"], "Bytes": f["size_bytes"], "Formula cells": f["formula_count"], "Sheets": len(f["sheets"])} for f in files])
    st.dataframe(fdf, use_container_width=True, hide_index=True)
    st.markdown("### Wording-based conclusion")
    if analysis["evidence"]:
        st.error("The supplied project materials themselves contain simulated/illustrative wording. Therefore this app cannot honestly certify the full dataset as real-world observations. The evidence should be disclosed and the project described accurately unless the team can provide contrary primary-source records and clarify which columns were simulated versus collected.")
    else:
        st.warning("The files contain no detected explicit simulation wording, but authenticity remains unverified without original source records.")

with guide:
    st.subheader("How the audit works")
    st.markdown("### Reading the statuses")
    st.markdown("- **PASS:** the exact reproducibility/structure condition passed.\n- **WARN:** missing documentation, partial overlap, or a non-fatal condition that needs human review.\n- **FAIL:** a check found a mismatch, duplicate, range issue, or other violation.\n- **INFO:** observation with no positive/negative judgment.")
    st.markdown("### How the decimal trace works")
    st.markdown("1. Finds the source file, sheet, row, column, and Excel cell address.\n2. Reports whether the value is numeric and how many decimal places are stored.\n3. Recomputes only formulas documented by the data/codebook.\n4. If the exact formula is missing, labels it as undocumented instead of reverse-engineering a convenient explanation.")
    st.markdown("### Known formula coverage")
    st.markdown("**Can be checked:** the health symptom score is documented as the sum of 10 binary symptom indicators, so the app recomputes it row by row.\n\n**Not fully reproducible from current documentation:** Exposure Index, Safety Practice Score, PPE Compliance, and Performance Impairment Score. Their workbooks describe these as derived/illustrative but do not provide complete equations, weights, and rounding rules.")
    st.markdown("### Provisional team responsibility map")
    st.caption("This mapping is inferred from file names and worksheet contents. Confirm it against the exact division your instructor approved; roles are not asserted as a historical record.")
    task_map = pd.DataFrame([
        {"Member": "Member 1", "Likely dataset/workstream": "Dataset 1 — worker profile and exposure variables", "Audit evidence": "ID uniqueness, tenure logic, exposure-zone allocation, source/codebook scan"},
        {"Member": "Member 2", "Likely dataset/workstream": "Dataset 2 — health symptoms and absence", "Audit evidence": "Binary symptom coding; recomputation of the 10-item symptom score"},
        {"Member": "Member 3", "Likely dataset/workstream": "Dataset 3 — safety practices and PPE", "Audit evidence": "Scale/range checks; require exact formula for safety/PPE composites"},
        {"Member": "Member 4", "Likely dataset/workstream": "Dataset 4 — work performance (XLSX and Orange CSV)", "Audit evidence": "Compare CSV/XLSX and master performance fields; formula not documented"},
        {"Member": "Member 5", "Likely dataset/workstream": "Dataset 5 — master integration", "Audit evidence": "Cross-file IDs, field matches, duplicated/derived columns and hashes"},
        {"Member": "Member 6", "Likely dataset/workstream": "Dataset 6 — statistical analysis", "Audit evidence": "Recomputed descriptives, zone means and Pearson correlations"},
        {"Member": "Group integration", "Likely dataset/workstream": "Dataset 7 charts and Dataset 8 final results", "Audit evidence": "Reported outputs should match recomputed statistics; clarify simulated/illustrative labels"},
    ])
    st.dataframe(task_map, use_container_width=True, hide_index=True)
    st.markdown("### Evidence required to substantiate real field data")
    st.markdown("For a defensible claim, retain anonymized original questionnaires/observation forms, dates and collection protocol, sampling notes, any lab or instrument readings and calibration information, a data-entry/change log, consent/ethics approval if applicable, and a written formula sheet for every derived score. The evidence should be consistent with privacy and research-ethics requirements.")

st.divider()
report = report_markdown(analysis, tables)
col1, col2, col3 = st.columns([1, 1, 2])
with col1:
    st.download_button("Download audit report (.md)", report.encode("utf-8"), file_name="research_data_audit_report.md", mime="text/markdown", use_container_width=True)
with col2:
    check_csv = checks_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download checks (.csv)", check_csv, file_name="research_data_audit_checks.csv", mime="text/csv", use_container_width=True)
with col3:
    st.caption("Report generation is based on the files loaded in the current session. Download it along with the file hashes if you need to document the audit run.")
