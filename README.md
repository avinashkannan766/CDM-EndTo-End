
# End-to-End Clinical Data Management Workflow

## Project Overview

This project demonstrates a simulated, end-to-end Clinical Data Management (CDM) workflow using synthetic clinical-trial data, Microsoft Excel, and Python.

The project covers study documentation, CRF design, data dictionary development, synthetic data generation, automated edit checks, query generation, query lifecycle management, visit reconciliation, and consolidated reporting.

**Purpose:** Build practical, demonstrable skills relevant to entry-level Clinical Data Management roles.

> **Disclaimer:** This is an educational portfolio project. All study details and participant records are fictional and synthetically generated. This project does not represent work on a real clinical trial, use of a production Electronic Data Capture (EDC) system, or a claim of GCP/CDISC certification.

## Study Scenario

| Attribute | Description |
|---|---|
| Study ID | DM-001 |
| Study type | Fictional educational study |
| Study population | Adults with type 2 diabetes |
| Study arms | Drug X and placebo |
| Synthetic participants | 30 |
| Planned visits | Screening, Baseline, Week 4, Week 8, Week 12 |
| Tools | Python, Excel, openpyxl |

The scenario is used to demonstrate data-management processes rather than to make clinical or treatment claims.

## Project Objectives

- Define a fictional study specification and CRF structure.
- Document data fields, types, and code lists.
- Generate clean synthetic data and a separate test dataset with deliberate defects.
- Implement automated checks for selected data-quality issues.
- Convert eligible findings into draft data queries.
- Demonstrate controlled query-status transitions.
- Reconcile participant visit records against expected visits and form tracking.
- Consolidate results into Excel reports.

## Workflow

### 1. Study Setup and Documentation

Prepared the following study documents:

- Study specification
- CRF design
- Data dictionary and code lists
- Edit-check specification
- Seeded-defect register

These documents define the fictional study structure, selected data fields, validation rules, and expected test defects.

### 2. Synthetic Data Generation

Generated two workbooks:

- `data/clinical_study_clean.xlsx`
- `data/clinical_study_test.xlsx`

The clean dataset contains 30 participants and five planned visits per participant, giving 150 expected participant-visit combinations.

The test dataset introduces eight deliberate defects to evaluate selected validation rules. The original clean dataset is retained separately.

### 3. Automated Edit Checks

The Python edit-check engine evaluates selected data-quality conditions and produces Excel findings reports.

| Test result | Count |
|---|---:|
| Clean dataset findings | 0 |
| Test dataset findings | 8 |
| Test findings classified as errors | 6 |
| Test findings classified for review | 2 |

The test findings include examples involving missing consent dates, missing visit dates, duplicate visits, missing laboratory results, unexpected units, date inconsistencies, medication date ordering, and an age-related review flag.

The severity assigned by a programmed check is a prompt for investigation, not an automatic decision to change source data.

### 4. Draft Query Generation

The query-generation script converts eligible test findings into draft queries.

- Seven query drafts were generated.
- The drafts include a subject reference, check ID, query description, and workflow status.
- A review finding can also generate a query when its specification requires one.
- Query generation does not modify the underlying study data.

These are fictional training examples, not actual communications with clinical sites.

### 5. Query Lifecycle Demonstration

A separate working copy demonstrates the following query lifecycle:

`Open → Answered → Resolved → Closed`

The script validates allowed transitions and required response or resolution information. It also demonstrates rejection of an invalid direct transition from Open to Closed.

The demonstration uses fictional response and resolution text. It does not represent a real site-query exchange.

### 6. Visit and Form-Tracking Reconciliation

The reconciliation script compares expected participant-visit combinations with visit records and form-tracking records.

**Clean dataset results:**
- 30 participants
- 150 expected visit combinations
- 150 visit rows
- 150 unique visit combinations
- 150 form-tracking rows
- 150 unique tracking combinations
- Zero findings from the implemented reconciliation checks

**Test dataset demonstration:**

The script detected a duplicate Week 8 visit for participant `DM001-005`, reporting two rows for the same participant-visit combination.

This confirms that the implemented duplicate-visit check can identify the deliberately seeded discrepancy. Reconciliation is limited to the checks currently implemented; it is not a comprehensive review of every clinical field or domain.

### 7. Consolidated Reporting

The consolidated workbook is saved as:

`reports/cdm_summary_report.xlsx`

It brings together:
- Edit-check findings
- Draft queries
- Query lifecycle demonstration records
- Reconciliation results
- A project overview and interpretation notes

The summary script reads existing reports and consolidates them. It does not independently rerun the validation, query, or reconciliation checks.

## Repository Structure

```text
CDM_EndToEnd/
├── data/
│   ├── clinical_study_clean.xlsx
│   ├── clinical_study_test.xlsx
│   └── query_log.xlsx
├── docs/
│   ├── study_specification.md
│   ├── crf_design.md
│   ├── data_dictionary.xlsx
│   ├── edit_check_specification.md
│   └── seeded_defect_register.xlsx
├── reports/
│   ├── clean_edit_check_findings.xlsx
│   ├── test_edit_check_findings.xlsx
│   ├── generated_query_log.xlsx
│   ├── query_lifecycle_working.xlsx
│   ├── reconciliation_report.xlsx
│   └── cdm_summary_report.xlsx
├── scripts/
│   ├── generate_study_data.py
│   ├── run_edit_checks.py
│   ├── generate_queries.py
│   ├── manage_query_lifecycle.py
│   ├── run_reconciliation.py
│   └── create_cdm_summary.py
└── README.md
```

*This tree reflects the expected project layout; confirm filenames against the current local repository before publishing.*

## How to Run

### Requirements

- Python 3.10 or later
- `openpyxl`

Install the dependency:

```bash
python -m pip install openpyxl
```

Run the scripts from the project root, in this order:

```bash
python scripts/generate_study_data.py
python scripts/run_edit_checks.py
python scripts/generate_queries.py
python scripts/manage_query_lifecycle.py
python scripts/run_reconciliation.py
python scripts/create_cdm_summary.py
```

**Important:** The reconciliation script reads the file selected by `INPUT_FILE` in `scripts/run_reconciliation.py`. Select the clean dataset for clean-data results or the test dataset to reproduce the duplicate-visit finding.

The lifecycle script and summary script use existing outputs. Run the relevant upstream scripts first if you want to regenerate those outputs.

## Skills Demonstrated

- Clinical Data Management workflow fundamentals
- CRF and data-dictionary documentation
- Python scripting and Excel automation
- Data-quality checks and defect classification
- Draft query generation and status-transition validation
- Visit and form-tracking reconciliation
- Traceable reporting and separation of source data from generated outputs
- Reproducible testing with deliberately seeded defects

## Scope and Limitations

This project is a learning simulation, not a validated clinical system.

- The data are synthetic.
- The checks cover only the rules implemented in the scripts.
- The project does not use a production EDC system or demonstrate real clinical-trial operations.
- It does not establish GCP or CDISC compliance or certification.
- It does not establish readiness for a real clinical database lock.
- Human review remains necessary before interpreting findings or changing any source data.

## Future Improvements

- Add automated unit tests for individual validation rules.
- Expand cross-domain reconciliation for laboratory, adverse-event, and medication records.
- Add explicit traceability from each edit check to its specification and test case.
- Add a reproducible test runner and expected-results comparison.
- Improve report versioning and record script execution metadata.

## Author

Created as a practical portfolio project to develop Python-based data-quality and workflow skills for entry-level Clinical Data Management opportunities.
