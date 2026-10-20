
# End-to-End Clinical Data Management Workflow

## Project Overview

This project demonstrates a simulated, end-to-end Clinical Data Management (CDM) workflow using synthetic clinical-trial data, Microsoft Excel, and Python.

The workflow covers study documentation, CRF design, data dictionary development, synthetic data generation, automated edit checks, query generation, query lifecycle management, correction audit-trail documentation, visit reconciliation, consolidated reporting, and automated testing.

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
| Tools | Python, Excel, openpyxl, unittest |

The scenario demonstrates data-management processes rather than making clinical or treatment claims.

## Project Objectives

- Define a fictional study specification and CRF structure.
- Document data fields, types, and code lists.
- Generate clean synthetic data and a separate test dataset containing deliberate defects.
- Implement automated checks for selected data-quality issues.
- Convert eligible findings into draft data queries.
- Demonstrate controlled query-status transitions.
- Document proposed data corrections in a separate audit-trail workbook.
- Reconcile participant visit records against expected visits and form tracking.
- Consolidate results into Excel reports.
- Verify expected validation behavior with automated tests.

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

The test findings include missing consent dates, missing visit dates, duplicate visits, missing laboratory results, unexpected units, date inconsistencies, medication date ordering, and an age-related review flag.

A programmed finding is a prompt for investigation, not an automatic decision to change source data.

### 4. Draft Query Generation

The query-generation script converts eligible findings into draft queries.

- Seven query drafts were generated.
- The drafts include query identifiers, subject references, edit-check identifiers, and query descriptions.
- Findings that do not require a query are excluded according to the configured specification.
- Query generation does not modify the underlying study data.

These are fictional training examples, not actual communications with clinical sites.

### 5. Query Lifecycle Demonstration

A separate working copy demonstrates the query lifecycle:

`Open → Answered → Resolved → Closed`

The script validates allowed status transitions and required response or resolution information. It also demonstrates rejection of an invalid direct transition from Open to Closed.

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

Reconciliation is limited to the checks currently implemented; it is not a comprehensive review of every clinical field or domain.

### 7. Correction Audit-Trail Documentation

The script `scripts/create_correction_audit_log.py` creates:

`reports/data_correction_audit_log.xlsx`

The workbook contains three worksheets:

- **Correction Audit:** a template for documenting original values, proposed replacements, reasons, review status, reviewer, and relevant dates.
- **Query Reference:** references the seven query records in the query lifecycle workbook to support traceability.
- **Instructions:** describes the intended fields and their use.

The audit workbook is separate from the source clinical data and query lifecycle workbook. Creating it does not apply a correction or automatically approve a proposed change.

This is a documentation template for synthetic training data. It does not implement an authorized production correction process or establish a complete, immutable audit trail.

### 8. Consolidated Reporting

The consolidated workbook is saved as:

`reports/cdm_summary_report.xlsx`

It brings together:
- Edit-check findings
- Draft queries
- Query lifecycle demonstration records
- Reconciliation results
- A project overview and interpretation notes

The summary script reads existing reports and consolidates them. It does not independently rerun the validation, query, or reconciliation checks.

### 9. Automated Testing

A Python `unittest` suite tests the edit-check engine against the clean and deliberately seeded test datasets.

The current suite contains six automated tests:

| Test | Expected result |
|---|---|
| Clean dataset | Zero findings |
| Test dataset | Eight findings |
| Expected edit-check IDs | All eight expected check IDs detected |
| Severity classification | Six errors and two review flags |
| Finding IDs | Unique, sequential, and repeatable |
| Source data integrity | Source workbook SHA-256 fingerprints remain unchanged |

The latest local test run passed all six tests.

Run the suite from the project root:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

The captured test output is saved in:

`reports/test_results.txt`

These tests verify the current implementation against the expected synthetic cases. They do not establish that the validation engine is comprehensive or suitable for production clinical-trial use.

## Repository Structure

```text
CDM_EndToEnd/
├── .gitignore
├── README.md
├── requirements.txt
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
│   ├── data_correction_audit_log.xlsx
│   ├── reconciliation_report.xlsx
│   ├── cdm_summary_report.xlsx
│   └── test_results.txt
├── scripts/
│   ├── generate_study_data.py
│   ├── create_data_dictionary.py
│   ├── run_edit_checks.py
│   ├── create_query_log.py
│   ├── generate_queries.py
│   ├── manage_query_lifecycle.py
│   ├── create_correction_audit_log.py
│   ├── run_reconciliation.py
│   └── create_cdm_summary.py
└── tests/
    └── test_edit_checks.py
```

This tree represents the intended layout. Confirm filenames against the current local repository before publishing, because generated reports may vary depending on which scripts have been run.

## Requirements

- Python 3.10 or later
- `openpyxl`

Install dependencies from the project root:

```powershell
python -m pip install -r requirements.txt
```

Alternatively:

```powershell
python -m pip install openpyxl
```

## How to Run

Run commands from the project root.

### 1. Generate synthetic study data

```powershell
python scripts/generate_study_data.py
```

### 2. Generate or refresh the data dictionary

```powershell
python scripts/create_data_dictionary.py
```

### 3. Run edit checks

```powershell
python scripts/run_edit_checks.py
```

### 4. Create the query-log template, if needed

```powershell
python scripts/create_query_log.py
```

### 5. Generate draft queries

```powershell
python scripts/generate_queries.py
```

### 6. Demonstrate the query lifecycle

```powershell
python scripts/manage_query_lifecycle.py
```

### 7. Generate the correction audit-trail workbook

```powershell
python scripts/create_correction_audit_log.py
```

This step requires `reports/query_lifecycle_working.xlsx` to exist. The script stops rather than overwriting an existing correction audit workbook.

### 8. Run visit reconciliation

```powershell
python scripts/run_reconciliation.py
```

The reconciliation script reads the file selected by `INPUT_FILE` in `scripts/run_reconciliation.py`. Select the clean dataset for clean-data results or the test dataset to reproduce the duplicate-visit finding.

### 9. Create the consolidated report

```powershell
python scripts/create_cdm_summary.py
```

### 10. Run automated tests

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

The test suite reads the existing synthetic source datasets and does not regenerate or overwrite them.

**Execution notes:**

- The lifecycle, correction audit, and summary scripts depend on outputs from earlier steps.
- Some scripts overwrite their own generated outputs. Review their file paths and behavior before rerunning them.
- The correction audit generator protects an existing audit workbook by refusing to overwrite it.
- Keep source data, test data, working copies, and reports clearly separated.

## Skills Demonstrated

- Clinical Data Management workflow fundamentals
- CRF and data-dictionary documentation
- Python scripting and Excel automation
- Data-quality checks and defect classification
- Draft query generation and status-transition validation
- Visit and form-tracking reconciliation
- Correction audit-trail documentation and traceability
- Automated testing with Python's `unittest` framework
- Regression checks and source-file integrity verification
- Reproducible testing with deliberately seeded defects
- Separation of source data from generated outputs
- Git and GitHub version control

## Scope and Limitations

This project is a learning simulation, not a validated clinical system.

- All study records are synthetic.
- Checks cover only the rules implemented in the scripts.
- The project does not use a production EDC system or demonstrate real clinical-trial operations.
- It does not establish GCP or CDISC compliance or certification.
- It does not establish readiness for a real clinical database lock.
- The correction audit workbook is a template; it does not apply or approve corrections.
- Automated tests cover the current expected cases and do not guarantee that all defects will be detected.
- Human review remains necessary before interpreting findings or changing source data.

## Future Improvements

- Add unit tests for individual validation rules and edge cases.
- Expand cross-domain reconciliation for laboratory, adverse-event, and medication records.
- Add explicit traceability from each edit check to its specification and test case.
- Add a reproducible test runner and expected-results comparison.
- Improve report versioning and record script execution metadata.
- Develop a controlled correction workflow with validation, approval, and verification steps.
- Add automated tests for query lifecycle transitions and reconciliation logic.

## Author

Created as a practical portfolio project to develop Python-based data-quality and workflow skills for entry-level Clinical Data Management opportunities.
