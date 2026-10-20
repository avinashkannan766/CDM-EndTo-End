
from pathlib import Path
from datetime import date
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo


# --------------------------------------------------
# Project paths
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent
REPORT_DIR = BASE_DIR / "reports"
OUTPUT_FILE = REPORT_DIR / "generated_query_log.xlsx"
FINDINGS_FILE = REPORT_DIR / "test_edit_check_findings.xlsx"

REPORT_DIR.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------
# Query-generation rules
# --------------------------------------------------

QUERY_COLUMNS = [
    "Query_ID",
    "Subject_ID",
    "Form_Domain",
    "Record_Reference",
    "Edit_Check_ID",
    "Query_Type",
    "Query_Text",
    "Date_Raised",
    "Query_Status",
    "Response",
    "Resolution",
    "Date_Closed",
    "Reviewer_Notes",
]

QUERY_TEXTS = {
    "EC-004": (
        "Consent status is recorded as Yes, but the consent date is "
        "missing. Please verify the source information and provide or "
        "clarify the consent date. Do not change the record without "
        "appropriate source confirmation."
    ),
    "EC-009": (
        "The visit is marked Completed, but the actual visit date is "
        "missing. Please verify the visit date or clarify why it is "
        "unavailable."
    ),
    "EC-011": (
        "More than one record exists for this subject and visit. "
        "Please review the source information and confirm which record "
        "is correct. Do not delete or overwrite records without "
        "documented review."
    ),
    "EC-017": (
        "The laboratory result is missing, nonnumeric, or not positive "
        "according to this project's programmed check. Please verify "
        "the source result, units, and transcription."
    ),
    "EC-018": (
        "The laboratory unit differs from the unit expected by this "
        "project's data specification. Please verify the source report "
        "and confirm the correct unit."
    ),
    "EC-022": (
        "The adverse-event onset date is earlier than the recorded "
        "consent date. Please review the source information and clarify "
        "the chronology. This is a review flag, not an automatic "
        "determination that the event record is incorrect."
    ),
    "EC-025": (
        "The medication end date is earlier than its start date. "
        "Please verify both dates against the source information and "
        "clarify the correct chronology."
    ),
}

QUERY_TYPES = {
    "EC-004": "Missing Data",
    "EC-009": "Missing Data",
    "EC-011": "Inconsistency",
    "EC-017": "Missing Data",
    "EC-018": "Inconsistency",
    "EC-022": "Date Inconsistency",
    "EC-025": "Date Inconsistency",
}


# --------------------------------------------------
# Read findings
# --------------------------------------------------

def read_findings(path):
    if not path.exists():
        raise FileNotFoundError(
            f"Findings report not found: {path}"
        )

    workbook = load_workbook(path, data_only=True)

    if "Findings" not in workbook.sheetnames:
        workbook.close()
        raise ValueError(
            "The findings workbook has no 'Findings' sheet."
        )

    sheet = workbook["Findings"]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()

    if not rows:
        return []

    headers = [
        str(value).strip() if value is not None else ""
        for value in rows[0]
    ]

    findings = []

    for values in rows[1:]:
        if all(value is None for value in values):
            continue

        findings.append({
            headers[index]: values[index]
            for index in range(min(len(headers), len(values)))
            if headers[index]
        })

    return findings


# --------------------------------------------------
# Generate query records from eligible findings
# --------------------------------------------------

def generate_query_records(findings):
    query_records = []
    skipped_review_findings = []

    # Stable ordering keeps query IDs reproducible.
    findings = sorted(
        findings,
        key=lambda row: (
            str(row.get("Check_ID", "")),
            str(row.get("Subject_ID", "")),
            str(row.get("Record_Reference", "")),
        ),
    )

    for finding in findings:
        check_id = str(finding.get("Check_ID", "")).strip()
        severity = str(finding.get("Severity", "")).strip()
        requires_query = str(
            finding.get("Requires_Query", "")
        ).strip().lower()

        # Only findings explicitly marked for queries are eligible.
        if requires_query != "yes":
            continue

        # EC-003 age range is a review flag and is not automatically
        # converted into a query by this workflow.
        if check_id == "EC-003" or check_id not in QUERY_TEXTS:
            skipped_review_findings.append(finding)
            continue

        query_id = f"QRY-{len(query_records) + 1:03d}"

        query_records.append({
            "Query_ID": query_id,
            "Subject_ID": finding.get("Subject_ID"),
            "Form_Domain": finding.get("Domain"),
            "Record_Reference": finding.get("Record_Reference"),
            "Edit_Check_ID": check_id,
            "Query_Type": QUERY_TYPES[check_id],
            "Query_Text": QUERY_TEXTS[check_id],
            "Date_Raised": date.today(),
            "Query_Status": "Open",
            "Response": None,
            "Resolution": None,
            "Date_Closed": None,
            "Reviewer_Notes": (
                "Generated from a synthetic test finding. "
                "Requires human review; not a real clinical query."
            ),
        })

    return query_records, skipped_review_findings


# --------------------------------------------------
# Write generated query log
# --------------------------------------------------

def write_query_log(query_records, output_path):
    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Generated Queries"

    sheet.append(QUERY_COLUMNS)

    for record in query_records:
        sheet.append([
            record.get(column)
            for column in QUERY_COLUMNS
        ])

    # Format the header.
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            vertical="top",
            wrap_text=True
        )

    # Make dates display clearly.
    date_column = QUERY_COLUMNS.index("Date_Raised") + 1
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=date_column).number_format = (
            "yyyy-mm-dd"
        )

    # Make the sheet readable.
    widths = {
        "A": 15,
        "B": 16,
        "C": 20,
        "D": 25,
        "E": 16,
        "F": 20,
        "G": 75,
        "H": 16,
        "I": 16,
        "J": 30,
        "K": 30,
        "L": 16,
        "M": 55,
    }

    for column_letter, width in widths.items():
        sheet.column_dimensions[column_letter].width = width

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    # Add an Excel table only when there are data rows.
    if sheet.max_row >= 2:
        table = Table(
            displayName="GeneratedQueryTable",
            ref=sheet.dimensions
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)

    # Add a workflow note sheet.
    guide = workbook.create_sheet("Workflow Notes")
    guide.append(["Item", "Description"])
    guide.append([
        "Purpose",
        "Generated from the project's synthetic test findings."
    ])
    guide.append([
        "Initial status",
        "All generated queries start as Open."
    ])
    guide.append([
        "Data integrity",
        "The source dataset and findings report are not modified."
    ])
    guide.append([
        "Review requirement",
        "A generated query is a draft for human review, not proof "
        "that source data are incorrect."
    ])
    guide.append([
        "Limitations",
        "This is an educational Excel workflow, not a production EDC "
        "or validated clinical system."
    ])

    for cell in guide[1]:
        cell.font = Font(bold=True)

    guide.column_dimensions["A"].width = 24
    guide.column_dimensions["B"].width = 100

    for row in guide.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    guide.freeze_panes = "A2"

    workbook.save(output_path)
    workbook.close()


# --------------------------------------------------
# Main
# --------------------------------------------------

if __name__ == "__main__":
    findings = read_findings(FINDINGS_FILE)

    queries, skipped = generate_query_records(findings)

    write_query_log(queries, OUTPUT_FILE)

    print("\nQuery generation completed.")
    print(f"Findings reviewed: {len(findings)}")
    print(f"Draft queries generated: {len(queries)}")
    print(f"Findings not converted to queries: {len(skipped)}")
    print(f"Output saved: {OUTPUT_FILE}")

    print("\nGenerated query counts by status:")
    print(f"Open: {len(queries)}")
