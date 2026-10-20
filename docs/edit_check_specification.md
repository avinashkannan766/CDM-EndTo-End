# Edit Check Specification — Study DM-001

## 1. Purpose

This document defines the data review rules for the simulated DM-001 clinical data management project.

The checks identify missing, invalid, duplicate, or inconsistent records. Findings may require review, correction, or a data query.

All checks are educational examples applied to synthetic data.

## 2. Check Categories

- **Hard validation:** Identifies a structural or rule violation, such as a missing subject identifier.
- **Review flag:** Identifies a value or situation that needs human assessment; it does not automatically mean the data is wrong.
- **Data query:** A documented request for clarification or correction when the finding cannot be resolved from available information.
- **Cross-form check:** Compares information recorded on different forms.

## 3. Edit Check Inventory

| Check ID | Domain | Rule | Severity | Expected Action |
|---|---|---|---|---|
| EC-001 | Demographics | Subject_ID is blank | Error | Flag record; query if source clarification is required |
| EC-002 | Demographics | Subject_ID is duplicated | Error | Review duplicate records and query if unresolved |
| EC-003 | Demographics | Age is outside 18–70 | Review | Check eligibility and source data; do not automatically change |
| EC-004 | Consent | Consent_Status is Yes but Consent_Date is missing | Error | Raise query for clarification |
| EC-005 | Consent | Consent_Date is later than the participant's first study procedure date | Error | Review chronology and raise query if unresolved |
| EC-006 | Eligibility | Eligibility_Status is Eligible despite an unmet inclusion criterion | Error | Review eligibility responses and query |
| EC-007 | Eligibility | Eligibility_Status is Eligible despite a positive exclusion criterion | Error | Review eligibility responses and query |
| EC-008 | Visits | Visit_Name is not in the approved visit schedule | Error | Correct or clarify the visit name |
| EC-009 | Visits | Visit_Status is Completed but Actual_Date is missing | Error | Query the missing date |
| EC-010 | Visits | Actual_Date precedes the corresponding Planned_Date | Review | Check scheduling and source data; do not assume an error |
| EC-011 | Visits | The same subject has duplicate records for the same scheduled visit | Error | Review duplicate visit records |
| EC-012 | Vital Signs | A required measurement is missing | Error | Query missing data or document the applicable missing-data reason |
| EC-013 | Vital Signs | Measurement is zero, negative, or nonnumeric | Error | Verify the value and query if needed |
| EC-014 | Vital Signs | A value is outside a predefined training review range | Review | Assess plausibility; do not automatically classify as clinically abnormal |
| EC-015 | Laboratory | Lab_ID is blank or duplicated | Error | Review record identity and duplicates |
| EC-016 | Laboratory | Subject_ID does not exist in demographics | Error | Verify the participant identifier |
| EC-017 | Laboratory | Result_Value is missing or nonnumeric | Error | Verify the result or document why it is unavailable |
| EC-018 | Laboratory | Test_Name and Unit do not match the data dictionary | Error | Review test and unit definitions |
| EC-019 | Laboratory | Collection_Date does not align with the associated visit chronology | Review | Review dates and clarify if needed |
| EC-020 | Adverse Events | AE_ID is blank or duplicated | Error | Review record identity and duplicates |
| EC-021 | Adverse Events | Subject_ID does not exist in demographics | Error | Verify the participant identifier |
| EC-022 | Adverse Events | Onset_Date is before the participant's consent date | Review | Check dates and source information |
| EC-023 | Adverse Events | Seriousness or Outcome contains an unapproved code | Error | Verify and correct the coded value |
| EC-024 | Medications | Medication_ID is blank or duplicated | Error | Review record identity and duplicates |
| EC-025 | Medications | End_Date precedes Start_Date | Error | Verify dates and query if unresolved |
| EC-026 | Cross-form | A visit-based record references a visit that does not exist for that subject | Error | Review the visit and associated record |
| EC-027 | Cross-form | A required form is absent for a completed visit | Review | Confirm whether the form is missing, not applicable, or not done |
| EC-028 | Reconciliation | An adverse event requiring further review is not accounted for in the review tracker | Review | Reconcile the event and document the outcome |

## 4. Severity and Query Rules

### Error

An error indicates a violation of a defined data rule. It requires assessment and resolution before the record can be considered clean.

An error does not authorize an automatic change to the original data.

### Review

A review flag indicates a situation that requires human assessment. It may be valid after review and should not automatically generate a query.

### Query

A query is created when clarification or correction is needed from the simulated data-entry or source-review process. Each query should include:

- Query ID
- Subject ID and relevant form/record
- Related check ID
- Finding and query text
- Date raised
- Query status
- Response
- Resolution and closure date

## 5. Query Status Values

The project will use the following statuses:

- Open
- Answered
- Resolved
- Closed

A query should not be marked Closed until the documented resolution has been reviewed. An Answered query may still require further clarification.

## 6. Raw Data and Auditability

- Preserve the original raw dataset unchanged.
- Store proposed or accepted corrections separately from the raw dataset.
- Record the original value, corrected value, reason, and resolution status where applicable.
- Do not invent missing clinical values.
- Do not automatically change review flags into errors without a documented rule.

## 7. Implementation Notes

- The checks will be implemented against the synthetic datasets defined for DM-001.
- Each implemented check will be mapped to its Check ID.
- Checks that require information not represented in the dataset will be documented as planned checks until the necessary fields are available.
- A passing validation result means only that the implemented rules passed; it does not establish clinical validity or regulatory compliance.

## 8. Limitations

This is a portfolio training specification. The checks and query procedures are simplified examples and do not replace a real study protocol, medical review, safety process, or production EDC configuration.
