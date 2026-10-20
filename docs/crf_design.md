# CRF Design Specification — Study DM-001

## 1. Purpose

This document defines the Case Report Forms (CRFs) for the fictional DM-001 study. Each form specifies the fields to be collected, their data types, and whether they are required.

The CRFs are designed for an educational CDM simulation. They are not intended for clinical use.

## 2. General CRF Conventions

- Each participant is identified using a unique `Subject_ID`.
- Each visit-based form includes a `Visit_Name` where applicable.
- Dates use the format `YYYY-MM-DD`.
- Missing or unavailable data must not be silently replaced with invented values.
- Each submitted form should have a documented completion status.
- Corrections and query resolutions must remain traceable.
- Personal identifiers such as names, phone numbers, and addresses are not collected.

## 3. Form Inventory

| Form ID | Form Name | Purpose |
|---|---|---|
| DM001-DMOGRAPHICS | Demographics and Consent | Record participant characteristics and consent |
| DM001-ELIGIBILITY | Eligibility | Record inclusion and exclusion assessments |
| DM001-VISITS | Visit Tracking | Record planned and actual visits |
| DM001-VITALS | Vital Signs | Record measurements collected at visits |
| DM001-LABS | Laboratory Results | Record laboratory test results |
| DM001-AE | Adverse Events | Record reported adverse events |
| DM001-MEDICATIONS | Concomitant Medications | Record relevant medication history |

## 4. Demographics and Consent Form

**Form ID:** DM001-DMOGRAPHICS

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| Subject_ID | Text | Yes | Unique, nonblank identifier |
| Age | Integer | Yes | 18–70 for this simulation |
| Sex | Coded text | Yes | Use documented code list |
| Consent_Status | Coded text | Yes | Yes / No |
| Consent_Date | Date | Conditional | Required when consent is Yes |

## 5. Eligibility Form

**Form ID:** DM001-ELIGIBILITY

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| Subject_ID | Text | Yes | Must exist in demographics |
| Visit_Name | Coded text | Yes | Screening |
| Inclusion_1 | Coded text | Yes | Yes / No |
| Inclusion_2 | Coded text | Yes | Yes / No |
| Exclusion_1 | Coded text | Yes | Yes / No |
| Eligibility_Status | Coded text | Yes | Eligible / Not Eligible / Pending |
| Assessment_Date | Date | Yes | Valid date |

These criteria are simplified training fields, not medically validated eligibility criteria.

## 6. Visit Tracking Form

**Form ID:** DM001-VISITS

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| Subject_ID | Text | Yes | Must exist in demographics |
| Visit_Name | Coded text | Yes | Screening / Baseline / Week 4 / Week 8 / Week 12 |
| Planned_Date | Date | Yes | Must be a valid date |
| Actual_Date | Date | Conditional | Required when visit is completed |
| Visit_Status | Coded text | Yes | Not Started / Completed / Missed / Not Done |
| Form_Status | Coded text | Yes | Incomplete / Complete / Reviewed |

## 7. Vital Signs Form

**Form ID:** DM001-VITALS

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| Subject_ID | Text | Yes | Must exist in demographics |
| Visit_Name | Coded text | Yes | Must match a scheduled visit |
| Measurement_Date | Date | Yes | Must be a valid date |
| Systolic_BP | Numeric | Yes | Positive value; outliers flagged for review |
| Diastolic_BP | Numeric | Yes | Positive value; outliers flagged for review |
| Heart_Rate | Numeric | Yes | Positive value; outliers flagged for review |
| Weight_kg | Numeric | Yes | Positive value; outliers flagged for review |

Review flags are not automatic diagnoses or clinical judgments.

## 8. Laboratory Results Form

**Form ID:** DM001-LABS

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| Lab_ID | Text | Yes | Unique record identifier |
| Subject_ID | Text | Yes | Must exist in demographics |
| Visit_Name | Coded text | Yes | Must match a scheduled visit |
| Test_Name | Coded text | Yes | Must match the study test list |
| Result_Value | Numeric | Yes | Must be numeric when recorded |
| Unit | Text | Yes | Must match the specified unit for the test |
| Collection_Date | Date | Yes | Must be a valid date |

For this simulation, the laboratory test list will be finalized in the data dictionary.

## 9. Adverse Events Form

**Form ID:** DM001-AE

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| AE_ID | Text | Yes | Unique event identifier |
| Subject_ID | Text | Yes | Must exist in demographics |
| Event_Description | Text | Yes | Nonblank |
| Onset_Date | Date | Yes | Must be a valid date |
| Seriousness | Coded text | Yes | Yes / No |
| Outcome | Coded text | Yes | Resolved / Recovering / Ongoing / Unknown |

This is a simplified training form. Real safety reporting requires additional protocol-specific fields and procedures.

## 10. Concomitant Medications Form

**Form ID:** DM001-MEDICATIONS

| Field | Data Type | Required? | Validation |
|---|---|---|---|
| Medication_ID | Text | Yes | Unique medication record identifier |
| Subject_ID | Text | Yes | Must exist in demographics |
| Medication_Name | Text | Yes | Nonblank |
| Start_Date | Date | Yes | Valid date |
| End_Date | Date | Conditional | Blank if ongoing; otherwise valid date |
| Indication | Text | Yes | Nonblank |

## 11. Form Completion and Review

Each form will be tracked using a status such as:

- Not Started
- In Progress
- Complete
- Reviewed
- Requires Correction

The project will distinguish form completion from data review. A form marked Complete may still contain data queries that must be resolved before lock readiness.

## 12. Design Decisions and Limitations

- The project uses a simplified CRF structure suitable for a beginner portfolio.
- Field requirements and validation rules are training assumptions.
- All data will be synthetic.
- The final field definitions, code lists, and edit checks will be aligned with the data dictionary and study specification.
- This document does not claim compliance with a real protocol or production EDC configuration.
