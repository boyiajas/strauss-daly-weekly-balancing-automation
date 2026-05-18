# Automation Flow

## Purpose

This automation reconciles LegalSuite matters against Standard Bank source files, identifies exceptions, applies weekly-balancing fallback logic, and prepares follow-up actions for records that need correction or closure.

The goal is to keep LegalSuite aligned with the bank-side files while producing a clear audit trail of what was matched, what was missing, and what action was taken.

## Main Systems

- `LegalSuite`
  Source of internal matter records, including `Matter File Ref` and `Their Ref`.

- `FTP`
  Source of Standard Bank files and destination for correction files.

- `Claim Amount File`
  Daily file used as the primary comparison source.

- `Weekly Balancing File`
  Weekly file used to determine matter status and investigate anomalies.

- `Verification Outputs`
  Local artifacts that record what the automation concluded and why.

## Architecture Overview

This automation sits between three main operational layers:

1. `Source systems`
   - LegalSuite for internal matter data
   - FTP for external Standard Bank files

2. `Automation layer`
   - `compare_file_refs.py`
   - comparison logic
   - fallback anomaly handling
   - LegalSuite close/update logic
   - verification and audit generation

3. `Output and handoff layer`
   - local Excel outputs
   - verification artifacts
   - FTP upload of matter-ref correction files

At a high level, the architecture is a reconciliation pipeline with exception routing:

- LegalSuite or CSV provides the starting matter population.
- FTP claim files provide the primary reconciliation source.
- FTP weekly files provide exception-resolution and status context.
- LegalSuite provides the fallback lookup key `Their Ref` when `File Ref` fails.
- The automation writes outputs locally and, where needed, pushes correction files back to FTP.

## Field Mapping

### Primary Comparison Mapping

| Source | Source Field | Target | Target Field | Purpose |
| --- | --- | --- | --- | --- |
| LegalSuite report / CSV | `Matter File Ref` | Claim Amount file | `File Ref` | Primary reconciliation key |

### Weekly Lookup Mapping

| Source | Source Field | Target | Target Field | Purpose |
| --- | --- | --- | --- | --- |
| Missing matter list | `Matter File Ref` | Weekly Balancing file | `File Ref` | First weekly lookup attempt |
| LegalSuite matter | `Their Ref` | Weekly Balancing file | `Matter` | Fallback lookup when weekly `File Ref` does not match |

### Matter Ref Update File Mapping

| Source | Source Field | Output File | Output Column | Purpose |
| --- | --- | --- | --- | --- |
| LegalSuite matter | `Matter File Ref` | `Panel L Matter Ref Updates.xlsx` | `Matter File Ref` | Correct file ref to be restored downstream |
| LegalSuite matter | `Their Ref` | `Panel L Matter Ref Updates.xlsx` | `Their Ref` | Account number used to identify the weekly record |

### Weekly Status Mapping

| Source | Source Field | Internal Use | Purpose |
| --- | --- | --- | --- |
| Weekly Balancing file | `Current Status` | weekly status value | Determines whether the matter stays live, needs follow-up, or can be closed |

### LegalSuite Close Verification Mapping

| Source | Source Field | Verification Use | Purpose |
| --- | --- | --- | --- |
| LegalSuite matter after update | `archiveflag` | closure verification | Confirms archive state |
| LegalSuite matter after update | `archivestatus` | closure verification | Confirms archive status code |
| LegalSuite matter after update | `archivestatusdescription` | closure verification | Confirms archive status description |
| LegalSuite matter after update | `archiveno` | fallback / informational verification | Archive number assigned or preserved by LegalSuite |

## High-Level Flow

1. Collect the matter list from LegalSuite or from an input CSV.
2. Download the current Standard Bank claim amount file from FTP.
3. Compare LegalSuite `Matter File Ref` values to the claim file `File Ref` values.
4. Separate records into matched and missing groups.
5. Write the missing group to a local exception workbook.
6. If there are missing records, download the weekly balancing file.
7. Try to locate each missing record in weekly balancing by `File Ref`.
8. For records still not found, use LegalSuite `Their Ref` to search the weekly `Matter` column.
9. Decide the next action based on what is found and what status is returned.
10. Produce local outputs and verification artifacts.
11. If required, upload a correction workbook to FTP.
12. If required, update LegalSuite for matters that should be closed.

## Decision Paths

### Path 1: Direct Match In Claim File

If a LegalSuite `Matter File Ref` exists in the claim amount file, no exception action is needed.

Outcome:
- Matter is treated as reconciled for that run.

### Path 2: Missing From Claim File, Found In Weekly By File Ref

If a matter is missing from the claim file but appears in the weekly balancing file under the same `File Ref`, the automation uses the weekly record to determine the current status.

Outcome:
- Matter is included in the weekly found output.
- Weekly status is available for downstream action.

### Path 3: Missing From Claim File, Not Found In Weekly By File Ref, Found By Their Ref

If a matter cannot be found in weekly balancing by `File Ref`, the automation checks LegalSuite for that matter and reads `Their Ref` as the account number.

That account number is then searched in the weekly balancing `Matter` column.

If a match is found, it means the weekly record exists but is associated with an incorrect `File Ref`.

Outcome:
- A correction row is created using:
  - `Matter File Ref` from LegalSuite
  - `Their Ref` from LegalSuite
- That row is written to `Panel L Matter Ref Updates.xlsx`.
- The correction workbook is uploaded to the FTP `Matter Ref Updates` root folder.

### Path 4: Missing From Claim File, Not Found In Weekly At All

If a matter cannot be found in weekly balancing by either `File Ref` or `Their Ref`, it remains unresolved for that run.

Outcome:
- Matter stays in the weekly not-found output.
- Manual follow-up may be needed.

## Status Handling

The weekly balancing file is treated as the operational status source for exception handling.

### Open Or Live-Type Status

If the weekly status indicates the matter is still open, the matter should generally remain live in LegalSuite even if it was missing from the claim file.

Interpretation:
- The account may not yet be ready for closure.
- No closure action should be forced prematurely.

### Closed Status

If the weekly status shows the matter is closed, the automation can optionally trigger LegalSuite closure logic.

Outcome:
- LegalSuite is updated to archive the matter.
- If archive is rejected or does not stick, a fallback state such as Pending Deletion may be used.
- The final LegalSuite state is verified after the update.

## Outputs

### Operational Outputs

- `Report (SUP).xlsx`
  Normalized report used for comparison.

- `Missing_Matter_File_Ref.xlsx`
  All claim-file exceptions from the primary comparison.

- `Matter_Found_From_Weekly.xlsx`
  Missing matters later found in weekly balancing.

- `Matter_Not_Found_From_Weekly.xlsx`
  Missing matters still unresolved after weekly lookup.

- `Panel L Matter Ref Updates.xlsx`
  Correction file for weekly records located by `Their Ref`.

### Verification Outputs

- `verification/`
  Folder containing verification copies of result workbooks.

- `verification_summary.json`
  Machine-readable summary of the run.

- `verification_summary.txt`
  Human-readable summary of the run.

These outputs support traceability, review, and audit of what happened during a run.

## FTP Handoffs

### Files Downloaded From FTP

- Standard Bank daily claim amount file
- Standard Bank weekly balancing file

### Files Uploaded To FTP

- `Panel L Matter Ref Updates.xlsx`

Important rule:
- The correction workbook must be uploaded to the root `Matter Ref Updates` folder itself, not into a subfolder.

## LegalSuite Handoffs

### Data Read From LegalSuite

- `Matter File Ref`
- `Their Ref`
- matter state used for verification and closure logic

### Data Potentially Written Back To LegalSuite

- archive / close updates for matters confirmed as closed
- fallback archive-state handling when the primary archive action is rejected

## Control Principles

- Prefer direct reconciliation first.
- Use weekly balancing as the operational exception source.
- Use `Their Ref` as the fallback key when `File Ref` fails.
- Do not close matters only because they are absent from the claim file.
- Verify every significant automation outcome.
- Preserve artifacts for traceability and later review.

## Summary

At an abstract level, this automation is an exception-management pipeline.

It starts with a straight comparison, then moves through progressively deeper validation:

1. Compare by official file reference.
2. Validate against weekly balancing.
3. Investigate anomalies through LegalSuite account references.
4. Generate correction files where identifiers are inconsistent.
5. Update LegalSuite only when the weekly status supports that action.
6. Record the full result set in verification outputs.
