# compare_file_refs

`compare_file_refs.py` downloads matter data from LegalSuite, compares it against the Standard Bank claim amount file from FTP, writes missing file refs to Excel, performs weekly balancing checks, creates verification artifacts, and can optionally close matters in LegalSuite based on the weekly balancing file.

Credentials are loaded from a local `.env` file via `env_config.py`, following the same pattern as `ftp_download_today.py`.

## Flow

1. Pull the LegalSuite matter report with the configured client IDs, or use `--csv` to read `Report (SUP).CSV`.
2. Download the Standard Bank claim amount file from FTP, or use `--standard` to supply it directly.
3. Compare LegalSuite `Matter File Ref` values to claim file `File Ref` values.
4. Write missing refs to `Missing_Matter_File_Ref.xlsx`.
5. Copy generated result workbooks into `verification/` and stamp row-level verification columns.
6. If missing refs exist, download the weekly balancing file and look up their statuses by weekly `File Ref`.
7. If a ref is still not found in weekly balancing, fetch the matter from LegalSuite, take `Their Ref`, and search the weekly `Matter` column with that value.
8. If the weekly row is found by `Their Ref`, write it to `Panel L Matter Ref Updates.xlsx` with columns `Matter File Ref` and `Their Ref`.
9. Upload `Panel L Matter Ref Updates.xlsx` to the FTP folder root `Matter Ref Updates`.
10. Write weekly matches to `Matter_Found_From_Weekly.xlsx` and unresolved misses to `Matter_Not_Found_From_Weekly.xlsx`.
11. Write verification summary files to `verification/verification_summary.json` and `verification/verification_summary.txt`.
12. If `--close-closed` is enabled, close matters in LegalSuite when the weekly status is `Closed`.

## Weekly Fallback

When a missing ref cannot be found in the weekly balancing file by `File Ref`, the script uses the anomaly process:

1. Fetch the matter from LegalSuite by the missing `File Ref`.
2. Read `Their Ref` from the LegalSuite matter.
3. Search the weekly balancing `Matter` column using `Their Ref`.
4. If a row is found, treat it as a weekly match, keep its weekly status, and add the correct LegalSuite `Matter File Ref` plus `Their Ref` to `Panel L Matter Ref Updates.xlsx`.
5. Upload that workbook to the FTP folder `Matter Ref Updates` itself, not to a subfolder.

## Close Handling

When `--close-closed` is used, the script tries to archive the matching LegalSuite matter.

If the archive is rejected by LegalSuite, or if the update succeeds but the matter still comes back as `Live`, the script falls back to setting the matter to `Pending Deletion`.

This fallback is also used when the archive update call errors after the matter has already been fetched.

After either update path, the script fetches the matter again and verifies the expected archive fields. For the normal archive path, `archiveno` is not treated as fixed because LegalSuite may assign it during archiving. For the pending deletion fallback, the returned archive state is verified against the fallback payload.

### Close Process

1. Find missing refs that appear in the weekly file with a `Closed` status.
2. Fetch the current matter from LegalSuite by `FileRef`.
3. Build the archive payload and send the update.
4. Fetch the matter again and verify the archive fields match the requested archived state.
5. If the archive update is rejected, errors, or verification shows the matter is still `Live`, build a `Pending Deletion` payload instead.
6. Send the `Pending Deletion` update.
7. Fetch the matter again and verify the archive fields now match the pending state.
8. Only count the matter as successfully handled if one of those verified end states matches.

## Verification

The script creates a `verification/` folder and writes:

- verification copies of generated workbooks with added columns such as:
  - `Verification Status`
  - `Verification Timestamp`
  - `Verification Notes`
  - `Verification GET Response`
- `verification_summary.json`
- `verification_summary.txt`

The verification copies mirror the local output filenames, for example:

- `verification/Missing_Matter_File_Ref.xlsx`
- `verification/Matter_Found_From_Weekly.xlsx`
- `verification/Matter_Not_Found_From_Weekly.xlsx`
- `verification/Panel L Matter Ref Updates.xlsx`

## Common Flags

- `--csv PATH`: use a CSV file instead of downloading the LegalSuite report.
- `--standard PATH`: use a local Standard Bank claim file instead of FTP.
- `--day N`: use the claim file date from `N` days back.
- `--client-ids ID [ID ...]`: override the LegalSuite client list.
- `--close-closed`: update LegalSuite for refs found as closed in the weekly file.
- `--close-dry-run`: build the close payload without updating LegalSuite.
- `--close-verbose`: print LegalSuite request and response payloads for close updates.
- `--debug-legal-suite-output PATH`: dump the raw LegalSuite API response to JSON.
- `--verification-dir PATH`: choose where verification artifacts are written.
- `--matter-ref-updates-output PATH`: override the local `Panel L Matter Ref Updates.xlsx` path.

## Environment Variables

The script loads these from `.env` in the same folder:

- `FTP_HOST`
- `FTP_USER`
- `FTP_PASS`
- `LEGALSUITE_API_KEY`

## Example

```bash
python3 compare_file_refs.py --day 1 --close-closed
```

## Outputs

- `Report (SUP).xlsx`
- `Missing_Matter_File_Ref.xlsx`
- `Matter_Found_From_Weekly.xlsx`
- `Matter_Not_Found_From_Weekly.xlsx`
- `Panel L Matter Ref Updates.xlsx`
- `verification/verification_summary.json`
- `verification/verification_summary.txt`
