# Validate Translation Tables - User Guide (Plain English)

This guide explains how to use the app without technical jargon.

## What this tool does

The app supports 3 workflows:

1. `Validate` - check an existing translation table and fix errors.
2. `Create` - build a new translation table from Outcomes and myWSU.
3. `Join Preview` - join your translation table with Outcomes and myWSU source data for visual verification before publish.

A translation table maps one Outcomes key (`translate_input`) to one myWSU key (`translate_output`).

## Basic terms

- `Workflow`: Validate, Create, or Join Preview.
- `Sheet`: a tab in the Excel file.
- `Key`: unique ID column used for mapping.
- `Name matching`: uses school names (and location context) when key matching is not enough.

## Before you start

- Make sure Outcomes and myWSU key columns are unique.
- Include school name, city, state, and country columns if possible.
- If a file looks garbled after upload, change encoding and re-parse.

## Workflow 1: Validate an existing translation table

Use this when you already have a translation table and want to correct it.

### Files required

- Outcomes source file
- Translate table file
- myWSU source file
- **Optional:** Campus-family JSON rules file - to prefill parent keys for campus variants (for example, Texas A&M* -> TAMU-MAIN)
- **Optional (in editor panel after validation):**
  - Prior Validate workbook (Excel from a previous Validate run) - to re-apply prior decisions
  - Review session JSON - to resume saved in-app edits

### Steps in the app

1. Select `Validate`.
2. Upload all 3 required files. Optionally upload a **Campus-family JSON** file to prefill parent keys.
3. In `Select Columns, Keys, and Roles`:
   - Pick key columns for Outcomes, Translate input/output, and myWSU.
   - Pick included columns.
   - Optionally map roles (`School`, `City`, `State`, `Country`).
4. Choose validation mode:
   - `Key only`, or
   - `Key + name comparison` (default).
5. If using name comparison, pick name columns and adjust threshold/ambiguity gap.
6. Click `Validate Mappings`.
7. Review on-screen cards and counts.
8. Optional: click `Bulk edit before export` to apply `Decision` and/or `Manual_Suggested_Key` to filtered rows in one step.
   - New in-app review grid includes Outcomes/myWSU name + state + country context, per-row suggestion dropdown, and `Reason_Code`.
   - Use quick family chips (for example, `Texas A&M`, `Troy University`) to jump to common campus groups.
   - `Bulk edit before export` is a one-time opener for a run. After first click it stays open for that run and re-opens automatically after re-validate; `Start Over` resets this.
   - For performance, use pagination controls (100/200/400 rows per page) with synced controls above and below the grid. Bulk actions can target all filtered rows or only the current page.
   - For duplicate and one-to-many style review rows, the Suggested myWSU dropdown lists the top 5 location-valid name matches, ranked by similarity.
   - `Review Scope` controls both what appears in the in-app grid and what is included in the downloaded report:
     - `All review rows`
     - `Uploaded Translate rows only`
     - `Missing mappings only`
   - `Review Scope` export behavior:
     - `Uploaded Translate rows only` excludes Missing_Mapping rows from the export review queue and omits the `Missing_Mappings` sheet.
     - `Missing mappings only` keeps only Missing_Mapping rows in the export review queue, scopes `Final_Translation_Table` to missing-mapping review output, and suppresses non-missing diagnostic sheets.
   - In the editor panel, upload **Prior Validate workbook** and/or **Review session JSON** to resume prior work after validation.
   - Use `Save session` to download a JSON snapshot and `Load session` to resume later.
9. Click `Download Full Report`.

**Note:** The Validate Excel file requires **Excel 365 or 2021+** to display the Final_Translation_Table correctly. Excel 2016/2019 may show errors for the compact table.

### What the Excel file contains

The downloaded Excel file includes:

- **Auto-matched rows** - rows classified as `Valid` or `High_Confidence_Match` are already approved and appear in `Final_Translation_Table` automatically.
- **Rows needing your decision** - errors, mismatches, duplicates, and missing mappings appear in `Review_Workbench`. You must choose a Decision for each of these before the final table is complete.

**You must use Review_Workbench** to resolve all rows that need a decision. The Final_Translation_Table is built from your decisions plus the auto-matched rows.

### Validate Excel review order (left to right)

Recommended order:

1. `Review_Workbench` - main decision sheet (use this first). Sortable and filterable. When a row has multiple suggestions (C1, C2, C3), the **C1**, **C2**, **C3** columns show each option inline (Key: Name - City, State, Country | Score). Compare them on the same row, then enter C1, C2, or C3 in **Selected Candidate ID** to pick the best match. No need to switch sheets.
2. `Final_Translation_Table` - final publish-ready key table (auto-matched + your approved decisions).
3. `Translation_Key_Updates` - **What changed** verification sheet: only key pairs that differ from current. Use this to quickly verify all changes before publish.
4. `QA_Checks_Validate` - publish gate checks.

Hidden by default (diagnostic/internal):

- `Errors_in_Translate`
- `Output_Not_Found_Ambiguous` (if present)
- `Output_Not_Found_No_Replacement` (if present)
- `One_to_Many`
- `Missing_Mappings`
- `High_Confidence_Matches`
- `Valid_Mappings`
- `Action_Queue`
- `Candidate_Pool`
- `Candidate_Reference`
- `Approved_Mappings`
- `Final_Staging`

**Scope note:** When you export with `Missing mappings only`, non-missing diagnostic sheets are intentionally not created for that workbook.

**To unhide Candidate_Pool or Candidate_Reference:** Right-click the tab bar at the bottom of Excel, choose **Unhide**, then select the sheet you need. These sheets support the C1/C2/C3 candidate lookup; unhide them only when you need to inspect candidate details.

### How `Review_Workbench` works

**Your decisions here control which rows appear in the Final_Translation_Table.** Each row needing review shows:

- Outcomes and myWSU name/key context
- Source location columns when selected: Outcomes State, Outcomes Country; myWSU City, myWSU State, myWSU Country (blanks in source appear as blank)
- Current translate keys and suggested key/school/city/state/country (verify suggested location before applying Use Suggestion). When the best suggestion equals the current value, Suggested_Key is left blank to avoid confusion—use Manual_Suggested_Key or Keep As-Is instead.
- `Decision` (editable dropdown) - **this determines whether the row is copied to the Final_Translation_Table**
- Formula outputs: `Final_Input`, `Final_Output`, `Publish_Eligible`, `Decision_Warning` (these show whether the row is ready to publish)

Only `Decision` should be edited. The sheet is intentionally unprotected so sort/filter works reliably, so avoid editing formula/system columns.
The sheet freezes the header row only, so horizontal scrolling should remain usable.
**Review_Workbench is sortable and filterable** - use the header dropdowns to sort or filter rows as you work.

### Validate decision options (Review_Workbench)

Each decision determines whether the row is included in the Final_Translation_Table:

| Decision | Effect |
|----------|--------|
| **Keep As-Is** | Row is copied to Final_Translation_Table with current keys. No changes. The row still participates in duplicate-key QA checks—if it creates a duplicate final key with another row (that is not Allow One-to-Many), QA will fail. |
| **Use Suggestion** | Row is copied to Final_Translation_Table with the effective suggestion key applied on the side shown by Update Side. The effective key is `Manual_Suggested_Key` if you enter one, otherwise `Suggested_Key` from candidate selection. Verify the key and context before applying. |
| **Allow One-to-Many** | Row is copied to Final_Translation_Table as an intentional one-to-many or many-to-one exception. The row is **excluded** from duplicate-key QA checks—use this when you deliberately want multiple rows with the same final key (e.g., several campus keys mapping to one parent org). |
| **Ignore** | Row is **not** copied to Final_Translation_Table. Kept unresolved for later review. |

**When to choose Keep As-Is vs Allow One-to-Many for duplicate rows:** Use **Keep As-Is** when the row is correct and does not create a duplicate. Use **Allow One-to-Many** when the row is correct *and* you are intentionally approving a duplicate key (one-to-many or many-to-one mapping).

**Intentional many-to-one (Duplicate_Target + Keep As-Is):** When multiple Outcomes campus keys map to one myWSU parent key (e.g., Texas A&M campuses → Texas A&M main), you can use **Keep As-Is** on each Duplicate_Target row. The duplicate output QA check (B14) exempts rows where Source_Sheet=One_to_Many, Error_Type=Duplicate_Target, and Decision=Keep As-Is. Duplicate input (B13) and duplicate pair (B15) checks still block as usual.

### How to edit rows to get a clean final table

1. Go to `Review_Workbench`.
2. Work top to bottom, or filter to rows where `Decision` is blank.
3. For each row, choose one Decision from the table above.
4. If you choose `Use Suggestion`, confirm:
   - Either `Suggested_Key` is filled (from candidate selection), or you enter a valid key in **Manual_Suggested_Key**.
   - `Update Side` is `Input`, `Output`, or `Both` (not `None`).
   - For risky decisions, select a **Reason Code**:
     - `Use Suggestion` with manual key
     - `Allow One-to-Many`
     - `Keep As-Is` on `Duplicate_Target` rows (`Source_Sheet=One_to_Many`)
5. Check the same row:
   - `Decision_Warning` should be blank.
   - `Final_Input` and `Final_Output` should be filled for rows you want to publish.
6. Verify outputs:
   - `Final_Translation_Table` shows approved rows.
   - `Translation_Key_Updates` (What changed) shows only changed keys—use this to verify all updates before publish.
   - `QA_Checks_Validate` has no blocking checks.

### Manual override when candidates are blank or unusable

When Outcomes has many campuses but myWSU has only one parent key, the candidate list may be blank or unusable. You can still resolve the row:

1. Enter the correct myWSU key in **Manual_Suggested_Key** (must exist in the myWSU key list).
2. Set **Update Side** to `Output` (or `Both` if updating both sides).
3. Choose **Use Suggestion**.
4. Ensure `Decision_Warning` is blank. If you see "Invalid manual key: not found in valid keys", the key is not in the valid myWSU keys—correct it or use a different key.

### Resuming a prior session (re-import)

If you upload a **Prior Validate workbook** (Excel from a previous Validate run), the app will re-apply your prior decisions where `Review_Row_ID` matches. After download, the progress area shows: *Re-import: X applied, Y conflicts, Z new rows, W orphaned.* Conflicts occur when a prior "Use Suggestion" key is no longer valid in the new context. New rows are rows in the current run with no prior decision; orphaned rows are prior decisions for rows no longer in the current run.

### What happens automatically

- `Valid` and `High_Confidence_Match` rows are **auto-matched** and appear in `Final_Translation_Table` without any decision from you.
- For rows that need review, **default decisions** are pre-populated when the best choice is obvious (for example, Name_Mismatch with good score -> Keep As-Is; Output_Not_Found with no replacement -> Ignore). You can still change any decision.
- `Final_Translation_Table` = auto-matched rows + rows where your Review_Workbench decision is Keep As-Is, Use Suggestion, or Allow One-to-Many (not Ignore).
- Columns: Review Row ID, your selected Outcomes columns, Translate Input, Translate Output, your selected myWSU columns.
- `Translation_Key_Updates` (What changed) shows only rows where final keys differ from current keys—your primary verification sheet before publish.

### Final step before updating Outcomes Translation Table

When you are done reviewing and the QA gate passes:

1. **Copy the Final_Translation_Table** - select all data rows and copy.
2. **Paste as values only** into a new sheet (Paste Special -> Values). The Final_Translation_Table is formula-driven and cannot be sorted directly.
3. **Sort and double-check** the pasted values in the new sheet.
4. **Only then** update the Outcomes Translation Table with the verified data.

Do not update the Outcomes Translation Table until you have completed these steps.

### Publish rule of thumb

Use `QA_Checks_Validate` before publishing.
Publish only when gate checks are clean (`PASS`) or you intentionally accept documented exceptions.
Gate-blocking checks include:

- unresolved actions
- overflow beyond formula capacity
- blank finals on publish-eligible rows
- `Use Suggestion` without effective key (needs Manual_Suggested_Key or Selected_Candidate_ID + Suggested_Key)
- `Use Suggestion` with invalid Update Side
- `Use Suggestion` with invalid manual key (key not in valid myWSU/Outcomes keys)
- `Use Suggestion` no-op (effective key equals current value; no change)
- risky decisions without reason code (B10)
- duplicate final input keys (B13)
- duplicate final output keys (B14), excluding Allow One-to-Many and intentional Duplicate_Target+Keep As-Is
- duplicate (input, output) pairs (B15)

**Duplicate input (B13) and duplicate pair (B15) checks still block**—only the duplicate output check (B14) has the narrow exemption for intentional many-to-one.

## Workflow 2: Create a new translation table

Use this when you are starting from Outcomes + myWSU and need a new mapping table.

### Files required

- Outcomes source file
- myWSU source file

### Match method

- `Match by key columns`: use when both sides have reliable keys.
- `Match by name columns`: use when key mapping is missing/unreliable.

Important: in Create + name mode, key radio selections are optional and ignored.

### Steps in the app

1. Select `Create`.
2. Upload Outcomes and myWSU.
3. Choose included columns and optional role mapping.
4. Pick match method (`key` or `name`).
5. If name mode, pick name columns and set threshold/ambiguity gap.
6. Click `Generate Translation Table`.
7. Open the downloaded Excel workbook.

### Create Excel review order

See the `Review_Instructions_Create` sheet in the workbook for detailed guidance. Recommended order:

1. `Summary`
2. `New_Translation_Candidates`
3. `Ambiguous_Candidates`
4. `Missing_In_myWSU`
5. `Missing_In_Outcomes` (diagnostic)
6. `Review_Decisions`
7. `Final_Translation_Table`
8. `QA_Checks`

### Create decision meanings (`Review_Decisions`)

- `Accept`: keep proposed match.
- `Choose Alternate`: pick `Alt 1/2/3`.
- `No Match`: unresolved.

`Final_myWSU_Key` and `Final_myWSU_Name` are formula-driven from your decision.

## Workflow 3: Join Preview

Use this when you have updated a translation table (via Validate or manually) and want a final visual check before publishing to Outcomes. Join Preview produces a single flat table showing Outcomes data, translate keys, and myWSU data side by side—no validation logic, no decisions.

### Files required

- Outcomes source file
- Translation table file
- myWSU source file

### Steps in the app

1. Select `Join Preview`.
2. Upload all 3 files.
3. In `Select Columns, Keys, and Roles`:
   - Pick key columns for Outcomes, Translation input/output, and myWSU.
   - Pick which Outcomes and myWSU columns to include (e.g., name, state, country).
4. Click `Generate Join Preview`.
5. Download the Excel file.

### What the Join Preview Excel contains

- **Join_Preview** sheet: one row per translate row, with Outcomes columns | translate_input, translate_output | myWSU columns.
- Key-only matching. If a key has no match, that side's columns are blank.
- Scan the table to confirm Outcomes and myWSU data align correctly before you update the Outcomes Translation Table.

## Practical review tips

- Review by exception, not row by row.
- Start with high-priority structural/key issues.
- Process in passes (for example 500 to 1000 rows at a time).
- Always check QA sheets before publish.

## Common mistakes

- Wrong key column selected.
- Name comparison enabled with wrong name columns.
- Not mapping city/state/country roles when needed.
- Treating `Missing_In_Outcomes` as publishable rows (it is diagnostic).

## Privacy

Your uploaded files are processed locally in your browser and are not uploaded by this app. The page loads JavaScript libraries from public CDNs.
