#!/usr/bin/env python3
"""
Audit Review_Workbench decisions against Final_Translation_Table.
Verifies that decisions are accurately reflected in the final table.
"""
import pandas as pd
import sys
from pathlib import Path

def normalize_key(k):
    if pd.isna(k): return ''
    s = str(k).strip()
    # Normalize numbers: 11285619.0 and 11285619 and "11285619" -> same key
    if '.' in s and s.replace('.', '').isdigit():
        try:
            return str(int(float(s)))
        except (ValueError, TypeError):
            pass
    return s

def main():
    xlsx_path = Path(r"C:\Python Projects\wsu-gradschool-tools\WSU_Mapping_Validation_Report_gjc.xlsx")
    if not xlsx_path.exists():
        print(f"File not found: {xlsx_path}")
        sys.exit(1)

    # Load Review_Workbench - first 681 data rows (row 2+ since row 1 is header)
    review_df = pd.read_excel(xlsx_path, sheet_name="Review_Workbench", header=0)
    review_df = review_df.head(681)

    # Find column names (actual export uses "Final Translate Input", "Current Translate Input", etc.)
    def find_col(df, candidates):
        for c in candidates:
            for col in df.columns:
                if str(col).strip().lower() == c.lower():
                    return col
        return None

    decision_col = find_col(review_df, ['Decision', 'decision'])
    final_input_col = find_col(review_df, ['Final Translate Input', 'Final_Input'])
    final_output_col = find_col(review_df, ['Final Translate Output', 'Final_Output'])
    current_input_col = find_col(review_df, ['Current Translate Input', 'Current_Input', 'Input (Translate Input)'])
    current_output_col = find_col(review_df, ['Current Translate Output', 'Current_Output', 'Output (Translate Output)'])
    suggested_col = find_col(review_df, ['Suggested Key', 'Suggested_Key', 'suggested_key'])
    update_side_col = find_col(review_df, ['Update Side', 'Key_Update_Side', 'key_update_side'])
    publish_eligible_col = find_col(review_df, ['Publish Eligible (1=yes)', 'Publish_Eligible', 'publish_eligible'])

    print("Review_Workbench columns detected:")
    print(f"  Decision: {decision_col}")
    print(f"  Final_Input: {final_input_col}")
    print(f"  Final_Output: {final_output_col}")
    print(f"  Current_Input: {current_input_col}")
    print(f"  Current_Output: {current_output_col}")
    print(f"  Suggested_Key: {suggested_col}")
    print(f"  Key_Update_Side: {update_side_col}")
    print(f"  Publish_Eligible: {publish_eligible_col}")
    print()

    # Load Final_Translation_Table
    final_df = pd.read_excel(xlsx_path, sheet_name="Final_Translation_Table", header=0)
    final_input_ft = find_col(final_df, ['Translate Input', 'Final_Input', 'translate_input'])
    final_output_ft = find_col(final_df, ['Translate Output', 'Final_Output', 'translate_output'])
    if final_input_ft is None or final_output_ft is None:
        # Fallback
        for c in final_df.columns:
            if 'input' in str(c).lower() and final_input_ft is None:
                final_input_ft = c
            if 'output' in str(c).lower() and final_output_ft is None:
                final_output_ft = c

    print("Final_Translation_Table key columns:", final_input_ft, final_output_ft)
    print()

    # Build set of (Input, Output) in Final_Translation_Table
    final_pairs = set()
    for _, row in final_df.iterrows():
        inp = normalize_key(row[final_input_ft] if final_input_ft in row.index else '')
        out = normalize_key(row[final_output_ft] if final_output_ft in row.index else '')
        if inp or out:
            final_pairs.add((inp, out))

    print(f"Final_Translation_Table has {len(final_pairs)} unique (input, output) pairs")
    print()

    # Audit each review row
    publishable = {'Keep As-Is', 'Use Suggestion', 'Allow One-to-Many'}
    issues = []
    ignore_count = 0
    publishable_count = 0
    blank_decision = 0

    for idx, row in review_df.iterrows():
        review_row_num = idx + 2  # 1-based + header
        decision = str(row.get(decision_col, '')).strip()
        if not decision:
            blank_decision += 1
            continue
        if decision == 'Ignore':
            ignore_count += 1
            # For Ignore, Final columns are blank. Skip verification (pair should be excluded).
            continue
        if decision in publishable:
            publishable_count += 1
            fin_inp = normalize_key(row[final_input_col] if final_input_col and final_input_col in row.index else '')
            fin_out = normalize_key(row[final_output_col] if final_output_col and final_output_col in row.index else '')
            if not fin_inp and current_input_col:
                fin_inp = normalize_key(row.get(current_input_col, ''))
            if not fin_out and current_output_col:
                fin_out = normalize_key(row.get(current_output_col, ''))
            if not fin_inp and not fin_out:
                issues.append({
                    'row': review_row_num,
                    'type': 'BLANK_FINAL',
                    'decision': decision,
                    'msg': f"Row {review_row_num}: Decision={decision} but Final_Input and Final_Output are blank (formula issue?)"
                })
            elif (fin_inp, fin_out) not in final_pairs:
                issues.append({
                    'row': review_row_num,
                    'type': 'MISSING_FROM_FINAL',
                    'decision': decision,
                    'msg': f"Row {review_row_num}: Decision={decision} but (Final_Input, Final_Output)=({fin_inp}, {fin_out}) NOT in Final_Translation_Table"
                })

    print("=== AUDIT SUMMARY ===")
    print(f"Review rows with decisions: {publishable_count + ignore_count}")
    print(f"  Publishable (Keep As-Is / Use Suggestion / Allow One-to-Many): {publishable_count}")
    print(f"  Ignore: {ignore_count}")
    print(f"  Blank decision (skipped): {blank_decision}")
    print()
    print(f"Final_Translation_Table unique pairs: {len(final_pairs)}")
    print()
    if issues:
        print(f"=== ISSUES FOUND: {len(issues)} ===")
        for i in issues[:50]:
            print(i['msg'])
        if len(issues) > 50:
            print(f"... and {len(issues) - 50} more")
    else:
        print("=== NO ISSUES: Decisions accurately reflected in Final_Translation_Table ===")

if __name__ == "__main__":
    main()
