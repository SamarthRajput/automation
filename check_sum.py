import pandas as pd
import os


def build_checksum_str(indices):
    """
    Convert list of 0-based df indices to Excel row checksum string.
    Consecutive rows → 'first;last'  (e.g. rows 2,3,4 → '2;4')
    Non-consecutive  → 'r1 + r2'     (e.g. rows 2,5  → '2 + 5')
    Mixed segments   → '2;4 + 7;9'
    """
    if not indices:
        return ""
    excel_rows = sorted([i + 2 for i in indices])
    if len(excel_rows) == 1:
        return str(excel_rows[0])

    segments = []
    seg_start = excel_rows[0]
    seg_end   = excel_rows[0]
    for r in excel_rows[1:]:
        if r == seg_end + 1:
            seg_end = r
        else:
            segments.append((seg_start, seg_end))
            seg_start = r
            seg_end   = r
    segments.append((seg_start, seg_end))

    return " + ".join(f"{s};{e}" if s != e else str(s) for s, e in segments)


def vals_match(df, source_indices, target_idx, year_cols, tol=2):
    """Check if sum of source rows equals target row (within tolerance)."""
    target = df.loc[target_idx, year_cols].apply(pd.to_numeric, errors="coerce").dropna()
    if target.empty:
        return False
    total  = df.loc[source_indices, year_cols].apply(pd.to_numeric, errors="coerce").sum()
    common = target.index.intersection(total.index)
    if common.empty:
        return False
    return all(abs(target[c] - total[c]) < tol for c in common)


def find_checksum(df, target_idx, year_cols, tol=2):
    """
    Dynamically find which rows sum to the target row and return checksum string.

    Search order:
      1. Rows in same dim_1_name group (since last subtotal) — handles P&L / Cash Flow groups
      2. Rows in same dim_2_name group (column-based equity grouping)
      3. Sliding window: last K True (subtotal) rows + any ungrouped False rows
      4. Last N raw rows (small sliding window fallback)
    """
    target_vals = df.loc[target_idx, year_cols].apply(pd.to_numeric, errors="coerce").dropna()
    if target_vals.empty:
        return ""  # no numeric data — skip

    table_id   = df.loc[target_idx, "table_id"]
    dim1       = df.loc[target_idx, "dim_1_name"]
    dim2       = df.loc[target_idx, "dim_2_name"]
    same_table = df[df["table_id"] == table_id]

    def try_match(candidates):
        if not candidates:
            return None
        if vals_match(df, candidates, target_idx, year_cols, tol):
            return build_checksum_str(candidates)
        return None

    # ── Strategy 1: same dim_1_name group (P&L / Cash Flow sections) ──
    if pd.notna(dim1):
        dim1_above = same_table[
            (same_table["dim_1_name"] == dim1) & (same_table.index < target_idx)
        ].index.tolist()

        true_in_group = [i for i in dim1_above if df.loc[i, "comments"] == True]

        if true_in_group:
            last_true    = max(true_in_group)
            since_last   = [i for i in dim1_above if i > last_true]
            with_subtotal = sorted([last_true] + since_last)
            for c in [since_last, with_subtotal, dim1_above]:
                r = try_match(c)
                if r is not None:
                    return r
        else:
            r = try_match(dim1_above)
            if r is not None:
                return r

    # ── Strategy 2: same dim_2_name group (equity column slices) ──
    if pd.notna(dim2):
        dim2_above = same_table[
            (same_table["dim_2_name"] == dim2) & (same_table.index < target_idx)
        ].index.tolist()
        true_d2  = [i for i in dim2_above if df.loc[i, "comments"] == True]
        false_d2 = [i for i in dim2_above if df.loc[i, "comments"] == False]
        for c in [dim2_above, true_d2, false_d2]:
            r = try_match(c)
            if r is not None:
                return r

    # ── Strategy 3: K most-recent True rows + ungrouped False rows ──
    all_above      = same_table[same_table.index < target_idx].index.tolist()
    true_above     = [i for i in all_above if df.loc[i, "comments"] == True]
    ungrouped_false = [
        i for i in all_above
        if df.loc[i, "comments"] == False and pd.isna(df.loc[i, "dim_1_name"])
    ]

    for k in range(1, len(true_above) + 1):
        subset = sorted(set(true_above[-k:] + ungrouped_false))
        r = try_match(subset)
        if r is not None:
            return r

    # ── Strategy 4: raw sliding window (last N rows) ──
    for n in range(2, min(len(all_above) + 1, 10)):
        r = try_match(all_above[-n:])
        if r is not None:
            return r

    return ""  # standalone row — opening balances, reference rows, etc.


def fill_checksum(input_file, output_file):
    df = pd.read_excel(input_file, engine="openpyxl")
    year_cols = [c for c in df.columns if str(c).isdigit()]

    checksums = []
    for idx, row in df.iterrows():
        if row.get("comments") == True:
            cs = find_checksum(df, idx, year_cols)
            label = f"✅ '{cs}'" if cs else "⬜  standalone"
            print(f"  Excel Row {idx+2:>3} | {str(row['metric_name']):<58} | {label}")
            checksums.append(cs)
        else:
            checksums.append("")

    df["check_sum"] = checksums
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n✅ Done! Output saved to: {output_file}")


# ── Entry Point ──
input_path  = os.path.expanduser("~/Downloads/Untitled_spreadsheet.xlsx")
output_path = os.path.expanduser("~/Downloads/output_with_checksum.xlsx")

fill_checksum(input_path, output_path)