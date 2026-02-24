import pandas as pd
import os


def build_checksum_str(local_positions):
    """
    Build checksum string from local row positions within a table.
    Consecutive rows  → 'first;last'   e.g. [2,3,4] → '2;4'
    Non-consecutive   → 'r1 + r2'      e.g. [2,5]   → '2 + 5'
    Mixed             → '2;4 + 7;9'
    """
    if not local_positions:
        return ""
    rows = sorted(local_positions)
    if len(rows) == 1:
        return str(rows[0])

    segments = []
    seg_start = rows[0]
    seg_end   = rows[0]
    for r in rows[1:]:
        if r == seg_end + 1:
            seg_end = r
        else:
            segments.append((seg_start, seg_end))
            seg_start = r
            seg_end   = r
    segments.append((seg_start, seg_end))

    return " + ".join(f"{s};{e}" if s != e else str(s) for s, e in segments)


def vals_match(df, source_indices, target_idx, year_cols, tol=2):
    """Returns True if sum of source rows equals target row values (within tol)."""
    target = df.loc[target_idx, year_cols].apply(pd.to_numeric, errors="coerce").dropna()
    if target.empty:
        return False
    total  = df.loc[source_indices, year_cols].apply(pd.to_numeric, errors="coerce").sum()
    common = target.index.intersection(total.index)
    if common.empty:
        return False
    return all(abs(target[c] - total[c]) < tol for c in common)


def find_checksum(df, target_idx, year_cols, global_to_local, tol=2):
    """
    Dynamically finds which rows (within the SAME table_id only) sum to the
    target row, then returns a checksum string using LOCAL row numbers.

    Row numbers restart from 2 for each new table_id (row 1 = header).

    Search strategies (in order):
      1. Rows in same dim_1_name group since last subtotal
      2. Rows in same dim_2_name group (equity column slices)
      3. Last K subtotal (True) rows + ungrouped detail rows
      4. Raw sliding window of last N rows (fallback)
    """
    target_vals = df.loc[target_idx, year_cols].apply(pd.to_numeric, errors="coerce").dropna()
    if target_vals.empty:
        return ""

    table_id   = df.loc[target_idx, "table_id"]
    dim1       = df.loc[target_idx, "dim_1_name"]
    dim2       = df.loc[target_idx, "dim_2_name"]

    # ── STRICT: only search within the same table_id ──
    same_table = df[df["table_id"] == table_id]

    def to_local(indices):
        return [global_to_local[i] for i in indices]

    def try_match(candidates):
        if not candidates:
            return None
        if vals_match(df, candidates, target_idx, year_cols, tol):
            return build_checksum_str(to_local(candidates))
        return None

    # ── Strategy 1: same dim_1_name group (P&L / Cash Flow sections) ──
    if pd.notna(dim1):
        dim1_above = same_table[
            (same_table["dim_1_name"] == dim1) & (same_table.index < target_idx)
        ].index.tolist()
        true_in_group = [i for i in dim1_above if df.loc[i, "comments"] == True]

        if true_in_group:
            last_true     = max(true_in_group)
            since_last    = [i for i in dim1_above if i > last_true]
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

    # ── Strategy 3: last K True rows + ungrouped False rows (within same table) ──
    all_above       = same_table[same_table.index < target_idx].index.tolist()
    true_above      = [i for i in all_above if df.loc[i, "comments"] == True]
    ungrouped_false = [
        i for i in all_above
        if df.loc[i, "comments"] == False and pd.isna(df.loc[i, "dim_1_name"])
    ]

    for k in range(1, len(true_above) + 1):
        subset = sorted(set(true_above[-k:] + ungrouped_false))
        r = try_match(subset)
        if r is not None:
            return r

    # ── Strategy 4: raw sliding window fallback (within same table) ──
    for n in range(2, min(len(all_above) + 1, 10)):
        r = try_match(all_above[-n:])
        if r is not None:
            return r

    return ""  # standalone row — no source rows found (e.g. opening balances)


def fill_checksum(input_file, output_file):
    df = pd.read_excel(input_file, engine="openpyxl")
    year_cols = [c for c in df.columns if str(c).isdigit()]

    # ── Build global_index → local_row_number map (resets per table_id) ──
    global_to_local = {}
    for table_id, group in df.groupby("table_id", sort=False):
        for local_pos, global_idx in enumerate(group.index):
            global_to_local[global_idx] = local_pos + 2  # row 1 = header

    checksums = []
    for idx, row in df.iterrows():
        if row.get("comments") == True:
            cs = find_checksum(df, idx, year_cols, global_to_local)
            label = f"✅  '{cs}'" if cs else "⬜  standalone (no source rows)"
            print(f"  [{row['table_id']}]  Row {global_to_local[idx]:>3}  {str(row['metric_name']):<55} {label}")
            checksums.append(cs)
        else:
            checksums.append("")

    df["check_sum"] = checksums
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n✅ Done! Saved to: {output_file}")


# ── Entry Point ──
input_path  = os.path.expanduser("~/Downloads/Untitled-spreadsheet.xlsx")
output_path = os.path.expanduser("~/Downloads/output_with_checksum.xlsx")

fill_checksum(input_path, output_path)
