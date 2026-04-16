import pandas as pd
import os


def build_checksum_str(positions):
    if not positions:
        return ""
    rows = sorted(positions)
    if len(rows) == 1:
        return str(rows[0])

    segments = []
    seg_start = seg_end = rows[0]
    for r in rows[1:]:
        if r == seg_end + 1:
            seg_end = r
        else:
            segments.append((seg_start, seg_end))
            seg_start = seg_end = r
    segments.append((seg_start, seg_end))

    return " + ".join(f"{s};{e}" if s != e else str(s) for s, e in segments)


def is_true(val):
    """Robustly checks if a comments cell means True — handles bool, int, str, float."""
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val == 1
    if isinstance(val, str):
        return val.strip().lower() in ("true", "1", "yes")
    return False


def vals_match(df, source_indices, target_idx, year_cols, tol=2):
    target = df.loc[target_idx, year_cols].apply(pd.to_numeric, errors="coerce").dropna()
    if target.empty:
        return False
    total  = df.loc[source_indices, year_cols].apply(pd.to_numeric, errors="coerce").sum()
    common = target.index.intersection(total.index)
    if common.empty:
        return False
    return all(abs(target[c] - total[c]) < tol for c in common)


def find_checksum(df, target_idx, year_cols, global_to_excel_row, tol=2):
    target_vals = df.loc[target_idx, year_cols].apply(pd.to_numeric, errors="coerce").dropna()
    if target_vals.empty:
        return ""

    table_id = df.loc[target_idx, "table_id"]
    dim1     = df.loc[target_idx, "dim_1_name"]
    dim2     = df.loc[target_idx, "dim_2_name"]

    same_table = df[df["table_id"] == table_id]

    def to_excel_rows(indices):
        return [global_to_excel_row[i] for i in indices]

    def try_match(candidates):
        if not candidates:
            return None
        if vals_match(df, candidates, target_idx, year_cols, tol):
            return build_checksum_str(to_excel_rows(candidates))
        return None

    # ── Strategy 1: same dim_1_name group ──
    if pd.notna(dim1):
        dim1_above    = same_table[(same_table["dim_1_name"] == dim1) & (same_table.index < target_idx)].index.tolist()
        true_in_group = [i for i in dim1_above if is_true(df.loc[i, "comments"])]  # ← FIXED

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

    # ── Strategy 2: same dim_2_name group ──
    if pd.notna(dim2):
        dim2_above = same_table[(same_table["dim_2_name"] == dim2) & (same_table.index < target_idx)].index.tolist()
        true_d2    = [i for i in dim2_above if is_true(df.loc[i, "comments"])]      # ← FIXED
        false_d2   = [i for i in dim2_above if not is_true(df.loc[i, "comments"])]  # ← FIXED
        for c in [dim2_above, true_d2, false_d2]:
            r = try_match(c)
            if r is not None:
                return r

    # ── Strategy 3: last K True rows + ungrouped False rows ──
    all_above       = same_table[same_table.index < target_idx].index.tolist()
    true_above      = [i for i in all_above if is_true(df.loc[i, "comments"])]      # ← FIXED
    ungrouped_false = [
        i for i in all_above
        if not is_true(df.loc[i, "comments"]) and pd.isna(df.loc[i, "dim_1_name"])  # ← FIXED
    ]

    for k in range(1, len(true_above) + 1):
        subset = sorted(set(true_above[-k:] + ungrouped_false))
        r = try_match(subset)
        if r is not None:
            return r

    # ── Strategy 4: raw sliding window fallback ──
    for n in range(2, min(len(all_above) + 1, 10)):
        r = try_match(all_above[-n:])
        if r is not None:
            return r

    return ""


def fill_checksum(input_file, output_file):
    df = pd.read_excel(input_file, engine="openpyxl")
    year_cols = [c for c in df.columns if str(c).isdigit()]

    # ── Normalize comments column to proper booleans ──
    df["comments"] = df["comments"].apply(is_true)  # ← KEY FIX

    global_to_excel_row = {
        global_idx: pos + 2
        for pos, global_idx in enumerate(df.index)
    }

    checksums = []
    for idx, row in df.iterrows():
        if is_true(row.get("comments")):  # ← FIXED (was == True)
            cs = find_checksum(df, idx, year_cols, global_to_excel_row)
            label = f"✅  '{cs}'" if cs else "⬜  standalone (no source rows)"
            print(f"  [{row['table_id']}]  ExcelRow {global_to_excel_row[idx]:>3}  {str(row['metric_name']):<55} {label}")
            checksums.append(cs)
        else:
            checksums.append("")

    df["check_sum"] = checksums
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n✅ Done! Saved to: {output_file}")


# ── Entry Point ──
input_path  = os.path.expanduser("~/Downloads/abu_dhabi_port_company_quantitative.xlsx")
output_path = os.path.expanduser("~/Downloads/output_with_checksum.xlsx")

fill_checksum(input_path, output_path)