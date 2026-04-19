import pandas as pd
import numpy as np
import os
from itertools import combinations

# Optimised and fast way to mark checksum

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
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val == 1
    if isinstance(val, str):
        return val.strip().lower() in ("true", "1", "yes")
    return False


# ─────────────────────────────────────────────────────────────
# Pre-cache numeric matrix: index → numpy array of year values
# Built once per fill_checksum call, reused across all rows
# ─────────────────────────────────────────────────────────────
def build_numeric_cache(df, year_cols):
    cache = {}
    for idx in df.index:
        vals = pd.to_numeric(df.loc[idx, year_cols], errors="coerce").values.astype(float)
        cache[idx] = vals  # NaN where non-numeric
    return cache


def vals_match_fast(numeric_cache, source_indices, target_idx, tol=2):
    """Numpy-based sum check — avoids pandas overhead per call."""
    target = numeric_cache[target_idx]
    valid  = ~np.isnan(target)
    if not valid.any():
        return False

    total = np.zeros(len(target))
    for i in source_indices:
        src = numeric_cache[i]
        total += np.where(np.isnan(src), 0, src)

    diffs = np.abs(target[valid] - total[valid])
    return bool((diffs < tol).all())


COMBO_LIMIT = 15  # full subset search for pools ≤ this size


def search_all_subsets(pool, target_idx, numeric_cache, global_to_excel_row, tol=2):
    """
    For small pools (≤ COMBO_LIMIT): full combinatorial search, largest first.
    For large pools (> COMBO_LIMIT): contiguous window search only — O(n²).
    No redundant re-parsing — uses pre-cached numpy arrays.
    """
    if not pool:
        return None
    pool = sorted(set(pool))
    n    = len(pool)

    if n <= COMBO_LIMIT:
        for size in range(n, 0, -1):
            for combo in combinations(pool, size):
                if vals_match_fast(numeric_cache, list(combo), target_idx, tol):
                    return build_checksum_str([global_to_excel_row[i] for i in combo])
    else:
        # Large pool — contiguous windows (financial groups are almost always contiguous)
        for size in range(n, 0, -1):
            for start in range(n - size + 1):
                candidates = pool[start: start + size]
                if vals_match_fast(numeric_cache, candidates, target_idx, tol):
                    return build_checksum_str([global_to_excel_row[i] for i in candidates])

    return None


def find_checksum(df, target_idx, year_cols, global_to_excel_row, numeric_cache, tol=2):
    target_arr = numeric_cache[target_idx]
    if np.all(np.isnan(target_arr)):
        return ""

    table_id = df.loc[target_idx, "table_id"]
    dim1     = df.loc[target_idx, "dim_1_name"]
    dim2     = df.loc[target_idx, "dim_2_name"]

    same_table = df[df["table_id"] == table_id]
    all_above  = same_table[same_table.index < target_idx].index.tolist()

    if not all_above:
        return ""

    tried = set()

    def attempt(pool):
        pool = sorted(set(pool))
        key  = tuple(pool)
        if not pool or key in tried:
            return None
        tried.add(key)
        return search_all_subsets(pool, target_idx, numeric_cache, global_to_excel_row, tol)

    # ── STAGE 1: dim_1_name group ──
    if pd.notna(dim1):
        dim1_rows  = same_table[
            (same_table["dim_1_name"] == dim1) & (same_table.index < target_idx)
        ].index.tolist()
        dim1_false = [i for i in dim1_rows if not is_true(df.loc[i, "comments"])]
        dim1_true  = [i for i in dim1_rows if is_true(df.loc[i, "comments"])]

        r = attempt(dim1_false);                    
        if r is not None: return r
        r = attempt(dim1_true)
        if r is not None: return r

        if dim1_true:
            last_true  = max(dim1_true)
            since_last = [i for i in dim1_rows if i > last_true]
            r = attempt(since_last)
            if r is not None: return r
            r = attempt(sorted([last_true] + since_last))
            if r is not None: return r

        r = attempt(dim1_rows)
        if r is not None: return r

    # ── STAGE 2: dim_2_name group ──
    if pd.notna(dim2):
        dim2_rows  = same_table[
            (same_table["dim_2_name"] == dim2) & (same_table.index < target_idx)
        ].index.tolist()
        dim2_false = [i for i in dim2_rows if not is_true(df.loc[i, "comments"])]
        dim2_true  = [i for i in dim2_rows if is_true(df.loc[i, "comments"])]

        for pool in [dim2_false, dim2_true, dim2_rows]:
            r = attempt(pool)
            if r is not None: return r

    # ── STAGE 3: cross-dim ──
    true_above  = [i for i in all_above if is_true(df.loc[i, "comments"])]
    false_above = [i for i in all_above if not is_true(df.loc[i, "comments"])]

    r = attempt(true_above)
    if r is not None: return r
    r = attempt(false_above)
    if r is not None: return r

    ungrouped_false = [
        i for i in false_above
        if pd.isna(df.loc[i, "dim_1_name"]) and pd.isna(df.loc[i, "dim_2_name"])
    ]
    for k in range(1, len(true_above) + 1):
        r = attempt(sorted(set(true_above[-k:]) | set(ungrouped_false)))
        if r is not None: return r

    # ── STAGE 4: full table above ──
    r = attempt(all_above)
    if r is not None: return r

    return ""


def fill_checksum(input_file, output_file):
    df = pd.read_excel(input_file, engine="openpyxl")
    year_cols = [c for c in df.columns if str(c).isdigit()]

    df["comments"] = df["comments"].apply(is_true)

    # Build numeric cache ONCE — reused for every row
    numeric_cache = build_numeric_cache(df, year_cols)

    global_to_excel_row = {
        global_idx: pos + 2
        for pos, global_idx in enumerate(df.index)
    }

    checksums  = []
    standalone = []

    for idx, row in df.iterrows():
        if is_true(row.get("comments")):
            cs = find_checksum(df, idx, year_cols, global_to_excel_row, numeric_cache)
            if cs:
                label = f"✅  '{cs}'"
            else:
                label = "⬜  standalone"
                standalone.append((row["table_id"], global_to_excel_row[idx], row["metric_name"]))
            print(
                f"  [{row['table_id']}]  "
                f"ExcelRow {global_to_excel_row[idx]:>3}  "
                f"dim1={str(row.get('dim_1_name', '')):<20}  "
                f"{str(row['metric_name']):<50}  {label}"
            )
            checksums.append(cs)
        else:
            checksums.append("")

    df["check_sum"] = checksums
    df.to_excel(output_file, index=False, engine="openpyxl")
    print(f"\n✅ Done! Saved to: {output_file}")

    if standalone:
        print(f"\n⚠️  {len(standalone)} rows with no checksum:")
        for tid, erow, name in standalone:
            print(f"     [{tid}]  ExcelRow {erow:>3}  {name}")


# ── Entry Point ──
input_path  = os.path.expanduser("~/Downloads/abu_dhabi_port_company_quantitative.xlsx")
output_path = os.path.expanduser("~/Downloads/output_optimised.xlsx")

fill_checksum(input_path, output_path)