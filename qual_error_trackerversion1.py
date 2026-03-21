"""
Qual RI Error Rate Tracker  (v6 – hardcoded path, no UI popups)
================================================================
▶  CONFIGURE THESE TWO LINES BEFORE RUNNING:
       FILE_PATH  – full path to your Transition Excel sheet
       SHEET_NAME – exact sheet tab name inside that file

Usage:
    pip install pandas openpyxl xlsxwriter
    python qual_error_tracker.py
"""

import re
import pandas as pd
import os, sys

# ══════════════════════════════════════════════════════════════════════════════
# ▼▼▼  CHANGE THESE TWO LINES  ▼▼▼
# ══════════════════════════════════════════════════════════════════════════════

FILE_PATH  = r"C:\Users\SamarthRajput\Downloads\Transition_Sheet_06032026.xlsx"
SHEET_NAME = "Transformation Transition"   # exact tab name

# ══════════════════════════════════════════════════════════════════════════════
# ▲▲▲  THAT'S ALL YOU NEED TO EDIT  ▲▲▲
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
# ISSUE CATEGORISATION
#
# Priority order (checked top to bottom, first match wins):
#   1.  Wrong Page Number
#   2.  Wrong Table ID
#   3.  Mapping Issue          ← checked before underscore catch-all
#   4.  Wrong Note ID
#   5.  Missing Table
#   6.  Duplicate / Extra Table
#   7.  Wrong Data / Value
#   8.  Wrong Currency / Unit
#   9.  Wrong Period / Year
#  10.  Language / Translation Issue
#  11.  Wrong Report Number
#  12.  Should Be Excluded
#  13.  Other / Uncategorised
# ══════════════════════════════════════════════════════════════════════════════

def _has_underscore_value(text: str) -> bool:
    """True if text contains a snake_case token (word_word)."""
    return bool(re.search(r'\b\w+_\w+\b', text))


def categorise_issue(text: str) -> str:   # noqa: C901
    if not isinstance(text, str) or not text.strip():
        return "No Change"
    t = text.strip()

    # ── 1. Wrong Page Number ──────────────────────────────────────────────────
    PAGE_PATTERNS = [
        r"file\s*page",
        r"page\s*(num|no\.?|number)",
        r"change\s*(the\s*)?page",
        r"page.*from\s+\d+\s+to\s+\d+",
        r"from\s+\d+\s+to\s+\d+",
        r"\bpage\b.*\d+",
        r"\d+.*\bpage\b",
    ]
    if any(re.search(p, t, re.IGNORECASE) for p in PAGE_PATTERNS):
        return "Wrong Page Number"

    # ── 2. Wrong Table ID ─────────────────────────────────────────────────────
    TABLE_ID_PATTERNS = [
        r"change\s*(table[\s_]?id|table\s*id)",
        r"table[\s_]?id\s*(to|=|:)",
        r"change\s*(the\s*)?(name|id)\s+to\s+\S+_\S+",
        r"rename\s+(to\s+)?\S+_\S+",
        r"rename\s+\S+\s+to\s+\S+",
        r"change\s+name\s+to",
        r"change\s+id\s+to",
        r"update\s*(the\s*)?table[\s_]?id",
        r"replace\s*(the\s*)?table[\s_]?id",
        r"wrong\s*table[\s_]?id",
        r"incorrect\s*table[\s_]?id",
    ]
    if any(re.search(p, t, re.IGNORECASE) for p in TABLE_ID_PATTERNS):
        return "Wrong Table ID"

    # ── 3. Mapping Issue ──────────────────────────────────────────────────────
    # Must be BEFORE the underscore catch-all so that
    # "table present with old table_id" → Mapping Issue, not Wrong Table ID
    MAPPING_PATTERNS = [
        r"old\s*table[\s_]?id",
        r"table.*present.*old",
        r"present.*old.*table",
        r"add\s*(in|to)?\s*(the\s*)?current\s*year",
        r"present\s*(in)?\s*(the\s*)?current\s*year",
        r"current\s*year.*add",
        r"current\s*year.*present",
        r"wrong\s*mapping",
        r"mapping\s*wrong",
        r"remap",
        r"re-?map",
        r"wrong\s*structure",
        r"restructur",
        r"wrong\s*categor",
        r"mis-?categor",
        r"wrong\s*classif",
    ]
    if any(re.search(p, t, re.IGNORECASE) for p in MAPPING_PATTERNS):
        return "Mapping Issue"

    # snake_case value anywhere → Wrong Table ID (after Mapping check)
    if _has_underscore_value(t):
        return "Wrong Table ID"

    # ── 4. Wrong Note ID ──────────────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"note[\s_]?id", r"change\s*note", r"wrong\s*note",
            r"incorrect\s*note", r"update\s*note"]):
        return "Wrong Note ID"

    # ── 5. Missing Table ──────────────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"table.*missing", r"missing.*table", r"add.*table",
            r"table.*not\s*found", r"table.*absent", r"include.*table"]):
        return "Missing Table"

    # ── 6. Duplicate / Extra Table ────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"duplicate.*table", r"table.*duplicate", r"extra.*table",
            r"remove.*table", r"table.*already.*exists"]):
        return "Duplicate / Extra Table"

    # ── 7. Wrong Data / Value ─────────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"wrong\s*value", r"value\s*wrong", r"incorrect\s*value",
            r"incorrect\s*data", r"wrong\s*data", r"data\s*incorrect",
            r"wrong\s*(number|figure)", r"(number|figure)\s*wrong", r"mismatch"]):
        return "Wrong Data / Value"

    # ── 8. Wrong Currency / Unit ──────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"currency", r"wrong\s*unit", r"unit\s*wrong",
            r"denomination", r"wrong\s*scale", r"scale\s*wrong"]):
        return "Wrong Currency / Unit"

    # ── 9. Wrong Period / Year ────────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"wrong\s*period", r"period\s*wrong", r"wrong\s*year",
            r"year\s*wrong", r"incorrect\s*period", r"period\s*incorrect",
            r"change\s*period"]):
        return "Wrong Period / Year"

    # ── 10. Language / Translation ────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"translat", r"\blanguage\b", r"english\s*version"]):
        return "Language / Translation Issue"

    # ── 11. Wrong Report Number ───────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"report\s*(num|no\.?|number)", r"no\.?\s*of\s*report",
            r"wrong\s*report", r"report\s*wrong"]):
        return "Wrong Report Number"

    # ── 12. Should Be Excluded ────────────────────────────────────────────────
    if any(re.search(p, t, re.IGNORECASE) for p in [
            r"exclud", r"remove\s*row", r"not\s*applicable",
            r"should\s*not\s*be\s*includ"]):
        return "Should Be Excluded"

    # ── 13. Catch-all ─────────────────────────────────────────────────────────
    return "Other / Uncategorised"


ISSUE_COLORS = {
    "Wrong Page Number":              "#D6EAF8",
    "Wrong Table ID":                 "#D5F5E3",
    "Mapping Issue":                  "#FCF3CF",
    "Wrong Note ID":                  "#E8DAEF",
    "Missing Table":                  "#FDEBD0",
    "Duplicate / Extra Table":        "#FADBD8",
    "Wrong Data / Value":             "#D6DBDF",
    "Wrong Currency / Unit":          "#D0ECE7",
    "Wrong Period / Year":            "#FEF9E7",
    "Language / Translation Issue":   "#EBF5FB",
    "Wrong Report Number":            "#F9EBEA",
    "Should Be Excluded":             "#FDFEFE",
    "Other / Uncategorised":          "#F2F3F4",
}

HEADER_KEYWORDS = {"company_id", "qual ri", "qualri", "qual_ri", "table_id",
                   "key", "language", "period", "status", "qual changes"}


# ── Auto-detect real header row ────────────────────────────────────────────────

def find_header_row(filepath: str, sheet: str) -> int:
    print(f"\n🔍  Scanning for real header row in '{sheet}' …")
    raw = pd.read_excel(filepath, sheet_name=sheet, header=None, dtype=str, nrows=30)
    best_row, best_score = 0, 0
    for i, row in raw.iterrows():
        cells = {str(c).strip().lower() for c in row if pd.notna(c) and str(c).strip()}
        score = sum(1 for kw in HEADER_KEYWORDS if any(kw in cell for cell in cells))
        if score > best_score:
            best_score, best_row = score, i
        if score >= 4:
            break
    print(f"   Header detected at row index {best_row}  (score={best_score})")
    return int(best_row)


# ── Load sheet ─────────────────────────────────────────────────────────────────

def load_sheet(filepath: str, sheet: str) -> pd.DataFrame:
    header_row = find_header_row(filepath, sheet)
    df = pd.read_excel(filepath, sheet_name=sheet, header=header_row, dtype=str)
    df = df.dropna(axis=1, how="all")
    df.columns = (
        df.columns.astype(str).str.strip().str.lower()
        .str.replace(r"[\s./\-]+", "_", regex=True)
        .str.replace(r"[^a-z0-9_]", "", regex=True)
    )
    df = df.dropna(how="all").reset_index(drop=True)
    print(f"   Columns ({len(df.columns)}): {list(df.columns)}")
    print(f"   Rows after cleaning : {len(df)}")
    return df


# ── Column resolver ────────────────────────────────────────────────────────────

COLUMN_MAP = {
    "company":   ["company_id", "company"],
    "qual_ri":   ["qual_ri", "qualri", "qual_ri_", "ri", "reviewer"],
    "changes":   ["qual_changes", "qualchanges", "qual_change", "changes"],
    "table":     ["table_id", "tableid", "table"],
    "status":    ["status"],
    "period":    ["period"],
    "language":  ["language"],
    "page":      ["file_page_num", "page_num", "file_page", "page"],
    "note":      ["note_id"],
    "quant_ri":  ["quantri", "quant_ri"],
    "date":      ["date"],
    "no_report": ["no_of_report", "number_of_reports"],
}

def resolve(df, role):
    for c in COLUMN_MAP.get(role, []):
        if c in df.columns: return c
    return None


# ── Core analysis ──────────────────────────────────────────────────────────────

def analyse(df: pd.DataFrame) -> dict:
    col_company  = resolve(df, "company")
    col_qual_ri  = resolve(df, "qual_ri")
    col_changes  = resolve(df, "changes")
    col_table    = resolve(df, "table")
    col_status   = resolve(df, "status")
    col_period   = resolve(df, "period")
    col_language = resolve(df, "language")
    col_page     = resolve(df, "page")
    col_date     = resolve(df, "date")

    missing = [r for r, c in [("company_id", col_company), ("qual_ri", col_qual_ri)] if c is None]
    if missing:
        print(f"\n❌  Required columns not found: {missing}")
        print(f"    Available columns: {list(df.columns)}")
        sys.exit(1)

    print(f"\n✅  Column mapping:")
    for role, col in [("company", col_company), ("qual_ri", col_qual_ri),
                      ("changes", col_changes),  ("table",   col_table)]:
        print(f"    {role:12s} → '{col}'")

    df = df[df[col_qual_ri].notna() & (df[col_qual_ri].str.strip() != "")].copy()
    df[col_qual_ri] = df[col_qual_ri].str.strip().str.title()
    df[col_company] = df[col_company].str.strip()

    if col_changes:
        df["_has_change"]     = df[col_changes].notna() & (df[col_changes].str.strip() != "")
        df["_issue_category"] = df[col_changes].apply(
            lambda x: categorise_issue(x) if pd.notna(x) and str(x).strip() else "No Change"
        )
    else:
        df["_has_change"]     = False
        df["_issue_category"] = "No Change"
        print("⚠️  'Qual changes' column not found – change counts will be 0.")

    df["_table"] = df[col_table].str.strip() if col_table else "unknown"

    total_rows    = len(df)
    total_changes = int(df["_has_change"].sum())
    print(f"\n   Qual RI rows  : {total_rows}")
    print(f"   Qual changes  : {total_changes}")

    # A. Reviewer Summary
    ri_grp = df.groupby(col_qual_ri)
    ri_summary = pd.DataFrame({
        "Qual RI":       ri_grp[col_qual_ri].first(),
        "Total Tables":  ri_grp["_table"].count(),
        "Qual Changes":  ri_grp["_has_change"].sum().astype(int),
    }).reset_index(drop=True)
    ri_summary["Error Rate (%)"] = (ri_summary["Qual Changes"] / ri_summary["Total Tables"] * 100).round(2)
    ri_summary["Accuracy (%)"]   = (100 - ri_summary["Error Rate (%)"]).round(2)
    ri_summary = ri_summary.sort_values("Error Rate (%)", ascending=False).reset_index(drop=True)

    # B. Company Summary
    co_grp = df.groupby(col_company)
    co_summary = pd.DataFrame({
        "Company ID":   co_grp[col_company].first(),
        "Qual RIs":     co_grp[col_qual_ri].apply(lambda s: ", ".join(sorted(s.unique()))),
        "Total Tables": co_grp["_table"].count(),
        "Qual Changes": co_grp["_has_change"].sum().astype(int),
    }).reset_index(drop=True)
    co_summary["Error Rate (%)"] = (co_summary["Qual Changes"] / co_summary["Total Tables"] * 100).round(2)
    co_summary = co_summary.sort_values("Qual Changes", ascending=False).reset_index(drop=True)

    # C. Reviewer x Company
    ri_co = df.groupby([col_qual_ri, col_company]).agg(
        Total_Tables=("_table", "count"),
        Qual_Changes=("_has_change", "sum"),
    ).reset_index()
    ri_co.columns = ["Qual RI", "Company ID", "Total Tables", "Qual Changes"]
    ri_co["Qual Changes"]   = ri_co["Qual Changes"].astype(int)
    ri_co["Error Rate (%)"] = (ri_co["Qual Changes"] / ri_co["Total Tables"] * 100).round(2)
    ri_co = ri_co.sort_values(["Qual RI", "Error Rate (%)"], ascending=[True, False]).reset_index(drop=True)

    # D. Issue Category Breakdown
    issue_overall = (
        df[df["_has_change"]].groupby("_issue_category").size()
        .reset_index(name="Count")
        .rename(columns={"_issue_category": "Issue Category"})
        .sort_values("Count", ascending=False).reset_index(drop=True)
    )
    issue_overall["% of All Changes"] = (issue_overall["Count"] / total_changes * 100).round(2)

    issue_by_ri = (
        df[df["_has_change"]].groupby([col_qual_ri, "_issue_category"]).size()
        .reset_index(name="Count")
        .rename(columns={col_qual_ri: "Qual RI", "_issue_category": "Issue Category"})
        .sort_values(["Qual RI", "Count"], ascending=[True, False]).reset_index(drop=True)
    )

    issue_by_co = (
        df[df["_has_change"]].groupby([col_company, "_issue_category"]).size()
        .reset_index(name="Count")
        .rename(columns={col_company: "Company ID", "_issue_category": "Issue Category"})
        .sort_values(["Company ID", "Count"], ascending=[True, False]).reset_index(drop=True)
    )

    # E. Change Detail
    dmap = {col_company: "Company ID", col_qual_ri: "Qual RI", "_table": "Table ID"}
    if col_page:     dmap[col_page]     = "File Page"
    if col_changes:  dmap[col_changes]  = "Qual Change Description"
    dmap["_issue_category"] = "Issue Category"
    if col_status:   dmap[col_status]   = "Status"
    if col_date:     dmap[col_date]      = "Date"
    if col_period:   dmap[col_period]    = "Period"
    if col_language: dmap[col_language]  = "Language"

    change_detail = (
        df[df["_has_change"]][list(dmap.keys())]
        .rename(columns=dmap)
        .sort_values(["Qual RI", "Issue Category", "Company ID"])
        .reset_index(drop=True)
    )

    return dict(
        ri_summary    = ri_summary,
        co_summary    = co_summary,
        ri_co_detail  = ri_co,
        change_detail = change_detail,
        issue_overall = issue_overall,
        issue_by_ri   = issue_by_ri,
        issue_by_co   = issue_by_co,
        total_rows    = total_rows,
        total_changes = total_changes,
    )


# ── Export Excel ───────────────────────────────────────────────────────────────

def export_excel(results: dict, source_path: str) -> str:
    out_name = os.path.splitext(os.path.basename(source_path))[0] + "_qual_error_report.xlsx"
    out_path = os.path.join(os.path.dirname(source_path), out_name)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb  = writer.book
        hdr = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1e3a5f",
                              "border":1,"align":"center","valign":"vcenter"})
        ctr = wb.add_format({"align":"center"})
        red = wb.add_format({"bg_color":"#FFB3B3","align":"center","num_format":"0.00"})
        ora = wb.add_format({"bg_color":"#FFE0B2","align":"center","num_format":"0.00"})
        grn = wb.add_format({"bg_color":"#C8E6C9","align":"center","num_format":"0.00"})
        ttl = wb.add_format({"bold":True,"font_size":14,"font_color":"#1e3a5f"})

        def rfmt(v): return grn if v == 0 else (ora if v < 20 else red)

        def write_df(ws_name, df, rate_col="Error Rate (%)", issue_col=None):
            df.to_excel(writer, sheet_name=ws_name, index=False, startrow=1, header=False)
            ws = writer.sheets[ws_name]
            for ci, cn in enumerate(df.columns):
                ws.write(0, ci, cn, hdr)
                ws.set_column(ci, ci, max(len(str(cn)) + 4, 20))
            if rate_col in df.columns:
                rc = list(df.columns).index(rate_col)
                for ri_idx in range(len(df)):
                    ws.write(ri_idx + 1, rc, float(df[rate_col].iloc[ri_idx]),
                             rfmt(float(df[rate_col].iloc[ri_idx])))
            if issue_col and issue_col in df.columns:
                ic = list(df.columns).index(issue_col)
                for ri_idx in range(len(df)):
                    val = str(df[issue_col].iloc[ri_idx])
                    ws.write(ri_idx + 1, ic, val,
                             wb.add_format({"bg_color": ISSUE_COLORS.get(val, "#FFFFFF"),
                                            "align": "left"}))

        # Dashboard
        ws_d = wb.add_worksheet("Dashboard")
        ws_d.set_column(0, 0, 32); ws_d.set_column(1, 1, 22)
        ws_d.write(0, 0, "Qual RI Error Rate Tracker", ttl)
        ws_d.write(1, 0, f"Source: {os.path.basename(source_path)}")
        ws_d.write(2, 0, f"Sheet:  {SHEET_NAME}")

        overall = round(results["total_changes"] / max(results["total_rows"], 1) * 100, 2)
        stats = [
            ("Total Qual RI Rows",     results["total_rows"]),
            ("Total Qual Changes",     results["total_changes"]),
            ("Overall Error Rate (%)", overall),
            ("Unique Reviewers",       results["ri_summary"].shape[0]),
            ("Unique Companies",       results["co_summary"].shape[0]),
        ]
        ws_d.write(4, 0, "Metric", hdr); ws_d.write(4, 1, "Value", hdr)
        for i, (k, v) in enumerate(stats, 5):
            ws_d.write(i, 0, k); ws_d.write(i, 1, v, ctr)

        ws_d.write(11, 0, "Reviewer Rankings (by Error Rate %)", hdr)
        for j in range(1, 4): ws_d.write(11, j, "", hdr)
        for ci, cn in enumerate(["Qual RI", "Total Tables", "Qual Changes", "Error Rate (%)"]):
            ws_d.write(12, ci, cn, hdr)
        ri_df = results["ri_summary"]
        for i in range(len(ri_df)):
            ws_d.write(13+i, 0, ri_df["Qual RI"].iloc[i])
            ws_d.write(13+i, 1, int(ri_df["Total Tables"].iloc[i]), ctr)
            ws_d.write(13+i, 2, int(ri_df["Qual Changes"].iloc[i]), ctr)
            er = float(ri_df["Error Rate (%)"].iloc[i])
            ws_d.write(13+i, 3, er, rfmt(er))

        start = 13 + len(ri_df) + 2
        ws_d.write(start,   0, "Top Issue Categories (Overall)", hdr)
        for j in range(1, 3): ws_d.write(start, j, "", hdr)
        for ci, cn in enumerate(["Issue Category", "Count", "% of All Changes"]):
            ws_d.write(start+1, ci, cn, hdr)
        io_df = results["issue_overall"]
        for i in range(len(io_df)):
            cat = str(io_df["Issue Category"].iloc[i])
            ws_d.write(start+2+i, 0, cat,
                       wb.add_format({"bg_color": ISSUE_COLORS.get(cat, "#FFFFFF")}))
            ws_d.write(start+2+i, 1, int(io_df["Count"].iloc[i]), ctr)
            ws_d.write(start+2+i, 2, float(io_df["% of All Changes"].iloc[i]), ctr)

        write_df("Reviewer Summary",   results["ri_summary"])
        write_df("Company Summary",    results["co_summary"])
        write_df("Reviewer x Company", results["ri_co_detail"])
        write_df("Change Detail",      results["change_detail"],
                 rate_col="__none__",  issue_col="Issue Category")
        write_df("Issue Overview",     results["issue_overall"],  rate_col="__none__")
        write_df("Issues by Reviewer", results["issue_by_ri"],    rate_col="__none__")
        write_df("Issues by Company",  results["issue_by_co"],    rate_col="__none__")

        pivot_ri = results["issue_by_ri"].pivot_table(
            index="Qual RI", columns="Issue Category",
            values="Count", aggfunc="sum", fill_value=0
        ).reset_index()
        pivot_ri.to_excel(writer, sheet_name="Reviewer Issue Pivot",
                          index=False, startrow=1, header=False)
        ws_p = writer.sheets["Reviewer Issue Pivot"]
        for ci, cn in enumerate(pivot_ri.columns):
            ws_p.write(0, ci, str(cn), hdr)
            ws_p.set_column(ci, ci, max(len(str(cn)) + 4, 14))

        ws_l = wb.add_worksheet("Legend")
        ws_l.set_column(0, 0, 42)
        ws_l.write(0, 0, "Error Rate Colour Legend", hdr)
        ws_l.write(1, 0, "0%    – No errors",         grn)
        ws_l.write(2, 0, "< 20% – Low / acceptable",  ora)
        ws_l.write(3, 0, "≥ 20% – High error rate",   red)
        ws_l.write(5, 0, "Error Rate = Qual Changes ÷ Total Tables × 100")
        ws_l.write(7, 0, "Issue Category Legend", hdr)
        for i, (cat, color) in enumerate(ISSUE_COLORS.items(), 8):
            ws_l.write(i, 0, cat, wb.add_format({"bg_color": color, "border": 1}))

    print(f"\n✅  Report saved → {out_path}")
    return out_path


# ── Terminal print ─────────────────────────────────────────────────────────────

def print_results(r):
    overall = round(r["total_changes"] / max(r["total_rows"], 1) * 100, 2)
    sep = "═" * 65
    print(f"\n{sep}\n  QUAL RI ERROR RATE TRACKER\n{sep}")
    print(f"  Rows: {r['total_rows']}  |  Changes: {r['total_changes']}  |  Overall: {overall}%")
    print("\n── REVIEWER SUMMARY " + "─" * 45)
    print(r["ri_summary"].to_string(index=False))
    print("\n── ISSUE CATEGORY OVERVIEW " + "─" * 38)
    print(r["issue_overall"].to_string(index=False))
    print("\n── ISSUES BY REVIEWER " + "─" * 43)
    print(r["issue_by_ri"].to_string(index=False))
    print("\n── COMPANY SUMMARY (top 20) " + "─" * 37)
    print(r["co_summary"].head(20).to_string(index=False))


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    # Validate inputs
    if not os.path.isfile(FILE_PATH):
        print(f"❌  File not found: {FILE_PATH}")
        print("    Update FILE_PATH at the top of the script and try again.")
        sys.exit(1)

    xl     = pd.ExcelFile(FILE_PATH)
    sheets = xl.sheet_names
    print(f"\n📋  Sheets in workbook: {sheets}")

    if SHEET_NAME not in sheets:
        print(f"❌  Sheet '{SHEET_NAME}' not found.")
        print(f"    Available sheets: {sheets}")
        print("    Update SHEET_NAME at the top of the script and try again.")
        sys.exit(1)

    print(f"📂  File  : {FILE_PATH}")
    print(f"📄  Sheet : {SHEET_NAME}")

    df      = load_sheet(FILE_PATH, SHEET_NAME)
    results = analyse(df)
    print_results(results)
    export_excel(results, FILE_PATH)
    print("\nDone! Open _qual_error_report.xlsx to view your full dashboard.\n")


if __name__ == "__main__":
    main()
