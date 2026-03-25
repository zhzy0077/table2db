"""New evaluation scorer: tiered scoring (data 0-0.6, columns 0.6-0.8, tables 0.8-1.0)."""

import os
import re
import sqlite3
import json
import sys
from collections import defaultdict
from difflib import SequenceMatcher

EVAL_DIR = os.path.dirname(__file__)
GOLDEN_DIRS = [
    os.path.join(EVAL_DIR, "golden_dbs", "references"),
    os.path.join(EVAL_DIR, "golden_dbs", "deliverables"),
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _safe_connect(path: str) -> sqlite3.Connection | None:
    try:
        conn = sqlite3.connect(path)
        conn.execute("SELECT 1")
        return conn
    except Exception:
        return None


def _get_tables(conn: sqlite3.Connection) -> list[str]:
    cur = conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
    return [r[0] for r in cur.fetchall() if not r[0].startswith("_")]


def _get_meta(conn: sqlite3.Connection, meta_table: str) -> list[tuple[str, str]]:
    """Read _meta table as list of (source, target) pairs."""
    try:
        cur = conn.execute(f"SELECT source, target FROM [{meta_table}]")
        return cur.fetchall()
    except Exception:
        return []


def _table_source_sheet(meta: list[tuple[str, str]], table_name: str) -> str | None:
    """Derive source sheet for a table from its meta entries.

    Each entry has source like 'Sheet1!B2:B100' and target like 'table/col'.
    """
    for source, target in meta:
        tbl = target.split("/", 1)[0] if "/" in target else target
        if tbl == table_name and "!" in source:
            return source.split("!", 1)[0]
    return None


def _get_columns(conn: sqlite3.Connection, table: str) -> list[tuple[str, str]]:
    try:
        cur = conn.execute(f"PRAGMA table_info([{table}])")
        return [(r[1], r[2]) for r in cur.fetchall()]
    except Exception:
        return []


def _get_row_count(conn: sqlite3.Connection, table: str) -> int:
    try:
        return conn.execute(f"SELECT COUNT(*) FROM [{table}]").fetchone()[0]
    except Exception:
        return 0


def _get_all_rows(conn: sqlite3.Connection, table: str) -> list[tuple]:
    try:
        return conn.execute(f"SELECT * FROM [{table}]").fetchall()
    except Exception:
        return []


def _normalize_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def _normalize_source_sheet(sheet: str) -> str:
    return re.sub(r"_table_\d+$", "", sheet).strip()


# ---------------------------------------------------------------------------
# Table matching via _meta source_sheet (primary) + fallbacks
# ---------------------------------------------------------------------------

def match_tables(
    golden_conn: sqlite3.Connection,
    test_conn: sqlite3.Connection,
    golden_meta: list[tuple[str, str]],
    test_meta: list[tuple[str, str]],
) -> list[tuple[str, str | None]]:
    """Match golden tables to test tables. Returns [(golden_table, test_table_or_None)]."""
    golden_tables = _get_tables(golden_conn)
    test_tables = _get_tables(test_conn)

    if not golden_tables:
        return []

    matched: list[tuple[str, str | None]] = []
    used_test: set[str] = set()

    # Pass 1: match by source_sheet from _meta
    test_sheet_map: dict[str, list[str]] = defaultdict(list)
    for tt in test_tables:
        sheet = _table_source_sheet(test_meta, tt)
        if sheet:
            norm = _normalize_source_sheet(sheet).lower()
            test_sheet_map[norm].append(tt)

    for gt in golden_tables:
        g_sheet = _table_source_sheet(golden_meta, gt)
        if g_sheet:
            candidates = test_sheet_map.get(g_sheet.lower(), [])
            best, best_score = None, -1.0
            for c in candidates:
                if c in used_test:
                    continue
                g_cols = {_normalize_name(n) for n, _ in _get_columns(golden_conn, gt)}
                t_cols = {_normalize_name(n) for n, _ in _get_columns(test_conn, c)}
                score = len(g_cols & t_cols) / max(len(g_cols), 1)
                if score > best_score:
                    best_score = score
                    best = c
            if best is not None:
                matched.append((gt, best))
                used_test.add(best)
                continue
        matched.append((gt, None))

    # Pass 2: fuzzy name matching for unmatched
    for i, (gt, tt) in enumerate(matched):
        if tt is not None:
            continue
        best, best_ratio = None, 0.0
        gt_norm = _normalize_name(gt)
        for cand in test_tables:
            if cand in used_test:
                continue
            ratio = SequenceMatcher(None, gt_norm, _normalize_name(cand)).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best = cand
        if best_ratio > 0.6 and best is not None:
            matched[i] = (gt, best)
            used_test.add(best)

    # Pass 3: column-overlap matching for remaining unmatched
    for i, (gt, tt) in enumerate(matched):
        if tt is not None:
            continue
        g_cols = {_normalize_name(n) for n, _ in _get_columns(golden_conn, gt)}
        if not g_cols:
            continue
        best, best_score = None, 0.0
        for cand in test_tables:
            if cand in used_test:
                continue
            t_cols = {_normalize_name(n) for n, _ in _get_columns(test_conn, cand)}
            score = len(g_cols & t_cols) / len(g_cols)
            if score > best_score:
                best_score = score
                best = cand
        if best_score >= 0.3 and best is not None:
            matched[i] = (gt, best)
            used_test.add(best)

    # Pass 4: if only 1 golden and 1 test table remain unmatched, force-match
    unmatched_golden = [(i, gt) for i, (gt, tt) in enumerate(matched) if tt is None]
    unmatched_test = [t for t in test_tables if t not in used_test]
    if len(unmatched_golden) == 1 and len(unmatched_test) == 1:
        idx = unmatched_golden[0][0]
        matched[idx] = (matched[idx][0], unmatched_test[0])

    return matched


# ---------------------------------------------------------------------------
# Column matching via _meta source ranges (strict, no data-overlap fallback)
# ---------------------------------------------------------------------------

def _get_column_source_map(
    meta: list[tuple[str, str]], table_name: str,
) -> dict[str, str]:
    """Return {source_range: col_name} for a table from _meta entries.

    Each _meta row: source='Sheet1!B2:B100', target='table_name/col_name'.
    """
    result: dict[str, str] = {}
    for source, target in meta:
        if "/" not in target:
            continue
        tbl, col = target.split("/", 1)
        if tbl == table_name:
            result[source] = col
    return result


def _match_columns_via_meta(
    golden_meta: list[tuple[str, str]],
    test_meta: list[tuple[str, str]],
    golden_table: str,
    test_table: str,
) -> list[tuple[str, str]]:
    """Match golden→test columns by exact source range. Returns [(golden_col, test_col)]."""
    golden_map = _get_column_source_map(golden_meta, golden_table)
    test_map = _get_column_source_map(test_meta, test_table)
    matched = []
    for source_range, golden_col in golden_map.items():
        if source_range in test_map:
            matched.append((golden_col, test_map[source_range]))
    return matched


# ---------------------------------------------------------------------------
# Data matching using _meta-aligned columns
# ---------------------------------------------------------------------------

def _values_match(golden_val, test_val) -> bool:
    if golden_val is None and test_val is None:
        return True
    if golden_val is None or test_val is None:
        other = test_val if golden_val is None else golden_val
        if isinstance(other, str) and other.strip() == "":
            return True
        return False

    # Both numeric
    try:
        g_num = float(golden_val)
        t_num = float(test_val)
        if g_num == 0 and t_num == 0:
            return True
        if g_num == 0:
            return abs(t_num) < 0.01
        return abs(g_num - t_num) / max(abs(g_num), 1e-10) <= 0.01
    except (TypeError, ValueError):
        pass

    g_str = str(golden_val).strip().lower()
    t_str = str(test_val).strip().lower()
    return g_str == t_str


def _table_data_is_correct(
    golden_conn: sqlite3.Connection,
    test_conn: sqlite3.Connection,
    golden_table: str,
    test_table: str,
    col_pairs: list[tuple[str, str]],
) -> bool:
    """Binary check: does test table contain all golden data?

    col_pairs: [(golden_col_name, test_col_name)] from _meta matching.
    """
    g_rows = _get_all_rows(golden_conn, golden_table)
    t_rows = _get_all_rows(test_conn, test_table)

    if not g_rows:
        return len(t_rows) == 0
    if len(g_rows) != len(t_rows):
        return False
    if not col_pairs:
        return False

    # Map column names → indices
    g_col_names = [n for n, _ in _get_columns(golden_conn, golden_table)]
    t_col_names = [n for n, _ in _get_columns(test_conn, test_table)]
    g_idx = {name: i for i, name in enumerate(g_col_names)}
    t_idx = {name: i for i, name in enumerate(t_col_names)}

    idx_pairs = []
    for g_col, t_col in col_pairs:
        gi = g_idx.get(g_col)
        ti = t_idx.get(t_col)
        if gi is not None and ti is not None:
            idx_pairs.append((gi, ti))

    if not idx_pairs:
        return False

    # Greedy row matching: every golden row must find a perfect match
    used_t: set[int] = set()
    for g_row in g_rows:
        found = False
        for ti, t_row in enumerate(t_rows):
            if ti in used_t:
                continue
            all_match = all(
                _values_match(
                    g_row[gi] if gi < len(g_row) else None,
                    t_row[tci] if tci < len(t_row) else None,
                )
                for gi, tci in idx_pairs
            )
            if all_match:
                used_t.add(ti)
                found = True
                break
        if not found:
            return False

    return True


# ---------------------------------------------------------------------------
# Scoring tiers
# ---------------------------------------------------------------------------

def score_file(golden_db_path: str, test_db_path: str | None) -> dict:
    basename = os.path.basename(golden_db_path)
    result: dict = {
        "file": basename,
        "converted": test_db_path is not None,
        "data_score": 0.0,
        "col_score": 0.0,
        "table_score": 0.0,
        "final_score": 0.0,
        "details": {},
    }

    golden_conn = _safe_connect(golden_db_path)
    if golden_conn is None:
        result["error"] = "Could not open golden DB"
        return result

    golden_tables = _get_tables(golden_conn)
    result["details"]["golden_tables"] = len(golden_tables)

    if not test_db_path or not os.path.exists(test_db_path):
        result["details"]["test_tables"] = 0
        golden_conn.close()
        return result

    test_conn = _safe_connect(test_db_path)
    if test_conn is None:
        result["error"] = "Could not open test DB"
        golden_conn.close()
        return result

    try:
        golden_meta = _get_meta(golden_conn, "_meta")
        test_meta = _get_meta(test_conn, "_meta")
        matches = match_tables(golden_conn, test_conn, golden_meta, test_meta)

        test_table_count = len(_get_tables(test_conn))
        matched_count = sum(1 for _, tt in matches if tt is not None)

        result["details"]["test_tables"] = test_table_count
        result["details"]["matched_tables"] = matched_count
        result["details"]["matches"] = []

        n_golden = len(golden_tables) or 1
        total_golden_meta_cols = 0
        correct_col_names = 0
        correct_data_tables = 0
        correct_table_names = 0

        for gt, tt in matches:
            match_detail = {"golden": gt, "test": tt, "data_correct": False,
                            "col_names_correct": [], "table_name_correct": False}

            # Count golden columns from _meta for this table
            golden_col_map = _get_column_source_map(golden_meta, gt)
            n_golden_cols = len(golden_col_map)
            total_golden_meta_cols += n_golden_cols

            if tt is None:
                result["details"]["matches"].append(match_detail)
                continue

            # Match columns via _meta source ranges
            col_pairs = _match_columns_via_meta(golden_meta, test_meta, gt, tt)
            match_detail["meta_matched_cols"] = len(col_pairs)
            match_detail["golden_meta_cols"] = n_golden_cols

            # --- Tier 1: Data correctness (binary per table) ---
            # All golden meta columns must be matched for data to be correct
            if len(col_pairs) == n_golden_cols and n_golden_cols > 0:
                data_ok = _table_data_is_correct(
                    golden_conn, test_conn, gt, tt, col_pairs,
                )
            else:
                data_ok = False
            match_detail["data_correct"] = data_ok
            if data_ok:
                correct_data_tables += 1

            # --- Tier 2: Column name correctness (among meta-matched pairs) ---
            for g_col, t_col in col_pairs:
                if _normalize_name(g_col) == _normalize_name(t_col):
                    correct_col_names += 1
                    match_detail["col_names_correct"].append(g_col)

            # --- Tier 3: Table name correctness ---
            if _normalize_name(gt) == _normalize_name(tt):
                correct_table_names += 1
                match_detail["table_name_correct"] = True

            result["details"]["matches"].append(match_detail)

        result["data_score"] = round((correct_data_tables / n_golden) * 0.6, 4)
        result["col_score"] = round(
            (correct_col_names / max(total_golden_meta_cols, 1)) * 0.2, 4,
        )
        result["table_score"] = round((correct_table_names / n_golden) * 0.2, 4)
        result["final_score"] = round(
            result["data_score"] + result["col_score"] + result["table_score"], 4
        )

        result["details"]["correct_data_tables"] = correct_data_tables
        result["details"]["correct_cols"] = correct_col_names
        result["details"]["total_golden_cols"] = total_golden_meta_cols
        result["details"]["correct_table_names"] = correct_table_names

    except Exception as e:
        result["error"] = str(e)
        import traceback
        traceback.print_exc()
    finally:
        golden_conn.close()
        test_conn.close()

    return result


# ---------------------------------------------------------------------------
# Full evaluation
# ---------------------------------------------------------------------------

def _find_test_output(golden_basename: str, test_dirs: list[str]) -> str | None:
    stem = os.path.splitext(golden_basename)[0]
    candidates = [
        stem.replace(" ", "_") + ".db",
        golden_basename,
    ]
    for sdir in test_dirs:
        for name in candidates:
            path = os.path.join(sdir, name)
            if os.path.exists(path):
                return path
    return None


def run_evaluation(test_output_dir: str) -> dict:
    test_dirs = []
    for subdir in ["references", "deliverables"]:
        sub = os.path.join(test_output_dir, subdir)
        if os.path.isdir(sub):
            test_dirs.append(sub)
    if not test_dirs:
        test_dirs = [test_output_dir]

    golden_files = []
    for gdir in GOLDEN_DIRS:
        if not os.path.isdir(gdir):
            continue
        category = "references" if "references" in gdir else "deliverables"
        for f in os.listdir(gdir):
            if f.endswith(".db"):
                golden_files.append((f, os.path.join(gdir, f), category))
    golden_files.sort(key=lambda x: x[0])

    results: dict = {
        "total_files": len(golden_files),
        "converted": 0,
        "failed": 0,
        "file_results": [],
    }

    for gf, golden_path, category in golden_files:
        test_path = _find_test_output(gf, test_dirs)
        file_result = score_file(golden_path, test_path)
        file_result["category"] = category
        results["file_results"].append(file_result)

        if file_result["converted"]:
            results["converted"] += 1
        else:
            results["failed"] += 1

    # Overall averages
    n = max(results["total_files"], 1)
    results["overall"] = {
        "data_score": round(sum(fr["data_score"] for fr in results["file_results"]) / n, 4),
        "col_score": round(sum(fr["col_score"] for fr in results["file_results"]) / n, 4),
        "table_score": round(sum(fr["table_score"] for fr in results["file_results"]) / n, 4),
        "final_score": round(sum(fr["final_score"] for fr in results["file_results"]) / n, 4),
    }
    return results


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------

def print_report(results: dict) -> None:
    total = results["total_files"]
    converted = results["converted"]
    failed = results["failed"]
    ov = results["overall"]

    print("=== EVALUATION REPORT (new scorer) ===")
    print(f"Files: {total} total, {converted} converted, {failed} not converted")
    print()
    print("Overall Scores (averaged across all files):")
    print(f"  Data correctness (0-0.6):  {ov['data_score']:.4f}")
    print(f"  Column names (0-0.2):      {ov['col_score']:.4f}")
    print(f"  Table names (0-0.2):       {ov['table_score']:.4f}")
    print(f"  FINAL SCORE:               {ov['final_score']:.4f}")
    print()

    file_results = results["file_results"]
    converted_results = [fr for fr in file_results if fr["converted"]]
    if converted_results:
        n_conv = len(converted_results)
        print(f"Scores for converted files only ({n_conv} files):")
        avg_data = sum(fr["data_score"] for fr in converted_results) / n_conv
        avg_col = sum(fr["col_score"] for fr in converted_results) / n_conv
        avg_tbl = sum(fr["table_score"] for fr in converted_results) / n_conv
        avg_final = sum(fr["final_score"] for fr in converted_results) / n_conv
        print(f"  Data correctness:  {avg_data:.4f}")
        print(f"  Column names:      {avg_col:.4f}")
        print(f"  Table names:       {avg_tbl:.4f}")
        print(f"  FINAL SCORE:       {avg_final:.4f}")
        print()

    # Sort by final_score
    sorted_results = sorted(file_results, key=lambda x: x["final_score"])

    # Per-file details for converted files
    print("Per-file scores (converted only):")
    for fr in sorted(converted_results or [], key=lambda x: x["final_score"]):
        d = fr.get("details", {})
        data_tables = d.get("correct_data_tables", 0)
        gt = d.get("golden_tables", 0)
        cc = d.get("correct_cols", 0)
        tc = d.get("total_golden_cols", 0)
        tn = d.get("correct_table_names", 0)
        print(
            f"  {fr['final_score']:.2f}  {fr['file']:<55s} "
            f"data={data_tables}/{gt}  cols={cc}/{tc}  tblnames={tn}/{gt}"
        )
    print()

    # Grade distribution
    grades = {"A (>=0.8)": 0, "B (0.6-0.8)": 0, "C (0.4-0.6)": 0,
              "D (0.2-0.4)": 0, "F (<0.2)": 0}
    for fr in file_results:
        w = fr["final_score"]
        if w >= 0.8:
            grades["A (>=0.8)"] += 1
        elif w >= 0.6:
            grades["B (0.6-0.8)"] += 1
        elif w >= 0.4:
            grades["C (0.4-0.6)"] += 1
        elif w >= 0.2:
            grades["D (0.2-0.4)"] += 1
        else:
            grades["F (<0.2)"] += 1

    print("Score Distribution:")
    for grade, count in grades.items():
        print(f"  {grade}: {count} files")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <test_output_dir>")
        sys.exit(1)

    test_dir = sys.argv[1]
    results = run_evaluation(test_dir)
    print_report(results)

    json_path = os.path.join(EVAL_DIR, "new_eval_results.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nDetailed results: {json_path}")
