"""Stage 2: Detect structure — find headers, normalise names, extract data rows."""
from __future__ import annotations

from table2db.models import SheetData, WorkbookData
from table2db.pipeline.island_detector import detect_table_islands


def detect_structure(
    wb: WorkbookData,
    header_min_fill_ratio: float = 0.5,
    header_min_string_ratio: float = 0.7,
) -> tuple[WorkbookData, list[str]]:
    """Detect headers and data regions in each sheet.

    Args:
        header_min_fill_ratio: Min ratio of non-empty cells in a header row (default 0.5).
        header_min_string_ratio: Min ratio of string cells in a header row (default 0.7).

    Returns the modified WorkbookData and a list of warning strings.
    """
    warnings: list[str] = []
    kept_sheets: list[SheetData] = []

    for sheet in wb.sheets:
        # Detect islands first
        islands = detect_table_islands(sheet.rows)

        if len(islands) == 0:
            sheet_warnings = _process_sheet(sheet, header_min_fill_ratio, header_min_string_ratio)
            warnings.extend(sheet_warnings)
            if sheet.headers and sheet.rows:
                sheet.metadata.setdefault("island_confidence", 1.0)
                kept_sheets.append(sheet)
            else:
                warnings.append(
                    f"Sheet '{sheet.name}' removed: no headers or no data rows"
                )
        elif len(islands) == 1:
            island = islands[0]
            sheet.rows = [
                row[island.col_start:island.col_end]
                for row in sheet.rows[island.row_start:island.row_end]
            ]
            sheet_warnings = _process_sheet(sheet, header_min_fill_ratio, header_min_string_ratio)
            warnings.extend(sheet_warnings)
            if sheet.headers and sheet.rows:
                sheet.metadata.setdefault("island_confidence", island.confidence)
                kept_sheets.append(sheet)
            else:
                warnings.append(
                    f"Sheet '{sheet.name}' removed: no headers or no data rows"
                )
        else:
            warnings.append(
                f"Sheet '{sheet.name}': detected {len(islands)} tables, splitting"
            )
            for i, island in enumerate(islands):
                sub_name = f"{sheet.name}_table_{i+1}"
                sub_sheet = SheetData(
                    name=sub_name,
                    rows=[row[island.col_start:island.col_end]
                          for row in sheet.rows[island.row_start:island.row_end]],
                    merge_map=sheet.merge_map,
                    metadata={**sheet.metadata, "island_confidence": island.confidence},
                    row_styles=sheet.row_styles,
                )
                sub_warnings = _process_sheet(sub_sheet, header_min_fill_ratio, header_min_string_ratio)
                warnings.extend(sub_warnings)
                if sub_sheet.headers and sub_sheet.rows:
                    kept_sheets.append(sub_sheet)

    wb.sheets = kept_sheets
    return wb, warnings


def _process_sheet(
    sheet: SheetData,
    header_min_fill_ratio: float = 0.5,
    header_min_string_ratio: float = 0.7,
) -> list[str]:
    """Process a single sheet: find header, extract data, normalise."""
    warnings: list[str] = []
    rows = sheet.rows

    if not rows:
        return warnings

    max_cols = max((len(r) for r in rows), default=0)
    if max_cols == 0:
        return warnings

    # Pad short rows
    for i, row in enumerate(rows):
        if len(row) < max_cols:
            rows[i] = row + [None] * (max_cols - len(row))

    # Find header row
    header_start = _find_header_row(rows, max_cols, sheet.merge_map,
                                     header_min_fill_ratio, header_min_string_ratio)
    if header_start is None:
        sheet.headers = []
        sheet.rows = []
        return warnings

    sheet.header_row_start = header_start

    # Compute header confidence
    header_row = rows[header_start]
    non_none = [v for v in header_row if v is not None]
    non_none_count = len(non_none)
    string_count = sum(1 for v in non_none if isinstance(v, str))
    if non_none_count > 0 and max_cols > 0:
        header_confidence = (non_none_count / max_cols) * (string_count / non_none_count)
    else:
        header_confidence = 0.0
    sheet.metadata["header_confidence"] = round(header_confidence, 2)

    # Multi-level header detection
    header_end = header_start
    next_row = header_start + 1
    if next_row < len(rows):
        if _is_string_row(rows[next_row], threshold=0.7) and not _is_data_row(
            rows, next_row, max_cols
        ):
            # Guard: if rows below the candidate also have no numerics,
            # then next_row is likely data, not a second header level
            check_range = list(range(next_row + 1, min(next_row + 4, len(rows))))
            all_string_below = check_range and all(
                _is_string_row(rows[r], threshold=0.5)
                and not _is_data_row(rows, r, max_cols)
                for r in check_range
            )
            if not all_string_below:
                header_end = next_row

    sheet.header_row_end = header_end

    # Extract headers
    if header_end > header_start:
        headers = _merge_multi_level_headers(
            rows[header_start], rows[header_end], max_cols, sheet.merge_map,
            header_start,
        )
    else:
        headers = [
            v if v is not None else None for v in rows[header_start]
        ]

    # Normalise column names
    headers = _normalize_headers(headers, max_cols)
    sheet.headers = headers

    # Extract data rows (after header_end)
    data_start = header_end + 1
    data_rows = rows[data_start:]

    # Trim trailing all-None rows
    while data_rows and all(v is None for v in data_rows[-1]):
        data_rows.pop()

    sheet.rows = data_rows
    return warnings


def _find_header_row(
    rows: list[list], max_cols: int, merge_map: dict[tuple, object],
    min_fill_ratio: float = 0.5, min_string_ratio: float = 0.7,
) -> int | None:
    """Find the first row that looks like a header."""
    for idx, row in enumerate(rows):
        non_none = [v for v in row if v is not None]
        non_none_count = len(non_none)

        # Enough cells filled?
        if max_cols > 0 and non_none_count / max_cols < min_fill_ratio:
            continue

        # Mostly strings?
        string_count = sum(1 for v in non_none if isinstance(v, str))
        if non_none_count > 0 and string_count / non_none_count < min_string_ratio:
            continue

        # Skip title rows: ≤2 cells filled in a wide table
        if non_none_count <= 2 and max_cols > 2:
            continue

        # Skip title rows: single merged cell spanning ≥80%
        if _is_title_merge(idx, max_cols, merge_map):
            continue

        # Check rows below have data (≥3 rows with at least 1 non-None)
        data_rows_below = 0
        for check_idx in range(idx + 1, min(idx + 6, len(rows))):
            if any(v is not None for v in rows[check_idx]):
                data_rows_below += 1
        if data_rows_below < 3:
            # Relax: if there's at least 1 data row below, still accept
            if data_rows_below < 1:
                continue

        return idx

    return None


def _is_title_merge(
    row_idx: int, max_cols: int, merge_map: dict[tuple, object]
) -> bool:
    """Check if a single merged group spans ≥80% of columns in this row."""
    # Group columns by their merge value to find individual merge ranges
    value_to_cols: dict[str, set[int]] = {}
    for (r, c), val in merge_map.items():
        if r == row_idx and val is not None:
            key = str(val)
            value_to_cols.setdefault(key, set()).add(c)

    # A title row has ONE merge group spanning ≥80% of columns
    for cols in value_to_cols.values():
        if max_cols > 0 and len(cols) >= 0.8 * max_cols:
            return True
    return False


def _is_string_row(row: list, threshold: float = 0.7) -> bool:
    """Check if ≥threshold of non-None cells are strings."""
    non_none = [v for v in row if v is not None]
    if not non_none:
        return False
    string_count = sum(1 for v in non_none if isinstance(v, str))
    return string_count / len(non_none) >= threshold


def _is_data_row(rows: list[list], idx: int, max_cols: int) -> bool:
    """Heuristic: a row looks like data if it has numeric values."""
    if idx >= len(rows):
        return False
    row = rows[idx]
    non_none = [v for v in row if v is not None]
    if not non_none:
        return False
    numeric_count = sum(1 for v in non_none if isinstance(v, (int, float)))
    return numeric_count / len(non_none) > 0.3


def _merge_multi_level_headers(
    top_row: list, bottom_row: list, max_cols: int,
    merge_map: dict[tuple, object], header_start: int,
) -> list[str]:
    """Merge two header rows into combined names like 'Parent_Child'."""
    # Build a mapping of which top-level header covers each column
    # by checking horizontal merges in merge_map at header_start row
    parent_for_col: dict[int, str] = {}

    # First pass: direct values from top row
    for c in range(max_cols):
        val = top_row[c] if c < len(top_row) else None
        if val is not None:
            parent_for_col[c] = str(val).strip()

    # For merged cells in the top row, the merge_map will have the same value
    # for all cells in the merge range. We use that to propagate parent names.
    # Group consecutive columns with the same merge_map value in the top row
    for c in range(max_cols):
        key = (header_start, c)
        if key in merge_map and merge_map[key] is not None:
            parent_for_col[c] = str(merge_map[key]).strip()

    headers = []
    for c in range(max_cols):
        parent = parent_for_col.get(c)
        child = bottom_row[c] if c < len(bottom_row) else None

        if child is not None:
            child_str = str(child).strip()
        else:
            child_str = None

        if parent and child_str:
            # If parent == child (e.g. "ID" spanning both rows), just use parent
            if parent == child_str:
                headers.append(parent)
            else:
                headers.append(f"{parent}_{child_str}")
        elif parent:
            headers.append(parent)
        elif child_str:
            headers.append(child_str)
        else:
            headers.append(None)

    return headers


def _normalize_headers(headers: list, max_cols: int) -> list[str]:
    """Normalize: strip whitespace, fill blanks, deduplicate."""
    result: list[str] = []
    for i, h in enumerate(headers):
        if h is None or (isinstance(h, str) and h.strip() == ""):
            result.append(f"column_{i + 1}")
        else:
            result.append(str(h).strip().replace("\n", " "))

    # Deduplicate
    seen: dict[str, int] = {}
    final: list[str] = []
    for name in result:
        if name in seen:
            count = seen[name]
            new_name = f"{name}_{count}"
            seen[name] = count + 1
            # Handle edge case where new_name also exists
            while new_name in seen:
                count += 1
                new_name = f"{name}_{count}"
                seen[name] = count + 1
            seen[new_name] = 1
            final.append(new_name)
        else:
            seen[name] = 1
            final.append(name)

    return final
