"""
Excel data loader — streaming architecture.

Instead of loading the entire Excel into memory, rows are streamed one at a
time via openpyxl's read_only iterator.  Merged-cell values are resolved
on-the-fly using a small metadata dict parsed from the XLSX XML (no full
data copy required).  Rows are flushed to SQLite in batches of BATCH_SIZE.

Memory footprint stays under ~100 MB even for 20 MB+ Excel files.
"""

import os
import re
import gc
import sys
import sqlite3
import datetime
import zipfile
import xml.etree.ElementTree as ET

import openpyxl

BATCH_SIZE = 2000
FTS_CHUNK_SIZE = 5000


def _log(msg):
    print(f"[DATA] {msg}", file=sys.stderr, flush=True)

_SLUG_RE = re.compile(r"[^a-zA-Z0-9\u4e00-\u9fff]+")
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_COL_RE = re.compile(r"([A-Z]+)(\d+)")


def _slugify(name: str) -> str:
    return _SLUG_RE.sub("_", name).strip("_")


def _cell_to_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M:%S")
    try:
        return str(val).strip()
    except Exception:
        return ""


def _safe_col_name(name, idx: int) -> str:
    if not name or not str(name).strip():
        return f"col_{idx}"
    return str(name).strip()


def _col_letter_to_idx(letters: str) -> int:
    result = 0
    for ch in letters:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


# ── Merge-range parsing (fast XML, no data loaded) ──────────────────────

def _parse_merge_ranges_from_xml(xlsx_path: str) -> dict[int, list[tuple]]:
    result: dict[int, list[tuple]] = {}
    try:
        zf = zipfile.ZipFile(xlsx_path, "r")
    except Exception:
        return result
    try:
        for name in zf.namelist():
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            idx_str = name.replace("xl/worksheets/sheet", "").replace(".xml", "")
            try:
                sheet_idx = int(idx_str)
            except ValueError:
                continue
            tree = ET.parse(zf.open(name))
            merge_elems = tree.findall(f".//{_NS}mergeCell")
            if not merge_elems:
                continue
            ranges = []
            for elem in merge_elems:
                ref = elem.get("ref", "")
                parts = ref.split(":")
                if len(parts) != 2:
                    continue
                m1 = _COL_RE.match(parts[0])
                m2 = _COL_RE.match(parts[1])
                if not m1 or not m2:
                    continue
                ranges.append((
                    int(m1.group(2)),                    # min_row (1-based)
                    int(m2.group(2)),                    # max_row
                    _col_letter_to_idx(m1.group(1)),     # min_col (0-based)
                    _col_letter_to_idx(m2.group(1)),     # max_col
                ))
            if ranges:
                result[sheet_idx] = ranges
    finally:
        zf.close()
    return result


def _build_merge_lookup(merges: list[tuple]):
    """Build two small dicts for O(1) merge resolution during streaming.

    Returns (source_cells, fill_cells):
      source_cells : {(row_1based, col_0based): None}  — top-left of each range
      fill_cells   : {(row_1based, col_0based): (src_row, src_col)} — cells to fill
    """
    source_cells = {}
    fill_cells = {}
    for min_row, max_row, min_col, max_col in merges:
        source_cells[(min_row, min_col)] = None
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if r == min_row and c == min_col:
                    continue
                fill_cells[(r, c)] = (min_row, min_col)
    return source_cells, fill_cells


# ── Streaming loader ────────────────────────────────────────────────────

def load_excel_to_db(
    model_id: str,
    excel_path: str,
    db_path: str,
    mode: str = "replace",
    on_progress=None,
):
    """Load Excel into SQLite using streaming row iteration.

    *on_progress* is an optional callback:
        on_progress(sheet_index, total_sheets, sheet_name, rows_processed)
    Called every BATCH_SIZE rows and at end of each sheet.
    """
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    merge_map = _parse_merge_ranges_from_xml(excel_path)

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    sheet_names = list(wb.sheetnames)
    total_sheets = len(sheet_names)
    wb.close()
    del wb
    gc.collect()

    _log(f"Processing {excel_path}: {total_sheets} sheets")

    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")

    failed_sheets: list[str] = []

    try:
        if mode == "replace":
            _drop_model_tables(conn, model_id)

        for sheet_idx_0, sheet_name in enumerate(sheet_names):
            sheet_idx = sheet_idx_0 + 1

            if on_progress:
                on_progress(sheet_idx, total_sheets, sheet_name, 0)

            try:
                _process_sheet(
                    conn, model_id, excel_path, sheet_name,
                    sheet_idx, total_sheets, merge_map, on_progress,
                )
            except Exception as exc:
                _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — FAILED: {exc}")
                failed_sheets.append(sheet_name)
                gc.collect()

        _save_model_meta(conn, model_id, excel_path, sheet_names)
    finally:
        conn.close()

    if failed_sheets:
        _log(f"Finished with {len(failed_sheets)} failed sheet(s): {failed_sheets}")


def _find_header_row(ws, max_scan: int = 10) -> tuple[int, tuple]:
    """Detect the real header row by scanning the first few rows.

    Heuristic: if a row has <=2 non-empty cells, it's likely a title/meta row.
    The header is the first row with >=3 non-empty cells, or the row with the
    most non-empty string-like cells among the first ``max_scan`` rows.
    """
    candidates: list[tuple[int, tuple, int]] = []
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)):
        row_tuple = tuple(row)
        non_empty = sum(1 for v in row_tuple if v is not None and str(v).strip())
        candidates.append((i + 1, row_tuple, non_empty))

    if not candidates:
        return 1, ()

    for row_num, row_tuple, count in candidates:
        if count >= 3:
            return row_num, row_tuple

    best = max(candidates, key=lambda x: x[2])
    return best[0], best[1]


_PLACEHOLDER_RE = re.compile(
    r"^(col_\d+|第\s*.+\s*列.*|column\s*\d+|\.\.*|_{2,})$", re.IGNORECASE
)


def _is_placeholder_header(name) -> bool:
    """Check if a header name is an auto-generated placeholder."""
    if name is None:
        return True
    s = str(name).strip()
    if not s:
        return True
    return bool(_PLACEHOLDER_RE.match(s))


def _identify_empty_columns(header_row: tuple, sample_rows: list[tuple]) -> set[int]:
    """Return column indices to skip: placeholder/empty header with no sample data."""
    empty = set()
    for i, h in enumerate(header_row):
        if not _is_placeholder_header(h):
            continue
        has_data = False
        for row in sample_rows:
            if i < len(row) and row[i] is not None and str(row[i]).strip():
                has_data = True
                break
        if not has_data:
            empty.add(i)
    return empty


def _process_sheet(
    conn: sqlite3.Connection,
    model_id: str,
    excel_path: str,
    sheet_name: str,
    sheet_idx: int,
    total_sheets: int,
    merge_map: dict,
    on_progress,
):
    _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — opening workbook")

    wb = openpyxl.load_workbook(
        excel_path, read_only=True, data_only=True,
    )
    try:
        ws = wb[sheet_name]

        header_row_num, header_raw = _find_header_row(ws)
        if not header_raw or all(v is None for v in header_raw):
            if on_progress:
                on_progress(sheet_idx, total_sheets, sheet_name, 0)
            return

        if header_row_num > 1:
            _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — "
                 f"auto-detected header at row {header_row_num}")

        # Trim trailing None columns from header
        header_list = list(header_raw)
        while header_list and header_list[-1] is None:
            header_list.pop()
        if not header_list:
            if on_progress:
                on_progress(sheet_idx, total_sheets, sheet_name, 0)
            return

        # Sample a few data rows to detect empty columns
        sample_rows: list[tuple] = []
        data_start = header_row_num + 1
        for row in ws.iter_rows(min_row=data_start, max_row=data_start + 99,
                                values_only=True):
            sample_rows.append(tuple(row))

        empty_cols = _identify_empty_columns(tuple(header_list), sample_rows)

        # Build final column list, excluding empty columns
        keep_indices: list[int] = []
        raw_columns: list[str] = []
        for i, h in enumerate(header_list):
            if i in empty_cols:
                continue
            keep_indices.append(i)
            raw_columns.append(_safe_col_name(h, i))

        if not keep_indices:
            if on_progress:
                on_progress(sheet_idx, total_sheets, sheet_name, 0)
            return

        # Deduplicate column names
        seen: dict[str, int] = {}
        columns: list[str] = []
        for c in raw_columns:
            if c in seen:
                seen[c] += 1
                columns.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                columns.append(c)

        num_cols = len(columns)
        full_width = len(header_list)

        if empty_cols:
            _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — "
                 f"skipped {len(empty_cols)} empty column(s), keeping {num_cols}")

        table = f"{model_id}__{_slugify(sheet_name)}"
        fts_table = f"{table}__fts"

        col_defs = ", ".join(f'"{c}" TEXT' for c in columns)
        conn.execute(
            f'CREATE TABLE IF NOT EXISTS "{table}" '
            f"(_rowid INTEGER PRIMARY KEY AUTOINCREMENT, {col_defs})"
        )

        placeholders = ", ".join(["?"] * num_cols)
        col_names = ", ".join(f'"{c}"' for c in columns)
        insert_sql = f'INSERT INTO "{table}" ({col_names}) VALUES ({placeholders})'

        merges = merge_map.get(sheet_idx, [])
        source_cells, fill_cells = _build_merge_lookup(merges) if merges else ({}, {})

        batch: list[tuple] = []
        rows_processed = 0

        for row in ws.iter_rows(min_row=data_start, values_only=True):
            excel_row = header_row_num + rows_processed + 1
            row_list = list(row)

            while len(row_list) < full_width:
                row_list.append(None)

            if source_cells or fill_cells:
                for col_idx in range(full_width):
                    cell_key = (excel_row, col_idx)
                    if cell_key in source_cells:
                        source_cells[cell_key] = row_list[col_idx]
                    elif row_list[col_idx] is None and cell_key in fill_cells:
                        src_key = fill_cells[cell_key]
                        src_val = source_cells.get(src_key)
                        if src_val is not None:
                            row_list[col_idx] = src_val

            vals = tuple(_cell_to_str(row_list[i]) for i in keep_indices)
            if all(v == "" for v in vals):
                rows_processed += 1
                continue

            batch.append(vals)
            rows_processed += 1

            if len(batch) >= BATCH_SIZE:
                conn.executemany(insert_sql, batch)
                conn.commit()
                batch.clear()
                if on_progress:
                    on_progress(sheet_idx, total_sheets, sheet_name, rows_processed)

        if batch:
            conn.executemany(insert_sql, batch)
            conn.commit()
            batch.clear()

        source_cells.clear()
        fill_cells.clear()
    finally:
        wb.close()
        del wb
        gc.collect()

    _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — {rows_processed} rows, building FTS")

    conn.execute(f'DROP TABLE IF EXISTS "{fts_table}"')
    conn.execute(
        f'CREATE VIRTUAL TABLE "{fts_table}" USING fts5('
        f"_all_text, content=\"{table}\", content_rowid=_rowid, "
        f"tokenize=\"unicode61 remove_diacritics 2\")"
    )

    concat = _concat_expr(columns)
    last_rowid = 0
    while True:
        rows = conn.execute(
            f'SELECT _rowid, {concat} FROM "{table}" '
            f"WHERE _rowid > ? ORDER BY _rowid LIMIT ?",
            (last_rowid, FTS_CHUNK_SIZE),
        ).fetchall()
        if not rows:
            break
        conn.executemany(
            f'INSERT INTO "{fts_table}" (rowid, _all_text) VALUES (?, ?)',
            rows,
        )
        conn.commit()
        last_rowid = rows[-1][0]
        del rows
    gc.collect()

    if on_progress:
        on_progress(sheet_idx, total_sheets, sheet_name, rows_processed)

    _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — done, memory released")


def _concat_expr(columns: list[str]) -> str:
    return " || ' ' || ".join(f'COALESCE("{c}", \'\')' for c in columns)


def _drop_model_tables(conn: sqlite3.Connection, model_id: str):
    cur = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE ?",
        (f"{model_id}__%",),
    )
    for (tbl,) in cur.fetchall():
        conn.execute(f'DROP TABLE IF EXISTS "{tbl}"')
    try:
        conn.execute("DELETE FROM model_meta WHERE model_id = ?", (model_id,))
    except sqlite3.OperationalError:
        pass


def _save_model_meta(conn, model_id, excel_path, sheet_names):
    conn.execute(
        "CREATE TABLE IF NOT EXISTS model_meta ("
        "  model_id TEXT, sheet_name TEXT, table_name TEXT, excel_path TEXT,"
        "  PRIMARY KEY (model_id, sheet_name))"
    )
    for sn in sheet_names:
        tbl = f"{model_id}__{_slugify(sn)}"
        cur = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (tbl,)
        )
        if not cur.fetchone():
            continue
        conn.execute(
            "INSERT OR REPLACE INTO model_meta VALUES (?, ?, ?, ?)",
            (model_id, sn, tbl, excel_path),
        )
    conn.commit()


def get_model_sheets(model_id: str, db_path: str) -> list[dict]:
    if not os.path.isfile(db_path):
        return []
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS model_meta ("
            "  model_id TEXT, sheet_name TEXT, table_name TEXT, excel_path TEXT,"
            "  PRIMARY KEY (model_id, sheet_name))"
        )
        cur = conn.execute(
            "SELECT sheet_name, table_name FROM model_meta WHERE model_id = ? ORDER BY rowid",
            (model_id,),
        )
        results = []
        for sheet_name, table_name in cur.fetchall():
            col_cur = conn.execute(f'PRAGMA table_info("{table_name}")')
            cols = [r[1] for r in col_cur.fetchall() if r[1] != "_rowid"]
            cnt = conn.execute(f'SELECT COUNT(*) FROM "{table_name}"').fetchone()[0]
            results.append({
                "sheet_name": sheet_name,
                "table_name": table_name,
                "columns": cols,
                "row_count": cnt,
            })
        return results
    finally:
        conn.close()


def get_all_loaded_models(db_path: str) -> list[str]:
    if not os.path.isfile(db_path):
        return []
    conn = sqlite3.connect(db_path)
    try:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS model_meta ("
            "  model_id TEXT, sheet_name TEXT, table_name TEXT, excel_path TEXT,"
            "  PRIMARY KEY (model_id, sheet_name))"
        )
        cur = conn.execute("SELECT DISTINCT model_id FROM model_meta")
        return [r[0] for r in cur.fetchall()]
    finally:
        conn.close()


# ── Custom models (user-defined, stored in per-user SQLite) ──────────

_CUSTOM_MODELS_DDL = (
    "CREATE TABLE IF NOT EXISTS custom_models ("
    "  model_id TEXT PRIMARY KEY, name TEXT NOT NULL, created_at TEXT)"
)


def _ensure_custom_models_table(conn: sqlite3.Connection):
    conn.execute(_CUSTOM_MODELS_DDL)


def create_custom_model(db_path: str, model_id: str, name: str) -> dict:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        _ensure_custom_models_table(conn)
        now = datetime.datetime.now(datetime.timezone.utc).isoformat()
        conn.execute(
            "INSERT OR REPLACE INTO custom_models (model_id, name, created_at) VALUES (?, ?, ?)",
            (model_id, name, now),
        )
        conn.commit()
        return {"id": model_id, "name": name, "created_at": now}
    finally:
        conn.close()


def get_custom_models(db_path: str) -> list[dict]:
    if not os.path.isfile(db_path):
        return []
    conn = sqlite3.connect(db_path)
    try:
        _ensure_custom_models_table(conn)
        cur = conn.execute("SELECT model_id, name, created_at FROM custom_models ORDER BY created_at")
        return [{"id": r[0], "name": r[1], "created_at": r[2]} for r in cur.fetchall()]
    finally:
        conn.close()


def delete_custom_model(db_path: str, model_id: str) -> bool:
    if not os.path.isfile(db_path):
        return False
    conn = sqlite3.connect(db_path)
    try:
        _ensure_custom_models_table(conn)
        _drop_model_tables(conn, model_id)
        conn.execute("DELETE FROM custom_models WHERE model_id = ?", (model_id,))
        conn.commit()
        return True
    finally:
        conn.close()
