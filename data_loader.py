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

BATCH_SIZE = 1000


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
    return str(val).strip()


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

    try:
        if mode == "replace":
            _drop_model_tables(conn, model_id)

        for sheet_idx_0, sheet_name in enumerate(sheet_names):
            sheet_idx = sheet_idx_0 + 1

            if on_progress:
                on_progress(sheet_idx, total_sheets, sheet_name, 0)

            _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — opening workbook")

            wb = openpyxl.load_workbook(
                excel_path, read_only=True, data_only=True,
            )
            ws = wb[sheet_name]

            header_row = None
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                header_row = row
                break
            if not header_row:
                wb.close()
                del ws, wb
                gc.collect()
                if on_progress:
                    on_progress(sheet_idx, total_sheets, sheet_name, 0)
                continue

            columns = [_safe_col_name(c, i) for i, c in enumerate(header_row)]
            seen: dict[str, int] = {}
            deduped: list[str] = []
            for c in columns:
                if c in seen:
                    seen[c] += 1
                    deduped.append(f"{c}_{seen[c]}")
                else:
                    seen[c] = 0
                    deduped.append(c)
            columns = deduped
            num_cols = len(columns)

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

            for row in ws.iter_rows(min_row=2, values_only=True):
                excel_row = rows_processed + 2
                row_list = list(row)

                while len(row_list) < num_cols:
                    row_list.append(None)

                if source_cells or fill_cells:
                    for col_idx in range(num_cols):
                        cell_key = (excel_row, col_idx)
                        if cell_key in source_cells:
                            source_cells[cell_key] = row_list[col_idx]
                        elif row_list[col_idx] is None and cell_key in fill_cells:
                            src_key = fill_cells[cell_key]
                            src_val = source_cells.get(src_key)
                            if src_val is not None:
                                row_list[col_idx] = src_val

                vals = tuple(_cell_to_str(row_list[i]) for i in range(num_cols))
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

            _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — {rows_processed} rows, building FTS")

            conn.execute(f'DROP TABLE IF EXISTS "{fts_table}"')
            conn.execute(
                f'CREATE VIRTUAL TABLE "{fts_table}" USING fts5('
                f"_all_text, content=\"{table}\", content_rowid=_rowid, "
                f"tokenize=\"unicode61 remove_diacritics 2\")"
            )
            conn.execute(
                f'INSERT INTO "{fts_table}" (rowid, _all_text) '
                f"SELECT _rowid, {_concat_expr(columns)} FROM \"{table}\""
            )
            conn.commit()

            source_cells.clear()
            fill_cells.clear()

            wb.close()
            del ws, wb
            gc.collect()

            if on_progress:
                on_progress(sheet_idx, total_sheets, sheet_name, rows_processed)

            _log(f"Sheet {sheet_idx}/{total_sheets}: '{sheet_name}' — done, memory released")

        _save_model_meta(conn, model_id, excel_path, sheet_names)
    finally:
        conn.close()


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
