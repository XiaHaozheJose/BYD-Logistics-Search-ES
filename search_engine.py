"""
Full-text search engine backed by SQLite FTS5.
"""

import re
import sqlite3

_SAFE_TABLE_RE = re.compile(r'^[a-zA-Z0-9_\u4e00-\u9fff]+$')


def search(
    table_name: str,
    query: str,
    db_path: str,
    limit: int = 200,
    offset: int = 0,
) -> dict:
    if not _SAFE_TABLE_RE.match(table_name):
        raise ValueError(f"Invalid table name: {table_name}")
    limit = min(max(limit, 1), 1000)
    offset = max(offset, 0)
    fts_table = f"{table_name}__fts"
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        col_cur = conn.execute(f'PRAGMA table_info("{table_name}")')
        columns = [r[1] for r in col_cur.fetchall() if r[1] != "_rowid"]

        fts_expr = _build_fts_expr(query)
        if not fts_expr:
            return {"columns": columns, "rows": [], "total": 0}

        col_select = ", ".join(f't."{c}"' for c in columns)

        count_sql = (
            f'SELECT COUNT(*) FROM "{fts_table}" f '
            f'JOIN "{table_name}" t ON t._rowid = f.rowid '
            f"WHERE f._all_text MATCH ?"
        )
        total = conn.execute(count_sql, (fts_expr,)).fetchone()[0]

        data_sql = (
            f"SELECT t._rowid AS _id, {col_select} "
            f'FROM "{fts_table}" f '
            f'JOIN "{table_name}" t ON t._rowid = f.rowid '
            f"WHERE f._all_text MATCH ? "
            f"ORDER BY f.rank "
            f"LIMIT ? OFFSET ?"
        )
        rows = []
        for r in conn.execute(data_sql, (fts_expr, limit, offset)):
            row_dict = {c: r[c] for c in columns}
            row_dict["_id"] = r["_id"]
            rows.append(row_dict)

        return {"columns": columns, "rows": rows, "total": total}
    finally:
        conn.close()


def _build_fts_expr(query: str) -> str:
    tokens = query.strip().split()
    tokens = [t for t in tokens if len(t) >= 1]
    if not tokens:
        return ""
    parts = []
    for t in tokens:
        safe = t.replace('"', '""')
        parts.append(f'"{safe}"')
    return " AND ".join(parts)
