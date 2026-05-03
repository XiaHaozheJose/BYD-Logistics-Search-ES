"""
Print-template management: CRUD for user-defined templates and rendering
rows into them.

Templates are stored as JSON files in PRINT_TEMPLATES_DIR, each containing:
    { "name": "...", "body": "...(HTML with {{col}} placeholders)..." }

Rendering uses Python's str-based replacement (not Jinja2) so that
double-brace syntax in user templates is kept simple and predictable.
"""

import json
import os
import re
import uuid

from config import PRINT_TEMPLATES_DIR

_PLACEHOLDER_RE = re.compile(r"\{\{(.+?)\}\}")

os.makedirs(PRINT_TEMPLATES_DIR, exist_ok=True)


def _tpl_path(tpl_id: str) -> str:
    return os.path.join(PRINT_TEMPLATES_DIR, f"{tpl_id}.json")


def list_templates() -> list[dict]:
    results = []
    for fname in sorted(os.listdir(PRINT_TEMPLATES_DIR)):
        if not fname.endswith(".json"):
            continue
        tpl_id = fname[:-5]
        try:
            with open(os.path.join(PRINT_TEMPLATES_DIR, fname), "r", encoding="utf-8") as f:
                data = json.load(f)
            results.append({"id": tpl_id, "name": data.get("name", tpl_id), "body": data.get("body", "")})
        except Exception:
            continue
    return results


def get_template(tpl_id: str) -> dict | None:
    path = _tpl_path(tpl_id)
    if not os.path.isfile(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    data["id"] = tpl_id
    return data


def save_template(name: str, body: str, tpl_id: str | None = None) -> dict:
    if not tpl_id:
        tpl_id = uuid.uuid4().hex[:12]
    path = _tpl_path(tpl_id)
    data = {"name": name, "body": body}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"id": tpl_id, **data}


def delete_template(tpl_id: str) -> bool:
    path = _tpl_path(tpl_id)
    if os.path.isfile(path):
        os.remove(path)
        return True
    return False


def render_template(tpl_id: str, rows: list[dict]) -> str:
    """Render one or more *rows* into the template identified by *tpl_id*.

    Each ``{{column_name}}`` in the template body is replaced with the
    corresponding value from the row dict.  The rendered blocks for each
    row are separated by ``<hr>``.
    """
    tpl = get_template(tpl_id)
    if not tpl:
        return "<p>模板未找到</p>"

    body = tpl["body"]
    rendered_parts: list[str] = []
    for row in rows:
        text = body
        for match in _PLACEHOLDER_RE.finditer(body):
            col_name = match.group(1).strip()
            value = row.get(col_name, "")
            text = text.replace(match.group(0), str(value) if value else "")
        rendered_parts.append(text)

    return "\n<hr class='my-4 border-gray-300'>\n".join(rendered_parts)


def render_template_raw(body: str, rows: list[dict]) -> str:
    """Render rows into an ad-hoc template body (not saved)."""
    rendered_parts: list[str] = []
    for row in rows:
        text = body
        for match in _PLACEHOLDER_RE.finditer(body):
            col_name = match.group(1).strip()
            value = row.get(col_name, "")
            text = text.replace(match.group(0), str(value) if value else "")
        rendered_parts.append(text)
    return "\n<hr class='my-4 border-gray-300'>\n".join(rendered_parts)


# Seed a default template if none exist
if not list_templates():
    save_template(
        name="默认模板",
        body=(
            "<h3>订单查询结果</h3>\n"
            "<table class='table-auto border-collapse border border-gray-400 w-full'>\n"
            "  <tbody>\n"
            "    {{#each}}\n"
            "  </tbody>\n"
            "</table>"
        ),
        tpl_id="default",
    )
