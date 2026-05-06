"""
Flask application – API routes for the BYD logistics search & print tool.

Performance architecture:
  - Upload: file → GCS (Excel) → background thread processes → SQLite
  - After processing: SQLite DB → GCS (cached DB)
  - Cold start: GCS (cached DB) → local SQLite (seconds, no re-parsing)
  - In-memory set tracks which users' DBs are already loaded locally

Local mode (LOCAL_MODE=1):
  - Skips Firebase Auth, GCS, and Firestore
  - Uses a fixed UID "local_user" for all requests
  - All data stays on the local filesystem
  - No cloud dependencies required
"""

import os
import sys
import uuid
import time
import gc
import sqlite3
import functools
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from flask import Flask, jsonify, request, render_template, send_from_directory
from werkzeug.utils import secure_filename

LOCAL_MODE = os.environ.get("LOCAL_MODE", "").lower() in ("1", "true")

if not LOCAL_MODE:
    import firebase_admin
    from firebase_admin import auth as fb_auth, firestore as fb_firestore
    from gcs_helper import (
        upload_excel as gcs_upload,
        download_excel as gcs_download,
        download_db as gcs_download_db,
        upload_db as gcs_upload_db,
        list_user_models,
        delete_model_files as gcs_delete_model,
        get_model_file_info as gcs_file_info,
        download_from_firebase_storage,
    )
    firebase_admin.initialize_app()

from config import MODELS, UPLOAD_DIR, get_user_db_path, get_user_upload_dir
from data_loader import load_excel_to_db, get_model_sheets, get_all_loaded_models
from search_engine import search
from template_engine import (
    list_templates,
    get_template,
    save_template,
    delete_template,
    render_template as render_print_template,
    render_template_raw,
)

app = Flask(__name__, static_folder="public", static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024 if LOCAL_MODE else 200 * 1024 * 1024

ALLOWED_ORIGINS = os.environ.get("CORS_ORIGINS", "*")
LOCAL_UID = "local_user"

_local_jobs: dict[str, dict] = {}

# In-memory cache: set of UIDs whose SQLite DB is already present locally
_db_ready: set[str] = set()
_db_lock = threading.Lock()
_executor = ThreadPoolExecutor(max_workers=3, thread_name_prefix="upload")


def _log(msg):
    print(f"[BYD] {msg}", file=sys.stderr, flush=True)


@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = ALLOWED_ORIGINS
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
    return response


os.makedirs(UPLOAD_DIR, exist_ok=True)


def _get_uid() -> str | None:
    if LOCAL_MODE:
        return LOCAL_UID
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return None
    token = auth_header[7:]
    try:
        decoded = fb_auth.verify_id_token(token)
        return decoded["uid"]
    except Exception:
        return None


def require_auth(fn):
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        uid = _get_uid()
        if not uid:
            return jsonify({"error": "Unauthorized"}), 401
        kwargs["uid"] = uid
        return fn(*args, **kwargs)
    return wrapper


# ── DB recovery from GCS (fast path) ───────────────────────────────────

def _ensure_user_db(uid: str):
    with _db_lock:
        if uid in _db_ready:
            return

    db_path = get_user_db_path(uid)
    if os.path.isfile(db_path):
        with _db_lock:
            _db_ready.add(uid)
        return

    if not LOCAL_MODE:
        t0 = time.time()
        ok = gcs_download_db(uid, db_path)
        elapsed = time.time() - t0
        if ok:
            _log(f"DB restored from GCS for user {uid[:8]}... in {elapsed:.1f}s")
    with _db_lock:
        _db_ready.add(uid)


# ── Background processing ───────────────────────────────────────────────

def _update_job(uid: str, job_id: str, data: dict):
    data["updatedAt"] = datetime.now(timezone.utc).isoformat()
    if LOCAL_MODE:
        if job_id not in _local_jobs:
            _local_jobs[job_id] = {}
        _local_jobs[job_id].update(data)
        return
    try:
        db = fb_firestore.client()
        ref = db.collection("users").document(uid).collection("jobs").document(job_id)
        ref.set(data, merge=True)
    except Exception as e:
        _log(f"Firestore update FAILED for job {job_id}: {e}")


def _process_upload(uid: str, model_id: str, excel_path: str, job_id: str, mode: str):
    _log(f"Processing started: job={job_id}, model={model_id}")
    db_path = get_user_db_path(uid)
    last_update = [0.0]

    def on_progress(sheet_idx, total_sheets, sheet_name, rows_processed):
        now = time.time()
        if now - last_update[0] < 2.0:
            return
        last_update[0] = now
        _update_job(uid, job_id, {
            "status": "processing",
            "currentSheet": sheet_idx,
            "totalSheets": total_sheets,
            "currentSheetName": sheet_name,
            "rowsProcessed": rows_processed,
        })

    try:
        _update_job(uid, job_id, {
            "status": "processing",
            "currentSheet": 0,
            "totalSheets": 0,
            "rowsProcessed": 0,
        })

        load_excel_to_db(
            model_id, excel_path, db_path, mode=mode,
            on_progress=on_progress,
        )
        gc.collect()

        if not LOCAL_MODE:
            gcs_upload_db(uid, db_path)

        with _db_lock:
            _db_ready.add(uid)

        sheets = get_model_sheets(model_id, db_path)
        sheet_data = [
            {"sheet_name": s["sheet_name"], "table_name": s["table_name"],
             "columns": s["columns"], "row_count": s["row_count"]}
            for s in sheets
        ]

        # Clean up the local temporary Excel file
        try:
            if os.path.isfile(excel_path):
                os.remove(excel_path)
                _log(f"Cleaned up local Excel: {excel_path}")
        except OSError:
            pass

        _log(f"Processing done: job={job_id}, {len(sheet_data)} sheets")
        _update_job(uid, job_id, {"status": "done", "sheets": sheet_data})
    except Exception as e:
        _log(f"Processing FAILED: job={job_id}: {traceback.format_exc()}")
        _update_job(uid, job_id, {"status": "error", "error": str(e)})


# ── Pages ────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    if LOCAL_MODE:
        return send_from_directory("public", "index.html")
    return render_template("index.html")


@app.route("/api/config")
def api_config():
    return jsonify({"localMode": LOCAL_MODE})


@app.route("/api/jobs/<job_id>")
@require_auth
def api_job_status(job_id, uid):
    """Poll job status (local mode only; cloud uses Firestore realtime)."""
    job = _local_jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


# ── Model / Data API ────────────────────────────────────────────────────

@app.route("/api/models")
@require_auth
def api_models(uid):
    _ensure_user_db(uid)
    db_path = get_user_db_path(uid)
    loaded_local = set(get_all_loaded_models(db_path))

    gcs_models = set()
    if not LOCAL_MODE and not loaded_local:
        try:
            gcs_models = set(list_user_models(uid))
        except Exception:
            pass

    result = []
    for mid, info in MODELS.items():
        result.append({
            "id": mid,
            "name": info["name"],
            "loaded": mid in loaded_local or mid in gcs_models,
        })
    return jsonify(result)


@app.route("/api/models/<model_id>/upload", methods=["POST", "OPTIONS"])
@require_auth
def api_upload_model(model_id, uid):
    if request.method == "OPTIONS":
        return jsonify({}), 200
    if model_id not in MODELS:
        return jsonify({"error": "Unknown model"}), 404

    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "No file uploaded"}), 400

    upload_dir = get_user_upload_dir(uid)
    fname = secure_filename(uploaded.filename)
    excel_path = os.path.join(upload_dir, fname)
    uploaded.save(excel_path)
    mode = request.form.get("mode", "replace")
    if mode not in ("replace", "append"):
        mode = "replace"

    if not LOCAL_MODE:
        try:
            gcs_upload(uid, model_id, excel_path)
        except Exception as e:
            _log(f"GCS upload failed: {e}")

    job_id = uuid.uuid4().hex[:16]
    _update_job(uid, job_id, {
        "status": "queued",
        "modelId": model_id,
        "fileName": fname,
        "currentSheet": 0,
        "totalSheets": 0,
        "rowsProcessed": 0,
    })

    _executor.submit(_process_upload, uid, model_id, excel_path, job_id, mode)
    return jsonify({"ok": True, "jobId": job_id})


@app.route("/api/models/<model_id>/process", methods=["POST", "OPTIONS"])
@require_auth
def api_process_model(model_id, uid):
    """Process an Excel file that was uploaded directly to Firebase Storage."""
    if request.method == "OPTIONS":
        return jsonify({}), 200
    if model_id not in MODELS:
        return jsonify({"error": "Unknown model"}), 404

    data = request.get_json(force=True)
    storage_path = data.get("storagePath", "")
    mode = data.get("mode", "replace")
    if mode not in ("replace", "append"):
        mode = "replace"

    if not storage_path or not storage_path.startswith(f"uploads/{uid}/"):
        return jsonify({"error": "Invalid storage path"}), 400

    upload_dir = get_user_upload_dir(uid)
    excel_path = os.path.join(upload_dir, "latest.xlsx")

    try:
        ok = download_from_firebase_storage(storage_path, excel_path)
        if not ok:
            return jsonify({"error": "File not found in Storage"}), 404
    except Exception as e:
        _log(f"Firebase Storage download failed: {e}")
        return jsonify({"error": "Failed to download file from Storage"}), 500

    job_id = uuid.uuid4().hex[:16]
    _update_job(uid, job_id, {
        "status": "queued",
        "modelId": model_id,
        "fileName": storage_path.split("/")[-1],
        "currentSheet": 0,
        "totalSheets": 0,
        "rowsProcessed": 0,
    })

    _executor.submit(_process_upload, uid, model_id, excel_path, job_id, mode)
    return jsonify({"ok": True, "jobId": job_id})


@app.route("/api/models/<model_id>/sheets")
@require_auth
def api_sheets(model_id, uid):
    _ensure_user_db(uid)
    db_path = get_user_db_path(uid)
    sheets = get_model_sheets(model_id, db_path)
    return jsonify(sheets)


@app.route("/api/models/<model_id>/info")
@require_auth
def api_model_info(model_id, uid):
    """Get file info (size, last updated) for a model's stored Excel."""
    if LOCAL_MODE:
        upload_dir = get_user_upload_dir(uid)
        for f in os.listdir(upload_dir) if os.path.isdir(upload_dir) else []:
            if f.endswith(".xlsx"):
                fp = os.path.join(upload_dir, f)
                stat = os.stat(fp)
                return jsonify({"exists": True, "size": stat.st_size,
                                "updated": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat()})
        return jsonify({"exists": False})
    info = None
    try:
        info = gcs_file_info(uid, model_id)
    except Exception as e:
        _log(f"GCS file info failed: {e}")
    if not info:
        return jsonify({"exists": False})
    return jsonify({"exists": True, **info})


@app.route("/api/models/<model_id>/data", methods=["DELETE", "OPTIONS"])
@require_auth
def api_delete_model_data(model_id, uid):
    """Delete all data for a model: SQLite tables + GCS files."""
    if request.method == "OPTIONS":
        return jsonify({}), 200

    db_path = get_user_db_path(uid)

    # Drop SQLite tables for this model
    if os.path.isfile(db_path):
        try:
            conn = sqlite3.connect(db_path)
            from data_loader import _drop_model_tables
            _drop_model_tables(conn, model_id)
            conn.commit()
            conn.close()
            _log(f"Dropped SQLite tables for model={model_id}, user={uid[:8]}")
        except Exception as e:
            _log(f"SQLite cleanup error: {e}")

    if not LOCAL_MODE:
        if os.path.isfile(db_path):
            try:
                gcs_upload_db(uid, db_path)
            except Exception as e:
                _log(f"GCS DB re-upload failed: {e}")
        try:
            gcs_delete_model(uid, model_id)
        except Exception as e:
            _log(f"GCS delete failed: {e}")
        try:
            db = fb_firestore.client()
            uploads_ref = db.collection("users").document(uid).collection("uploads")
            docs = uploads_ref.where("model", "==", model_id).stream()
            for doc in docs:
                doc.reference.delete()
        except Exception as e:
            _log(f"Firestore cleanup error: {e}")

    return jsonify({"ok": True, "message": f"Model {model_id} data deleted"})


# ── Search API ───────────────────────────────────────────────────────────

def _is_valid_table(uid: str, table: str) -> bool:
    """Check that the table exists in the user's loaded model metadata."""
    db_path = get_user_db_path(uid)
    if not os.path.isfile(db_path):
        return False
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.execute(
            "SELECT 1 FROM model_meta WHERE table_name = ? LIMIT 1", (table,)
        )
        found = cur.fetchone() is not None
        conn.close()
        return found
    except Exception:
        return False


@app.route("/api/search")
@require_auth
def api_search(uid):
    table = request.args.get("table", "")
    query = request.args.get("q", "")

    try:
        limit = min(int(request.args.get("limit", 200)), 1000)
        offset = max(int(request.args.get("offset", 0)), 0)
    except (ValueError, TypeError):
        limit, offset = 200, 0

    if not table or not query:
        return jsonify({"columns": [], "rows": [], "total": 0})

    _ensure_user_db(uid)

    if not _is_valid_table(uid, table):
        return jsonify({"error": "Invalid table"}), 400

    db_path = get_user_db_path(uid)

    try:
        result = search(table, query, db_path, limit=limit, offset=offset)
    except Exception as e:
        _log(f"Search error: {e}")
        return jsonify({"error": "Search failed"}), 500

    return jsonify(result)


# ── Template API ─────────────────────────────────────────────────────────

@app.route("/api/templates")
@require_auth
def api_list_templates(uid):
    return jsonify(list_templates())

@app.route("/api/templates/<tpl_id>")
@require_auth
def api_get_template(tpl_id, uid):
    tpl = get_template(tpl_id)
    if not tpl:
        return jsonify({"error": "Template not found"}), 404
    return jsonify(tpl)

@app.route("/api/templates", methods=["POST"])
@require_auth
def api_save_template(uid):
    data = request.get_json(force=True)
    tpl = save_template(name=data.get("name", "Untitled"), body=data.get("body", ""), tpl_id=data.get("id"))
    return jsonify(tpl)

@app.route("/api/templates/<tpl_id>", methods=["PUT"])
@require_auth
def api_update_template(tpl_id, uid):
    data = request.get_json(force=True)
    tpl = save_template(name=data.get("name", "Untitled"), body=data.get("body", ""), tpl_id=tpl_id)
    return jsonify(tpl)

@app.route("/api/templates/<tpl_id>", methods=["DELETE"])
@require_auth
def api_delete_template(tpl_id, uid):
    return jsonify({"ok": delete_template(tpl_id)})


# ── Render API ───────────────────────────────────────────────────────────

@app.route("/api/render", methods=["POST"])
@require_auth
def api_render(uid):
    data = request.get_json(force=True)
    rows = data.get("rows", [])
    body = data.get("body")
    tpl_id = data.get("template_id")
    if body:
        html = render_template_raw(body, rows)
    elif tpl_id:
        html = render_print_template(tpl_id, rows)
    else:
        return jsonify({"error": "Provide template_id or body"}), 400
    return jsonify({"html": html})


if __name__ == "__main__":
    if LOCAL_MODE:
        _log("=" * 50)
        _log("RUNNING IN LOCAL MODE")
        _log("No Firebase Auth / GCS / Firestore required")
        _log("Open http://localhost:5000 in your browser")
        _log("=" * 50)
    app.run(debug=True, port=5000)
