"""
Google Cloud Storage helper for persisting user data.

Stores:
  gs://BUCKET/{uid}/{model_id}/latest.xlsx   — raw Excel file
  gs://BUCKET/{uid}/db/byd_search.db         — pre-built SQLite database
"""

import os
import sys
from google.cloud import storage

BUCKET_NAME = os.environ.get("GCS_BUCKET", "byd-search-user-data")
FIREBASE_STORAGE_BUCKET = os.environ.get(
    "FIREBASE_STORAGE_BUCKET",
    "project-d5a525d1-72ba-431c-80c.firebasestorage.app",
)

_client = None


def _log(msg):
    print(f"[GCS] {msg}", file=sys.stderr, flush=True)


def _get_client():
    global _client
    if _client is None:
        _client = storage.Client()
    return _client


def upload_excel(uid: str, model_id: str, local_path: str):
    blob_path = f"{uid}/{model_id}/latest.xlsx"
    bucket = _get_client().bucket(BUCKET_NAME)
    blob = bucket.blob(blob_path)
    blob.upload_from_filename(local_path)
    _log(f"Uploaded Excel: {blob_path}")


def download_excel(uid: str, model_id: str, local_path: str) -> bool:
    blob_path = f"{uid}/{model_id}/latest.xlsx"
    bucket = _get_client().bucket(BUCKET_NAME)
    blob = bucket.blob(blob_path)
    if not blob.exists():
        return False
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    blob.download_to_filename(local_path)
    _log(f"Downloaded Excel: {blob_path}")
    return True


def upload_db(uid: str, local_db_path: str):
    """Upload the pre-built SQLite database to GCS for fast cold-start recovery."""
    if not os.path.isfile(local_db_path):
        return
    blob_path = f"{uid}/db/byd_search.db"
    bucket = _get_client().bucket(BUCKET_NAME)
    blob = bucket.blob(blob_path)
    blob.upload_from_filename(local_db_path)
    size_mb = os.path.getsize(local_db_path) / (1024 * 1024)
    _log(f"Uploaded DB: {blob_path} ({size_mb:.1f} MB)")


def download_db(uid: str, local_db_path: str) -> bool:
    """Download the pre-built SQLite database from GCS. Much faster than rebuilding from Excel."""
    blob_path = f"{uid}/db/byd_search.db"
    bucket = _get_client().bucket(BUCKET_NAME)
    blob = bucket.blob(blob_path)
    if not blob.exists():
        return False
    os.makedirs(os.path.dirname(local_db_path), exist_ok=True)
    blob.download_to_filename(local_db_path)
    size_mb = os.path.getsize(local_db_path) / (1024 * 1024)
    _log(f"Downloaded DB: {blob_path} ({size_mb:.1f} MB)")
    return True


def delete_model_files(uid: str, model_id: str):
    """Delete the Excel file for a model from GCS.

    Note: The SQLite DB blob is shared across all models and is NOT deleted here.
    It is re-uploaded after local cleanup by the caller.
    """
    bucket = _get_client().bucket(BUCKET_NAME)
    blob_path = f"{uid}/{model_id}/latest.xlsx"
    blob = bucket.blob(blob_path)
    if blob.exists():
        blob.delete()
        _log(f"Deleted: {blob_path}")


def get_model_file_info(uid: str, model_id: str) -> dict | None:
    """Get metadata for the stored Excel file."""
    blob_path = f"{uid}/{model_id}/latest.xlsx"
    bucket = _get_client().bucket(BUCKET_NAME)
    blob = bucket.blob(blob_path)
    if not blob.exists():
        return None
    blob.reload()
    return {
        "size_bytes": blob.size or 0,
        "size_mb": round((blob.size or 0) / (1024 * 1024), 2),
        "updated": blob.updated.isoformat() if blob.updated else None,
    }


def download_from_firebase_storage(storage_path: str, local_path: str) -> bool:
    """Download a file from the Firebase Storage bucket (used for direct uploads)."""
    bucket = _get_client().bucket(FIREBASE_STORAGE_BUCKET)
    blob = bucket.blob(storage_path)
    if not blob.exists():
        _log(f"Firebase Storage blob not found: {storage_path}")
        return False
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    blob.download_to_filename(local_path)
    size_mb = os.path.getsize(local_path) / (1024 * 1024)
    _log(f"Downloaded from Firebase Storage: {storage_path} ({size_mb:.1f} MB)")
    return True


def list_user_models(uid: str) -> list[str]:
    prefix = f"{uid}/"
    bucket = _get_client().bucket(BUCKET_NAME)
    blobs = bucket.list_blobs(prefix=prefix)
    models = set()
    for blob in blobs:
        parts = blob.name.split("/")
        if len(parts) >= 3 and parts[2] == "latest.xlsx":
            models.add(parts[1])
    return list(models)
