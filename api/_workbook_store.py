from __future__ import annotations

import json
import os
import tempfile
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from werkzeug.datastructures import FileStorage

try:
    from vercel.blob import list_objects, put
except Exception:  # pragma: no cover - optional dependency when running locally without vercel package
    list_objects = None  # type: ignore[assignment]
    put = None  # type: ignore[assignment]

DATASETS = {"main", "sales"}
DEFAULT_MAX_UPLOAD_BYTES = 30 * 1024 * 1024
WORKBOOK_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _timestamp_ms() -> int:
    return int(datetime.now(timezone.utc).timestamp() * 1000)


def _default_data_root() -> Path:
    configured = os.getenv("CAPACITY_SHARED_DATA_DIR", "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    return Path("/tmp/capacity_shared_store").resolve()


def _sanitize_file_name(file_name: str) -> str:
    base = Path(file_name).name.strip()
    safe = "".join(char if char.isalnum() or char in {"-", "_", "."} else "_" for char in base)
    return safe or "workbook.xlsx"


def _obj_get(obj: Any, key: str) -> Any:
    if isinstance(obj, dict):
        return obj.get(key)
    value = getattr(obj, key, None)
    if value is not None:
        return value
    if key == "uploadedAt":
        return getattr(obj, "uploaded_at", None)
    return None


def _as_blob_dict(blob_obj: Any) -> Dict[str, Any]:
    if isinstance(blob_obj, dict):
        return blob_obj
    try:
        return dict(blob_obj)
    except Exception:
        return {
            "url": _obj_get(blob_obj, "url"),
            "pathname": _obj_get(blob_obj, "pathname"),
            "uploadedAt": _obj_get(blob_obj, "uploadedAt"),
            "size": _obj_get(blob_obj, "size"),
        }


class WorkbookStore:
    def __init__(self) -> None:
        self.data_root = _default_data_root()
        self.manifest_dir = self.data_root / "manifests"
        self.max_upload_bytes = self._read_max_upload_bytes()
        self.blob_token = os.getenv("BLOB_READ_WRITE_TOKEN", "").strip()
        self.use_blob = bool(self.blob_token and put is not None and list_objects is not None)
        self._ensure_store()

    def _read_max_upload_bytes(self) -> int:
        raw = os.getenv("CAPACITY_MAX_UPLOAD_BYTES", str(DEFAULT_MAX_UPLOAD_BYTES))
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_UPLOAD_BYTES

    def _ensure_store(self) -> None:
        self.data_root.mkdir(parents=True, exist_ok=True)
        self.manifest_dir.mkdir(parents=True, exist_ok=True)
        for dataset in DATASETS:
            if not self._local_manifest_path(dataset).exists():
                self._write_local_manifest(dataset, self._default_manifest(dataset))

    def _validate_dataset(self, dataset: str) -> str:
        normalized = (dataset or "").strip().lower()
        if normalized not in DATASETS:
            raise ValueError("Invalid dataset. Allowed values: main, sales.")
        return normalized

    def _default_manifest(self, dataset: str) -> Dict[str, Any]:
        return {
            "version": 1,
            "dataset": dataset,
            "fileName": "",
            "workbookUrl": "",
            "workbookPath": "",
            "uploadedAt": None,
            "sizeBytes": 0,
        }

    def _local_workbook_path(self, dataset: str) -> Path:
        return self.data_root / f"active_{dataset}.xlsx"

    def _local_manifest_path(self, dataset: str) -> Path:
        return self.manifest_dir / f"{dataset}.json"

    def _write_json_atomic(self, path: Path, payload: Dict[str, Any]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        fd, temp_path = tempfile.mkstemp(prefix=f"{path.name}.", suffix=".tmp", dir=str(path.parent))
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=True, separators=(",", ":"))
            os.replace(temp_path, path)
        except Exception:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
            raise

    def _write_local_manifest(self, dataset: str, payload: Dict[str, Any]) -> None:
        self._write_json_atomic(self._local_manifest_path(dataset), payload)

    def _read_local_manifest(self, dataset: str) -> Dict[str, Any]:
        path = self._local_manifest_path(dataset)
        if not path.exists():
            return self._default_manifest(dataset)
        try:
            with path.open("r", encoding="utf-8") as handle:
                payload = json.load(handle)
                if isinstance(payload, dict):
                    return payload
        except Exception:
            pass
        return self._default_manifest(dataset)

    def _blob_common_kwargs(self) -> Dict[str, Any]:
        kwargs: Dict[str, Any] = {}
        if self.blob_token:
            kwargs["token"] = self.blob_token
        return kwargs

    def _blob_put(self, pathname: str, body: bytes, content_type: str) -> Dict[str, Any]:
        if not self.use_blob:
            raise RuntimeError("Blob storage is not configured.")
        uploaded = put(
            pathname,
            body,
            access="public",
            add_random_suffix=False,
            content_type=content_type,
            **self._blob_common_kwargs(),
        )
        return _as_blob_dict(uploaded)

    def _blob_latest_manifest(self, dataset: str) -> Optional[Dict[str, Any]]:
        if not self.use_blob:
            return None
        prefix = f"capacity-dashboard/manifests/{dataset}/"
        listing = list_objects(prefix=prefix, limit=1000, **self._blob_common_kwargs())
        blobs = _obj_get(listing, "blobs") or []
        if not isinstance(blobs, list) or len(blobs) == 0:
            return None
        latest_blob = max(blobs, key=lambda item: str(_obj_get(item, "pathname") or ""))
        latest_blob_url = str(_obj_get(latest_blob, "url") or "")
        if not latest_blob_url:
            return None
        with urllib.request.urlopen(latest_blob_url, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8"))
            if isinstance(payload, dict):
                return payload
        return None

    def _blob_fetch_workbook_bytes(self, workbook_url: str) -> bytes:
        with urllib.request.urlopen(workbook_url, timeout=30) as response:
            return response.read()

    def workbook_state(self, dataset: str) -> Dict[str, Any]:
        normalized = self._validate_dataset(dataset)
        if self.use_blob:
            manifest = self._blob_latest_manifest(normalized)
            if not manifest:
                return {
                    "dataset": normalized,
                    "hasWorkbook": False,
                    "fileName": "",
                    "uploadedAt": None,
                    "sizeBytes": 0,
                }
            return {
                "dataset": normalized,
                "hasWorkbook": True,
                "fileName": str(manifest.get("fileName") or ""),
                "uploadedAt": manifest.get("uploadedAt"),
                "sizeBytes": int(manifest.get("sizeBytes") or 0),
            }

        workbook_path = self._local_workbook_path(normalized)
        manifest = self._read_local_manifest(normalized)
        has_workbook = workbook_path.exists()
        size_bytes = int(manifest.get("sizeBytes") or 0)
        if has_workbook and size_bytes <= 0:
            size_bytes = workbook_path.stat().st_size
        return {
            "dataset": normalized,
            "hasWorkbook": has_workbook,
            "fileName": str(manifest.get("fileName") or (workbook_path.name if has_workbook else "")),
            "uploadedAt": manifest.get("uploadedAt"),
            "sizeBytes": size_bytes if has_workbook else 0,
        }

    def save_upload(self, dataset: str, uploaded_file: FileStorage) -> Dict[str, Any]:
        normalized = self._validate_dataset(dataset)
        if uploaded_file is None or not uploaded_file.filename:
            raise ValueError("Missing upload file. Use form field 'file' with a .xlsx workbook.")

        original_name = _sanitize_file_name(uploaded_file.filename)
        if not original_name.lower().endswith(".xlsx"):
            raise ValueError("Only .xlsx files are allowed.")

        temp_path: Optional[Path] = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, dir=str(self.data_root), prefix=f"{normalized}_", suffix=".tmp") as temp_file:
                uploaded_file.save(temp_file)
                temp_path = Path(temp_file.name)

            if temp_path is None or not temp_path.exists():
                raise RuntimeError("Saving the workbook file failed.")

            size_bytes = temp_path.stat().st_size
            if size_bytes <= 0:
                raise ValueError("Uploaded workbook is empty.")
            if size_bytes > self.max_upload_bytes:
                raise ValueError(f"Workbook exceeds maximum size of {self.max_upload_bytes} bytes.")

            workbook_bytes = temp_path.read_bytes()
            if workbook_bytes[:2] != b"PK":
                raise ValueError("Invalid workbook payload. Expected .xlsx content.")

            uploaded_at = _utc_now_iso()
            if self.use_blob:
                timestamp = _timestamp_ms()
                workbook_path = f"capacity-dashboard/uploads/{normalized}/{timestamp}-{original_name}"
                workbook_blob = self._blob_put(workbook_path, workbook_bytes, WORKBOOK_MIME)
                workbook_url = str(workbook_blob.get("url") or "")
                workbook_storage_path = str(workbook_blob.get("pathname") or workbook_path)
                manifest_payload = {
                    "version": 1,
                    "dataset": normalized,
                    "fileName": original_name,
                    "workbookUrl": workbook_url,
                    "workbookPath": workbook_storage_path,
                    "uploadedAt": uploaded_at,
                    "sizeBytes": size_bytes,
                }
                manifest_path = f"capacity-dashboard/manifests/{normalized}/{timestamp}.json"
                self._blob_put(manifest_path, json.dumps(manifest_payload, separators=(",", ":")).encode("utf-8"), "application/json")
                return {
                    "dataset": normalized,
                    "hasWorkbook": True,
                    "fileName": original_name,
                    "uploadedAt": uploaded_at,
                    "sizeBytes": size_bytes,
                }

            local_workbook = self._local_workbook_path(normalized)
            os.replace(temp_path, local_workbook)
            manifest_payload = {
                "version": 1,
                "dataset": normalized,
                "fileName": original_name,
                "workbookUrl": "",
                "workbookPath": str(local_workbook),
                "uploadedAt": uploaded_at,
                "sizeBytes": size_bytes,
            }
            self._write_local_manifest(normalized, manifest_payload)
            return {
                "dataset": normalized,
                "hasWorkbook": True,
                "fileName": original_name,
                "uploadedAt": uploaded_at,
                "sizeBytes": size_bytes,
            }
        finally:
            if temp_path is not None and temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass

    def workbook_content(self, dataset: str) -> Tuple[bytes, str]:
        normalized = self._validate_dataset(dataset)
        if self.use_blob:
            manifest = self._blob_latest_manifest(normalized)
            if not manifest:
                raise FileNotFoundError(f"No workbook stored for dataset '{normalized}'.")
            workbook_url = str(manifest.get("workbookUrl") or "")
            if not workbook_url:
                raise FileNotFoundError(f"No workbook URL stored for dataset '{normalized}'.")
            workbook_bytes = self._blob_fetch_workbook_bytes(workbook_url)
            workbook_name = str(manifest.get("fileName") or f"{normalized}.xlsx")
            return workbook_bytes, workbook_name

        local_workbook = self._local_workbook_path(normalized)
        if not local_workbook.exists():
            raise FileNotFoundError(f"No workbook stored for dataset '{normalized}'.")
        manifest = self._read_local_manifest(normalized)
        workbook_name = str(manifest.get("fileName") or local_workbook.name)
        return local_workbook.read_bytes(), workbook_name


store = WorkbookStore()
