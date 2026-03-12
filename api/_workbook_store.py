from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

from werkzeug.datastructures import FileStorage

DEFAULT_MAX_UPLOAD_BYTES = 30 * 1024 * 1024


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _default_data_root() -> Path:
    configured = os.getenv("CAPACITY_SHARED_DATA_DIR", "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    return Path("/tmp/capacity_shared_store").resolve()


def _default_manifest() -> Dict[str, Any]:
    return {"version": 1, "fileName": "", "uploadedAt": None, "sizeBytes": 0}


class WorkbookStore:
    def __init__(self) -> None:
        self.data_root = _default_data_root()
        self.active_workbook_path = self.data_root / "active_workbook.xlsx"
        self.manifest_path = self.data_root / "manifest.json"
        max_bytes_raw = os.getenv("CAPACITY_MAX_UPLOAD_BYTES", str(DEFAULT_MAX_UPLOAD_BYTES))
        try:
            self.max_upload_bytes = max(1, int(max_bytes_raw))
        except (TypeError, ValueError):
            self.max_upload_bytes = DEFAULT_MAX_UPLOAD_BYTES
        self._ensure_store()

    def _ensure_store(self) -> None:
        self.data_root.mkdir(parents=True, exist_ok=True)
        if not self.manifest_path.exists():
            self._write_manifest(_default_manifest())

    def _read_manifest(self) -> Dict[str, Any]:
        if not self.manifest_path.exists():
            return _default_manifest()
        try:
            with self.manifest_path.open("r", encoding="utf-8") as handle:
                payload = json.load(handle)
                if isinstance(payload, dict):
                    return payload
        except Exception:
            pass
        return _default_manifest()

    def _write_manifest(self, payload: Dict[str, Any]) -> None:
        fd, temp_path = tempfile.mkstemp(prefix=f"{self.manifest_path.name}.", suffix=".tmp", dir=str(self.data_root))
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=True, separators=(",", ":"))
            os.replace(temp_path, self.manifest_path)
        except Exception:
            try:
                os.unlink(temp_path)
            except OSError:
                pass
            raise

    def workbook_state(self) -> Dict[str, Any]:
        manifest = self._read_manifest()
        has_workbook = self.active_workbook_path.exists()
        size_bytes = int(manifest.get("sizeBytes") or 0)
        if has_workbook and size_bytes <= 0:
            size_bytes = self.active_workbook_path.stat().st_size
        return {
            "hasWorkbook": has_workbook,
            "fileName": str(manifest.get("fileName") or (self.active_workbook_path.name if has_workbook else "")),
            "uploadedAt": manifest.get("uploadedAt"),
            "sizeBytes": size_bytes if has_workbook else 0,
        }

    def save_upload(self, uploaded_file: FileStorage) -> Dict[str, Any]:
        if uploaded_file is None or not uploaded_file.filename:
            raise ValueError("Missing upload file. Use form field 'file' with a .xlsx workbook.")

        original_name = Path(uploaded_file.filename).name
        if not original_name.lower().endswith(".xlsx"):
            raise ValueError("Only .xlsx files are allowed.")

        temp_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, dir=str(self.data_root), prefix="active_", suffix=".tmp") as temp_file:
                uploaded_file.save(temp_file)
                temp_path = Path(temp_file.name)

            if temp_path is None or not temp_path.exists():
                raise RuntimeError("Saving the workbook file failed.")

            size_bytes = temp_path.stat().st_size
            if size_bytes <= 0:
                raise ValueError("Uploaded workbook is empty.")

            if size_bytes > self.max_upload_bytes:
                raise ValueError(f"Workbook exceeds maximum size of {self.max_upload_bytes} bytes.")

            with temp_path.open("rb") as handle:
                signature = handle.read(2)
                if signature != b"PK":
                    raise ValueError("Invalid workbook payload. Expected .xlsx content.")

            os.replace(temp_path, self.active_workbook_path)
            manifest = {
                "version": 1,
                "fileName": original_name,
                "uploadedAt": _utc_now_iso(),
                "sizeBytes": size_bytes,
            }
            self._write_manifest(manifest)
            return {
                "hasWorkbook": True,
                "fileName": original_name,
                "uploadedAt": manifest["uploadedAt"],
                "sizeBytes": size_bytes,
            }
        finally:
            if temp_path is not None and temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass


store = WorkbookStore()
