from __future__ import annotations

import json
import math
import os
import tempfile
import urllib.request
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from werkzeug.datastructures import FileStorage

try:
    from vercel.blob import list_objects, put
except Exception:  # pragma: no cover - optional dependency when running locally without vercel package
    list_objects = None  # type: ignore[assignment]
    put = None  # type: ignore[assignment]

DATASETS = {"main", "sales"}
DEFAULT_MAX_UPLOAD_BYTES = 30 * 1024 * 1024
WORKBOOK_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PLANNING_STATE_MIME = "application/json"
DEFAULT_MAX_PLANNING_OVERRIDES = 100_000
DEFAULT_MAX_WEEK_CAPACITY_SCHEDULE_ENTRIES = 100_000
DEFAULT_MAX_OVERRIDE_HOURS = 1_000_000.0
DEFAULT_MAX_RATE_PROJECTS = 10_000
DEFAULT_MAX_RATE_PER_HOUR = 1_000_000.0


class PlanningStateConflictError(RuntimeError):
    """Raised when a planning state save uses a stale version."""


class RevenueRatesConflictError(RuntimeError):
    """Raised when a revenue rates save uses a stale version."""


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
        self.max_planning_overrides = self._read_max_planning_overrides()
        self.max_week_capacity_schedule_entries = self._read_max_week_capacity_schedule_entries()
        self.max_override_hours = self._read_max_override_hours()
        self.max_rate_projects = self._read_max_rate_projects()
        self.max_rate_per_hour = self._read_max_rate_per_hour()
        self.blob_token = os.getenv("BLOB_READ_WRITE_TOKEN", "").strip()
        self.use_blob = bool(self.blob_token and put is not None and list_objects is not None)
        self._ensure_store()

    def _read_max_upload_bytes(self) -> int:
        raw = os.getenv("CAPACITY_MAX_UPLOAD_BYTES", str(DEFAULT_MAX_UPLOAD_BYTES))
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_UPLOAD_BYTES

    def _read_max_planning_overrides(self) -> int:
        raw = os.getenv("CAPACITY_MAX_PLANNING_OVERRIDES", str(DEFAULT_MAX_PLANNING_OVERRIDES))
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_PLANNING_OVERRIDES

    def _read_max_week_capacity_schedule_entries(self) -> int:
        raw = os.getenv(
            "CAPACITY_MAX_WEEK_CAPACITY_SCHEDULE_ENTRIES",
            os.getenv(
                "CAPACITY_MAX_WEEK_CAPACITY_OVERRIDES",
                str(DEFAULT_MAX_WEEK_CAPACITY_SCHEDULE_ENTRIES),
            ),
        )
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_WEEK_CAPACITY_SCHEDULE_ENTRIES

    def _read_max_override_hours(self) -> float:
        raw = os.getenv("CAPACITY_MAX_OVERRIDE_HOURS", str(DEFAULT_MAX_OVERRIDE_HOURS))
        try:
            return max(0.0, float(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_OVERRIDE_HOURS

    def _read_max_rate_projects(self) -> int:
        raw = os.getenv("CAPACITY_MAX_RATE_PROJECTS", str(DEFAULT_MAX_RATE_PROJECTS))
        try:
            return max(1, int(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_RATE_PROJECTS

    def _read_max_rate_per_hour(self) -> float:
        raw = os.getenv("CAPACITY_MAX_RATE_PER_HOUR", str(DEFAULT_MAX_RATE_PER_HOUR))
        try:
            return max(0.0, float(raw))
        except (TypeError, ValueError):
            return DEFAULT_MAX_RATE_PER_HOUR

    def _ensure_store(self) -> None:
        self.data_root.mkdir(parents=True, exist_ok=True)
        self.manifest_dir.mkdir(parents=True, exist_ok=True)
        for dataset in DATASETS:
            if not self._local_manifest_path(dataset).exists():
                self._write_local_manifest(dataset, self._default_manifest(dataset))
            if not self._local_revenue_rates_path(dataset).exists():
                self._write_local_revenue_rates_state(dataset, self._default_revenue_rates_state(dataset))

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

    def _local_planning_state_path(self, dataset: str) -> Path:
        return self.data_root / f"planning_state_{dataset}.json"

    def _local_revenue_rates_path(self, dataset: str) -> Path:
        return self.data_root / f"revenue_rates_{dataset}.json"

    def _default_planning_state(self, dataset: str) -> Dict[str, Any]:
        return {
            "dataset": dataset,
            "version": 0,
            "updatedAt": None,
            "source": "system",
            "overrideCount": 0,
            "overrides": {},
            "weekCapacityScheduleCount": 0,
            "weekCapacitySchedule": {},
        }

    def _default_revenue_rates_state(self, dataset: str) -> Dict[str, Any]:
        return {
            "dataset": dataset,
            "version": 0,
            "updatedAt": None,
            "source": "system",
            "rateCount": 0,
            "rates": {},
        }

    def _normalize_source(self, source: Any) -> str:
        if not isinstance(source, str):
            return "ui"
        trimmed = source.strip()
        if not trimmed:
            return "ui"
        safe = "".join(char if char.isalnum() or char in {"-", "_", "."} else "_" for char in trimmed)
        return safe[:80] or "ui"

    def _normalize_base_version(self, base_version: Any) -> Optional[int]:
        if base_version is None:
            return None
        if isinstance(base_version, bool):
            raise ValueError("baseVersion must be an integer.")
        try:
            parsed = int(base_version)
        except (TypeError, ValueError) as exc:
            raise ValueError("baseVersion must be an integer.") from exc
        if parsed < 0:
            raise ValueError("baseVersion must be greater than or equal to 0.")
        return parsed

    def _normalize_overrides(self, overrides: Any) -> Dict[str, float]:
        if overrides is None:
            return {}
        if not isinstance(overrides, dict):
            raise ValueError("Invalid planning state payload. 'overrides' must be an object.")
        if len(overrides) > self.max_planning_overrides:
            raise ValueError(f"Planning state exceeds maximum of {self.max_planning_overrides} override entries.")
        normalized: Dict[str, float] = {}
        for raw_key, raw_value in overrides.items():
            if not isinstance(raw_key, str):
                raise ValueError("Invalid planning override key.")
            key = raw_key.strip()
            if not key:
                raise ValueError("Planning override keys cannot be empty.")
            if len(key) > 512 or "\n" in key or "\r" in key:
                raise ValueError("Invalid planning override key format.")
            try:
                value = float(raw_value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Invalid planning override value for key '{key}'.") from exc
            if not math.isfinite(value) or value < 0:
                raise ValueError(f"Planning override value for key '{key}' must be a finite non-negative number.")
            if value > self.max_override_hours:
                raise ValueError(f"Planning override value for key '{key}' exceeds maximum of {self.max_override_hours}.")
            normalized[key] = value
        return normalized

    def _normalize_week_capacity_schedule(self, schedule: Any) -> Dict[str, float]:
        if schedule is None:
            return {}
        if not isinstance(schedule, dict):
            raise ValueError("Invalid planning state payload. 'weekCapacitySchedule' must be an object.")

        normalized: Dict[str, float] = {}
        total_entries = 0
        for raw_week_key, raw_value in schedule.items():
            if not isinstance(raw_week_key, str):
                raise ValueError("Invalid week capacity schedule key.")
            week_key = raw_week_key.strip()
            if not week_key:
                raise ValueError("Week capacity schedule keys cannot be empty.")
            try:
                datetime.strptime(week_key, "%Y-%m-%d")
            except ValueError as exc:
                raise ValueError(f"Invalid week capacity schedule key '{week_key}'.") from exc

            try:
                value = float(raw_value)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Invalid week capacity schedule value for week '{week_key}'.") from exc
            if not math.isfinite(value) or value < 0:
                raise ValueError(
                    f"Week capacity schedule for week '{week_key}' must be a finite non-negative number."
                )
            if value > self.max_override_hours:
                raise ValueError(
                    f"Week capacity schedule for week '{week_key}' exceeds maximum of {self.max_override_hours}."
                )

            normalized[week_key] = value
            total_entries += 1
            if total_entries > self.max_week_capacity_schedule_entries:
                raise ValueError(
                    "Week capacity schedule exceeds maximum of "
                    f"{self.max_week_capacity_schedule_entries} entries."
                )

        return normalized

    def _normalize_rate_value(self, value: Any, project: str, field_name: str) -> float:
        try:
            numeric = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid {field_name} for project '{project}'.") from exc
        if not math.isfinite(numeric) or numeric < 0:
            raise ValueError(f"{field_name} for project '{project}' must be a finite non-negative number.")
        if numeric > self.max_rate_per_hour:
            raise ValueError(
                f"{field_name} for project '{project}' exceeds maximum of {self.max_rate_per_hour}."
            )
        return numeric

    def _normalize_revenue_rates(self, rates: Any) -> Dict[str, Dict[str, float]]:
        if rates is None:
            return {}
        if not isinstance(rates, dict):
            raise ValueError("Invalid revenue rates payload. 'rates' must be an object.")
        if len(rates) > self.max_rate_projects:
            raise ValueError(f"Revenue rates exceed maximum of {self.max_rate_projects} projects.")

        normalized: Dict[str, Dict[str, float]] = {}
        for raw_project, raw_rate in rates.items():
            if not isinstance(raw_project, str):
                raise ValueError("Invalid revenue rate project key.")
            project = raw_project.strip()
            if not project:
                raise ValueError("Revenue rate project keys cannot be empty.")
            if len(project) > 256 or "\n" in project or "\r" in project:
                raise ValueError("Invalid revenue rate project key format.")
            if not isinstance(raw_rate, dict):
                raise ValueError(f"Invalid revenue rate payload for project '{project}'.")

            revenue_per_hour = self._normalize_rate_value(raw_rate.get("revenuePerHour", 0), project, "revenuePerHour")
            gross_profit_per_hour = self._normalize_rate_value(
                raw_rate.get("grossProfitPerHour", 0),
                project,
                "grossProfitPerHour",
            )
            normalized[project] = {
                "revenuePerHour": revenue_per_hour,
                "grossProfitPerHour": gross_profit_per_hour,
            }
        return normalized

    def _coerce_planning_state(self, dataset: str, payload: Any) -> Dict[str, Any]:
        default_state = self._default_planning_state(dataset)
        if not isinstance(payload, dict):
            return default_state

        version = payload.get("version")
        if isinstance(version, bool) or not isinstance(version, int) or version < 0:
            version = 0

        updated_at = payload.get("updatedAt")
        if not isinstance(updated_at, str):
            updated_at = None

        overrides_payload = payload.get("overrides")
        try:
            overrides = self._normalize_overrides(overrides_payload if overrides_payload is not None else {})
        except ValueError:
            overrides = {}

        week_capacity_schedule_payload = payload.get("weekCapacitySchedule")
        try:
            week_capacity_schedule = self._normalize_week_capacity_schedule(
                week_capacity_schedule_payload if week_capacity_schedule_payload is not None else {}
            )
        except ValueError:
            week_capacity_schedule = {}

        return {
            "dataset": dataset,
            "version": version,
            "updatedAt": updated_at,
            "source": self._normalize_source(payload.get("source")),
            "overrideCount": len(overrides),
            "overrides": overrides,
            "weekCapacityScheduleCount": len(week_capacity_schedule),
            "weekCapacitySchedule": week_capacity_schedule,
        }

    def _coerce_revenue_rates_state(self, dataset: str, payload: Any) -> Dict[str, Any]:
        default_state = self._default_revenue_rates_state(dataset)
        if not isinstance(payload, dict):
            return default_state

        version = payload.get("version")
        if isinstance(version, bool) or not isinstance(version, int) or version < 0:
            version = 0

        updated_at = payload.get("updatedAt")
        if not isinstance(updated_at, str):
            updated_at = None

        rates_payload = payload.get("rates")
        try:
            rates = self._normalize_revenue_rates(rates_payload if rates_payload is not None else {})
        except ValueError:
            rates = {}

        return {
            "dataset": dataset,
            "version": version,
            "updatedAt": updated_at,
            "source": self._normalize_source(payload.get("source")),
            "rateCount": len(rates),
            "rates": rates,
        }

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

    def _read_local_planning_state(self, dataset: str) -> Dict[str, Any]:
        path = self._local_planning_state_path(dataset)
        if not path.exists():
            return self._default_planning_state(dataset)
        try:
            with path.open("r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except Exception:
            return self._default_planning_state(dataset)
        return self._coerce_planning_state(dataset, payload)

    def _write_local_planning_state(self, dataset: str, payload: Dict[str, Any]) -> None:
        self._write_json_atomic(self._local_planning_state_path(dataset), payload)

    def _read_local_revenue_rates_state(self, dataset: str) -> Dict[str, Any]:
        path = self._local_revenue_rates_path(dataset)
        if not path.exists():
            return self._default_revenue_rates_state(dataset)
        try:
            with path.open("r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except Exception:
            return self._default_revenue_rates_state(dataset)
        return self._coerce_revenue_rates_state(dataset, payload)

    def _write_local_revenue_rates_state(self, dataset: str, payload: Dict[str, Any]) -> None:
        self._write_json_atomic(self._local_revenue_rates_path(dataset), payload)

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

    def _blob_latest_json_manifest(self, prefix: str) -> Optional[Dict[str, Any]]:
        if not self.use_blob:
            return None
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

    def _blob_latest_manifest(self, dataset: str) -> Optional[Dict[str, Any]]:
        return self._blob_latest_json_manifest(f"capacity-dashboard/manifests/{dataset}/")

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

    def planning_state(self, dataset: str) -> Dict[str, Any]:
        normalized = self._validate_dataset(dataset)
        if self.use_blob:
            manifest = self._blob_latest_json_manifest(f"capacity-dashboard/planning-manifests/{normalized}/")
            if not manifest:
                return self._default_planning_state(normalized)
            state_url = str(manifest.get("stateUrl") or "")
            if not state_url:
                return self._default_planning_state(normalized)
            with urllib.request.urlopen(state_url, timeout=20) as response:
                payload = json.loads(response.read().decode("utf-8"))
            return self._coerce_planning_state(normalized, payload)
        return self._read_local_planning_state(normalized)

    def save_planning_state(
        self,
        dataset: str,
        overrides: Any,
        week_capacity_schedule: Any = None,
        base_version: Any = None,
        source: Any = "ui",
    ) -> Dict[str, Any]:
        normalized = self._validate_dataset(dataset)
        normalized_overrides = self._normalize_overrides(overrides)
        normalized_week_capacity_schedule = self._normalize_week_capacity_schedule(week_capacity_schedule)
        normalized_base_version = self._normalize_base_version(base_version)
        normalized_source = self._normalize_source(source)
        uploaded_at = _utc_now_iso()

        current_state = self.planning_state(normalized)
        current_version = int(current_state.get("version") or 0)
        if normalized_base_version is not None and normalized_base_version != current_version:
            raise PlanningStateConflictError(
                f"Planning state version conflict for dataset '{normalized}'. "
                f"Expected {normalized_base_version}, current is {current_version}."
            )

        next_state = {
            "dataset": normalized,
            "version": current_version + 1,
            "updatedAt": uploaded_at,
            "source": normalized_source,
            "overrideCount": len(normalized_overrides),
            "overrides": normalized_overrides,
            "weekCapacityScheduleCount": len(normalized_week_capacity_schedule),
            "weekCapacitySchedule": normalized_week_capacity_schedule,
        }

        if self.use_blob:
            timestamp = _timestamp_ms()
            state_path = f"capacity-dashboard/planning/{normalized}/{timestamp}.json"
            state_blob = self._blob_put(
                state_path,
                json.dumps(next_state, separators=(",", ":")).encode("utf-8"),
                PLANNING_STATE_MIME,
            )
            state_url = str(state_blob.get("url") or "")
            state_storage_path = str(state_blob.get("pathname") or state_path)
            manifest_payload = {
                "dataset": normalized,
                "version": next_state["version"],
                "updatedAt": next_state["updatedAt"],
                "source": normalized_source,
                "overrideCount": next_state["overrideCount"],
                "weekCapacityScheduleCount": next_state["weekCapacityScheduleCount"],
                "stateUrl": state_url,
                "statePath": state_storage_path,
            }
            manifest_path = f"capacity-dashboard/planning-manifests/{normalized}/{timestamp}.json"
            self._blob_put(
                manifest_path,
                json.dumps(manifest_payload, separators=(",", ":")).encode("utf-8"),
                PLANNING_STATE_MIME,
            )
            return next_state

        self._write_local_planning_state(normalized, next_state)
        return next_state

    def revenue_rates(self, dataset: str) -> Dict[str, Any]:
        normalized = self._validate_dataset(dataset)
        if self.use_blob:
            manifest = self._blob_latest_json_manifest(f"capacity-dashboard/revenue-manifests/{normalized}/")
            if not manifest:
                return self._default_revenue_rates_state(normalized)
            state_url = str(manifest.get("stateUrl") or "")
            if not state_url:
                return self._default_revenue_rates_state(normalized)
            with urllib.request.urlopen(state_url, timeout=20) as response:
                payload = json.loads(response.read().decode("utf-8"))
            return self._coerce_revenue_rates_state(normalized, payload)
        return self._read_local_revenue_rates_state(normalized)

    def save_revenue_rates(
        self,
        dataset: str,
        rates: Any,
        base_version: Any = None,
        source: Any = "ui",
    ) -> Dict[str, Any]:
        normalized = self._validate_dataset(dataset)
        normalized_rates = self._normalize_revenue_rates(rates)
        normalized_base_version = self._normalize_base_version(base_version)
        normalized_source = self._normalize_source(source)
        updated_at = _utc_now_iso()

        current_state = self.revenue_rates(normalized)
        current_version = int(current_state.get("version") or 0)
        if normalized_base_version is not None and normalized_base_version != current_version:
            raise RevenueRatesConflictError(
                f"Revenue rates version conflict for dataset '{normalized}'. "
                f"Expected {normalized_base_version}, current is {current_version}."
            )

        next_state = {
            "dataset": normalized,
            "version": current_version + 1,
            "updatedAt": updated_at,
            "source": normalized_source,
            "rateCount": len(normalized_rates),
            "rates": normalized_rates,
        }

        if self.use_blob:
            timestamp = _timestamp_ms()
            state_path = f"capacity-dashboard/revenue-rates/{normalized}/{timestamp}.json"
            state_blob = self._blob_put(
                state_path,
                json.dumps(next_state, separators=(",", ":")).encode("utf-8"),
                PLANNING_STATE_MIME,
            )
            state_url = str(state_blob.get("url") or "")
            state_storage_path = str(state_blob.get("pathname") or state_path)
            manifest_payload = {
                "dataset": normalized,
                "version": next_state["version"],
                "updatedAt": next_state["updatedAt"],
                "source": normalized_source,
                "rateCount": next_state["rateCount"],
                "stateUrl": state_url,
                "statePath": state_storage_path,
            }
            manifest_path = f"capacity-dashboard/revenue-manifests/{normalized}/{timestamp}.json"
            self._blob_put(
                manifest_path,
                json.dumps(manifest_payload, separators=(",", ":")).encode("utf-8"),
                PLANNING_STATE_MIME,
            )
            return next_state

        self._write_local_revenue_rates_state(normalized, next_state)
        return next_state


store = WorkbookStore()
