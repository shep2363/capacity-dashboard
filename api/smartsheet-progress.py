from __future__ import annotations

import json
import logging
import os
import re
import urllib.error
import urllib.request
from datetime import datetime, timezone
from typing import Any, Iterable, Optional

from flask import Flask, Response, jsonify

app = Flask(__name__)
logger = logging.getLogger(__name__)

DEFAULT_SMARTSHEET_SHEET_ID = "6903479281864580"
SMARTSHEET_API_BASE = "https://api.smartsheet.com/2.0"


def _json_no_store(payload: dict, status: int = 200) -> Response:
    response = jsonify(payload)
    response.status_code = status
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response


def _normalize_title(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _pick_column(columns: Iterable[dict], candidates: Iterable[str]) -> Optional[dict]:
    normalized_candidates = {_normalize_title(candidate) for candidate in candidates}
    for column in columns:
        if _normalize_title(column.get("title")) in normalized_candidates:
            return column
    return None


def _fetch_sheet(sheet_id: str, token: str) -> dict:
    request = urllib.request.Request(
        f"{SMARTSHEET_API_BASE}/sheets/{sheet_id}",
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def _cell_value(cell: dict) -> Any:
    if "displayValue" in cell and cell["displayValue"] not in (None, ""):
        return cell["displayValue"]
    if "value" in cell and cell["value"] not in (None, ""):
        return cell["value"]
    object_value = cell.get("objectValue")
    if isinstance(object_value, dict):
        for key in ("displayValue", "value", "name", "title"):
            if object_value.get(key) not in (None, ""):
                return object_value.get(key)
    return None


def _to_text(value: Any) -> Optional[str]:
    if value in (None, ""):
        return None
    text = str(value).strip()
    return text or None


def _parse_percent(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
        if 0 <= numeric <= 1:
            return round(numeric * 100, 2)
        return round(numeric, 2)
    text = str(value).strip()
    cleaned = re.sub(r"[^0-9.\-]+", "", text)
    if not cleaned:
        return None
    try:
        numeric = float(cleaned)
    except ValueError:
        return None
    if "%" not in text and 0 <= numeric <= 1:
        numeric *= 100
    return round(numeric, 2)


def _clamp_percent(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return max(0.0, min(100.0, round(value, 2)))


@app.get("/")
@app.get("/api/smartsheet-progress")
def get_smartsheet_progress() -> Response:
    token = os.getenv("SMARTSHEET_API_TOKEN", "").strip()
    if not token:
        return _json_no_store({"error": "SMARTSHEET_API_TOKEN is not configured."}, 500)

    sheet_id = os.getenv("SMARTSHEET_SHEET_ID", DEFAULT_SMARTSHEET_SHEET_ID).strip() or DEFAULT_SMARTSHEET_SHEET_ID

    try:
        sheet = _fetch_sheet(sheet_id, token)
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="ignore")
        logger.exception("Smartsheet request failed. sheet_id=%s status=%s body=%s", sheet_id, exc.code, details)
        return _json_no_store({"error": f"Failed loading Smartsheet sheet ({exc.code})."}, 502)
    except Exception:
        logger.exception("Unexpected Smartsheet request failure. sheet_id=%s", sheet_id)
        return _json_no_store({"error": "Failed loading Smartsheet sheet."}, 500)

    columns = sheet.get("columns", []) if isinstance(sheet, dict) else []
    rows = sheet.get("rows", []) if isinstance(sheet, dict) else []

    job_column = _pick_column(columns, ["job", "job number", "job #", "job no"])
    sequence_column = _pick_column(columns, ["sequence", "sequence name", "task", "task name", "operation", "item", "name"])
    process_column = _pick_column(columns, ["process"])
    manual_process_column = _pick_column(columns, ["manual process", "manualprocess"])
    weld_column = _pick_column(columns, ["weld"])
    paint_column = _pick_column(columns, ["paint"])
    qc_column = _pick_column(columns, ["qc", "quality control"])
    ship_column = _pick_column(columns, ["ship", "shipping"])
    weight_column = _pick_column(columns, ["weight", "weight t", "weight (t)"])
    source_sheet_column = _pick_column(columns, ["source sheet", "sourcesheet"])

    if job_column is None or sequence_column is None:
        logger.error(
            "Smartsheet Job/Sequence columns not found. sheet_id=%s available_columns=%s",
            sheet_id,
            [column.get("title") for column in columns],
        )
        return _json_no_store({"error": "Could not find required Job and Sequence columns in Smartsheet."}, 500)

    parsed_rows = []
    skipped_without_identifiers = 0
    skipped_without_station_progress = 0

    for row in rows:
        cell_map = {cell.get("columnId"): cell for cell in row.get("cells", []) if isinstance(cell, dict)}
        job = _to_text(_cell_value(cell_map.get(job_column.get("id")))) if job_column else None
        sequence = _to_text(_cell_value(cell_map.get(sequence_column.get("id")))) if sequence_column else None

        if not job or not sequence:
            skipped_without_identifiers += 1
            continue

        station_progress = {
            "process": _clamp_percent(_parse_percent(_cell_value(cell_map.get(process_column.get("id"))))) if process_column else None,
            "manualProcess": _clamp_percent(_parse_percent(_cell_value(cell_map.get(manual_process_column.get("id"))))) if manual_process_column else None,
            "weld": _clamp_percent(_parse_percent(_cell_value(cell_map.get(weld_column.get("id"))))) if weld_column else None,
            "paint": _clamp_percent(_parse_percent(_cell_value(cell_map.get(paint_column.get("id"))))) if paint_column else None,
            "qc": _clamp_percent(_parse_percent(_cell_value(cell_map.get(qc_column.get("id"))))) if qc_column else None,
            "ship": _clamp_percent(_parse_percent(_cell_value(cell_map.get(ship_column.get("id"))))) if ship_column else None,
        }
        if not any(value is not None for value in station_progress.values()):
            skipped_without_station_progress += 1
            continue

        parsed_rows.append(
            {
                "rowId": str(row.get("id", "")),
                "rowNumber": int(row.get("rowNumber") or 0),
                "job": job,
                "sequence": sequence,
                "weight": _to_text(_cell_value(cell_map.get(weight_column.get("id")))) if weight_column else None,
                "sourceSheet": _to_text(_cell_value(cell_map.get(source_sheet_column.get("id")))) if source_sheet_column else None,
                "stationProgress": station_progress,
            }
        )

    logger.info(
        "Loaded Smartsheet progress. sheet_id=%s rows=%s parsed=%s skipped_without_identifiers=%s skipped_without_station_progress=%s matched_columns=%s",
        sheet_id,
        len(rows),
        len(parsed_rows),
        skipped_without_identifiers,
        skipped_without_station_progress,
        {
            "job": job_column.get("title") if job_column else None,
            "sequence": sequence_column.get("title") if sequence_column else None,
            "process": process_column.get("title") if process_column else None,
            "manualProcess": manual_process_column.get("title") if manual_process_column else None,
            "weld": weld_column.get("title") if weld_column else None,
            "paint": paint_column.get("title") if paint_column else None,
            "qc": qc_column.get("title") if qc_column else None,
            "ship": ship_column.get("title") if ship_column else None,
            "sourceSheet": source_sheet_column.get("title") if source_sheet_column else None,
        },
    )

    return _json_no_store(
        {
            "sheetId": sheet_id,
            "updatedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "rowCount": len(parsed_rows),
            "rows": parsed_rows,
            "matchedColumns": {
                "job": job_column.get("title") if job_column else None,
                "sequence": sequence_column.get("title") if sequence_column else None,
                "process": process_column.get("title") if process_column else None,
                "manualProcess": manual_process_column.get("title") if manual_process_column else None,
                "weld": weld_column.get("title") if weld_column else None,
                "paint": paint_column.get("title") if paint_column else None,
                "qc": qc_column.get("title") if qc_column else None,
                "ship": ship_column.get("title") if ship_column else None,
                "weight": weight_column.get("title") if weight_column else None,
                "sourceSheet": source_sheet_column.get("title") if source_sheet_column else None,
            },
        }
    )
