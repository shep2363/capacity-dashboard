from __future__ import annotations

import json
import logging
import os
import re
import urllib.error
import urllib.request
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, Optional

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


def _compose_project_label(project: Optional[str], sir: Optional[str], quote: Optional[str], title: Optional[str], job_number: Optional[str]) -> Optional[str]:
    if project:
        return project
    combined = " - ".join(part for part in (sir, quote, title) if part)
    if combined:
        return combined
    return job_number


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

    percent_column = _pick_column(
        columns,
        ["percent complete", "% complete", "percent done", "percentage complete", "progress", "pct complete"],
    )
    sequence_column = _pick_column(
        columns,
        ["sequence", "sequence name", "task", "task name", "operation", "description", "item", "name"],
    )
    project_column = _pick_column(columns, ["project", "project name", "job", "job name"])
    resource_column = _pick_column(columns, ["department", "resource", "shop", "area", "trade", "crew", "discipline"])
    sir_column = _pick_column(columns, ["sir"])
    quote_column = _pick_column(columns, ["quote", "quote number", "quote #"])
    title_column = _pick_column(columns, ["title", "project title"])
    job_number_column = _pick_column(columns, ["job number", "job #", "job no"])

    if percent_column is None:
        logger.error(
            "Smartsheet percent complete column not found. sheet_id=%s available_columns=%s",
            sheet_id,
            [column.get("title") for column in columns],
        )
        return _json_no_store({"error": "Could not find a Percent Complete column in Smartsheet."}, 500)

    parsed_rows = []
    skipped_without_percent = 0

    for row in rows:
        cell_map = {cell.get("columnId"): cell for cell in row.get("cells", []) if isinstance(cell, dict)}
        percent = _parse_percent(_cell_value(cell_map.get(percent_column.get("id"))))
        if percent is None:
            skipped_without_percent += 1
            continue

        project = _to_text(_cell_value(cell_map.get(project_column.get("id")))) if project_column else None
        sequence = _to_text(_cell_value(cell_map.get(sequence_column.get("id")))) if sequence_column else None
        resource = _to_text(_cell_value(cell_map.get(resource_column.get("id")))) if resource_column else None
        sir = _to_text(_cell_value(cell_map.get(sir_column.get("id")))) if sir_column else None
        quote = _to_text(_cell_value(cell_map.get(quote_column.get("id")))) if quote_column else None
        title = _to_text(_cell_value(cell_map.get(title_column.get("id")))) if title_column else None
        job_number = _to_text(_cell_value(cell_map.get(job_number_column.get("id")))) if job_number_column else None

        project_label = _compose_project_label(project, sir, quote, title, job_number)
        if not any([project_label, sequence, resource]):
            logger.warning(
                "Skipping Smartsheet row without usable identifiers. sheet_id=%s row_id=%s row_number=%s",
                sheet_id,
                row.get("id"),
                row.get("rowNumber"),
            )
            continue

        parsed_rows.append(
            {
                "rowId": str(row.get("id", "")),
                "rowNumber": int(row.get("rowNumber") or 0),
                "percentComplete": percent,
                "project": project_label,
                "sequence": sequence,
                "resource": resource,
                "identifiers": {
                    "sir": sir,
                    "quote": quote,
                    "title": title,
                    "jobNumber": job_number,
                },
            }
        )

    logger.info(
        "Loaded Smartsheet progress. sheet_id=%s rows=%s parsed=%s skipped_without_percent=%s",
        sheet_id,
        len(rows),
        len(parsed_rows),
        skipped_without_percent,
    )

    return _json_no_store(
        {
            "sheetId": sheet_id,
            "updatedAt": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "rowCount": len(parsed_rows),
            "rows": parsed_rows,
            "matchedColumns": {
                "percentComplete": percent_column.get("title") if percent_column else None,
                "project": project_column.get("title") if project_column else None,
                "sequence": sequence_column.get("title") if sequence_column else None,
                "resource": resource_column.get("title") if resource_column else None,
                "sir": sir_column.get("title") if sir_column else None,
                "quote": quote_column.get("title") if quote_column else None,
                "title": title_column.get("title") if title_column else None,
                "jobNumber": job_number_column.get("title") if job_number_column else None,
            },
        }
    )
