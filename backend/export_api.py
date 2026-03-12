from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List

from flask import Flask, Response, jsonify, request, send_file
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": "*"}})

DATASETS = {"main", "sales"}
DEFAULT_MAX_UPLOAD_BYTES = 30 * 1024 * 1024
MAX_UPLOAD_BYTES = int(os.getenv("CAPACITY_MAX_UPLOAD_BYTES", str(DEFAULT_MAX_UPLOAD_BYTES)))
DATA_ROOT = Path(os.getenv("CAPACITY_SHARED_DATA_DIR", Path(__file__).resolve().parent / "shared_store")).resolve()
WORKBOOK_DIR = DATA_ROOT / "workbooks"
MANIFEST_FILE = DATA_ROOT / "manifest.json"
SHARED_STATE_FILE = DATA_ROOT / "shared_state.json"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _default_manifest() -> Dict[str, Any]:
    return {"version": 1, "datasets": {"main": {}, "sales": {}}}


def _default_shared_state() -> Dict[str, Any]:
    return {
        "version": 1,
        "main": {
            "file": "",
            "overrides": {},
            "enabled": {},
            "weeklyCaps": {},
            "filters": {"dateFrom": "", "dateTo": "", "year": "", "resources": []},
            "weekendDates": [],
            "weekendExtras": {},
        },
        "sales": {
            "file": "",
            "overrides": {},
            "enabled": {},
        },
    }


def _ensure_store() -> None:
    WORKBOOK_DIR.mkdir(parents=True, exist_ok=True)
    if not MANIFEST_FILE.exists():
        _write_json_atomic(MANIFEST_FILE, _default_manifest())
    if not SHARED_STATE_FILE.exists():
        _write_json_atomic(SHARED_STATE_FILE, _default_shared_state())


def _write_json_atomic(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=f"{path.name}.", suffix=".tmp", dir=str(path.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=True, separators=(",", ":"))
        os.replace(tmp_path, path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise


def _read_json(path: Path, default_payload: Dict[str, Any]) -> Dict[str, Any]:
    if not path.exists():
        return default_payload
    try:
        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
            if isinstance(payload, dict):
                return payload
    except Exception:
        app.logger.exception("Failed reading JSON payload from %s", path)
    return default_payload


def _validate_dataset(dataset: str) -> str:
    normalized = (dataset or "").strip().lower()
    if normalized not in DATASETS:
        raise ValueError("Invalid dataset. Allowed values are: main, sales.")
    return normalized


def _sanitize_number_map(raw: Any) -> Dict[str, float]:
    if not isinstance(raw, dict):
        return {}
    cleaned: Dict[str, float] = {}
    for key, value in raw.items():
        key_str = str(key).strip()
        if not key_str:
            continue
        try:
            number = float(value)
        except (TypeError, ValueError):
            continue
        if number != number or number < 0:
            continue
        cleaned[key_str] = number
    return cleaned


def _sanitize_boolean_map(raw: Any) -> Dict[str, bool]:
    if not isinstance(raw, dict):
        return {}
    return {str(key).strip(): bool(value) for key, value in raw.items() if str(key).strip()}


def _sanitize_filters(raw: Any) -> Dict[str, Any]:
    default_filters = {"dateFrom": "", "dateTo": "", "year": "", "resources": []}
    if not isinstance(raw, dict):
        return default_filters
    resources_raw = raw.get("resources")
    resources = [str(item) for item in resources_raw if isinstance(item, str)] if isinstance(resources_raw, list) else []
    return {
        "dateFrom": str(raw.get("dateFrom") or ""),
        "dateTo": str(raw.get("dateTo") or ""),
        "year": str(raw.get("year") or ""),
        "resources": resources,
    }


def _sanitize_shared_state(raw: Dict[str, Any]) -> Dict[str, Any]:
    default_state = _default_shared_state()
    main_raw = raw.get("main") if isinstance(raw.get("main"), dict) else {}
    sales_raw = raw.get("sales") if isinstance(raw.get("sales"), dict) else {}

    cleaned_state = {
        "version": 1,
        "main": {
            "file": str(main_raw.get("file") or ""),
            "overrides": _sanitize_number_map(main_raw.get("overrides")),
            "enabled": _sanitize_boolean_map(main_raw.get("enabled")),
            "weeklyCaps": _sanitize_number_map(main_raw.get("weeklyCaps")),
            "filters": _sanitize_filters(main_raw.get("filters")),
            "weekendDates": [
                str(item)
                for item in main_raw.get("weekendDates", [])
                if isinstance(main_raw.get("weekendDates"), list) and isinstance(item, str)
            ],
            "weekendExtras": _sanitize_number_map(main_raw.get("weekendExtras")),
        },
        "sales": {
            "file": str(sales_raw.get("file") or ""),
            "overrides": _sanitize_number_map(sales_raw.get("overrides")),
            "enabled": _sanitize_boolean_map(sales_raw.get("enabled")),
        },
        "updatedAt": _utc_now_iso(),
    }

    # Keep compatibility if a client sends empty structure.
    return {**default_state, **cleaned_state}


def _autosize_columns(ws) -> None:
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_len = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_len:
                max_len = len(value)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 52)


def _style_header(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"


def _append_rows(ws, header: List[str], rows: List[List[Any]]) -> None:
    ws.append(header)
    for row in rows:
        ws.append(row)
    _style_header(ws)
    _autosize_columns(ws)


def _build_workbook(payload: Dict[str, Any]) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    chart_rows = payload.get("weeklyCapacityChartRows", [])
    weekly_rows = payload.get("weeklyForecastRows", [])
    monthly_rows = payload.get("monthlyForecastRows", [])
    summary_rows = payload.get("summaryRows", [])
    chart_category_keys = payload.get("chartCategoryKeys", [])
    if not isinstance(chart_category_keys, list):
        chart_category_keys = []
    chart_category_keys = [str(key) for key in chart_category_keys]
    chart_category_count = len(chart_category_keys)

    ws_chart = wb.create_sheet("Weekly Capacity Chart")
    _append_rows(
        ws_chart,
        [
            "Week Start",
            "Week End",
            "Week Label",
            "Total Forecast Hours",
            "Capacity",
            "Variance",
            "Status",
            *chart_category_keys,
        ],
        chart_rows,
    )

    if len(chart_rows) > 0 and chart_category_count > 0:
        min_row = 1
        max_row = len(chart_rows) + 1
        cat_start_col = 8
        cat_end_col = cat_start_col + chart_category_count - 1

        bar = BarChart()
        bar.type = "col"
        bar.grouping = "stacked"
        bar.overlap = 100
        bar.title = "Weekly Capacity Forecast"
        bar.y_axis.title = "Hours"
        bar.x_axis.title = "Week"

        data = Reference(ws_chart, min_col=cat_start_col, max_col=cat_end_col, min_row=min_row, max_row=max_row)
        categories = Reference(ws_chart, min_col=3, min_row=2, max_row=max_row)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(categories)

        line = LineChart()
        line.y_axis.title = "Capacity"
        line_data = Reference(ws_chart, min_col=5, max_col=5, min_row=min_row, max_row=max_row)
        line.add_data(line_data, titles_from_data=True)
        line.set_categories(categories)
        line.y_axis.axId = 200
        line.x_axis = bar.x_axis

        bar += line
        bar.width = 28
        bar.height = 12
        ws_chart.add_chart(bar, "A3")

    ws_weekly = wb.create_sheet("Weekly Forecast")
    _append_rows(
        ws_weekly,
        ["Week Start", "Week End", "Week Label", "Forecast Hours", "Capacity", "Variance", "Status"],
        weekly_rows,
    )

    ws_monthly = wb.create_sheet("Monthly Forecast")
    _append_rows(
        ws_monthly,
        ["Month", "Planned Hours", "Capacity", "Variance", "Status"],
        monthly_rows,
    )

    ws_summary = wb.create_sheet("Summary")
    _append_rows(ws_summary, ["Metric", "Value"], summary_rows)

    return wb


@app.get("/api/workbook-state")
def workbook_state() -> Response:
    dataset_param = request.args.get("dataset", "main")
    try:
        dataset = _validate_dataset(dataset_param)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    manifest = _read_json(MANIFEST_FILE, _default_manifest())
    dataset_entry = manifest.get("datasets", {}).get(dataset, {})
    stored_path = WORKBOOK_DIR / f"{dataset}.xlsx"
    has_workbook = stored_path.exists()
    file_name = str(dataset_entry.get("fileName") or (stored_path.name if has_workbook else ""))
    uploaded_at = dataset_entry.get("uploadedAt")
    size_bytes = int(dataset_entry.get("sizeBytes") or (stored_path.stat().st_size if has_workbook else 0))

    return jsonify(
        {
            "dataset": dataset,
            "hasWorkbook": has_workbook,
            "fileName": file_name,
            "uploadedAt": uploaded_at,
            "sizeBytes": size_bytes,
        }
    )


@app.get("/api/workbook-file")
def workbook_file() -> Response:
    dataset_param = request.args.get("dataset", "main")
    try:
        dataset = _validate_dataset(dataset_param)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    stored_path = WORKBOOK_DIR / f"{dataset}.xlsx"
    if not stored_path.exists():
        return jsonify({"error": f"No workbook stored for dataset '{dataset}'."}), 404

    manifest = _read_json(MANIFEST_FILE, _default_manifest())
    dataset_entry = manifest.get("datasets", {}).get(dataset, {})
    download_name = str(dataset_entry.get("fileName") or stored_path.name)

    return send_file(
        stored_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=False,
        download_name=download_name,
        max_age=0,
    )


@app.post("/api/upload-workbook")
def upload_workbook() -> Response:
    dataset_param = request.args.get("dataset", "main")
    try:
        dataset = _validate_dataset(dataset_param)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "Missing upload file. Use form field 'file' with a .xlsx workbook."}), 400

    original_name = Path(upload.filename).name
    if not original_name.lower().endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are allowed."}), 400

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, dir=str(WORKBOOK_DIR), prefix=f"{dataset}_", suffix=".tmp") as temp:
            upload.save(temp)
            temp_path = Path(temp.name)

        if temp_path is None or not temp_path.exists():
            return jsonify({"error": "Saving the workbook file failed."}), 500

        size_bytes = temp_path.stat().st_size
        if size_bytes <= 0:
            temp_path.unlink(missing_ok=True)
            return jsonify({"error": "Uploaded workbook is empty."}), 400
        if size_bytes > MAX_UPLOAD_BYTES:
            temp_path.unlink(missing_ok=True)
            return jsonify({"error": f"Workbook exceeds maximum size of {MAX_UPLOAD_BYTES} bytes."}), 413

        with temp_path.open("rb") as handle:
            signature = handle.read(2)
            if signature != b"PK":
                temp_path.unlink(missing_ok=True)
                return jsonify({"error": "Invalid workbook payload. Expected .xlsx file content."}), 400

        destination = WORKBOOK_DIR / f"{dataset}.xlsx"
        os.replace(temp_path, destination)

        manifest = _read_json(MANIFEST_FILE, _default_manifest())
        datasets = manifest.setdefault("datasets", {})
        datasets[dataset] = {
            "fileName": original_name,
            "uploadedAt": _utc_now_iso(),
            "sizeBytes": size_bytes,
        }
        _write_json_atomic(MANIFEST_FILE, manifest)

        return jsonify(
            {
                "dataset": dataset,
                "hasWorkbook": True,
                "fileName": original_name,
                "uploadedAt": datasets[dataset]["uploadedAt"],
                "sizeBytes": size_bytes,
            }
        )
    except Exception:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
        app.logger.exception("Failed saving uploaded workbook for dataset '%s'", dataset)
        return jsonify({"error": "Saving the workbook file failed."}), 500


@app.get("/api/shared-state")
def get_shared_state() -> Response:
    state = _read_json(SHARED_STATE_FILE, _default_shared_state())
    return jsonify(state)


@app.put("/api/shared-state")
def put_shared_state() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "Invalid JSON payload."}), 400

    cleaned_state = _sanitize_shared_state(payload)
    _write_json_atomic(SHARED_STATE_FILE, cleaned_state)
    return jsonify(cleaned_state)


@app.get("/api/shared-health")
def shared_health() -> Response:
    main_path = WORKBOOK_DIR / "main.xlsx"
    sales_path = WORKBOOK_DIR / "sales.xlsx"
    return jsonify(
        {
            "status": "ok",
            "dataRoot": str(DATA_ROOT),
            "mainWorkbookPresent": main_path.exists(),
            "salesWorkbookPresent": sales_path.exists(),
        }
    )


@app.post("/api/export-report")
def export_report() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "Invalid JSON payload"}), 400

    file_name = str(payload.get("fileName") or "capacity-report.xlsx")
    if not file_name.endswith(".xlsx"):
        file_name = f"{file_name}.xlsx"

    workbook = _build_workbook(payload)
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{file_name}"'},
    )


_ensure_store()

if __name__ == "__main__":
    host = os.getenv("CAPACITY_API_HOST", "0.0.0.0")
    port = int(os.getenv("CAPACITY_API_PORT", "8000"))
    debug = os.getenv("CAPACITY_API_DEBUG", "false").strip().lower() == "true"
    app.run(host=host, port=port, debug=debug)
