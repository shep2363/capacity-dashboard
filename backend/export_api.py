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

DEFAULT_MAX_UPLOAD_BYTES = 30 * 1024 * 1024
MAX_UPLOAD_BYTES = int(os.getenv("CAPACITY_MAX_UPLOAD_BYTES", str(DEFAULT_MAX_UPLOAD_BYTES)))
DATA_ROOT = Path(os.getenv("CAPACITY_SHARED_DATA_DIR", Path(__file__).resolve().parent / "shared_store")).resolve()
DATASETS = {"main", "sales"}
MANIFEST_DIR = DATA_ROOT / "manifests"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _default_manifest(dataset: str) -> Dict[str, Any]:
    return {"version": 1, "dataset": dataset, "fileName": "", "uploadedAt": None, "sizeBytes": 0}


def _validate_dataset(dataset: str) -> str:
    normalized = (dataset or "").strip().lower()
    if normalized not in DATASETS:
        raise ValueError("Invalid dataset. Allowed values: main, sales.")
    return normalized


def _active_workbook_path(dataset: str) -> Path:
    return DATA_ROOT / f"active_{dataset}.xlsx"


def _manifest_path(dataset: str) -> Path:
    return MANIFEST_DIR / f"{dataset}.json"


def _write_json_atomic(path: Path, payload: Dict[str, Any]) -> None:
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


def _read_manifest(dataset: str) -> Dict[str, Any]:
    path = _manifest_path(dataset)
    if not path.exists():
        return _default_manifest(dataset)
    try:
        with path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
            if isinstance(payload, dict):
                return payload
    except Exception:
        app.logger.exception("Failed reading manifest.")
    return _default_manifest(dataset)


def _ensure_store() -> None:
    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    MANIFEST_DIR.mkdir(parents=True, exist_ok=True)
    for dataset in DATASETS:
        path = _manifest_path(dataset)
        if not path.exists():
            _write_json_atomic(path, _default_manifest(dataset))


def _workbook_state_response(dataset: str) -> Dict[str, Any]:
    normalized = _validate_dataset(dataset)
    workbook_path = _active_workbook_path(normalized)
    manifest = _read_manifest(normalized)
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
    dataset = request.args.get("dataset", "main")
    try:
        return jsonify(_workbook_state_response(dataset))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400


@app.get("/api/workbook-file")
def workbook_file() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        normalized = _validate_dataset(dataset)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    workbook_path = _active_workbook_path(normalized)
    if not workbook_path.exists():
        return jsonify({"error": f"No workbook has been uploaded for dataset '{normalized}'."}), 404

    manifest = _read_manifest(normalized)
    download_name = str(manifest.get("fileName") or workbook_path.name)
    return send_file(
        workbook_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=False,
        download_name=download_name,
        max_age=0,
    )


@app.post("/api/upload-workbook")
def upload_workbook() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        normalized = _validate_dataset(dataset)
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
        with tempfile.NamedTemporaryFile(delete=False, dir=str(DATA_ROOT), prefix="active_", suffix=".tmp") as temp_file:
            upload.save(temp_file)
            temp_path = Path(temp_file.name)

        if temp_path is None or not temp_path.exists():
            return jsonify({"error": "Saving the workbook file failed."}), 500

        file_size = temp_path.stat().st_size
        if file_size <= 0:
            temp_path.unlink(missing_ok=True)
            return jsonify({"error": "Uploaded workbook is empty."}), 400

        if file_size > MAX_UPLOAD_BYTES:
            temp_path.unlink(missing_ok=True)
            return jsonify({"error": f"Workbook exceeds maximum size of {MAX_UPLOAD_BYTES} bytes."}), 413

        with temp_path.open("rb") as handle:
            if handle.read(2) != b"PK":
                temp_path.unlink(missing_ok=True)
                return jsonify({"error": "Invalid workbook payload. Expected .xlsx content."}), 400

        os.replace(temp_path, _active_workbook_path(normalized))
        uploaded_at = _utc_now_iso()

        manifest = {
            "version": 1,
            "dataset": normalized,
            "fileName": original_name,
            "uploadedAt": uploaded_at,
            "sizeBytes": file_size,
        }
        _write_json_atomic(_manifest_path(normalized), manifest)

        return jsonify(
            {
                "dataset": normalized,
                "hasWorkbook": True,
                "fileName": original_name,
                "uploadedAt": uploaded_at,
                "sizeBytes": file_size,
            }
        )
    except Exception:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
        app.logger.exception("Failed uploading active workbook.")
        return jsonify({"error": "Saving the workbook file failed."}), 500


@app.get("/api/shared-health")
def shared_health() -> Response:
    main_state = _workbook_state_response("main")
    sales_state = _workbook_state_response("sales")
    return jsonify(
        {
            "status": "ok",
            "dataRoot": str(DATA_ROOT),
            "datasets": {
                "main": main_state,
                "sales": sales_state,
            },
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
    app.run(host="0.0.0.0", port=8000, debug=True)
