from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response, jsonify, send_file

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _workbook_store import store

app = Flask(__name__)


@app.get("/")
@app.get("/api/workbook-file")
def workbook_file() -> Response:
    if not store.active_workbook_path.exists():
        return jsonify({"error": "No active workbook has been uploaded yet."}), 404

    state = store.workbook_state()
    response = send_file(
        store.active_workbook_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=False,
        download_name=state.get("fileName") or "active_workbook.xlsx",
        max_age=0,
    )
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response
