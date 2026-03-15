from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _workbook_store import store

app = Flask(__name__)


@app.post("/")
@app.post("/api/upload-workbook")
def upload_workbook() -> Response:
    dataset = request.args.get("dataset", "main")
    uploaded_file = request.files.get("file")
    try:
        payload = store.save_upload(dataset, uploaded_file)
        response = jsonify(payload)
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
    except ValueError as exc:
        message = str(exc)
        if message.startswith("Workbook exceeds maximum size"):
            response = jsonify({"error": message})
            response.status_code = 413
            response.headers["Cache-Control"] = "no-store, max-age=0"
            return response
        response = jsonify({"error": message})
        response.status_code = 400
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
    except Exception:
        response = jsonify({"error": "Saving the workbook file failed."})
        response.status_code = 500
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
