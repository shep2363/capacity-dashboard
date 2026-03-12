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
        return jsonify(payload)
    except ValueError as exc:
        message = str(exc)
        if message.startswith("Workbook exceeds maximum size"):
            return jsonify({"error": message}), 413
        return jsonify({"error": message}), 400
    except Exception:
        return jsonify({"error": "Saving the workbook file failed."}), 500
