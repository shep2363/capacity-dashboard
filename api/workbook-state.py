from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _workbook_store import store

app = Flask(__name__)


@app.get("/")
@app.get("/api/workbook-state")
def workbook_state() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        return jsonify(store.workbook_state(dataset))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
