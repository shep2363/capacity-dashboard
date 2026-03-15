from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _auth import READ_ONLY_ROLES, require_auth
from _workbook_store import store

app = Flask(__name__)


@app.get("/")
@app.get("/api/workbook-state")
@require_auth(READ_ONLY_ROLES)
def workbook_state() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        response = jsonify(store.workbook_state(dataset))
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
    except ValueError as exc:
        response = jsonify({"error": str(exc)})
        response.status_code = 400
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
