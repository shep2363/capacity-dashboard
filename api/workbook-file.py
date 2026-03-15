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
@app.get("/api/workbook-file")
@require_auth(READ_ONLY_ROLES)
def workbook_file() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        workbook_bytes, workbook_name = store.workbook_content(dataset)
    except ValueError as exc:
        response = jsonify({"error": str(exc)})
        response.status_code = 400
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
    except FileNotFoundError:
        response = jsonify({"error": f"No workbook has been uploaded for dataset '{dataset}'."})
        response.status_code = 404
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response
    except Exception:
        response = jsonify({"error": "Failed loading workbook file."})
        response.status_code = 500
        response.headers["Cache-Control"] = "no-store, max-age=0"
        return response

    response = Response(
        workbook_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = f'inline; filename="{workbook_name}"'
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response
