from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _workbook_store import store

app = Flask(__name__)


@app.get("/")
@app.get("/api/workbook-file")
def workbook_file() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        workbook_bytes, workbook_name = store.workbook_content(dataset)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except FileNotFoundError:
        return jsonify({"error": f"No workbook has been uploaded for dataset '{dataset}'."}), 404
    except Exception:
        return jsonify({"error": "Failed loading workbook file."}), 500

    response = Response(
        workbook_bytes,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["Content-Disposition"] = f'inline; filename="{workbook_name}"'
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response
