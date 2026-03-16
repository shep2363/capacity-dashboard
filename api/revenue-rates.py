from __future__ import annotations

import logging
import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _workbook_store import RevenueRatesConflictError, store

app = Flask(__name__)
logger = logging.getLogger(__name__)


def _json_no_store(payload: dict, status: int = 200) -> Response:
    response = jsonify(payload)
    response.status_code = status
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response


@app.get("/")
@app.get("/api/revenue-rates")
def get_revenue_rates() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        payload = store.revenue_rates(dataset)
        return _json_no_store(payload)
    except ValueError as exc:
        return _json_no_store({"error": str(exc)}, 400)
    except Exception:
        logger.exception("Failed loading revenue rates. dataset=%s", dataset)
        return _json_no_store({"error": "Failed loading revenue rates."}, 500)


@app.post("/")
@app.post("/api/revenue-rates")
def save_revenue_rates() -> Response:
    dataset = request.args.get("dataset", "main")
    payload = request.get_json(silent=True)
    if payload is None:
        payload = {}
    if not isinstance(payload, dict):
        return _json_no_store({"error": "Invalid JSON payload."}, 400)

    rates = payload.get("rates", {})
    base_version = payload.get("baseVersion")
    source = payload.get("source", "ui")

    try:
        saved = store.save_revenue_rates(dataset, rates, base_version=base_version, source=source)
        return _json_no_store(saved)
    except RevenueRatesConflictError as exc:
        return _json_no_store({"error": str(exc)}, 409)
    except ValueError as exc:
        return _json_no_store({"error": str(exc)}, 400)
    except Exception:
        logger.exception("Failed saving revenue rates. dataset=%s", dataset)
        return _json_no_store({"error": "Failed saving revenue rates."}, 500)
