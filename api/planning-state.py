from __future__ import annotations

import logging
import sys
from pathlib import Path

from flask import Flask, Response, jsonify, request

# Ensure sibling helper modules inside /api are importable in serverless runtime.
sys.path.append(str(Path(__file__).resolve().parent))
from _workbook_store import PlanningStateConflictError, store

app = Flask(__name__)
logger = logging.getLogger(__name__)


def _json_no_store(payload: dict, status: int = 200) -> Response:
    response = jsonify(payload)
    response.status_code = status
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response


@app.get("/")
@app.get("/api/planning-state")
def get_planning_state() -> Response:
    dataset = request.args.get("dataset", "main")
    try:
        payload = store.planning_state(dataset)
        return _json_no_store(payload)
    except ValueError as exc:
        return _json_no_store({"error": str(exc)}, 400)
    except Exception:
        logger.exception("Failed loading planning state. dataset=%s", dataset)
        return _json_no_store({"error": "Failed loading planning state."}, 500)


@app.post("/")
@app.post("/api/planning-state")
def save_planning_state() -> Response:
    dataset = request.args.get("dataset", "main")
    payload = request.get_json(silent=True)
    if payload is None:
        payload = {}
    if not isinstance(payload, dict):
        return _json_no_store({"error": "Invalid JSON payload."}, 400)

    overrides = payload.get("overrides", {})
    week_capacity_overrides = payload.get("weekCapacityOverrides", {})
    base_version = payload.get("baseVersion")
    source = payload.get("source", "ui")

    try:
        saved = store.save_planning_state(
            dataset,
            overrides,
            week_capacity_overrides=week_capacity_overrides,
            base_version=base_version,
            source=source,
        )
        return _json_no_store(saved)
    except PlanningStateConflictError as exc:
        return _json_no_store({"error": str(exc)}, 409)
    except ValueError as exc:
        return _json_no_store({"error": str(exc)}, 400)
    except Exception:
        logger.exception("Failed saving planning state. dataset=%s", dataset)
        return _json_no_store({"error": "Failed saving planning state."}, 500)
