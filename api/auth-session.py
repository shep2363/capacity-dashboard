from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response

sys.path.append(str(Path(__file__).resolve().parent))
from _auth import auth_configuration_error, build_session_payload, current_session, json_no_store

app = Flask(__name__)


@app.get("/")
@app.get("/api/auth/session")
def auth_session() -> Response:
    config_error = auth_configuration_error()
    if config_error:
        return json_no_store({"error": config_error}, 503)
    return json_no_store(build_session_payload(current_session()))
