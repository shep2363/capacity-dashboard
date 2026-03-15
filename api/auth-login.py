from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response, request

sys.path.append(str(Path(__file__).resolve().parent))
from _auth import (
    AuthSession,
    auth_configuration_error,
    authenticate_password,
    build_session_payload,
    json_no_store,
    set_session_cookie,
)

app = Flask(__name__)


@app.post("/")
@app.post("/api/auth/login")
def auth_login() -> Response:
    config_error = auth_configuration_error()
    if config_error:
        return json_no_store({"error": config_error}, 503)

    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return json_no_store({"error": "Invalid JSON payload."}, 400)

    role = authenticate_password(payload.get("password"))
    if role is None:
        return json_no_store({"error": "Incorrect password. Please try again."}, 401)

    response = json_no_store(build_session_payload(AuthSession(role=role)))
    return set_session_cookie(response, role)
