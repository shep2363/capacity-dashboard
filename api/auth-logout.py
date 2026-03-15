from __future__ import annotations

import sys
from pathlib import Path

from flask import Flask, Response

sys.path.append(str(Path(__file__).resolve().parent))
from _auth import clear_session_cookie, json_no_store

app = Flask(__name__)


@app.post("/")
@app.post("/api/auth/logout")
def auth_logout() -> Response:
    response = json_no_store({"authenticated": False, "role": None})
    return clear_session_cookie(response)
