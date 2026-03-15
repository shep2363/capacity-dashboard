from __future__ import annotations

import os
from dataclasses import dataclass
from functools import wraps
from typing import Any, Callable, Dict, Iterable, Optional, TypeVar, cast

import bcrypt
from flask import Response, jsonify, request
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

AuthRole = str
READ_ONLY_ROLES = frozenset({"admin", "user"})
ADMIN_ONLY_ROLES = frozenset({"admin"})

SESSION_COOKIE_NAME = "capacity_dashboard_session"
SESSION_SALT = "capacity-dashboard-auth-v1"
DEFAULT_SESSION_MAX_AGE_SECONDS = 12 * 60 * 60
DEFAULT_ALLOWED_ORIGINS = (
    "http://localhost:5173,http://127.0.0.1:5173,http://localhost:4173,http://127.0.0.1:4173"
)


@dataclass(frozen=True)
class AuthSession:
    role: AuthRole


ResponseHandler = TypeVar("ResponseHandler", bound=Callable[..., Response])


def json_no_store(payload: Dict[str, Any], status: int = 200) -> Response:
    response = jsonify(payload)
    response.status_code = status
    response.headers["Cache-Control"] = "no-store, max-age=0"
    return response


def allowed_frontend_origins() -> list[str]:
    raw = os.getenv("CAPACITY_ALLOWED_ORIGINS", DEFAULT_ALLOWED_ORIGINS)
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


def _cookie_name() -> str:
    return os.getenv("CAPACITY_SESSION_COOKIE_NAME", SESSION_COOKIE_NAME).strip() or SESSION_COOKIE_NAME


def _session_secret() -> str:
    return os.getenv("CAPACITY_SESSION_SECRET", "").strip()


def _serializer() -> URLSafeTimedSerializer | None:
    secret = _session_secret()
    if not secret:
        return None
    return URLSafeTimedSerializer(secret, salt=SESSION_SALT)


def _session_max_age_seconds() -> int:
    raw = os.getenv("CAPACITY_SESSION_MAX_AGE_SECONDS", str(DEFAULT_SESSION_MAX_AGE_SECONDS)).strip()
    try:
        return max(300, int(raw))
    except (TypeError, ValueError):
        return DEFAULT_SESSION_MAX_AGE_SECONDS


def _cookie_secure() -> bool:
    override = os.getenv("CAPACITY_SESSION_COOKIE_SECURE", "auto").strip().lower()
    if override in {"1", "true", "yes", "on"}:
        return True
    if override in {"0", "false", "no", "off"}:
        return False
    return bool(request.is_secure)


def _cookie_samesite() -> str:
    configured = os.getenv("CAPACITY_SESSION_COOKIE_SAMESITE", "Lax").strip().lower()
    if configured in {"none", "strict", "lax"}:
        return configured.capitalize()
    return "Lax"


def _password_hashes_by_role() -> Dict[AuthRole, bytes]:
    hashes: Dict[AuthRole, bytes] = {}
    admin_hash = os.getenv("CAPACITY_ADMIN_PASSWORD_HASH", "").strip()
    user_hash = os.getenv("CAPACITY_USER_PASSWORD_HASH", "").strip()
    if admin_hash:
        hashes["admin"] = admin_hash.encode("utf-8")
    if user_hash:
        hashes["user"] = user_hash.encode("utf-8")
    return hashes


def auth_configuration_error() -> str | None:
    if not _session_secret():
        return "Server auth is not configured. Missing CAPACITY_SESSION_SECRET."
    hashes = _password_hashes_by_role()
    if "admin" not in hashes and "user" not in hashes:
        return "Server auth is not configured. Set CAPACITY_ADMIN_PASSWORD_HASH and/or CAPACITY_USER_PASSWORD_HASH."
    return None


def _make_signed_session(role: AuthRole) -> str:
    serializer = _serializer()
    if serializer is None:
        raise RuntimeError("Missing session serializer.")
    return serializer.dumps({"role": role})


def current_session() -> AuthSession | None:
    serializer = _serializer()
    if serializer is None:
        return None

    cookie = request.cookies.get(_cookie_name(), "").strip()
    if not cookie:
        return None

    try:
        payload = serializer.loads(cookie, max_age=_session_max_age_seconds())
    except (BadSignature, SignatureExpired):
        return None

    if not isinstance(payload, dict):
        return None
    role = payload.get("role")
    if role not in READ_ONLY_ROLES:
        return None
    return AuthSession(role=cast(AuthRole, role))


def set_session_cookie(response: Response, role: AuthRole) -> Response:
    token = _make_signed_session(role)
    response.set_cookie(
        _cookie_name(),
        token,
        httponly=True,
        secure=_cookie_secure(),
        samesite=_cookie_samesite(),
        path="/",
    )
    return response


def clear_session_cookie(response: Response) -> Response:
    response.set_cookie(
        _cookie_name(),
        "",
        httponly=True,
        secure=_cookie_secure(),
        samesite=_cookie_samesite(),
        path="/",
        max_age=0,
        expires=0,
    )
    return response


def authenticate_password(password: Any) -> AuthRole | None:
    if not isinstance(password, str):
        return None
    candidate = password.encode("utf-8")
    if not candidate:
        return None

    for role, hashed in _password_hashes_by_role().items():
        try:
            if bcrypt.checkpw(candidate, hashed):
                return role
        except ValueError:
            continue
    return None


def build_session_payload(session: AuthSession | None) -> Dict[str, Any]:
    return {"authenticated": session is not None, "role": session.role if session else None}


def require_auth(roles: Optional[Iterable[AuthRole]] = None) -> Callable[[ResponseHandler], ResponseHandler]:
    allowed_roles = set(roles or READ_ONLY_ROLES)

    def decorator(func: ResponseHandler) -> ResponseHandler:
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> Response:
            config_error = auth_configuration_error()
            if config_error:
                return json_no_store({"error": config_error}, 503)

            session = current_session()
            if session is None:
                return json_no_store({"error": "Authentication required."}, 401)
            if session.role not in allowed_roles:
                return json_no_store({"error": "Forbidden."}, 403)
            return func(*args, **kwargs)

        return cast(ResponseHandler, wrapper)

    return decorator
