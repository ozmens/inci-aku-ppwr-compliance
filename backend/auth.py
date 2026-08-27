"""Built-in auth — no Supabase required.

JWT in HTTP-only cookie (+ optional Bearer). Users in workspace/auth/users.json.
Seed admin from INCI_PPWR_ADMIN_USER / INCI_PPWR_ADMIN_PASSWORD.
"""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from fastapi import HTTPException, Request, Response
from pydantic import BaseModel, Field

ROOT = Path(__file__).resolve().parents[1]
WS_ROOT = Path(os.environ.get("INCI_PPWR_WORKSPACE_ROOT", str(ROOT / "workspace"))).resolve()
AUTH_DIR = WS_ROOT / "auth"
USERS_FILE = AUTH_DIR / "users.json"

COOKIE_NAME = "inci_ppwr_session"
TOKEN_TTL_HOURS = int(os.environ.get("INCI_PPWR_TOKEN_HOURS", "72"))
AUTH_DISABLED = os.environ.get("INCI_PPWR_AUTH", "1").strip().lower() in {"0", "false", "off", "no"}

# Secret: env first, else stable file under auth dir
_SECRET_ENV = os.environ.get("INCI_PPWR_JWT_SECRET", "").strip()


def _secret() -> bytes:
    if _SECRET_ENV:
        return _SECRET_ENV.encode("utf-8")
    AUTH_DIR.mkdir(parents=True, exist_ok=True)
    secret_file = AUTH_DIR / ".jwt_secret"
    if secret_file.exists():
        return secret_file.read_bytes().strip()
    val = secrets.token_hex(32).encode("utf-8")
    secret_file.write_bytes(val)
    try:
        secret_file.chmod(0o600)
    except OSError:
        pass
    return val


def _b64url(data: bytes) -> str:
    import base64

    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


def _b64url_decode(s: str) -> bytes:
    import base64

    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s + pad)


def _hash_password(password: str, salt: str | None = None) -> tuple[str, str]:
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120_000,
    )
    return salt, digest.hex()


def _verify_password(password: str, salt: str, password_hash: str) -> bool:
    _, dig = _hash_password(password, salt)
    return hmac.compare_digest(dig, password_hash)


def _ensure_users() -> list[dict[str, Any]]:
    AUTH_DIR.mkdir(parents=True, exist_ok=True)
    if USERS_FILE.exists():
        data = json.loads(USERS_FILE.read_text(encoding="utf-8"))
        users = data.get("users") or []
        if users:
            return users

    admin_user = os.environ.get("INCI_PPWR_ADMIN_USER", "admin").strip() or "admin"
    admin_pass = os.environ.get("INCI_PPWR_ADMIN_PASSWORD", "160616").strip() or "160616"
    salt, pw_hash = _hash_password(admin_pass)
    users = [
        {
            "id": "usr_admin",
            "username": admin_user,
            "display_name": "Administrator",
            "role": "admin",
            "salt": salt,
            "password_hash": pw_hash,
            "active": True,
            "created_at": datetime.now(timezone.utc).isoformat(),
        }
    ]
    USERS_FILE.write_text(
        json.dumps({"users": users, "updated_at": datetime.now(timezone.utc).isoformat()}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return users


def _save_users(users: list[dict[str, Any]]) -> None:
    AUTH_DIR.mkdir(parents=True, exist_ok=True)
    USERS_FILE.write_text(
        json.dumps({"users": users, "updated_at": datetime.now(timezone.utc).isoformat()}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _find_user(username: str) -> dict[str, Any] | None:
    u = username.strip().lower()
    for row in _ensure_users():
        if str(row.get("username", "")).lower() == u:
            return row
    return None


def mint_token(user: dict[str, Any]) -> str:
    """Compact HMAC token: header.payload.sig (no external JWT lib)."""
    now = datetime.now(timezone.utc)
    payload = {
        "sub": user.get("id"),
        "username": user.get("username"),
        "role": user.get("role", "user"),
        "iat": int(now.timestamp()),
        "exp": int((now + timedelta(hours=TOKEN_TTL_HOURS)).timestamp()),
    }
    body = _b64url(json.dumps(payload, separators=(",", ":")).encode("utf-8"))
    head = _b64url(b'{"alg":"HS256","typ":"JWT"}')
    msg = f"{head}.{body}".encode("ascii")
    sig = _b64url(hmac.new(_secret(), msg, hashlib.sha256).digest())
    return f"{head}.{body}.{sig}"


def parse_token(token: str) -> dict[str, Any]:
    try:
        head, body, sig = token.split(".")
    except ValueError as e:
        raise HTTPException(401, "Geçersiz oturum") from e
    msg = f"{head}.{body}".encode("ascii")
    expect = _b64url(hmac.new(_secret(), msg, hashlib.sha256).digest())
    if not hmac.compare_digest(expect, sig):
        raise HTTPException(401, "Geçersiz oturum imzası")
    payload = json.loads(_b64url_decode(body))
    if int(payload.get("exp", 0)) < int(datetime.now(timezone.utc).timestamp()):
        raise HTTPException(401, "Oturum süresi doldu")
    return payload


class LoginRequest(BaseModel):
    username: str = Field(min_length=1)
    password: str = Field(min_length=1)


class CreateUserRequest(BaseModel):
    username: str = Field(min_length=2)
    password: str = Field(min_length=6)
    display_name: str = ""
    role: str = "user"


def login(username: str, password: str) -> dict[str, Any]:
    user = _find_user(username)
    if not user or not user.get("active", True):
        raise HTTPException(401, "Kullanıcı adı veya şifre hatalı")
    if not _verify_password(password, user["salt"], user["password_hash"]):
        raise HTTPException(401, "Kullanıcı adı veya şifre hatalı")
    token = mint_token(user)
    return {
        "token": token,
        "user": {
            "id": user["id"],
            "username": user["username"],
            "display_name": user.get("display_name") or user["username"],
            "role": user.get("role", "user"),
        },
        "expires_hours": TOKEN_TTL_HOURS,
    }


def set_session_cookie(response: Response, token: str) -> None:
    secure = bool(os.environ.get("RENDER")) or os.environ.get("INCI_PPWR_COOKIE_SECURE", "").lower() in {
        "1",
        "true",
    }
    response.set_cookie(
        key=COOKIE_NAME,
        value=token,
        httponly=True,
        samesite="lax",
        secure=secure,
        max_age=TOKEN_TTL_HOURS * 3600,
        path="/",
    )


def clear_session_cookie(response: Response) -> None:
    response.delete_cookie(COOKIE_NAME, path="/")


def extract_token(request: Request) -> str | None:
    auth = request.headers.get("Authorization") or ""
    if auth.lower().startswith("bearer "):
        return auth[7:].strip()
    q = request.query_params.get("access_token")
    if q:
        return q.strip()
    return request.cookies.get(COOKIE_NAME)


def require_user(request: Request) -> dict[str, Any]:
    if AUTH_DISABLED:
        return {"id": "dev", "username": "dev", "role": "admin"}
    token = extract_token(request)
    if not token:
        raise HTTPException(401, "Giriş gerekli")
    return parse_token(token)


def require_admin(request: Request) -> dict[str, Any]:
    user = require_user(request)
    if user.get("role") != "admin" and not AUTH_DISABLED:
        raise HTTPException(403, "Yalnızca yönetici")
    return user


def list_users_public() -> list[dict[str, Any]]:
    out = []
    for u in _ensure_users():
        out.append(
            {
                "id": u.get("id"),
                "username": u.get("username"),
                "display_name": u.get("display_name") or u.get("username"),
                "role": u.get("role", "user"),
                "active": bool(u.get("active", True)),
            }
        )
    return out


def public_user_from_payload(payload: dict[str, Any]) -> dict[str, Any]:
    uid = payload.get("sub") or payload.get("id")
    for u in _ensure_users():
        if u.get("id") == uid:
            return {
                "id": u.get("id"),
                "username": u.get("username"),
                "display_name": u.get("display_name") or u.get("username"),
                "role": u.get("role", "user"),
                "active": bool(u.get("active", True)),
            }
    return {
        "id": uid,
        "username": payload.get("username"),
        "display_name": payload.get("username"),
        "role": payload.get("role", "user"),
        "active": True,
    }


def create_user(req: CreateUserRequest) -> dict[str, Any]:
    users = _ensure_users()
    if any(str(u.get("username", "")).lower() == req.username.strip().lower() for u in users):
        raise HTTPException(400, "Bu kullanıcı adı zaten var")
    salt, pw_hash = _hash_password(req.password)
    row = {
        "id": f"usr_{secrets.token_hex(4)}",
        "username": req.username.strip(),
        "display_name": (req.display_name or req.username).strip(),
        "role": req.role if req.role in {"admin", "user"} else "user",
        "salt": salt,
        "password_hash": pw_hash,
        "active": True,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    users.append(row)
    _save_users(users)
    return {
        "id": row["id"],
        "username": row["username"],
        "display_name": row["display_name"],
        "role": row["role"],
        "active": True,
    }


class ResetPasswordRequest(BaseModel):
    password: str = Field(min_length=6)


class ChangePasswordRequest(BaseModel):
    current_password: str = Field(min_length=1)
    new_password: str = Field(min_length=6)


def change_own_password(user_id: str, current_password: str, new_password: str) -> dict[str, Any]:
    users = _ensure_users()
    for u in users:
        if u.get("id") == user_id:
            if not _verify_password(current_password, u["salt"], u["password_hash"]):
                raise HTTPException(400, "Mevcut şifre hatalı")
            salt, pw_hash = _hash_password(new_password)
            u["salt"] = salt
            u["password_hash"] = pw_hash
            _save_users(users)
            return {"ok": True, "username": u["username"]}
    raise HTTPException(404, "Kullanıcı bulunamadı")


def reset_password(user_id: str, password: str) -> dict[str, Any]:
    users = _ensure_users()
    for u in users:
        if u.get("id") == user_id:
            salt, pw_hash = _hash_password(password)
            u["salt"] = salt
            u["password_hash"] = pw_hash
            _save_users(users)
            return {"id": u["id"], "username": u["username"], "ok": True}
    raise HTTPException(404, "Kullanıcı bulunamadı")


def set_user_active(user_id: str, active: bool) -> dict[str, Any]:
    users = _ensure_users()
    for u in users:
        if u.get("id") == user_id:
            if u.get("username") == "admin" and not active:
                raise HTTPException(400, "admin hesabı kapatılamaz")
            u["active"] = bool(active)
            _save_users(users)
            return {
                "id": u["id"],
                "username": u["username"],
                "display_name": u.get("display_name") or u["username"],
                "role": u.get("role", "user"),
                "active": bool(u["active"]),
            }
    raise HTTPException(404, "Kullanıcı bulunamadı")


def auth_status() -> dict[str, Any]:
    _ensure_users()
    return {
        "auth_required": not AUTH_DISABLED,
        "auth_disabled": AUTH_DISABLED,
        "cookie": COOKIE_NAME,
        "users": len(list_users_public()),
    }


PUBLIC_API_PREFIXES = (
    "/api/health",
    "/api/auth/login",
    "/api/auth/status",
    "/docs",
    "/redoc",
    "/openapi.json",
)


def is_public_path(path: str) -> bool:
    if AUTH_DISABLED:
        return True
    if path == "/" or not path.startswith("/api"):
        return True
    if path.startswith("/assets"):
        return True
    for p in PUBLIC_API_PREFIXES:
        if path == p or path.startswith(p + "?"):
            return True
    return False
