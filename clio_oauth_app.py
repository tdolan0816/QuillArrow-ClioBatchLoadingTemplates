#!/usr/bin/env python3
"""
Minimal OAuth helper for Clio.

Endpoints:
- GET /login: starts OAuth flow
- GET /oauth/callback: exchanges code for tokens and stores them to a file
- POST /clio/deauth: deauthorization callback (optional)
"""

from __future__ import annotations

import json
import os
import secrets
import time
from pathlib import Path
from urllib.parse import urlencode

import requests
from flask import Flask, redirect, request


DEFAULT_AUTH_BASE = "https://app.clio.com"
DEFAULT_TOKEN_FILE = "clio_tokens.json"

app = Flask(__name__)


def require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def get_auth_base() -> str:
    return os.environ.get("CLIO_AUTH_BASE", DEFAULT_AUTH_BASE).rstrip("/")


def get_token_file() -> Path:
    token_file = os.environ.get("CLIO_TOKEN_FILE", DEFAULT_TOKEN_FILE)
    return Path(token_file).resolve()


def save_tokens(payload: dict) -> Path:
    now = int(time.time())
    payload["created_at"] = now
    if "expires_in" in payload:
        payload["expires_at"] = now + int(payload["expires_in"])

    token_path = get_token_file()
    token_path.parent.mkdir(parents=True, exist_ok=True)
    with token_path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)
    return token_path


@app.get("/")
def index() -> str:
    return "Clio OAuth helper is running. Visit /login to authorize."


@app.get("/login")
def login():
    client_id = require_env("CLIO_CLIENT_ID")
    redirect_uri = require_env("CLIO_REDIRECT_URI")
    auth_base = get_auth_base()

    state = secrets.token_urlsafe(16)
    app.config["OAUTH_STATE"] = state

    params = {
        "response_type": "code",
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "state": state,
    }
    scope = os.environ.get("CLIO_SCOPE")
    if scope:
        params["scope"] = scope

    auth_url = f"{auth_base}/oauth/authorize?{urlencode(params)}"
    return redirect(auth_url)


@app.get("/oauth/callback")
def oauth_callback():
    if request.args.get("error"):
        return f"OAuth error: {request.args.get('error')}", 400

    code = request.args.get("code")
    if not code:
        return "Missing authorization code.", 400

    expected_state = app.config.get("OAUTH_STATE")
    state = request.args.get("state")
    if expected_state and state != expected_state:
        return "Invalid OAuth state.", 400

    client_id = require_env("CLIO_CLIENT_ID")
    client_secret = require_env("CLIO_CLIENT_SECRET")
    redirect_uri = require_env("CLIO_REDIRECT_URI")
    auth_base = get_auth_base()

    token_url = f"{auth_base}/oauth/token"
    response = requests.post(
        token_url,
        data={
            "grant_type": "authorization_code",
            "client_id": client_id,
            "client_secret": client_secret,
            "redirect_uri": redirect_uri,
            "code": code,
        },
        timeout=30,
    )

    if response.status_code != 200:
        return f"Token request failed ({response.status_code}): {response.text}", 400

    payload = response.json()
    token_path = save_tokens(payload)

    return (
        "OAuth complete. Tokens saved to "
        f"{token_path}. You can now run clio_templates_sync.py."
    )


@app.post("/clio/deauth")
def deauthorize():
    return "", 204


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8787")))
