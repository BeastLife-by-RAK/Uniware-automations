import os
import requests
from dotenv import load_dotenv, set_key

load_dotenv()

TENANT_URL    = os.getenv("UNIWARE_TENANT_URL")
USERNAME      = os.getenv("UNIWARE_USERNAME")
PASSWORD      = os.getenv("UNIWARE_PASSWORD")
CLIENT_ID     = os.getenv("UNIWARE_CLIENT_ID", "my-trusted-client")
FACILITY_CODE = os.getenv("UNIWARE_FACILITY_CODE", "")

_access_token  = os.getenv("UNIWARE_ACCESS_TOKEN", "")
_refresh_token = os.getenv("UNIWARE_REFRESH_TOKEN", "")

_ENV_PATH      = ".env"
_HAS_ENV_FILE  = os.path.isfile(_ENV_PATH)


def _save_tokens(access: str, refresh: str) -> None:
    global _access_token, _refresh_token
    _access_token  = access
    _refresh_token = refresh
    if _HAS_ENV_FILE:
        set_key(_ENV_PATH, "UNIWARE_ACCESS_TOKEN",  access)
        set_key(_ENV_PATH, "UNIWARE_REFRESH_TOKEN", refresh)


def fetch_token_with_password() -> str:
    resp = requests.get(
        f"{TENANT_URL}/oauth/token",
        params={
            "grant_type": "password",
            "client_id":  CLIENT_ID,
            "username":   USERNAME,
            "password":   PASSWORD,
        },
        timeout=30,
    )
    resp.raise_for_status()
    tokens = resp.json()
    _save_tokens(tokens["access_token"], tokens["refresh_token"])
    print("✓ Token fetched via username/password")
    return tokens["access_token"]


def refresh_access_token() -> str:
    resp = requests.get(
        f"{TENANT_URL}/oauth/token",
        params={
            "grant_type":    "refresh_token",
            "client_id":     CLIENT_ID,
            "refresh_token": _refresh_token,
        },
        timeout=30,
    )
    if resp.status_code == 400:
        print("Refresh token expired — falling back to password login...")
        return fetch_token_with_password()
    resp.raise_for_status()
    tokens = resp.json()
    _save_tokens(tokens["access_token"], tokens.get("refresh_token", _refresh_token))
    print("✓ Token refreshed")
    return tokens["access_token"]


def get_valid_token() -> str:
    if _refresh_token:
        return refresh_access_token()
    return fetch_token_with_password()


def _headers() -> dict:
    global _access_token
    if not _access_token:
        _access_token = get_valid_token()
    return {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {_access_token}",
    }


def api_post(url: str, payload: dict, facility: str = None) -> dict:
    global _access_token

    def _hdrs():
        h = {
            "Content-Type":  "application/json",
            "Authorization": f"Bearer {_access_token}",
        }
        if facility:
            h["Facility"] = facility
        return h

    resp = requests.post(url, json=payload, headers=_hdrs(), timeout=60)

    if resp.status_code == 401:
        print("Token rejected — refreshing...")
        _access_token = get_valid_token()
        resp = requests.post(url, json=payload, headers=_hdrs(), timeout=60)

    if not resp.ok:
        raise Exception(f"HTTP {resp.status_code} from Unicommerce: {resp.text[:300]}")

    return resp.json()


def api_get(url: str, params: dict = None) -> dict:
    global _access_token

    resp = requests.get(url, params=params, headers=_headers(), timeout=60)

    if resp.status_code == 401:
        print("Token rejected — refreshing...")
        _access_token = get_valid_token()
        resp = requests.get(url, params=params, headers=_headers(), timeout=60)

    if not resp.ok:
        raise Exception(f"HTTP {resp.status_code} from Unicommerce: {resp.text[:300]}")

    return resp.json()