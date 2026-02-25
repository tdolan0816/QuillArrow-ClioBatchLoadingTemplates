#!/usr/bin/env python3
"""
Bulk download (and optional upload) Clio templates.

This tool supports two sources:
- document-templates: uses /document_templates.json (if available)
- documents-folder: uses /documents.json filtered by a templates folder

It writes a manifest JSON file that can be reused for uploads.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import requests


# -------------------------------
# Default configuration values
# -------------------------------
DEFAULT_BASE_URL = "https://app.clio.com/api/v4"
DEFAULT_AUTH_BASE = "https://app.clio.com"
DEFAULT_LIMIT = 200
DEFAULT_FOLDER_NAME = "Templates"
DEFAULT_MANIFEST = "clio_templates_manifest.json"
DEFAULT_OUTPUT_DIR = "Template_Download"
DEFAULT_TOKEN_FILE = "clio_tokens.json"
TOKEN_EXPIRY_SKEW_SECONDS = 60


def build_session(access_token: str, api_version: str | None) -> requests.Session:
    """Create a requests session preloaded with Clio auth headers."""
    session = requests.Session()
    session.headers.update(
        {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
        }
    )
    if api_version:
        session.headers["X-API-VERSION"] = api_version
    return session


def request_json(
    session: requests.Session,
    method: str,
    url: str,
    params: Optional[Dict[str, str]] = None,
    data: Optional[Dict[str, str]] = None,
    files: Optional[Dict] = None,
    max_retries: int = 3,
) -> requests.Response:
    """
    Make an HTTP request with retry handling for 429 and 5xx responses.

    - 429: respects Retry-After header.
    - 5xx: retries with exponential backoff.
    """
    for attempt in range(max_retries + 1):
        response = session.request(method, url, params=params, data=data, files=files)
        if response.status_code == 429:
            retry_after = int(response.headers.get("Retry-After", "5"))
            time.sleep(retry_after)
            continue
        if response.status_code >= 500 and attempt < max_retries:
            time.sleep(2**attempt)
            continue
        return response
    return response


def load_token_file(path: Path) -> Dict:
    """Load OAuth tokens from a JSON file on disk."""
    if not path.exists():
        raise FileNotFoundError(f"Token file not found: {path}")
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def write_token_file(path: Path, payload: Dict) -> None:
    """Persist OAuth tokens to disk as JSON."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)


def request_token(
    url: str, data: Dict[str, str], max_retries: int
) -> requests.Response:
    """POST to the OAuth token endpoint with basic retry handling."""
    for attempt in range(max_retries + 1):
        response = requests.post(url, data=data, timeout=30)
        if response.status_code == 429:
            retry_after = int(response.headers.get("Retry-After", "5"))
            time.sleep(retry_after)
            continue
        if response.status_code >= 500 and attempt < max_retries:
            time.sleep(2**attempt)
            continue
        return response
    return response


def refresh_access_token(
    auth_base: str,
    client_id: str,
    client_secret: str,
    refresh_token: str,
    max_retries: int,
) -> Dict:
    """Refresh an expired access token using the stored refresh token."""
    token_url = f"{auth_base.rstrip('/')}/oauth/token"
    response = request_token(
        token_url,
        {
            "grant_type": "refresh_token",
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
        },
        max_retries,
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"Token refresh failed ({response.status_code}): {response.text}"
        )
    payload = response.json()
    now = int(time.time())
    payload["created_at"] = now
    if "expires_in" in payload:
        payload["expires_at"] = now + int(payload["expires_in"])
    return payload


def token_expired(payload: Dict) -> bool:
    """Return True if the token is expired (with a small safety skew)."""
    expires_at = payload.get("expires_at")
    if not expires_at:
        return False
    try:
        return time.time() >= float(expires_at) - TOKEN_EXPIRY_SKEW_SECONDS
    except (TypeError, ValueError):
        return False


def resolve_access_token(args: argparse.Namespace) -> str:
    """
    Resolve an access token from (in order):
    1) CLI flag or environment variable.
    2) Token file, with optional refresh if expired.
    """
    access_token = args.access_token or os.environ.get("CLIO_ACCESS_TOKEN")
    if access_token:
        return access_token

    # Fall back to a token file produced by the OAuth helper.
    token_path = Path(
        args.token_file
        or os.environ.get("CLIO_TOKEN_FILE", DEFAULT_TOKEN_FILE)
    ).resolve()
    payload = load_token_file(token_path)

    if token_expired(payload):
        # Attempt a refresh if the access token is expired.
        refresh_token = payload.get("refresh_token")
        if not refresh_token:
            raise RuntimeError(
                "Access token expired and no refresh token found in token file."
            )
        client_id = os.environ.get("CLIO_CLIENT_ID")
        client_secret = os.environ.get("CLIO_CLIENT_SECRET")
        if not client_id or not client_secret:
            raise RuntimeError(
                "Access token expired. Set CLIO_CLIENT_ID and CLIO_CLIENT_SECRET "
                "to refresh automatically."
            )
        auth_base = args.auth_base or os.environ.get("CLIO_AUTH_BASE", DEFAULT_AUTH_BASE)
        refreshed = refresh_access_token(
            auth_base, client_id, client_secret, refresh_token, args.max_retries
        )
        # Store refreshed tokens back to the same file.
        payload.update(refreshed)
        write_token_file(token_path, payload)

    access_token = payload.get("access_token")
    if not access_token:
        raise RuntimeError("Token file does not include an access_token.")
    return access_token


def extract_next_page_token(meta: Dict) -> Optional[str]:
    """Return the next page token from a Clio paging metadata object."""
    paging = meta.get("paging", {}) if isinstance(meta, dict) else {}
    for key in ("next_page_token", "next_page", "next", "page_token"):
        token = paging.get(key)
        if token:
            return token
    return None


def find_payload_value(payload: object, keys: Tuple[str, ...]) -> Optional[object]:
    """Recursively search a JSON payload for the first matching key."""
    if isinstance(payload, dict):
        for key, value in payload.items():
            if key in keys:
                return value
            found = find_payload_value(value, keys)
            if found is not None:
                return found
    elif isinstance(payload, list):
        for item in payload:
            found = find_payload_value(item, keys)
            if found is not None:
                return found
    return None


def iter_pages(
    session: requests.Session,
    url: str,
    params: Dict[str, str],
    max_retries: int,
) -> Iterable[Dict]:
    """
    Yield items across all pages for a list endpoint.

    Uses Clio's paging token in the response metadata.
    """
    page_token: Optional[str] = None
    while True:
        page_params = dict(params)
        if page_token:
            page_params["page_token"] = page_token

        response = request_json(
            session, "GET", url, params=page_params, max_retries=max_retries
        )
        if response.status_code != 200:
            raise RuntimeError(
                f"Request failed ({response.status_code}): {response.text}"
            )
        payload = response.json()
        data = payload.get("data", [])
        if isinstance(data, list):
            for item in data:
                yield item

        page_token = extract_next_page_token(payload.get("meta", {}))
        if not page_token:
            break


def sanitize_filename(name: str) -> str:
    """Make a safe filename for Windows by removing invalid characters."""
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned or "template"


def unique_path(path: Path) -> Path:
    """Return a path that does not overwrite an existing file."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    counter = 1
    while True:
        candidate = path.with_name(f"{stem}__{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def try_list_document_templates(
    session: requests.Session,
    base_url: str,
    limit: int,
    max_retries: int,
) -> Tuple[Optional[List[Dict]], Optional[str]]:
    """
    Attempt to list templates using the document_templates endpoint.

    Returns (items, None) if supported, or (None, error_message) if not.
    """
    url = f"{base_url}/document_templates.json"
    response = request_json(
        session, "GET", url, params={"limit": str(limit)}, max_retries=max_retries
    )
    if response.status_code in (404, 403):
        return None, response.text
    if response.status_code != 200:
        raise RuntimeError(
            f"Document templates list failed ({response.status_code}): {response.text}"
        )
    items = list(iter_pages(session, url, {"limit": str(limit)}, max_retries))
    return items, None


def resolve_folder_id(
    session: requests.Session,
    base_url: str,
    folder_id: Optional[str],
    folder_name: str,
    limit: int,
    max_retries: int,
) -> str:
    """
    Resolve a folder ID by name when an explicit folder_id is not provided.

    Raises if the folder name is missing or ambiguous.
    """
    if folder_id:
        return folder_id

    url = f"{base_url}/folders.json"
    matches: List[Dict] = []
    for item in iter_pages(
        session,
        url,
        {"limit": str(limit)},
        max_retries,
    ):
        name = (item.get("name") or "").strip()
        if name.lower() == folder_name.lower():
            matches.append(item)

    if not matches:
        raise RuntimeError(
            f"No folders named '{folder_name}' were found. "
            "Provide --folder-id to use a specific folder."
        )
    if len(matches) > 1:
        raise RuntimeError(
            f"Multiple folders named '{folder_name}' found; "
            "provide --folder-id to disambiguate."
        )
    return str(matches[0]["id"])


def list_documents_in_folder(
    session: requests.Session,
    base_url: str,
    folder_id: str,
    limit: int,
    max_retries: int,
) -> List[Dict]:
    """List documents inside a single Clio folder."""
    url = f"{base_url}/documents.json"
    return list(
        iter_pages(
            session,
            url,
            {"limit": str(limit), "folder_id": str(folder_id)},
            max_retries,
        )
    )


def download_document(
    session: requests.Session,
    base_url: str,
    document_id: str,
    target_path: Path,
    max_retries: int,
) -> None:
    """Download a document file by ID to a local path."""
    url = f"{base_url}/documents/{document_id}/download"
    for attempt in range(max_retries + 1):
        response = session.get(url, allow_redirects=True, stream=True)
        if response.status_code == 429:
            retry_after = int(response.headers.get("Retry-After", "5"))
            time.sleep(retry_after)
            continue
        if response.status_code >= 500 and attempt < max_retries:
            time.sleep(2**attempt)
            continue
        if response.status_code != 200:
            raise RuntimeError(
                f"Download failed ({response.status_code}): {response.text}"
            )

        target_path.parent.mkdir(parents=True, exist_ok=True)
        with target_path.open("wb") as handle:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    handle.write(chunk)
        return


def download_document_template(
    session: requests.Session,
    base_url: str,
    template_id: str,
    target_path: Path,
    max_retries: int,
) -> None:
    """
    Download a document template by trying common template endpoints.

    This is separate from documents because some accounts expose templates
    via a dedicated endpoint.
    """
    candidate_paths = [
        f"{base_url}/document_templates/{template_id}/contents",
        f"{base_url}/document_templates/{template_id}/download",
    ]
    last_error: Optional[str] = None
    for url in candidate_paths:
        for attempt in range(max_retries + 1):
            response = session.get(url, allow_redirects=True, stream=True)
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", "5"))
                time.sleep(retry_after)
                continue
            if response.status_code >= 500 and attempt < max_retries:
                time.sleep(2**attempt)
                continue
            if response.status_code == 200:
                target_path.parent.mkdir(parents=True, exist_ok=True)
                with target_path.open("wb") as handle:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            handle.write(chunk)
                return
            last_error = f"{url} -> {response.status_code}"
            break
    raise RuntimeError(f"Document template download failed: {last_error}")


def write_manifest(path: Path, entries: List[Dict]) -> None:
    """Write a manifest JSON describing downloaded templates."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(entries, handle, indent=2)


def load_manifest(path: Path) -> List[Dict]:
    """Load a manifest JSON file created during download."""
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def upload_document_version(
    session: requests.Session,
    base_url: str,
    document_id: str,
    file_path: Path,
    max_retries: int,
    verbose: bool,
) -> None:
    """
    Upload a new version of a document using common endpoint variants.

    Some Clio accounts expect different field names; this tries a few.
    """
    candidate_paths = [
        f"{base_url}/documents/{document_id}/document_versions.json",
        f"{base_url}/documents/{document_id}/versions.json",
        f"{base_url}/document_versions.json",
    ]
    payload_options = [
        ("file", {"document_id": document_id}),
        ("file", {"document_version[document_id]": document_id}),
        ("document_version[file]", {"document_id": document_id}),
        ("document_version[file]", {"document_version[document_id]": document_id}),
    ]

    attempts: List[str] = []
    for url in candidate_paths:
        for file_field, data in payload_options:
            with file_path.open("rb") as handle:
                files = {file_field: (file_path.name, handle)}
                response = request_json(
                    session,
                    "POST",
                    url,
                    data=data,
                    files=files,
                    max_retries=max_retries,
                )
                if response.status_code in (200, 201, 204):
                    return
            attempts.append(
                f"POST {url} field={file_field} -> {response.status_code} "
                f"{response.text[:300]}"
            )
            if response.status_code not in (404, 422):
                break
    raise RuntimeError(
        "Upload failed. The document version endpoint may require different fields. "
        "Check Clio API documentation or contact api@clio.com."
    )


def upload_document_template(
    session: requests.Session,
    base_url: str,
    template_id: str,
    file_path: Path,
    max_retries: int,
    verbose: bool,
) -> None:
    """
    Upload updated contents for a document template.

    Clio's template upload endpoints are not documented in the public OpenAPI,
    so this attempts several common endpoint and field combinations.
    """
    candidate_requests = [
        ("PUT", f"{base_url}/document_templates/{template_id}/contents"),
        ("POST", f"{base_url}/document_templates/{template_id}/contents"),
        ("PATCH", f"{base_url}/document_templates/{template_id}.json"),
        ("POST", f"{base_url}/document_templates/{template_id}.json"),
    ]
    payload_options = [
        ("file", {}),
        ("document_template[file]", {}),
        ("document_template[file]", {"document_template[id]": template_id}),
        ("file", {"id": template_id}),
    ]

    attempts: List[str] = []
    for method, url in candidate_requests:
        for file_field, data in payload_options:
            with file_path.open("rb") as handle:
                files = {file_field: (file_path.name, handle)}
                response = request_json(
                    session,
                    method,
                    url,
                    data=data,
                    files=files,
                    max_retries=max_retries,
                )
                if response.status_code in (200, 201, 204):
                    upload_url = None
                    payload = None
                    if response.headers.get("Content-Type", "").startswith(
                        "application/json"
                    ):
                        try:
                            payload = response.json()
                        except ValueError:
                            payload = None
                        upload_url = find_payload_value(payload, ("upload_url", "put_url"))

                    # If the API provided a pre-signed upload URL, follow the 3-step flow.
                    if upload_url:
                        content_type = (
                            "application/vnd.openxmlformats-officedocument."
                            "wordprocessingml.document"
                        )
                        with file_path.open("rb") as upload_handle:
                            put_response = requests.put(
                                str(upload_url),
                                data=upload_handle,
                                headers={"Content-Type": content_type},
                                timeout=60,
                            )
                        if put_response.status_code in (200, 201, 204):
                            finalize_urls = [
                                f"{base_url}/document_templates/{template_id}.json",
                                f"{base_url}/document_templates/{template_id}/contents",
                            ]
                            finalize_payloads = [
                                {"document_template[fully_uploaded]": "true"},
                                {"fully_uploaded": "true"},
                            ]
                            for finalize_url in finalize_urls:
                                for finalize_data in finalize_payloads:
                                    finalize_response = request_json(
                                        session,
                                        "PATCH",
                                        finalize_url,
                                        data=finalize_data,
                                        max_retries=max_retries,
                                    )
                                    if finalize_response.status_code in (200, 204):
                                        return
                                    attempts.append(
                                        "PATCH "
                                        f"{finalize_url} -> {finalize_response.status_code} "
                                        f"{finalize_response.text[:300]}"
                                    )
                        attempts.append(
                            f"PUT {upload_url} -> {put_response.status_code} "
                            f"{put_response.text[:300]}"
                        )
                    else:
                        return
            attempts.append(
                f"{method} {url} field={file_field} -> {response.status_code} "
                f"{response.text[:300]}"
            )
            if response.status_code not in (404, 405, 415, 422):
                break
    raise RuntimeError(
        "Template upload failed. The template upload endpoint may be unavailable "
        "or require different fields. Contact api@clio.com for the correct route."
        + (f"\nAttempts:\n- " + "\n- ".join(attempts) if verbose else "")
    )


def parse_args() -> argparse.Namespace:
    """Define and parse CLI arguments for list/download/upload commands."""
    parser = argparse.ArgumentParser(
        description="Bulk download and upload Clio templates via API."
    )
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL)
    parser.add_argument(
        "--auth-base",
        default=os.environ.get("CLIO_AUTH_BASE", DEFAULT_AUTH_BASE),
        help="OAuth base URL (e.g., https://app.clio.com).",
    )
    parser.add_argument("--api-version", help="Optional X-API-VERSION header value.")
    parser.add_argument(
        "--access-token",
        help="OAuth access token (or set CLIO_ACCESS_TOKEN env var).",
    )
    parser.add_argument(
        "--token-file",
        help="Path to token file (or set CLIO_TOKEN_FILE env var).",
    )
    parser.add_argument("--limit", type=int, default=DEFAULT_LIMIT)
    parser.add_argument("--max-retries", type=int, default=3)
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Include detailed API attempt info in error messages.",
    )

    sub = parser.add_subparsers(dest="command", required=True)

    list_cmd = sub.add_parser("list", help="List templates available in Clio.")
    list_cmd.add_argument(
        "--source",
        choices=["auto", "document-templates", "documents-folder"],
        default="auto",
    )
    list_cmd.add_argument("--folder-id")
    list_cmd.add_argument("--folder-name", default=DEFAULT_FOLDER_NAME)

    download_cmd = sub.add_parser("download", help="Download templates to a folder.")
    download_cmd.add_argument(
        "--source",
        choices=["auto", "document-templates", "documents-folder"],
        default="auto",
    )
    download_cmd.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
    )
    download_cmd.add_argument(
        "--manifest",
        default=DEFAULT_MANIFEST,
    )
    download_cmd.add_argument("--folder-id")
    download_cmd.add_argument("--folder-name", default=DEFAULT_FOLDER_NAME)
    download_cmd.add_argument("--dry-run", action="store_true")

    upload_cmd = sub.add_parser("upload", help="Upload updated templates from manifest.")
    upload_cmd.add_argument(
        "--manifest",
        default=DEFAULT_MANIFEST,
    )
    upload_cmd.add_argument(
        "--upload-dir",
        help=(
            "Override manifest paths and upload files from this directory. "
            "Uses the manifest file_name (or file_path basename)."
        ),
    )
    upload_cmd.add_argument("--dry-run", action="store_true")

    return parser.parse_args()


def main() -> int:
    """Entry point: resolve auth, then run list/download/upload flows."""
    args = parse_args()
    try:
        access_token = resolve_access_token(args)
    except Exception as exc:
        print(str(exc))
        return 1

    # Build an HTTP session with auth headers.
    session = build_session(access_token, args.api_version)

    if args.command in ("list", "download"):
        # Determine which source to use for templates.
        items: List[Dict] = []
        source_used = args.source

        if args.source in ("auto", "document-templates"):
            templates, error = try_list_document_templates(
                session, args.base_url, args.limit, args.max_retries
            )
            if templates is not None:
                items = templates
                source_used = "document-templates"
            elif args.source == "document-templates":
                print(
                    "Document templates endpoint not available for this account. "
                    f"Server response: {error}"
                )
                return 1

        if not items and args.source in ("auto", "documents-folder"):
            # Fall back to documents stored in a "Templates" folder.
            folder_id = resolve_folder_id(
                session,
                args.base_url,
                args.folder_id,
                args.folder_name,
                args.limit,
                args.max_retries,
            )
            items = list_documents_in_folder(
                session, args.base_url, folder_id, args.limit, args.max_retries
            )
            source_used = "documents-folder"

        if args.command == "list":
            print(f"Source: {source_used}")
            print(f"Templates found: {len(items)}")
            return 0

        output_dir = Path(args.output_dir).resolve()
        manifest_path = Path(args.manifest).resolve()
        manifest_entries: List[Dict] = []

        for item in items:
            # Normalize a safe filename and plan an output path.
            template_id = str(item.get("id"))
            name = item.get("name") or item.get("filename") or f"template_{template_id}"
            filename = sanitize_filename(str(name))
            if not filename.lower().endswith(".docx"):
                filename = f"{filename}.docx"

            target_path = unique_path(output_dir / filename)
            manifest_entries.append(
                {
                    "id": template_id,
                    "name": name,
                    "source": source_used,
                    "file_name": filename,
                    "file_path": str(target_path),
                }
            )

            if args.dry_run:
                continue

            if source_used == "documents-folder":
                # Document download endpoint for regular files.
                download_document(
                    session,
                    args.base_url,
                    template_id,
                    target_path,
                    args.max_retries,
                )
            else:
                # Template download endpoint if supported by this account.
                download_document_template(
                    session,
                    args.base_url,
                    template_id,
                    target_path,
                    args.max_retries,
                )

        write_manifest(manifest_path, manifest_entries)
        print(f"Downloaded {len(items)} templates to {output_dir}")
        print(f"Manifest written to {manifest_path}")
        return 0

    if args.command == "upload":
        # Upload updated templates by reading the manifest file.
        manifest_path = Path(args.manifest).resolve()
        entries = load_manifest(manifest_path)
        for entry in entries:
            source = entry.get("source")
            template_id = entry.get("id")
            file_path = Path(entry.get("file_path", ""))
            if args.upload_dir:
                file_name = entry.get("file_name") or file_path.name
                if not file_name:
                    raise RuntimeError(
                        "Upload dir provided but manifest does not contain "
                        "file_name or file_path."
                    )
                file_path = Path(args.upload_dir).resolve() / file_name

            if args.dry_run:
                continue

            if source != "documents-folder":
                upload_document_template(
                    session,
                    args.base_url,
                    str(template_id),
                    file_path,
                    args.max_retries,
                    args.verbose,
                )
                continue

            upload_document_version(
                session,
                args.base_url,
                str(template_id),
                file_path,
                args.max_retries,
                args.verbose,
            )

        print(f"Uploaded {len(entries)} templates from manifest.")
        return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
