#!/usr/bin/env python3
"""
SharePoint Online version cleanup worker + HTTP API.

Design intent:
- Power Automate triggers this service by HTTP.
- Python handles recursive traversal and version cleanup.
- Supports sync and background execution.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import threading
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Dict, Iterator, List, Optional, Tuple
from urllib.parse import urlparse

import msal
import requests
import uvicorn
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field


UTC = timezone.utc
DEFAULT_API_KEY_HEADER = "X-API-Key"


def utc_now() -> datetime:
    return datetime.now(UTC)


def parse_sp_datetime(value: str) -> datetime:
    raw = value.strip()
    if raw.endswith("Z"):
        raw = raw[:-1] + "+00:00"
    try:
        return datetime.fromisoformat(raw).astimezone(UTC)
    except ValueError:
        for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"):
            try:
                return datetime.strptime(raw, fmt).astimezone(UTC)
            except ValueError:
                continue
    raise ValueError(f"Unsupported SharePoint datetime format: {value}")


def sp_literal(server_relative_url: str) -> str:
    return "'" + server_relative_url.replace("'", "''") + "'"


def parse_bool(value: str, default: bool) -> bool:
    if value is None:
        return default
    return value.strip().lower() in ("1", "true", "yes", "on")


@dataclass
class Config:
    tenant_id: str
    client_id: str
    thumbprint: str
    private_key_path: str
    site_url: str
    root_folder_server_relative_url: str
    days_to_keep: int = 30
    dry_run: bool = True
    skip_folder_names: Tuple[str, ...] = ("Forms",)
    request_timeout_seconds: int = 60
    sleep_seconds_between_requests: float = 0.0


class SharePointVersionCleaner:
    def __init__(self, config: Config) -> None:
        self.config = config
        self.site_origin = self._site_origin(config.site_url)
        self.token = self._acquire_token()
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {self.token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
            }
        )
        self.stats: Dict[str, int] = {
            "folders_seen": 0,
            "files_seen": 0,
            "versions_seen": 0,
            "versions_selected": 0,
            "versions_deleted": 0,
            "delete_failures": 0,
        }

    @staticmethod
    def _site_origin(site_url: str) -> str:
        p = urlparse(site_url)
        return f"{p.scheme}://{p.netloc}"

    def _acquire_token(self) -> str:
        authority = f"https://login.microsoftonline.com/{self.config.tenant_id}"
        with open(self.config.private_key_path, "r", encoding="utf-8") as f:
            private_key = f.read()

        app = msal.ConfidentialClientApplication(
            client_id=self.config.client_id,
            authority=authority,
            client_credential={
                "private_key": private_key,
                "thumbprint": self.config.thumbprint,
            },
        )
        scope = [f"{self.site_origin}/.default"]
        result = app.acquire_token_for_client(scopes=scope)
        if "access_token" not in result:
            raise RuntimeError(
                "Failed to acquire token.\n"
                f"Error: {result.get('error')}\n"
                f"Description: {result.get('error_description')}\n"
                f"Correlation ID: {result.get('correlation_id')}"
            )
        return result["access_token"]

    def _request(self, method: str, api_path: str, *, expected=(200,), **kwargs) -> requests.Response:
        url = f"{self.config.site_url.rstrip('/')}/{api_path.lstrip('/')}"
        timeout = kwargs.pop("timeout", self.config.request_timeout_seconds)
        resp = self.session.request(method, url, timeout=timeout, **kwargs)
        if resp.status_code not in expected:
            body = resp.text[:2000]
            raise RuntimeError(f"{method} {url} failed: {resp.status_code}\n{body}")
        if self.config.sleep_seconds_between_requests > 0:
            time.sleep(self.config.sleep_seconds_between_requests)
        return resp

    def list_subfolders(self, folder_server_relative_url: str) -> List[Dict]:
        api_path = (
            "_api/web/getfolderbyserverrelativeurl("
            f"{sp_literal(folder_server_relative_url)}"
            ")/folders"
        )
        resp = self._request("GET", api_path)
        results = resp.json()["d"]["results"]
        return [item for item in results if item.get("Name") not in self.config.skip_folder_names]

    def list_files(self, folder_server_relative_url: str) -> List[Dict]:
        api_path = (
            "_api/web/getfolderbyserverrelativeurl("
            f"{sp_literal(folder_server_relative_url)}"
            ")/files"
        )
        resp = self._request("GET", api_path)
        return resp.json()["d"]["results"]

    def list_versions(self, file_server_relative_url: str) -> List[Dict]:
        api_path = (
            "_api/web/getfilebyserverrelativeurl("
            f"{sp_literal(file_server_relative_url)}"
            ")/versions"
        )
        resp = self._request("GET", api_path)
        return resp.json()["d"]["results"]

    def delete_version(self, file_server_relative_url: str, version_id: int) -> None:
        api_path = (
            "_api/web/getfilebyserverrelativeurl("
            f"{sp_literal(file_server_relative_url)}"
            f")/versions({version_id})"
        )
        self._request("DELETE", api_path, expected=(200, 204))

    def walk_folders(self, root_server_relative_url: str) -> Iterator[str]:
        stack = [root_server_relative_url]
        while stack:
            current = stack.pop()
            self.stats["folders_seen"] += 1
            yield current
            try:
                subfolders = self.list_subfolders(current)
            except Exception:
                logging.exception("Failed to list subfolders: %s", current)
                continue
            for folder in reversed(subfolders):
                sr = folder.get("ServerRelativeUrl")
                if sr:
                    stack.append(sr)

    def eligible_versions(self, versions: List[Dict], *, cutoff: datetime) -> List[Dict]:
        normalized: List[Tuple[datetime, Dict]] = []
        for version in versions:
            created_raw = version.get("Created")
            if not created_raw:
                continue
            try:
                created_at = parse_sp_datetime(created_raw)
            except Exception:
                logging.warning("Could not parse version timestamp: %s", created_raw)
                continue
            normalized.append((created_at, version))

        normalized.sort(key=lambda x: x[0], reverse=True)
        selected: List[Dict] = []
        for idx, (created_at, version) in enumerate(normalized):
            if idx == 0:
                continue
            if created_at < cutoff:
                selected.append(version)
        return selected

    def process_file(self, file_server_relative_url: str, *, cutoff: datetime) -> None:
        self.stats["files_seen"] += 1
        versions = self.list_versions(file_server_relative_url)
        self.stats["versions_seen"] += len(versions)
        candidates = self.eligible_versions(versions, cutoff=cutoff)
        self.stats["versions_selected"] += len(candidates)

        if not candidates:
            return

        for version in candidates:
            version_id = version.get("ID")
            if version_id is None:
                logging.warning("Version without ID skipped on %s: %s", file_server_relative_url, version)
                continue

            if self.config.dry_run:
                logging.info(
                    "[DRY RUN] Would delete version ID=%s for %s",
                    version_id,
                    file_server_relative_url,
                )
                continue

            try:
                self.delete_version(file_server_relative_url, int(version_id))
                self.stats["versions_deleted"] += 1
            except Exception:
                self.stats["delete_failures"] += 1
                logging.exception("Failed to delete version ID=%s from %s", version_id, file_server_relative_url)

    def run(self) -> Dict[str, int]:
        cutoff = utc_now() - timedelta(days=self.config.days_to_keep)
        logging.info("Cutoff date (UTC): %s", cutoff.isoformat())
        logging.info("Root folder: %s", self.config.root_folder_server_relative_url)
        logging.info("Dry run: %s", self.config.dry_run)

        for folder_sr in self.walk_folders(self.config.root_folder_server_relative_url):
            try:
                files = self.list_files(folder_sr)
            except Exception:
                logging.exception("Failed to list files in folder: %s", folder_sr)
                continue

            for file_info in files:
                file_sr = file_info.get("ServerRelativeUrl")
                if not file_sr:
                    continue
                try:
                    self.process_file(file_sr, cutoff=cutoff)
                except Exception:
                    logging.exception("Failed to process file: %s", file_sr)

        return self.stats


class CleanupRequest(BaseModel):
    root_folder_server_relative_url: Optional[str] = Field(default=None)
    days_to_keep: Optional[int] = Field(default=None, ge=0)
    dry_run: Optional[bool] = Field(default=None)


class CleanupResponse(BaseModel):
    status: str
    stats: Dict[str, int]
    root_folder_server_relative_url: str
    days_to_keep: int
    dry_run: bool


class JobResponse(BaseModel):
    job_id: str
    status: str


class JobDetailResponse(BaseModel):
    job_id: str
    status: str
    requested_at_utc: str
    finished_at_utc: Optional[str] = None
    request: CleanupRequest
    result: Optional[CleanupResponse] = None
    error: Optional[str] = None


class JobStore:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._jobs: Dict[str, Dict] = {}

    def create(self, req: CleanupRequest) -> str:
        job_id = str(uuid.uuid4())
        with self._lock:
            self._jobs[job_id] = {
                "job_id": job_id,
                "status": "queued",
                "requested_at_utc": utc_now().isoformat(),
                "finished_at_utc": None,
                "request": req.model_dump(),
                "result": None,
                "error": None,
            }
        return job_id

    def update(self, job_id: str, **changes: object) -> None:
        with self._lock:
            if job_id not in self._jobs:
                return
            self._jobs[job_id].update(changes)

    def get(self, job_id: str) -> Optional[Dict]:
        with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return None
            return dict(job)


def load_base_config_from_env() -> Config:
    required = [
        "SP_TENANT_ID",
        "SP_CLIENT_ID",
        "SP_CERT_THUMBPRINT",
        "SP_PRIVATE_KEY_PATH",
        "SP_SITE_URL",
        "SP_ROOT_FOLDER_SERVER_RELATIVE_URL",
    ]
    missing = [name for name in required if not os.getenv(name)]
    if missing:
        raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

    skip_raw = os.getenv("SP_SKIP_FOLDER_NAMES", "Forms")
    skip_folders = tuple([x.strip() for x in skip_raw.split(",") if x.strip()])

    return Config(
        tenant_id=os.environ["SP_TENANT_ID"],
        client_id=os.environ["SP_CLIENT_ID"],
        thumbprint=os.environ["SP_CERT_THUMBPRINT"],
        private_key_path=os.environ["SP_PRIVATE_KEY_PATH"],
        site_url=os.environ["SP_SITE_URL"].rstrip("/"),
        root_folder_server_relative_url=os.environ["SP_ROOT_FOLDER_SERVER_RELATIVE_URL"],
        days_to_keep=int(os.getenv("SP_DAYS_TO_KEEP", "30")),
        dry_run=parse_bool(os.getenv("SP_DRY_RUN"), True),
        request_timeout_seconds=int(os.getenv("SP_TIMEOUT_SECONDS", "60")),
        sleep_seconds_between_requests=float(os.getenv("SP_SLEEP_BETWEEN_REQUESTS", "0")),
        skip_folder_names=skip_folders,
    )


def merge_config(base: Config, req: CleanupRequest) -> Config:
    root = req.root_folder_server_relative_url or base.root_folder_server_relative_url
    days = req.days_to_keep if req.days_to_keep is not None else base.days_to_keep
    dry = req.dry_run if req.dry_run is not None else base.dry_run

    return Config(
        tenant_id=base.tenant_id,
        client_id=base.client_id,
        thumbprint=base.thumbprint,
        private_key_path=base.private_key_path,
        site_url=base.site_url,
        root_folder_server_relative_url=root,
        days_to_keep=days,
        dry_run=dry,
        skip_folder_names=base.skip_folder_names,
        request_timeout_seconds=base.request_timeout_seconds,
        sleep_seconds_between_requests=base.sleep_seconds_between_requests,
    )


def execute_cleanup(req: CleanupRequest) -> CleanupResponse:
    base = load_base_config_from_env()
    cfg = merge_config(base, req)
    cleaner = SharePointVersionCleaner(cfg)
    stats = cleaner.run()
    return CleanupResponse(
        status="completed",
        stats=stats,
        root_folder_server_relative_url=cfg.root_folder_server_relative_url,
        days_to_keep=cfg.days_to_keep,
        dry_run=cfg.dry_run,
    )


app = FastAPI(title="SharePoint Version Cleanup API", version="1.0.0")
job_store = JobStore()


def assert_api_key(x_api_key: Optional[str]) -> None:
    required_key = os.getenv("SP_API_KEY")
    if not required_key:
        return
    if not x_api_key or x_api_key != required_key:
        raise HTTPException(status_code=401, detail="Invalid API key")


@app.get("/health")
def health() -> Dict[str, str]:
    return {
        "status": "ok",
        "time_utc": utc_now().isoformat(),
    }


@app.post("/cleanup", response_model=CleanupResponse)
def cleanup(
    request: CleanupRequest,
    x_api_key: Optional[str] = Header(default=None, alias=DEFAULT_API_KEY_HEADER),
) -> CleanupResponse:
    assert_api_key(x_api_key)
    try:
        return execute_cleanup(request)
    except HTTPException:
        raise
    except Exception as exc:
        logging.exception("Cleanup failed")
        raise HTTPException(status_code=500, detail=str(exc)) from exc


def run_background_job(job_id: str, request: CleanupRequest) -> None:
    job_store.update(job_id, status="running")
    try:
        result = execute_cleanup(request)
        job_store.update(
            job_id,
            status="completed",
            result=result.model_dump(),
            finished_at_utc=utc_now().isoformat(),
        )
    except Exception as exc:
        logging.exception("Background cleanup failed job_id=%s", job_id)
        job_store.update(
            job_id,
            status="failed",
            error=str(exc),
            finished_at_utc=utc_now().isoformat(),
        )


@app.post("/cleanup/background", response_model=JobResponse)
def cleanup_background(
    request: CleanupRequest,
    x_api_key: Optional[str] = Header(default=None, alias=DEFAULT_API_KEY_HEADER),
) -> JobResponse:
    assert_api_key(x_api_key)
    job_id = job_store.create(request)
    thread = threading.Thread(target=run_background_job, args=(job_id, request), daemon=True)
    thread.start()
    return JobResponse(job_id=job_id, status="queued")


@app.get("/cleanup/background/{job_id}", response_model=JobDetailResponse)
def cleanup_background_status(
    job_id: str,
    x_api_key: Optional[str] = Header(default=None, alias=DEFAULT_API_KEY_HEADER),
) -> JobDetailResponse:
    assert_api_key(x_api_key)
    job = job_store.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    return JobDetailResponse(
        job_id=job["job_id"],
        status=job["status"],
        requested_at_utc=job["requested_at_utc"],
        finished_at_utc=job["finished_at_utc"],
        request=CleanupRequest(**job["request"]),
        result=CleanupResponse(**job["result"]) if job["result"] else None,
        error=job["error"],
    )


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="SharePoint version cleanup tool (CLI + HTTP API)."
    )
    p.add_argument("--serve", action="store_true", help="Start FastAPI server.")
    p.add_argument("--host", default=os.getenv("SP_API_HOST", "0.0.0.0"))
    p.add_argument("--port", type=int, default=int(os.getenv("SP_API_PORT", "8000")))
    p.add_argument("--verbose", action="store_true", help="Enable debug logging.")
    p.add_argument("--execute", action="store_true", help="Disable dry-run in CLI mode.")
    p.add_argument("--days-to-keep", type=int, help="Override retention period in CLI mode.")
    p.add_argument("--root-folder", type=str, help="Override root folder in CLI mode.")
    return p


def run_cli(args: argparse.Namespace) -> int:
    req = CleanupRequest(
        root_folder_server_relative_url=args.root_folder,
        days_to_keep=args.days_to_keep,
        dry_run=False if args.execute else None,
    )
    result = execute_cleanup(req)
    print(json.dumps(result.model_dump(), ensure_ascii=False, indent=2))
    return 0


def main() -> int:
    args = build_arg_parser().parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    if args.serve:
        uvicorn.run(
            "sharepoint_version_cleanup:app",
            host=args.host,
            port=args.port,
            reload=False,
            log_level="debug" if args.verbose else "info",
        )
        return 0
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
