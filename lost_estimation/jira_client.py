"""Thin Jira REST API client for the LOST AI estimation workflow.

Uses HTTP Basic auth with email + API token. No external Jira SDK; just `requests`.
"""

from __future__ import annotations

import os
import time
from typing import Any, Iterable

import requests


class JiraClient:
    def __init__(self) -> None:
        self.base = os.environ["JIRA_BASE_URL"].rstrip("/")
        self.auth = (os.environ["JIRA_EMAIL"], os.environ["JIRA_API_TOKEN"])
        self.session = requests.Session()
        self.session.auth = self.auth
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
        })

    # ---- low-level ---------------------------------------------------------

    def _request(self, method: str, path: str, **kwargs: Any) -> requests.Response:
        url = f"{self.base}{path}"
        for attempt in range(4):
            r = self.session.request(method, url, timeout=30, **kwargs)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", "5"))
                print(f"Jira rate-limited; sleeping {wait}s")
                time.sleep(wait)
                continue
            if r.status_code >= 500 and attempt < 3:
                time.sleep(2 ** attempt)
                continue
            return r
        return r

    # ---- search ------------------------------------------------------------

    def search(self, jql: str, fields: list[str], page_size: int = 100) -> Iterable[dict]:
        """Yield issues matching the JQL, paginated via /rest/api/3/search/jql."""
        next_token: str | None = None
        while True:
            body: dict[str, Any] = {
                "jql": jql,
                "fields": fields,
                "maxResults": page_size,
            }
            if next_token:
                body["nextPageToken"] = next_token
            r = self._request("POST", "/rest/api/3/search/jql", json=body)
            r.raise_for_status()
            data = r.json()
            for issue in data.get("issues", []):
                yield issue
            next_token = data.get("nextPageToken")
            if not next_token:
                return

    def get_issue(self, key: str, fields: list[str] | None = None) -> dict:
        params = {"fields": ",".join(fields)} if fields else {}
        r = self._request("GET", f"/rest/api/3/issue/{key}", params=params)
        r.raise_for_status()
        return r.json()

    # ---- write -------------------------------------------------------------

    def set_field(self, key: str, field_id: str, value: Any) -> bool:
        """Update a single custom field. Returns True on 204 success."""
        body = {"fields": {field_id: value}}
        r = self._request("PUT", f"/rest/api/3/issue/{key}", json=body)
        if r.status_code == 204:
            return True
        print(f"  ! Jira write failed for {key} ({field_id}={value}): {r.status_code} {r.text[:200]}")
        return False

    # ---- helpers -----------------------------------------------------------

    @staticmethod
    def adf_to_text(adf: Any) -> str:
        """Flatten Atlassian Document Format (or dict-of-content) to plain text."""
        if adf is None:
            return ""
        if isinstance(adf, str):
            return adf
        if isinstance(adf, dict):
            if "content" in adf:
                return "".join(JiraClient.adf_to_text(c) for c in adf.get("content", []))
            if adf.get("type") == "text":
                return adf.get("text", "")
            return ""
        if isinstance(adf, list):
            return "".join(JiraClient.adf_to_text(c) for c in adf)
        return ""
