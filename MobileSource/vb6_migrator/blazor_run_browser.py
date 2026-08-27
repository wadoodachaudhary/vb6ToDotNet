#!/usr/bin/env python3
"""Open the Blazor app URL when `dotnet run` reports listening endpoints."""

from __future__ import annotations

import re
import threading
import webbrowser


class BlazorRunBrowser:
    """Tracks run output and opens the first discovered local web URL once."""

    _LISTEN_RE = re.compile(r"Now listening on:\s*(https?://[^\s]+)", re.IGNORECASE)
    _LOCAL_RE = re.compile(r"(https?://(?:localhost|127\.0\.0\.1):\d+[^\s]*)", re.IGNORECASE)

    def __init__(self) -> None:
        self._opened = False
        self._opened_url = ""

    @property
    def opened_url(self) -> str:
        return self._opened_url

    def reset(self) -> None:
        self._opened = False
        self._opened_url = ""

    def _extract_url(self, text: str) -> str:
        m = self._LISTEN_RE.search(text)
        if m:
            return m.group(1).strip()
        m = self._LOCAL_RE.search(text)
        if m:
            return m.group(1).strip()
        return ""

    def maybe_open_from_text(self, text: str) -> str:
        """Open browser once if a URL is found. Returns opened URL or empty string."""
        if self._opened or not text:
            return ""
        url = self._extract_url(text)
        if not url:
            return ""
        self._opened = True
        self._opened_url = url
        threading.Thread(
            target=lambda: webbrowser.open(url, new=1, autoraise=True),
            daemon=True,
        ).start()
        return url
