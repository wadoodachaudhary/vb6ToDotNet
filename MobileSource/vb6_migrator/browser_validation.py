#!/usr/bin/env python3
"""Launch a local Blazor app in Chromium and validate browser input behavior.

Without a JSON spec this module runs a smoke probe:
- load the target URL
- enumerate visible interactive controls
- if possible, type into the first editable control and restore its value

A JSON spec can drive richer checks later without changing Python code.
Example:
{
  "url": "/counter",
  "steps": [
    { "action": "wait_for", "selector": "input" },
    { "action": "fill", "selector": "input", "value": "abc" },
    { "action": "expect_value", "selector": "input", "value": "abc" }
  ]
}
"""

from __future__ import annotations

import argparse
import json
import sys
import time
import traceback
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urljoin

LOG_ROOT = Path(__file__).resolve().parent / ".logs" / "browser_validation"
DEFAULT_TIMEOUT_MS = 15000


class BrowserValidationError(RuntimeError):
    """Raised when browser validation cannot complete successfully."""

    def __init__(self, message: str, result: "BrowserValidationResult | None" = None):
        super().__init__(message)
        self.result = result


@dataclass
class ValidationStepResult:
    index: int
    action: str
    status: str
    detail: str
    selector: str = ""
    duration_ms: int = 0


@dataclass
class BrowserValidationResult:
    ok: bool
    url: str
    visited_url: str
    headless: bool
    spec_path: str = ""
    summary: str = ""
    started_at_utc: str = ""
    finished_at_utc: str = ""
    log_path: str = ""
    screenshot_path: str = ""
    steps: List[ValidationStepResult] = field(default_factory=list)

    def to_text(self) -> str:
        lines = [
            f"Browser validation {'passed' if self.ok else 'failed'}.",
            f"Target URL: {self.url}",
            f"Visited URL: {self.visited_url or self.url}",
            f"Mode: {'headless' if self.headless else 'headed'}",
        ]
        if self.spec_path:
            lines.append(f"Spec: {self.spec_path}")
        if self.summary:
            lines.append(f"Summary: {self.summary}")
        if self.log_path:
            lines.append(f"Log: {self.log_path}")
        if self.screenshot_path:
            lines.append(f"Screenshot: {self.screenshot_path}")
        if self.steps:
            lines.append("")
            lines.append("Steps:")
            for step in self.steps:
                label = f"{step.index + 1}. [{step.status}] {step.action}"
                if step.selector:
                    label += f" :: {step.selector}"
                if step.duration_ms:
                    label += f" ({step.duration_ms} ms)"
                lines.append(label)
                if step.detail:
                    lines.append(f"   {step.detail}")
        return "\n".join(lines).strip() + "\n"


def _load_playwright():
    try:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        from playwright.sync_api import sync_playwright
    except Exception as exc:  # pragma: no cover - import-time environment issue
        raise BrowserValidationError(
            "Playwright is not installed. Run `pip install -r requirements.txt` and "
            "`python -m playwright install chromium` in vb6_migrator."
        ) from exc
    return sync_playwright, PlaywrightTimeoutError


def _truthy(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _now_utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def _step_result(index: int, action: str, detail: str, selector: str = "", *, status: str = "ok", started: float | None = None) -> ValidationStepResult:
    duration_ms = 0
    if started is not None:
        duration_ms = max(0, int((time.perf_counter() - started) * 1000))
    return ValidationStepResult(index=index, action=action, status=status, detail=detail, selector=selector, duration_ms=duration_ms)


def _write_result(result: BrowserValidationResult) -> None:
    payload = asdict(result)
    if result.log_path:
        log_path = Path(result.log_path)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _resolve_target_url(base_url: str, spec: Dict[str, Any]) -> str:
    override = str(spec.get("url") or spec.get("path") or "").strip()
    if not override:
        return base_url
    if override.startswith("http://") or override.startswith("https://"):
        return override
    return urljoin(base_url.rstrip("/") + "/", override)


def _load_spec(spec_path: str | Path | None) -> Dict[str, Any]:
    if not spec_path:
        return {}
    path = Path(spec_path).expanduser().resolve()
    if not path.exists():
        raise BrowserValidationError(f"Validation spec not found: {path}")
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise BrowserValidationError(f"Validation spec is not valid JSON: {path}\n{exc}") from exc
    if not isinstance(raw, dict):
        raise BrowserValidationError(f"Validation spec must be a JSON object: {path}")
    return raw


def _interactive_summary(page) -> List[Dict[str, str]]:
    return page.evaluate(
        """
        () => {
          const visible = (el) => {
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return !!style && style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const hint = (el) => {
            if (el.id) return `#${el.id}`;
            if (el.getAttribute('name')) return `${el.tagName.toLowerCase()}[name="${el.getAttribute('name')}"]`;
            if (el.getAttribute('aria-label')) return `${el.tagName.toLowerCase()}[aria-label="${el.getAttribute('aria-label')}"]`;
            return el.tagName.toLowerCase();
          };
          return Array.from(document.querySelectorAll('input, textarea, select, button, [contenteditable="true"]'))
            .filter(visible)
            .slice(0, 24)
            .map((el) => ({
              hint: hint(el),
              tag: el.tagName.toLowerCase(),
              type: (el.getAttribute('type') || '').toLowerCase(),
              text: ((el.innerText || el.value || el.textContent || '').trim()).slice(0, 80),
              editable: String(!!el.isContentEditable || ((el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') && !el.disabled && !el.readOnly))
            }));
        }
        """
    )


def _editable_locator(page):
    selector = (
        "textarea:not([disabled]):not([readonly]), "
        "input:not([type='hidden']):not([type='checkbox']):not([type='radio']):not([type='submit']):not([type='button']):not([disabled]):not([readonly]), "
        "[contenteditable='true']"
    )
    return page.locator(selector).first


def _read_editable_value(locator) -> str:
    return locator.evaluate(
        """
        (el) => {
          if (el.isContentEditable) return el.innerText || '';
          if ('value' in el) return el.value || '';
          return el.textContent || '';
        }
        """
    )


def _restore_editable_value(locator, value: str) -> None:
    locator.evaluate(
        """
        (el, restored) => {
          if (el.isContentEditable) {
            el.innerText = restored;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return;
          }
          if ('value' in el) {
            el.value = restored;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
          }
        }
        """,
        value,
    )


def _run_smoke_probe(page, timeout_ms: int) -> List[ValidationStepResult]:
    steps: List[ValidationStepResult] = []

    started = time.perf_counter()
    page.wait_for_selector("body", timeout=timeout_ms)
    page.wait_for_timeout(500)
    interactive = _interactive_summary(page)
    sample = ", ".join(item["hint"] for item in interactive[:8]) or "no visible controls"
    steps.append(
        _step_result(
            0,
            "smoke-scan",
            f"Found {len(interactive)} visible interactive control(s): {sample}",
            started=started,
        )
    )

    started = time.perf_counter()
    editable = _editable_locator(page)
    if editable.count() == 0:
        steps.append(
            _step_result(
                1,
                "smoke-input",
                "No editable input was found. Page load succeeded, but input typing was skipped.",
                started=started,
            )
        )
        return steps

    locator = editable
    original = _read_editable_value(locator)
    sentinel = f"smoke-{int(time.time())}"
    locator.click()
    is_contenteditable = locator.evaluate("(el) => !!el.isContentEditable")
    if is_contenteditable:
        page.keyboard.press("Meta+A")
        page.keyboard.insert_text(sentinel)
    else:
        locator.fill(sentinel)
    observed = _read_editable_value(locator)
    if sentinel not in observed:
        raise BrowserValidationError("Smoke input probe could not read back the typed value from the first editable control.")
    _restore_editable_value(locator, original)
    restored = _read_editable_value(locator)
    if restored != original:
        raise BrowserValidationError("Smoke input probe changed the field but could not restore its original value.")
    steps.append(
        _step_result(
            1,
            "smoke-input",
            f"Typed sentinel into the first editable control and restored the original value ({len(original)} original chars).",
            started=started,
        )
    )
    return steps


def _locator_for_step(page, step: Dict[str, Any], *, allow_missing: bool = False):
    selector = str(step.get("selector") or "").strip()
    if not selector and not allow_missing:
        raise BrowserValidationError(f"Validation step requires 'selector': {step}")
    locator = page.locator(selector)
    index = int(step.get("index") or 0)
    if index > 0:
        locator = locator.nth(index)
    else:
        locator = locator.first
    return selector, locator


def _step_timeout_ms(step: Dict[str, Any], default_timeout_ms: int) -> int:
    raw = step.get("timeout_ms")
    if raw in (None, ""):
        return default_timeout_ms
    return max(1, int(raw))


def _run_spec_step(page, step: Dict[str, Any], index: int, artifact_dir: Path, default_timeout_ms: int) -> ValidationStepResult:
    action = str(step.get("action") or step.get("type") or "").strip().lower()
    if not action:
        raise BrowserValidationError(f"Validation step is missing 'action': {step}")

    started = time.perf_counter()
    timeout_ms = _step_timeout_ms(step, default_timeout_ms)
    selector, locator = _locator_for_step(page, step, allow_missing=action in {"wait", "sleep", "expect_url_contains", "goto", "screenshot"})

    if action == "goto":
        target = str(step.get("url") or step.get("value") or "").strip()
        if not target:
            raise BrowserValidationError("goto step requires 'url' or 'value'.")
        page.goto(target, wait_until="domcontentloaded", timeout=timeout_ms)
        return _step_result(index, action, f"Navigated to {page.url}", started=started)

    if action == "wait_for":
        state = str(step.get("state") or "visible").strip() or "visible"
        locator.wait_for(state=state, timeout=timeout_ms)
        return _step_result(index, action, f"Locator reached state '{state}'.", selector=selector, started=started)

    if action == "click":
        locator.click(timeout=timeout_ms)
        return _step_result(index, action, "Clicked locator.", selector=selector, started=started)

    if action == "dblclick":
        locator.dblclick(timeout=timeout_ms)
        return _step_result(index, action, "Double-clicked locator.", selector=selector, started=started)

    if action == "fill":
        value = str(step.get("value") or "")
        locator.fill(value, timeout=timeout_ms)
        return _step_result(index, action, f"Filled {len(value)} character(s).", selector=selector, started=started)

    if action == "type":
        value = str(step.get("value") or "")
        locator.click(timeout=timeout_ms)
        locator.type(value, timeout=timeout_ms)
        return _step_result(index, action, f"Typed {len(value)} character(s).", selector=selector, started=started)

    if action == "press":
        key = str(step.get("key") or step.get("value") or "").strip()
        if not key:
            raise BrowserValidationError("press step requires 'key' or 'value'.")
        locator.press(key, timeout=timeout_ms)
        return _step_result(index, action, f"Pressed {key}.", selector=selector, started=started)

    if action == "check":
        locator.check(timeout=timeout_ms)
        return _step_result(index, action, "Checked locator.", selector=selector, started=started)

    if action == "uncheck":
        locator.uncheck(timeout=timeout_ms)
        return _step_result(index, action, "Unchecked locator.", selector=selector, started=started)

    if action == "select":
        value = step.get("value")
        label = step.get("label")
        if value is not None:
            locator.select_option(value=str(value), timeout=timeout_ms)
            detail = f"Selected option value '{value}'."
        elif label is not None:
            locator.select_option(label=str(label), timeout=timeout_ms)
            detail = f"Selected option label '{label}'."
        else:
            raise BrowserValidationError("select step requires 'value' or 'label'.")
        return _step_result(index, action, detail, selector=selector, started=started)

    if action == "expect_visible":
        if not locator.is_visible(timeout=timeout_ms):
            raise BrowserValidationError(f"Locator is not visible: {selector}")
        return _step_result(index, action, "Locator is visible.", selector=selector, started=started)

    if action == "expect_text":
        actual = (locator.inner_text(timeout=timeout_ms) or "").strip()
        expected = str(step.get("value") or "").strip()
        contains = str(step.get("contains") or "").strip()
        if expected:
            if actual != expected:
                raise BrowserValidationError(f"Expected exact text '{expected}' but found '{actual}'.")
            detail = f"Matched exact text '{expected}'."
        elif contains:
            if contains not in actual:
                raise BrowserValidationError(f"Expected text containing '{contains}' but found '{actual}'.")
            detail = f"Matched text containing '{contains}'."
        else:
            raise BrowserValidationError("expect_text step requires 'value' or 'contains'.")
        return _step_result(index, action, detail, selector=selector, started=started)

    if action == "expect_value":
        actual = _read_editable_value(locator)
        expected = str(step.get("value") or "")
        contains = str(step.get("contains") or "")
        if expected:
            if actual != expected:
                raise BrowserValidationError(f"Expected value '{expected}' but found '{actual}'.")
            detail = f"Matched exact value '{expected}'."
        elif contains:
            if contains not in actual:
                raise BrowserValidationError(f"Expected value containing '{contains}' but found '{actual}'.")
            detail = f"Matched value containing '{contains}'."
        else:
            raise BrowserValidationError("expect_value step requires 'value' or 'contains'.")
        return _step_result(index, action, detail, selector=selector, started=started)

    if action == "expect_checked":
        expected = _truthy(step.get("value"), True)
        actual = locator.is_checked(timeout=timeout_ms)
        if actual != expected:
            raise BrowserValidationError(f"Expected checked={expected} but found checked={actual}.")
        return _step_result(index, action, f"Matched checked={expected}.", selector=selector, started=started)

    if action == "expect_count":
        raw_selector = str(step.get("selector") or "").strip()
        if not raw_selector:
            raise BrowserValidationError("expect_count step requires 'selector'.")
        actual = page.locator(raw_selector).count()
        expected = int(step.get("value") or step.get("count") or 0)
        if actual != expected:
            raise BrowserValidationError(f"Expected {expected} match(es) for '{raw_selector}' but found {actual}.")
        return _step_result(index, action, f"Matched count {expected}.", selector=raw_selector, started=started)

    if action == "expect_url_contains":
        fragment = str(step.get("value") or step.get("contains") or "").strip()
        if not fragment:
            raise BrowserValidationError("expect_url_contains step requires 'value' or 'contains'.")
        if fragment not in page.url:
            raise BrowserValidationError(f"Expected current URL to contain '{fragment}' but found '{page.url}'.")
        return _step_result(index, action, f"Current URL contains '{fragment}'.", started=started)

    if action in {"wait", "sleep"}:
        delay_ms = max(0, int(step.get("ms") or step.get("value") or 500))
        page.wait_for_timeout(delay_ms)
        return _step_result(index, action, f"Waited {delay_ms} ms.", started=started)

    if action == "screenshot":
        file_name = str(step.get("file") or f"step-{index + 1:02d}.png").strip()
        shot_path = artifact_dir / file_name
        shot_path.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(shot_path), full_page=_truthy(step.get("full_page"), False))
        return _step_result(index, action, f"Saved screenshot to {shot_path}.", selector=selector, started=started)

    raise BrowserValidationError(f"Unsupported validation action '{action}'.")


def run_browser_validation(url: str, spec_path: str | Path | None = None, *, headless: bool = False, timeout_ms: int = DEFAULT_TIMEOUT_MS) -> BrowserValidationResult:
    sync_playwright, PlaywrightTimeoutError = _load_playwright()
    spec = _load_spec(spec_path)
    target_url = _resolve_target_url(url, spec)

    stamp = time.strftime("%Y%m%d-%H%M%S")
    artifact_dir = LOG_ROOT / stamp
    artifact_dir.mkdir(parents=True, exist_ok=True)
    screenshot_path = artifact_dir / "final.png"
    result = BrowserValidationResult(
        ok=False,
        url=url,
        visited_url=target_url,
        headless=_truthy(spec.get("headless"), headless),
        spec_path=str(Path(spec_path).expanduser().resolve()) if spec_path else "",
        started_at_utc=_now_utc(),
        log_path=str((artifact_dir / "result.json").resolve()),
        screenshot_path=str(screenshot_path.resolve()),
    )

    try:
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=result.headless)
            context = browser.new_context(viewport={"width": 1440, "height": 1024})
            page = context.new_page()
            try:
                page.goto(target_url, wait_until="domcontentloaded", timeout=timeout_ms)
                page.wait_for_timeout(int(spec.get("initial_wait_ms") or 600))

                steps_data = spec.get("steps") if isinstance(spec.get("steps"), list) else []
                if steps_data:
                    for index, step in enumerate(steps_data):
                        if not isinstance(step, dict):
                            raise BrowserValidationError(f"Validation step {index + 1} is not an object.")
                        result.steps.append(_run_spec_step(page, step, index, artifact_dir, timeout_ms))
                else:
                    result.steps.extend(_run_smoke_probe(page, timeout_ms))

                page.screenshot(path=str(screenshot_path), full_page=True)
                result.visited_url = page.url
                result.ok = True
                result.summary = f"Completed {len(result.steps)} validation step(s)."
                return result
            finally:
                try:
                    browser.close()
                except Exception:
                    pass
    except PlaywrightTimeoutError as exc:
        result.summary = f"Timed out while validating browser behavior: {exc}"
        raise BrowserValidationError(result.summary, result) from exc
    except BrowserValidationError as exc:
        if exc.result is not None:
            result = exc.result
        result.summary = str(exc)
        raise BrowserValidationError(result.summary, result) from exc
    except Exception as exc:
        result.summary = str(exc)
        raise BrowserValidationError(str(exc), result) from exc
    finally:
        result.finished_at_utc = _now_utc()
        if not screenshot_path.exists():
            result.screenshot_path = ""
        _write_result(result)


def _print_failure(exc: BrowserValidationError) -> None:
    result = exc.result
    if result is not None:
        text = result.to_text()
    else:
        text = f"Browser validation failed.\nSummary: {exc}\n"
    sys.stderr.write(text)
    if result is not None:
        sys.stderr.write(json.dumps(asdict(result), indent=2) + "\n")
    else:
        traceback.print_exc()


def main() -> int:
    parser = argparse.ArgumentParser(description="Launch Chromium against a local Blazor URL and validate browser input behavior.")
    parser.add_argument("--url", required=True, help="Base URL to open, for example http://127.0.0.1:5159")
    parser.add_argument("--spec", default="", help="Optional JSON validation spec path")
    parser.add_argument("--headless", default="false", choices=["true", "false"], help="Run Chromium headless instead of showing the browser")
    parser.add_argument("--timeout-ms", type=int, default=DEFAULT_TIMEOUT_MS, help="Per-step timeout in milliseconds")
    args = parser.parse_args()

    try:
        result = run_browser_validation(
            args.url,
            args.spec or None,
            headless=_truthy(args.headless),
            timeout_ms=args.timeout_ms,
        )
    except BrowserValidationError as exc:
        _print_failure(exc)
        return 1

    sys.stdout.write(result.to_text())
    sys.stdout.write(json.dumps(asdict(result), indent=2) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
