#!/usr/bin/env python3
"""Helpers that add basic Blazor logging to selected .razor and .cs files."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception:
        try:
            return path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return path.read_text(encoding="latin-1", errors="ignore")


def _find_matching_brace(text: str, open_index: int) -> int:
    """Find the matching closing brace for a brace-open index."""
    if open_index < 0 or open_index >= len(text) or text[open_index] != "{":
        return -1

    depth = 0
    i = open_index
    in_string = False
    in_char = False
    in_verbatim = False
    in_line_comment = False
    in_block_comment = False

    while i < len(text):
        ch = text[i]
        nxt = text[i + 1] if i + 1 < len(text) else ""

        if in_line_comment:
            if ch == "\n":
                in_line_comment = False
            i += 1
            continue

        if in_block_comment:
            if ch == "*" and nxt == "/":
                in_block_comment = False
                i += 2
                continue
            i += 1
            continue

        if in_string:
            if in_verbatim:
                if ch == '"' and nxt == '"':
                    i += 2
                    continue
                if ch == '"':
                    in_string = False
                    in_verbatim = False
                i += 1
                continue
            if ch == "\\":
                i += 2
                continue
            if ch == '"':
                in_string = False
            i += 1
            continue

        if in_char:
            if ch == "\\":
                i += 2
                continue
            if ch == "'":
                in_char = False
            i += 1
            continue

        if ch == "/" and nxt == "/":
            in_line_comment = True
            i += 2
            continue
        if ch == "/" and nxt == "*":
            in_block_comment = True
            i += 2
            continue
        if ch == "@" and nxt == '"':
            in_string = True
            in_verbatim = True
            i += 2
            continue
        if ch == '"':
            in_string = True
            in_verbatim = False
            i += 1
            continue
        if ch == "'":
            in_char = True
            i += 1
            continue

        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1

    return -1


def _inject_log_into_method(source: str, method_name: str, log_call: str) -> tuple[str, bool, bool]:
    """
    Inject a log call into an existing method.
    Returns: (new_source, changed, method_found)
    """
    sig_re = re.compile(
        rf"(?m)^[ \t]*(?:public|protected|private|internal)[^\n]*\b{method_name}\s*\([^)]*\)"
    )
    match = sig_re.search(source)
    if not match:
        return source, False, False

    sig_end = match.end()
    pos = sig_end
    while pos < len(source) and source[pos].isspace():
        pos += 1
    if pos >= len(source):
        return source, False, True

    if source.startswith("=>", pos):
        semi = source.find(";", pos + 2)
        if semi == -1:
            return source, False, True
        expr = source[pos + 2 : semi].strip().rstrip(";")
        if "LogInformation(" in expr:
            return source, False, True
        line_start = source.rfind("\n", 0, match.start()) + 1
        indent = source[line_start:match.start()]
        signature = source[match.start() : sig_end].rstrip()
        is_void = method_name == "OnInitialized" and " void " in f" {signature} "
        expr_line = f"{expr};" if is_void else f"return {expr};"
        block = (
            f"{signature}\n"
            f"{indent}{{\n"
            f"{indent}    {log_call}\n"
            f"{indent}    {expr_line}\n"
            f"{indent}}}"
        )
        return source[: match.start()] + block + source[semi + 1 :], True, True

    brace_open = source.find("{", pos)
    if brace_open == -1:
        return source, False, True
    brace_close = _find_matching_brace(source, brace_open)
    if brace_close == -1:
        return source, False, True

    body = source[brace_open + 1 : brace_close]
    if "LogInformation(" in body:
        return source, False, True

    line_start = source.rfind("\n", 0, match.start()) + 1
    indent = source[line_start:match.start()]
    insertion = f"\n{indent}    {log_call}"
    return source[: brace_open + 1] + insertion + source[brace_open + 1 :], True, True


def _find_razor_code_block(source: str) -> tuple[int, int] | None:
    match = re.search(r"@code\s*{", source)
    if not match:
        return None
    brace_open = match.end() - 1
    brace_close = _find_matching_brace(source, brace_open)
    if brace_close == -1:
        return None
    return brace_open, brace_close


def _ensure_razor_logger(source: str, component_name: str) -> tuple[str, str, bool]:
    inject_match = re.search(
        r"@inject\s+ILogger(?:<[^>]+>)?\s+([A-Za-z_][A-Za-z0-9_]*)",
        source,
    )
    if inject_match:
        return source, inject_match.group(1), False

    lines = source.splitlines()
    insert_at = 0
    while insert_at < len(lines) and lines[insert_at].lstrip().startswith("@"):
        insert_at += 1

    add_lines: list[str] = []
    if "Microsoft.Extensions.Logging" not in source:
        add_lines.append("@using Microsoft.Extensions.Logging")
    add_lines.append(f"@inject ILogger<{component_name}> Logger")
    lines[insert_at:insert_at] = add_lines

    updated = "\n".join(lines)
    if source.endswith("\n"):
        updated += "\n"
    return updated, "Logger", True


def _instrument_razor(source: str, path: Path) -> tuple[str, bool, str]:
    component_name = path.stem
    updated = source
    changed = False

    updated, logger_name, logger_changed = _ensure_razor_logger(updated, component_name)
    changed = changed or logger_changed
    log_call = f'{logger_name}.LogInformation("{component_name} initialized.");'

    method_found = False
    for method_name in ("OnInitialized", "OnInitializedAsync"):
        updated, injected, found = _inject_log_into_method(updated, method_name, log_call)
        changed = changed or injected
        method_found = method_found or found
        if injected:
            return updated, True, "updated method"

    if method_found:
        return updated, changed, "logger already present"

    code_block = _find_razor_code_block(updated)
    method_text = (
        "\n"
        "    protected override void OnInitialized()\n"
        "    {\n"
        f"        {log_call}\n"
        "        base.OnInitialized();\n"
        "    }\n"
    )
    if code_block:
        _, block_close = code_block
        updated = updated[:block_close] + method_text + updated[block_close:]
    else:
        updated = (
            updated.rstrip()
            + "\n\n@code {\n"
            "    protected override void OnInitialized()\n"
            "    {\n"
            f"        {log_call}\n"
            "        base.OnInitialized();\n"
            "    }\n"
            "}\n"
        )
    return updated, True, "added OnInitialized"


def _ensure_using(source: str, namespace: str) -> tuple[str, bool]:
    using_re = re.compile(rf"(?m)^\s*using\s+{re.escape(namespace)}\s*;\s*$")
    if using_re.search(source):
        return source, False
    using_lines = list(re.finditer(r"(?m)^\s*using\s+[^\n;]+;\s*$", source))
    insert_text = f"\nusing {namespace};"
    if using_lines:
        idx = using_lines[-1].end()
        return source[:idx] + insert_text + source[idx:], True
    return f"using {namespace};\n{source}", True


def _find_first_class(source: str) -> tuple[str, int, int, str] | None:
    class_match = re.search(
        r"(?m)^([ \t]*)(?:public|protected|private|internal|abstract|sealed|static|partial|\s)*class\s+([A-Za-z_][A-Za-z0-9_]*)",
        source,
    )
    if not class_match:
        return None
    class_name = class_match.group(2)
    class_indent = class_match.group(1)
    brace_open = source.find("{", class_match.end())
    if brace_open == -1:
        return None
    brace_close = _find_matching_brace(source, brace_open)
    if brace_close == -1:
        return None
    return class_name, brace_open, brace_close, class_indent


def _find_cs_logger_name(source: str) -> str | None:
    match = re.search(
        r"ILogger(?:<[^>]+>)?\s*\??\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\{|=|;)",
        source,
    )
    if match:
        return match.group(1)
    return None


def _instrument_cs(source: str, path: Path) -> tuple[str, bool, str]:
    # Keep this conservative so plain non-component classes are not modified unexpectedly.
    is_component = (
        path.name.lower().endswith(".razor.cs")
        or "ComponentBase" in source
        or "OnInitialized" in source
    )
    if not is_component:
        return source, False, "skipped non-component .cs file"

    class_info = _find_first_class(source)
    if class_info is None:
        return source, False, "no class found"
    class_name, class_open, class_close, class_indent = class_info

    updated = source
    changed = False

    updated, using_changed = _ensure_using(updated, "Microsoft.Extensions.Logging")
    changed = changed or using_changed

    logger_name = _find_cs_logger_name(updated)
    if logger_name is None:
        logger_name = "Logger"
        class_info = _find_first_class(updated)
        if class_info is None:
            return source, False, "no class found after using insert"
        class_name, class_open, class_close, class_indent = class_info
        prop_indent = class_indent + "    "
        logger_prop = (
            "\n"
            f"{prop_indent}[Microsoft.AspNetCore.Components.Inject]\n"
            f"{prop_indent}protected ILogger<{class_name}>? {logger_name} {{ get; set; }}\n"
        )
        updated = updated[: class_open + 1] + logger_prop + updated[class_open + 1 :]
        changed = True

    log_call = f'{logger_name}?.LogInformation("{class_name} initialized.");'

    method_found = False
    for method_name in ("OnInitialized", "OnInitializedAsync"):
        updated, injected, found = _inject_log_into_method(updated, method_name, log_call)
        changed = changed or injected
        method_found = method_found or found
        if injected:
            return updated, True, "updated method"

    if method_found:
        return updated, changed, "logger already present"

    class_info = _find_first_class(updated)
    if class_info is None:
        return source, False, "no class found for method insert"
    class_name, _class_open, class_close, class_indent = class_info
    method_indent = class_indent + "    "
    method_text = (
        "\n"
        f"{method_indent}protected override void OnInitialized()\n"
        f"{method_indent}{{\n"
        f"{method_indent}    {log_call}\n"
        f"{method_indent}    base.OnInitialized();\n"
        f"{method_indent}}}\n"
    )
    updated = updated[:class_close] + method_text + updated[class_close:]
    changed = True
    return updated, changed, "added OnInitialized"


def _instrument_single_file(path: Path) -> tuple[str, str]:
    suffix = path.suffix.lower()
    if suffix == ".razor":
        source = _read_text(path)
        updated, changed, reason = _instrument_razor(source, path)
        if changed and updated != source:
            path.write_text(updated, encoding="utf-8")
            return "updated", reason
        return "skipped", reason

    if suffix == ".cs":
        source = _read_text(path)
        updated, changed, reason = _instrument_cs(source, path)
        if changed and updated != source:
            path.write_text(updated, encoding="utf-8")
            return "updated", reason
        return "skipped", reason

    return "skipped", f"unsupported extension: {suffix}"


def apply_logging_to_files(paths: Iterable[Path | str]) -> dict:
    """
    Add logging to selected files and return an operation summary.
    """
    unique_paths: list[Path] = []
    seen: set[str] = set()

    for raw in paths:
        path = Path(raw).expanduser()
        key = str(path.resolve()) if path.exists() else str(path)
        if key in seen:
            continue
        seen.add(key)
        unique_paths.append(path)

    results: list[dict] = []
    for path in unique_paths:
        if not path.exists():
            results.append({"path": str(path), "status": "skipped", "reason": "file not found"})
            continue
        if not path.is_file():
            results.append({"path": str(path), "status": "skipped", "reason": "not a file"})
            continue

        try:
            status, reason = _instrument_single_file(path)
            results.append({"path": str(path), "status": status, "reason": reason})
        except Exception as exc:  # pragma: no cover - defensive path
            results.append({"path": str(path), "status": "error", "reason": str(exc)})

    updated = sum(1 for r in results if r["status"] == "updated")
    skipped = sum(1 for r in results if r["status"] == "skipped")
    errors = sum(1 for r in results if r["status"] == "error")
    return {
        "processed": len(results),
        "updated": updated,
        "skipped": skipped,
        "errors": errors,
        "results": results,
    }

