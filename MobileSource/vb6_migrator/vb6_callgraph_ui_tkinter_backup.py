#!/usr/bin/env python3
"""
VB6 call graph generator with optional UI tree viewer.

Outputs (under <HomeFront>/vb6_graphs):
  - vb6_callgraph_prefer_hfest.csv
  - vb6_callgraph_prefer_hfest.mmd
  - vb6_calltree_prefer_hfest.txt

Usage:
  python3 vb6_callgraph_ui.py \
      --root /Users/wadood/projects/VBToCSharp/HomeFront/HomeFrontVB6 \
      --prefer-hfest true \
      --ui true
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import socket
import signal
import subprocess
import shutil
import sys
import threading
import time
from datetime import datetime, timezone
import urllib.error
import urllib.request
import hashlib
import difflib
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

try:
    from blazor_run_browser import BlazorRunBrowser
except Exception:
    class BlazorRunBrowser:  # type: ignore[no-redef]
        def __init__(self) -> None:
            self._opened_url = ""

        @property
        def opened_url(self) -> str:
            return self._opened_url

        def reset(self) -> None:
            self._opened_url = ""

        def maybe_open_from_text(self, text: str) -> str:
            return ""

try:
    from PIL import Image, ImageChops, ImageStat, ImageTk  # type: ignore
except Exception:
    Image = None
    ImageChops = None
    ImageStat = None
    ImageTk = None

# macOS: promote Python to a foreground app so tkinter windows can receive focus
def _activate_macos_app() -> None:
    if sys.platform != "darwin":
        return
    try:
        pid = os.getpid()
        subprocess.Popen(
            [
                "osascript",
                "-e",
                f'tell application "System Events" to set frontmost of the first process whose unix id is {pid} to true',
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        pass

_activate_macos_app()


def _pick_free_local_port() -> int:
    """Pick an available localhost TCP port for dotnet run."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return int(s.getsockname()[1])


RUN_BIND_FAIL_RE = re.compile(r"failed to bind to address\s+https?://[^:\s]+:(\d+)", re.IGNORECASE)


def _extract_bind_fail_port(text: str) -> int:
    m = RUN_BIND_FAIL_RE.search(text or "")
    if not m:
        return 0
    try:
        return int(m.group(1))
    except Exception:
        return 0


def _kill_listeners_on_port(port: int) -> List[int]:
    if port <= 0:
        return []
    try:
        probe = subprocess.run(
            ["lsof", "-tiTCP:" + str(port), "-sTCP:LISTEN"],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        return []
    pids: List[int] = []
    me = os.getpid()
    for line in (probe.stdout or "").splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            pid = int(line)
        except Exception:
            continue
        if pid != me:
            pids.append(pid)
    if not pids:
        return []
    for pid in pids:
        try:
            os.kill(pid, 15)
        except Exception:
            pass
    time.sleep(0.35)
    for pid in pids:
        try:
            os.kill(pid, 0)
            os.kill(pid, 9)
        except Exception:
            pass
    return pids


def _terminate_process(proc: subprocess.Popen, timeout_sec: float = 2.0) -> None:
    """Terminate a process and its group if possible, then force kill."""
    if proc is None:
        return
    try:
        if proc.poll() is not None:
            return
    except Exception:
        return

    terminated = False
    try:
        os.killpg(proc.pid, signal.SIGTERM)
        terminated = True
    except Exception:
        pass
    if not terminated:
        try:
            proc.terminate()
            terminated = True
        except Exception:
            pass

    end = time.time() + timeout_sec
    while time.time() < end:
        try:
            if proc.poll() is not None:
                return
        except Exception:
            return
        time.sleep(0.05)

    try:
        os.killpg(proc.pid, signal.SIGKILL)
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass

VB_NAME_RE = re.compile(r'Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.IGNORECASE)
MAX_FILES_TO_LOAD = 0  # Debug: limit total VB files read (0 = all files)

RX_LOAD = re.compile(r'(?i)\bLoad\s+([A-Za-z_][A-Za-z0-9_]*)')
RX_UNLOAD = re.compile(r'(?i)\bUnload\s+([A-Za-z_][A-Za-z0-9_]*)')
RX_SET = re.compile(r'(?i)\bSet\s+\w+\s*=\s*([A-Za-z_][A-Za-z0-9_]*)')
RX_NEW = re.compile(r'(?i)\bNew\s+([A-Za-z_][A-Za-z0-9_]*)')
RX_DOT = re.compile(r'(?i)\b([A-Za-z_][A-Za-z0-9_]*)\s*\.')
SUB_MAIN_RE = re.compile(r'(?i)\bSub\s+Main\b')

DEFAULT_PROJECT_NAME = "PrescionBuilder"
PROJECT_MANIFEST_NAME = "project.json"
PROJECT_SOURCE_DIR = "source"
PROJECT_TARGET_DIR = "target"
PROJECT_CACHE_DIR = "cache"
PROJECT_GRAPHS_DIR = "graphs"
EMPTY_PROJECT_NAME = "__NoProject__"

# VB6 files used for dependency scanning and display.
VB6_COMPONENT_TYPES: Dict[str, str] = {
    ".bas": "Module",
    ".cls": "Class",
    ".frm": "Form",
    ".ctl": "Control",
    ".ctrl": "Control",
    ".vb": "VB",
}

# Additional sidecar files copied into project source packages.
VB6_COPY_EXTENSIONS = set(VB6_COMPONENT_TYPES.keys()) | {".frx", ".ctx", ".vbp", ".vbg"}
IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"]


@dataclass(frozen=True)
class Component:
    project: str
    name: str
    path: Path
    type: str


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="latin-1", errors="ignore")
    except Exception:
        return path.read_text(errors="ignore")


def get_vb_name(path: Path) -> str:
    txt = read_text(path)
    m = VB_NAME_RE.search(txt)
    return m.group(1).strip() if m else path.stem


def normalize_rel(rel: str) -> str:
    return rel.replace("\\", "/")


def compute_signature(components: List[Component]) -> str:
    items = []
    for comp in components:
        try:
            stat = comp.path.stat()
            items.append(f"{comp.path}|{stat.st_mtime}|{stat.st_size}")
        except Exception:
            items.append(f"{comp.path}|0|0")
    items.sort()
    raw = "\n".join(items).encode("utf-8", errors="ignore")
    return hashlib.sha1(raw).hexdigest()


def load_cache_index(cache_dir: Path) -> dict:
    cache_dir.mkdir(parents=True, exist_ok=True)
    index_path = cache_dir / "index.json"
    if not index_path.exists():
        return {"version": 1, "files": {}, "tree": {}}
    try:
        return json.loads(index_path.read_text(encoding="utf-8"))
    except Exception:
        return {"version": 1, "files": {}, "tree": {}}


def save_cache_index(cache_dir: Path, index: dict) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    index_path = cache_dir / "index.json"
    index_path.write_text(json.dumps(index, indent=2), encoding="utf-8")


def load_tree_cache(cache_dir: Path, signature: str, prefer_hfest: bool) -> dict | None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    key = f"tree_{signature}_{'hfest' if prefer_hfest else 'combined'}.json"
    path = cache_dir / key
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def save_tree_cache(cache_dir: Path, signature: str, prefer_hfest: bool, edges: dict, roots: list[str]) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    key = f"tree_{signature}_{'hfest' if prefer_hfest else 'combined'}.json"
    path = cache_dir / key
    edges_list = []
    for (caller, callee), meta in edges.items():
        edges_list.append([caller, callee, bool(meta.get("ambiguous"))])
    payload = {"signature": signature, "edges": edges_list, "roots": roots}
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

def chunk_vb6_source(comp: Component, src: str) -> str:
    lines = src.splitlines()
    out: List[str] = []

    # Form/control layout chunks
    if comp.path.suffix.lower() in {".frm", ".ctl"}:
        begin_re = re.compile(r"^\s*Begin\s+([A-Za-z0-9_.]+)\s+([A-Za-z0-9_]+)", re.IGNORECASE)
        for line in lines:
            m = begin_re.match(line)
            if m:
                ctl_type = m.group(1)
                ctl_name = m.group(2)
                out.append(f"' === [Form Element] {ctl_type} {ctl_name} ===")
            out.append(line)
    else:
        out = lines[:]

    # Code chunks (functions/subs/properties/types/enums)
    code_out: List[str] = []
    module_kind = ""
    ext = comp.path.suffix.lower()
    if ext == ".cls":
        module_kind = "Class"
    elif ext == ".bas":
        module_kind = "Module"
    elif ext in {".frm", ".ctl"}:
        module_kind = "Form Code"
    if module_kind:
        code_out.append(f"' === [{module_kind}] {comp.name} ===")

    sub_fun_re = re.compile(r"^\s*(Public|Private|Friend|Static)?\s*(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
    prop_re = re.compile(r"^\s*(Public|Private|Friend|Static)?\s*Property\s+(Get|Let|Set)\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
    type_re = re.compile(r"^\s*Type\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
    enum_re = re.compile(r"^\s*Enum\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)
    decl_re = re.compile(r"^\s*(Public|Private)?\s*Declare\s+(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)", re.IGNORECASE)

    for line in out:
        m = sub_fun_re.match(line)
        if m:
            kind = m.group(2).capitalize()
            name = m.group(3)
            code_out.append(f"' === [{kind}] {name} ===")
        else:
            m = prop_re.match(line)
            if m:
                kind = f"Property {m.group(2).capitalize()}"
                name = m.group(3)
                code_out.append(f"' === [{kind}] {name} ===")
            else:
                m = type_re.match(line)
                if m:
                    code_out.append(f"' === [Type] {m.group(1)} ===")
                else:
                    m = enum_re.match(line)
                    if m:
                        code_out.append(f"' === [Enum] {m.group(1)} ===")
                    else:
                        m = decl_re.match(line)
                        if m:
                            kind = f"Declare {m.group(2).capitalize()}"
                            name = m.group(3)
                            code_out.append(f"' === [{kind}] {name} ===")
        code_out.append(line)

    return "\n".join(code_out)


def count_lines_file(path: Path) -> int:
    count = 0
    try:
        with path.open("r", encoding="latin-1", errors="ignore") as f:
            for _ in f:
                count += 1
    except Exception:
        return 0
    return count


def extract_response_text(payload: dict) -> str:
    if isinstance(payload, dict):
        if isinstance(payload.get("output_text"), str):
            return payload["output_text"]
        if "output" in payload:
            parts: List[str] = []
            for item in payload.get("output", []):
                for content in item.get("content", []):
                    if content.get("type") == "output_text" and "text" in content:
                        parts.append(content["text"])
            if parts:
                return "\n".join(parts)
        if "choices" in payload:
            choice = payload.get("choices", [{}])[0]
            msg = choice.get("message", {})
            if isinstance(msg, dict):
                return msg.get("content", "")
    return ""


def compute_migration_order(
    components: List[Component], edges: Dict[Tuple[str, str], Dict]
) -> tuple[list[str], set[str]]:
    comp_ids = [f"{c.project}::{c.name}" for c in components]
    comp_by_id = {f"{c.project}::{c.name}": c for c in components}

    rev_adj: Dict[str, List[str]] = {cid: [] for cid in comp_ids}
    indeg: Dict[str, int] = {cid: 0 for cid in comp_ids}

    for (caller, callee), _meta in edges.items():
        if callee not in rev_adj:
            rev_adj[callee] = []
            indeg[callee] = 0
        if caller not in rev_adj:
            rev_adj[caller] = []
            indeg[caller] = 0
        rev_adj[callee].append(caller)
        indeg[caller] += 1

    def priority(cid: str) -> tuple:
        comp = comp_by_id.get(cid)
        ext = comp.path.suffix.lower() if comp else ""
        ext_rank = {".bas": 0, ".cls": 1, ".ctl": 2, ".frm": 3, ".vbp": 4}.get(ext, 5)
        name = comp.name if comp else cid
        return (ext_rank, name.lower(), cid)

    import heapq

    heap = []
    for cid, deg in indeg.items():
        if deg == 0:
            heapq.heappush(heap, (priority(cid), cid))

    order: List[str] = []
    while heap:
        _prio, cid = heapq.heappop(heap)
        order.append(cid)
        for nxt in rev_adj.get(cid, []):
            indeg[nxt] -= 1
            if indeg[nxt] == 0:
                heapq.heappush(heap, (priority(nxt), nxt))

    remaining = [cid for cid, deg in indeg.items() if deg > 0]
    remaining.sort(key=priority)
    order.extend(remaining)
    cycle_set = set(remaining)
    return order, cycle_set


def collect_context_files(context_dir: Path) -> List[Path]:
    if not context_dir or not context_dir.exists():
        return []
    files: List[Path] = []
    for pattern in ("*.razor", "*.cs"):
        files.extend(context_dir.rglob(pattern))
    return sorted([p for p in files if p.is_file()])


def compute_context_signature(files: List[Path]) -> str:
    items = []
    for path in files:
        try:
            stat = path.stat()
            items.append(f"{path}|{stat.st_mtime}|{stat.st_size}")
        except Exception:
            items.append(f"{path}|0|0")
    raw = "\n".join(items).encode("utf-8", errors="ignore")
    return hashlib.sha1(raw).hexdigest()


def build_context_blob(context_dir: Path, files: List[Path]) -> str:
    if not files:
        return ""
    parts = ["CONTEXT FILES (read-only):"]
    for path in files:
        try:
            content = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            content = ""
        try:
            rel = str(path.relative_to(context_dir))
        except Exception:
            rel = str(path)
        ext = path.suffix.lstrip(".") or "text"
        parts.append(f"FILE: {rel}\n```{ext}\n{content}\n```")
    return "\n\n".join(parts)


def find_component_image(image_dir: Path, component_name: str) -> Path | None:
    """Resolve a screenshot path for a component name using common naming patterns."""
    if not image_dir.exists():
        return None
    base = component_name.strip()
    if not base:
        return None
    candidates = [
        base,
        f"{base}Razor",
        f"{base}Frm",
        f"Screen_{base}",
        f"Screen_{base}Razor",
    ]
    for stem in candidates:
        for ext in IMAGE_EXTENSIONS:
            p = image_dir / f"{stem}{ext}"
            if p.exists():
                return p
    return None


def image_dimensions(path: Path) -> tuple[int, int]:
    """Read image dimensions with Pillow when available, fallback to Tk for PNG/GIF."""
    if Image is not None:
        try:
            with Image.open(path) as img:
                return int(img.width), int(img.height)
        except Exception:
            pass
    return (0, 0)


def run_pixel_pass(vb6_image: Path, dotnet_image: Path) -> dict:
    """Compute image-diff metrics used for screenshot alignment prompts."""
    vb_w, vb_h = image_dimensions(vb6_image)
    net_w, net_h = image_dimensions(dotnet_image)
    metrics = {
        "vb6_size": (vb_w, vb_h),
        "dotnet_size": (net_w, net_h),
        "mae_255": None,
        "mae_pct": None,
        "changed_pct": None,
        "used_pillow": False,
    }

    if not (Image and ImageChops and ImageStat):
        return metrics

    try:
        with Image.open(vb6_image).convert("RGB") as vb_img, Image.open(dotnet_image).convert("RGB") as net_img:
            if vb_img.width <= 0 or vb_img.height <= 0 or net_img.width <= 0 or net_img.height <= 0:
                return metrics

            # Compare at a common resolution to normalize differences and keep this fast.
            common_w = max(64, min(vb_img.width, net_img.width, 640))
            common_h = max(64, min(vb_img.height, net_img.height, 480))
            vb_cmp = vb_img.resize((common_w, common_h), Image.BILINEAR)
            net_cmp = net_img.resize((common_w, common_h), Image.BILINEAR)
            diff = ImageChops.difference(vb_cmp, net_cmp)
            stat = ImageStat.Stat(diff)
            mae = float(sum(stat.mean) / 3.0)
            gray = diff.convert("L")
            hist = gray.histogram()
            changed = sum(hist[21:])  # ignore tiny anti-aliased differences
            total = max(1, sum(hist))
            metrics["mae_255"] = round(mae, 3)
            metrics["mae_pct"] = round((mae / 255.0) * 100.0, 3)
            metrics["changed_pct"] = round((changed / float(total)) * 100.0, 3)
            metrics["used_pillow"] = True
    except Exception:
        return metrics

    return metrics


def build_alignment_report(vb6_image: Path, dotnet_image: Path, metrics: dict) -> str:
    vb_w, vb_h = metrics.get("vb6_size", (0, 0))
    net_w, net_h = metrics.get("dotnet_size", (0, 0))
    mae_255 = metrics.get("mae_255")
    mae_pct = metrics.get("mae_pct")
    changed_pct = metrics.get("changed_pct")
    lines = [
        "SCREEN ALIGNMENT REQUEST (pixel pass)",
        f"VB6 screenshot: {vb6_image}",
        f".NET screenshot: {dotnet_image}",
        f"VB6 size: {vb_w}x{vb_h}",
        f".NET size: {net_w}x{net_h}",
    ]
    if mae_255 is not None:
        lines.append(f"Mean absolute pixel diff (0-255): {mae_255}")
    if mae_pct is not None:
        lines.append(f"Mean absolute pixel diff (%): {mae_pct}")
    if changed_pct is not None:
        lines.append(f"Changed pixels (% thresholded): {changed_pct}")
    lines += [
        "",
        "Update the Razor UI to more closely match the VB6 screenshot.",
        "Focus on spacing, alignment, row heights, field widths, and control positioning.",
        "Preserve existing behavior and data logic.",
    ]
    return "\n".join(lines)



def call_codex_convert(
    source_text: str,
    source_path: Path,
    output_name: str,
    model: str,
    temperature: float,
    instructions: str = "",
    context_blob: str = "",
) -> str:

    api_key = os.environ.get("OPENAI_API_KEY") or os.environ.get("CODEX_API_KEY")
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY or CODEX_API_KEY in environment.")

    endpoint = os.environ.get("CODEX_ENDPOINT", "https://api.openai.com/v1/responses")

    system_prompt = (
        "You are Codex. Convert VB6 code to a .NET Blazor component.\n"
        "Return only the .razor file contents with proper Blazor syntax and @code block.\n"
        "Prefer idiomatic C# and avoid extra explanations."
    )
    user_prompt = (
        f"Convert the following VB6 source file into a Blazor component.\n"
        f"Source path: {source_path}\n"
        f"Output file name: {output_name}\n\n"
        "VB6 SOURCE:\n"
        "```vb\n"
        f"{source_text}\n"
        "```\n"
    )
    if instructions.strip():
        user_prompt += f"\nINSTRUCTIONS:\n{instructions.strip()}\n"
    if context_blob.strip():
        user_prompt += f"\n{context_blob.strip()}\n"

    payload = {
        "model": model,
        "input": [
            {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
            {"role": "user", "content": [{"type": "input_text", "text": user_prompt}]},
        ],
    }
    model_lower = model.lower()
    if "codex-5.2" in model_lower or "gpt-5.2" in model_lower:
        payload["reasoning"] = {"effort": "high"}
    else:
        payload["temperature"] = temperature

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Codex API error: {exc.code} {exc.reason}\n{detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Codex API error: {exc}") from exc

    text = extract_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Codex API returned empty output.")
    return text


def call_codex_fix(blazor_text: str, errors: str, source_path: Path, model: str, temperature: float) -> str:
    api_key = os.environ.get("OPENAI_API_KEY") or os.environ.get("CODEX_API_KEY")
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY or CODEX_API_KEY in environment.")

    endpoint = os.environ.get("CODEX_ENDPOINT", "https://api.openai.com/v1/responses")

    system_prompt = (
        "You are Codex. Fix a .NET Blazor component to resolve compile errors.\n"
        "Return only the corrected .razor file contents with proper Blazor syntax and @code block.\n"
        "Do not include explanations."
    )
    user_prompt = (
        f"Fix the following Blazor component based on the compile errors.\n"
        f"Source path: {source_path}\n\n"
        "COMPILE ERRORS:\n"
        "```\n"
        f"{errors}\n"
        "```\n\n"
        "BLAZOR SOURCE:\n"
        "```razor\n"
        f"{blazor_text}\n"
        "```\n"
    )

    payload = {
        "model": model,
        "input": [
            {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
            {"role": "user", "content": [{"type": "input_text", "text": user_prompt}]},
        ],
    }
    model_lower = model.lower()
    if "codex-5.2" in model_lower or "gpt-5.2" in model_lower:
        payload["reasoning"] = {"effort": "high"}
    else:
        payload["temperature"] = temperature

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Codex API error: {exc.code} {exc.reason}\n{detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Codex API error: {exc}") from exc

    text = extract_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Codex API returned empty output.")
    return text


def call_codex_align(
    blazor_text: str,
    source_path: Path,
    alignment_report: str,
    model: str,
    temperature: float,
) -> str:
    api_key = os.environ.get("OPENAI_API_KEY") or os.environ.get("CODEX_API_KEY")
    if not api_key:
        raise RuntimeError("Missing OPENAI_API_KEY or CODEX_API_KEY in environment.")

    endpoint = os.environ.get("CODEX_ENDPOINT", "https://api.openai.com/v1/responses")

    system_prompt = (
        "You are Codex. Align a .NET Blazor UI to match a VB6 screenshot.\n"
        "Return only the corrected .razor file contents.\n"
        "Do not include explanations."
    )
    user_prompt = (
        f"Align this Blazor component to match VB6 layout and spacing.\n"
        f"Source path: {source_path}\n\n"
        f"{alignment_report}\n\n"
        "BLAZOR SOURCE:\n"
        "```razor\n"
        f"{blazor_text}\n"
        "```\n"
    )

    payload = {
        "model": model,
        "input": [
            {"role": "system", "content": [{"type": "input_text", "text": system_prompt}]},
            {"role": "user", "content": [{"type": "input_text", "text": user_prompt}]},
        ],
    }
    model_lower = model.lower()
    if "codex-5.2" in model_lower or "gpt-5.2" in model_lower:
        payload["reasoning"] = {"effort": "high"}
    else:
        payload["temperature"] = temperature

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Codex API error: {exc.code} {exc.reason}\n{detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Codex API error: {exc}") from exc

    text = extract_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Codex API returned empty output.")
    return text


def is_claude_model(model: str) -> bool:
    """Return True when the selected model should be routed to the Anthropic API."""
    m = model.lower()
    return "claude" in m


def normalize_model_name(model: str) -> str:
    """Normalize model aliases to API-friendly names."""
    m = model.strip()
    if "claude" in m.lower():
        m = m.replace("_", "-").replace(".", "-")
    return m


def extract_claude_response_text(payload: dict) -> str:
    """Extract the assistant text from an Anthropic Messages API response."""
    if not isinstance(payload, dict):
        return ""
    content_list = payload.get("content", [])
    parts: List[str] = []
    for block in content_list:
        if isinstance(block, dict) and block.get("type") == "text":
            parts.append(block.get("text", ""))
    return "\n".join(parts)


def call_claude_convert(
    source_text: str,
    source_path: Path,
    output_name: str,
    model: str,
    temperature: float,
    instructions: str = "",
    context_blob: str = "",
) -> str:
    """Call the Anthropic Messages API to convert VB6 source to a Blazor component.

    Each VB6 file is converted in its own independent API call (one session per file).
    The context_blob is included in every call unchanged.
    """

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("Missing ANTHROPIC_API_KEY in environment.")

    endpoint = os.environ.get(
        "ANTHROPIC_ENDPOINT", "https://api.anthropic.com/v1/messages"
    )

    system_prompt = (
        "You are an expert VB6-to-.NET Blazor migration assistant.\n"
        "Convert VB6 code to a .NET Blazor component.\n"
        "Return only the .razor file contents with proper Blazor syntax and @code block.\n"
        "Prefer idiomatic C# and avoid extra explanations."
    )
    user_prompt = (
        f"Convert the following VB6 source file into a Blazor component.\n"
        f"Source path: {source_path}\n"
        f"Output file name: {output_name}\n\n"
        "VB6 SOURCE:\n"
        "```vb\n"
        f"{source_text}\n"
        "```\n"
    )
    if instructions.strip():
        user_prompt += f"\nINSTRUCTIONS:\n{instructions.strip()}\n"
    if context_blob.strip():
        user_prompt += f"\n{context_blob.strip()}\n"

    payload: dict = {
        "model": normalize_model_name(model),
        "max_tokens": 16384,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt},
        ],
    }
    if temperature > 0:
        payload["temperature"] = temperature

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(
            f"Claude API error: {exc.code} {exc.reason}\n{detail}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Claude API error: {exc}") from exc

    text = extract_claude_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Claude API returned empty output.")
    return text


def call_claude_fix(
    blazor_text: str,
    errors: str,
    source_path: Path,
    model: str,
    temperature: float,
) -> str:
    """Call the Anthropic Messages API to fix compile errors in a Blazor component."""

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("Missing ANTHROPIC_API_KEY in environment.")

    endpoint = os.environ.get(
        "ANTHROPIC_ENDPOINT", "https://api.anthropic.com/v1/messages"
    )

    system_prompt = (
        "You are an expert .NET Blazor developer.\n"
        "Fix the Blazor component to resolve compile errors.\n"
        "Return only the corrected .razor file contents with proper Blazor syntax and @code block.\n"
        "Do not include explanations."
    )
    user_prompt = (
        f"Fix the following Blazor component based on the compile errors.\n"
        f"Source path: {source_path}\n\n"
        "COMPILE ERRORS:\n"
        "```\n"
        f"{errors}\n"
        "```\n\n"
        "BLAZOR SOURCE:\n"
        "```razor\n"
        f"{blazor_text}\n"
        "```\n"
    )

    payload: dict = {
        "model": normalize_model_name(model),
        "max_tokens": 16384,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt},
        ],
    }
    if temperature > 0:
        payload["temperature"] = temperature

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(
            f"Claude API error: {exc.code} {exc.reason}\n{detail}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Claude API error: {exc}") from exc

    text = extract_claude_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Claude API returned empty output.")
    return text


def call_claude_align(
    blazor_text: str,
    source_path: Path,
    alignment_report: str,
    model: str,
    temperature: float,
) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("Missing ANTHROPIC_API_KEY in environment.")

    endpoint = os.environ.get(
        "ANTHROPIC_ENDPOINT", "https://api.anthropic.com/v1/messages"
    )

    system_prompt = (
        "You are an expert .NET Blazor UI migration assistant.\n"
        "Align a Blazor page layout to match a VB6 screenshot while preserving behavior.\n"
        "Return only the corrected .razor file contents.\n"
    )
    user_prompt = (
        f"Align this Blazor component to match VB6 layout and spacing.\n"
        f"Source path: {source_path}\n\n"
        f"{alignment_report}\n\n"
        "BLAZOR SOURCE:\n"
        "```razor\n"
        f"{blazor_text}\n"
        "```\n"
    )

    payload: dict = {
        "model": normalize_model_name(model),
        "max_tokens": 16384,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt},
        ],
    }
    if temperature > 0:
        payload["temperature"] = temperature

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(
            f"Claude API error: {exc.code} {exc.reason}\n{detail}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Claude API error: {exc}") from exc

    text = extract_claude_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Claude API returned empty output.")
    return text


def parse_vb6_blocks(source: str) -> List[dict]:
    """Parse a VB6 source file into code blocks delimited by Sub/Function/Property boundaries.

    Returns a list of dicts:
        {"start": int, "end": int, "header": str, "code": str, "comment_start": int|None}
    where start/end are 0-based line indices (inclusive), header is the Sub/Function line,
    and comment_start is the line index where a contiguous comment block immediately
    preceding the header begins (or None if no such block exists).
    """
    lines = source.splitlines()
    block_start_re = re.compile(
        r"^\s*(Public\s+|Private\s+|Friend\s+|Static\s+)?"
        r"(Sub|Function|Property\s+(Get|Let|Set))\s+([A-Za-z_]\w*)",
        re.IGNORECASE,
    )
    block_end_re = re.compile(
        r"^\s*End\s+(Sub|Function|Property)", re.IGNORECASE
    )
    blocks: List[dict] = []
    i = 0
    while i < len(lines):
        m = block_start_re.match(lines[i])
        if m:
            header_line = i
            # Find contiguous comment block immediately above the header
            comment_start = None
            j = header_line - 1
            while j >= 0:
                stripped = lines[j].lstrip()
                if stripped.startswith("'") or stripped.lower().startswith("rem "):
                    comment_start = j
                    j -= 1
                elif stripped == "":
                    j -= 1  # skip blank lines between comments and header
                else:
                    break
            # Find the End Sub/Function/Property
            end_line = i
            for k in range(i + 1, len(lines)):
                if block_end_re.match(lines[k]):
                    end_line = k
                    break
            else:
                end_line = len(lines) - 1
            block_code = "\n".join(lines[i : end_line + 1])
            blocks.append({
                "start": comment_start if comment_start is not None else header_line,
                "header_line": header_line,
                "end": end_line,
                "header": lines[i].strip(),
                "code": block_code,
                "comment_start": comment_start,
            })
            i = end_line + 1
        else:
            i += 1
    return blocks


def call_claude_tag(
    source_text: str,
    source_path: Path,
    model: str,
    context_blob: str = "",
    migrated_source: str = "",
) -> str:
    """Call the Anthropic Messages API to tag VB6 code blocks and form elements.

    For each Sub/Function/Property block AND each Form UI Element (Begin block) the LLM writes:
      1. A short description of what the block does / what the control is.
      2. Up to 5 lines of .NET conversion instructions.
    Returns the full file content with comment blocks inserted/replaced.
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("Missing ANTHROPIC_API_KEY in environment.")

    endpoint = os.environ.get(
        "ANTHROPIC_ENDPOINT", "https://api.anthropic.com/v1/messages"
    )

    system_prompt = (
        "You are an expert VB6 code analyst preparing files for .NET Blazor migration.\n\n"
        "=== CODE BLOCKS ===\n"
        "For EVERY Sub, Function, and Property block:\n"
        "1. Write a one-line description of what the block does.\n"
        "2. Write up to 5 short instruction lines for converting this block to .NET Blazor.\n"
        "Format each block's tag as a VB6 comment block placed IMMEDIATELY before the Sub/Function/Property line:\n"
        "' TAG: <description>\n"
        "' CONVERT: <instruction 1>\n"
        "' CONVERT: <instruction 2>\n"
        "... (up to 5 CONVERT lines)\n\n"
        "=== FORM UI ELEMENTS ===\n"
        "For EVERY Form UI Element (lines starting with 'Begin <ControlType> <ControlName>'),\n"
        "add a tag comment block IMMEDIATELY before the Begin line:\n"
        "' TAG: <what this control is and its purpose>\n"
        "' CONVERT: <instruction 1>\n"
        "' CONVERT: <instruction 2>\n"
        "... (up to 5 CONVERT lines)\n"
        "Common VB6 controls and their Blazor equivalents:\n"
        "  - VSFlex8Ctl.VSFlexGrid / MSFlexGridLib.MSFlexGrid -> HfGrid with HfGridColumn components\n"
        "    Use <HfGrid TValue=\"RowType\"> with <HfGridColumnsBase> containing <HfGridColumn> for each column.\n"
        "    Load column layout via MMain.IniGetGridAsync(formName, gridName, 0, staticNames: true).\n"
        "  - VB.TextBox -> <input> or <InputText> with @bind\n"
        "  - VB.ComboBox / VBCombo -> <select> or <InputSelect> with @bind\n"
        "  - VB.CheckBox -> <input type=\"checkbox\"> with @bind\n"
        "  - VB.CommandButton -> <button @onclick=\"Handler\">\n"
        "  - VB.Label -> <label> or <span>\n"
        "  - ComctlLib.Toolbar -> Toolbar div with <button> elements\n"
        "  - VB.ListBox -> <select multiple> or custom list component\n"
        "  - TabDlg.SSTab / ComctlLib.TabStrip -> Tab component or CSS tabs\n"
        "  - VB.Frame -> <div> or <fieldset> grouping\n"
        "  - VB.PictureBox -> <img> or container div\n\n"
        "=== .NET BLAZOR CONVERSION PATTERNS ===\n"
        "Database: Use DbWrapperSqlServer with QueryAsync<T>(), ExecuteAsync(), QueryFirstOrDefaultAsync<T>()\n"
        "  Example: var rows = await Db.QueryAsync<MyRow>(\"SELECT * FROM Table WHERE Id=@Id\", new { Id = id });\n"
        "Grid Events: Use HfGridEvents<T> with RowSelected, OnCellSave callbacks.\n"
        "Grid Columns: Use HfGridColumn with Field, HeaderText, Width, Type, Format, AllowEditing, Visible.\n"
        "State: Use private fields + StateHasChanged() for UI updates.\n"
        "Modal Dialogs: Use _isVisible bool + overlay div pattern with Open()/Close() methods.\n"
        "Navigation: Use NavigationManager.NavigateTo() instead of Form.Show.\n\n"
        "=== RULES ===\n"
        "Return the COMPLETE file with all original code intact and the tag comments inserted.\n"
        "Do NOT change any code. Only add comment blocks before each Sub/Function/Property and each Begin block.\n"
        "If a TAG comment block already exists before a block, replace it with the new one.\n"
        "Preserve all other existing comments exactly as they are."
    )
    user_prompt = (
        f"Tag every Sub/Function/Property block AND every Form UI Element (Begin block) in this VB6 file.\n"
        f"Source path: {source_path}\n\n"
        "VB6 SOURCE:\n"
        "```vb\n"
        f"{source_text}\n"
        "```\n"
    )
    if migrated_source.strip():
        user_prompt += (
            "\nMIGRATED BLAZOR FILE (for reference — shows conversion patterns already used):\n"
            "```razor\n"
            f"{migrated_source.strip()}\n"
            "```\n"
        )
    if context_blob.strip():
        user_prompt += f"\n{context_blob.strip()}\n"

    payload: dict = {
        "model": normalize_model_name(model),
        "max_tokens": 16384,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.2,
    }

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=data,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=600) as resp:
            resp_payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(
            f"Claude API error: {exc.code} {exc.reason}\n{detail}"
        ) from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Claude API error: {exc}") from exc

    text = extract_claude_response_text(resp_payload)
    if not text.strip():
        raise RuntimeError("Claude API returned empty output.")

    # Strip markdown code fences if the LLM wrapped its response
    stripped = text.strip()
    if stripped.startswith("```"):
        first_nl = stripped.index("\n") if "\n" in stripped else len(stripped)
        stripped = stripped[first_nl + 1 :]
        if stripped.endswith("```"):
            stripped = stripped[: -3].rstrip()
        text = stripped

    return text


def ensure_csproj_includes(csproj_path: Path, file_path: Path) -> None:
    if not csproj_path.exists():
        raise RuntimeError(f"Missing csproj at {csproj_path}")

    rel = file_path.relative_to(csproj_path.parent).as_posix().replace("/", "\\")
    csproj_text = csproj_path.read_text(encoding="utf-8", errors="ignore")
    if rel in csproj_text:
        return

    insert = f"""  <ItemGroup>\n    <Content Include=\"{rel}\" />\n  </ItemGroup>\n"""

    if "</Project>" in csproj_text:
        csproj_text = csproj_text.replace("</Project>", insert + "\n</Project>")
    else:
        csproj_text += "\n" + insert

    csproj_path.write_text(csproj_text, encoding="utf-8")

def load_hfsystem_components(hf_system_dir: Path, max_files: int | None = None) -> List[Component]:
    vbp = hf_system_dir / "HFSystem.vbp"
    if not vbp.exists():
        raise SystemExit("HFSystem.vbp not found")
    lines = read_text(vbp).splitlines()
    comps: List[Component] = []
    count = 0
    for line in lines:
        if max_files is not None and count >= max_files:
            break
        line = line.strip()
        if not line:
            continue
        if line.startswith("Form="):
            rel = normalize_rel(line.split("=", 1)[1].strip().strip('"'))
            path = (hf_system_dir / rel).resolve()
            if path.exists():
                comps.append(Component("HFSystem", get_vb_name(path), path, "Form"))
                count += 1
        elif line.startswith("Class="):
            body = line.split("=", 1)[1]
            parts = [p.strip() for p in body.split(";")]
            if len(parts) == 2:
                rel = normalize_rel(parts[1].strip().strip('"'))
                path = (hf_system_dir / rel).resolve()
                if path.exists():
                    comps.append(Component("HFSystem", get_vb_name(path), path, "Class"))
                    count += 1
        elif line.startswith("Module="):
            body = line.split("=", 1)[1]
            parts = [p.strip() for p in body.split(";")]
            if len(parts) == 2:
                rel = normalize_rel(parts[1].strip().strip('"'))
                path = (hf_system_dir / rel).resolve()
                if path.exists():
                    comps.append(Component("HFSystem", get_vb_name(path), path, "Module"))
                    count += 1
    return comps


def load_hfest_components(hf_est_dir: Path, max_files: int | None = None) -> List[Component]:
    comps: List[Component] = []
    count = 0
    for ext, ctype in [(".frm", "Form"), (".bas", "Module"), (".cls", "Class")]:
        for path in hf_est_dir.glob(f"*{ext}"):
            comps.append(Component("HFEst", get_vb_name(path), path, ctype))
            count += 1
            if max_files is not None and count >= max_files:
                return comps
    return comps


def sanitize_project_name(name: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9._-]+", "_", (name or "").strip())
    clean = clean.strip("._-")
    return clean or "Project"


def iter_matching_files(root: Path, allowed_extensions: set[str]) -> list[Path]:
    if not root.exists():
        return []
    files: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() in allowed_extensions:
            files.append(path)
    files.sort(key=lambda p: str(p).lower())
    return files


def copy_matching_files(source_root: Path, dest_root: Path, allowed_extensions: set[str]) -> int:
    source_root = source_root.expanduser().resolve()
    dest_root = dest_root.expanduser().resolve()
    if not source_root.exists():
        raise RuntimeError(f"Source folder not found: {source_root}")
    if dest_root == source_root:
        raise RuntimeError("Destination folder must be different from source folder.")
    if str(dest_root).startswith(str(source_root) + os.sep):
        raise RuntimeError("Destination folder cannot be inside source folder.")

    copied = 0
    for src in iter_matching_files(source_root, allowed_extensions):
        rel = src.relative_to(source_root)
        dst = dest_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        copied += 1
    return copied


def copy_all_files(source_root: Path, dest_root: Path) -> int:
    source_root = source_root.expanduser().resolve()
    dest_root = dest_root.expanduser().resolve()
    if not source_root.exists():
        return 0
    if dest_root == source_root:
        return 0
    if str(dest_root).startswith(str(source_root) + os.sep):
        raise RuntimeError("Destination folder cannot be inside source folder.")

    copied = 0
    for src in source_root.rglob("*"):
        if not src.is_file():
            continue
        rel = src.relative_to(source_root)
        dst = dest_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        copied += 1
    return copied


def read_project_manifest(project_dir: Path) -> dict:
    manifest_path = project_dir / PROJECT_MANIFEST_NAME
    if not manifest_path.exists():
        return {}
    try:
        payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return payload if isinstance(payload, dict) else {}


def to_project_relative(base_dir: Path, path: Path) -> str:
    try:
        return path.resolve().relative_to(base_dir.resolve()).as_posix()
    except Exception:
        return str(path.resolve())


def resolve_project_path(project_dir: Path, raw_value: str | None, default_name: str) -> Path:
    raw = (raw_value or "").strip()
    if not raw:
        return (project_dir / default_name).resolve()
    parsed = Path(raw).expanduser()
    if parsed.is_absolute():
        return parsed.resolve()
    return (project_dir / parsed).resolve()


def write_project_manifest(
    project_dir: Path,
    *,
    name: str,
    source_dir: Path,
    target_dir: Path,
    cache_dir: Path,
    graphs_dir: Path,
    output_project: str = "",
    source_origin: str = "",
    target_origin: str = "",
) -> None:
    project_dir = project_dir.expanduser().resolve()
    payload = {
        "version": 1,
        "name": name,
        "source_dir": to_project_relative(project_dir, source_dir),
        "target_dir": to_project_relative(project_dir, target_dir),
        "cache_dir": to_project_relative(project_dir, cache_dir),
        "graphs_dir": to_project_relative(project_dir, graphs_dir),
        "output_project": output_project,
        "source_origin": source_origin,
        "target_origin": target_origin,
        "updated_at_utc": datetime.now(timezone.utc).isoformat(),
    }
    manifest_path = project_dir / PROJECT_MANIFEST_NAME
    manifest_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def get_project_layout(project_dir: Path) -> dict:
    project_dir = project_dir.expanduser().resolve()
    project_dir.mkdir(parents=True, exist_ok=True)
    manifest = read_project_manifest(project_dir)
    source_dir = resolve_project_path(project_dir, manifest.get("source_dir"), PROJECT_SOURCE_DIR)
    target_dir = resolve_project_path(project_dir, manifest.get("target_dir"), PROJECT_TARGET_DIR)
    cache_dir = resolve_project_path(project_dir, manifest.get("cache_dir"), PROJECT_CACHE_DIR)
    graphs_dir = resolve_project_path(project_dir, manifest.get("graphs_dir"), PROJECT_GRAPHS_DIR)
    for folder in (source_dir, target_dir, cache_dir, graphs_dir):
        folder.mkdir(parents=True, exist_ok=True)
    output_project = (manifest.get("output_project") or "").strip()
    if output_project:
        output_path = Path(output_project).expanduser()
        if not output_path.is_absolute():
            output_path = (project_dir / output_path).resolve()
        output_project = str(output_path)
    return {
        "project_dir": project_dir,
        "name": (manifest.get("name") or project_dir.name).strip() or project_dir.name,
        "source_dir": source_dir,
        "target_dir": target_dir,
        "cache_dir": cache_dir,
        "graphs_dir": graphs_dir,
        "output_project": output_project,
        "source_origin": (manifest.get("source_origin") or "").strip(),
        "target_origin": (manifest.get("target_origin") or "").strip(),
    }


def count_component_files(source_dir: Path) -> int:
    return len(iter_matching_files(source_dir, set(VB6_COMPONENT_TYPES.keys())))


def create_project_package(
    projects_root: Path,
    project_name: str,
    vb6_source_root: Path,
    output_project: str = "",
    seed_target_dir: Path | None = None,
    overwrite: bool = False,
) -> tuple[Path, int, int]:
    projects_root = projects_root.expanduser().resolve()
    projects_root.mkdir(parents=True, exist_ok=True)

    normalized_name = sanitize_project_name(project_name)
    project_dir = (projects_root / normalized_name).resolve()
    if project_dir.exists() and any(project_dir.iterdir()) and not overwrite:
        raise FileExistsError(f"Project folder already exists: {project_dir}")
    if overwrite and project_dir.exists():
        shutil.rmtree(project_dir)

    source_dir = project_dir / PROJECT_SOURCE_DIR
    target_dir = project_dir / PROJECT_TARGET_DIR
    cache_dir = project_dir / PROJECT_CACHE_DIR
    graphs_dir = project_dir / PROJECT_GRAPHS_DIR
    for folder in (source_dir, target_dir, cache_dir, graphs_dir):
        folder.mkdir(parents=True, exist_ok=True)

    copied_source = copy_matching_files(
        vb6_source_root,
        source_dir,
        allowed_extensions=VB6_COPY_EXTENSIONS,
    )
    if copied_source == 0:
        raise RuntimeError(f"No VB6 files found under {vb6_source_root}")

    copied_target = 0
    if seed_target_dir:
        copied_target = copy_all_files(seed_target_dir, target_dir)

    write_project_manifest(
        project_dir,
        name=normalized_name,
        source_dir=source_dir,
        target_dir=target_dir,
        cache_dir=cache_dir,
        graphs_dir=graphs_dir,
        output_project=output_project,
        source_origin=str(vb6_source_root.expanduser().resolve()),
        target_origin=str(seed_target_dir.expanduser().resolve()) if seed_target_dir else "",
    )
    return project_dir, copied_source, copied_target


def ensure_default_project_package(
    projects_root: Path,
    default_name: str,
    vb6_source_root: Path,
    output_project: str = "",
    seed_target_dir: Path | None = None,
) -> Path:
    projects_root = projects_root.expanduser().resolve()
    project_dir = (projects_root / sanitize_project_name(default_name)).resolve()
    if project_dir.exists():
        layout = get_project_layout(project_dir)
        if count_component_files(layout["source_dir"]) > 0:
            return layout["project_dir"]
    created_dir, _, _ = create_project_package(
        projects_root=projects_root,
        project_name=default_name,
        vb6_source_root=vb6_source_root,
        output_project=output_project,
        seed_target_dir=seed_target_dir,
        overwrite=project_dir.exists(),
    )
    return created_dir


def list_existing_project_dirs(projects_root: Path, include_default: bool = True) -> list[Path]:
    projects_root = projects_root.expanduser().resolve()
    projects_root.mkdir(parents=True, exist_ok=True)
    seen: set[str] = set()
    dirs: list[Path] = []
    for child in sorted(projects_root.iterdir(), key=lambda p: p.name.lower()):
        if not child.is_dir():
            continue
        if child.name.startswith("__"):
            continue
        key = str(child.resolve())
        if key in seen:
            continue
        dirs.append(child.resolve())
        seen.add(key)

    if include_default:
        default_dir = (projects_root / sanitize_project_name(DEFAULT_PROJECT_NAME)).resolve()
        key = str(default_dir)
        if key not in seen:
            dirs.insert(0, default_dir)

    return dirs


def ensure_empty_project(projects_root: Path) -> Path:
    projects_root = projects_root.expanduser().resolve()
    project_dir = (projects_root / EMPTY_PROJECT_NAME).resolve()
    source_dir = project_dir / PROJECT_SOURCE_DIR
    target_dir = project_dir / PROJECT_TARGET_DIR
    cache_dir = project_dir / PROJECT_CACHE_DIR
    graphs_dir = project_dir / PROJECT_GRAPHS_DIR

    for folder in (project_dir, source_dir, target_dir, cache_dir, graphs_dir):
        folder.mkdir(parents=True, exist_ok=True)

    # A closed project should always reopen with no source rows.
    shutil.rmtree(source_dir, ignore_errors=True)
    source_dir.mkdir(parents=True, exist_ok=True)

    write_project_manifest(
        project_dir,
        name="No Project",
        source_dir=source_dir,
        target_dir=target_dir,
        cache_dir=cache_dir,
        graphs_dir=graphs_dir,
        output_project="",
        source_origin="",
        target_origin="",
    )
    return project_dir


def load_project_components(
    source_dir: Path,
    *,
    fallback_project_name: str,
    max_files: int | None = None,
) -> List[Component]:
    components: List[Component] = []
    count = 0
    component_exts = set(VB6_COMPONENT_TYPES.keys())
    for path in iter_matching_files(source_dir, component_exts):
        rel = path.relative_to(source_dir)
        project_name = rel.parts[0] if len(rel.parts) > 1 else fallback_project_name
        ext = path.suffix.lower()
        ctype = VB6_COMPONENT_TYPES.get(ext, "VB6")
        if ext in {".bas", ".cls", ".frm", ".ctl", ".ctrl"}:
            comp_name = get_vb_name(path)
        else:
            comp_name = path.stem
        components.append(Component(project_name, comp_name, path.resolve(), ctype))
        count += 1
        if max_files is not None and count >= max_files:
            break
    return components


def normalize_lines(path: Path) -> List[str]:
    txt = read_text(path)
    lines = []
    for line in txt.splitlines():
        s = line.strip()
        if s:
            lines.append(s)
    return lines


def dedupe_by_content(
    components: List[Component],
    similarity_threshold: float = 0.9,
    prefer_hfest: bool = True,
) -> tuple[list[Component], list[Component]]:
    max_fuzzy_lines = 2000
    infos = []
    for comp in components:
        lines = normalize_lines(comp.path)
        h = hashlib.sha1()
        for line in lines:
            h.update(line.encode("utf-8", errors="ignore"))
            h.update(b"\n")
        infos.append(
            {
                "comp": comp,
                "lines": lines,
                "line_count": len(lines),
                "ext": comp.path.suffix.lower(),
                "hash": h.hexdigest(),
            }
        )
    n = len(infos)
    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    ext_groups: Dict[str, List[int]] = {}
    for idx, info in enumerate(infos):
        ext_groups.setdefault(info["ext"], []).append(idx)

    for _ext, idxs in ext_groups.items():
        idxs.sort(key=lambda i: infos[i]["line_count"])
        # Exact-match union by content hash
        hash_groups: Dict[str, List[int]] = {}
        for idx in idxs:
            hash_groups.setdefault(infos[idx]["hash"], []).append(idx)
        for group in hash_groups.values():
            if len(group) > 1:
                root = group[0]
                for other in group[1:]:
                    union(root, other)

        for i, ia in enumerate(idxs):
            a = infos[ia]
            a_count = a["line_count"]
            if a_count == 0 or a_count > max_fuzzy_lines:
                continue
            max_count = int(a_count * 1.25)
            for ib in idxs[i + 1 :]:
                b = infos[ib]
                b_count = b["line_count"]
                if b_count > max_count:
                    break
                if b_count == 0 or b_count > max_fuzzy_lines:
                    continue
                sm = difflib.SequenceMatcher(None, a["lines"], b["lines"])
                if sm.quick_ratio() < similarity_threshold:
                    continue
                ratio = sm.ratio()
                if ratio >= similarity_threshold:
                    union(ia, ib)

    clusters: Dict[int, List[int]] = {}
    for idx in range(n):
        clusters.setdefault(find(idx), []).append(idx)

    def rank(idx: int) -> tuple:
        info = infos[idx]
        comp = info["comp"]
        prefer = 0 if (prefer_hfest and comp.project == "HFEst") else 1
        return (-info["line_count"], prefer, str(comp.path))

    winners: List[Component] = []
    dropped: List[Component] = []
    for members in clusters.values():
        if len(members) == 1:
            winners.append(infos[members[0]]["comp"])
            continue
        members.sort(key=rank)
        winner = members[0]
        winners.append(infos[winner]["comp"])
        for idx in members[1:]:
            dropped.append(infos[idx]["comp"])

    return winners, dropped


def build_indices(components: List[Component], prefer_hfest: bool):
    by_name: Dict[str, List[Component]] = {}
    for c in components:
        by_name.setdefault(c.name, []).append(c)

    duplicates = {
        name for name, comps in by_name.items()
        if len({c.project for c in comps}) > 1
    }

    preferred: Dict[str, Component] = {}
    for name, comps in by_name.items():
        if prefer_hfest:
            hfest = [c for c in comps if c.project == "HFEst"]
            preferred[name] = hfest[0] if hfest else comps[0]
        else:
            preferred[name] = comps[0]

    return by_name, preferred, duplicates


def build_edges(components: List[Component], preferred: Dict[str, Component], duplicates: set):
    edges: Dict[Tuple[str, str], Dict] = {}
    for comp in components:
        caller_id = f"{comp.project}::{comp.name}"
        txt = read_text(comp.path)
        for line in txt.splitlines():
            line_strip = line.lstrip()
            if line_strip.startswith("'"):
                continue
            for label, rx in [("Load", RX_LOAD), ("Unload", RX_UNLOAD), ("Set", RX_SET), ("New", RX_NEW), ("Dot", RX_DOT)]:
                for m in rx.finditer(line_strip):
                    ref = m.group(1)
                    if ref == comp.name:
                        continue
                    if ref not in preferred:
                        continue
                    t = preferred[ref]
                    callee_id = f"{t.project}::{t.name}"
                    key = (caller_id, callee_id)
                    if key not in edges:
                        edges[key] = {"types": set(), "ambiguous": False}
                    edges[key]["types"].add(label)
                    if ref in duplicates:
                        edges[key]["ambiguous"] = True
    return edges


def build_roots(components: List[Component]) -> List[str]:
    roots = []
    for comp in components:
        if SUB_MAIN_RE.search(read_text(comp.path)):
            roots.append(f"{comp.project}::{comp.name}")
    return sorted(set(roots))


def write_outputs(out_dir: Path, edges: Dict[Tuple[str, str], Dict], components: List[Component],
                  roots: List[str], suffix: str) -> Tuple[Path, Path, Path,List[str],set[str],set[str]]:
    out_dir = out_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    id_to_file = {f"{c.project}::{c.name}": str(c.path) for c in components}

    csv_path = out_dir / f"vb6_callgraph_{suffix}.csv"
    with csv_path.open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["caller", "callee", "types", "ambiguous", "caller_file", "callee_file"])
        for (caller, callee), meta in sorted(edges.items()):
            writer.writerow([
                caller, callee, ";".join(sorted(meta["types"])),
                "true" if meta.get("ambiguous") else "false",
                id_to_file.get(caller, ""), id_to_file.get(callee, "")
            ])

    mmd_path = out_dir / f"vb6_callgraph_{suffix}.mmd"

    def safe_id(node_id: str) -> str:
        return re.sub(r"[^A-Za-z0-9_]", "_", node_id)

    with mmd_path.open("w") as f:
        f.write("graph LR\n")
        for (caller, callee), meta in sorted(edges.items()):
            label = ",".join(sorted(meta["types"]))
            if meta.get("ambiguous"):
                label += " (ambig)"
            f.write(f'  {safe_id(caller)}["{caller}"] -->|"{label}"| {safe_id(callee)}["{callee}"]\n')

    # Full-depth tree
    adj: Dict[str, List[Tuple[str, bool]]] = {}
    for (caller, callee), meta in edges.items():
        adj.setdefault(caller, []).append((callee, meta.get("ambiguous", False)))
    for k in adj:
        adj[k] = sorted(set(adj[k]), key=lambda x: x[0])

    migration_order, cycle_set = compute_migration_order(components, edges)
    comp_ids = [f"{c.project}::{c.name}" for c in components]
    # Orphans: not referenced by any other program (exclude entry points).
    entry_points = set(roots)
    if "HFSystem::FLogin" in adj:
        entry_points.add("HFSystem::FLogin")
    inbound_by_name: Dict[str, int] = {}
    for (_caller, callee), _meta in edges.items():
        name = callee.split("::", 1)[-1]
        inbound_by_name[name] = inbound_by_name.get(name, 0) + 1
    orphan_set = set()
    comp_by_id = {f"{c.project}::{c.name}": c for c in components}
    for cid in comp_ids:
        if cid in entry_points:
            continue
        comp = comp_by_id.get(cid)
        if not comp:
            continue
        if comp.path.suffix.lower() != ".frm":
            continue
        name = cid.split("::", 1)[-1]
        if inbound_by_name.get(name, 0) == 0:
            orphan_set.add(cid)

    def tree_lines_full(start: str) -> List[str]:
        lines: List[str] = []
        stack_path: List[str] = []

        def rec(node: str, prefix: str):
            lines.append(prefix + node)
            stack_path.append(node)
            for child, ambig in adj.get(node, []):
                label = child + (" [ambig]" if ambig else "")
                if child in stack_path:
                    lines.append(prefix + "  " + label + " [cycle]")
                    continue
                rec(label, prefix + "  ")
            stack_path.pop()

        rec(start, "")
        return lines

    text_path = out_dir / f"vb6_calltree_{suffix}.txt"
    with text_path.open("w") as f:
        f.write("Roots (Sub Main): " + ", ".join(roots) + "\n\n")
        if "HFSystem::FLogin" in adj:
            f.write("FLogin tree (full depth):\n")
            f.write("\n".join(tree_lines_full("HFSystem::FLogin")) + "\n\n")
        if "HFEst::MMain" in adj:
            f.write("MMain tree (full depth):\n")
            f.write("\n".join(tree_lines_full("HFEst::MMain")) + "\n")

    return csv_path, mmd_path, text_path, migration_order,cycle_set,orphan_set


def run_ui(
    components: List[Component],
    edges: Dict[Tuple[str, str], Dict],
    roots: List[str],
    duplicates: set,
    ui_defaults: Dict[str, str],
    migration_order: List[str],
    cycle_set: set[str],
    orphan_set: set[str]
):
    import time
    import tkinter as tk
    from tkinter import ttk
    from tkinter import font as tkfont
    from tkinter import filedialog, messagebox, simpledialog

    t_start = time.perf_counter()
    chunk_times: list[tuple[str, float]] = []

    def mark_chunk(name: str, last_time: list[float]) -> None:
        now = time.perf_counter()
        duration = now - last_time[0]
        chunk_times.append((name, duration))
        print(f"[timing] {name}: {duration:.3f}s")
        last_time[0] = now

    last_time = [t_start]

    id_to_component = {f"{c.project}::{c.name}": c for c in components}
    comp_id_by_path = {str(c.path.resolve()): f"{c.project}::{c.name}" for c in components}
    comp_id_by_name = {c.name.lower(): f"{c.project}::{c.name}" for c in components}
    cache_dir = Path(ui_defaults.get("cache_dir", "") or "").expanduser()
    cache_index = load_cache_index(cache_dir)
    files_index = cache_index.setdefault("files", {})
    if "chunks_enabled" not in cache_index:
        cache_index["chunks_enabled"] = True
        save_cache_index(cache_dir, cache_index)
    valid_paths = {str(c.path) for c in components if c.path.exists()}
    invalid_keys = [k for k in list(files_index.keys()) if k not in valid_paths or not Path(k).exists()]
    if invalid_keys:
        for k in invalid_keys:
            files_index.pop(k, None)
        save_cache_index(cache_dir, cache_index)
    source_cache: Dict[str, str] = {}
    blazor_cache: Dict[str, str] = {}
    pre_output_dir = Path(ui_defaults.get("output_dir", "") or "").expanduser()
    total_lines = 0
    line_count_by_id: Dict[str, int | None] = {}
    file_size_by_id: Dict[str, int] = {}
    type_set = set()
    cache_dirty = False
    for comp in components:
        comp_id = f"{comp.project}::{comp.name}"
        try:
            stat = comp.path.stat()
            file_size_by_id[comp_id] = stat.st_size
            entry = files_index.get(str(comp.path))
            if entry and entry.get("mtime") == stat.st_mtime and entry.get("size") == stat.st_size:
                line_count = entry.get("line_count")
                line_count_by_id[comp_id] = line_count
                if isinstance(line_count, int):
                    total_lines += line_count
            else:
                line_count = count_lines_file(comp.path)
                line_count_by_id[comp_id] = line_count
                total_lines += line_count
                files_index[str(comp.path)] = {
                    "mtime": stat.st_mtime,
                    "size": stat.st_size,
                    "line_count": line_count,
                }
                cache_dirty = True
        except Exception:
            file_size_by_id[comp_id] = 0
            line_count_by_id[comp_id] = None
        ext = comp.path.suffix.lower() or "unknown"
        type_set.add(ext)
    if cache_dirty:
        save_cache_index(cache_dir, cache_index)
    total_files = len(components)
    program_types = sorted(type_set)
    mark_chunk("chunk1_cache_and_counts", last_time)

    adj: Dict[str, List[Tuple[str, bool]]] = {}
    for (caller, callee), meta in edges.items():
        adj.setdefault(caller, []).append((callee, meta.get("ambiguous", False)))
    for k in adj:
        adj[k] = sorted(set(adj[k]), key=lambda x: x[0])
    mark_chunk("chunk2_adj_build", last_time)

    root = tk.Tk()
    project_title = ui_defaults.get("project_name", "").strip()
    if project_title:
        root.title(f"VB6 Call Tree Migrator - {project_title}")
    else:
        root.title("VB6 Call Tree Migrator")
    root.withdraw()
    root.update_idletasks()
    # 1/2 of 4K (3840x2160) => 1920x1080, start at top-left
    root.geometry("1920x1080+0+0")
    mark_chunk("chunk3_root_window", last_time)

    # Configuration state
    root_var = tk.StringVar(value=ui_defaults.get("root", ""))
    project_var = tk.StringVar(value=ui_defaults.get("project", ""))
    output_dir_var = tk.StringVar(value=ui_defaults.get("output_dir", ""))
    output_name_var = tk.StringVar(value=ui_defaults.get("output_name", ""))
    context_var = tk.StringVar(
        value=ui_defaults.get(
            "context_dir",
            "/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/Components/Pages/Migrated/ContextDir/",
        )
    )
    project_root_guess = Path(project_var.get()).expanduser().parent if project_var.get().strip() else Path.home()
    vb6_image_dir_var = tk.StringVar(
        value=ui_defaults.get(
            "vb6_image_dir",
            str(project_root_guess / "wwwroot" / "images" / "vb6"),
        )
    )
    dotnet_image_dir_var = tk.StringVar(
        value=ui_defaults.get(
            "dotnet_image_dir",
            str(project_root_guess / "wwwroot" / "images" / "dotnet"),
        )
    )
    search_var = tk.StringVar(value="")
    instructions_default = ui_defaults.get(
        "instructions",
        "Convert teh attached VB6 files to .NET Blazor using the HF Controls. Employ and DBWrapperSqlServer.cs for database queries.",
    )
    model_var = tk.StringVar(value=ui_defaults.get("model", "claude-opus-4-6"))
    temp_var = tk.StringVar(value=ui_defaults.get("temperature", "0.2"))
    prefer_var = tk.BooleanVar(value=ui_defaults.get("prefer_hfest", True))
    status_var = tk.StringVar(value="Ready")
    last_error_var = tk.StringVar(value="")
    ui_action = {"action": "exit", "project_dir": ""}
    current_output_path: Path | None = None
    blazor_header_lines = 0
    converted_map: Dict[str, Path] = {}
    run_process: subprocess.Popen | None = None
    run_project_path: Path | None = None
    run_log_path: Path | None = None
    run_tail_offset = 0
    run_tail_buffer = ""
    run_log_header_lines: List[str] = []
    run_url_announced = ""
    run_exit_announced = False
    run_restart_count = 0
    run_max_restarts = 2
    run_session_id = ""
    tail_job_id: str | None = None
    error_panel_visible = False
    error_panel_mode = "build"  # build | run
    run_browser = BlazorRunBrowser()

    instructions_text = tk.Text(root, height=4, wrap="word")
    instructions_text.insert("1.0", instructions_default)

    # Top bar (status + search)
    controls = ttk.Frame(root)
    controls.pack(fill=tk.X, padx=8, pady=6)

    ttk.Label(controls, text="Search").grid(row=0, column=0, sticky="w")
    search_entry = ttk.Entry(controls, textvariable=search_var, width=40)
    search_entry.grid(row=0, column=1, sticky="w", padx=(6, 12))
    search_btn = ttk.Button(controls, text="Search")
    search_btn.grid(row=0, column=2, sticky="w")

    status_entry = ttk.Entry(controls, textvariable=status_var, state="readonly")
    status_entry.grid(row=0, column=3, sticky="we", padx=(12, 6))

    def on_copy_error():
        err = last_error_var.get()
        if not err:
            status_var.set("No error to copy.")
            return
        root.clipboard_clear()
        root.clipboard_append(err)
        status_var.set("Error copied to clipboard.")

    copy_error_btn = ttk.Button(controls, text="Copy Error", command=on_copy_error)
    copy_error_btn.grid(row=0, column=4, sticky="e")

    controls.columnconfigure(3, weight=1)

    def open_config_dialog():
        dlg = tk.Toplevel(root)
        dlg.title("Configuration")
        dlg.geometry("900x600")
        dlg.transient(root)

        frame = ttk.Frame(dlg, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        def browse_dir(var: tk.StringVar):
            path = filedialog.askdirectory(initialdir=var.get() or "/")
            if path:
                var.set(path)

        row = 0
        ttk.Label(frame, text="VB6 Root").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=root_var, width=60).grid(row=row, column=1, sticky="we", padx=(6, 6))
        ttk.Button(frame, text="Browse", command=lambda: browse_dir(root_var)).grid(row=row, column=2, sticky="w")
        row += 1

        ttk.Label(frame, text="Prefer HFEst").grid(row=row, column=0, sticky="w")
        ttk.Checkbutton(frame, variable=prefer_var).grid(row=row, column=1, sticky="w", padx=(6, 0))
        row += 1

        ttk.Label(frame, text="HomeFront.csproj").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=project_var, width=60).grid(row=row, column=1, sticky="we", padx=(6, 6))
        ttk.Button(frame, text="Browse", command=lambda: browse_dir(project_var)).grid(row=row, column=2, sticky="w")
        row += 1

        ttk.Label(frame, text="Output Dir").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=output_dir_var, width=60).grid(row=row, column=1, sticky="we", padx=(6, 6))
        ttk.Button(frame, text="Browse", command=lambda: browse_dir(output_dir_var)).grid(row=row, column=2, sticky="w")
        row += 1

        ttk.Label(frame, text="Output File").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=output_name_var, width=40).grid(row=row, column=1, sticky="w", padx=(6, 6))
        row += 1

        ttk.Label(frame, text="Model").grid(row=row, column=0, sticky="w")
        model_frame = ttk.Frame(frame)
        model_frame.grid(row=row, column=1, sticky="w", padx=(6, 0))
        ttk.Radiobutton(model_frame, text="claude-opus-4-6", variable=model_var, value="claude-opus-4-6").pack(side="left", padx=(0, 10))
        ttk.Radiobutton(model_frame, text="gpt-5.2-codex", variable=model_var, value="gpt-5.2-codex").pack(side="left")
        row += 1

        ttk.Label(frame, text="Temp").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=temp_var, width=10).grid(row=row, column=1, sticky="w", padx=(6, 0))
        row += 1

        ttk.Label(frame, text="Context Dir").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=context_var, width=60).grid(row=row, column=1, sticky="we", padx=(6, 6))
        ttk.Button(frame, text="Browse", command=lambda: browse_dir(context_var)).grid(row=row, column=2, sticky="w")
        row += 1

        ttk.Label(frame, text="VB6 Image Dir").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=vb6_image_dir_var, width=60).grid(row=row, column=1, sticky="we", padx=(6, 6))
        ttk.Button(frame, text="Browse", command=lambda: browse_dir(vb6_image_dir_var)).grid(row=row, column=2, sticky="w")
        row += 1

        ttk.Label(frame, text=".NET Image Dir").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=dotnet_image_dir_var, width=60).grid(row=row, column=1, sticky="we", padx=(6, 6))
        ttk.Button(frame, text="Browse", command=lambda: browse_dir(dotnet_image_dir_var)).grid(row=row, column=2, sticky="w")
        row += 1

        ttk.Label(frame, text="Instructions").grid(row=row, column=0, sticky="nw")
        instr = tk.Text(frame, height=6, wrap="word")
        instr.grid(row=row, column=1, columnspan=2, sticky="we", padx=(6, 0))
        instr.insert("1.0", instructions_text.get("1.0", tk.END))
        row += 1

        def on_save():
            instructions_text.delete("1.0", tk.END)
            instructions_text.insert("1.0", instr.get("1.0", tk.END).strip())
            status_var.set("Configuration updated.")
            dlg.destroy()

        btns = ttk.Frame(frame)
        btns.grid(row=row, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Save", command=on_save).pack(side=tk.RIGHT, padx=(0, 8))

        frame.columnconfigure(1, weight=1)

    closing_state = {"active": False}

    def stop_all_jobs():
        nonlocal run_process, tail_job_id, run_session_id, run_project_path
        run_session_id = f"shutdown:{time.time():.6f}"
        if tail_job_id is not None:
            try:
                root.after_cancel(tail_job_id)
            except Exception:
                pass
            tail_job_id = None
        if run_process is not None:
            proc = run_process
            run_process = None
            _terminate_process(proc)
        run_project_path = None

    def on_app_exit():
        if closing_state["active"]:
            return
        closing_state["active"] = True
        stop_all_jobs()
        root.destroy()

    def queue_project_switch(project_dir: Path) -> None:
        ui_action["action"] = "load_project"
        ui_action["project_dir"] = str(project_dir.expanduser().resolve())
        on_app_exit()

    def on_new_project():
        initial_source = root_var.get().strip() or str(Path.home())
        source_selected = filedialog.askdirectory(
            parent=root,
            title="Select VB6 Source Folder",
            initialdir=initial_source,
        )
        if not source_selected:
            return

        suggested_name = sanitize_project_name(Path(source_selected).name or "Project")
        entered_name = simpledialog.askstring(
            "New Project",
            "Project name folder:",
            parent=root,
            initialvalue=suggested_name,
        )
        if entered_name is None:
            status_var.set("New project cancelled.")
            return
        project_name = sanitize_project_name(entered_name)

        projects_root_raw = ui_defaults.get("projects_root", "").strip()
        projects_root = Path(projects_root_raw).expanduser() if projects_root_raw else (Path(__file__).resolve().parent / "projects")
        candidate = projects_root / project_name
        overwrite = False
        if candidate.exists() and any(candidate.iterdir()):
            overwrite = messagebox.askyesno(
                "Overwrite Project",
                f"Project folder already exists:\n{candidate}\n\nOverwrite it?",
                parent=root,
            )
            if not overwrite:
                status_var.set("New project cancelled.")
                return

        try:
            project_dir, source_count, _target_count = create_project_package(
                projects_root=projects_root,
                project_name=project_name,
                vb6_source_root=Path(source_selected),
                output_project=project_var.get().strip(),
                seed_target_dir=None,
                overwrite=overwrite,
            )
        except Exception as exc:
            messagebox.showerror("New Project", f"Failed to create project:\n{exc}", parent=root)
            status_var.set(f"New project failed: {exc}")
            last_error_var.set(str(exc))
            return

        status_var.set(f"Created project '{project_dir.name}' with {source_count} VB6 files.")
        queue_project_switch(project_dir)

    def on_load_project():
        projects_root_raw = ui_defaults.get("projects_root", "").strip()
        projects_root = Path(projects_root_raw).expanduser() if projects_root_raw else (Path(__file__).resolve().parent / "projects")
        candidates = list_existing_project_dirs(projects_root, include_default=True)
        if not candidates:
            messagebox.showinfo("Load Project", "No existing projects found.", parent=root)
            return

        chooser = tk.Toplevel(root)
        chooser.title("Load Existing Project")
        chooser.geometry("560x170")
        chooser.transient(root)
        chooser.grab_set()

        frame = ttk.Frame(chooser, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Existing Projects").grid(row=0, column=0, sticky="w")

        labels: list[str] = []
        by_label: Dict[str, Path] = {}
        default_label = ""
        for idx, path in enumerate(candidates):
            label = path.name
            if not path.exists():
                label += " (not created)"
            elif not (path / PROJECT_MANIFEST_NAME).exists():
                label += " (no manifest)"
            labels.append(label)
            by_label[label] = path
            if idx == 0:
                default_label = label

        selected_var = tk.StringVar(value=default_label)
        combo = ttk.Combobox(frame, textvariable=selected_var, values=labels, state="readonly", width=56)
        combo.grid(row=1, column=0, sticky="we", pady=(6, 10))

        chosen_path: Dict[str, Path | None] = {"value": None}

        def accept():
            label = selected_var.get().strip()
            selected_path = by_label.get(label)
            if selected_path is None:
                return
            chosen_path["value"] = selected_path
            chooser.destroy()

        def cancel():
            chooser.destroy()

        buttons = ttk.Frame(frame)
        buttons.grid(row=2, column=0, sticky="e")
        ttk.Button(buttons, text="Cancel", command=cancel).pack(side=tk.RIGHT)
        ttk.Button(buttons, text="Load", command=accept).pack(side=tk.RIGHT, padx=(0, 8))

        frame.columnconfigure(0, weight=1)
        combo.focus_set()
        chooser.bind("<Return>", lambda _e: accept())
        chooser.wait_window()

        project_dir = chosen_path["value"]
        if project_dir is None:
            return

        try:
            layout = get_project_layout(project_dir)
            source_count = count_component_files(layout["source_dir"])
            if source_count == 0:
                raise RuntimeError(
                    f"No VB6 files found in project source: {layout['source_dir']}"
                )
        except Exception as exc:
            messagebox.showerror("Load Project", f"Invalid project folder:\n{exc}", parent=root)
            status_var.set(f"Load project failed: {exc}")
            last_error_var.set(str(exc))
            return

        status_var.set(f"Loading project '{layout['name']}'...")
        queue_project_switch(layout["project_dir"])

    def on_close_project():
        if not messagebox.askyesno(
            "Close Project",
            "Close the current project?",
            parent=root,
        ):
            return
        projects_root_raw = ui_defaults.get("projects_root", "").strip()
        projects_root = Path(projects_root_raw).expanduser() if projects_root_raw else (Path(__file__).resolve().parent / "projects")
        empty_project = ensure_empty_project(projects_root)
        status_var.set("Project closed.")
        queue_project_switch(empty_project)

    context_cache = {"dir": "", "signature": "", "blob": ""}
    context_cache_lock = threading.Lock()

    def get_context_blob_for(context_dir_str: str) -> str:
        context_dir = Path(context_dir_str).expanduser()
        if not context_dir.exists():
            return ""
        files = collect_context_files(context_dir)
        signature = compute_context_signature(files)
        dir_key = str(context_dir.resolve())
        with context_cache_lock:
            if context_cache["dir"] == dir_key and context_cache["signature"] == signature:
                return context_cache["blob"]
        blob = build_context_blob(context_dir, files)
        with context_cache_lock:
            context_cache["dir"] = dir_key
            context_cache["signature"] = signature
            context_cache["blob"] = blob
        return blob

    def run_async(task_label: str, work_fn, success_fn, error_fn=None):
        def _worker():
            try:
                result = work_fn()
                root.after(0, lambda: success_fn(result))
            except Exception as exc:
                def _err(e=exc):
                    if error_fn:
                        error_fn(e)
                    else:
                        status_var.set(f"{task_label} failed: {e}")
                        last_error_var.set(str(e))
                root.after(0, _err)

        threading.Thread(target=_worker, daemon=True).start()

    def on_rebuild_cache():
        nonlocal total_lines, cache_dirty, cache_index, chunks_enabled
        status_var.set("Rebuilding cache...")
        root.update_idletasks()
        try:
            shutil.rmtree(cache_dir, ignore_errors=True)
            cache_dir.mkdir(parents=True, exist_ok=True)
            cache_index = {"version": 1, "files": {}, "tree": {}, "chunks_enabled": True}
            files_index.clear()
            total_lines = 0
            for comp in components:
                comp_id = f"{comp.project}::{comp.name}"
                try:
                    stat = comp.path.stat()
                    file_size_by_id[comp_id] = stat.st_size
                    line_count = count_lines_file(comp.path)
                    line_count_by_id[comp_id] = line_count
                    files_index[str(comp.path)] = {
                        "mtime": stat.st_mtime,
                        "size": stat.st_size,
                        "line_count": line_count,
                    }
                    total_lines += line_count
                except Exception:
                    line_count_by_id[comp_id] = None
                    file_size_by_id[comp_id] = 0
            save_cache_index(cache_dir, cache_index)
            # Update stats and order list
            refresh_stats_labels()
            for cid, row_id in order_row_by_comp_id.items():
                line_count = line_count_by_id.get(cid)
                order_tree.set(row_id, "lines", str(line_count) if isinstance(line_count, int) else "")
                order_tree.set(row_id, "size", str(file_size_by_id.get(cid, 0)))

            signature = compute_signature(components)
            save_tree_cache(cache_dir, signature, prefer_var.get(), edges, roots)
            chunks_enabled = True
            status_var.set("Cache rebuilt.")
        except Exception as exc:
            status_var.set(f"Cache rebuild failed: {exc}")
            last_error_var.set(str(exc))

    def on_refresh_stats():
        nonlocal total_lines, total_files, program_types, migration_order
        status_var.set("Refreshing project stats...")
        root.update_idletasks()
        try:
            valid_paths = {str(c.path) for c in components if c.path.exists()}
            stale_keys = [k for k in list(files_index.keys()) if k not in valid_paths or not Path(k).exists()]
            for key in stale_keys:
                files_index.pop(key, None)

            total_lines = 0
            total_files = len(components)
            refreshed_types: set[str] = set()
            for comp in components:
                comp_id = f"{comp.project}::{comp.name}"
                path_key = str(comp.path)
                try:
                    stat = comp.path.stat()
                    line_count = count_lines_file(comp.path)
                    file_size_by_id[comp_id] = stat.st_size
                    line_count_by_id[comp_id] = line_count
                    total_lines += line_count
                    files_index[path_key] = {
                        "mtime": stat.st_mtime,
                        "size": stat.st_size,
                        "line_count": line_count,
                    }
                except Exception:
                    file_size_by_id[comp_id] = 0
                    line_count_by_id[comp_id] = None
                    files_index.pop(path_key, None)
                refreshed_types.add(comp.path.suffix.lower() or "unknown")

            program_types = sorted(refreshed_types)
            save_cache_index(cache_dir, cache_index)

            rebuild_order_rows(recompute_order=True)
            populate_order(order_rows)
            sort_state.clear()
            for c, title in header_titles.items():
                order_tree.heading(c, text=title)

            refresh_stats_labels()
            status_var.set("Stats refreshed. Grid organized by conversion order.")
        except Exception as exc:
            status_var.set(f"Stats refresh failed: {exc}")
            last_error_var.set(str(exc))

    rebuild_btn = ttk.Button(controls, text="Rebuild Cache", command=on_rebuild_cache)
    rebuild_btn.grid(row=0, column=5, sticky="e", padx=(6, 0))
    refresh_stats_btn = ttk.Button(controls, text="Refresh Stats", command=on_refresh_stats)
    refresh_stats_btn.grid(row=0, column=6, sticky="e", padx=(6, 0))
    align_btn = ttk.Button(controls, text="Align Screens")
    align_btn.grid(row=0, column=7, sticky="e", padx=(6, 0))

    for i in range(8):
        controls.columnconfigure(i, weight=1)
    mark_chunk("chunk4_controls", last_time)

    main = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
    main.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

    project_frame = ttk.Frame(main)
    left_frame = ttk.Frame(main)
    right_frame = ttk.Frame(main)
    main.add(project_frame, weight=1)
    main.add(left_frame, weight=1)
    main.add(right_frame, weight=3)

    project_header = ttk.Frame(project_frame)
    project_header.pack(fill=tk.X)
    ttk.Label(project_header, text=".NET Blazor Project").pack(side=tk.LEFT)

    project_tree = ttk.Treeview(project_frame, show="tree")
    project_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    project_scroll = ttk.Scrollbar(project_frame, orient=tk.VERTICAL, command=project_tree.yview)
    project_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    project_tree.configure(yscrollcommand=project_scroll.set)

    project_item_to_path: Dict[str, Path] = {}
    project_loaded_dirs: set[str] = set()

    project_root_path: Path | None = None
    if project_var.get().strip():
        project_csproj = Path(project_var.get()).expanduser()
        if project_csproj.exists():
            project_root_path = project_csproj.parent
    if project_root_path is None:
        project_root_path = project_root_guess if project_root_guess.exists() else None

    def project_should_skip(path: Path) -> bool:
        name = path.name
        if name in {"bin", "obj", ".git", ".idea", ".vs", "__pycache__"}:
            return True
        if name.startswith(".") and name not in {".vscode"}:
            return True
        if path.is_file() and name == ".DS_Store":
            return True
        return False

    def project_iter_children(dir_path: Path) -> List[Path]:
        try:
            entries = list(dir_path.iterdir())
        except Exception:
            return []
        items = [p for p in entries if not project_should_skip(p)]
        return sorted(items, key=lambda p: (0 if p.is_dir() else 1, p.name.lower()))

    def project_has_children(dir_path: Path) -> bool:
        try:
            for p in dir_path.iterdir():
                if not project_should_skip(p):
                    return True
        except Exception:
            return False
        return False

    def project_add_placeholder(item_id: str) -> None:
        placeholder = project_tree.insert(item_id, "end", text="...", tags=("placeholder",))
        project_item_to_path.pop(placeholder, None)

    def project_populate(item_id: str) -> None:
        if item_id in project_loaded_dirs:
            return
        dir_path = project_item_to_path.get(item_id)
        if not dir_path or not dir_path.exists() or not dir_path.is_dir():
            return

        for child in project_tree.get_children(item_id):
            project_item_to_path.pop(child, None)
            project_tree.delete(child)

        for child_path in project_iter_children(dir_path):
            child_id = project_tree.insert(item_id, "end", text=child_path.name)
            project_item_to_path[child_id] = child_path
            if child_path.is_dir() and project_has_children(child_path):
                project_add_placeholder(child_id)
        project_loaded_dirs.add(item_id)

    def refresh_project_tree() -> None:
        project_tree.delete(*project_tree.get_children(""))
        project_item_to_path.clear()
        project_loaded_dirs.clear()
        if project_root_path is None or not project_root_path.exists():
            return
        root_id = project_tree.insert("", "end", text=project_root_path.name or str(project_root_path), open=True)
        project_item_to_path[root_id] = project_root_path
        project_populate(root_id)

    def show_dotnet_project_file(path: Path) -> None:
        nonlocal current_output_path, blazor_header_lines
        if not path.exists() or not path.is_file():
            return
        header_text = f"{path.name} :: {path}\n\n"
        blazor_header_lines = header_text.count("\n")
        binary_suffixes = {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".webp", ".svg",
            ".frx", ".ctx", ".dll", ".exe", ".zip", ".pdf"
        }
        if path.suffix.lower() in binary_suffixes:
            highlight_blazor(blazor_text, header_text + "[Binary file preview is disabled.]")
            blazor_text.tag_add("hdr", "1.0", "2.0")
            current_output_path = None
            return
        try:
            content = path.read_text(encoding="utf-8", errors="ignore")
        except Exception as exc:
            highlight_blazor(blazor_text, header_text + f"[Unable to read file: {exc}]")
            blazor_text.tag_add("hdr", "1.0", "2.0")
            current_output_path = None
            return
        highlight_blazor(blazor_text, header_text + content)
        blazor_text.tag_add("hdr", "1.0", "2.0")
        current_output_path = path

    def on_project_tree_open(_event):
        item_id = project_tree.focus()
        if item_id:
            project_populate(item_id)

    def on_project_tree_select(_event):
        sel = project_tree.selection()
        if not sel:
            return
        item_id = sel[0]
        path = project_item_to_path.get(item_id)
        if not path:
            return
        if path.is_file():
            show_dotnet_project_file(path)

    project_tree.bind("<<TreeviewOpen>>", on_project_tree_open)
    project_tree.bind("<<TreeviewSelect>>", on_project_tree_select)

    refresh_project_btn = ttk.Button(project_header, text="Refresh", command=refresh_project_tree)
    refresh_project_btn.pack(side=tk.RIGHT)
    refresh_project_tree()

    stats_frame = ttk.Frame(left_frame)
    stats_frame.pack(fill=tk.X)
    stats_text = f"Files: {total_files} | Program types: {len(program_types)} | Lines (cached): {total_lines}"
    types_text = ", ".join(program_types)
    stats_label = ttk.Label(stats_frame, text=stats_text)
    stats_label.pack(anchor="w")
    types_label = ttk.Label(stats_frame, text=f"Types: {types_text}")
    types_label.pack(anchor="w")

    def refresh_stats_labels() -> None:
        stats_label.configure(
            text=f"Files: {total_files} | Program types: {len(program_types)} | Lines (cached): {total_lines}"
        )
        types_label.configure(text=f"Types: {', '.join(program_types)}")

    view_toggle = ttk.Frame(left_frame)
    view_toggle.pack(fill=tk.X, pady=(6, 4))
    view_mode_var = tk.StringVar(value="tree")
    ttk.Label(view_toggle, text="View").pack(side=tk.LEFT)
    tree_rb = ttk.Radiobutton(view_toggle, text="Tree", variable=view_mode_var, value="tree")
    grid_rb = ttk.Radiobutton(view_toggle, text="Grid", variable=view_mode_var, value="grid")
    tree_rb.pack(side=tk.LEFT, padx=(6, 0))
    grid_rb.pack(side=tk.LEFT, padx=(6, 0))

    tree_container = ttk.Frame(left_frame)
    grid_container = ttk.Frame(left_frame)
    tree_container.pack(fill=tk.BOTH, expand=True)

    tree_frame = ttk.Frame(tree_container)
    tree_frame.pack(fill=tk.BOTH, expand=True)
    order_frame = ttk.Frame(grid_container)
    order_frame.pack(fill=tk.BOTH, expand=True)

    style = ttk.Style()
    style.configure("Migrator.Treeview", rowheight=36)
    tree = ttk.Treeview(tree_frame, show="tree", style="Migrator.Treeview")
    tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=tree_scroll.set)

    order_header = ttk.Frame(order_frame)
    order_header.pack(fill=tk.X)
    order_label = ttk.Label(order_header, text="VB6 Files")
    order_label.pack(side=tk.LEFT)
    order_tree = ttk.Treeview(
        order_frame,
        columns=("convertred", "order", "component", "type", "project", "size", "lines"),
        show="headings",
        height=20,
        selectmode="extended",
    )
    tree.tag_configure("orphan", foreground="#cc0000")
    order_tree.tag_configure("orphan", foreground="#cc0000")
    header_titles = {
        "convertred": "Convertred",
        "order": "#",
        "component": "Component",
        "type": "Type",
        "project": "Project",
        "size": "File Size",
        "lines": "Line Count",
    }
    for col, title in header_titles.items():
        order_tree.heading(col, text=title)
    order_tree.column("convertred", width=80, anchor="center")
    order_tree.column("order", width=40, anchor="e")
    order_tree.column("component", width=160, anchor="w")
    order_tree.column("type", width=60, anchor="w")
    order_tree.column("project", width=70, anchor="w")
    order_tree.column("size", width=80, anchor="e")
    order_tree.column("lines", width=80, anchor="e")
    order_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    order_scroll = ttk.Scrollbar(order_frame, orient=tk.VERTICAL, command=order_tree.yview)
    order_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    order_tree.configure(yscrollcommand=order_scroll.set)

    def update_left_view():
        if view_mode_var.get() == "tree":
            grid_container.pack_forget()
            tree_container.pack(fill=tk.BOTH, expand=True)
        else:
            tree_container.pack_forget()
            grid_container.pack(fill=tk.BOTH, expand=True)

    tree_rb.configure(command=update_left_view)
    grid_rb.configure(command=update_left_view)
    update_left_view()

    comp_by_id = {f"{c.project}::{c.name}": c for c in components}
    order_row_by_comp_id: Dict[str, str] = {}
    order_comp_by_row_id: Dict[str, str] = {}

    def alphanum_key(text: str):
        return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\\d+)", text)]

    cycle_ids = set(cycle_set)
    orphan_ids = set(orphan_set)
    order_rows = []
    converted_rows: set[str] = set()
    for comp in components:
        comp_id = f"{comp.project}::{comp.name}"
        candidate = pre_output_dir / f"{comp.name}.razor"
        if candidate.exists():
            converted_rows.add(comp_id)

    def compute_orphan_ids() -> set[str]:
        comp_ids_local = [f"{c.project}::{c.name}" for c in components]
        entry_points = set(roots)
        if "HFSystem::FLogin" in adj:
            entry_points.add("HFSystem::FLogin")
        inbound_by_name: Dict[str, int] = {}
        for (_caller, callee), _meta in edges.items():
            name = callee.split("::", 1)[-1]
            inbound_by_name[name] = inbound_by_name.get(name, 0) + 1
        local_orphans: set[str] = set()
        for cid_local in comp_ids_local:
            if cid_local in entry_points:
                continue
            comp_local = comp_by_id.get(cid_local)
            if not comp_local:
                continue
            if comp_local.path.suffix.lower() != ".frm":
                continue
            name_local = cid_local.split("::", 1)[-1]
            if inbound_by_name.get(name_local, 0) == 0:
                local_orphans.add(cid_local)
        return local_orphans

    def rebuild_order_rows(recompute_order: bool = False) -> None:
        nonlocal order_rows, migration_order, cycle_ids, orphan_ids
        if recompute_order:
            migration_order_local, cycle_set_local = compute_migration_order(components, edges)
            migration_order = migration_order_local
            cycle_ids = set(cycle_set_local)
            orphan_ids = compute_orphan_ids()

        rebuilt_rows = []
        for idx, cid in enumerate(migration_order, start=1):
            comp = comp_by_id.get(cid)
            if not comp:
                continue
            ext = comp.path.suffix.lower().lstrip(".")
            name = comp.name
            display_name = f"{name} (cycle)" if cid in cycle_ids else name
            tags = ("orphan",) if cid in orphan_ids else ()
            size = file_size_by_id.get(cid, 0)
            line_count = line_count_by_id.get(cid)
            line_display = line_count if isinstance(line_count, int) else ""
            rebuilt_rows.append(
                {
                    "cid": cid,
                    "values": ("", idx, display_name, ext, comp.project, size, line_display),
                    "tags": tags,
                    "sort_key": alphanum_key(name),
                }
            )
        order_rows = rebuilt_rows

    rebuild_order_rows(recompute_order=False)

    sort_state: Dict[str, bool] = {}

    def populate_order(rows):
        order_tree.delete(*order_tree.get_children())
        order_row_by_comp_id.clear()
        order_comp_by_row_id.clear()
        for row in rows:
            cid = row["cid"]
            chk = "✔" if cid in converted_rows else ""
            values = (chk,) + tuple(row["values"][1:])
            row_id = order_tree.insert("", "end", values=values, tags=row["tags"])
            order_row_by_comp_id[row["cid"]] = row_id
            order_comp_by_row_id[row_id] = row["cid"]

    populate_order(order_rows)
    mark_chunk("chunk5_left_panes", last_time)

    def on_tag():
        """Tag selected VB6 files: send each to LLM to annotate code blocks with
        descriptions and .NET conversion instructions, then write back and refresh cache."""
        selected_rows = order_tree.selection()
        if not selected_rows:
            status_var.set("No selected items to tag.")
            return
        checked = [order_comp_by_row_id.get(rid) for rid in selected_rows]
        checked = [cid for cid in checked if cid]
        if not checked:
            status_var.set("No valid components selected.")
            return

        context_dir_str = context_var.get().strip()
        model = model_var.get().strip() or "claude-opus-4-6"
        total = len(checked)
        status_var.set(f"Tagging 0/{total}...")
        root.update_idletasks()

        def work():
            context_blob = get_context_blob_for(context_dir_str)
            pause = 5 if is_claude_model(model) else 2
            results = []
            for idx, comp_id in enumerate(checked, 1):
                comp = id_to_component.get(comp_id)
                if not comp:
                    results.append((comp_id, None, "Component not found"))
                    continue

                root.after(0, lambda i=idx, c=comp: status_var.set(
                    f"Tagging {i}/{total}: {c.name}..."
                ))

                try:
                    vb6_source = read_text(comp.path)

                    # Look for matching migrated .razor file in the output directory
                    migrated_source = ""
                    out_dir_str = output_dir_var.get().strip()
                    if out_dir_str:
                        migrated_path = Path(out_dir_str).expanduser() / f"{comp.name}.razor"
                        if migrated_path.exists():
                            try:
                                migrated_source = migrated_path.read_text(encoding="utf-8", errors="ignore")
                            except Exception:
                                migrated_source = ""

                    tagged_source = call_claude_tag(
                        source_text=vb6_source,
                        source_path=comp.path,
                        model=model,
                        context_blob=context_blob,
                        migrated_source=migrated_source,
                    )
                    # Write tagged source back to the VB6 file
                    comp.path.write_text(tagged_source, encoding="latin-1", errors="ignore")

                    # Invalidate source_cache and chunk cache so next view picks up changes
                    source_cache.pop(comp_id, None)
                    cp = chunk_cache_path(comp.path)
                    if cp.exists():
                        cp.unlink()
                    # Remove from files_index to force re-cache
                    files_index.pop(str(comp.path), None)
                    save_cache_index(cache_dir, cache_index)

                    results.append((comp_id, comp, None))
                except Exception as exc:
                    results.append((comp_id, comp, str(exc)))

                if idx < total:
                    time.sleep(pause)

            return results

        def on_success(results):
            ok = sum(1 for _, _, err in results if err is None)
            fail = sum(1 for _, _, err in results if err is not None)
            msg = f"Tagging complete: {ok} succeeded"
            if fail:
                msg += f", {fail} failed"
                # Show first error
                for cid, comp, err in results:
                    if err:
                        last_error_var.set(f"Tag error ({cid}): {err}")
                        break
            status_var.set(msg)
            # Refresh the currently displayed component if it was tagged
            sel = order_tree.selection()
            if sel:
                row_id = sel[0]
                cid = order_comp_by_row_id.get(row_id)
                if cid:
                    show_component(cid)

        def on_error(exc):
            status_var.set(f"Tag failed: {exc}")
            last_error_var.set(str(exc))

        run_async("Tag", work, on_success, on_error)

    def on_sort_original():
        populate_order(order_rows)

    def sort_by_column(col: str):
        reverse = sort_state.get(col, False)

        def key_fn(row):
            if col == "convertred":
                return 1 if row["cid"] in converted_rows else 0
            if col == "order":
                return row["values"][1]
            if col == "component":
                return row["values"][2].lower()
            if col == "type":
                return row["values"][3].lower()
            if col == "project":
                return row["values"][4].lower()
            if col == "size":
                return row["values"][5]
            if col == "lines":
                return row["values"][6] if isinstance(row["values"][6], int) else 0
            return row["values"][1]

        sorted_rows = sorted(order_rows, key=key_fn, reverse=reverse)
        populate_order(sorted_rows)
        sort_state[col] = not reverse
        for c, title in header_titles.items():
            if c == col:
                arrow = " ▼" if reverse else " ▲"
                order_tree.heading(c, text=title + arrow)
            else:
                order_tree.heading(c, text=title)

    for col in ("convertred", "order", "component", "type", "project", "size", "lines"):
        order_tree.heading(col, command=lambda c=col: sort_by_column(c))

    def on_select_all():
        rows = order_tree.get_children()
        order_tree.selection_set(rows)
        status_var.set("Selected all items.")

    def on_clear_all():
        order_tree.selection_remove(order_tree.selection())
        status_var.set("Cleared all selections.")

    def build_search_results(paths: list[Path], pattern: str, base_dir: Path | None, max_matches: int):
        try:
            rx = re.compile(pattern, re.IGNORECASE)
        except re.error:
            rx = re.compile(re.escape(pattern), re.IGNORECASE)
        results: list[dict] = []
        match_count = 0
        for path in paths:
            try:
                content = path.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                content = read_text(path)
            lines = content.splitlines()
            for idx, line in enumerate(lines, start=1):
                if not rx.search(line):
                    continue
                match_count += 1
                header = str(path)
                if base_dir:
                    try:
                        header = str(path.relative_to(base_dir))
                    except Exception:
                        pass
                results.append({"gutter": "", "text": header, "file": path, "line": None})
                start = max(1, idx - 2)
                end = min(len(lines), idx + 2)
                for ln in range(start, end + 1):
                    results.append({"gutter": str(ln), "text": lines[ln - 1], "file": path, "line": ln})
                if match_count >= max_matches:
                    results.append({"gutter": "", "text": f"... truncated after {max_matches} matches", "file": None, "line": None})
                    return results
        if not results:
            results.append({"gutter": "", "text": "No matches.", "file": None, "line": None})
        return results

    def on_search():
        popup = tk.Toplevel(root)
        popup.title("Search")
        popup.geometry("900x600")
        popup.transient(root)
        popup.grab_set()

        pattern_var = tk.StringVar(value=search_var.get())
        target_var = tk.BooleanVar(value=False)  # False=VB6 (default), True=.NET

        header = ttk.Frame(popup)
        header.pack(fill=tk.X, padx=8, pady=6)
        ttk.Label(header, text="Pattern").pack(side=tk.LEFT)
        pattern_entry = ttk.Entry(header, textvariable=pattern_var, width=50)
        pattern_entry.pack(side=tk.LEFT, padx=(6, 12))
        ttk.Checkbutton(header, text="Search target (.NET)", variable=target_var).pack(side=tk.LEFT)
        run_btn = ttk.Button(header, text="Search")
        run_btn.pack(side=tk.RIGHT)

        results_frame = ttk.Frame(popup)
        results_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        ttk.Label(results_frame, text="Search Results").pack(anchor="w")

        results_body = ttk.Frame(results_frame)
        results_body.pack(fill=tk.BOTH, expand=True)
        search_gutter = tk.Text(results_body, width=6, padx=4, takefocus=0, borderwidth=0)
        search_gutter.pack(side=tk.LEFT, fill=tk.Y)
        search_text = tk.Text(results_body, wrap="none")
        search_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        def on_search_scroll(*args):
            search_text.yview(*args)
            search_gutter.yview(*args)

        search_scroll_y = ttk.Scrollbar(results_body, orient=tk.VERTICAL, command=on_search_scroll)
        search_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        search_scroll_x = ttk.Scrollbar(results_frame, orient=tk.HORIZONTAL, command=search_text.xview)
        search_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        def on_search_yscroll(first, last):
            search_scroll_y.set(first, last)
            search_gutter.yview_moveto(first)

        search_text.configure(yscrollcommand=on_search_yscroll, xscrollcommand=search_scroll_x.set)
        search_text.bind("<Key>", lambda _e: "break")
        search_gutter.bind("<Key>", lambda _e: "break")
        search_gutter.configure(font=mono, background="#000000", foreground="#999999")
        search_text.configure(font=mono)

        search_line_map: list[dict] = []

        def set_results(rows: list[dict]):
            nonlocal search_line_map
            search_line_map = rows
            search_text.configure(state="normal")
            search_gutter.configure(state="normal")
            search_text.delete("1.0", tk.END)
            search_gutter.delete("1.0", tk.END)
            search_text.insert("1.0", "\n".join(row["text"] for row in rows))
            search_gutter.insert("1.0", "\n".join(row["gutter"] for row in rows))
            search_text.configure(state="normal")
            search_gutter.configure(state="normal")

        def show_vb6_file(path: Path):
            comp_id = comp_id_by_path.get(str(path.resolve()))
            if comp_id:
                show_component(comp_id)
                return
            content = read_text(path)
            highlight_vb(vb_text, content)
            highlight_blazor(blazor_text, "Not converted yet.")

        def show_net_file(path: Path):
            nonlocal current_output_path, blazor_header_lines
            comp_id = comp_id_by_name.get(path.stem.lower())
            if comp_id:
                show_component(comp_id)
                return
            try:
                content = path.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                content = ""
            header_text = f"{path.name} :: {path}\n\n"
            blazor_header_lines = header_text.count("\\n")
            highlight_blazor(blazor_text, header_text + content)
            blazor_text.tag_add("hdr", "1.0", "2.0")
            current_output_path = path
            highlight_vb(vb_text, f"No VB6 source found for {path.name}")

        def on_gutter_click(event):
            index = search_gutter.index(f"@{event.x},{event.y}")
            line_no = int(index.split(".")[0])
            if line_no - 1 >= len(search_line_map):
                return
            row = search_line_map[line_no - 1]
            file_path = row.get("file")
            if not file_path:
                return
            if target_var.get():
                show_net_file(file_path)
            else:
                show_vb6_file(file_path)
            popup.destroy()

        search_gutter.bind("<Button-1>", on_gutter_click)

        def do_search():
            pattern = pattern_var.get().strip()
            if not pattern:
                status_var.set("Enter a search pattern.")
                return
            search_var.set(pattern)
            status_var.set("Searching...")
            root.update_idletasks()
            is_target = target_var.get()
            if is_target:
                context_dir = Path(context_var.get()).expanduser()
                paths = collect_context_files(context_dir) if context_dir.exists() else []
                base_dir = context_dir if context_dir.exists() else None
            else:
                paths = [c.path for c in components]
                base_dir = Path(root_var.get()).expanduser()

            def work():
                return build_search_results(paths, pattern, base_dir, MAX_SEARCH_MATCHES)

            def success(rows):
                set_results(rows)
                status_var.set("Search complete.")

            def error(exc: Exception):
                status_var.set(f"Search failed: {exc}")
                last_error_var.set(str(exc))

            run_async("Search", work, success, error)

        run_btn.configure(command=do_search)
        popup.bind("<Return>", lambda _e: do_search())
        pattern_entry.focus_set()

    tag_btn = ttk.Button(order_header, text="Tag", command=on_tag)
    tag_btn.pack(side=tk.RIGHT)
    reset_btn = ttk.Button(order_header, text="Original Order", command=on_sort_original)
    reset_btn.pack(side=tk.RIGHT, padx=(6, 6))
    convert_checked_btn = ttk.Button(order_header, text="Convert")
    convert_checked_btn.pack(side=tk.RIGHT, padx=(6, 6))
    select_all_btn = ttk.Button(order_header, text="Select All", command=on_select_all)
    select_all_btn.pack(side=tk.RIGHT, padx=(6, 0))
    clear_all_btn = ttk.Button(order_header, text="Clear All", command=on_clear_all)
    clear_all_btn.pack(side=tk.RIGHT, padx=(6, 0))
    mark_chunk("chunk6_search_actions", last_time)

    content_split = ttk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
    content_split.pack(fill=tk.BOTH, expand=True)

    vb_frame = ttk.Frame(content_split)
    blazor_frame = ttk.Frame(content_split)
    content_split.add(vb_frame, weight=1)
    content_split.add(blazor_frame, weight=1)

    current_comp_state = {"id": None}
    vb_image_overrides: Dict[str, Path] = {}
    dotnet_image_overrides: Dict[str, Path] = {}
    vb_image_state = {"photo": None, "path": None}
    dotnet_image_state = {"photo": None, "path": None}
    vb_view_mode = tk.StringVar(value="source")
    dotnet_view_mode = tk.StringVar(value="source")

    def resolve_image_dir(raw_dir: str) -> Path:
        p = Path(raw_dir).expanduser()
        if p.exists():
            return p
        # Common case-insensitive variants used in this workspace.
        raw = str(p)
        alt = raw.replace("/VB6", "/vb6").replace("/dotNET", "/dotnet")
        p_alt = Path(alt).expanduser()
        if p_alt.exists():
            return p_alt
        alt2 = raw.replace("/vb6", "/VB6").replace("/dotnet", "/dotNET")
        p_alt2 = Path(alt2).expanduser()
        if p_alt2.exists():
            return p_alt2
        return p

    def resolve_panel_image(comp_id: str | None, is_vb6: bool) -> Path | None:
        if not comp_id:
            return None
        overrides = vb_image_overrides if is_vb6 else dotnet_image_overrides
        override = overrides.get(comp_id)
        if override and override.exists():
            return override
        comp = id_to_component.get(comp_id)
        if not comp:
            return None
        image_dir = resolve_image_dir(vb6_image_dir_var.get()) if is_vb6 else resolve_image_dir(dotnet_image_dir_var.get())
        return find_component_image(image_dir, comp.name)

    def load_canvas_image(canvas, state_dict: dict, image_path: Path | None, missing_text: str) -> None:
        canvas.delete("all")
        state_dict["photo"] = None
        state_dict["path"] = image_path
        if image_path is None or not image_path.exists():
            canvas.configure(scrollregion=(0, 0, 1, 1))
            canvas.create_text(8, 8, anchor="nw", fill="#888888", text=missing_text)
            return
        try:
            photo = None
            if Image is not None and ImageTk is not None:
                with Image.open(image_path) as img:
                    photo = ImageTk.PhotoImage(img.copy())
            else:
                photo = tk.PhotoImage(file=str(image_path))
            state_dict["photo"] = photo
            canvas.create_image(0, 0, anchor="nw", image=photo)
            canvas.configure(scrollregion=(0, 0, photo.width(), photo.height()))
        except Exception as exc:
            canvas.configure(scrollregion=(0, 0, 1, 1))
            canvas.create_text(8, 8, anchor="nw", fill="#cc6666", text=f"Unable to load image:\n{image_path}\n{exc}")

    def refresh_vb_image() -> None:
        comp_id = current_comp_state.get("id")
        path = resolve_panel_image(comp_id, is_vb6=True)
        load_canvas_image(vb_image_canvas, vb_image_state, path, "No VB6 image found for selected component.")
        if path:
            status_var.set(f"VB6 image: {path}")

    def refresh_dotnet_image() -> None:
        comp_id = current_comp_state.get("id")
        path = resolve_panel_image(comp_id, is_vb6=False)
        load_canvas_image(dotnet_image_canvas, dotnet_image_state, path, "No .NET image found for selected component.")
        if path:
            status_var.set(f".NET image: {path}")

    def set_vb_view(mode: str) -> None:
        vb_view_mode.set(mode)
        if mode == "image":
            vb_source_container.pack_forget()
            vb_image_container.pack(fill=tk.BOTH, expand=True)
            vb_toggle_btn.configure(text="Show Source")
            refresh_vb_image()
        else:
            vb_image_container.pack_forget()
            vb_source_container.pack(fill=tk.BOTH, expand=True)
            vb_toggle_btn.configure(text="Show Image")

    def set_dotnet_view(mode: str) -> None:
        dotnet_view_mode.set(mode)
        if mode == "image":
            dotnet_source_container.pack_forget()
            dotnet_image_container.pack(fill=tk.BOTH, expand=True)
            dotnet_toggle_btn.configure(text="Show Source")
            refresh_dotnet_image()
        else:
            dotnet_image_container.pack_forget()
            dotnet_source_container.pack(fill=tk.BOTH, expand=True)
            dotnet_toggle_btn.configure(text="Show Image")

    def on_browse_vb_image() -> None:
        comp_id = current_comp_state.get("id")
        if not comp_id:
            status_var.set("Select a component before choosing a VB6 image.")
            return
        initial = str(resolve_image_dir(vb6_image_dir_var.get()))
        path = filedialog.askopenfilename(
            parent=root,
            title="Select VB6 screenshot",
            initialdir=initial if Path(initial).exists() else str(Path.home()),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("All files", "*.*")],
        )
        if not path:
            return
        vb_image_overrides[comp_id] = Path(path).expanduser().resolve()
        set_vb_view("image")

    def on_browse_dotnet_image() -> None:
        comp_id = current_comp_state.get("id")
        if not comp_id:
            status_var.set("Select a component before choosing a .NET image.")
            return
        initial = str(resolve_image_dir(dotnet_image_dir_var.get()))
        path = filedialog.askopenfilename(
            parent=root,
            title="Select .NET screenshot",
            initialdir=initial if Path(initial).exists() else str(Path.home()),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp *.webp"), ("All files", "*.*")],
        )
        if not path:
            return
        dotnet_image_overrides[comp_id] = Path(path).expanduser().resolve()
        set_dotnet_view("image")

    vb_header = ttk.Frame(vb_frame)
    vb_header.pack(fill=tk.X)
    ttk.Label(vb_header, text="VB6 Source").pack(side=tk.LEFT)
    vb_browse_btn = ttk.Button(vb_header, text="Image...", command=on_browse_vb_image)
    vb_browse_btn.pack(side=tk.RIGHT, padx=(6, 0))
    vb_toggle_btn = ttk.Button(
        vb_header,
        text="Show Image",
        command=lambda: set_vb_view("image" if vb_view_mode.get() == "source" else "source"),
    )
    vb_toggle_btn.pack(side=tk.RIGHT)

    vb_stack = ttk.Frame(vb_frame)
    vb_stack.pack(fill=tk.BOTH, expand=True)
    vb_source_container = ttk.Frame(vb_stack)
    vb_source_container.pack(fill=tk.BOTH, expand=True)
    vb_body = ttk.Frame(vb_source_container)
    vb_body.pack(fill=tk.BOTH, expand=True)
    vb_gutter = tk.Text(vb_body, width=6, padx=4, takefocus=0, borderwidth=0)
    vb_gutter.pack(side=tk.LEFT, fill=tk.Y)
    vb_text = tk.Text(vb_body, wrap="none")
    vb_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

    def on_vb_scroll(*args):
        vb_text.yview(*args)
        vb_gutter.yview(*args)

    vb_scroll_y = ttk.Scrollbar(vb_body, orient=tk.VERTICAL, command=on_vb_scroll)
    vb_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    vb_scroll_x = ttk.Scrollbar(vb_source_container, orient=tk.HORIZONTAL, command=vb_text.xview)
    vb_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    def on_vb_yscroll(first, last):
        vb_scroll_y.set(first, last)
        vb_gutter.yview_moveto(first)

    vb_text.configure(yscrollcommand=on_vb_yscroll, xscrollcommand=vb_scroll_x.set)

    vb_image_container = ttk.Frame(vb_stack)
    vb_image_canvas = tk.Canvas(vb_image_container, background="#101010", highlightthickness=0)
    vb_image_canvas.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    vb_image_scroll_y = ttk.Scrollbar(vb_image_container, orient=tk.VERTICAL, command=vb_image_canvas.yview)
    vb_image_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    vb_image_scroll_x = ttk.Scrollbar(vb_image_container, orient=tk.HORIZONTAL, command=vb_image_canvas.xview)
    vb_image_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    vb_image_canvas.configure(yscrollcommand=vb_image_scroll_y.set, xscrollcommand=vb_image_scroll_x.set)

    dotnet_header = ttk.Frame(blazor_frame)
    dotnet_header.pack(fill=tk.X)
    ttk.Label(dotnet_header, text=".NET Source").pack(side=tk.LEFT)
    dotnet_browse_btn = ttk.Button(dotnet_header, text="Image...", command=on_browse_dotnet_image)
    dotnet_browse_btn.pack(side=tk.RIGHT, padx=(6, 0))
    dotnet_toggle_btn = ttk.Button(
        dotnet_header,
        text="Show Image",
        command=lambda: set_dotnet_view("image" if dotnet_view_mode.get() == "source" else "source"),
    )
    dotnet_toggle_btn.pack(side=tk.RIGHT)

    dotnet_stack = ttk.Frame(blazor_frame)
    dotnet_stack.pack(fill=tk.BOTH, expand=True)
    dotnet_source_container = ttk.Frame(dotnet_stack)
    dotnet_source_container.pack(fill=tk.BOTH, expand=True)
    output_body = ttk.Frame(dotnet_source_container)
    output_body.pack(fill=tk.BOTH, expand=True)
    blazor_gutter = tk.Text(output_body, width=6, padx=4, takefocus=0, borderwidth=0)
    blazor_gutter.pack(side=tk.LEFT, fill=tk.Y)
    blazor_text = tk.Text(output_body, wrap="none")
    blazor_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

    def on_blazor_scroll(*args):
        blazor_text.yview(*args)
        blazor_gutter.yview(*args)

    blazor_scroll_y = ttk.Scrollbar(output_body, orient=tk.VERTICAL, command=on_blazor_scroll)
    blazor_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    blazor_scroll_x = ttk.Scrollbar(dotnet_source_container, orient=tk.HORIZONTAL, command=blazor_text.xview)
    blazor_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    def on_blazor_yscroll(first, last):
        blazor_scroll_y.set(first, last)
        blazor_gutter.yview_moveto(first)

    blazor_text.configure(yscrollcommand=on_blazor_yscroll, xscrollcommand=blazor_scroll_x.set)

    dotnet_image_container = ttk.Frame(dotnet_stack)
    dotnet_image_canvas = tk.Canvas(dotnet_image_container, background="#101010", highlightthickness=0)
    dotnet_image_canvas.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    dotnet_image_scroll_y = ttk.Scrollbar(dotnet_image_container, orient=tk.VERTICAL, command=dotnet_image_canvas.yview)
    dotnet_image_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    dotnet_image_scroll_x = ttk.Scrollbar(dotnet_image_container, orient=tk.HORIZONTAL, command=dotnet_image_canvas.xview)
    dotnet_image_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    dotnet_image_canvas.configure(yscrollcommand=dotnet_image_scroll_y.set, xscrollcommand=dotnet_image_scroll_x.set)

    # Compile errors panel (collapsible, below .NET source)
    error_frame = ttk.Frame(blazor_frame)
    error_header = ttk.Frame(error_frame)
    error_header.pack(fill=tk.X)
    error_title = ttk.Label(error_header, text="Compile Errors")
    error_title.pack(side=tk.LEFT)
    toggle_error_btn = ttk.Button(error_header, text="Hide")
    toggle_error_btn.pack(side=tk.RIGHT)

    error_text = tk.Text(error_frame, wrap="none", height=10)
    error_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
    error_scroll_y = ttk.Scrollbar(error_frame, orient=tk.VERTICAL, command=error_text.yview)
    error_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
    error_scroll_x = ttk.Scrollbar(error_frame, orient=tk.HORIZONTAL, command=error_text.xview)
    error_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
    error_text.configure(yscrollcommand=error_scroll_y.set, xscrollcommand=error_scroll_x.set)
    mark_chunk("chunk7_right_panes", last_time)

    mono = tkfont.Font(family="Menlo", size=11)
    # VB6 pane: classic Visual Studio colors (light theme)
    vb_text.configure(
        font=mono,
        background="#ffffff",
        foreground="#000000",
        insertbackground="#000000",
        selectbackground="#cfe8ff",
        selectforeground="#000000",
    )
    vb_gutter.configure(font=mono, background="#f3f3f3", foreground="#666666", state="disabled")
    # VS Code Dark+ style for .NET output panel
    blazor_text.configure(
        font=mono,
        background="#1e1e1e",
        foreground="#d4d4d4",
        insertbackground="#d4d4d4",
        selectbackground="#264f78",
        selectforeground="#ffffff",
    )
    blazor_gutter.configure(font=mono, background="#1e1e1e", foreground="#6b6b6b", state="disabled")
    error_text.configure(font=mono)

    # Syntax highlighting tags
    # VB6 tag colors (classic Visual Studio)
    vb_text.tag_configure("kw", foreground="#0000ff")
    vb_text.tag_configure("str", foreground="#800000")
    vb_text.tag_configure("cmt", foreground="#008000")
    vb_text.tag_configure("hdr", foreground="#555555")

    error_text.tag_configure("kw", foreground="#1f5fbf")
    error_text.tag_configure("str", foreground="#a31515")
    error_text.tag_configure("cmt", foreground="#008000")
    error_text.tag_configure("hdr", foreground="#555555")

    # VS Code Dark+ palette
    blazor_text.tag_configure("kw", foreground="#569cd6")
    blazor_text.tag_configure("str", foreground="#ce9178")
    blazor_text.tag_configure("cmt", foreground="#6a9955")
    blazor_text.tag_configure("hdr", foreground="#9aa0a6")
    blazor_text.tag_configure("razor", foreground="#c586c0")
    blazor_text.tag_configure("num", foreground="#b5cea8")
    blazor_text.tag_configure("type", foreground="#4ec9b0")
    blazor_text.tag_configure("func", foreground="#dcdcaa")
    blazor_text.tag_configure("htmltag", foreground="#569cd6")
    blazor_text.tag_configure("attr", foreground="#d7ba7d")
    blazor_text.tag_configure("err", background="#ffd6d6")
    error_text.tag_configure("filehdr", foreground="#1f5fbf")
    vb_text.tag_configure("find_match", background="#ffe58f")
    blazor_text.tag_configure("find_match", background="#ffe58f")

    MAX_HIGHLIGHT_CHARS = 200000
    MAX_HIGHLIGHT_LINES = 4000
    MAX_GUTTER_LINES = 2000
    MAX_SEARCH_MATCHES = 200

    vb_keywords = {
        "if","then","else","elseif","end","sub","function","dim","set","new","for","next","while",
        "wend","do","loop","until","select","case","exit","on","error","resume","goto","call","with",
        "me","private","public","friend","property","get","let","type","enum","const","as","byval",
        "byref","optional","not","and","or","xor","mod","true","false","nothing","is","like","to",
        "each","in","stop","debug","print","input","open","close"
    }

    def update_gutter(gutter: tk.Text, content: str):
        line_count = content.count("\n") + 1
        if line_count > MAX_GUTTER_LINES:
            gutter.configure(state="normal")
            gutter.delete("1.0", tk.END)
            gutter.insert("1.0", "...")
            gutter.configure(state="disabled")
            return
        numbers = "\n".join(str(i) for i in range(1, line_count + 1))
        gutter.configure(state="normal")
        gutter.delete("1.0", tk.END)
        gutter.insert("1.0", numbers)
        gutter.configure(state="disabled")

    # ── In-panel Find (Ctrl/Cmd+F) ─────────────────────────────────────
    find_state: dict[int, dict] = {}
    last_find_target = {"widget": vb_text}

    def open_find_popup(target: tk.Text, title: str):
        widget_id = id(target)
        state = find_state.get(widget_id)
        if state and state.get("popup") and state["popup"].winfo_exists():
            popup = state["popup"]
            popup.deiconify()
            popup.lift()
            state["entry"].focus_set()
            return

        popup = tk.Toplevel(root)
        popup.title(f"Find — {title}")
        popup.geometry("420x120")
        popup.transient(root)

        container = ttk.Frame(popup, padding=8)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Find").grid(row=0, column=0, sticky="w")
        pattern_var = tk.StringVar()
        entry = ttk.Entry(container, textvariable=pattern_var, width=40)
        entry.grid(row=0, column=1, columnspan=3, sticky="we", padx=(6, 0))

        status = ttk.Label(container, text="", foreground="#666666")
        status.grid(row=1, column=0, columnspan=4, sticky="w", pady=(4, 0))

        btn_search = ttk.Button(container, text="Search")
        btn_cancel = ttk.Button(container, text="Cancel", command=popup.destroy)
        btn_next = ttk.Button(container, text="Next")

        btn_search.grid(row=2, column=1, sticky="e", pady=(8, 0))
        btn_next.grid(row=2, column=2, sticky="e", padx=(6, 0), pady=(8, 0))
        btn_cancel.grid(row=2, column=3, sticky="e", padx=(6, 0), pady=(8, 0))

        btn_next.grid_remove()

        container.columnconfigure(1, weight=1)

        def clear_highlight():
            target.tag_remove("find_match", "1.0", tk.END)

        def do_search(start_index: str = "1.0"):
            pattern = pattern_var.get()
            if not pattern:
                status.configure(text="Enter a search term.")
                return None
            clear_highlight()
            idx = target.search(pattern, start_index, tk.END, nocase=True)
            if not idx:
                status.configure(text="No match.")
                return None
            end = f"{idx}+{len(pattern)}c"
            target.tag_add("find_match", idx, end)
            target.see(idx)
            target.mark_set(tk.INSERT, end)
            target.focus_set()
            status.configure(text=f"Found at {idx}")
            return end

        def on_search():
            end = do_search("1.0")
            if end:
                btn_next.grid()
                find_state[widget_id]["last_index"] = end

        def on_next():
            last_index = find_state[widget_id].get("last_index") or "1.0"
            end = do_search(last_index)
            if end:
                find_state[widget_id]["last_index"] = end

        btn_search.configure(command=on_search)
        btn_next.configure(command=on_next)
        popup.bind("<Return>", lambda _e: on_search())
        popup.bind("<Escape>", lambda _e: popup.destroy())

        find_state[widget_id] = {
            "popup": popup,
            "entry": entry,
            "last_index": None,
        }

        entry.focus_set()

    def bind_find_shortcuts(widget: tk.Text, title: str):
        widget.bind("<Control-f>", lambda _e: open_find_popup(widget, title))
        widget.bind("<Control-F>", lambda _e: open_find_popup(widget, title))
        widget.bind("<Command-f>", lambda _e: open_find_popup(widget, title))
        widget.bind("<Command-F>", lambda _e: open_find_popup(widget, title))
        widget.bind("<FocusIn>", lambda _e: last_find_target.__setitem__("widget", widget))

    bind_find_shortcuts(vb_text, "VB6 Source")
    bind_find_shortcuts(blazor_text, ".NET Output")

    def open_find_for_last():
        target = last_find_target.get("widget") or vb_text
        title = "VB6 Source" if target is vb_text else ".NET Output"
        open_find_popup(target, title)

    root.bind("<Control-f>", lambda _e: open_find_for_last())
    root.bind("<Control-F>", lambda _e: open_find_for_last())
    root.bind("<Command-f>", lambda _e: open_find_for_last())
    root.bind("<Command-F>", lambda _e: open_find_for_last())

    def highlight_vb(text_widget: tk.Text, content: str):
        text_widget.delete("1.0", tk.END)
        text_widget.insert("1.0", content)
        update_gutter(vb_gutter, content)
        if len(content) > MAX_HIGHLIGHT_CHARS or content.count("\n") + 1 > MAX_HIGHLIGHT_LINES:
            return

        # Comments: lines starting with ' or Rem
        for i, line in enumerate(content.splitlines(), start=1):
            stripped = line.lstrip()
            if stripped.startswith("'") or stripped.lower().startswith("rem "):
                text_widget.tag_add("cmt", f"{i}.0", f"{i}.end")

        # Strings
        for m in re.finditer(r'\"([^\"\n]|\"\")*\"', content):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("str", start, end)

        # Keywords
        for m in re.finditer(r"\b[A-Za-z_][A-Za-z0-9_]*\b", content):
            word = m.group(0).lower()
            if word in vb_keywords:
                start = f"1.0 + {m.start()} chars"
                end = f"1.0 + {m.end()} chars"
                text_widget.tag_add("kw", start, end)

    cs_keywords = {
        "using","namespace","class","struct","interface","enum","record","public","private","protected",
        "internal","static","void","int","string","bool","var","new","return","if","else","switch",
        "case","break","continue","for","foreach","while","do","try","catch","finally","throw","async",
        "await","Task","partial","get","set","true","false","null","this","base","typeof","default",
        "in","out","ref","is","as","operator","override","virtual","sealed","readonly","const"
    }

    def highlight_blazor(text_widget: tk.Text, content: str):
        text_widget.delete("1.0", tk.END)
        text_widget.insert("1.0", content)
        for tag in ("err", "kw", "str", "cmt", "razor", "num", "type", "func", "htmltag", "attr"):
            text_widget.tag_remove(tag, "1.0", tk.END)
        update_gutter(blazor_gutter, content)
        if len(content) > MAX_HIGHLIGHT_CHARS or content.count("\n") + 1 > MAX_HIGHLIGHT_LINES:
            return

        # Razor comments
        for m in re.finditer(r"@\*.*?\*@", content, flags=re.S):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("cmt", start, end)

        # C# line comments
        for m in re.finditer(r"//.*", content):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("cmt", start, end)

        # C# block comments
        for m in re.finditer(r"/\*.*?\*/", content, flags=re.S):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("cmt", start, end)

        # Strings
        for m in re.finditer(r'"([^"\n]|\\")*"', content):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("str", start, end)

        # Numbers
        for m in re.finditer(r"\b\d+(?:\.\d+)?\b", content):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("num", start, end)

        # Razor directives
        for m in re.finditer(r"@\w+", content):
            start = f"1.0 + {m.start()} chars"
            end = f"1.0 + {m.end()} chars"
            text_widget.tag_add("razor", start, end)

        # HTML tag names
        for m in re.finditer(r"</?\s*([A-Za-z0-9_:\-]+)", content):
            start = f"1.0 + {m.start(1)} chars"
            end = f"1.0 + {m.end(1)} chars"
            text_widget.tag_add("htmltag", start, end)

        # HTML attribute names
        for m in re.finditer(r"\b([A-Za-z_:][A-Za-z0-9_:\-]*)(?==)", content):
            start = f"1.0 + {m.start(1)} chars"
            end = f"1.0 + {m.end(1)} chars"
            text_widget.tag_add("attr", start, end)

        # Keywords
        for m in re.finditer(r"\b[A-Za-z_][A-Za-z0-9_]*\b", content):
            word = m.group(0)
            if word in cs_keywords:
                start = f"1.0 + {m.start()} chars"
                end = f"1.0 + {m.end()} chars"
                text_widget.tag_add("kw", start, end)

        # Type declarations
        for m in re.finditer(r"\b(class|struct|interface|enum|record)\s+([A-Za-z_][A-Za-z0-9_]*)", content):
            start = f"1.0 + {m.start(2)} chars"
            end = f"1.0 + {m.end(2)} chars"
            text_widget.tag_add("type", start, end)

        # Function identifiers (simple heuristic)
        for m in re.finditer(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*(?=\()", content):
            start = f"1.0 + {m.start(1)} chars"
            end = f"1.0 + {m.end(1)} chars"
            text_widget.tag_add("func", start, end)
    mark_chunk("chunk8_highlighting", last_time)

    def show_vb_only(comp_id: str):
        comp = id_to_component.get(comp_id)
        if not comp:
            return
        vb_content = get_chunked_source(comp_id, comp)
        highlight_vb(vb_text, vb_content)
        vb_text.tag_add("hdr", "1.0", "2.0")

    def show_error_in_panels(file_path: Path, line_no: int):
        nonlocal current_output_path, blazor_header_lines
        # Load VB6 (best effort)
        comp_id = comp_id_by_name.get(file_path.stem.lower())
        if comp_id:
            current_comp_state["id"] = comp_id
            show_vb_only(comp_id)
            try:
                vb_text.see(f"{line_no + 2}.0")
            except Exception:
                pass
        else:
            highlight_vb(vb_text, f"No VB6 source found for {file_path.name}")

        # Load .NET file
        try:
            content = file_path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            content = ""
        header_text = f"{file_path.name} :: {file_path}\n\n"
        blazor_header_lines = header_text.count("\\n")
        highlight_blazor(blazor_text, header_text + content)
        blazor_text.tag_add("hdr", "1.0", "2.0")
        current_output_path = file_path
        try:
            blazor_text.see(f"{line_no + blazor_header_lines}.0")
        except Exception:
            pass
        if vb_view_mode.get() == "image":
            refresh_vb_image()
        if dotnet_view_mode.get() == "image":
            refresh_dotnet_image()

    def show_error_text():
        if error_text.winfo_ismapped():
            return
        error_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        error_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        error_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    def set_error_panel_visible(visible: bool):
        nonlocal error_panel_visible
        if visible == error_panel_visible:
            return
        error_panel_visible = visible
        if visible:
            if not error_frame.winfo_ismapped():
                error_frame.pack(fill=tk.BOTH, expand=False, pady=(6, 0))
            toggle_error_btn.configure(text="Hide")
            show_error_text()
        else:
            if error_frame.winfo_ismapped():
                error_frame.pack_forget()
            toggle_error_btn.configure(text="Show")

    def on_toggle_error():
        set_error_panel_visible(not error_panel_visible)

    toggle_error_btn.configure(command=on_toggle_error)

    def update_error_panel_title():
        if error_panel_mode == "run":
            error_title.configure(text="Run Log (tail)")
        else:
            error_title.configure(text="Compile Errors")
        show_error_text()

    def render_run_log():
        if error_panel_mode != "run":
            return
        lines: List[str] = []
        if run_log_header_lines:
            lines.extend(run_log_header_lines)
        if run_tail_buffer.strip():
            if lines:
                lines.append("")
            lines.append(run_tail_buffer.rstrip())
        if not lines:
            lines = ["Waiting for run output..."]
        text = "\n".join(lines).rstrip() + "\n"
        error_text.delete("1.0", tk.END)
        error_text.insert("1.0", text)

    def launch_run_process(project_path: Path, is_retry: bool = False, retry_reason: str = "", killed_pids: List[int] | None = None):
        nonlocal run_process, run_project_path, run_log_path, run_tail_offset, run_tail_buffer
        nonlocal run_log_header_lines, run_url_announced, run_exit_announced, run_session_id
        nonlocal error_panel_mode
        if run_process is not None and run_process.poll() is None:
            _terminate_process(run_process)
        log_dir = Path(output_dir_var.get()).expanduser()
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "run.log"
        run_project_path = project_path
        run_log_path = log_path
        run_browser.reset()
        run_tail_offset = 0
        run_tail_buffer = ""
        run_url_announced = ""
        run_exit_announced = False
        run_session_id = f"{time.time():.6f}:{run_restart_count}"
        run_port = _pick_free_local_port()
        requested_url = f"http://127.0.0.1:{run_port}"
        log_file = log_path.open("w", encoding="utf-8")
        run_env = os.environ.copy()
        run_env["ASPNETCORE_URLS"] = requested_url
        run_env["ASPNETCORE_ENVIRONMENT"] = "Development"
        run_env["DOTNET_ENVIRONMENT"] = "Development"
        run_process = subprocess.Popen(
            ["dotnet", "run", "--no-launch-profile", "--project", str(project_path)],
            cwd=str(project_path.parent),
            stdout=log_file,
            stderr=subprocess.STDOUT,
            env=run_env,
            start_new_session=True,
        )
        log_file.close()
        run_log_header_lines = [
            f"Running... PID {run_process.pid}",
            f"URL: {requested_url} (starting...)",
            f"Log: {log_path}",
        ]
        if is_retry:
            note = f"Auto-retry {run_restart_count}/{run_max_restarts}"
            if retry_reason:
                note += f": {retry_reason}"
            run_log_header_lines.insert(0, note)
            if killed_pids:
                run_log_header_lines.append(
                    "Killed listener PID(s): " + ", ".join(str(p) for p in killed_pids)
                )
        status_var.set(f"Program started: {requested_url}")
        error_panel_mode = "run"
        update_error_panel_title()
        set_error_panel_visible(True)
        render_run_log()
        start_tail()

    def start_tail():
        nonlocal tail_job_id, run_tail_offset, run_tail_buffer, run_url_announced, run_exit_announced
        nonlocal run_restart_count
        if run_log_path is None:
            return
        if tail_job_id is not None:
            return
        session_snapshot = run_session_id

        def read_new_run_log() -> bool:
            nonlocal run_tail_offset, run_tail_buffer, run_url_announced
            if run_log_path is None or not run_log_path.exists():
                return False
            chunk = ""
            with run_log_path.open("r", encoding="utf-8", errors="ignore") as f:
                f.seek(run_tail_offset)
                chunk = f.read()
                run_tail_offset = f.tell()
            if not chunk:
                return False
            run_tail_buffer += chunk
            if len(run_tail_buffer) > 20000:
                run_tail_buffer = run_tail_buffer[-20000:]
            opened_url = run_browser.maybe_open_from_text(chunk)
            if opened_url and opened_url != run_url_announced:
                run_url_announced = opened_url
                url_line = f"URL: {opened_url}"
                if len(run_log_header_lines) >= 2 and run_log_header_lines[1].startswith("URL:"):
                    run_log_header_lines[1] = url_line
                else:
                    run_log_header_lines.insert(1, url_line)
                status_var.set(f"Opened browser: {opened_url}")
            return True

        def tick():
            nonlocal tail_job_id, run_exit_announced
            if run_session_id != session_snapshot:
                tail_job_id = None
                return
            if run_log_path is None:
                tail_job_id = None
                return
            try:
                updated = read_new_run_log()
                if updated:
                    render_run_log()
            except Exception:
                pass

            if run_process is not None and run_process.poll() is None:
                tail_job_id = root.after(1000, tick)
            else:
                # one final read after process ends
                tail_job_id = None
                if run_process is not None and run_process.poll() is not None:
                    try:
                        if read_new_run_log():
                            render_run_log()
                    except Exception:
                        pass
                    if not run_exit_announced:
                        exit_code = run_process.poll()
                        bind_port = _extract_bind_fail_port(run_tail_buffer)
                        has_bind_error = bind_port > 0 or "address already in use" in run_tail_buffer.lower()
                        if (
                            exit_code not in (0, None)
                            and has_bind_error
                            and run_restart_count < run_max_restarts
                            and run_project_path is not None
                        ):
                            run_restart_count += 1
                            killed_pids = _kill_listeners_on_port(bind_port) if bind_port > 0 else []
                            port_label = str(bind_port) if bind_port > 0 else "unknown"
                            run_log_header_lines.append(
                                f"Detected bind error on port {port_label}. Relaunching..."
                            )
                            if bind_port > 0 and not killed_pids:
                                run_log_header_lines.append(
                                    f"No external listener found on {bind_port}; retrying on a new port."
                                )
                            status_var.set("Bind error detected; relaunching...")
                            run_exit_announced = True
                            render_run_log()
                            retry_project = run_project_path
                            root.after(
                                300,
                                lambda p=retry_project, k=killed_pids, b=bind_port: launch_run_process(
                                    p,
                                    is_retry=True,
                                    retry_reason=f"bind failure on {b}" if b > 0 else "address already in use",
                                    killed_pids=k,
                                ),
                            )
                            return

                        run_exit_announced = True
                        if exit_code == 0:
                            run_log_header_lines.append("Program exited (code 0).")
                            status_var.set("Program exited (code 0).")
                        else:
                            run_log_header_lines.append(f"Program exited with errors (code {exit_code}).")
                            status_var.set(f"Program exited with errors (code {exit_code}).")
                            last_error_var.set(run_tail_buffer[-2000:].strip())
                        render_run_log()

        tail_job_id = root.after(500, tick)

    # Tree icons removed (simple tree)

    item_to_component: Dict[str, str] = {}
    last_click_element = ""

    def node_label(comp_id: str, ambig: bool, cycle: bool = False) -> str:
        label = comp_id
        if ambig:
            label += " [ambig]"
        if cycle:
            label += " [cycle]"
        return label

    def insert_node_simple(parent_id: str, comp_id: str, ambig: bool, cycle: bool) -> str:
        label = node_label(comp_id, ambig, cycle)
        tags = ("orphan",) if comp_id in orphan_set else ()
        item_id = tree.insert(parent_id, "end", text=label, open=True, tags=tags)
        item_to_component[item_id] = comp_id
        return item_id

    # Root selection: Sub Main roots + fallback
    root_nodes = roots[:]
    if "HFSystem::FLogin" not in root_nodes and "HFSystem::FLogin" in adj:
        root_nodes.append("HFSystem::FLogin")

    from collections import deque
    build_queue = deque()
    for r in root_nodes:
        build_queue.append(("", r, r.split("::")[1] in duplicates, False, []))

    def process_tree_batch():
        start = time.perf_counter()
        processed = 0
        while build_queue and (time.perf_counter() - start) < 0.01:
            parent_id, comp_id, ambig, cycle, path_stack = build_queue.popleft()
            item_id = insert_node_simple(parent_id, comp_id, ambig, cycle)
            if not cycle:
                new_stack = path_stack + [comp_id]
                for child_id, child_ambig in adj.get(comp_id, []):
                    is_cycle = child_id in new_stack
                    build_queue.append((item_id, child_id, child_ambig, is_cycle, new_stack))
            processed += 1
        if build_queue:
            if processed > 0:
                status_var.set(f"Building tree... {len(build_queue)} remaining")
            root.after(5, process_tree_batch)
        else:
            status_var.set("Ready")
            mark_chunk("chunk9_tree_build", last_time)

    root.after(50, process_tree_batch)

    chunks_dir = cache_dir / "chunks"
    chunks_dir.mkdir(parents=True, exist_ok=True)
    chunks_enabled = bool(cache_index.get("chunks_enabled", True))

    def chunk_cache_path(path: Path) -> Path:
        h = hashlib.sha1(str(path).encode("utf-8", errors="ignore")).hexdigest()
        return chunks_dir / f"{h}.txt"

    def get_chunked_source(comp_id: str, comp: Component) -> str:
        nonlocal total_lines
        if comp_id in source_cache:
            return source_cache[comp_id]
        path_key = str(comp.path)
        try:
            stat = comp.path.stat()
        except Exception:
            stat = None
        entry = files_index.get(path_key)
        cache_path = chunk_cache_path(comp.path)
        if (
            entry
            and stat is not None
            and entry.get("mtime") == stat.st_mtime
            and entry.get("size") == stat.st_size
            and cache_path.exists()
        ):
            try:
                text = cache_path.read_text(encoding="utf-8", errors="ignore")
                source_cache[comp_id] = text
                return text
            except Exception:
                pass

        # Load from disk and build chunked view only when enabled
        src = read_text(comp.path)
        if chunks_enabled:
            display_src = chunk_vb6_source(comp, src)
            text = f"{comp_id} :: {comp.path}\n\n{display_src}"
            try:
                cache_path.write_text(text, encoding="utf-8")
            except Exception:
                pass
        else:
            text = f"{comp_id} :: {comp.path}\n\n{src}"
        line_count = len(src.splitlines())
        if stat is not None:
            entry = {
                "mtime": stat.st_mtime,
                "size": stat.st_size,
                "line_count": line_count,
            }
            files_index[path_key] = entry
            save_cache_index(cache_dir, cache_index)
        if line_count_by_id.get(comp_id) is None:
            line_count_by_id[comp_id] = line_count
            total_lines += line_count
            refresh_stats_labels()
            row_id = order_row_by_comp_id.get(comp_id)
            if row_id:
                order_tree.set(row_id, "lines", str(line_count))
        source_cache[comp_id] = text
        return text

    def on_click(event):
        nonlocal last_click_element
        element = tree.identify("element", event.x, event.y) or ""
        last_click_element = element

    def show_component(comp_id: str):
        nonlocal current_output_path, blazor_header_lines
        comp = id_to_component.get(comp_id)
        if not comp:
            return
        current_comp_state["id"] = comp_id
        vb_content = get_chunked_source(comp_id, comp)
        highlight_vb(vb_text, vb_content)
        vb_text.tag_add("hdr", "1.0", "2.0")
        output_name_var.set(f"{comp.name}.razor")

        cached_blazor = blazor_cache.get(comp_id)
        candidate = converted_map.get(comp_id)
        shown_blazor = False
        if cached_blazor:
            blazor_header_lines = 2
            highlight_blazor(blazor_text, cached_blazor)
            blazor_text.tag_add("hdr", "1.0", "2.0")
            current_output_path = candidate
            shown_blazor = True
            if candidate and candidate.exists():
                converted_rows.add(comp_id)
                row_id = order_row_by_comp_id.get(comp_id)
                if row_id:
                    values = list(order_tree.item(row_id, "values"))
                    if values:
                        values[0] = "✔"
                        order_tree.item(row_id, values=values)
        else:
            if candidate is None:
                out_dir = Path(output_dir_var.get()).expanduser()
                candidate = out_dir / f"{comp.name}.razor"
            if candidate and candidate.exists():
                try:
                    blazor_content = candidate.read_text(encoding="utf-8", errors="ignore")
                    blazor_header = f"{candidate.name} :: {candidate}\n\n"
                    blazor_header_lines = blazor_header.count("\n")
                    blazor_cache[comp_id] = blazor_header + blazor_content
                    highlight_blazor(blazor_text, blazor_header + blazor_content)
                    blazor_text.tag_add("hdr", "1.0", "2.0")
                    current_output_path = candidate
                    shown_blazor = True
                    converted_rows.add(comp_id)
                    row_id = order_row_by_comp_id.get(comp_id)
                    if row_id:
                        values = list(order_tree.item(row_id, "values"))
                        if values:
                            values[0] = "✔"
                            order_tree.item(row_id, values=values)
                except Exception:
                    pass
        if not shown_blazor:
            highlight_blazor(blazor_text, "Not converted yet.")
            current_output_path = None
            blazor_header_lines = 0
        if vb_view_mode.get() == "image":
            refresh_vb_image()
        if dotnet_view_mode.get() == "image":
            refresh_dotnet_image()

    def on_select(_event):
        nonlocal last_click_element, current_output_path, blazor_header_lines
        if last_click_element in ("indicator", "image"):
            last_click_element = ""
            return
        last_click_element = ""
        sel = tree.selection()
        if not sel:
            return
        item_id = sel[0]
        comp_id = item_to_component.get(item_id)
        if not comp_id:
            return
        print(f"[select] Tree selected: {comp_id}")
        show_component(comp_id)

    def on_order_select(_event):
        sel = order_tree.selection()
        if not sel:
            return
        row_id = sel[0]
        comp_id = order_comp_by_row_id.get(row_id)
        if comp_id:
            print(f"[select] Grid selected: {comp_id}")
            show_component(comp_id)

    def get_selected_comp_id() -> str | None:
        if view_mode_var.get() == "grid":
            sel = order_tree.selection()
            if not sel:
                return None
            row_id = sel[0]
            return order_comp_by_row_id.get(row_id)
        sel = tree.selection()
        if not sel:
            return None
        item_id = sel[0]
        return item_to_component.get(item_id)

    def on_convert():
        nonlocal current_output_path, blazor_header_lines, error_panel_mode
        comp_id = get_selected_comp_id()
        if not comp_id:
            status_var.set("Select a node to convert.")
            return
        comp = id_to_component.get(comp_id)
        if not comp:
            status_var.set("Invalid component selection.")
            return

        out_dir = Path(output_dir_var.get()).expanduser()
        out_name = output_name_var.get().strip() or f"{comp.name}.razor"
        if "." not in out_name:
            out_name = out_name + ".razor"
        project_path = Path(project_var.get()).expanduser()
        model = model_var.get().strip() or "gpt-4.1"
        try:
            temperature = float(temp_var.get().strip() or "0.2")
        except ValueError:
            temperature = 0.2

        instructions = instructions_text.get("1.0", tk.END).strip()

        api_label = "Claude" if is_claude_model(model) else "Codex"
        status_var.set(f"Converting with {api_label}...")
        root.update_idletasks()

        def work():
            vb6_source = read_text(comp.path)
            context_blob = get_context_blob_for(context_var.get().strip())
            convert_fn = call_claude_convert if is_claude_model(model) else call_codex_convert
            result = convert_fn(
                vb6_source,
                comp.path,
                out_name,
                model,
                temperature,
                instructions,
                context_blob,
            )
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / out_name
            out_path.write_text(result, encoding="utf-8")
            ensure_csproj_includes(project_path, out_path)
            header = f"{out_name} :: {out_path}\n\n"
            return {
                "out_path": out_path,
                "header": header,
                "result": result,
                "comp_id": comp_id,
            }

        def success(payload):
            nonlocal current_output_path, blazor_header_lines
            header = payload["header"]
            result = payload["result"]
            out_path = payload["out_path"]
            comp_id_local = payload["comp_id"]
            blazor_header_lines = header.count("\n")
            highlight_blazor(blazor_text, header + result)
            blazor_text.tag_add("hdr", "1.0", "2.0")
            current_output_path = out_path
            converted_map[comp_id_local] = out_path
            blazor_cache[comp_id_local] = header + result
            converted_rows.add(comp_id_local)
            row_id = order_row_by_comp_id.get(comp_id_local)
            if row_id:
                values = list(order_tree.item(row_id, "values"))
                if values:
                    values[0] = "✔"
                    order_tree.item(row_id, values=values)
            status_var.set(f"Converted and saved to {out_path}")
            last_error_var.set("")
            error_text.delete("1.0", tk.END)
            error_panel_mode = "build"
            update_error_panel_title()

        def error(exc: Exception):
            nonlocal error_panel_mode
            err_msg = f"Conversion failed: {exc}"
            status_var.set(err_msg)
            last_error_var.set(err_msg)
            highlight_blazor(blazor_text, f"{err_msg}")
            error_text.delete("1.0", tk.END)
            error_text.insert("1.0", err_msg)
            error_panel_mode = "build"
            update_error_panel_title()

        run_async("Convert", work, success, error)

    def on_compile():
        nonlocal blazor_header_lines, error_panel_mode
        project_path = Path(project_var.get()).expanduser()
        if not project_path.exists():
            status_var.set(f"Project not found: {project_path}")
            return

        set_error_panel_visible(True)
        status_var.set("Compiling...")
        root.update_idletasks()

        def work():
            result = subprocess.run(
                ["dotnet", "build", str(project_path)],
                cwd=str(project_path.parent),
                capture_output=True,
                text=True,
            )
            output = (result.stdout or "") + "\n" + (result.stderr or "")
            return {"returncode": result.returncode, "output": output}

        def success(payload):
            nonlocal error_panel_mode
            output = payload["output"]
            if payload["returncode"] == 0:
                status_var.set("Build succeeded.")
                last_error_var.set("")
                blazor_text.tag_remove("err", "1.0", tk.END)
                error_text.delete("1.0", tk.END)
                error_text.insert("1.0", "Build succeeded.")
                error_panel_mode = "build"
                update_error_panel_title()
                set_error_panel_visible(False)
                return

            error_regex_full = re.compile(
                r"^(?P<file>.+)\\((?P<line>\\d+),(?P<col>\\d+)\\):\\s+error\\s+(?P<code>[^:]+):\\s+(?P<msg>.+)$",
                re.IGNORECASE,
            )
            error_regex_line = re.compile(
                r"^(?P<file>.+)\\((?P<line>\\d+)\\):\\s+error\\s+(?P<code>[^:]+):\\s+(?P<msg>.+)$",
                re.IGNORECASE,
            )
            error_regex_simple = re.compile(
                r"^(?P<file>.+):\\s+error\\s+(?P<code>[^:]+):\\s+(?P<msg>.+)$",
                re.IGNORECASE,
            )
            any_match = False
            grouped: Dict[str, List[Tuple[int | None, str, Path | None]]] = {}
            for line in output.splitlines():
                stripped = line.strip()
                m = error_regex_full.match(stripped) or error_regex_line.match(stripped) or error_regex_simple.match(stripped)
                if m:
                    file_raw = m.group("file").strip()
                    file_path = None
                    try:
                        file_path = Path(file_raw).resolve()
                        file_label = file_path.name
                    except Exception:
                        file_label = Path(file_raw).name if file_raw else "<unknown>"
                    line_no = int(m.group("line")) if "line" in m.groupdict() and m.group("line") else None
                    msg = f"{m.group('code').strip()}: {m.group('msg').strip()}"
                    grouped.setdefault(file_label, []).append((line_no, msg, file_path))
                    if (
                        line_no is not None
                        and current_output_path
                        and file_path
                        and file_path == current_output_path.resolve()
                    ):
                        any_match = True
                        hl_line = line_no + blazor_header_lines
                        blazor_text.tag_add("err", f"{hl_line}.0", f"{hl_line}.end")
                    continue
                if " error " in stripped.lower():
                    grouped.setdefault("Other", []).append((None, stripped, None))

            status_var.set("Build failed. Errors highlighted." if any_match else "Build failed. See output.")
            last_error_var.set(output.strip()[:2000])
            error_text.delete("1.0", tk.END)
            if grouped:
                for file_label, items in grouped.items():
                    error_text.insert(tk.END, f"{file_label}\\n", ("filehdr",))
                    for line_no, msg, _file_path in items:
                        prefix = f"  L{line_no}: " if line_no else "  "
                        error_text.insert(tk.END, f"{prefix}{msg}\\n")
                    error_text.insert(tk.END, "\\n")
            else:
                error_text.insert("1.0", output.strip())
            error_panel_mode = "build"
            update_error_panel_title()

        def error(exc: Exception):
            err_msg = f"Compile failed: {exc}"
            status_var.set(err_msg)
            last_error_var.set(err_msg)

        run_async("Compile", work, success, error)

    def convert_component(
        comp: Component,
        instructions: str,
        context_blob: str,
        out_dir: Path,
        project_path: Path,
        model: str,
        temperature: float,
    ):
        out_name = f"{comp.name}.razor"

        vb6_source = read_text(comp.path)
        convert_fn = call_claude_convert if is_claude_model(model) else call_codex_convert
        result = convert_fn(
            vb6_source,
            comp.path,
            out_name,
            model,
            temperature,
            instructions,
            context_blob,
        )

        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / out_name
        out_path.write_text(result, encoding="utf-8")
        ensure_csproj_includes(project_path, out_path)

        header = f"{out_name} :: {out_path}\n\n"
        return {
            "out_path": out_path,
            "header": header,
            "result": result,
            "comp_id": f"{comp.project}::{comp.name}",
        }

    def on_convert_checked():
        selected_rows = order_tree.selection()
        if not selected_rows:
            status_var.set("No selected items to convert.")
            return
        checked = [order_comp_by_row_id.get(rid) for rid in selected_rows]
        checked = [cid for cid in checked if cid]
        instructions = instructions_text.get("1.0", tk.END).strip()
        context_dir = context_var.get().strip()
        out_dir = Path(output_dir_var.get()).expanduser()
        project_path = Path(project_var.get()).expanduser()
        model = model_var.get().strip() or "gpt-5.2-codex"
        try:
            temperature = float(temp_var.get().strip() or "0.2")
        except ValueError:
            temperature = 0.2
        status_var.set(f"Converting {len(checked)} items...")
        root.update_idletasks()

        def work():
            context_blob = get_context_blob_for(context_dir)
            last_payload = None
            total = len(checked)
            pause = 180 if is_claude_model(model) else 5
            for idx, cid in enumerate(checked, start=1):
                comp = id_to_component.get(cid)
                if not comp:
                    continue
                # Wait between calls to avoid rate-limit (one session per file)
                if idx > 1:
                    root.after(
                        0,
                        lambda i=idx, t=total, n=comp.name, p=pause: status_var.set(
                            f"Waiting {p}s before {i}/{t}: {n}..."
                        ),
                    )
                    time.sleep(pause)
                root.after(
                    0,
                    lambda i=idx, t=total, n=comp.name: status_var.set(f"Converting {i}/{t}: {n}"),
                )
                last_payload = convert_component(
                    comp,
                    instructions,
                    context_blob,
                    out_dir,
                    project_path,
                    model,
                    temperature,
                )
            return last_payload

        def success(payload):
            nonlocal current_output_path, blazor_header_lines, error_panel_mode
            if payload:
                header = payload["header"]
                result = payload["result"]
                out_path = payload["out_path"]
                comp_id_local = payload["comp_id"]
                blazor_header_lines = header.count("\n")
                highlight_blazor(blazor_text, header + result)
                blazor_text.tag_add("hdr", "1.0", "2.0")
                current_output_path = out_path
                converted_map[comp_id_local] = out_path
                blazor_cache[comp_id_local] = header + result
                converted_rows.add(comp_id_local)
                row_id = order_row_by_comp_id.get(comp_id_local)
                if row_id:
                    values = list(order_tree.item(row_id, "values"))
                    if values:
                        values[0] = "✔"
                        order_tree.item(row_id, values=values)
                error_panel_mode = "build"
                update_error_panel_title()
            status_var.set(f"Converted {len(checked)} items.")

        def error(exc: Exception):
            err_msg = f"Conversion failed: {exc}"
            status_var.set(err_msg)
            last_error_var.set(err_msg)
            error_text.delete("1.0", tk.END)
            error_text.insert("1.0", err_msg)

        run_async("Convert Checked", work, success, error)

    def on_fix():
        nonlocal current_output_path, blazor_header_lines, error_panel_mode
        if current_output_path is None or not current_output_path.exists():
            status_var.set("No converted file to fix.")
            return
        errors = error_text.get("1.0", tk.END).strip() or last_error_var.get().strip()
        if not errors:
            status_var.set("No compile errors to fix.")
            return

        model = model_var.get().strip() or "gpt-4.1"
        try:
            temperature = float(temp_var.get().strip() or "0.2")
        except ValueError:
            temperature = 0.2

        api_label = "Claude" if is_claude_model(model) else "Codex"
        status_var.set(f"Fixing with {api_label}...")
        root.update_idletasks()

        def work():
            blazor_source = current_output_path.read_text(encoding="utf-8", errors="ignore")
            fix_fn = call_claude_fix if is_claude_model(model) else call_codex_fix
            fixed = fix_fn(blazor_source, errors, current_output_path, model, temperature)
            current_output_path.write_text(fixed, encoding="utf-8")
            header = f"{current_output_path.name} :: {current_output_path}\n\n"
            return {"header": header, "fixed": fixed}

        def success(payload):
            nonlocal error_panel_mode, blazor_header_lines
            header = payload["header"]
            fixed = payload["fixed"]
            blazor_header_lines = header.count("\n")
            highlight_blazor(blazor_text, header + fixed)
            blazor_text.tag_add("hdr", "1.0", "2.0")
            status_var.set("Fix applied.")
            last_error_var.set("")
            error_text.delete("1.0", tk.END)
            error_panel_mode = "build"
            update_error_panel_title()

        def error(exc: Exception):
            nonlocal error_panel_mode
            err_msg = f"Fix failed: {exc}"
            status_var.set(err_msg)
            last_error_var.set(err_msg)
            error_text.delete("1.0", tk.END)
            error_text.insert("1.0", err_msg)
            error_panel_mode = "build"
            update_error_panel_title()

        run_async("Fix", work, success, error)

    def on_align_screens():
        nonlocal current_output_path, blazor_header_lines, error_panel_mode
        comp_id = get_selected_comp_id()
        if not comp_id:
            status_var.set("Select a component to align.")
            return
        comp = id_to_component.get(comp_id)
        if not comp:
            status_var.set("Invalid component selection.")
            return

        vb_img = resolve_panel_image(comp_id, is_vb6=True)
        net_img = resolve_panel_image(comp_id, is_vb6=False)
        if vb_img is None or not vb_img.exists():
            status_var.set("VB6 screenshot not found. Configure VB6 image folder or browse image.")
            return
        if net_img is None or not net_img.exists():
            status_var.set(".NET screenshot not found. Configure .NET image folder or browse image.")
            return

        out_path = converted_map.get(comp_id)
        if out_path is None:
            out_path = Path(output_dir_var.get()).expanduser() / f"{comp.name}.razor"
        if not out_path.exists():
            status_var.set(f"Converted file not found for alignment: {out_path}")
            return

        model = model_var.get().strip() or "gpt-5.2-codex"
        try:
            temperature = float(temp_var.get().strip() or "0.2")
        except ValueError:
            temperature = 0.2

        api_label = "Claude" if is_claude_model(model) else "Codex"
        status_var.set(f"Running pixel pass + alignment with {api_label}...")
        root.update_idletasks()

        def work():
            metrics = run_pixel_pass(vb_img, net_img)
            report = build_alignment_report(vb_img, net_img, metrics)
            blazor_source = out_path.read_text(encoding="utf-8", errors="ignore")
            align_fn = call_claude_align if is_claude_model(model) else call_codex_align
            aligned = align_fn(blazor_source, out_path, report, model, temperature)
            out_path.write_text(aligned, encoding="utf-8")
            header = f"{out_path.name} :: {out_path}\n\n"
            return {
                "out_path": out_path,
                "header": header,
                "aligned": aligned,
                "metrics": metrics,
                "report": report,
                "comp_id": comp_id,
            }

        def success(payload):
            nonlocal current_output_path, blazor_header_lines, error_panel_mode
            header = payload["header"]
            aligned = payload["aligned"]
            metrics = payload["metrics"]
            report = payload["report"]
            comp_id_local = payload["comp_id"]
            blazor_header_lines = header.count("\n")
            highlight_blazor(blazor_text, header + aligned)
            blazor_text.tag_add("hdr", "1.0", "2.0")
            current_output_path = payload["out_path"]
            converted_map[comp_id_local] = payload["out_path"]
            blazor_cache[comp_id_local] = header + aligned
            if dotnet_view_mode.get() == "image":
                refresh_dotnet_image()
            err_lines = [
                "Alignment completed.",
                f"VB6 size: {metrics.get('vb6_size')}",
                f".NET size: {metrics.get('dotnet_size')}",
                f"MAE(0-255): {metrics.get('mae_255')}",
                f"MAE(%): {metrics.get('mae_pct')}",
                f"Changed pixels(%): {metrics.get('changed_pct')}",
                "",
                report,
            ]
            error_text.delete("1.0", tk.END)
            error_text.insert("1.0", "\n".join(err_lines))
            error_panel_mode = "build"
            update_error_panel_title()
            set_error_panel_visible(True)
            status_var.set("Alignment pass completed and Razor file updated.")
            last_error_var.set("")

        def error(exc: Exception):
            nonlocal error_panel_mode
            err_msg = f"Alignment failed: {exc}"
            status_var.set(err_msg)
            last_error_var.set(err_msg)
            error_text.delete("1.0", tk.END)
            error_text.insert("1.0", err_msg)
            error_panel_mode = "build"
            update_error_panel_title()
            set_error_panel_visible(True)

        run_async("Align Screens", work, success, error)

    def on_run():
        nonlocal run_restart_count, run_project_path, run_session_id
        project_path = Path(project_var.get()).expanduser()
        if not project_path.exists():
            status_var.set(f"Project not found: {project_path}")
            return
        if run_process and run_process.poll() is None:
            status_var.set("Program is already running.")
            return
        try:
            run_restart_count = 0
            run_project_path = project_path
            run_session_id = ""
            launch_run_process(project_path, is_retry=False)
        except Exception as exc:
            err_msg = f"Run failed: {exc}"
            status_var.set(err_msg)
            last_error_var.set(err_msg)
            error_text.delete("1.0", tk.END)
            error_text.insert("1.0", err_msg)
            error_panel_mode = "build"
            update_error_panel_title()

    def build_menu():
        menubar = tk.Menu(root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Project...", command=on_new_project)
        file_menu.add_command(label="Load Existing Project...", command=on_load_project)
        file_menu.add_command(label="Close Project", command=on_close_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=on_app_exit)
        menubar.add_cascade(label="File", menu=file_menu)

        project_menu = tk.Menu(menubar, tearoff=0)
        project_menu.add_command(label="Refresh Stats", command=on_refresh_stats)
        project_menu.add_command(label="Rebuild Cache", command=on_rebuild_cache)
        project_menu.add_command(label="Convert Selected", command=on_convert)
        project_menu.add_command(label="Convert Selected (Grid)", command=on_convert_checked)
        menubar.add_cascade(label="Project", menu=project_menu)

        build_menu = tk.Menu(menubar, tearoff=0)
        build_menu.add_command(label="Compile", command=on_compile)
        build_menu.add_command(label="Fix", command=on_fix)
        build_menu.add_command(label="Run", command=on_run)
        build_menu.add_command(label="Align Screens", command=on_align_screens)
        menubar.add_cascade(label="Build", menu=build_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Search", command=on_search)
        tools_menu.add_command(label="Tag Selected", command=on_tag)
        tools_menu.add_command(label="Select All (Grid)", command=on_select_all)
        tools_menu.add_command(label="Clear All (Grid)", command=on_clear_all)
        menubar.add_cascade(label="Tools", menu=tools_menu)

        config_menu = tk.Menu(menubar, tearoff=0)
        config_menu.add_command(label="Settings...", command=open_config_dialog)
        menubar.add_cascade(label="Configuration", menu=config_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=lambda: status_var.set("VB6 Migrator"))
        menubar.add_cascade(label="Help", menu=help_menu)

        root.config(menu=menubar)

    build_menu()

    convert_checked_btn.configure(command=on_convert_checked)
    search_btn.configure(command=on_search)
    align_btn.configure(command=on_align_screens)

    tree.bind("<Button-1>", on_click)
    tree.bind("<<TreeviewSelect>>", on_select)
    order_tree.bind("<<TreeviewSelect>>", on_order_select)
    # no checkbox toggle; selection drives conversion

    mark_chunk("chunk10_bindings", last_time)
    total_time = time.perf_counter() - t_start
    print("UI init timings:")
    for name, dur in chunk_times:
        print(f"  {name}: {dur:.3f}s")
    print(f"  total: {total_time:.3f}s")
    print("End of UI")

    root.deiconify()
    root.lift()
    try:
        root.focus_force()
    except Exception:
        pass
    root.protocol("WM_DELETE_WINDOW", on_app_exit)
    if sys.platform == "darwin":
        root.after(50, _activate_macos_app)
        root.after(100, root.lift)
        root.after(150, lambda: root.focus_force())
    try:
        root.mainloop()
    finally:
        stop_all_jobs()
    return ui_action


def load_project_runtime(project_dir: Path, prefer_hfest: bool) -> dict:
    layout = get_project_layout(project_dir)
    max_files = MAX_FILES_TO_LOAD if MAX_FILES_TO_LOAD and MAX_FILES_TO_LOAD > 0 else None
    all_components = load_project_components(
        layout["source_dir"],
        fallback_project_name=layout["name"],
        max_files=max_files,
    )
    if all_components:
        deduped_components, _dropped = dedupe_by_content(all_components, prefer_hfest=prefer_hfest)
        _, preferred, duplicates = build_indices(deduped_components, prefer_hfest=prefer_hfest)
        components = list(preferred.values())
    else:
        preferred = {}
        duplicates = set()
        components = []

    cache_dir = layout["cache_dir"]
    signature = compute_signature(components)
    tree_cache = load_tree_cache(cache_dir, signature, prefer_hfest)
    if tree_cache:
        edges: Dict[Tuple[str, str], Dict] = {}
        for caller, callee, ambig in tree_cache.get("edges", []):
            edges[(caller, callee)] = {"types": set(), "ambiguous": bool(ambig)}
        roots = tree_cache.get("roots", [])
    else:
        edges = build_edges(components, preferred, duplicates)
        roots = build_roots(components)
        save_tree_cache(cache_dir, signature, prefer_hfest, edges, roots)

    suffix = "prefer_hfest" if prefer_hfest else "combined"
    csv_path, mmd_path, text_path, migration_order, cycle_set, orphan_set = write_outputs(
        layout["graphs_dir"],
        edges,
        components,
        roots,
        suffix,
    )
    return {
        "layout": layout,
        "components": components,
        "edges": edges,
        "roots": roots,
        "duplicates": duplicates,
        "migration_order": migration_order,
        "cycle_set": cycle_set,
        "orphan_set": orphan_set,
        "csv_path": csv_path,
        "mmd_path": mmd_path,
        "text_path": text_path,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", help="Path to VB6 source folder (used to seed default project)")
    ap.add_argument("--project-dir", default="", help="Path to an existing packaged project folder")
    ap.add_argument(
        "--projects-root",
        default=str((Path(__file__).resolve().parent / "projects").resolve()),
        help="Folder where packaged projects are stored",
    )
    ap.add_argument("--project-name", default=DEFAULT_PROJECT_NAME, help="Default packaged project name")
    ap.add_argument("--prefer-hfest", default="true", choices=["true", "false"])
    ap.add_argument("--ui", default="true", choices=["true", "false"])
    ap.add_argument(
        "--output-project",
        default="/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/HomeFront.csproj",
        help="Path to HomeFront.csproj",
    )
    ap.add_argument(
        "--output-dir",
        default="/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/Components/Pages/Migrated",
        help="Directory containing existing converted target files to seed the default project",
    )
    ap.add_argument(
        "--icon-dir",
        default="/Users/wadood/projects/VBToCSharp/HomeFront/Syncfusion/31.2.12/Blazor/Samples/Blazor-Server-Demos/wwwroot/images/32",
        help="Directory containing icon files",
    )
    ap.add_argument(
        "--cache-dir",
        default="/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/vb6_migrator/cache",
        help="Legacy cache directory (project-local cache is used for packaged projects)",
    )
    ap.add_argument(
        "--context-dir",
        default="/Users/wadood/projects/VBToCSharp/HomeFront/MobileSource/HomeFront/Components/Pages/Migrated/ContextDir/",
        help="Directory containing .razor/.cs files to include as context",
    )
    ap.add_argument(
        "--instructions",
        default="Convert the attached VB6 files to .NET Blazor using the HF Controls. Employ DBWrapperSqlServer.cs for database queries.",
        help="Additional instructions passed to Codex during conversion",
    )
    ap.add_argument(
        "--vb6-image-dir",
        default="",
        help="VB6 screenshot folder (defaults to <BlazorProject>/wwwroot/images/vb6)",
    )
    ap.add_argument(
        "--dotnet-image-dir",
        default="",
        help=".NET screenshot folder (defaults to <BlazorProject>/wwwroot/images/dotnet)",
    )
    ap.add_argument("--model", default="claude-opus-4-6", help="Model name (claude-opus-4-6 or gpt-5.2-codex)")
    ap.add_argument("--temperature", default="0.2", help="Codex temperature")
    args = ap.parse_args()

    prefer_hfest = args.prefer_hfest.lower() == "true"
    projects_root = Path(args.projects_root).expanduser().resolve()

    if args.project_dir.strip():
        active_project_dir = Path(args.project_dir).expanduser().resolve()
    else:
        if not args.root:
            raise SystemExit("Either --project-dir or --root is required.")
        root_dir = Path(args.root).expanduser().resolve()
        seed_target_dir: Path | None = None
        if (args.output_dir or "").strip():
            target_candidate = Path(args.output_dir).expanduser().resolve()
            if target_candidate.exists():
                seed_target_dir = target_candidate
        active_project_dir = ensure_default_project_package(
            projects_root=projects_root,
            default_name=args.project_name or DEFAULT_PROJECT_NAME,
            vb6_source_root=root_dir,
            output_project=args.output_project,
            seed_target_dir=seed_target_dir,
        )

    while True:
        try:
            runtime = load_project_runtime(active_project_dir, prefer_hfest=prefer_hfest)
        except Exception as exc:
            print(f"Project load error: {exc}", file=sys.stderr)
            sys.exit(1)
        print(runtime["csv_path"])
        print(runtime["mmd_path"])
        print(runtime["text_path"])

        if args.ui.lower() != "true":
            return

        try:
            project_output_csproj = runtime["layout"]["output_project"] or args.output_project
            project_root = Path(project_output_csproj).expanduser().parent
            vb6_image_dir = args.vb6_image_dir.strip() or str(project_root / "wwwroot" / "images" / "vb6")
            dotnet_image_dir = args.dotnet_image_dir.strip() or str(project_root / "wwwroot" / "images" / "dotnet")
            ui_defaults = {
                "root": str(runtime["layout"]["source_dir"]),
                "project": project_output_csproj,
                "output_dir": str(runtime["layout"]["target_dir"]),
                "output_name": "",
                "context_dir": args.context_dir,
                "instructions": args.instructions,
                "model": args.model,
                "temperature": args.temperature,
                "prefer_hfest": prefer_hfest,
                "icon_dir": args.icon_dir,
                "cache_dir": str(runtime["layout"]["cache_dir"]),
                "project_name": runtime["layout"]["name"],
                "project_dir": str(runtime["layout"]["project_dir"]),
                "projects_root": str(projects_root),
                "vb6_image_dir": vb6_image_dir,
                "dotnet_image_dir": dotnet_image_dir,
            }
            action = run_ui(
                runtime["components"],
                runtime["edges"],
                runtime["roots"],
                runtime["duplicates"],
                ui_defaults,
                runtime["migration_order"],
                runtime["cycle_set"],
                runtime["orphan_set"],
            )
        except Exception as exc:
            print(f"UI error: {exc}", file=sys.stderr)
            sys.exit(1)

        if not isinstance(action, dict):
            break
        if action.get("action") != "load_project":
            break
        next_dir = (action.get("project_dir") or "").strip()
        if not next_dir:
            break
        active_project_dir = Path(next_dir).expanduser().resolve()

if __name__ == "__main__":
    if "OPENAI_API_KEY" not in os.environ:
        os.environ["OPENAI_API_KEY"] = "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"
    if "ANTHROPIC_API_KEY" not in os.environ:
        if "ANTHROPIC_API_KEY" not in os.environ:
            os.environ["ANTHROPIC_API_KEY"] = "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"
        pass
    main()


'''
1. In Python Program modify the Compiler Error Panel. display all the errors under compiler errors. Have a single column display. Don't show the path information. FIle Name is sufficent. Display should have filename followed by all errors in that program. Filename shoud be in different colors than error.

'''
