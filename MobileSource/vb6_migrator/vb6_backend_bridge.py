#!/usr/bin/env python3
"""Minimal JSON bridge for Blazor migrator fallback paths.

This script intentionally keeps the contract small:
- analyze: load packaged project runtime via vb6_callgraph_ui.load_project_runtime
           and emit JSON consumable by HomeFront.Services.Migrator.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Tuple

from vb6_callgraph_ui import load_project_runtime

SKIP_DIRS = {"bin", "obj", ".git", ".idea", ".vs", "node_modules", "packages"}


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="latin-1", errors="ignore")
    except Exception:
        try:
            return path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return ""


def count_lines(path: Path) -> int:
    txt = read_text(path)
    if not txt:
        return 0
    return len(txt.replace("\r\n", "\n").split("\n"))


def iter_razor_files(root: Path) -> List[Path]:
    files: List[Path] = []
    if not root.exists():
        return files
    stack = [root]
    while stack:
        current = stack.pop()
        try:
            for child in current.iterdir():
                if child.is_dir():
                    if child.name.lower() in SKIP_DIRS:
                        continue
                    stack.append(child)
                    continue
                if child.is_file() and child.suffix.lower() == ".razor":
                    files.append(child.resolve())
        except Exception:
            continue
    files.sort(key=lambda p: str(p).lower())
    return files


def build_razor_index(layout: dict) -> Dict[str, List[Path]]:
    roots: List[Path] = []
    target = Path(str(layout.get("target_dir") or "")).expanduser()
    if target.exists():
        roots.append(target.resolve())

    output_project = str(layout.get("output_project") or "").strip()
    if output_project:
        out = Path(output_project).expanduser()
        if out.exists():
            roots.append(out.resolve().parent)

    index: Dict[str, List[Path]] = {}
    for root in roots:
        for path in iter_razor_files(root):
            key = path.stem.lower()
            index.setdefault(key, []).append(path)

    for key in list(index.keys()):
        index[key] = sorted(index[key], key=lambda p: (razor_score(p), str(p).lower()))
    return index


def razor_score(path: Path) -> int:
    normalized = str(path).replace("\\", "/").lower()
    if "/components/pages/migrated/" in normalized:
        return 0
    if "/components/pages/" in normalized:
        return 1
    if "/target/" in normalized:
        return 2
    return 3


def resolve_migrated_path(component_name: str, source_stem: str, index: Dict[str, List[Path]]) -> str:
    keys: List[str] = []

    def add(value: str) -> None:
        if not value:
            return
        lowered = value.lower()
        if lowered not in keys:
            keys.append(lowered)

    add(component_name)
    add(source_stem)
    if component_name.lower().startswith("frm") and len(component_name) > 3:
        add(component_name[3:])

    for key in keys:
        hits = index.get(key)
        if not hits:
            continue
        for path in hits:
            if path.exists():
                return str(path)
    return ""


def build_tree_rows(components: list, migration_order: list[str], cycle_set: set[str], orphan_set: set[str], roots: list[str]):
    order_lookup = {cid: idx for idx, cid in enumerate(migration_order)}
    by_id = {f"{c.project}::{c.name}": c for c in components}

    rows = []
    node_id = 1
    ordered_ids = sorted(by_id.keys(), key=lambda cid: (order_lookup.get(cid, 10**9), cid.lower()))
    for cid in ordered_ids:
        comp = by_id[cid]
        notes = []
        if cid in roots:
            notes.append("Root")
        if cid in cycle_set:
            notes.append("Cycle")
        if cid in orphan_set:
            notes.append("Orphan")
        rows.append(
            {
                "nodeId": node_id,
                "parentId": None,
                "componentId": cid,
                "displayName": comp.name,
                "project": comp.project,
                "type": comp.type,
                "status": "Pending",
                "notes": ", ".join(notes),
            }
        )
        node_id += 1
    return rows


def analyze(project_dir: Path, prefer_hfest: bool) -> int:
    runtime = load_project_runtime(project_dir.expanduser().resolve(), prefer_hfest=prefer_hfest)
    layout = runtime["layout"]
    components = runtime["components"]
    edges = runtime["edges"]
    migration_order = list(runtime.get("migration_order") or [])
    cycle_set = set(runtime.get("cycle_set") or set())
    orphan_set = set(runtime.get("orphan_set") or set())
    roots = list(runtime.get("roots") or [])

    incoming: Dict[str, int] = {}
    outgoing: Dict[str, int] = {}
    edge_rows = []
    for (caller, callee), meta in edges.items():
        outgoing[caller] = outgoing.get(caller, 0) + 1
        incoming[callee] = incoming.get(callee, 0) + 1
        edge_rows.append(
            {
                "caller": caller,
                "callee": callee,
                "ambiguous": bool(meta.get("ambiguous")),
                "types": sorted(list(meta.get("types") or [])),
            }
        )

    order_lookup = {cid: idx + 1 for idx, cid in enumerate(migration_order)}
    razor_index = build_razor_index(layout)

    component_rows = []
    converted_count = 0
    total_lines = 0
    for comp in components:
        comp_id = f"{comp.project}::{comp.name}"
        source_path = comp.path.resolve()
        source_lines = count_lines(source_path)
        total_lines += source_lines

        migrated_path = resolve_migrated_path(comp.name, source_path.stem, razor_index)
        is_converted = bool(migrated_path)
        converted_at_utc = None
        if is_converted:
            converted_count += 1
            try:
                converted_at_utc = Path(migrated_path).stat().st_mtime
            except Exception:
                converted_at_utc = None

        component_rows.append(
            {
                "componentId": comp_id,
                "project": comp.project,
                "name": comp.name,
                "type": comp.type,
                "sourcePath": str(source_path),
                "relativeSourcePath": str(source_path.relative_to(Path(layout["source_dir"]))),
                "migratedPath": migrated_path,
                "isConverted": is_converted,
                "convertedAtUtc": None,
                "migrationOrder": order_lookup.get(comp_id, 2**31 - 1),
                "sourceLines": source_lines,
                "callsOut": outgoing.get(comp_id, 0),
                "calledBy": incoming.get(comp_id, 0),
                "isCycleMember": comp_id in cycle_set,
                "isOrphanForm": comp_id in orphan_set,
            }
        )

    component_rows.sort(
        key=lambda row: (
            0 if row["isConverted"] else 1,
            row["migrationOrder"],
            row["name"].lower(),
        )
    )

    tree_rows = build_tree_rows(components, migration_order, cycle_set, orphan_set, roots)

    payload = {
        "layout": {
            "name": str(layout.get("name") or Path(layout["project_dir"]).name),
            "projectDirectory": str(layout.get("project_dir") or project_dir),
            "sourceDirectory": str(layout.get("source_dir") or ""),
            "targetDirectory": str(layout.get("target_dir") or ""),
            "cacheDirectory": str(layout.get("cache_dir") or ""),
            "graphsDirectory": str(layout.get("graphs_dir") or ""),
            "outputProject": str(layout.get("output_project") or ""),
            "sourceOrigin": str(layout.get("source_origin") or ""),
            "targetOrigin": str(layout.get("target_origin") or ""),
        },
        "components": component_rows,
        "edges": edge_rows,
        "treeRows": tree_rows,
        "stats": {
            "totalFiles": len(component_rows),
            "convertedFiles": converted_count,
            "pendingFiles": max(0, len(component_rows) - converted_count),
            "cycleFiles": len(cycle_set),
            "orphanForms": len(orphan_set),
            "totalSourceLines": total_lines,
        },
        "migrationOrder": migration_order,
        "cycleSet": sorted(list(cycle_set)),
        "orphanSet": sorted(list(orphan_set)),
    }

    print(json.dumps(payload, ensure_ascii=False))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="cmd", required=True)

    analyze_cmd = sub.add_parser("analyze")
    analyze_cmd.add_argument("--project-dir", required=True)
    analyze_cmd.add_argument("--prefer-hfest", default="true", choices=["true", "false"])

    args = parser.parse_args()

    if args.cmd == "analyze":
        return analyze(Path(args.project_dir), prefer_hfest=(args.prefer_hfest.lower() == "true"))

    return 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # pragma: no cover - fallback safety
        print(f"{{\"error\": {json.dumps(str(exc))}}}", file=sys.stderr)
        raise
