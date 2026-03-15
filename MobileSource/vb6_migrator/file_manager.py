"""
File manager — handles reading VB6 input, collecting context files,
extracting code from Claude responses, and writing converted .NET files.
"""

import re
import logging
from pathlib import Path

from config import (
    INPUT_DIR, OUTPUT_DIR, CONTEXT_DIR,
    VB6_EXTENSIONS, DOTNET_EXTENSIONS,
)

logger = logging.getLogger(__name__)


# ── Listing / Discovery ──────────────────────────────────────────────────

def list_vb6_files(directory: Path | None = None) -> list[Path]:
    """Return all VB6 source files in the input directory."""
    d = directory or INPUT_DIR
    files = [f for f in d.rglob("*") if f.suffix.lower() in VB6_EXTENSIONS]
    files.sort(key=lambda f: f.name.lower())
    logger.info("Found %d VB6 file(s) in %s", len(files), d)
    return files


def list_context_files(directory: Path | None = None) -> list[Path]:
    """Return all already-converted .NET files to use as context."""
    d = directory or CONTEXT_DIR
    files = [f for f in d.rglob("*") if f.suffix.lower() in DOTNET_EXTENSIONS]
    files.sort(key=lambda f: f.name.lower())
    logger.info("Found %d context file(s) in %s", len(files), d)
    return files


def list_output_files(directory: Path | None = None) -> list[Path]:
    """Return all converted files in the output directory."""
    d = directory or OUTPUT_DIR
    files = [f for f in d.rglob("*") if f.suffix.lower() in DOTNET_EXTENSIONS]
    files.sort(key=lambda f: f.name.lower())
    return files


# ── Code Extraction ──────────────────────────────────────────────────────

def extract_code_from_response(response: str) -> tuple[str, str]:
    """
    Extract the code and target filename from Claude's fenced-block response.

    Returns:
        (filename, code) — filename is parsed from the first '// File:' comment
        inside the code block; code is the full content.
    """
    # Match fenced code blocks (```csharp ... ``` or ```razor ... ``` or ``` ... ```)
    pattern = r"```(?:csharp|cs|razor|xml|html)?\s*\n(.*?)```"
    match = re.search(pattern, response, re.DOTALL)

    if not match:
        # Fallback: try to find any fenced block
        pattern = r"```\s*\n(.*?)```"
        match = re.search(pattern, response, re.DOTALL)

    if not match:
        logger.warning("No fenced code block found in response; using raw text")
        code = response.strip()
    else:
        code = match.group(1).strip()

    # Extract filename from // File: comment
    filename = "ConvertedFile.cs"
    file_pattern = r"//\s*File:\s*(.+?)(?:\s*$|\n)"
    file_match = re.search(file_pattern, code)
    if file_match:
        filename = file_match.group(1).strip()
        # Remove the // File: line from the code
        code = code[file_match.end():].strip()

    # Determine extension from content if filename doesn't have one
    if not Path(filename).suffix:
        if "@page" in code or "@code" in code:
            filename += ".razor"
        else:
            filename += ".cs"

    logger.info("Extracted filename: %s (%d chars of code)", filename, len(code))
    return filename, code


# ── File I/O ─────────────────────────────────────────────────────────────

def read_file(path: Path) -> str:
    """Read a file and return its content."""
    return path.read_text(encoding="utf-8", errors="replace")


def write_output_file(filename: str, code: str, directory: Path | None = None) -> Path:
    """
    Write converted code to the output directory.

    Returns:
        The full path to the written file.
    """
    d = directory or OUTPUT_DIR
    out_path = d / filename
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(code, encoding="utf-8")
    logger.info("Wrote %s (%d bytes)", out_path, len(code))
    return out_path


def copy_to_project(output_file: Path, project_dir: Path, subfolder: str = "") -> Path:
    """
    Copy a converted file into the .NET project tree.

    Args:
        output_file: The file in the output directory.
        project_dir: Root of the .NET project.
        subfolder:   Optional subfolder within the project (e.g., 'Components/Pages').

    Returns:
        The destination path.
    """
    dest_dir = project_dir / subfolder if subfolder else project_dir
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / output_file.name

    dest.write_text(output_file.read_text(encoding="utf-8"), encoding="utf-8")
    logger.info("Copied %s → %s", output_file.name, dest)
    return dest


# ── Summary ──────────────────────────────────────────────────────────────

def file_summary(path: Path) -> dict:
    """Return metadata about a file."""
    content = read_file(path)
    return {
        "name": path.name,
        "path": str(path),
        "extension": path.suffix,
        "size_bytes": path.stat().st_size,
        "line_count": content.count("\n") + 1,
    }
