"""
.NET build integration — compiles the project, captures errors,
and parses them into structured records.
"""

import re
import subprocess
import logging
from pathlib import Path
from dataclasses import dataclass, field

from config import DOTNET_CLI, DOTNET_PROJECT_DIR

logger = logging.getLogger(__name__)


@dataclass
class BuildError:
    """One compiler error or warning."""
    file: str
    line: int
    column: int
    code: str        # e.g. CS0246, RZ1001
    severity: str    # "error" or "warning"
    message: str

    def __str__(self):
        return f"{self.file}({self.line},{self.column}): {self.severity} {self.code}: {self.message}"


@dataclass
class BuildResult:
    """Result of a dotnet build invocation."""
    success: bool
    raw_output: str
    errors: list[BuildError] = field(default_factory=list)
    warnings: list[BuildError] = field(default_factory=list)
    error_count: int = 0
    warning_count: int = 0

    @property
    def errors_text(self) -> str:
        """Format all errors as a string suitable for sending to Claude."""
        return "\n".join(str(e) for e in self.errors)

    @property
    def summary(self) -> str:
        return (
            f"Build {'SUCCEEDED' if self.success else 'FAILED'}: "
            f"{self.error_count} error(s), {self.warning_count} warning(s)"
        )


# ── Error Parsing ─────────────────────────────────────────────────────────

# Pattern matches MSBuild-style error lines:
# /path/File.cs(10,5): error CS1002: ; expected
_ERROR_PATTERN = re.compile(
    r"^(?P<file>.+?)\((?P<line>\d+),(?P<col>\d+)\):\s+"
    r"(?P<severity>error|warning)\s+(?P<code>\w+):\s+"
    r"(?P<message>.+)$",
    re.MULTILINE,
)

# Also catch Razor errors:
# /path/File.razor(5,10): error RZ1001: ...
_RAZOR_ERROR_PATTERN = re.compile(
    r"^(?P<file>.+?\.razor)\((?P<line>\d+),(?P<col>\d+)\):\s+"
    r"(?P<severity>error|warning)\s+(?P<code>\w+):\s+"
    r"(?P<message>.+)$",
    re.MULTILINE,
)

# Summary line: "Build FAILED." / "Build succeeded."  +  "X Warning(s)  Y Error(s)"
_SUMMARY_PATTERN = re.compile(
    r"(?P<warnings>\d+)\s+Warning\(s\)\s+(?P<errors>\d+)\s+Error\(s\)",
    re.IGNORECASE,
)


def parse_build_output(raw: str) -> BuildResult:
    """Parse raw dotnet build output into a structured BuildResult."""
    errors: list[BuildError] = []
    warnings: list[BuildError] = []

    for m in _ERROR_PATTERN.finditer(raw):
        entry = BuildError(
            file=m.group("file").strip(),
            line=int(m.group("line")),
            column=int(m.group("col")),
            code=m.group("code"),
            severity=m.group("severity"),
            message=m.group("message").strip(),
        )
        if entry.severity == "error":
            errors.append(entry)
        else:
            warnings.append(entry)

    # De-duplicate (MSBuild sometimes prints errors twice)
    seen = set()
    unique_errors = []
    for e in errors:
        key = (e.file, e.line, e.column, e.code)
        if key not in seen:
            seen.add(key)
            unique_errors.append(e)

    seen.clear()
    unique_warnings = []
    for w in warnings:
        key = (w.file, w.line, w.column, w.code)
        if key not in seen:
            seen.add(key)
            unique_warnings.append(w)

    success = "Build succeeded" in raw and len(unique_errors) == 0

    # Try to parse the summary line
    error_count = len(unique_errors)
    warning_count = len(unique_warnings)
    sm = _SUMMARY_PATTERN.search(raw)
    if sm:
        error_count = max(error_count, int(sm.group("errors")))
        warning_count = max(warning_count, int(sm.group("warnings")))

    return BuildResult(
        success=success,
        raw_output=raw,
        errors=unique_errors,
        warnings=unique_warnings,
        error_count=error_count,
        warning_count=warning_count,
    )


# ── Build Execution ──────────────────────────────────────────────────────

def build_project(project_dir: str | Path | None = None) -> BuildResult:
    """
    Run `dotnet build` on the project and return structured results.

    Args:
        project_dir: Path to the .NET project directory (or .csproj file).
                     Defaults to config.DOTNET_PROJECT_DIR.

    Returns:
        BuildResult with parsed errors and warnings.
    """
    proj = Path(project_dir or DOTNET_PROJECT_DIR)

    # Find the .csproj file
    if proj.is_dir():
        csproj_files = list(proj.glob("*.csproj"))
        if csproj_files:
            proj_arg = str(csproj_files[0])
        else:
            proj_arg = str(proj)
    else:
        proj_arg = str(proj)

    cmd = [DOTNET_CLI, "build", proj_arg, "--no-restore"]
    logger.info("Running: %s", " ".join(cmd))

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            cwd=str(Path(proj_arg).parent if Path(proj_arg).is_file() else proj_arg),
        )
        raw = result.stdout + "\n" + result.stderr
    except subprocess.TimeoutExpired:
        raw = "Build timed out after 120 seconds."
    except FileNotFoundError:
        raw = f"Error: '{DOTNET_CLI}' not found. Is .NET SDK installed?"

    build_result = parse_build_output(raw)
    logger.info(build_result.summary)
    return build_result


def restore_project(project_dir: str | Path | None = None) -> bool:
    """Run dotnet restore before building."""
    proj = Path(project_dir or DOTNET_PROJECT_DIR)
    cmd = [DOTNET_CLI, "restore", str(proj)]
    logger.info("Running: %s", " ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    return result.returncode == 0


def filter_errors_for_file(build_result: BuildResult, filename: str) -> list[BuildError]:
    """Return only the errors that belong to a specific file."""
    return [e for e in build_result.errors if filename in e.file]
