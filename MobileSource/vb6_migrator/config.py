"""
Configuration for the VB6-to-.NET migration tool.
"""

import os
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent.resolve()
INPUT_DIR = BASE_DIR / "input"          # VB6 files to convert
OUTPUT_DIR = BASE_DIR / "output"        # Converted .NET files
CONTEXT_DIR = BASE_DIR / "context"      # Already-converted files for context
LOG_DIR = BASE_DIR / "logs"

# ── Claude API ────────────────────────────────────────────────────────────
if "ANTHROPIC_API_KEY" not in os.environ:
      os.environ["ANTHROPIC_API_KEY"] = "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"
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = "claude-sonnet-4-20250514"
#CLAUDE_MODEL = "claude-opus-6"
MAX_TOKENS = 16384
MAX_FIX_ATTEMPTS = 5       # Max rounds of error-fix loop

# ── .NET Build ────────────────────────────────────────────────────────────
DOTNET_PROJECT_DIR = os.environ.get(
    "DOTNET_PROJECT_DIR",
    str(Path(__file__).parent.parent / "HomeFront")
)
DOTNET_CLI = "dotnet"

# ── File Extensions ──────────────────────────────────────────────────────
VB6_EXTENSIONS = {".bas", ".cls", ".frm", ".vbp", ".ctl", ".dsr"}
DOTNET_EXTENSIONS = {".cs", ".razor", ".csproj", ".xaml"}

# ── Ensure directories exist ─────────────────────────────────────────────
for d in [INPUT_DIR, OUTPUT_DIR, CONTEXT_DIR, LOG_DIR]:
    d.mkdir(parents=True, exist_ok=True)
