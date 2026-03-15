"""
Claude API client for VB6 → .NET conversion.
Handles: sending VB6 source + context files, receiving converted C#/Razor,
and iterating on build errors.
"""

import json
import logging
from pathlib import Path
from anthropic import Anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL, MAX_TOKENS

logger = logging.getLogger(__name__)


SYSTEM_PROMPT = """\
You are an expert VB6-to-.NET migration engineer. You convert Visual Basic 6 \
code into modern C# / .NET 10 Blazor code.

RULES:
1. Produce COMPLETE, compilable C# (.cs) or Razor (.razor) files.
2. Use modern C# 14 idioms: records, pattern matching, nullable reference types, \
   file-scoped namespaces, primary constructors where appropriate.
3. Preserve the original business logic exactly.
4. Map VB6 types to .NET equivalents (Integer→int, Long→long, String→string, \
   Variant→object, Collection→List<T>, etc.).
5. Convert VB6 Forms/Controls to Blazor components when the target is a UI file.
6. Convert VB6 Modules (.bas) to static classes, Class Modules (.cls) to classes.
7. Wrap your output in a single fenced code block with the target filename on the \
   first line as a comment:  // File: ClassName.cs
8. If the file references types from the context files provided, use the correct \
   namespaces and type names — do NOT re-declare them.
9. Do NOT add explanatory text outside the code block.
"""

FIX_ERRORS_PROMPT = """\
The converted file failed to compile. Fix ALL of the build errors listed below.

Return the COMPLETE corrected file in a single fenced code block with \
// File: <filename> on the first line. Do NOT add explanatory text outside the \
code block.

BUILD ERRORS:
{errors}
"""


class ClaudeClient:
    """Thin wrapper around the Anthropic SDK for migration tasks."""

    def __init__(self, api_key: str | None = None):
        key = api_key or ANTHROPIC_API_KEY
        if not key:
            raise ValueError(
                "ANTHROPIC_API_KEY is not set. "
                "Export it as an environment variable or pass it explicitly."
            )
        self.client = Anthropic(api_key=key)
        self.conversation: list[dict] = []   # running conversation history

    # ── Helpers ───────────────────────────────────────────────────────────

    @staticmethod
    def _read_file(path: Path) -> str:
        return path.read_text(encoding="utf-8", errors="replace")

    @staticmethod
    def _build_file_block(path: Path, content: str | None = None) -> str:
        """Format a file as a labelled block for the prompt."""
        text = content or ClaudeClient._read_file(path)
        return f"### {path.name}\n```\n{text}\n```"

    def _call(self, messages: list[dict], system: str = SYSTEM_PROMPT) -> str:
        """Send messages to Claude and return the assistant text."""
        logger.info("Calling Claude (%s) — %d message(s)", CLAUDE_MODEL, len(messages))
        response = self.client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=MAX_TOKENS,
            system=system,
            messages=messages,
        )
        text = response.content[0].text
        logger.debug("Response length: %d chars", len(text))
        return text

    # ── Public API ────────────────────────────────────────────────────────

    def convert_file(
        self,
        vb6_file: Path,
        context_files: list[Path] | None = None,
        extra_instructions: str = "",
    ) -> str:
        """
        Send a VB6 file (+ optional already-converted context files) to Claude
        and return the converted .NET source code.

        Args:
            vb6_file:           Path to the VB6 source file to convert.
            context_files:      Paths to already-converted .NET files for context.
            extra_instructions: Additional conversion instructions from the user.

        Returns:
            The raw assistant response containing the converted code.
        """
        # Build the user message
        parts: list[str] = []

        # 1 — Context files (already converted)
        if context_files:
            parts.append("## Already-converted project files (for reference / context):\n")
            for cf in context_files:
                parts.append(self._build_file_block(cf))
            parts.append("")

        # 2 — The VB6 file to convert
        parts.append("## VB6 file to convert:\n")
        parts.append(self._build_file_block(vb6_file))
        parts.append("")

        # 3 — Extra instructions
        if extra_instructions:
            parts.append(f"## Additional instructions:\n{extra_instructions}\n")

        parts.append(
            "Convert the VB6 file above into a modern C# / .NET 10 Blazor file. "
            "Return ONLY the complete code in a fenced code block."
        )

        user_content = "\n".join(parts)

        # Start a fresh conversation
        self.conversation = [{"role": "user", "content": user_content}]
        assistant_text = self._call(self.conversation)
        self.conversation.append({"role": "assistant", "content": assistant_text})

        return assistant_text

    def fix_errors(self, errors: str, current_code: str) -> str:
        """
        Send build errors back to Claude and ask for a corrected version.
        Continues the existing conversation so Claude has full context.

        Args:
            errors:       The raw compiler error output.
            current_code: The current (broken) version of the converted file.

        Returns:
            The raw assistant response with corrected code.
        """
        user_content = (
            FIX_ERRORS_PROMPT.format(errors=errors)
            + "\n\n## Current code:\n```csharp\n"
            + current_code
            + "\n```"
        )

        self.conversation.append({"role": "user", "content": user_content})
        assistant_text = self._call(self.conversation)
        self.conversation.append({"role": "assistant", "content": assistant_text})

        return assistant_text

    def reset_conversation(self):
        """Clear conversation history for a new file."""
        self.conversation.clear()
