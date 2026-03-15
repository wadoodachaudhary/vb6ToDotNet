"""
Orchestrator — ties together Claude conversion, file I/O, building,
and the error-fix loop into a single pipeline.
"""

import logging
import datetime
from pathlib import Path

from config import MAX_FIX_ATTEMPTS, DOTNET_PROJECT_DIR, LOG_DIR
from claude_client import ClaudeClient
from file_manager import (
    list_vb6_files,
    list_context_files,
    extract_code_from_response,
    write_output_file,
    copy_to_project,
    read_file,
)
from dotnet_builder import (
    build_project,
    filter_errors_for_file,
    restore_project,
    BuildResult,
)

logger = logging.getLogger(__name__)


class MigrationResult:
    """Result of migrating a single VB6 file."""

    def __init__(self, vb6_file: Path):
        self.vb6_file = vb6_file
        self.output_file: Path | None = None
        self.project_file: Path | None = None
        self.success: bool = False
        self.attempts: int = 0
        self.build_results: list[BuildResult] = []
        self.error: str | None = None

    def __str__(self):
        status = "SUCCESS" if self.success else "FAILED"
        return (
            f"[{status}] {self.vb6_file.name} → "
            f"{self.output_file.name if self.output_file else '???'} "
            f"({self.attempts} attempt(s))"
        )


class Migrator:
    """
    Full migration pipeline:
      1. Read VB6 file
      2. Collect context files
      3. Send to Claude for conversion
      4. Write converted file to output and project
      5. Build the project
      6. If errors, send them back to Claude (up to MAX_FIX_ATTEMPTS)
      7. Report results
    """

    def __init__(
        self,
        api_key: str | None = None,
        project_dir: str | Path | None = None,
        project_subfolder: str = "",
        extra_instructions: str = "",
    ):
        self.claude = ClaudeClient(api_key=api_key)
        self.project_dir = Path(project_dir or DOTNET_PROJECT_DIR)
        self.project_subfolder = project_subfolder
        self.extra_instructions = extra_instructions
        self.results: list[MigrationResult] = []

    def migrate_file(
        self,
        vb6_file: Path,
        context_files: list[Path] | None = None,
    ) -> MigrationResult:
        """
        Migrate a single VB6 file through the full pipeline.

        Steps:
          1. Send VB6 + context to Claude → get converted code
          2. Extract code, write to output/ and project
          3. Build the project
          4. If errors, send them back to Claude, write fix, rebuild
          5. Repeat up to MAX_FIX_ATTEMPTS times
        """
        result = MigrationResult(vb6_file)
        self.claude.reset_conversation()

        logger.info("=" * 70)
        logger.info("MIGRATING: %s", vb6_file.name)
        logger.info("=" * 70)

        # ── Step 1: Initial conversion ────────────────────────────────────
        try:
            logger.info("Step 1: Sending to Claude for conversion...")
            response = self.claude.convert_file(
                vb6_file=vb6_file,
                context_files=context_files,
                extra_instructions=self.extra_instructions,
            )
            result.attempts = 1
        except Exception as e:
            result.error = f"Claude API error: {e}"
            logger.error(result.error)
            self.results.append(result)
            return result

        # ── Step 2: Extract and write ─────────────────────────────────────
        filename, code = extract_code_from_response(response)
        result.output_file = write_output_file(filename, code)

        # Copy into the .NET project
        result.project_file = copy_to_project(
            result.output_file, self.project_dir, self.project_subfolder
        )

        # ── Step 3–6: Build → Fix loop ────────────────────────────────────
        for attempt in range(1, MAX_FIX_ATTEMPTS + 1):
            logger.info("Build attempt %d / %d ...", attempt, MAX_FIX_ATTEMPTS)
            build = build_project(self.project_dir)
            result.build_results.append(build)

            # Filter to only errors from OUR file
            file_errors = filter_errors_for_file(build, filename)

            if build.success or not file_errors:
                logger.info("Build PASSED (or no errors from our file)!")
                result.success = True
                break

            logger.warning(
                "Build FAILED with %d error(s) from %s",
                len(file_errors), filename,
            )
            for err in file_errors:
                logger.warning("  %s", err)

            if attempt >= MAX_FIX_ATTEMPTS:
                logger.error("Max fix attempts reached. Giving up.")
                result.error = f"Still {len(file_errors)} error(s) after {attempt} fix attempt(s)"
                break

            # Send errors back to Claude
            logger.info("Sending %d error(s) to Claude for fixing...", len(file_errors))
            errors_text = "\n".join(str(e) for e in file_errors)
            try:
                fix_response = self.claude.fix_errors(errors_text, code)
                result.attempts += 1
            except Exception as e:
                result.error = f"Claude API error on fix attempt: {e}"
                logger.error(result.error)
                break

            # Extract fixed code and overwrite
            _, code = extract_code_from_response(fix_response)
            result.output_file = write_output_file(filename, code)
            result.project_file = copy_to_project(
                result.output_file, self.project_dir, self.project_subfolder
            )

        self.results.append(result)
        self._save_log(result)
        return result

    def migrate_all(
        self,
        vb6_dir: Path | None = None,
        context_dir: Path | None = None,
    ) -> list[MigrationResult]:
        """
        Migrate all VB6 files in the input directory.

        Args:
            vb6_dir:     Directory containing VB6 files (default: config.INPUT_DIR)
            context_dir: Directory containing already-converted files (default: config.CONTEXT_DIR)

        Returns:
            List of MigrationResult for each file.
        """
        vb6_files = list_vb6_files(vb6_dir)
        context_files = list_context_files(context_dir)

        if not vb6_files:
            logger.warning("No VB6 files found in input directory.")
            return []

        logger.info("Starting migration of %d file(s)...", len(vb6_files))
        logger.info("Context files: %d", len(context_files))

        # Restore NuGet packages once before building
        restore_project(self.project_dir)

        for vb6_file in vb6_files:
            self.migrate_file(vb6_file, context_files)

            # After each successful conversion, add the output to context
            # so subsequent files can reference it
            if self.results[-1].success and self.results[-1].output_file:
                context_files.append(self.results[-1].output_file)

        self._print_summary()
        return self.results

    # ── Reporting ─────────────────────────────────────────────────────────

    def _save_log(self, result: MigrationResult):
        """Save a migration log for a single file."""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = LOG_DIR / f"{result.vb6_file.stem}_{timestamp}.log"

        lines = [
            f"Migration Log: {result.vb6_file.name}",
            f"Timestamp: {timestamp}",
            f"Status: {'SUCCESS' if result.success else 'FAILED'}",
            f"Attempts: {result.attempts}",
            f"Output: {result.output_file}",
            f"Project: {result.project_file}",
            "",
        ]

        if result.error:
            lines.append(f"Error: {result.error}")
            lines.append("")

        for i, br in enumerate(result.build_results, 1):
            lines.append(f"--- Build Attempt {i} ---")
            lines.append(br.summary)
            if br.errors:
                for e in br.errors:
                    lines.append(f"  {e}")
            lines.append("")

        log_file.write_text("\n".join(lines), encoding="utf-8")
        logger.info("Log saved: %s", log_file)

    def _print_summary(self):
        """Print a summary table of all migration results."""
        print("\n" + "=" * 70)
        print("MIGRATION SUMMARY")
        print("=" * 70)

        total = len(self.results)
        passed = sum(1 for r in self.results if r.success)
        failed = total - passed

        for r in self.results:
            print(f"  {r}")

        print("-" * 70)
        print(f"  Total: {total}  |  Passed: {passed}  |  Failed: {failed}")
        print("=" * 70)
