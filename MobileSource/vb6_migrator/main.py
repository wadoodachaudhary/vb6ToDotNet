#!/usr/bin/env python3
"""
VB6 → .NET Migration Tool (powered by Claude API)

Usage:
  # Convert a single file:
  python main.py convert input/MyModule.bas

  # Convert a single file with context:
  python main.py convert input/MyModule.bas --context context/

  # Convert all files in input/:
  python main.py convert-all

  # Just build the project and show errors:
  python main.py build

  # List available files:
  python main.py list
"""

import sys
import argparse
import logging
from pathlib import Path

# Ensure the script's directory is on the path
sys.path.insert(0, str(Path(__file__).parent))

from config import INPUT_DIR, OUTPUT_DIR, CONTEXT_DIR, DOTNET_PROJECT_DIR, LOG_DIR
from migrator import Migrator
from file_manager import list_vb6_files, list_context_files, list_output_files, file_summary
from dotnet_builder import build_project


def setup_logging(verbose: bool = False):
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


# ── Commands ──────────────────────────────────────────────────────────────

def cmd_convert(args):
    """Convert a single VB6 file."""
    vb6_file = Path(args.file)
    if not vb6_file.exists():
        print(f"Error: File not found: {vb6_file}")
        sys.exit(1)

    context_dir = Path(args.context) if args.context else None
    context_files = list_context_files(context_dir) if context_dir else list_context_files()

    migrator = Migrator(
        project_dir=args.project or DOTNET_PROJECT_DIR,
        project_subfolder=args.subfolder or "",
        extra_instructions=args.instructions or "",
    )

    result = migrator.migrate_file(vb6_file, context_files)
    print(f"\n{result}")

    if result.success:
        print(f"\nConverted file: {result.output_file}")
        print(f"Project file:   {result.project_file}")
    else:
        print(f"\nConversion failed: {result.error}")
        if result.build_results:
            last = result.build_results[-1]
            print(f"Last build: {last.summary}")
            for e in last.errors[:10]:
                print(f"  {e}")
        sys.exit(1)


def cmd_convert_all(args):
    """Convert all VB6 files in the input directory."""
    migrator = Migrator(
        project_dir=args.project or DOTNET_PROJECT_DIR,
        project_subfolder=args.subfolder or "",
        extra_instructions=args.instructions or "",
    )
    results = migrator.migrate_all()

    failed = [r for r in results if not r.success]
    if failed:
        sys.exit(1)


def cmd_build(args):
    """Build the .NET project and show errors."""
    project = args.project or DOTNET_PROJECT_DIR
    print(f"Building: {project}")

    result = build_project(project)
    print(f"\n{result.summary}")

    if result.errors:
        print("\nErrors:")
        for e in result.errors:
            print(f"  {e}")

    if result.warnings:
        print(f"\nWarnings: {len(result.warnings)}")
        for w in result.warnings[:5]:
            print(f"  {w}")
        if len(result.warnings) > 5:
            print(f"  ... and {len(result.warnings) - 5} more")

    sys.exit(0 if result.success else 1)


def cmd_list(args):
    """List VB6, context, and output files."""
    print("=" * 60)
    print("VB6 INPUT FILES")
    print("=" * 60)
    for f in list_vb6_files():
        s = file_summary(f)
        print(f"  {s['name']:30s}  {s['line_count']:>5d} lines  {s['size_bytes']:>8d} bytes")

    print()
    print("=" * 60)
    print("CONTEXT FILES (already converted)")
    print("=" * 60)
    for f in list_context_files():
        s = file_summary(f)
        print(f"  {s['name']:30s}  {s['line_count']:>5d} lines  {s['size_bytes']:>8d} bytes")

    print()
    print("=" * 60)
    print("OUTPUT FILES (converted)")
    print("=" * 60)
    for f in list_output_files():
        s = file_summary(f)
        print(f"  {s['name']:30s}  {s['line_count']:>5d} lines  {s['size_bytes']:>8d} bytes")


def cmd_info(args):
    """Show configuration."""
    print("VB6 → .NET Migration Tool")
    print("-" * 40)
    print(f"  Input dir:     {INPUT_DIR}")
    print(f"  Output dir:    {OUTPUT_DIR}")
    print(f"  Context dir:   {CONTEXT_DIR}")
    print(f"  Log dir:       {LOG_DIR}")
    print(f"  .NET project:  {DOTNET_PROJECT_DIR}")
    print(f"  Claude model:  {__import__('config').CLAUDE_MODEL}")


# ── CLI ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="VB6 → .NET Migration Tool (powered by Claude API)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  python main.py convert input/MyModule.bas
  python main.py convert input/MyForm.frm --subfolder Components/Pages
  python main.py convert-all --instructions "Use HomeFront namespace"
  python main.py build
  python main.py list
  python main.py info
""",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging")

    sub = parser.add_subparsers(dest="command", help="Command to run")

    # convert
    p_conv = sub.add_parser("convert", help="Convert a single VB6 file")
    p_conv.add_argument("file", help="Path to the VB6 file")
    p_conv.add_argument("--context", help="Context directory (default: context/)")
    p_conv.add_argument("--project", help="Path to .NET project directory")
    p_conv.add_argument("--subfolder", help="Subfolder in project to place file")
    p_conv.add_argument("--instructions", help="Extra instructions for Claude")

    # convert-all
    p_all = sub.add_parser("convert-all", help="Convert all VB6 files in input/")
    p_all.add_argument("--project", help="Path to .NET project directory")
    p_all.add_argument("--subfolder", help="Subfolder in project to place files")
    p_all.add_argument("--instructions", help="Extra instructions for Claude")

    # build
    p_build = sub.add_parser("build", help="Build .NET project and show errors")
    p_build.add_argument("--project", help="Path to .NET project directory")

    # list
    sub.add_parser("list", help="List VB6, context, and output files")

    # info
    sub.add_parser("info", help="Show configuration")

    args = parser.parse_args()
    setup_logging(args.verbose)

    if args.command is None:
        parser.print_help()
        sys.exit(0)

    commands = {
        "convert": cmd_convert,
        "convert-all": cmd_convert_all,
        "build": cmd_build,
        "list": cmd_list,
        "info": cmd_info,
    }

    commands[args.command](args)


if __name__ == "__main__":
    main()
