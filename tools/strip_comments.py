#!/usr/bin/env python3
"""
Strip comments from .razor / .razor.cs / .razor.css / .cs / .css files.
Preserves the FIRST block comment in each file (header/banner).

Used by deploy_to_repos.sh so repositories get clean code without LLM
marginalia, VB6 notes, inline TODOs, etc.

Edits files IN PLACE — run against a staging copy, never the working tree.
"""
import re
import sys
from pathlib import Path

RAZOR_COMMENT_RE  = re.compile(r'@\*.*?\*@', re.DOTALL)
HTML_COMMENT_RE   = re.compile(r'<!--.*?-->', re.DOTALL)
C_BLOCK_RE_DOTALL = re.compile(r'/\*.*?\*/', re.DOTALL)
C_BLOCK_RE_LINE   = re.compile(r'/\*.*?\*/')
C_LINE_RE         = re.compile(r'^[ \t]*//[^\r\n]*\r?\n?', re.MULTILINE)


def _find_first_block(text, pattern):
    m = pattern.search(text)
    return m.group(0) if m else None


def _strip_all_but_first(text, pattern):
    first_found = [False]
    def _replacer(m):
        if not first_found[0]:
            first_found[0] = True
            return m.group(0)
        return ""
    return pattern.sub(_replacer, text)


def strip_file(path: Path) -> bool:
    text = path.read_text(encoding="utf-8")
    original = text
    name = path.name
    is_razor = name.endswith(".razor") and not name.endswith(".razor.cs") and not name.endswith(".razor.css")
    is_cs = name.endswith(".cs")
    is_css = name.endswith(".css")

    if is_razor:
        text = _strip_all_but_first(text, RAZOR_COMMENT_RE)
        text = HTML_COMMENT_RE.sub("", text)
        text = C_BLOCK_RE_LINE.sub("", text)
    elif is_cs:
        text = _strip_all_but_first(text, C_BLOCK_RE_DOTALL)
    elif is_css:
        text = _strip_all_but_first(text, C_BLOCK_RE_DOTALL)

    text = C_LINE_RE.sub("", text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = text.rstrip('\n') + '\n'

    if text != original:
        path.write_text(text, encoding="utf-8")
        return True
    return False


def main():
    if len(sys.argv) < 2:
        print(f"usage: {sys.argv[0]} <directory>", file=sys.stderr)
        sys.exit(2)
    root = Path(sys.argv[1])
    if not root.is_dir():
        print(f"not a directory: {root}", file=sys.stderr)
        sys.exit(2)

    changed = 0
    total = 0
    for ext in ("*.razor", "*.razor.cs", "*.razor.css", "*.cs", "*.css"):
        for p in root.rglob(ext):
            if "/bin/" in str(p) or "/obj/" in str(p):
                continue
            total += 1
            if strip_file(p):
                changed += 1
    print(f"Processed {total} files; stripped comments in {changed}.")


if __name__ == "__main__":
    main()
