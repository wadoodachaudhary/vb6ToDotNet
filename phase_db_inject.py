#!/usr/bin/env python3
"""
Replace `private DbWrapperSqlServer db = new()` field declarations
with `@inject DbWrapperSqlServer db` directive in .razor files.

Also handles variants:
- private DbWrapperSqlServer db = new DbWrapperSqlServer();
- private readonly DbWrapperSqlServer db = new();
- private HomeFront.Data.DbWrapperSqlServer db = new();
"""

import re
import os
import sys

def process_file(filepath, dry_run=False):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    original = content

    # Match field declarations for DbWrapperSqlServer
    field_patterns = [
        r'^\s*private\s+(?:readonly\s+)?(?:HomeFront\.Data\.)?DbWrapperSqlServer\s+db\s*=\s*new\s*(?:DbWrapperSqlServer)?\s*\(\s*\)\s*;\s*\n?',
    ]

    has_field = any(re.search(p, content, re.MULTILINE) for p in field_patterns)
    if not has_field:
        return None

    # Check if already has @inject DbWrapperSqlServer
    has_inject = re.search(r'@inject\s+(?:HomeFront\.Data\.)?DbWrapperSqlServer\s+', content)
    if has_inject:
        # Already injected — just remove the field declaration
        for p in field_patterns:
            content = re.sub(p, '', content, flags=re.MULTILINE)
        if content != original and not dry_run:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
        return 'Removed duplicate db field (inject already exists)'

    # Remove the field declaration
    for p in field_patterns:
        content = re.sub(p, '', content, flags=re.MULTILINE)

    # Add @inject directive after last existing @inject or @using line
    lines = content.split('\n')
    insert_idx = 0
    has_data_using = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith('@inject ') or stripped.startswith('@using '):
            insert_idx = i + 1
        if 'HomeFront.Data' in stripped:
            has_data_using = True

    # Add @using HomeFront.Data if needed
    if not has_data_using:
        lines.insert(insert_idx, '@using HomeFront.Data')
        insert_idx += 1

    lines.insert(insert_idx, '@inject DbWrapperSqlServer db')
    content = '\n'.join(lines)

    if content != original and not dry_run:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)

    return 'Replaced db field with @inject DbWrapperSqlServer db'


def main():
    dry_run = '--dry-run' in sys.argv
    mode = "DRY RUN" if dry_run else "APPLYING"

    dirs = [
        ('HomeFrontPOC', 'HomeFrontPOC/Components/Pages'),
        ('Main', 'MobileSource/HomeFront/Components/Pages'),
    ]

    print(f"\n=== DB Injection Migration ({mode}) ===\n")

    for name, directory in dirs:
        print(f"--- {name} ---")
        count = 0
        for root, _, files in os.walk(directory):
            if 'vb6_migrator' in root:
                continue
            for fname in sorted(files):
                if not fname.endswith('.razor'):
                    continue
                fp = os.path.join(root, fname)
                result = process_file(fp, dry_run)
                if result:
                    rel = os.path.relpath(fp, directory)
                    print(f"  {rel}: {result}")
                    count += 1
        print(f"  Total: {count} files\n")


if __name__ == '__main__':
    main()
