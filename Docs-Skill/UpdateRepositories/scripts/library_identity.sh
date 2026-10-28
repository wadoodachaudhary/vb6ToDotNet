#!/bin/bash
# library_identity.sh — read-only. Shows every difference between the FlexKit and FlexCore source trees that is
# not an expected identity difference (owner 2026-10-27: the two libraries carry the same source; only project,
# package and assembly identity differ). Untracked files count: a new file is untracked until committed.
#
#   bash library_identity.sh        # prints the differences, then IDENTICAL or N DIFFERENCE(S); exit 0 / 1
#
# Expected to differ, and not reported: FlexKit.csproj / FlexCore.csproj, the documents project file
# (FlexCore.Documents.csproj; FlexKit compiles FlexKit.Documents/ into FlexKit.dll), README.md, CLAUDE.md,
# and FlexCore's own tests/, FlexCore.Llm/ and FlexCore.Migration/. FlexKit.Documents/ is compared with
# FlexCore.Documents/ file by file. The documents' web assets stay in FlexKit's own wwwroot (the Razor SDK
# serves a project's wwwroot only), so FlexKit wwwroot/pdf-viewer.js and wwwroot/vendor/** are compared with
# FlexCore.Documents/wwwroot/. A listed file missing on disk (deleted, not yet committed) is reported.
set -u
R=/Users/wadood/projects/VBToCSharp
list() {  # $1 = FlexKit | FlexCore -> "normalised-path<TAB>real-path", one per tracked or untracked file
  git -C "$R/$1" ls-files --cached --others --exclude-standard | while IFS= read -r f; do
    n=$(printf '%s' "$f" | sed -E 's#^(FlexKit|FlexCore)\.Documents/#BRAND.Documents/#')
    [ "$1" = FlexKit ] && n=$(printf '%s' "$n" | sed -E 's#^wwwroot/(pdf-viewer\.js|vendor/)#BRAND.Documents/wwwroot/\1#')
    printf '%s\t%s\n' "$n" "$f"
  done | grep -vE $'^(tests/|FlexCore\\.Llm/|FlexCore\\.Migration/|BRAND\\.Documents/[^/]+\\.csproj\t|(FlexKit|FlexCore)\\.csproj\t|README\\.md\t|CLAUDE\\.md\t)' | sort
}
T=$(mktemp -d); trap 'rm -rf "$T"' EXIT
list FlexKit > "$T/kit"; list FlexCore > "$T/core"
n=0
while IFS= read -r p; do echo "ONLY IN FlexKit:  $p"; n=$((n+1)); done < <(comm -23 <(cut -f1 "$T/kit") <(cut -f1 "$T/core"))
while IFS= read -r p; do echo "ONLY IN FlexCore: $p"; n=$((n+1)); done < <(comm -13 <(cut -f1 "$T/kit") <(cut -f1 "$T/core"))
join -t $'\t' "$T/kit" "$T/core" > "$T/both"
while IFS=$'\t' read -r p k c; do
  if [ ! -f "$R/FlexKit/$k" ] || [ ! -f "$R/FlexCore/$c" ]; then echo "MISSING ON DISK:  $p"; n=$((n+1)); continue; fi
  cmp -s "$R/FlexKit/$k" "$R/FlexCore/$c" || { echo "DIFFER:           $p"; n=$((n+1)); }
done < "$T/both"
# A literal _content/FlexKit or _content/FlexCore path in library code fails at run time in the other library.
lit=$(grep -rlE '_content/Flex(Kit|Core)/' "$R/FlexKit" "$R/FlexCore" --include='*.cs' --include='*.razor' --exclude-dir=bin --exclude-dir=obj 2>/dev/null | head -5)
[ -n "$lit" ] && { echo "LITERAL _content PATH: $lit"; n=$((n+1)); }
echo "shared paths: $(wc -l < "$T/both" | tr -d ' ')"
if [ "$n" -eq 0 ]; then echo "IDENTICAL"; exit 0; else echo "$n DIFFERENCE(S)"; exit 1; fi
