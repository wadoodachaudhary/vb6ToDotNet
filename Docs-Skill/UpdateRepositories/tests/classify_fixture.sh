#!/bin/bash
# Fixture test for classify_file/hunks_present in scripts/_common.sh. Builds throwaway repos in a temp
# dir; touches nothing real. Run after changing _common.sh:  bash tests/classify_fixture.sh
set -uo pipefail
T=$(mktemp -d); export UR_D=$T/repos UR_HF=$T/src UR_H=$T/H; mkdir -p $T/H/Deploy
git init -q --bare $T/remote.git
git clone -q $T/remote.git $T/w 2>/dev/null; cd $T/w; git config user.email a@a; git config user.name a
mkdir -p P
printf 'line one of a long enough file\n<div class="teammate-edits-here">original body text</div>\n}\n' > P/Ours.razor
printf 'keep this old named page body\n' > P/Old.razor
printf 'a file the teammate will delete entirely\n' > P/Del.razor
printf '<Project>\n  <ItemGroup>\n  </ItemGroup>\n</Project>\n' > HomeFront.csproj
printf 'alpha line that stays\nobsolete configuration line to remove\n' > P/Rm.razor
printf 'historic original content here\n' > P/Hist.razor
printf 'generic file body line here\n' > P/Gen.razor
printf 'superseded original body line\n' > P/Sup.razor
printf 'acknowledged original body line\n' > P/Ack.razor
git add -A; git commit -qm "Deploy hyphen-pb — base"
printf 'historic original content here\nhistoric teammate addition line\n' > P/Hist.razor; git commit -qam "teammate: historic fix"
printf 'superseded original body line\nsuperseded teammate addition line\n' > P/Sup.razor; git commit -qam "teammate: historic sup"
printf 'historic original content here\nour own unrelated edit in hist\n' > P/Hist.razor
printf 'superseded original body line\nOUR later rewrite of that addition\n' > P/Sup.razor
git commit -qam "Deploy hyphen-pb — later (overwrites hist and sup)"
printf 'line one of a long enough file\n<div class="teammate-edits-here">TEAMMATE improved body text</div>\n}\n' > P/Ours.razor; git commit -qam "teammate: ours fix"
git mv P/Old.razor P/New.razor; git commit -qm "teammate: rename"
git rm -q P/Del.razor; git commit -qm "teammate: delete"
printf '<Project>\n  <ItemGroup>\n    <Compile Remove="verification/**" />\n  </ItemGroup>\n</Project>\n' > HomeFront.csproj; git commit -qam "teammate: csproj"
printf 'alpha line that stays\n' > P/Rm.razor; git commit -qam "teammate: removal only"
printf 'generic file body line here\n}\n' > P/Gen.razor; git commit -qam "teammate: generic only"
printf 'acknowledged original body line\nteammate change we deliberately rejected\n' > P/Ack.razor; git commit -qam "teammate: ack me"
git push -q origin HEAD:main 2>/dev/null
mkdir -p $T/repos; git clone -q $T/remote.git $T/repos/hyphen-pb 2>/dev/null
mkdir -p $T/src/P; (cd $T/src && git init -q)
printf 'OUR extra header line added locally\nline one of a long enough file\n<div class="teammate-edits-here">TEAMMATE improved body text</div>\n}\n' > $T/src/P/Ours.razor
printf 'keep this old named page body\n' > $T/src/P/Old.razor; cp $T/src/P/Old.razor $T/src/P/New.razor
printf 'a file the teammate will delete entirely\n' > $T/src/P/Del.razor
printf '<Project>\n  <ItemGroup>\n    <PackageReference Include="FlexKit" />\n  </ItemGroup>\n</Project>\n' > $T/src/HomeFront.csproj
printf 'alpha line that stays\nobsolete configuration line to remove\n' > $T/src/P/Rm.razor
printf 'historic original content here\nour own unrelated edit in hist\n' > $T/src/P/Hist.razor
printf 'generic file body line here\n' > $T/src/P/Gen.razor
printf 'acknowledged original body line\n' > $T/src/P/Ack.razor
sleep 1
(cd $T/src && git config user.email o@o && git config user.name o \
  && printf 'superseded original body line\nsuperseded teammate addition line\n' > P/Sup.razor && git add P/Sup.razor && git commit -qm "pull: took teammate sup" \
  && printf 'superseded original body line\nOUR later rewrite of that addition\n' > P/Sup.razor && git commit -qam "our rewrite")
source "$(cd "$(dirname "$0")/.." && pwd)/scripts/_common.sh"
c=$D/hyphen-pb; fails=0
expect() { got=$(classify_file hyphen-pb "$1" "$2" | cut -d'|' -f1); ok=FAIL; [ "$got" = "$3" ] && ok=PASS || fails=$((fails+1)); printf '  %-40s %-18s %-10s expected %-10s %s\n' "$4" "$2" "$got" "$3" "$ok"; }
sha() { git -C $c log --format=%h --grep="^$1" -1 origin/main; }
expect "$(sha 'teammate: ours fix')"     P/Ours.razor     ok        "merged, plus our own extra edit"
expect "$(sha 'teammate: rename')"       P/New.razor      ok        "rename: new path present"
expect "$(sha 'teammate: rename')"       P/Old.razor      RESURRECT "rename: old path still here"
expect "$(sha 'teammate: delete')"       P/Del.razor      RESURRECT "deleted upstream, still here"
expect "$(sha 'teammate: csproj')"       HomeFront.csproj HANDMERGE "transformed file, hunk missing"
expect "$(sha 'teammate: removal only')" P/Rm.razor       REVERT    "removal-only, not merged"
expect "$(sha 'teammate: generic only')" P/Gen.razor      REVERT    "only a '}' added (can't judge)"
expect "$(sha 'teammate: historic fix')" P/Hist.razor     LOST      "historic, never ours, overwritten"
expect "$(sha 'teammate: historic sup')" P/Sup.razor      NOTE      "historic, we had it then rewrote it"
expect "$(sha 'teammate: ack me')"       P/Ack.razor      REVERT    "unshipped, deliberately rejected"
printf '%s P/Ack.razor rejected on purpose — owner decision\n' "$(sha 'teammate: ack me')" > $T/H/Deploy/preflight-ack.txt
expect "$(sha 'teammate: ack me')"       P/Ack.razor      ok        "same, after recording the ack"
printf '<Project>\n  <ItemGroup>\n    <PackageReference Include="FlexKit" />\n    <Compile Remove="verification/**" />\n  </ItemGroup>\n</Project>\n' > $T/src/HomeFront.csproj
printf 'alpha line that stays\n' > $T/src/P/Rm.razor
printf 'historic original content here\nour own unrelated edit in hist\nhistoric teammate addition line\n' > $T/src/P/Hist.razor
expect "$(sha 'teammate: csproj')"       HomeFront.csproj ok        "after hand-merging the hunk"
expect "$(sha 'teammate: removal only')" P/Rm.razor       ok        "after removing the line"
expect "$(sha 'teammate: historic fix')" P/Hist.razor     ok        "after restoring the historic line"
echo "  ---- $fails failure(s)"
rm -rf $T
exit $fails
