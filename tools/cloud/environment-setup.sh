#!/usr/bin/env bash
# Claude Code cloud environment setup for HomeFront (runs as root; its filesystem is
# snapshotted and reused by later sessions). Installs the .NET 10 SDK, pulls the
# SQL Server 2022 image, caches the private app and FlexKit mirrors in the Mac's
# layout, restores their NuGet packages, and downloads the database backup.
# Running processes are not cached: session-start.sh starts SQL Server per session.
set -uo pipefail
CACHE=/opt/hf-cache
ROOT="$CACHE/VBToCSharp"
DB_TAG=db-20260930
DB_FILE=HOMEFRONTSQL-20260930.bak
warn() { echo "WARN: $*" >&2; }
mkdir -p "$CACHE/db" "$ROOT/HomeFront/MobileSource"
# `gh release download` goes through GraphQL, which Claude Code sessions refuse; use REST.
release_asset() { # <owner/repo> <tag> <file> <dir>
    local id
    id="$(gh api "repos/$1/releases/tags/$2" --jq ".assets[] | select(.name == \"$3\") | .id")" && [ -n "$id" ] \
        && gh api -H 'Accept: application/octet-stream' "repos/$1/releases/assets/$id" > "$4/$3.part" \
        && mv "$4/$3.part" "$4/$3" || { rm -f "$4/$3.part"; return 1; }
}

if ! command -v dotnet >/dev/null 2>&1 || ! dotnet --list-sdks 2>/dev/null | grep -q '^10\.'; then
    curl -fsSL https://dot.net/v1/dotnet-install.sh -o /tmp/dotnet-install.sh \
        && bash /tmp/dotnet-install.sh --channel 10.0 --install-dir /usr/share/dotnet \
        && ln -sf /usr/share/dotnet/dotnet /usr/local/bin/dotnet \
        || warn ".NET 10 SDK install failed"
fi

if ! docker info >/dev/null 2>&1; then
    (dockerd >/tmp/dockerd-setup.log 2>&1 &)
    for _ in $(seq 60); do docker info >/dev/null 2>&1 && break; sleep 1; done
fi
# The browser checks launch Playwright with channel 'chrome' (Google Chrome, not the
# bundled Chromium). The sandbox already has the playwright package in /opt/node-tools.
if ! command -v google-chrome >/dev/null 2>&1; then
    curl -fsSL -o /tmp/chrome.deb https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb \
        && { apt-get update -qq || true; } \
        && DEBIAN_FRONTEND=noninteractive apt-get install -y -qq /tmp/chrome.deb \
        || warn "Google Chrome install failed"
    rm -f /tmp/chrome.deb
fi

docker pull -q mcr.microsoft.com/mssql/server:2022-latest \
    || warn "SQL Server image pull failed (network access must reach mcr.microsoft.com)"

clone() { # <dir> <url> <branch>
    [ -d "$1/.git" ] || git clone -q --branch "$3" "$2" "$1" || warn "could not clone $2 (is GitHub access granted for it?)"
}
clone "$ROOT/HomeFront/MobileSource/HomeFront" https://github.com/wadoodachaudhary/HomeFront.git main
clone "$ROOT/FlexKit" https://github.com/wadoodachaudhary/FlexKit.git telerik-parity-20260904
[ -f "$ROOT/HomeFront/MobileSource/HomeFront/HomeFront.csproj" ] && command -v dotnet >/dev/null 2>&1 \
    && { dotnet restore "$ROOT/HomeFront/MobileSource/HomeFront/HomeFront.csproj" -v q || warn "NuGet restore failed"; }

[ -f "$CACHE/db/$DB_FILE" ] || release_asset wadoodachaudhary/HomeFront "$DB_TAG" "$DB_FILE" "$CACHE/db" || warn "database backup download failed"
chmod 644 "$CACHE/db/$DB_FILE" 2>/dev/null || true
echo "HomeFront cloud environment setup finished."
