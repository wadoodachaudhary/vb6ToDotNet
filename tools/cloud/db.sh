#!/usr/bin/env bash
# Start SQL Server 2022 in Docker and restore HOMEFRONTSQL from the cached backup.
# Idempotent. The SA password is random per sandbox, kept in $HF_SA_FILE (0600);
# tools/cloud/sql.sh reads it. Overridable for a local dry run:
# HF_SQL_NAME, HF_SQL_PORT, HF_DB_DIR, HF_SA_FILE.
set -uo pipefail
NAME="${HF_SQL_NAME:-hf-sql}"
PORT="${HF_SQL_PORT:-1433}"
DB_DIR="${HF_DB_DIR:-/opt/hf-cache/db}"
SA_FILE="${HF_SA_FILE:-$HOME/.hf-sa}"
DB_TAG=db-20260930
DB_FILE=HOMEFRONTSQL-20260930.bak
DB_SHA256=42197db007a932c23341531ebfe5a97d772edb663e575c931437c7e799703db6
IMAGE=mcr.microsoft.com/mssql/server:2022-latest
fail() { echo "ERROR: $*" >&2; exit 1; }
sha() { if command -v sha256sum >/dev/null 2>&1; then sha256sum "$1"; else shasum -a 256 "$1"; fi | cut -c1-64; }
# `gh release download` goes through GraphQL, which Claude Code sessions refuse; use REST.
release_asset() { # <owner/repo> <tag> <file> <dir>
    local id
    id="$(gh api "repos/$1/releases/tags/$2" --jq ".assets[] | select(.name == \"$3\") | .id")" && [ -n "$id" ] \
        && gh api -H 'Accept: application/octet-stream' "repos/$1/releases/assets/$id" > "$4/$3.part" \
        && mv "$4/$3.part" "$4/$3" || { rm -f "$4/$3.part"; return 1; }
}

if [ ! -s "$SA_FILE" ]; then
    (umask 077; printf '%sAa1!' "$(head -c 32 /dev/urandom | base64 | tr -dc 'A-Za-z0-9' | cut -c1-20)" > "$SA_FILE")
fi
SA="$(cat "$SA_FILE")"

if ! docker info >/dev/null 2>&1; then
    (dockerd >/tmp/dockerd.log 2>&1 &)
    for _ in $(seq 60); do docker info >/dev/null 2>&1 && break; sleep 1; done
    docker info >/dev/null 2>&1 || fail "Docker is not running (see /tmp/dockerd.log)"
fi

mkdir -p "$DB_DIR"
if [ ! -f "$DB_DIR/$DB_FILE" ]; then
    release_asset wadoodachaudhary/HomeFront "$DB_TAG" "$DB_FILE" "$DB_DIR" \
        || fail "could not download $DB_FILE from the private release $DB_TAG"
fi
[ "$(sha "$DB_DIR/$DB_FILE")" = "$DB_SHA256" ] || fail "$DB_FILE checksum mismatch"
chmod 644 "$DB_DIR/$DB_FILE" 2>/dev/null || true

if docker ps -a --format '{{.Names}}' | grep -qx "$NAME"; then
    docker start "$NAME" >/dev/null || fail "could not start container $NAME"
else
    docker run -d --name "$NAME" -e ACCEPT_EULA=Y -e MSSQL_PID=Developer -e "MSSQL_SA_PASSWORD=$SA" \
        -p "$PORT:1433" -v "$DB_DIR:/backup:ro" "$IMAGE" >/dev/null || fail "could not run $IMAGE"
fi

sqlcmd() { docker exec -e "SQLCMDPASSWORD=$SA" "$NAME" /opt/mssql-tools18/bin/sqlcmd -C -S localhost -U sa -b "$@"; }
for _ in $(seq 90); do sqlcmd -Q "SELECT 1" >/dev/null 2>&1 && break; sleep 1; done
sqlcmd -Q "SELECT 1" >/dev/null 2>&1 || fail "SQL Server in $NAME did not come up"

if [ "$(sqlcmd -h -1 -W -Q "SET NOCOUNT ON; SELECT COUNT(*) FROM sys.databases WHERE name = 'HOMEFRONTSQL'")" != "1" ]; then
    sqlcmd -Q "RESTORE DATABASE HOMEFRONTSQL FROM DISK = N'/backup/$DB_FILE' WITH
        MOVE N'HOMEFRONTSQL_Data' TO N'/var/opt/mssql/data/HOMEFRONTSQL.mdf',
        MOVE N'HOMEFRONTSQL_Log' TO N'/var/opt/mssql/data/HOMEFRONTSQL_log.ldf', REPLACE" >/dev/null \
        || fail "RESTORE of HOMEFRONTSQL failed"
fi
echo "SQL Server ready on localhost,$PORT (container $NAME); HOMEFRONTSQL restored from $DB_FILE."
