#!/usr/bin/env bash
# Query the cloud HOMEFRONTSQL: bash tools/cloud/sql.sh "SELECT ..."  (or pipe the query on stdin).
# Never prints the SA password.
set -uo pipefail
NAME="${HF_SQL_NAME:-hf-sql}"
SA_FILE="${HF_SA_FILE:-$HOME/.hf-sa}"
q="${1:-$(cat)}"
docker exec -e "SQLCMDPASSWORD=$(cat "$SA_FILE")" "$NAME" /opt/mssql-tools18/bin/sqlcmd -C -S localhost -U sa \
    -d HOMEFRONTSQL -W -s '|' -Q "$q"
