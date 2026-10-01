set -euo pipefail
err(){ echo "ERR: $*"; }; ok(){ echo "OK: $*"; }
scan() {
  local git_dir="$1" label="$2"
  local hits rc
  set +e
  hits=$(grep -rl --include='*.cs' --include='*.razor' -e 'UseDevAutoLogin' "$git_dir" 2>/dev/null)
  rc=$?
  set -e
  case $rc in
    0) err "$label: login skip survived the strip"; echo "$hits" | sed 's/^/    /'; return 1 ;;
    1) ok  "$label: no trace of the login skip" ;;
    *) err "$label: grep failed (rc=$rc)"; return 1 ;;
  esac
}
scan "$1" "t1" || echo "  (scan returned nonzero, handled)"
echo "rec3: reached end"
