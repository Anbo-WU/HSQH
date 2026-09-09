#!/usr/bin/env bash
# Optional Windows Git Bash entry point; uses the same account virtual environment.
set -eu
account_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
account_python="$account_dir/.venv/Scripts/python.exe"
if [[ ! -f "$account_python" ]]; then
    printf 'Python environment not found: %s\n' "$account_python" >&2
    exit 1
fi
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
exec "$account_python" "$account_dir/run.py" "$@"
