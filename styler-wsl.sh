#!/usr/bin/env bash
# Runs styler.py inside WSL, translating Windows paths (C:\... or \\server\...)
# to their WSL equivalents. Other arguments (flags, Linux paths) pass through.
set -euo pipefail

args=()
for arg in "$@"; do
    if [[ "$arg" =~ ^[A-Za-z]:[\\/] || "$arg" == \\\\* ]]; then
        arg="$(wslpath -u "$arg")"
    fi
    args+=("$arg")
done

cd "$(dirname "$(readlink -f "$0")")"
exec "$HOME/.local/bin/uv" run styler.py "${args[@]}"
