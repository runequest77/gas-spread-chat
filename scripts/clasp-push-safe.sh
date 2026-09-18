#!/usr/bin/env bash
set -euo pipefail

repo_root="$(git rev-parse --show-toplevel)"
cd "$repo_root"

if [[ ! -f .clasp.json ]]; then
  echo ".clasp.json does not exist." >&2
  exit 1
fi

if ! git ls-files --error-unmatch .clasp.json >/dev/null 2>&1; then
  echo ".clasp.json is not tracked by git. Commit the deployment target before pushing to GAS." >&2
  exit 1
fi

if ! git diff --quiet -- .clasp.json || ! git diff --cached --quiet -- .clasp.json; then
  echo ".clasp.json differs from HEAD. Commit the deployment target before pushing to GAS." >&2
  exit 1
fi

exec clasp push "$@"
