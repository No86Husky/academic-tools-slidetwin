#!/usr/bin/env bash
set -euo pipefail

INSTALL_DIR="${1:-$HOME/.codex/marketplaces/ppt-visual-tools}"
REPOSITORY="https://github.com/No86Husky/academic-tools.git"
PLUGIN="ppt-visual-reconstructor@ppt-visual-tools"

for command in git node python3 codex; do
  if ! command -v "$command" >/dev/null 2>&1; then
    echo "Required command not found: $command" >&2
    exit 1
  fi
done

mkdir -p "$(dirname "$INSTALL_DIR")"
if [[ -d "$INSTALL_DIR/.git" ]]; then
  git -C "$INSTALL_DIR" pull --ff-only
elif [[ -e "$INSTALL_DIR" ]]; then
  echo "Install directory exists but is not a Git checkout: $INSTALL_DIR" >&2
  exit 1
else
  git clone "$REPOSITORY" "$INSTALL_DIR"
fi

python3 -m pip install --user -r "$INSTALL_DIR/requirements.txt"
codex plugin marketplace add "$INSTALL_DIR"
codex plugin add "$PLUGIN"

echo "Installation complete. Restart Codex and start a new thread."
echo 'Upload one slide image and ask: Use $ppt-visual-reconstructor to recreate this image as an editable PowerPoint slide.'
