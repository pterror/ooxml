#!/usr/bin/env bash
# Share one target/ dir across the main checkout and every git worktree of
# this repo (mac/linux). For contributors who don't use direnv or
# `nix develop` -- both of those already set CARGO_TARGET_DIR dynamically,
# see .envrc and flake.nix. This script gets the same effect via
# .cargo/config.toml's [build] target-dir for anyone running plain `cargo`
# outside of direnv/nix.
#
# Run once per worktree, from inside that worktree:
#   scripts/setup-worktree-target.sh
#
# Safe to re-run (idempotent) and safe to run in a worktree that already
# gets CARGO_TARGET_DIR from direnv/nix -- CARGO_TARGET_DIR, if set, takes
# precedence over this file's target-dir, so this is inert for those users.
#
# NOTE: this edits .cargo/config.toml *in the worktree you run it from*.
# That file is git-tracked, so this intentionally makes it locally dirty
# (git status will show it modified). This is expected and is why the
# tracked copy of .cargo/config.toml deliberately omits target-dir -- see
# the comment at the top of that file. Do not commit this edit; it bakes
# in a machine-specific absolute path.
set -euo pipefail

cd "$(git rev-parse --show-toplevel)"

git_common_dir=$(git rev-parse --git-common-dir)
# Resolve symlinks the same way .envrc's `realpath` does, without depending
# on a `realpath` binary being present (not guaranteed on older macOS).
common_dir_abs=$(cd "$git_common_dir" && pwd -P)
target_dir="$(dirname "$common_dir_abs")/target"

config_file=".cargo/config.toml"
mkdir -p .cargo
touch "$config_file"

target_line="target-dir = \"$target_dir\""

if grep -qF "$target_line" "$config_file"; then
  echo "Already set: [build] $target_line in $config_file"
  exit 0
fi

if grep -qE '^target-dir[[:space:]]*=' "$config_file"; then
  # Stale target-dir line present (e.g. copied from a different machine or
  # a moved repo) -- replace it in place rather than duplicating the key.
  tmp=$(mktemp)
  sed -E "s|^target-dir[[:space:]]*=.*|$target_line|" "$config_file" > "$tmp"
  mv "$tmp" "$config_file"
  echo "Updated: [build] $target_line in $config_file"
elif grep -qE '^\[build\]' "$config_file"; then
  # [build] table exists but has no target-dir key -- insert right after
  # the table header.
  tmp=$(mktemp)
  awk -v line="$target_line" '
    { print }
    /^\[build\]/ && !done { print line; done=1 }
  ' "$config_file" > "$tmp"
  mv "$tmp" "$config_file"
  echo "Added to existing [build] table: $target_line in $config_file"
else
  # No [build] table at all -- append a new one.
  {
    echo ""
    echo "# Local-only: shared target dir across worktrees of this repo."
    echo "# Set by scripts/setup-worktree-target.sh or .ps1 -- do not commit this edit."
    echo "[build]"
    echo "$target_line"
  } >> "$config_file"
  echo "Appended new [build] table: $target_line in $config_file"
fi
