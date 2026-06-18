#!/usr/bin/env bash
# UserPromptSubmit hook — inject orchestrator rules into the main session only.
#
# Subagents carry a top-level agent_id field in the hook stdin JSON.
# If detected, exit silently (no additionalContext injected).
# If absent (main session), cat orchestrator-rules.md to stdout so the harness
# includes it as additionalContext for the prompt.
#
# No jq, python, or node — pure bash + sed/tr (matches exploration hook pattern).

set -euo pipefail

dir="$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")"
input=$(head -c $((1024 * 1024)))

# shellcheck source=lib/agent-id.sh
source "$dir/lib/agent-id.sh"

if is_subagent "$input"; then
    exit 0  # subagent — emit nothing
fi

# Main session — inject orchestrator rules as additionalContext
cat "$dir/orchestrator-rules.md"
