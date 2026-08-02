#!/usr/bin/env bash
# UserPromptSubmit hook — inject short style-discipline rules into every prompt.
#
# Unlike inject-orchestrator-rules.sh, this fires unconditionally (main session
# AND subagents): the rules it carries (concise, no overclaiming, no banned
# words) apply to any response the harness produces, not just orchestration.
#
# cat style-rules.md to stdout so the harness includes it as additionalContext
# for the prompt.
#
# No jq, python, or node — pure bash (matches the other hooks' pattern).

set -euo pipefail

dir="$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")"

cat "$dir/style-rules.md"
