#!/usr/bin/env bash
# PreToolUse hook for Agent. Requires every subagent spawn to name an explicit
# subagent_type — no falling through to whatever default the harness would
# otherwise pick when the field is omitted.
#
# Context: the ecosystem replaced the default general-purpose persona with a
# much more restrictive one (literal instruction-executor, no unprompted
# action). Callers must now consciously choose between that restrictive
# default and any other more capable agent type for a given task, rather than
# silently getting whichever type the harness falls back to when the field is
# left out.
#
# Jq-free by design (matches block-blocking-bash.sh convention: the harness
# doesn't always have jq on PATH). We parse the minimal slice we need
# (tool_input.subagent_type) with sed.

set -euo pipefail

input=$(cat)

# Flatten input to one line, then extract the value of "subagent_type" inside
# "tool_input". Same escaped-string capture as block-blocking-bash.sh's
# "command" extraction. If extraction fails (no "subagent_type" found,
# malformed JSON, etc.), $subagent_type stays empty and the deny path below
# fires — fail CLOSED here, since an omitted field is exactly the case this
# hook exists to catch.
subagent_type=$(printf '%s' "$input" \
  | tr '\n' ' ' \
  | sed -nE 's/.*"subagent_type"[[:space:]]*:[[:space:]]*"((\\\\|\\"|[^"])*)".*/\1/p' \
  | head -1)

if [ -z "$subagent_type" ]; then
  printf '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"deny","permissionDecisionReason":"Agent call is missing subagent_type. The default persona is now a restrictive literal instruction-executor (no unprompted action) — name an explicit agent type (e.g. general-purpose, Explore, Plan, peer, or another listed type) so you consciously choose it over that restrictive default."}}\n'
  exit 0
fi
