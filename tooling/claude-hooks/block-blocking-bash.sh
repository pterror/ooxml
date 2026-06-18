#!/usr/bin/env bash
# PreToolUse hook for Bash. Denies commands that never return on their own
# (follow/stream/watch). They hang the session until timeout.
#
# Claude can opt into a long-running command by passing run_in_background:true
# on the Bash tool call — that bypasses the hang because output is streamed.
#
# Jq-free by design: the Claude Code harness doesn't always have jq on PATH
# (on NixOS-from-flake setups it's only in the project devshell). We parse
# the minimal slice we need (tool_input.command) with sed.

set -euo pipefail

input=$(cat)

# Flatten input to one line, then extract the value of "command" inside
# "tool_input". The regex captures an escaped JSON string body:
#   (\\\\|\\"|[^"])* — escaped backslash, escaped quote, or any non-quote
# This is enough for the harness's well-defined payload shape. If extraction
# fails (no "command" found, malformed JSON, etc.), $cmd stays empty and the
# hook is a no-op (fail-open).
cmd=$(printf '%s' "$input" \
  | tr '\n' ' ' \
  | sed -nE 's/.*"command"[[:space:]]*:[[:space:]]*"((\\\\|\\"|[^"])*)".*/\1/p' \
  | head -1)

# Strip quoted regions and heredoc bodies from a scan copy so that
# `echo "tail -f"` or `git commit -m "block tail -f"` don't trip the rules.
# This is a 90% heuristic — ANSI-C $'...' quoting and escaped quotes inside
# regular strings can still leak — but it covers the cases that bite.
scan=$(printf '%s' "$cmd" \
  | sed -E "s/<<-?[[:space:]]*'?[A-Za-z_][A-Za-z0-9_]*'?.*$//" \
  | sed -E 's/"[^"]*"//g' \
  | sed -E "s/'[^']*'//g")

reason=""

# tail -f / -F / --follow (catches -f, -fn 100, -F, --follow=name)
if printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])tail\b[^|;&]*([[:space:]]-[a-zA-Z]*[fF][a-zA-Z]*\b|[[:space:]]--follow\b)'; then
  reason="tail follow mode"
fi

# journalctl -f / --follow
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])journalctl\b[^|;&]*([[:space:]]-[a-zA-Z]*f[a-zA-Z]*\b|[[:space:]]--follow\b)'; then
  reason="journalctl follow mode"
fi

# docker/podman/kubectl logs|attach -f / --follow
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])(docker|podman|kubectl)\b[^|;&]*\b(logs|attach)\b[^|;&]*([[:space:]]-[a-zA-Z]*f[a-zA-Z]*\b|[[:space:]]--follow\b)'; then
  reason="container logs/attach follow mode"
fi

# watch (always blocks)
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])watch[[:space:]]'; then
  reason="watch"
fi

# entr / fswatch / inotifywait without timeout
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])(entr|fswatch)\b'; then
  reason="filewatch (entr/fswatch)"
fi
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])inotifywait\b' && ! printf '%s' "$scan" | grep -qE '[[:space:]](-t|--timeout)[[:space:]=]'; then
  reason="inotifywait without -t timeout"
fi

# ping without -c (count) — finite usage requires -c N
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])ping\b' && ! printf '%s' "$scan" | grep -qE '[[:space:]]-c[[:space:]=]?[0-9]'; then
  reason="ping without -c N"
fi

# nc / netcat in listen mode without timeout (-w)
if [ -z "$reason" ] && printf '%s' "$scan" | grep -qE '(^|[[:space:];&|])(nc|netcat|ncat)\b[^|;&]*[[:space:]]-[a-zA-Z]*l\b' && ! printf '%s' "$scan" | grep -qE '[[:space:]]-w[[:space:]=]?[0-9]'; then
  reason="nc -l without -w timeout"
fi

if [ -n "$reason" ]; then
  # Hand-build JSON. The reason text is constant (no shell-special chars),
  # so printf %s embedding is safe. We deliberately don't echo $cmd back —
  # it could contain quotes/newlines that would need JSON-escaping.
  printf '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"deny","permissionDecisionReason":"Refused: %s — these never return and hang the session. Use a snapshot form (--lines N, --since, -c N, --max-events), or pass run_in_background:true to the Bash tool if you actually want a streaming background process."}}\n' "$reason"
fi
