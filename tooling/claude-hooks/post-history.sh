#!/usr/bin/env bash
set -euo pipefail

input=$(cat)

# Self-contained top-level agent_id detector (inlined from lib/agent-id.sh).
is_subagent() {
    local skeleton
    skeleton=$(printf '%s' "$1" \
        | sed 's/\\"//g' \
        | sed 's/"agent_id"\([[:space:]]*\):/\x01\1:/g' \
        | sed 's/"[^"]*"/""/g' \
        | tr -cd '{}\001')
    local i ch depth=0 n=${#skeleton}
    for (( i = 0; i < n; i++ )); do
        ch="${skeleton:i:1}"
        case "$ch" in
            '{') (( depth++ )) ;;
            '}') (( depth-- )) ;;
            $'\001') (( depth == 1 )) && return 0 ;;
        esac
    done
    return 1
}

if is_subagent "$input"; then
    cat <<'EOF'
You're a subagent: you can't delegate further or ask the user. Gather with your own tools, don't invent to fill gaps. Return calibrated evidence — separate what you verified from what you inferred from what you couldn't confirm, and surface uncertainty. False completeness reported upward poisons the caller's context.
EOF
else
    cat <<'EOF'
CONSTRAINT — NEVER GUESS. Ask when unclear. On rejection: discard and re-derive, don't patch forward. Tradeoffs not verdicts: name options and costs, user chooses. Verify evidence or admit gaps. (Full version in CLAUDE.md.)
EOF
fi
