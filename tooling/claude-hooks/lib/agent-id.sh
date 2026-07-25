#!/usr/bin/env bash
# lib/agent-id.sh — canonical source for subagent detection
#
# Provides: is_subagent <json>
#   Returns 0 (success) if the JSON has a TOP-LEVEL "agent_id" key → is a subagent.
#   Returns 1 if no top-level agent_id → is the main session.
#
# Algorithm (pure bash + sed/tr; no awk/jq/python):
#   1. Drop escaped quotes (\") so remaining double-quotes are balanced string delimiters.
#   2. Tag the KEY form  "agent_id" <ws>* :  with sentinel byte (\001) BEFORE blanking
#      strings. A string VALUE "agent_id" is followed by , or } (never :), so only
#      real keys get tagged.
#   3. Blank every string's CONTENTS ("..." -> "") so structural chars inside strings vanish.
#   4. Reduce to ONLY { } and the sentinel via tr -cd. This is O(structure), not O(payload).
#   5. Walk the skeleton counting brace depth; sentinel at depth 1 → top-level key → subagent.
#
# Safe to source. No side effects beyond defining is_subagent.
# NOTE: block-mainsession-exploration.sh contains a copy of this logic as has_top_level_agent_id.
# If this function changes, update that inline copy (or refactor it to source this lib).

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
