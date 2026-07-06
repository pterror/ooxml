#!/usr/bin/env bash
# PreToolUse hook — main session is a pure orchestrator.
#
# Architecture: bash parameter-expansion split on "tool_input" separates
# harness-controlled fields (tool_name, agent_id) from model-controlled
# content (tool_input.command). No JSON depth-walker; no brace counters.
#
# Subagent detection: a genuine TOP-LEVEL agent_id key (brace-depth 1) → bypass.
# Order-independent and immune to the literal text "agent_id" inside tool_input
# (see has_top_level_agent_id). tool_name is still extracted only from $prefix.
#
# Bash allowlist: every semicolon/&&/||/pipe/newline-separated segment must
# start with an allowed (verb, subverb, ...) tuple.  Forbidden constructs
# ($( ` ${ eval source dot-source) are rejected before tokenizing.
#
# Awk scripts live in lib/ next to this file for independent testability.
# No python, jq, node, perl, ruby, nix-shell, or compiled binaries.

set -euo pipefail

dir="$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")"
input=$(head -c $((1024 * 1024)))

# ── debug log ────────────────────────────────────────────────────────────────
if [[ "${CLAUDE_HOOK_DEBUG:-}" == "1" ]]; then
    state_dir="${CLAUDE_HOOK_STATE_DIR:-/tmp/claude-state}"
    mkdir -p "$state_dir"
    debug_log="$state_dir/hook-input.debug.log"
    touch "$debug_log"
    chmod 600 "$debug_log"
    printf '\n=== %s ===\n' "$(date -Iseconds)" >> "$debug_log"
    printf '%s\n' "$input" >> "$debug_log"
    # Trim to 2000 lines
    tmp_log=$(mktemp)
    tail -n 2000 "$debug_log" > "$tmp_log" && mv "$tmp_log" "$debug_log"
    chmod 600 "$debug_log"
fi

# ── denial helper ─────────────────────────────────────────────────────────────
DENY_MSG="Main session is orchestrator only. Allowed: Agent/SendMessage/Task*/AskUserQuestion/EnterPlanMode/ExitPlanMode/SendUserFile/Skill/ToolSearch/ScheduleWakeup; Bash limited to git commit, git push, git status, git log --oneline (no chaining, no command substitution, no eval/source). Delegate everything else to a subagent."

deny() {
    local tool_name="$1"
    local extra="${2:-}"
    local reason="$DENY_MSG Denied tool: $tool_name."
    if [[ -n "$extra" ]]; then
        reason="$reason $extra"
    fi
    # JSON-escape the reason: \, then ", then tab/CR/LF via tr placeholders
    # Use awk to handle all control-char substitutions safely
    local escaped
    escaped=$(printf '%s' "$reason" | awk '
        {
            gsub(/\\/, "\\\\")
            gsub(/"/, "\\\"")
            gsub(/\t/, "\\t")
            gsub(/\r/, "\\r")
            # awk RS splits on \n; print adds \n between records but not in ORS
            printf "%s\\n", $0
        }
    ' | sed '$ s/\\n$//')
    printf '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"deny","permissionDecisionReason":"%s"}}\n' "$escaped"
    exit 0
}

# ── split on "tool_input" ─────────────────────────────────────────────────────
# prefix  = everything before the first occurrence of "tool_input"
# rest    = everything after  "tool_input":
prefix="${input%%\"tool_input\"*}"
rest="${input#*\"tool_input\":}"

# ── extract tool_name (only from prefix) ─────────────────────────────────────
tool_name=$(printf '%s' "$prefix" | grep -oE '"tool_name"\s*:\s*"[^"]*"' | head -1 | grep -oE '"[^"]*"$' | tr -d '"' || true)
if [[ -z "$tool_name" ]]; then
    exit 0  # no tool_name — fail open
fi

# ── subagent detection: TOP-LEVEL agent_id present → bypass ──────────────────
# JSON field order is NOT guaranteed: agent_id may be serialized before OR
# after tool_input. The old code only searched $prefix (text before
# "tool_input"), so when the harness emitted agent_id after tool_input the
# bypass silently missed and a subagent's Read fell through to deny.
#
# We must also NOT be fooled by the literal text "agent_id" appearing INSIDE
# tool_input (e.g. inside a subagent prompt that discusses agent_id). Only a
# genuine TOP-LEVEL "agent_id" key counts — i.e. one at JSON brace-depth 1.
#
# Technique (pure bash + grep/sed/tr; no awk/jq/python on this path):
#   1. Drop escaped quotes (\") so remaining double-quotes are balanced
#      string delimiters.
#   2. Tag the KEY form  "agent_id" <ws>* :  with a sentinel byte (\001)
#      BEFORE blanking strings. A string VALUE "agent_id" is followed by , or
#      } (never :), so only real keys get tagged.
#   3. Blank every string's CONTENTS ("..." -> "") so no structural-looking
#      char ({ } and stray text) survives inside string data. This kills the
#      false positive: braces/colons/"agent_id" text inside a prompt vanish.
#   4. Reduce to ONLY { } and the sentinel via `tr -cd`. The result is a tiny
#      structural skeleton regardless of payload size, so the depth scan below
#      is O(structure), not O(payload) — a 1 MiB prompt costs <40 ms.
#   5. Walk the skeleton counting brace depth; a sentinel seen at depth 1 is a
#      top-level agent_id key → subagent.
has_top_level_agent_id() {
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

if has_top_level_agent_id "$input"; then
    exit 0  # subagent — pass unconditionally
fi

# ── plan mode (permission_mode == "plan") → stand down ───────────────────────
# In plan mode the main session legitimately needs to read/write its own plan
# file; orchestrator delegation-enforcement is moot. Extract permission_mode
# only from $prefix (it is serialized before tool_input) so a "permission_mode"
# string inside tool_input can't false-positive.
perm_mode=$(printf '%s' "$prefix" | grep -oE '"permission_mode"[[:space:]]*:[[:space:]]*"[^"]*"' | head -1 | sed -E 's/.*:[[:space:]]*"([^"]*)"/\1/' || true)
if [[ "$perm_mode" == "plan" ]]; then
    exit 0
fi

# ── cost-tier enforcement: Agent / Workflow ──────────────────────────────────
# Cheapest-adequate-model discipline: no silent default to a frontier tier.
# COST_MSG text matches the marker checked for below ([frontier-approved]).
COST_MSG="Name the tier: cheapest adequate model (haiku for mechanical/extraction, sonnet for scripted implementation). Frontier tiers require user-approved cost: add model plus [frontier-approved] in the prompt after the user approves a cost estimate."

if [[ "$tool_name" == "Agent" ]]; then
    model_val=$(printf '%s' "$rest" | awk -v field="model" -f "$dir/lib/extract-field.awk")

    if [[ -z "$model_val" ]]; then
        deny "$tool_name" "$COST_MSG"
    fi

    if [[ "$model_val" == "fable" || "$model_val" == "opus" ]]; then
        if ! printf '%s' "$rest" | grep -qF '[frontier-approved]'; then
            deny "$tool_name" "$COST_MSG"
        fi
    fi
fi

# Workflow tool is disabled unconditionally by owner directive: unpredictable
# cost amplification (resume double-runs, echo stages) is not tier-gateable
# the way a single Agent call is. Re-enable only by owner editing this hook.
if [[ "$tool_name" == "Workflow" ]]; then
    deny "$tool_name" "Workflow tool disabled by owner directive 2026-07-03 (unpredictable cost amplification: resume double-runs, echo stages). Use individual tiered Agent calls. Re-enable only by owner editing this hook."
fi

# ── orchestration tools (always allowed) ─────────────────────────────────────
case "$tool_name" in
    Agent|SendMessage|Task|TaskCreate|TaskUpdate|TaskList|TaskGet|TaskOutput|TaskStop|\
    AskUserQuestion|EnterPlanMode|ExitPlanMode|SendUserFile|Skill|ToolSearch|ScheduleWakeup)
        exit 0
        ;;
esac

# ── mutation tools (never allowed in main) ───────────────────────────────────
case "$tool_name" in
    Edit|Write|NotebookEdit)
        deny "$tool_name"
        ;;
esac

# ── Bash (limited allowlist) ──────────────────────────────────────────────────
if [[ "$tool_name" == "Bash" ]]; then

    # Extract raw (JSON-encoded) command string from $rest via awk state machine
    cmd_raw=$(printf '%s' "$rest" | awk -f "$dir/lib/extract-command.awk")

    # Reject empty command
    if [[ -z "$cmd_raw" ]]; then
        deny "$tool_name" "Empty command."
    fi

    # JSON-decode the command string via awk.
    # Order of substitutions (to avoid double-decoding):
    # 1. \\ → placeholder (chr(1)) first
    # 2. \" → "
    # 3. \n → newline
    # 4. \t → tab
    # 5. \r → CR
    # 6. \b → backspace
    # 7. \f → form-feed
    # 8. \/ → /
    # 9. placeholder → \
    # If \uXXXX appears, deny conservatively (no Unicode support needed for git cmds).
    if printf '%s' "$cmd_raw" | grep -qE '\\u[0-9a-fA-F]{4}'; then
        deny "$tool_name" "Command contains \\uXXXX escape — denied conservatively."
    fi

    command=$(printf '%s' "$cmd_raw" | awk '
        BEGIN { RS = ""; ORS = "" }
        {
            gsub(/\\\\/, "\001")
            gsub(/\\"/, "\"")
            gsub(/\\n/, "\n")
            gsub(/\\t/, "\t")
            gsub(/\\r/, "\r")
            gsub(/\\b/, "\010")
            gsub(/\\f/, "\014")
            gsub(/\\\//, "/")
            gsub(/\001/, "\\")
            printf "%s", $0
        }
    ')

    # Forbidden constructs — check decoded command (conservative: includes inside quotes)
    if printf '%s' "$command" | grep -qF '$(' ; then
        deny "$tool_name" "Forbidden construct: \$( in command."
    fi
    if printf '%s' "$command" | grep -qF '`' ; then
        deny "$tool_name" "Forbidden construct: backtick in command."
    fi
    if printf '%s' "$command" | grep -qF '${' ; then
        deny "$tool_name" "Forbidden construct: \${ in command."
    fi
    if printf '%s' "$command" | grep -qE '(^|[[:space:];&|])eval([[:space:];&|]|$)' ; then
        deny "$tool_name" "Forbidden construct: eval in command."
    fi
    if printf '%s' "$command" | grep -qE '(^|[[:space:];&|])source([[:space:];&|]|$)' ; then
        deny "$tool_name" "Forbidden construct: source in command."
    fi
    if printf '%s' "$command" | grep -qE '(^|[[:space:];&|])\.([[:space:]/~]|$)' ; then
        deny "$tool_name" "Forbidden construct: dot-source in command."
    fi

    # Tokenize and allowlist-check each segment
    result=$(printf '%s' "$command" | awk -f "$dir/lib/tokenize-bash.awk")
    if [[ "$result" != "OK" ]]; then
        deny "$tool_name" "$result"
    fi

    exit 0
fi

# ── everything else (Read, Grep, Glob, NotebookRead, …) ──────────────────────
deny "$tool_name"
