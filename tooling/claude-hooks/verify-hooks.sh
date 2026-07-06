#!/usr/bin/env bash
# verify-hooks.sh [hooks-dir]
#
# Verifies the propagated behavioral hooks actually EXECUTE in a receiver —
# catches the class of failure where a hook ships without a helper it needs
# (e.g. lib/extract-field.awk omitted from propagation, 2026-07: every Agent
# call in every receiver crashed with `awk: fatal: cannot open source file`).
#
#   STATIC : every `$dir/<path>` file a hook references relative to its own
#            directory (lib/*.awk, sourced .sh helpers, cat'd .md) exists.
#   DYNAMIC: run each hook against the benign fixture at
#            lib/smoke/<hook-basename>.json — a payload its own allow/deny
#            contract should ALLOW cleanly (read the hook to learn its
#            contract before trusting a fixture). Fixtures ship as real files
#            beside this script, not inline literals, so they travel with it
#            through propagation and a receiver can inspect/extend them.
#            A few extra inline payloads below exercise deny paths and the
#            subagent bypass for additional regression coverage — a
#            deliberate, well-formed deny there is success, not failure.
#            A CRASH is always failure: nonzero exit, anything on stderr
#            (awk fatal, missing file, command not found), a benign fixture
#            denied, or an expected deny not produced.
#
# Runs anywhere: defaults to verifying the hooks next to itself, or pass a
# hooks directory explicitly (propagate-harness.sh runs the canonical copy
# against each receiver's tooling/claude-hooks). Fixtures always come from
# lib/smoke/ next to THIS script (the canonical copy), so an old or
# not-yet-propagated receiver can still be checked from the hub. Read-only:
# writes nothing.
#
# Exit: 0 all hooks pass; 1 any failure (per-case FAIL lines on stdout).
#
# No jq, python, or node — pure bash + grep/sed (matches the hooks' own rule).

set -uo pipefail

SCRIPT_DIR="$(dirname "$(readlink -f "${BASH_SOURCE[0]}")")"
FIXTURES_DIR="$SCRIPT_DIR/lib/smoke"
HOOKS_DIR="${1:-$SCRIPT_DIR}"
fail=0

note() { printf '%s\n' "$*"; }

# ── STATIC: every $dir-relative reference inside a hook must exist ────────────
static_check() {
    local hook="$1" name ref missing=0
    name="$(basename "$hook")"
    while IFS= read -r ref; do
        [ -n "$ref" ] || continue
        if [ ! -f "$HOOKS_DIR/$ref" ]; then
            note "FAIL static  $name: missing dependency $ref"
            fail=1; missing=1
        fi
    done < <(grep -oE '\$dir/[A-Za-z0-9_./-]+' "$hook" | sed 's|^\$dir/||' | sort -u)
    [ "$missing" -eq 0 ] && note "ok   static  $name"
}

# ── DYNAMIC: run the hook on a payload, classify allow / deny / crash ────────
#   expect=allow : exit 0, empty stderr, stdout carries no deny decision
#   expect=deny  : exit 0, empty stderr, stdout carries a deny decision
run_case() {
    local hook="$1" expect="$2" label="$3" payload="$4"
    local out err rc=0 errf
    [ -f "$HOOKS_DIR/$hook" ] || return 0   # absence already flagged
    errf="$(mktemp)"
    out=$(printf '%s' "$payload" | bash "$HOOKS_DIR/$hook" 2>"$errf") || rc=$?
    err="$(cat "$errf")"; rm -f "$errf"
    if [ "$rc" -ne 0 ]; then
        note "FAIL dynamic $hook [$label]: CRASH exit $rc${err:+ — stderr: $err}"
        fail=1; return
    fi
    if [ -n "$err" ]; then
        note "FAIL dynamic $hook [$label]: CRASH — stderr: $err"
        fail=1; return
    fi
    case "$expect" in
        allow)
            if printf '%s' "$out" | grep -qF '"permissionDecision":"deny"'; then
                note "FAIL dynamic $hook [$label]: benign payload was denied"
                fail=1; return
            fi ;;
        deny)
            if ! printf '%s' "$out" | grep -qF '"permissionDecision":"deny"'; then
                note "FAIL dynamic $hook [$label]: expected deny decision, got none"
                fail=1; return
            fi ;;
    esac
    note "ok   dynamic $hook [$label]"
}

# ── file-fixture wrapper: the required "benign payload this hook should
#    ALLOW" case, read from lib/smoke/<hook-basename>.json rather than an
#    inline literal ─────────────────────────────────────────────────────────
run_fixture() {
    local hook="$1" name fixture
    name="$(basename "$hook" .sh)"
    fixture="$FIXTURES_DIR/$name.json"
    if [ ! -f "$fixture" ]; then
        note "FAIL fixture $hook: no smoke fixture at $fixture"
        fail=1
        return
    fi
    run_case "$hook" allow fixture "$(cat "$fixture")"
}

# ── the propagated hook set (must match HOOK_FILES in propagate-harness.sh) ──
HOOKS="inject-orchestrator-rules.sh block-blocking-bash.sh block-mainsession-exploration.sh post-history.sh"

for h in $HOOKS; do
    if [ ! -f "$HOOKS_DIR/$h" ]; then
        note "FAIL missing $h (hook file absent — settings.json wires it)"
        fail=1
        continue
    fi
    static_check "$HOOKS_DIR/$h"
done

# Required per-hook smoke fixture (lib/smoke/<hook>.json): a benign payload
# each hook's own contract should ALLOW cleanly.
run_fixture inject-orchestrator-rules.sh
run_fixture post-history.sh
run_fixture block-blocking-bash.sh
run_fixture block-mainsession-exploration.sh

# Extra inline cases: bonus regression coverage beyond the required fixture,
# exercising deny paths and the subagent bypass. A deliberate, well-formed
# deny here is success, not failure.
run_case inject-orchestrator-rules.sh allow subagent \
    '{"session_id":"verify-smoke","agent_id":"verify-smoke","prompt":"verify smoke"}'
run_case block-blocking-bash.sh deny tail-follow \
    '{"tool_name":"Bash","tool_input":{"command":"tail -f /var/log/syslog"}}'

# PreToolUse (all tools): each case exercises a distinct lib/ helper.
#   subagent-bypass    → agent_id skeleton scan
#   read-mainsession   → deny path (deliberate deny is success, not failure)
run_case block-mainsession-exploration.sh allow subagent-bypass \
    '{"tool_name":"Read","agent_id":"verify-smoke","tool_input":{"file_path":"/x"}}'
run_case block-mainsession-exploration.sh deny read-mainsession \
    '{"tool_name":"Read","tool_input":{"file_path":"/x"}}'

if [ "$fail" -ne 0 ]; then
    note "hook verification FAILED in $HOOKS_DIR"
    exit 1
fi
note "hook verification passed in $HOOKS_DIR"
