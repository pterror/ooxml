# claude-hooks/

Behavioral hook scripts for the Claude Code harness, propagated to every
ecosystem repo by `tooling/propagate-harness.sh` and wired into
`.claude/settings.json`. They enforce the main-session-as-orchestrator model and
record session history. All shells are jq/python/node-free by design (the
harness does not always have those on PATH; on NixOS-from-flake setups they live
only in the project devshell).

- `inject-orchestrator-rules.sh` — UserPromptSubmit hook. In the main session it
  emits `orchestrator-rules.md` as additionalContext; in a subagent (detected via
  a top-level `agent_id` in the hook JSON) it exits silently.
- `post-history.sh` — UserPromptSubmit hook. Records session history;
  self-contained, with an inlined copy of the subagent detector.
- `block-blocking-bash.sh` — PreToolUse(Bash) hook. Denies commands that never
  return on their own (follow/stream/watch) and would hang the session until
  timeout; `run_in_background:true` is the sanctioned escape hatch.
- `block-mainsession-exploration.sh` — PreToolUse hook. Enforces that the main
  session is a pure orchestrator: only an allow-listed set of git verbs
  (commit/push/status/log) may run as Bash; subagents (top-level `agent_id`)
  bypass. Parses the payload via bash parameter expansion and the `lib/` awk
  scripts.
- `orchestrator-rules.md` — the rules text injected into the main session by
  `inject-orchestrator-rules.sh`.
- `orchestrator-workflows.md` — lessons that apply when running a Workflow in the
  main session; read before running one.
- `lib/agent-id.sh` — canonical `is_subagent <json>` subagent detector (pure
  bash + sed/tr). Sourced by `inject-orchestrator-rules.sh`.
- `lib/extract-command.awk` — extracts the `tool_input.command` string from the
  harness JSON for `block-mainsession-exploration.sh`.
- `lib/tokenize-bash.awk` — splits a decoded bash command into quote-aware
  segments and checks each against the git-verb allowlist.
- `SUMMARY.md` — this file.
