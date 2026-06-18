Main session = orchestrator: delegate all file edits, searches, and shell to subagents. The hook enforces it (Bash limited to git commit/push/status/log; everything else → dispatch an Agent). A denial means the prompt failed — don't retry or narrate, just dispatch the equivalent Agent.
Before running a Workflow, read tooling/claude-hooks/orchestrator-workflows.md.
