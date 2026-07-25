# Orchestrator Workflows

These lessons apply when running a Workflow in the main session (orchestration tool). Read before running a Workflow. Lessons (observed 2026-05-30):

- **Resume does not adopt newly-passed `args`.** `resumeFromRunId` reuses the original run's args; args you pass on resume are ignored. Never branch run-mode (e.g. dry-run vs write) on an arg you intend to flip across a resume — it won't flip. Bake the mode into a script constant (the script IS re-read on resume) or use a separate script.
- **Never route large content through one agent for verbatim reproduction.** An agent asked to echo ~100k tokens is slow, costly, and silently truncates. The workflow JS sandbox cannot write files, so all writes go through agents — keep each agent's write payload small and batch many small files per agent, not one giant blob through one agent. For review data, prefer the workflow's structured return value over having an agent transcribe a report file.
- **A resume that produces no expected output is a signal — find the cause before patching a symptom.** (Here: the first write-resume wrote nothing and re-ran a giant report agent; the real cause was args not flipping across resume, not the report agent. Guarding the report agent alone did not fix it.)
- **Gate expensive fan-outs behind a dry-run, and confirm cache reuse before the costly stage.** Mining/read fan-out is the dominant cost; verify it's cached (not re-running) before resuming into write.

Lessons (observed 2026-07-03):

- **`agent()` defaults to the session-inherit model — name the tier explicitly.** For extraction/mechanical stages pass `model: 'haiku'` (small chunks equalize tiers; the positional-recall drop that motivates a bigger model vanishes at small chunk size); `model: 'sonnet'` for scripted implementation work. State a token/meter cost estimate to the user before launching any fan-out — do not let a frontier default happen silently.
- **A null `agent()` return (quota death, error) is a failure record, never a smoothed-over empty result.** A run that silently absorbed nulls into "no data" fabricated per-doc counts. Every null must be recorded as an explicit failure, not folded into the aggregate as if the item were processed.
- **Expensive stages need data-level checkpointing.** Persist per-item results to disk as they complete. The call-cache makes resume cheap, but a stopped run's unpersisted aggregations (sums, joins, reports built in-memory) are lost regardless of cache.
- **Writer/landing stages must be deterministic transforms, never a model echoing large payloads.** Build the output with jq/script over `journal.jsonl` or transcripts. Embedding a large payload in a model's prompt and asking it to reproduce it burned ~50k output tokens at frontier rates in one run and risks silent truncation besides.
- **A resume that appears to "re-run" agents may actually be completing in-flight units that never resulted.** Before diagnosing cache breakage, check the journal for duplicate keys — a duplicate can mean the unit is legitimately being finished, not recomputed from scratch.
