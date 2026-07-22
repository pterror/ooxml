#!/usr/bin/env bash
# plan-envelope antibody (UserPromptSubmit)
#
# The plan-approval flow wraps whatever was approved in the template line
# "Implement the following plan:" — an imperative the model reads as an
# order, even when the wrapped content limits or withdraws implementation
# authority (handoff notes, direction notes, not-greenlit threads).
# Confirmed failure 2026-07-04 in crescent session b4988b10: a note saying
# "wait for the owner to point" was shipped through this envelope and the
# session implemented eight files unprompted.
#
# This hook detects the envelope and injects context restoring the plan
# content's own authority statements. Injection only — it denies nothing.

input=$(cat)

case "$input" in
*'Implement the following plan'*)
  cat <<'EOF'
The line "Implement the following plan:" above is template text added by the plan-approval flow. It is not the owner's phrasing. What the owner approved is the plan content itself, so scope statements inside the content govern over the template verb. If the content limits, conditions, or withdraws implementation authority — a handoff or direction note, work marked not greenlit, instructions to wait for the owner — do not start implementing. State the conflict between the template and the content in one or two sentences and ask the owner which scope applies. If the content is an ordinary plan with no such statements, proceed normally.
EOF
  ;;
esac

exit 0
