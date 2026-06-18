# tokenize-bash.awk
# Reads a decoded (non-JSON-encoded) bash command on stdin.
# Walks char-by-char tracking single-quote / double-quote / backslash state.
# On segment boundaries (;, &&, ||, |, newline) outside quotes, emits a segment
# for allowlist checking.
#
# Allowed leading-token tuples (every non-empty segment must match one):
#   git [-C <path>] commit ...
#   git [-C <path>] push ...
#   git [-C <path>] status ...
#   git [-C <path>] log --oneline ...
#
# The -C flag may appear as a separate token ("git -C /path status") or as an
# attached form ("git -C/path status"). In both cases the real subcommand is
# validated against the same allowlist as bare "git <sub>" commands.
#
# Prints "OK" if all segments pass, "DENY:<reason>" if any fail.

BEGIN {
    in_single = 0
    in_double = 0
    esc_next  = 0
    segment   = ""
    result    = "OK"
    RS        = ""     # slurp
    FS        = ""
}

function check_segment(seg,    tokens, n, t0, idx, t_sub, t_after) {
    # Strip leading/trailing whitespace
    gsub(/^[[:space:]]+|[[:space:]]+$/, "", seg)
    if (seg == "") return "OK"

    n = split(seg, tokens, /[[:space:]]+/)
    t0 = (n >= 1) ? tokens[1] : ""

    if (t0 == "git") {
        # Determine which token index holds the subcommand.
        # Accept an optional -C <path> flag before the subcommand:
        #   "git -C <path> <sub>"  → tokens: git, -C, <path>, <sub>  → idx 4
        #   "git -C<path> <sub>"   → tokens: git, -C<path>, <sub>    → idx 3
        #   "git <sub>"            → tokens: git, <sub>               → idx 2
        idx = 2
        if (n >= idx && tokens[idx] == "-C") {
            # Separate form: consume "-C" token and the following path token
            idx = idx + 2   # skip "-C" and "<path>", land on subcommand
        } else if (n >= idx && substr(tokens[idx], 1, 2) == "-C" && length(tokens[idx]) > 2) {
            # Attached form: "-C<path>" is one token, next token is subcommand
            idx = idx + 1
        }

        t_sub   = (n >= idx)     ? tokens[idx]     : ""
        t_after = (n >= idx + 1) ? tokens[idx + 1] : ""

        if (t_sub == "commit") return "OK"
        if (t_sub == "push")   return "OK"
        if (t_sub == "status") return "OK"
        if (t_sub == "log" && t_after == "--oneline") return "OK"
        return "DENY:git subcommand not allowed: " t_sub
    }
    return "DENY:command not allowed: " t0
}

function flush_segment(    r) {
    if (result != "OK") return
    r = check_segment(segment)
    if (r != "OK") result = r
    segment = ""
}

{
    n = split($0, chars, "")
    for (i = 1; i <= n; i++) {
        c = chars[i]
        next_c = (i < n) ? chars[i+1] : ""

        if (esc_next) {
            segment = segment c
            esc_next = 0
            continue
        }

        if (in_single) {
            if (c == "'") {
                in_single = 0
            } else {
                segment = segment c
            }
            continue
        }

        if (in_double) {
            if (c == "\\") {
                esc_next = 1
                segment = segment c
            } else if (c == "\"") {
                in_double = 0
            } else {
                segment = segment c
            }
            continue
        }

        # Outside quotes
        if (c == "\\") {
            esc_next = 1
            segment = segment c
            continue
        }
        if (c == "'") {
            in_single = 1
            continue
        }
        if (c == "\"") {
            in_double = 1
            continue
        }
        if (c == ";") {
            flush_segment()
            continue
        }
        if (c == "|") {
            flush_segment()
            # Consume second '|' if present (||)
            if (next_c == "|") i++
            continue
        }
        if (c == "&" && next_c == "&") {
            flush_segment()
            i++
            continue
        }
        if (c == "\n") {
            flush_segment()
            continue
        }
        segment = segment c
    }
    flush_segment()
}

END {
    print result
}
