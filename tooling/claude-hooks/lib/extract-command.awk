# extract-command.awk
# Reads the portion of harness JSON AFTER the "tool_input": split point on stdin.
# Finds the first "command":"..." value, handling \" and \\ escapes.
# Prints the raw (still JSON-encoded) string content — no final newline appended
# beyond what awk's print provides.
#
# State machine:
#   0 = scanning for "command"
#   1 = seen "command", waiting for ':'
#   2 = seen ':', waiting for '"'
#   3 = inside string value, collecting chars
#   4 = done (printed, exit)
#
# We emit the raw bytes of the JSON string (including escape sequences like \n, \",
# \\) so the caller can apply JSON string decoding in a subsequent step.

BEGIN {
    state = 0
    result = ""
    RS = ""          # slurp whole input as one record
    FS = ""
}

{
    n = split($0, chars, "")
    for (i = 1; i <= n; i++) {
        c = chars[i]

        if (state == 0) {
            # Look for the literal substring "command"
            # We accumulate a window of 9 chars (including the quote)
            window = window c
            if (length(window) > 9) {
                window = substr(window, length(window) - 8)
            }
            if (window == "\"command\"") {
                state = 1
                window = ""
            }

        } else if (state == 1) {
            # Skip whitespace, wait for ':'
            if (c == ":") {
                state = 2
            } else if (c != " " && c != "\t" && c != "\n" && c != "\r") {
                # Not whitespace and not ':', must be something else (e.g. another key
                # starting with "command") — reset
                state = 0
                window = ""
            }

        } else if (state == 2) {
            # Skip whitespace, wait for opening '"'
            if (c == "\"") {
                state = 3
            } else if (c != " " && c != "\t" && c != "\n" && c != "\r") {
                state = 0
                window = ""
            }

        } else if (state == 3) {
            # Inside the string value
            if (escaped) {
                result = result "\\" c
                escaped = 0
            } else if (c == "\\") {
                escaped = 1
            } else if (c == "\"") {
                # End of string
                printf "%s", result
                exit
            } else {
                result = result c
            }
        }
    }
}
