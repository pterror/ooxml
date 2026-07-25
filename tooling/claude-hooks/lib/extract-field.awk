# extract-field.awk
# Generalization of extract-command.awk: reads the portion of harness JSON
# AFTER the "tool_input": split point on stdin and prints the first
# "<field>":"..." string value it finds, raw (still JSON-encoded, i.e. any
# \n \" \\ escapes are left un-decoded — same convention as extract-command.awk).
# Empty output (no exit-triggered print) if the field is absent.
#
# Usage: awk -v field="model" -f extract-field.awk
#
# Same state machine as extract-command.awk, parameterized on the target key
# instead of hardcoding "command".

BEGIN {
    state = 0
    result = ""
    RS = ""          # slurp whole input as one record
    FS = ""
    target = "\"" field "\""
    tlen = length(target)
}

{
    n = split($0, chars, "")
    for (i = 1; i <= n; i++) {
        c = chars[i]

        if (state == 0) {
            window = window c
            if (length(window) > tlen) {
                window = substr(window, length(window) - tlen + 1)
            }
            if (window == target) {
                state = 1
                window = ""
            }

        } else if (state == 1) {
            # Skip whitespace, wait for ':'
            if (c == ":") {
                state = 2
            } else if (c != " " && c != "\t" && c != "\n" && c != "\r") {
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
                printf "%s", result
                exit
            } else {
                result = result c
            }
        }
    }
}
