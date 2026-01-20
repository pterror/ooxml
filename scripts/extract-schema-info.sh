#!/usr/bin/env bash
# Extract element names, types, and enum values from XSD schemas
set -euo pipefail

SCHEMA_DIR="$(dirname "$0")/../spec/OfficeOpenXML-XMLSchema-Transitional"
OUTPUT_DIR="$(dirname "$0")/../specs"

mkdir -p "$OUTPUT_DIR/elements" "$OUTPUT_DIR/types" "$OUTPUT_DIR/enums"

# Function to extract from a single XSD file
extract_from_xsd() {
    local xsd="$1"
    local basename="${xsd##*/}"
    local name="${basename%.xsd}"

    # Extract element names (from xsd:element name="...")
    grep -oP '<xsd:element[^>]* name="\K[^"]+' "$xsd" 2>/dev/null | sort -u > "$OUTPUT_DIR/elements/${name}.txt" || true

    # Extract complex type names
    grep -oP '<xsd:complexType[^>]* name="\K[^"]+' "$xsd" 2>/dev/null | sort -u > "$OUTPUT_DIR/types/${name}.txt" || true

    # Extract simple type names (often enums)
    grep -oP '<xsd:simpleType[^>]* name="\K[^"]+' "$xsd" 2>/dev/null | sort -u >> "$OUTPUT_DIR/types/${name}.txt" || true
    sort -u -o "$OUTPUT_DIR/types/${name}.txt" "$OUTPUT_DIR/types/${name}.txt" 2>/dev/null || true

    # Extract enum values
    grep -oP '<xsd:enumeration value="\K[^"]+' "$xsd" 2>/dev/null | sort -u > "$OUTPUT_DIR/enums/${name}.txt" || true

    # Count stats
    local elem_count=$(wc -l < "$OUTPUT_DIR/elements/${name}.txt" 2>/dev/null || echo 0)
    local type_count=$(wc -l < "$OUTPUT_DIR/types/${name}.txt" 2>/dev/null || echo 0)
    local enum_count=$(wc -l < "$OUTPUT_DIR/enums/${name}.txt" 2>/dev/null || echo 0)

    echo "$name: $elem_count elements, $type_count types, $enum_count enum values"

    # Remove empty files
    [[ -s "$OUTPUT_DIR/elements/${name}.txt" ]] || rm -f "$OUTPUT_DIR/elements/${name}.txt"
    [[ -s "$OUTPUT_DIR/types/${name}.txt" ]] || rm -f "$OUTPUT_DIR/types/${name}.txt"
    [[ -s "$OUTPUT_DIR/enums/${name}.txt" ]] || rm -f "$OUTPUT_DIR/enums/${name}.txt"
}

echo "Extracting schema information..."
echo ""

for xsd in "$SCHEMA_DIR"/*.xsd; do
    extract_from_xsd "$xsd"
done

echo ""
echo "Generating combined lists..."

# Combined element list with namespace prefixes
cat "$OUTPUT_DIR/elements/"*.txt 2>/dev/null | sort -u > "$OUTPUT_DIR/all-elements.txt" || true
cat "$OUTPUT_DIR/types/"*.txt 2>/dev/null | sort -u > "$OUTPUT_DIR/all-types.txt" || true
cat "$OUTPUT_DIR/enums/"*.txt 2>/dev/null | sort -u > "$OUTPUT_DIR/all-enums.txt" || true

echo "  all-elements.txt: $(wc -l < "$OUTPUT_DIR/all-elements.txt") elements"
echo "  all-types.txt: $(wc -l < "$OUTPUT_DIR/all-types.txt") types"
echo "  all-enums.txt: $(wc -l < "$OUTPUT_DIR/all-enums.txt") enum values"

echo ""
echo "Output written to: $OUTPUT_DIR"
