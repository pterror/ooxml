#!/usr/bin/env bash
# Download ECMA-376 specification files
set -euo pipefail

SPEC_DIR="$(dirname "$0")/../spec"
mkdir -p "$SPEC_DIR"

# 5th edition (current) - individual parts
FIFTH_EDITION=(
    "https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip"
    "https://ecma-international.org/wp-content/uploads/ECMA-376-2_5th_edition_december_2021.zip"
    "https://ecma-international.org/wp-content/uploads/ECMA-376-3_5th_edition_december_2015.zip"
    "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip"
)

# Legacy editions (combined zips)
LEGACY_EDITIONS=(
    "https://ecma-international.org/wp-content/uploads/ECMA-376_1st_edition_december_2006.zip"
    "https://ecma-international.org/wp-content/uploads/ECMA-376_2nd_edition_december_2008.zip"
    "https://ecma-international.org/wp-content/uploads/ECMA-376_3rd_edition_june_2011.zip"
    "https://ecma-international.org/wp-content/uploads/ECMA-376_4th_edition_december_2012.zip"
)

download() {
    local url="$1"
    local filename="${url##*/}"
    local dest="$SPEC_DIR/$filename"

    if [[ -f "$dest" ]]; then
        echo "Already exists: $filename"
    else
        echo "Downloading: $filename"
        curl -fSL "$url" -o "$dest"
    fi
}

echo "Downloading 5th edition (current)..."
for url in "${FIFTH_EDITION[@]}"; do
    download "$url"
done

if [[ "${1:-}" == "--all" ]]; then
    echo ""
    echo "Downloading legacy editions..."
    for url in "${LEGACY_EDITIONS[@]}"; do
        download "$url"
    done
fi

echo ""
echo "Done. Specs downloaded to: $SPEC_DIR"
echo ""
echo "To extract (requires unzip or python):"
echo "  cd $SPEC_DIR && unzip -o '*.zip'"
echo "  # or: cd $SPEC_DIR && python -c \"import zipfile, glob; [zipfile.ZipFile(f).extractall() for f in glob.glob('*.zip')]\""
