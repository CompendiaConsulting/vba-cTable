#!/bin/bash
# Run with: bash ./tools/scripts/make-txt-copies.sh
# Compatible with macOS's bundled Bash 3.2. Copies bytes without conversion.
set -euo pipefail

script_dir=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)
repo_root=$(cd -- "$script_dir/../.." && pwd -P)
output="$repo_root/txt-copies"

# Refuse a redirected output location before deleting anything.
if [[ -L "$output" || ( -e "$output" && ! -d "$output" ) ]]; then
    printf 'Error: txt-copies must be a real directory, not a link or file.\n' >&2
    exit 1
fi

# Collect sources before cleanup; a failed scan must not look successful.
manifest=$(mktemp "${TMPDIR:-/tmp}/vba-txt-copies.XXXXXXXX")
trap 'rm -f -- "$manifest"' EXIT
find "$repo_root" \( -path "$output" -o -name .git \) -prune -o \
    -type f \( -iname '*.bas' -o -iname '*.cls' \) -print0 > "$manifest"

rm -rf -- "$output"
mkdir -p -- "$output"
count=0
while IFS= read -r -d '' source; do
    relative=${source#"$repo_root/"}
    extension=${relative##*.}
    # Normalize the suffix even for .BAS/.CLS sources.
    case "$extension" in
        [bB][aA][sS]) suffix=bas ;;
        [cC][lL][sS]) suffix=cls ;;
    esac
    destination="$output/${relative%.*}_${suffix}.txt"
    mkdir -p -- "$(dirname -- "$destination")"
    # Do not silently overwrite two sources mapping to the same output name.
    if [[ -e "$destination" ]]; then
        printf 'Error: duplicate destination: %s\n' "$destination" >&2
        exit 1
    fi
    cp -- "$source" "$destination"
    count=$((count + 1))
done < "$manifest"
printf 'Created %s text copies in %s\n' "$count" "$output"
