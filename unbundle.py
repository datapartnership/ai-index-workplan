#!/usr/bin/env python3
"""
Unbundle index.html → timeline.html for GitHub Pages.

The source file is a single-file bundle that uses JavaScript at runtime to
decode base64-encoded assets (fonts, images) and inject them as blob URLs.
GitHub Pages serves static HTML — so this script pre-resolves every asset
reference into a standard data-URI, producing a self-contained HTML file
that renders identically without any JavaScript unpacking.

All formatting, fonts, layout, colors, call-outs, and logos are preserved.

Usage:
    python3 unbundle.py                     # index.html → timeline.html
    python3 unbundle.py -i other.html       # custom input
    python3 unbundle.py -o output.html      # custom output
"""

import argparse
import base64
import gzip
import json
import re
import sys
from pathlib import Path


def extract_section(content: str, tag_type: str) -> str:
    """Extract the text content of a <script type="__bundler/..."> block."""
    opener = f'<script type="{tag_type}">'
    start = content.find(opener)
    if start < 0:
        return ""
    start += len(opener)
    end = content.find("</script>", start)
    if end < 0:
        return ""
    return content[start:end]


def build_data_uri(entry: dict) -> str:
    """Convert a manifest entry (base64 + optional gzip) to a data URI."""
    raw = base64.b64decode(entry["data"])
    if entry.get("compressed"):
        raw = gzip.decompress(raw)
    b64 = base64.b64encode(raw).decode("ascii")
    return f"data:{entry['mime']};base64,{b64}"


def unbundle(src: Path, dst: Path) -> None:
    content = src.read_text(encoding="utf-8")

    # ── 1. Parse manifest (uuid → {mime, data, compressed}) ──────────
    manifest_raw = extract_section(content, "__bundler/manifest")
    if not manifest_raw:
        sys.exit("Error: no __bundler/manifest found in " + str(src))
    manifest: dict = json.loads(manifest_raw)
    print(f"  Manifest: {len(manifest)} assets")

    # ── 2. Parse template (the actual HTML page, JSON-encoded) ───────
    template_raw = extract_section(content, "__bundler/template")
    if not template_raw:
        sys.exit("Error: no __bundler/template found in " + str(src))
    html: str = json.loads(template_raw, strict=False)
    print(f"  Template: {len(html):,} chars")

    # ── 3. Replace every UUID reference with its data URI ────────────
    replaced = 0
    for uuid, entry in manifest.items():
        if uuid not in html:
            print(f"  Warning: asset {uuid} ({entry['mime']}) not referenced in template")
            continue
        data_uri = build_data_uri(entry)
        html = html.replace(uuid, data_uri)
        replaced += 1
    print(f"  Replaced: {replaced} asset references")

    # ── 4. Verify no orphan UUIDs remain ─────────────────────────────
    leftover = re.findall(
        r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
        html,
    )
    if leftover:
        print(f"  Warning: {len(leftover)} UUID-like strings still present")

    # ── 5. Write output ──────────────────────────────────────────────
    dst.write_text(html, encoding="utf-8")
    size_kb = dst.stat().st_size / 1024
    print(f"  Output:   {dst}  ({size_kb:,.0f} KB)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Unbundle index.html into a GitHub-Pages-ready timeline.html"
    )
    parser.add_argument(
        "-i", "--input",
        default="index.html",
        help="Path to the bundled HTML file (default: index.html)",
    )
    parser.add_argument(
        "-o", "--output",
        default="timeline.html",
        help="Path for the output file (default: timeline.html)",
    )
    args = parser.parse_args()

    src = Path(args.input)
    dst = Path(args.output)

    if not src.exists():
        sys.exit(f"Error: {src} not found")

    print(f"Unbundling {src} → {dst}")
    unbundle(src, dst)
    print("Done.")


if __name__ == "__main__":
    main()
