#!/usr/bin/env python3
"""Download the OpenMoji illustration set from the npm registry and render
the icons used by the deck to transparent PNGs in `assets/`.

OpenMoji (https://openmoji.org) is CC BY-SA 4.0 — credit: OpenMoji project.

Usage:  pip install cairosvg && python3 fetch_assets.py
"""

import io
import os
import tarfile
import urllib.request

import cairosvg

NPM_TARBALL = "https://registry.npmjs.org/openmoji/-/openmoji-17.0.0.tgz"
OUT_DIR = "assets"
SIZE = 512

# semantic name -> unicode codepoint sequence (OpenMoji file name)
ICONS = {
    "globe": "1F30D", "island": "1F3DD", "star": "2B50",
    "flag_eu": "1F1EA-1F1FA", "medal": "1F3C5", "dove": "1F54A",
    "euro_note": "1F4B6", "crate": "1F4E6", "jeans": "1F456",
    "rivet": "1F529", "microphone": "1F3A4", "trophy": "1F3C6",
    "mountain": "1F3D4", "climber": "1F9D7", "butter": "1F9C8",
    "bacon": "1F953", "flag_dk": "1F1E9-1F1F0", "bonfire": "1F525",
    "fern": "1F33F", "parliament": "1F3DB", "ballot": "1F5F3",
    "wave": "1F30A", "amphora": "1F3FA", "dumpling": "1F95F",
    "blossom": "1F338", "soup": "1F963", "guitar": "1F3B8",
    "headphones": "1F3A7", "microscope": "1F52C", "orangutan": "1F9A7",
    "gorilla": "1F98D", "monkey": "1F412", "newspaper": "1F4F0",
    "scales": "2696", "tools": "1F6E0", "beaver": "1F9AB",
    "speech": "1F4AC", "recycle": "267B", "abacus": "1F9EE",
    "hand": "270B", "ship": "1F6A2", "scroll": "1F4DC",
    "deer": "1F98C", "tree": "1F332", "confetti": "1F389",
    "teams": "1F465", "target": "1F3AF", "memo": "1F4DD",
    "no_phone": "1F4F5", "brain": "1F9E0", "bulb": "1F4A1",
    "detective": "1F575",
}


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    print("Downloading OpenMoji from npm…")
    data = urllib.request.urlopen(NPM_TARBALL, timeout=120).read()
    wanted = {f"package/color/svg/{code}.svg": name
              for name, code in ICONS.items()}
    found = 0
    with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as tar:
        for member in tar:
            name = wanted.get(member.name)
            if not name:
                continue
            svg = tar.extractfile(member).read()
            out = os.path.join(OUT_DIR, f"{name}.png")
            cairosvg.svg2png(bytestring=svg, write_to=out,
                             output_width=SIZE, output_height=SIZE)
            found += 1
    missing = len(ICONS) - found
    print(f"Rendered {found} icons to {OUT_DIR}/" +
          (f" ({missing} MISSING!)" if missing else ""))


if __name__ == "__main__":
    main()
