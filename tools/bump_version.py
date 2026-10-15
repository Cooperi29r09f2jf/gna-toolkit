#!/usr/bin/env python3
"""Bump the toolkit version so PowerPoint picks up new code.

    python3 tools/bump_version.py 1.0.1

Updates the ?v= cache-busting suffix in taskpane.html and commands.html and
<Version> in manifest.xml (as 1.0.1.0). Push to GitHub afterwards. Only
re-upload manifest.xml in the admin center if you changed the manifest itself
(ribbon buttons, names, icons, URLs).
"""
import re
import sys
from pathlib import Path

if len(sys.argv) != 2 or not re.fullmatch(r"\d+\.\d+\.\d+", sys.argv[1]):
    sys.exit(__doc__)
v = sys.argv[1]
root = Path(__file__).resolve().parent.parent
for name in ("taskpane.html", "commands.html"):
    p = root / name
    p.write_text(re.sub(r"\?v=[\d.]+", "?v=" + v, p.read_text(encoding="utf-8")), encoding="utf-8")
m = root / "manifest.xml"
m.write_text(re.sub(r"<Version>[\d.]+</Version>", f"<Version>{v}.0</Version>", m.read_text(encoding="utf-8")), encoding="utf-8")
print("Version set to", v)
