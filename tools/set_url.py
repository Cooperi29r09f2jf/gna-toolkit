#!/usr/bin/env python3
"""Point manifest.xml at a new hosting URL.

    python3 tools/set_url.py https://your-name.github.io/gna-toolkit/

Then raise <Version> in manifest.xml and re-upload it in the Microsoft 365
admin center (Settings > Integrated apps).
"""
import re
import sys
from pathlib import Path

if len(sys.argv) != 2 or not sys.argv[1].startswith("https://"):
    sys.exit(__doc__)
new = sys.argv[1].rstrip("/") + "/"
p = Path(__file__).resolve().parent.parent / "manifest.xml"
x = p.read_text(encoding="utf-8")
old = re.search(r'<SourceLocation DefaultValue="(https://[^"]+/)taskpane\.html"', x).group(1)
x = x.replace(old, new)
origin = re.match(r"(https://[^/]+)", new).group(1)
x = re.sub(r"<AppDomain>[^<]*</AppDomain>", f"<AppDomain>{origin}</AppDomain>", x)
p.write_text(x, encoding="utf-8")
print(f"manifest.xml now points at {new} (was {old})")
