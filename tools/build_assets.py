#!/usr/bin/env python3
"""Build the GNA Toolkit slide assets from the GNA slide template.

Run this whenever the template changes:

    python3 tools/build_assets.py "path/to/THE Slide Template.pptx"

It writes, inside the repo:
    assets/slides/sNN.pptx     one file per template slide (the library)
    assets/starters/<id>.pptx  one file per deck starter (default lineups)
    assets/thumbs/sNN.jpg      library thumbnails       (needs LibreOffice)
    assets/guide/gNN.jpg       style guide images 1-3   (needs LibreOffice)
    src/catalog.js             catalog for the task pane (from tools/catalog.json)

Every generated slide carries two hidden PowerPoint tags:
    GNA_TEMPLATE_ID       the template slide number (e.g. 16)
    GNA_TEMPLATE_VERSION  the catalog version
The style checker uses them to know which slides need a source line.

Only the Python standard library is required. Thumbnails need LibreOffice
(soffice) and pdftoppm (poppler). Without them, use --skip-thumbs and export
thumbnails from PowerPoint instead (see README).
"""

import argparse
import json
import os
import shlex
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
CATALOG = REPO / "tools" / "catalog.json"

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
T_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
T_TAGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
CT_TAGS = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
DROP_REL_TYPES = (
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes",
    "http://schemas.microsoft.com/office/2015/10/relationships/revisionInfo",
    "http://schemas.microsoft.com/office/2016/11/relationships/changesInfo",
)

# Template text fixes applied to every generated asset: (slide number, old, new)
TEXT_FIXES = [
    (57, "Delive-rables", "Deliverables"),
]


# --------------------------------------------------------------------------
# Package helpers
# --------------------------------------------------------------------------

def rels_path_for(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def read_rels(root: Path, part: str):
    """Return list of (Id, Type, Target, TargetMode) for a part ('' = package)."""
    rp = root / ("_rels/.rels" if part == "" else rels_path_for(part))
    if not rp.exists():
        return []
    x = rp.read_text(encoding="utf-8")
    out = []
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", x):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        out.append((a.get("Id"), a.get("Type"), a.get("Target"), a.get("TargetMode", "")))
    return out


def resolve(part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = posixpath.dirname(part)
    return posixpath.normpath(posixpath.join(base, target))


def remove_rel(root: Path, part: str, rid_pred):
    rp = root / ("_rels/.rels" if part == "" else rels_path_for(part))
    x = rp.read_text(encoding="utf-8")

    def keep(m):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        return "" if rid_pred(a) else m.group(0)

    x = re.sub(r"<Relationship\b([^>]*)/>", keep, x)
    rp.write_text(x, encoding="utf-8")


def prune_unreachable(root: Path):
    """Delete every part that is not reachable from the package root."""
    seen = set()
    stack = [""]
    while stack:
        part = stack.pop()
        for _id, _type, target, mode in read_rels(root, part):
            if mode == "External" or not target:
                continue
            p = resolve(part, target) if part else target.lstrip("/")
            if p not in seen:
                seen.add(p)
                stack.append(p)
    keep = set(seen)
    keep |= {rels_path_for(p) for p in seen}
    keep |= {"[Content_Types].xml", "_rels/.rels"}
    for f in sorted(root.rglob("*")):
        if f.is_file():
            rel = f.relative_to(root).as_posix()
            if rel not in keep:
                f.unlink()
    for d in sorted([d for d in root.rglob("*") if d.is_dir()], reverse=True):
        if not any(d.iterdir()):
            d.rmdir()
    # Drop content-type overrides for parts that no longer exist
    ct = root / "[Content_Types].xml"
    x = ct.read_text(encoding="utf-8")
    x = re.sub(
        r'<Override PartName="/([^"]+)"[^>]*/>',
        lambda m: m.group(0) if (root / m.group(1)).exists() else "",
        x,
    )
    ct.write_text(x, encoding="utf-8")


def zip_dir(src: Path, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists():
        dest.unlink()
    files = sorted(p for p in src.rglob("*") if p.is_file())
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml first, as Office expects
        z.write(src / "[Content_Types].xml", "[Content_Types].xml")
        for p in files:
            arc = p.relative_to(src).as_posix()
            if arc != "[Content_Types].xml":
                z.write(p, arc)


# --------------------------------------------------------------------------
# Template-specific steps
# --------------------------------------------------------------------------

def slide_order(root: Path):
    """Return [(sldId id, rId, 'ppt/slides/slideX.xml'), ...] in presentation order."""
    pres = (root / "ppt/presentation.xml").read_text(encoding="utf-8")
    rels = {r[0]: r for r in read_rels(root, "ppt/presentation.xml")}
    out = []
    lst = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", pres, re.S).group(1)
    for m in re.finditer(r'<p:sldId id="(\d+)" r:id="([^"]+)"/>', lst):
        sid, rid = m.groups()
        out.append((sid, rid, resolve("ppt/presentation.xml", rels[rid][2])))
    return out


def prepare_base(template: Path, work: Path) -> Path:
    base = work / "base"
    with zipfile.ZipFile(template) as z:
        z.extractall(base)
    # Remove embedded add-in task panes, co-authoring history
    remove_rel(base, "", lambda a: a.get("Type") in DROP_REL_TYPES)
    remove_rel(base, "ppt/presentation.xml", lambda a: a.get("Type") in DROP_REL_TYPES)
    # Remove sections (they list slide ids that won't exist in the split files)
    pres_p = base / "ppt/presentation.xml"
    pres = pres_p.read_text(encoding="utf-8")
    pres = re.sub(r'<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}">.*?</p:ext>', "", pres, flags=re.S)
    pres = pres.replace("<p:extLst></p:extLst>", "")
    pres_p.write_text(pres, encoding="utf-8")
    # Template text fixes
    order = slide_order(base)
    for num, old, new in TEXT_FIXES:
        part = base / order[num - 1][2]
        x = part.read_text(encoding="utf-8")
        if old in x:
            part.write_text(x.replace(old, new), encoding="utf-8")
            print(f"  fixed '{old}' -> '{new}' on slide {num}")
    return base


def add_slide_tags(root: Path, slide_part: str, tags: dict):
    """Attach a PowerPoint tag list (p:custDataLst/p:tags) to a slide."""
    sp = root / slide_part
    x = sp.read_text(encoding="utf-8")
    body = "".join(f'<p:tag name="{k.upper()}" val="{v}"/>' for k, v in tags.items())
    # Existing slide-level tags? (custDataLst directly after spTree)
    m = re.search(r'</p:spTree><p:custDataLst>(.*?)</p:custDataLst>', x, re.S)
    existing = None
    if m:
        t = re.search(r'<p:tags r:id="([^"]+)"/>', m.group(1))
        if t:
            existing = t.group(1)
    if existing:
        rel = {r[0]: r for r in read_rels(root, slide_part)}[existing]
        tag_part = root / resolve(slide_part, rel[2])
        tx = tag_part.read_text(encoding="utf-8")
        tx = tx.replace("</p:tagLst>", body + "</p:tagLst>")
        tag_part.write_text(tx, encoding="utf-8")
        return
    # New tag part
    n = 1
    while (root / f"ppt/tags/gnatag{n}.xml").exists():
        n += 1
    tag_name = f"ppt/tags/gnatag{n}.xml"
    (root / "ppt/tags").mkdir(parents=True, exist_ok=True)
    (root / tag_name).write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        + body + "</p:tagLst>",
        encoding="utf-8",
    )
    ct = root / "[Content_Types].xml"
    cx = ct.read_text(encoding="utf-8")
    cx = cx.replace("</Types>", f'<Override PartName="/{tag_name}" ContentType="{CT_TAGS}"/></Types>')
    ct.write_text(cx, encoding="utf-8")
    # Relationship from slide to tag part
    rp = root / rels_path_for(slide_part)
    rx = rp.read_text(encoding="utf-8")
    rid = "rIdGna1"
    rx = rx.replace(
        "</Relationships>",
        f'<Relationship Id="{rid}" Type="{T_TAGS}" Target="../tags/gnatag{n}.xml"/></Relationships>',
    )
    rp.write_text(rx, encoding="utf-8")
    if 'xmlns:r=' not in x[:2000]:
        x = x.replace("<p:sld ", '<p:sld xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    if "</p:spTree><p:custDataLst>" in x:
        x = x.replace("</p:spTree><p:custDataLst>", f'</p:spTree><p:custDataLst><p:tags r:id="{rid}"/>', 1)
    else:
        x = x.replace("</p:spTree>", f'</p:spTree><p:custDataLst><p:tags r:id="{rid}"/></p:custDataLst>', 1)
    sp.write_text(x, encoding="utf-8")


T_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"


def prune_layouts(root: Path, slide_parts):
    """Keep only the slide layouts that the given slides use (keeps files small)."""
    used = set()
    for sp in slide_parts:
        for _id, typ, target, _m in read_rels(root, sp):
            if typ == T_LAYOUT:
                used.add(resolve(sp, target))
    for _id, typ, target, _m in read_rels(root, "ppt/presentation.xml"):
        if typ.endswith("/slideMaster"):
            mpart = resolve("ppt/presentation.xml", target)
            mrels = {r[0]: r for r in read_rels(root, mpart)}
            drop = {rid for rid, r in mrels.items() if r[1] == T_LAYOUT and resolve(mpart, r[2]) not in used}
            mp = root / mpart
            mx = mp.read_text(encoding="utf-8")
            mx = re.sub(r'<p:sldLayoutId id="\d+" r:id="([^"]+)"/>',
                        lambda m: "" if m.group(1) in drop else m.group(0), mx)
            mp.write_text(mx, encoding="utf-8")
            remove_rel(root, mpart, lambda a: a.get("Id") in drop)


def build_subset(base: Path, work: Path, numbers, dest: Path, version: str, slim: bool = False):
    """Write a .pptx containing only template slides `numbers`, in that order."""
    d = work / ("sub_" + dest.stem)
    if d.exists():
        shutil.rmtree(d)
    shutil.copytree(base, d)
    order = slide_order(d)
    pres_p = d / "ppt/presentation.xml"
    pres = pres_p.read_text(encoding="utf-8")
    entries = [f'<p:sldId id="{order[n - 1][0]}" r:id="{order[n - 1][1]}"/>' for n in numbers]
    pres = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "<p:sldIdLst>" + "".join(entries) + "</p:sldIdLst>", pres, flags=re.S)
    pres_p.write_text(pres, encoding="utf-8")
    keep_rids = {order[n - 1][1] for n in numbers}
    remove_rel(d, "ppt/presentation.xml", lambda a: a.get("Type") == T_SLIDE and a.get("Id") not in keep_rids)
    for n in numbers:
        add_slide_tags(d, order[n - 1][2], {"GNA_TEMPLATE_ID": n, "GNA_TEMPLATE_VERSION": version})
    if slim:
        prune_layouts(d, [order[n - 1][2] for n in numbers])
    prune_unreachable(d)
    app = d / "docProps/app.xml"
    if app.exists():
        ax = app.read_text(encoding="utf-8")
        ax = re.sub(r"<Slides>\d+</Slides>", f"<Slides>{len(numbers)}</Slides>", ax)
        app.write_text(ax, encoding="utf-8")
    zip_dir(d, dest)
    shutil.rmtree(d)
    return dest.stat().st_size


def build_thumbs(template: Path, work: Path, visible, guide):
    soffice = os.environ.get("GNA_SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice or not shutil.which("pdftoppm"):
        print("  LibreOffice or pdftoppm not found: skipping thumbnails (see README)")
        return
    pdf_dir = work / "pdf"
    pdf_dir.mkdir(exist_ok=True)
    src = work / "thumbsrc.pptx"
    zip_dir(work / "base", src)
    subprocess.run(shlex.split(soffice) + ["--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(src)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    pdf = pdf_dir / "thumbsrc.pdf"
    tdir = REPO / "assets/thumbs"
    gdir = REPO / "assets/guide"
    for d in (tdir, gdir):
        d.mkdir(parents=True, exist_ok=True)
    for n in visible:
        subprocess.run(["pdftoppm", "-jpeg", "-jpegopt", "quality=82", "-scale-to-x", "480", "-scale-to-y", "-1",
                        "-f", str(n), "-l", str(n), "-singlefile", str(pdf), str(tdir / f"s{n:02d}")], check=True)
    for n in guide:
        subprocess.run(["pdftoppm", "-jpeg", "-jpegopt", "quality=88", "-scale-to-x", "1280", "-scale-to-y", "-1",
                        "-f", str(n), "-l", str(n), "-singlefile", str(pdf), str(gdir / f"g{n:02d}")], check=True)
    print(f"  thumbnails: {len(visible)} library, {len(guide)} guide")


def write_catalog_js(cat):
    out = {
        "version": cat["version"],
        "noSourceSlides": cat["noSourceSlides"],
        "categories": cat["categories"],
        "slides": {k: v for k, v in cat["slides"].items()},
        "starters": cat["starters"],
    }
    js = ("/* Generated by tools/build_assets.py from tools/catalog.json. Do not edit by hand. */\n"
          "window.GNA_CATALOG = " + json.dumps(out, indent=1, ensure_ascii=False) + ";\n")
    (REPO / "src/catalog.js").write_text(js, encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("template", type=Path, nargs="?")
    ap.add_argument("--skip-thumbs", action="store_true")
    ap.add_argument("--only", type=int, nargs="*", help="only rebuild these slide numbers")
    ap.add_argument("--catalog-only", action="store_true", help="only regenerate src/catalog.js from tools/catalog.json")
    args = ap.parse_args()
    if args.catalog_only:
        write_catalog_js(json.loads(CATALOG.read_text(encoding="utf-8")))
        print("Wrote src/catalog.js.")
        return

    if not args.template:
        ap.error("template path is required (unless --catalog-only)")
    cat = json.loads(CATALOG.read_text(encoding="utf-8"))
    version = cat["version"]
    visible = sorted(int(k) for k in cat["slides"])
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        print("Preparing template...")
        base = prepare_base(args.template, work)
        total = len(slide_order(base))
        if total < max(visible):
            sys.exit(f"Template has {total} slides but the catalog references slide {max(visible)}.")
        print(f"  template has {total} slides")

        print("Building library slides...")
        sizes = []
        for n in (args.only or visible):
            sizes.append(build_subset(base, work, [n], REPO / f"assets/slides/s{n:02d}.pptx", version, slim=True))
        print(f"  {len(sizes)} files, avg {sum(sizes) / len(sizes) / 1e6:.2f} MB, max {max(sizes) / 1e6:.2f} MB")

        if not args.only:
            print("Building deck starters...")
            for st in cat["starters"]:
                size = build_subset(base, work, st["slides"], REPO / f"assets/starters/{st['id']}.pptx", version)
                print(f"  {st['id']}: {len(st['slides'])} slides, {size / 1e6:.2f} MB")

        if not args.skip_thumbs:
            print("Rendering thumbnails...")
            build_thumbs(args.template, work, args.only or visible, [1, 2, 3])

    write_catalog_js(cat)
    print("Wrote src/catalog.js. Done.")


if __name__ == "__main__":
    main()
