# -*- coding: utf-8 -*-
"""Merge generated slides into a PowerPoint-created skeleton.

Supports any number of slides: if the deck has more slides than the skeleton,
extra slide parts are cloned from the skeleton; if fewer, unused slide parts
are dropped from the package.

Usage:
    python assemble_pptx.py deck_config.json [--out outputs/xxx.pptx] [--skeleton work/_skeleton.pptx]
"""

import json
import os
import re
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

from build_pptx import make_deck

WORK = Path(__file__).parent
SKELETON = WORK / "_skeleton.pptx"
OUT_DIR = WORK.parent / "outputs"

args = [a for a in sys.argv[1:] if not a.startswith("--")]
config_path = Path(args[0]) if args else WORK / "deck_config.json"
config = json.loads(config_path.read_text(encoding="utf-8"))
out_name = config.get("output", "deck") + ".pptx"

if "--out" in sys.argv:
    i = sys.argv.index("--out")
    if i + 1 < len(sys.argv):
        out_name = Path(sys.argv[i + 1]).name
OUT = OUT_DIR / out_name

if "--skeleton" in sys.argv:
    i = sys.argv.index("--skeleton")
    if i + 1 < len(sys.argv):
        SKELETON = Path(sys.argv[i + 1])


def parse_sld_id_lst(xml):
    return [m for m in re.findall(r'<p:sldId id="(\d+)" r:id="(rId\d+)"/>', xml)]


def main():
    OUT_DIR.mkdir(exist_ok=True)
    slides, media = make_deck(config)
    n = len(slides)

    with zipfile.ZipFile(SKELETON) as zin:
        names = zin.namelist()
        parts = {name: zin.read(name) for name in names}

    skeleton_slides = sorted(
        int(re.search(r"slide(\d+)\.xml", p).group(1))
        for p in names if re.match(r"ppt/slides/slide\d+\.xml$", p)
    )
    sk_n = len(skeleton_slides)
    if sk_n == 0:
        raise SystemExit(f"skeleton has no slides: {SKELETON}")

    if media:
        ct = parts["[Content_Types].xml"].decode("utf-8")
        if 'Extension="png"' not in ct:
            ct = ct.replace('<Default Extension="jpeg" ContentType="image/jpeg"/>',
                            '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                            '<Default Extension="png" ContentType="image/png"/>')
        parts["[Content_Types].xml"] = ct.encode("utf-8")

    # Declare any extra slide parts in Content_Types (deck longer than skeleton).
    if n > sk_n:
        ct = parts["[Content_Types].xml"].decode("utf-8")
        for idx in range(sk_n + 1, n + 1):
            override = (f'<Override PartName="/ppt/slides/slide{idx}.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
            if f"/ppt/slides/slide{idx}.xml" not in ct:
                ct = ct.replace("</Types>", override + "</Types>")
        parts["[Content_Types].xml"] = ct.encode("utf-8")

    # Clone extra slide parts when the deck exceeds the skeleton slide count.
    for idx in range(sk_n + 1, n + 1):
        src = f"ppt/slides/slide{sk_n}.xml"
        if src not in parts:
            raise SystemExit(f"skeleton slide part missing: {src}")
        parts[f"ppt/slides/slide{idx}.xml"] = parts[src]
        src_rels = f"ppt/slides/_rels/slide{sk_n}.xml.rels"
        parts[f"ppt/slides/_rels/slide{idx}.xml.rels"] = parts.get(src_rels, b"")

    # Replace slide XML and append image relationships.
    for idx, xml, img in slides:
        parts[f"ppt/slides/slide{idx}.xml"] = xml.encode("utf-8")
        rels_name = f"ppt/slides/_rels/slide{idx}.xml.rels"
        rel = parts.get(rels_name, b"")
        if img:
            mi = media.index(img) + 1
            extra = (f"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
                     f"Target=\"../media/image{mi}.png\"/>").encode("utf-8")
            if b'Id="rId2"' not in rel:
                rel = rel.replace(b"</Relationships>", extra + b"</Relationships>")
        parts[rels_name] = rel

    for i, mp in enumerate(media, 1):
        parts[f"ppt/media/image{i}.png"] = open(mp, "rb").read()

    # Rebuild presentation.xml sldId list and rels to match the deck size.
    pres_xml = parts["ppt/presentation.xml"].decode("utf-8")
    sld_ids = parse_sld_id_lst(pres_xml)
    if len(sld_ids) != sk_n:
        raise SystemExit(f"unexpected skeleton slide count in presentation.xml: {len(sld_ids)}")
    keep_ids = list(sld_ids[:min(n, sk_n)])
    extra_rids = []
    if n > sk_n:
        max_id = max((int(iid) for iid, _ in sld_ids), default=256)
        pres_rels_raw = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
        used_rids = {int(m) for m in re.findall(r'Id="rId(\d+)"', pres_rels_raw)}
        next_rid = max(used_rids) + 1 if used_rids else 22
        for idx in range(sk_n + 1, n + 1):
            max_id += 1
            rid = f"rId{next_rid}"
            next_rid += 1
            keep_ids.append((str(max_id), rid))
            extra_rids.append((idx, rid))
    new_sld_lst = "<p:sldIdLst>" + "".join(
        f'<p:sldId id="{iid}" r:id="{rid}"/>' for iid, rid in keep_ids
    ) + "</p:sldIdLst>"
    pres_xml = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", new_sld_lst, pres_xml, flags=re.S)
    parts["ppt/presentation.xml"] = pres_xml.encode("utf-8")

    pres_rels = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    keep_rids = {rid for _, rid in keep_ids}
    rel_entries = re.findall(r'<Relationship ([^>]*?)/>', pres_rels)
    new_rels = []
    for entry in rel_entries:
        m = re.search(r'Id="(rId\d+)"', entry)
        rid = m.group(1) if m else None
        # Slide relationships have Type ending with ".../slide" (exact); the
        # master is ".../slideMaster" and must never be treated as a slide.
        is_slide = "/slide\"" in entry
        # Always keep master and non-slide rels (theme, presProps, viewProps,
        # tableStyles, ...); keep slide rels only for slides in the deck.
        if not is_slide or rid in keep_rids:
            new_rels.append(f"<Relationship {entry}/>")
    for idx, rid in extra_rids:
        new_rels.append(
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{idx}.xml"/>'
        )
    pres_rels = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 + "".join(new_rels) + "</Relationships>")
    parts["ppt/_rels/presentation.xml.rels"] = pres_rels.encode("utf-8")

    # Remove unused skeleton slide parts (when deck is shorter).
    for idx in range(n + 1, sk_n + 1):
        parts.pop(f"ppt/slides/slide{idx}.xml", None)
        parts.pop(f"ppt/slides/_rels/slide{idx}.xml.rels", None)
    if n < sk_n:
        ct = parts["[Content_Types].xml"].decode("utf-8")
        for idx in range(n + 1, sk_n + 1):
            ct = re.sub(
                r'<Override PartName="/ppt/slides/slide%d\.xml"[^>]*/>' % idx,
                "", ct
            )
        parts["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
    print("wrote", OUT, "%.1f MB" % (os.path.getsize(OUT) / 1048576),
          "| slides:", n, "| media:", len(media))


if __name__ == "__main__":
    main()
