# -*- coding: utf-8 -*-
"""Extract figure images from literature PDFs.

Default mode: crop embedded raster figures (real figures, not whole-page
screenshots). Pass `--whole-page` to fall back to rendering full pages.

Usage:
    python extract_figures.py KEY1,KEY2,...
    python extract_figures.py --keys-file keys.json
    python extract_figures.py --keys-file keys.json --whole-page

Outputs go to <work>/ppt_images/<KEY>_f1.png, <KEY>_f2.png ...
"""

import json
import sys
import urllib.parse
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

WORK = Path(__file__).parent
sys.path.insert(0, str(WORK / "pylibs"))
import fitz  # noqa: E402

OUT = WORK / "ppt_images"
OUT.mkdir(exist_ok=True)


def resolve_pdf(parent_key):
    """Find the first PDF attachment of an item via the local API."""
    url = f"http://127.0.0.1:23119/api/users/0/items/{parent_key}/children?format=json"
    try:
        raw = json.load(urllib.request.urlopen(url, timeout=30))
    except Exception as exc:  # noqa: BLE001
        print("  children error", parent_key, exc)
        return None
    for it in raw:
        d = it.get("data", {})
        if d.get("itemType") != "attachment" or not (d.get("contentType") or "").endswith("pdf"):
            continue
        akey = it["key"]
        try:
            u = urllib.request.urlopen(
                f"http://127.0.0.1:23119/api/users/0/items/{akey}/file/view/url",
                timeout=30,
            ).read().decode("utf-8").strip()
            p = Path(urllib.parse.unquote(u[len("file:///"):]))
            if p.exists():
                return p
        except Exception as exc:  # noqa: BLE001
            print("  attachment error", akey, exc)
    return None


def find_figures(pdf_path, max_figs=2):
    """Return up to max_figs (page, rect, width_pt, height_pt) candidates."""
    doc = fitz.open(pdf_path)
    candidates = []
    seen = set()
    for pno in range(len(doc)):
        page = doc[pno]
        for im in page.get_image_info(xrefs=True):
            bbox = im.get("bbox")
            if not bbox:
                continue
            x0, y0, x1, y1 = bbox
            w, h = x1 - x0, y1 - y0
            # Real figures are large; skip logos / icons / decorations.
            if w < 180 or h < 120:
                continue
            if w > 680 or h > 600:
                continue
            key = (round(x0), round(y0), round(x1), round(y1))
            if key in seen:
                continue
            seen.add(key)
            candidates.append((w * h, pno, (x0, y0, x1, y1), w, h))
    doc.close()
    candidates.sort(key=lambda c: -c[0])
    return candidates[:max_figs]


def best_pages(pdf_path, limit=6):
    """Whole-page mode: pages with the most embedded images."""
    doc = fitz.open(pdf_path)
    n = min(len(doc), limit)
    scored = []
    for i in range(n):
        imgs = doc[i].get_images(full=True)
        scored.append((len(imgs), i))
    scored.sort(key=lambda x: -x[0])
    doc.close()
    if scored and scored[0][0] > 0:
        return [i for c, i in scored if c > 0][:2] or [0]
    return [0]


def render_region(pdf_path, pno, rect, out_path, zoom=3.0):
    doc = fitz.open(pdf_path)
    page = doc[pno]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=fitz.Rect(rect))
    pix.save(str(out_path))
    doc.close()
    return out_path.stat().st_size


def render_page(pdf_path, page_no, out_path, zoom=1.8):
    doc = fitz.open(pdf_path)
    page = doc[page_no]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    pix.save(str(out_path))
    doc.close()
    return out_path.stat().st_size


def extract(key, whole_page=False, max_figs=2):
    pdf_path = resolve_pdf(key)
    if pdf_path is None:
        return "no pdf"
    try:
        if whole_page:
            saved = []
            for i, pno in enumerate(best_pages(pdf_path)[:max_figs], 1):
                out = OUT / f"{key}_f{i}.png"
                size = render_page(pdf_path, pno, out)
                saved.append(f"p{pno + 1}={size}")
            return f"{Path(pdf_path).name[:40]} -> {saved}"
        figs = find_figures(pdf_path, max_figs=max_figs)
        saved = []
        for i, (area, pno, rect, w, h) in enumerate(figs, 1):
            out = OUT / f"{key}_f{i}.png"
            size = render_region(pdf_path, pno, rect, out)
            saved.append(f"p{pno + 1}x{w:.0f}x{h:.0f}={size}")
        return f"{Path(pdf_path).name[:40]} -> {saved}"
    except Exception as exc:  # noqa: BLE001
        return f"ERROR {exc}"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    keys = []
    if args:
        keys = [k.strip() for k in args[0].split(",") if k.strip()]
    if "--keys-file" in sys.argv:
        i = sys.argv.index("--keys-file")
        if i + 1 < len(sys.argv):
            keys = json.loads(Path(sys.argv[i + 1]).read_text(encoding="utf-8"))
    whole_page = "--whole-page" in sys.argv
    if not keys:
        raise SystemExit("no keys given; pass KEY1,KEY2,... or --keys-file keys.json")
    for k in keys:
        print(k, "|", extract(k, whole_page=whole_page))


if __name__ == "__main__":
    main()
