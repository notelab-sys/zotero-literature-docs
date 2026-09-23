# -*- coding: utf-8 -*-
"""Extract readable text from the selected papers' PDFs (for reading).

Saves the first two pages and the last page of each PDF to
<work>/paper_texts/<KEY>.txt so the reviewer can read abstracts,
introductions and conclusions before writing.

Usage:
    python extract_texts.py KEY1,KEY2,...
    python extract_texts.py --keys-file keys.json
    python extract_texts.py            # uses all keys in work/data.json
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

OUT = WORK / "paper_texts"
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


def extract(key):
    pdf_path = resolve_pdf(key)
    if pdf_path is None:
        return "no pdf"
    doc = fitz.open(pdf_path)
    n = len(doc)
    pages = list(range(min(2, n)))
    if n > 2:
        pages.append(n - 1)
    text = []
    for pno in pages:
        t = doc[pno].get_text().strip()
        text.append(f"----- page {pno + 1}/{n} -----\n{t}")
    doc.close()
    out = OUT / f"{key}.txt"
    out.write_text("\n\n".join(text), encoding="utf-8")
    return f"{Path(pdf_path).name[:45]} | pages: {n} | chars: {sum(len(t) for t in text)}"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    keys = []
    if args:
        keys = [k.strip() for k in args[0].split(",") if k.strip()]
    if "--keys-file" in sys.argv:
        i = sys.argv.index("--keys-file")
        if i + 1 < len(sys.argv):
            keys = json.loads(Path(sys.argv[i + 1]).read_text(encoding="utf-8"))
    if not keys:
        data_path = WORK / "data.json"
        if data_path.exists():
            keys = [e["key"] for e in json.loads(data_path.read_text(encoding="utf-8"))]
        else:
            raise SystemExit("no keys given and no work/data.json found")
    for k in keys:
        print(k, "|", extract(k))


if __name__ == "__main__":
    main()
