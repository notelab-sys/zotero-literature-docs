# -*- coding: utf-8 -*-
"""Fetch clean item data from the Zotero local API.

Usage:
    python fetch_items.py KEY1,KEY2,...      -> writes work/data.json
    python fetch_items.py --keys-file keys.json

Stale/404 keys are skipped with a warning instead of crashing. Titles with
leading numbering noise (e.g. "1 Elucidation of ...") are cleaned.
"""

import json
import re
import sys
import unicodedata
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

WORK = Path(__file__).parent
OUT = WORK / "data.json"

DEFAULT_KEYS = [
    "6QE2L5E9", "PV5YPCL8", "XBVUUJWB", "AT3HAECV", "MSM4KA3M",
    "Y88RZPSH", "T9B93WYB", "NDL6R63T", "5VGFJIH5", "383V5M5C",
    "P79YVPTW", "WC56CA7U", "BXU37NHA", "86AC9X5X",
]


def normalize(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u2010", "-").replace("\u2011", "-")
    text = text.replace("\u2012", "-").replace("\u2013", "-")
    text = text.replace("\u2014", "-").replace("\u2015", "-")
    text = re.sub(r"[\uE000-\uF8FF]", "", text)
    # Repair GBK-style mojibake observed in some imported titles.
    text = text.replace("\u6c55", "\u03b2")   # 汕 -> 尾
    text = text.replace("\u6c19", "\u03b1")   # 汐 -> 伪
    text = text.replace("\u6c4e", "\u03b6")   # 汎 -> 味
    text = text.replace("\u576c", "-")        # 坼 -> hyphen
    text = re.sub(r"[ \t]+", " ", text).strip()
    return text


def clean_title(title):
    """Strip leading numbering noise from imported titles."""
    t = normalize(title)
    m = re.match(r"^\d+\s+(.{12,})", t)
    if m:
        candidate = m.group(1).strip()
        # Only strip when the remainder is long enough to be a real title.
        if len(candidate) >= 15:
            t = candidate
    return t


def authors_pretty(creators):
    names = []
    for c in creators:
        if c.get("creatorType") not in ("author", None):
            continue
        first = c.get("firstName", "").strip()
        last = c.get("lastName", "").strip()
        if first and last:
            names.append(f"{first} {last}")
        elif last:
            names.append(last)
    return ", ".join(names)


def fetch(keys, data_path=None, quiet=False):
    """Fetch items from the local API; skip missing keys with a warning."""
    keys = [k.strip() for k in keys if k.strip()]
    by_key = {}
    missing = []
    # The local API can silently drop items for very long itemKey lists;
    # keep requests small (25 keys each) to avoid missing entries.
    for i in range(0, len(keys), 25):
        chunk = keys[i:i + 25]
        url = ("http://127.0.0.1:23119/api/users/0/items?itemKey=" + ",".join(chunk)
               + "&format=json&limit=100")
        try:
            raw = json.load(urllib.request.urlopen(url, timeout=60))
        except Exception as exc:  # noqa: BLE001
            raise SystemExit(f"Zotero local API error: {exc}") from exc
        found = set()
        for it in raw:
            d = it.get("data", {})
            if d.get("itemType") == "attachment":
                continue
            k = it.get("key")
            if k:
                by_key[k] = d
                found.add(k)
        for k in chunk:
            if k not in found and k not in by_key:
                missing.append(k)

    entries = []
    for key in keys:
        d = by_key.get(key)
        if d is None:
            if not quiet:
                print(f"WARN: item {key} not found in library (skipped)")
            continue
        year = ""
        if d.get("date"):
            m = re.search(r"(\d{4})", d["date"])
            year = m.group(1) if m else ""
        entries.append(
            {
                "key": key,
                "itemType": d.get("itemType", ""),
                "title": clean_title(d.get("title", "")),
                "author": authors_pretty(d.get("creators", [])),
                "journal": normalize(d.get("publicationTitle", "")),
                "year": year,
                "volume": d.get("volume", ""),
                "number": d.get("issue", ""),
                "pages": d.get("pages", ""),
                "doi": d.get("DOI", ""),
                "abstract": normalize(d.get("abstractNote", "")),
            }
        )

    out = data_path or OUT
    out.write_text(json.dumps(entries, ensure_ascii=False, indent=1), encoding="utf-8")
    if not quiet:
        print(f"wrote {out} ({len(entries)} entries; {len(missing)} missing skipped)")
        for e in entries:
            print(e["key"], "|", e["title"][:55], "|", e["year"])
    return entries


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    keys = DEFAULT_KEYS
    if args:
        keys = [k.strip() for k in args[0].split(",") if k.strip()]
    for a in sys.argv[1:]:
        if a.startswith("--keys-file"):
            i = sys.argv.index(a)
            if i + 1 < len(sys.argv):
                keys = json.loads(Path(sys.argv[i + 1]).read_text(encoding="utf-8"))
    if not keys:
        print("no keys given; using defaults")
    fetch(keys)


if __name__ == "__main__":
    main()
