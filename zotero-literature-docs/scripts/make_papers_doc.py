# -*- coding: utf-8 -*-
"""Shared helpers for literature document generation.

make_docx_v2.py and make_pdf.py import load_entries() and clean() from here.
load_entries() prefers work/data.json (fetched from the Zotero local API by
fetch_items.py) and falls back to parsing a BibTeX file when present.
"""

import html
import json
import re
from pathlib import Path

WORK = Path(__file__).parent
BIB = WORK / "references.bib"


def extract_fields(body):
    """Extract BibTeX fields as {name: value}, handling nested braces."""
    fields = {}
    pos = 0
    while True:
        m = re.search(r"([A-Za-z]+)\s*=\s*\{", body[pos:])
        if not m:
            break
        name = m.group(1).lower()
        start = pos + m.end()
        depth = 1
        i = start
        while i < len(body) and depth:
            if body[i] == "{":
                depth += 1
            elif body[i] == "}":
                depth -= 1
            i += 1
        value = body[start : i - 1].strip()
        fields[name] = value
        pos = i
    return fields


def parse_bib(text):
    entries = []
    for m in re.finditer(r"@(\w+)\s*\{\s*([^,]+),\s*(.*?)\n\}", text, re.S):
        fields = extract_fields(m.group(3))
        entries.append(
            {
                "type": m.group(1),
                "key": m.group(2).strip(),
                "title": fields.get("title", ""),
                "author": fields.get("author", ""),
                "journal": fields.get("journal", ""),
                "year": fields.get("year", ""),
                "volume": fields.get("volume", ""),
                "number": fields.get("number", ""),
                "pages": fields.get("pages", ""),
                "doi": fields.get("doi", ""),
                "abstract": fields.get("abstract", ""),
            }
        )
    return entries


def load_entries():
    data_path = WORK / "data.json"
    if data_path.exists():
        return json.loads(data_path.read_text(encoding="utf-8"))
    if BIB.exists():
        return parse_bib(BIB.read_text(encoding="utf-8"))
    return []


def clean(value):
    """Remove BibTeX brace groups and LaTeX escapes for display."""
    value = re.sub(r"[{}]", "", value)
    value = value.replace(r"\%", "%").replace(r"\&", "&").replace(r"\_", "_")
    value = re.sub(r"\\[a-zA-Z]+", "", value)
    return html.escape(value.strip())
