# -*- coding: utf-8 -*-
"""Unified command line entry for the zotero-literature-docs pipeline.

Run from the project root (the directory that contains work/ and outputs/).
The first invocation copies the pipeline scripts into <cwd>/work/.

Commands:
    python zotero_docs.py setup                     copy scripts/skeleton into work/
    python zotero_docs.py status                    Zotero local API health
    python zotero_docs.py search <query>            search the library (read-only)
    python zotero_docs.py fetch KEY1,KEY2,...       write work/data.json
    python zotero_docs.py fetch --keys-file k.json
    python zotero_docs.py texts [KEY1,...]          extract PDF text for reading
    python zotero_docs.py review <content.json>     Word + PDF review (+validation)
    python zotero_docs.py figs [KEY1,...]           crop real figures from PDFs
    python zotero_docs.py figs ... --whole-page     old full-page rendering mode
    python zotero_docs.py ppt <deck_config.json>    build PPTX from a deck config
"""

import json
import shutil
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

ROOT = Path.cwd()
WORK = ROOT / "work"
SKILL = Path(__file__).resolve().parent
HAS_SKILL_MD = (SKILL.parent / "SKILL.md").exists()


def setup():
    if not HAS_SKILL_MD:
        print("setup was already run; scripts are present in work/")
        return
    WORK.mkdir(exist_ok=True)
    for py in SKILL.glob("*.py"):
        shutil.copyfile(py, WORK / py.name)
    skel = SKILL.parent / "assets" / "_skeleton.pptx"
    if skel.exists() and not (WORK / "_skeleton.pptx").exists():
        shutil.copyfile(skel, WORK / "_skeleton.pptx")
    print(f"scripts copied to {WORK}")
    print("note: copy pylibs/pylibs2 into work/ if not already present "
          "(see references/method.md)")


def run(script, *args):
    cmd = [sys.executable, str(WORK / script), *args]
    return subprocess.run(cmd, cwd=str(ROOT))


def status():
    try:
        req = urllib.request.urlopen("http://127.0.0.1:23119/api/users/0/items?format=json&limit=1", timeout=10)
        req.read()
        print("Zotero local API: OK")
    except Exception as exc:  # noqa: BLE001
        print("Zotero local API: NOT reachable ->", exc)
        return 1
    return 0


def search(query):
    url = ("http://127.0.0.1:23119/api/users/0/items?q="
           + urllib.parse.quote(query) + "&format=json&limit=50")
    try:
        raw = json.load(urllib.request.urlopen(url, timeout=30))
    except Exception as exc:  # noqa: BLE001
        print("search failed:", exc)
        return 1
    items = [it for it in raw if it.get("data", {}).get("itemType") != "attachment"]
    print(f"{len(items)} result(s) for: {query}")
    for it in items:
        d = it["data"]
        year = ""
        if d.get("date"):
            import re
            m = re.search(r"(\d{4})", d["date"])
            year = m.group(1) if m else ""
        print(it["key"], "|", year, "|", (d.get("title") or "")[:100])
    return 0


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 0
    cmd, rest = sys.argv[1], sys.argv[2:]

    if cmd == "setup":
        setup()
        return 0
    if cmd == "status":
        return status()
    if cmd == "search":
        if not rest:
            print("usage: zotero_docs.py search <query>")
            return 1
        return search(" ".join(rest))

    # All other commands need the copied scripts in work/.
    if not (WORK / "fetch_items.py").exists():
        setup()

    if cmd == "fetch":
        return run("fetch_items.py", *rest).returncode
    if cmd == "review":
        if not rest:
            print("usage: zotero_docs.py review <content.json> [--data work/data.json]")
            return 1
        return run("make_review_docs.py", *rest).returncode
    if cmd == "texts":
        return run("extract_texts.py", *rest).returncode
    if cmd == "figs":
        return run("extract_figures.py", *rest).returncode
    if cmd == "ppt":
        if not rest:
            print("usage: zotero_docs.py ppt <deck_config.json> [--out outputs/xxx.pptx]")
            return 1
        return run("assemble_pptx.py", *rest).returncode

    print(f"unknown command: {cmd}")
    print(__doc__)
    return 1


if __name__ == "__main__":
    sys.exit(main())
