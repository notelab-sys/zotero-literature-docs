"""One-shot runner: regenerate Word + PDF from work/data.json, then verify.

Usage: python run_all.py [config.json]
If a config file is given, it is copied to work/doc_config.json first.
"""
import json
import re
import shutil
import sys
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass
import zipfile
from pathlib import Path

WORK = Path(__file__).parent
sys.path.insert(0, str(WORK))
sys.path.insert(0, str(WORK / "pylibs"))
sys.path.insert(0, str(WORK / "pylibs2"))

import make_docx_v2  # noqa: E402
import make_pdf  # noqa: E402

if len(sys.argv) > 1:
    cfg_src = Path(sys.argv[1])
    shutil.copyfile(cfg_src, WORK / "doc_config.json")
    print("using config:", cfg_src)

make_docx_v2.main()
make_pdf.main()

OUT = WORK.parent / "outputs"
cfg = json.loads((WORK / "doc_config.json").read_text(encoding="utf-8"))
out_base = cfg.get("output", cfg.get("title", "文献"))
suspicious = set("\u9225\u604A\u94FF\u4E67\u788C\u4E6A\u4FF9\u50E3\u614F\u8133\u923C\u4E76\u617A\u4E7A\uE6A9\u20AC")
bad_pat = re.compile("[\uE000-\uF8FF\uFB00-\uFB06]")


def scan(name, text):
    hits = {ch for ch in text if ch in suspicious or bad_pat.match(ch)}
    status = "clean" if not hits else "HAS GARBLED: " + ",".join("U+%04X" % ord(h) for h in sorted(hits))
    print(name, status)


docx = OUT / (out_base + ".docx")
with zipfile.ZipFile(docx) as z:
    scan("docx", z.read("word/document.xml").decode("utf-8"))

import fitz  # noqa: E402

pdf = OUT / (out_base + ".pdf")
doc = fitz.open(pdf)
scan("pdf", "\n".join(p.get_text() for p in doc))
print("done")
