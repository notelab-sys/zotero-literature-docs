import json
import sys
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "pylibs"))

import make_papers_doc as m  # noqa: E402
from fpdf import FPDF  # noqa: E402

WORK = Path(__file__).parent
BIB = WORK.parent / "outputs" / "genetic_transformation.bib"
OUT = WORK.parent / "outputs" / "基因遗传转化文献.pdf"
(WORK.parent / "outputs").mkdir(exist_ok=True)
CONFIG = WORK / "doc_config.json"

DARK = (31, 78, 121)
GRAY = (90, 90, 90)
LIGHT = (110, 110, 110)
BLACK = (20, 20, 20)


class PapersPDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font("Deng", size=9)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"第 {self.page_no()} 页 / 共 {{nb}} 页", align="C")


def load_config():
    default_sections = [
        {"title": "一、综述与方法类", "keys": None},
        {"title": "二、兰花（蝴蝶兰等）转化研究", "keys": None},
    ]
    if CONFIG.exists():
        cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
        cfg.setdefault("title", "文献清单")
        cfg.setdefault("output", cfg["title"])
        cfg.setdefault("sections", default_sections)
        return cfg
    return {"title": "基因遗传转化相关文献", "output": "基因遗传转化文献.pdf", "sections": default_sections}


def split_sections(entries, sections):
    if all(s.get("keys") for s in sections):
        result = []
        for s in sections:
            wanted = set(s["keys"])
            result.append((s["title"], [e for e in entries if e["key"] in wanted]))
        return result
    half = len(entries) // 2
    return [(sections[0]["title"], entries[:half]), (sections[1]["title"], entries[half:])]


def main():
    entries = m.load_entries()
    cfg = load_config()
    sections = split_sections(entries, cfg["sections"])
    out = WORK.parent / "outputs" / (cfg["output"] + ".pdf")
    total = len(entries)
    assert total > 0

    pdf = PapersPDF(format="A4")
    pdf.set_margins(20, 18, 20)
    pdf.set_auto_page_break(True, margin=20)
    pdf.add_font("Deng", fname=r"C:\Windows\Fonts\Deng.ttf")
    pdf.add_font("Deng", style="B", fname=r"C:\Windows\Fonts\Dengb.ttf")
    pdf.set_title(cfg["title"])
    pdf.set_author("Zotero 文献整理")
    pdf.alias_nb_pages()
    pdf.add_page()

    pdf.set_font("Deng", "B", size=20)
    pdf.set_text_color(*DARK)
    pdf.multi_cell(0, 10, cfg["title"], align="C")
    pdf.ln(2)
    pdf.set_font("Deng", size=10)
    pdf.set_text_color(*LIGHT)
    pdf.multi_cell(0, 6, f"来源：Zotero 个人文献库     |     整理日期：2026-08-04     |     共 {total} 篇", align="C")
    pdf.ln(6)

    def section(title):
        pdf.ln(3)
        pdf.set_font("Deng", "B", size=14)
        pdf.set_text_color(*DARK)
        pdf.cell(0, 9, title, new_x="LMARGIN", new_y="NEXT")
        pdf.set_draw_color(*DARK)
        pdf.set_line_width(0.4)
        pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
        pdf.ln(4)

    def entry(idx, e):
        title = m.clean(e["title"])
        authors = e["author"]
        journal_bits = [m.clean(e["journal"])]
        if e["volume"]:
            vol = f'{m.clean(e["volume"])}({m.clean(e["number"])})' if e["number"] else m.clean(e["volume"])
            journal_bits.append(vol)
        if e["pages"]:
            journal_bits.append(f"pp. {m.clean(e['pages'])}")
        journal_bits.append(f"({m.clean(e['year'])})")
        meta = ", ".join(x for x in journal_bits if x)
        doi = f'https://doi.org/{m.clean(e["doi"])}' if e["doi"] else ""

        pdf.set_font("Deng", "B", size=10.5)
        pdf.set_text_color(*BLACK)
        pdf.multi_cell(0, 6, f"{idx}. {title}", align="L")
        pdf.set_x(pdf.l_margin)
        pdf.set_font("Deng", size=10)
        pdf.set_text_color(*GRAY)
        pdf.multi_cell(0, 5.5, authors, align="L")
        pdf.set_x(pdf.l_margin)
        pdf.multi_cell(0, 5.5, meta, align="L")
        pdf.set_x(pdf.l_margin)
        if doi:
            pdf.set_text_color(*DARK)
            pdf.write(5.5, "DOI: ")
            pdf.write(5.5, doi, link=doi)
            pdf.ln(0)
            pdf.set_x(pdf.l_margin)
        if e["abstract"]:
            ab = m.clean(e["abstract"])
            if len(ab) > 380:
                ab = ab[:380].rsplit(" ", 1)[0] + " ..."
            pdf.set_font("Deng", size=9)
            pdf.set_text_color(*LIGHT)
            pdf.multi_cell(0, 5, "摘要：" + ab, align="L")
            pdf.set_x(pdf.l_margin)
        pdf.ln(4)

    idx = 0
    for title, items in sections:
        section(title)
        for e in items:
            idx += 1
            entry(idx, e)

    pdf.output(str(out))
    print(f"wrote {out} ({out.stat().st_size} bytes, {pdf.pages_count} pages)")


if __name__ == "__main__":
    main()
