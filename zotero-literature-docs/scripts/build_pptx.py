# -*- coding: utf-8 -*-
"""Build a review PPTX (16:9) from a JSON deck config.

Slide content lives in <work>/deck_config.json (path can be overridden on the
command line). Figure slides are picture-only: title bar + centered figure +
citation caption. Text slides (cover, toc, conclusion, references) carry the
headings only.

Usage:
    python build_pptx.py [deck_config.json]   # writes outputs/<output>.pptx

Deck config schema:
{
  "output": "植物代谢与调控研究进展综述",
  "review_content": "review_content.json",
  "slides": [
    {"type": "cover", "title": "...", "subtitle": "...", "note1": "...", "note2": "..."},
    {"type": "toc", "items": ["一、...", "二、..."]},
    {"type": "pic", "title": "...", "subtitle": "...", "img": "WRS5RFG3", "fig": 1},
    {"type": "text", "title": "...", "lines": ["...", "..."]},
    {"type": "refs", "count": 10}
  ]
}
"""

import json
import os
import re
import struct
import sys
EN = False
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # noqa: BLE001
    pass

WORK = Path(__file__).parent
OUT_DIR = WORK.parent / "outputs"
IMG_DIR = WORK / "ppt_images"

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

SLIDE_W, SLIDE_H = 12192000, 6858000
MARGIN = 500000
TITLE_H = 1150000

DARK = "1F4E79"
GRAY = "595959"
BG = "EAF1FA"
BAR = "D6E4F0"

CONFIG_PATH = Path(sys.argv[1]) if len(sys.argv) > 1 else WORK / "deck_config.json"
CONFIG = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
REVIEW_CONTENT = CONFIG.get("review_content", "review_content.json")


def png_size(path):
    with open(path, "rb") as fh:
        head = fh.read(24)
    w, h = struct.unpack(">II", head[16:24])
    return w, h


def clean_title(title):
    title = re.sub(r"<[^>]+>", "", title or "")
    return title.replace("&amp;", "&").strip()


def ref_numbers(content_name):
    content = json.loads((WORK / content_name).read_text(encoding="utf-8"))
    order = []
    for sec in content["sections"]:
        for g in sec.get("groups", []):
            for item in g.get("items", []):
                for m in re.finditer(r"\[@([A-Z0-9]{8})\]", item):
                    if m.group(1) not in order:
                        order.append(m.group(1))
    return {k: i + 1 for i, k in enumerate(order)}


def short_author(author):
    if not author:
        return ""
    first = author.split(",")[0].strip()
    multi = author.count(",") > 0 or "  " in author
    if EN:
        return first + " et al." if multi else first
    return first + " 等" if multi else first


def surname_of(name):
    """Author display name: Chinese (CJK) names keep the full name; Latin /
    pinyin names are reduced to the surname (last token)."""
    name = (name or "").strip()
    if not name:
        return name
    if re.search(r"[\u4e00-\u9fff]", name):
        return name
    parts = name.split()
    if len(parts) > 1:
        return parts[-1]
    return name


def load_data():
    return {e["key"]: e for e in json.loads((WORK / "data.json").read_text(encoding="utf-8"))}


REFNO = ref_numbers(REVIEW_CONTENT)
DATA = load_data()


PPT_ITALIC = {
    "ARF", "ARF1", "ARF2", "ARF3", "ARF4", "ARF5", "ARF6", "ARF7", "ARF8",
    "ARF9", "ARF10", "ARF11", "ARF16", "ARF17", "ARF18", "ARF19", "ARF23",
    "SlARF2", "SlARF5", "SlARF7", "SlARF8", "SlARF9",
    "MONOPTEROS", "ETTIN", "TIR1", "AFB", "TPL", "TPR", "Aux/IAA",
    "IAA7", "IAA12", "IPT3", "IPT5", "IPT7", "LOG", "AHK4",
    "ARR5", "ARR7", "ARR15", "CLV3", "WUS", "WOX", "AG", "AP2", "LFY",
    "RhARF1", "RhARF8", "RhCHSa", "RhCHSc", "RhCHSa/c", "LCR", "miR394",
    "AUXIN RESPONSE FACTOR", "AUXIN RESPONSE FACTOR3", "AUXIN RESPONSE FACTOR 5",
}
PPT_GENERA = {
    "Arabidopsis", "Phalaenopsis", "Apostasia", "Dendrobium", "Vanilla",
    "Rosa", "Oryza", "Nicotiana", "Solanum", "Vitis", "Malus", "Prunus",
    "Citrus", "Triticum", "Zea", "Glycine", "Medicago", "Osmanthus",
    "Petunia", "Picea", "Pinus", "Gossypium", "Brassica", "Fragaria",
    "Vaccinium",
}
PPT_PROPER = {"Orchidaceae"}


def _merge_spans(spans):
    spans = sorted(set(spans))
    merged = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    return merged


def italic_spans(text):
    """Return merged (start, end) spans of gene names / Latin binomials."""
    spans = []
    for w in sorted(PPT_ITALIC, key=len, reverse=True):
        for m in re.finditer(r"(?<![A-Za-z0-9])" + re.escape(w) + r"(?![A-Za-z0-9])", text):
            spans.append((m.start(), m.end()))
    for g in PPT_GENERA:
        for m in re.finditer(r"\b" + g + r"\s+[a-z]{3,}\b", text):
            spans.append((m.start(), m.end()))
        for m in re.finditer(r"\b" + g + r"\b(?!\s+[a-z]{3,})", text):
            spans.append((m.start(), m.end()))
    for m in re.finditer(r"\b[A-Z]\.\s+[a-z]{3,}\b", text):
        spans.append((m.start(), m.end()))
    return _merge_spans(spans)


def en_scase(text, spans):
    """Sentence-case an English title, keeping italic spans and proper nouns."""
    def seg_case(seg, first):
        def repl(m):
            w = m.group(0)
            if w in PPT_PROPER:
                return w
            if not first[0]:
                first[0] = True
                return w[:1].upper() + w[1:].lower()
            return w.lower()
        return re.sub(r"[A-Za-z]+(?:['\u2019-][A-Za-z]+)*", repl, seg)

    out = []
    pos = 0
    first = [False]
    for s, e in spans:
        seg = text[pos:s]
        if seg:
            out.append(seg_case(seg, first))
            if re.search(r"[A-Za-z0-9]", seg):
                first[0] = True
        out.append(text[s:e])
        first[0] = True
        pos = e
    seg = text[pos:]
    if seg:
        out.append(seg_case(seg, first))
    return "".join(out)


def fmt_paragraph(text, size=1500, bold=False, color="333333", bullet=False, align="l", spacing=280,
                  numbered=False, ls=130, spans=None):
    """Return one paragraph XML (PPT format per SOP section 7, 2026-08-14)."""
    align = "ctr" if align == "c" else align

    def rpr(i=0):
        return [f"<a:rPr lang=\"zh-CN\" sz=\"{size}\" b=\"{1 if bold else 0}\" i=\"{i}\" dirty=\"0\">",
                "<a:latin typeface=\"Times New Roman\"/><a:ea typeface=\"SimSun\"/>",
                f"<a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:rPr>"]

    ppr = [f"<a:pPr algn=\"{align}\" marL=\"228600\" indent=\"-228600\">",
           f"<a:lnSpc><a:spcPct val=\"{ls}000\"/></a:lnSpc>",
           f"<a:spcBef><a:spcPts val=\"{spacing // 100}\"/></a:spcBef>"]
    if numbered:
        ppr.append("<a:buFont typeface=\"Times New Roman\"/><a:buAutoNum type=\"arabicPeriod\"/>")
    elif bullet:
        char = "*" if EN else "•"
        ppr.append(f"<a:buFont typeface=\"Arial\"/><a:buChar char=\"{char}\"/>")
    else:
        ppr.append("<a:buNone/>")
    ppr.append("</a:pPr>")
    if spans:
        runs = []
        pos = 0
        for s, e in spans:
            if s > pos:
                runs.append((text[pos:s], 0))
            runs.append((text[s:e], 1))
            pos = e
        if pos < len(text):
            runs.append((text[pos:], 0))
        body = "".join("<a:r>" + "".join(rpr(i)) + f"<a:t>{escape(t)}</a:t></a:r>"
                       for t, i in runs)
    else:
        body = "<a:r>" + "".join(rpr(0)) + f"<a:t>{escape(text)}</a:t></a:r>"
    return "<a:p>" + "".join(ppr) + body + "</a:p>"


def sp_text(x, y, cx, cy, paragraphs, sid=2):
    body = "<a:bodyPr wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"><a:noAutofit/></a:bodyPr><a:lstStyle/>"
    return (f"<p:sp><p:nvSpPr><p:cNvPr id=\"{sid}\" name=\"TextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>"
            f"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
            f"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>"
            f"<p:txBody>{body}{''.join(paragraphs)}</p:txBody></p:sp>")


def sp_rect(x, y, cx, cy, fill, sid=3):
    return (f"<p:sp><p:nvSpPr><p:cNvPr id=\"{sid}\" name=\"Rect\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
            f"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
            f"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            f"<a:solidFill><a:srgbClr val=\"{fill}\"/></a:solidFill></p:spPr>"
            f"<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"zh-CN\"/></a:p></p:txBody></p:sp>")


def sp_pic(x, y, cx, cy, rid, name="Figure", sid=5):
    return (f"<p:pic><p:nvPicPr><p:cNvPr id=\"{sid}\" name=\"{name}\"/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>"
            f"<p:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch/></p:blipFill>"
            f"<p:spPr><a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
            f"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>")


def slide_xml(shapes):
    body = ["<p:cSld><p:spTree>",
            "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>",
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
            "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>",
            *shapes,
            "</p:spTree></p:cSld>",
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"]
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<p:sld xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\">"
            + "".join(body) + "</p:sld>")


def pic_slide(spec):
    """Picture-only slide: title bar, centered figure, citation caption."""
    title = spec.get("title", "")
    subtitle = spec.get("subtitle")
    img_key = spec["img"]
    img_no = spec.get("fig", 1)
    caption_key = spec.get("caption_key", img_key)
    img = IMG_DIR / f"{img_key}_f{img_no}.png"
    title_paras = [fmt_paragraph(title, size=3000 if subtitle else 3200, bold=True, color=DARK,
                                 spans=italic_spans(title))]
    if subtitle:
        title_paras.append(fmt_paragraph(subtitle, size=2000, bold=False, color="2E75B6", spacing=120,
                                         spans=italic_spans(subtitle)))
    shapes = [sp_rect(0, 0, SLIDE_W, SLIDE_H, BG, sid=2),
              sp_rect(0, 0, SLIDE_W, TITLE_H, BAR, sid=3),
              sp_text(MARGIN, 100000, SLIDE_W - 2 * MARGIN, 900000, title_paras, sid=4)]
    img_path = None
    if img.exists():
        img_path = img
        iw, ih = png_size(img)
        box_w, box_h = 10800000, 4500000
        bx, by = (SLIDE_W - box_w) // 2, TITLE_H + 250000
        scale = min(box_w / iw, box_h / ih)
        dw, dh = int(iw * scale), int(ih * scale)
        ix, iy = bx + (box_w - dw) // 2, by + (box_h - dh) // 2
        shapes.append(sp_pic(ix, iy, dw, dh, "rId2", sid=5))
        cap = caption_for(caption_key)
        cap_paras = [fmt_paragraph(cap, size=1600, bold=True, color="404040", align="c", spacing=40,
                                   spans=italic_spans(cap))]
        src = spec.get("source")
        if src:
            cap_paras.append(fmt_paragraph(src, size=1100, color="595959", align="c", spacing=20,
                                           spans=italic_spans(src)))
        shapes.append(sp_text(MARGIN, by + box_h + 80000, SLIDE_W - 2 * MARGIN, 600000, cap_paras, sid=6))
    return slide_xml(shapes), img_path


def text_slide(spec):
    title = spec.get("title", "")
    subtitle = spec.get("subtitle")
    lines = spec.get("lines", [])
    title_paras = [fmt_paragraph(title, size=3000 if subtitle else 3200, bold=True, color=DARK,
                                 spans=italic_spans(title))]
    if subtitle:
        title_paras.append(fmt_paragraph(subtitle, size=2000, bold=False, color="2E75B6", spacing=120,
                                         spans=italic_spans(subtitle)))
    bullet_paras = [fmt_paragraph(b, size=2000, bold=True, bullet=True, spacing=700, ls=150,
                                  spans=italic_spans(b)) for b in lines]
    shapes = [sp_rect(0, 0, SLIDE_W, SLIDE_H, BG, sid=2),
              sp_rect(0, 0, SLIDE_W, TITLE_H, BAR, sid=3),
              sp_text(MARGIN, 100000, SLIDE_W - 2 * MARGIN, 900000, title_paras, sid=4),
              sp_text(MARGIN, TITLE_H + 250000, SLIDE_W - 2 * MARGIN, SLIDE_H - TITLE_H - 400000, bullet_paras, sid=5)]
    return slide_xml(shapes)


def cover_slide(spec):
    title_t = spec.get("title", "")
    subtitle_t = spec.get("subtitle", "")
    shapes = [sp_rect(0, 0, SLIDE_W, SLIDE_H, BG, sid=2),
              sp_rect(0, SLIDE_H - 140000, SLIDE_W, 140000, DARK, sid=3),
              sp_text(900000, 1600000, SLIDE_W - 1800000, 1700000,
                      [fmt_paragraph(title_t, size=4000, bold=True, color=DARK, align="c",
                                     spans=italic_spans(title_t))], sid=4),
              sp_text(900000, 3350000, SLIDE_W - 1800000, 700000,
                      [fmt_paragraph(subtitle_t, size=2200, color="2E75B6", align="c",
                                     spans=italic_spans(subtitle_t))], sid=5),
              sp_text(900000, 4300000, SLIDE_W - 1800000, 900000,
                      [fmt_paragraph(spec.get("note1", ""), size=1600, color="44546A", align="c"),
                       fmt_paragraph(spec.get("note2", ""), size=1400, color="44546A", align="c")], sid=6)]
    return slide_xml(shapes)


def toc_slide(spec):
    paras = [fmt_paragraph(it, size=2000, bullet=False, spacing=800, spans=italic_spans(it))
             for it in spec.get("items", [])]
    title = spec.get("title") or ("Contents" if EN else "目录")
    shapes = [sp_rect(0, 0, SLIDE_W, SLIDE_H, BG, sid=2),
              sp_rect(0, 0, SLIDE_W, TITLE_H, BAR, sid=3),
              sp_text(MARGIN, 200000, SLIDE_W - 2 * MARGIN, 700000,
                      [fmt_paragraph(title, size=3200, bold=True, color=DARK)], sid=4),
              sp_text(MARGIN, TITLE_H + 300000, SLIDE_W - 2 * MARGIN, SLIDE_H - TITLE_H - 600000, paras, sid=5)]
    return slide_xml(shapes)


def refs_slide(spec):
    count = spec.get("count", 10)
    title = spec.get("title") or ("References" if EN else "参考文献")
    def ref_sort(kv):
        e = DATA.get(kv[0], {})
        first = (e.get("author") or "").split(",")[0].strip()
        return (surname_of(first).lower(), int(e.get("year") or 0))
    refs_list = sorted(REFNO.items(), key=ref_sort)[:count]
    paras = []
    for key, n in refs_list:
        e = DATA.get(key, {})
        j = e.get("journal", "")
        v = e.get("volume", "")
        num = e.get("number", "")
        pg = e.get("pages", "")
        vnp = ""
        if v:
            vnp = v + (f"({num})" if num else "") + (f": {pg}" if pg else "")
        first = (e.get("author") or "").split(",")[0].strip()
        multi = "," in (e.get("author") or "")
        who = (surname_of(first) + " et al.") if multi else surname_of(first)
        raw_title = clean_title(e.get("title", ""))
        title_txt = en_scase(raw_title, italic_spans(raw_title))
        lead = f"{who}, {e.get('year', '')}. "
        tail = (" " + j + (f", {vnp}" if vnp else "")) if j else ""
        full = f"{lead}{title_txt}.{tail}."
        spans = [(s + len(lead), e + len(lead)) for s, e in italic_spans(title_txt)]
        paras.append(fmt_paragraph(full, size=1350, color="333333", bullet=False, spacing=200, spans=spans))
    shapes = [sp_rect(0, 0, SLIDE_W, SLIDE_H, BG, sid=2),
              sp_rect(0, 0, SLIDE_W, TITLE_H, BAR, sid=3),
              sp_text(MARGIN, 200000, SLIDE_W - 2 * MARGIN, 700000,
                      [fmt_paragraph(title, size=3200, bold=True, color=DARK)], sid=4),
              sp_text(MARGIN, TITLE_H + 200000, SLIDE_W - 2 * MARGIN, SLIDE_H - TITLE_H - 300000, paras, sid=5)]
    return slide_xml(shapes)


def caption_for(key):
    n = REFNO.get(key, "?")
    e = DATA.get(key, {})
    author = short_author(e.get("author", ""))
    year = e.get("year", "")
    raw_title = clean_title(e.get("title", ""))
    title = en_scase(raw_title, italic_spans(raw_title))

    if EN:
        journal = e.get("journal", "")
        first = (e.get("author") or "").split(",")[0].strip()
        multi = "," in (e.get("author") or "")
        who = (surname_of(first) + " et al.") if multi else surname_of(first)
        return f"{title} ({who}, {year}, {journal})"
    journal = e.get("journal", "")
    first = (e.get("author") or "").split(",")[0].strip()
    multi = "," in (e.get("author") or "")
    if multi:
        who = surname_of(first) + (" et al." if not re.search(r"[\u4e00-\u9fff]", first) else " 等")
    else:
        who = surname_of(first)
    return f"{title}（{who}，{year}，{journal}）"


def make_deck(config=None):
    global REFNO, DATA, REVIEW_CONTENT, EN
    cfg = config or CONFIG
    EN = str(cfg.get("lang", "")).lower() == "en"
    REVIEW_CONTENT = cfg.get("review_content", "review_content.json")
    REFNO = ref_numbers(REVIEW_CONTENT)
    DATA = load_data()
    slides = []
    media = []

    def add_slide(xml, img_path=None):
        idx = len(slides) + 1
        slides.append((idx, xml, img_path))
        if img_path:
            media.append(img_path)

    for spec in cfg.get("slides", []):
        stype = spec.get("type", "pic")
        if stype == "cover":
            add_slide(cover_slide(spec))
        elif stype == "toc":
            add_slide(toc_slide(spec))
        elif stype == "pic":
            add_slide(*pic_slide(spec))
        elif stype == "text":
            add_slide(text_slide(spec))
        elif stype == "refs":
            add_slide(refs_slide(spec))
        else:
            raise SystemExit(f"unknown slide type: {stype}")
    return slides, media


def build():
    slides, media = make_deck()
    write_pptx(slides, media)


def write_pptx(slides, media):
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / (CONFIG.get("output", "deck") + ".pptx")
    media_names = [(f"image{i}.png", mp) for i, mp in enumerate(media, 1)]

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(slides)))
        z.writestr("_rels/.rels", root_rels())
        z.writestr("docProps/core.xml", core_props())
        z.writestr("docProps/app.xml", app_props(len(slides)))
        z.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels(len(slides)))
        z.writestr("ppt/theme/theme1.xml", theme_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml", master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels())
        z.writestr("ppt/slideLayouts/slideLayout1.xml", layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels())
        for idx, xml, img in slides:
            z.writestr(f"ppt/slides/slide{idx}.xml", xml)
            rel = (f"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                   f"<Relationships xmlns=\"{PKG_NS}\">"
                   "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>")
            if img:
                mi = media.index(img) + 1
                rel += (f"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
                        f"Target=\"../media/image{mi}.png\"/>")
            rel += "</Relationships>"
            z.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", rel)
        for i, (name, mp) in enumerate(media_names, 1):
            z.writestr(f"ppt/media/{name}", open(mp, "rb").read())
    print("wrote", out, os.path.getsize(out), "bytes; slides:", len(slides))


def content_types(nslides):
    parts = [
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
        f"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
        "<Default Extension=\"png\" ContentType=\"image/png\"/>",
        "<Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>",
        "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>",
        "<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>",
        "<Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>",
        "<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>",
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
    ]
    for i in range(1, nslides + 1):
        parts.append(f"<Override PartName=\"/ppt/slides/slide{i}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>")
    parts.append("</Types>")
    return "".join(parts)


def root_rels():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<Relationships xmlns=\"{PKG_NS}\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>"
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
            "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
            "</Relationships>")


def core_props():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
            "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
            "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"
            f"<dc:title>{CONFIG.get('output', 'deck')}</dc:title>"
            "<dc:creator>Codex</dc:creator>"
            "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2026-08-05T00:00:00Z</dcterms:created>"
            "</cp:coreProperties>")


def app_props(nslides):
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
            "<Application>Codex</Application>"
            f"<Slides>{nslides}</Slides></Properties>")


def presentation_xml(nslides):
    ids = "".join(f"<p:sldId id=\"{256 + i}\" r:id=\"rId{i + 2}\"/>" for i in range(1, nslides + 1))
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<p:presentation xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\">"
            "<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>"
            f"<p:sldIdLst>{ids}</p:sldIdLst>"
            "<p:sldSz cx=\"12192000\" cy=\"6858000\" type=\"screen16x9\"/>"
            "<p:notesSz cx=\"6858000\" cy=\"9144000\"/>"
            "</p:presentation>")


def presentation_rels(nslides):
    parts = [f"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{PKG_NS}\">",
             "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>"]
    for i in range(1, nslides + 1):
        parts.append(f"<Relationship Id=\"rId{i + 2}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i}.xml\"/>")
    parts.append("</Relationships>")
    return "".join(parts)


def theme_xml():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<a:theme xmlns:a=\"{A_NS}\" name=\"Office\">"
            "<a:themeElements>"
            "<a:clrScheme name=\"Office\">"
            "<a:dk1><a:srgbClr val=\"000000\"/></a:dk1><a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1>"
            "<a:dk2><a:srgbClr val=\"1F4E79\"/></a:dk2><a:lt2><a:srgbClr val=\"D9E2F3\"/></a:lt2>"
            "<a:accent1><a:srgbClr val=\"1F4E79\"/></a:accent1><a:accent2><a:srgbClr val=\"2E75B6\"/></a:accent2>"
            "<a:accent3><a:srgbClr val=\"70AD47\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>"
            "<a:accent5><a:srgbClr val=\"ED7D31\"/></a:accent5><a:accent6><a:srgbClr val=\"A5A5A5\"/></a:accent6>"
            "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>"
            "</a:clrScheme>"
            "<a:fontScheme name=\"Office\">"
            "<a:majorFont><a:latin typeface=\"Times New Roman\"/><a:ea typeface=\"SimSun\"/><a:cs typeface=\"\"/></a:majorFont>"
            "<a:minorFont><a:latin typeface=\"Times New Roman\"/><a:ea typeface=\"SimSun\"/><a:cs typeface=\"\"/></a:minorFont>"
            "</a:fontScheme>"
            "<a:fmtScheme name=\"Office\">"
            "<a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>"
            "<a:lnStyleLst>"
            "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>"
            "<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>"
            "<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>"
            "</a:lnStyleLst>"
            "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>"
            "<a:effectStyle><a:effectLst/></a:effectStyle>"
            "<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
            "<a:bgFillStyleLst>"
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            "</a:bgFillStyleLst>"
            "</a:fmtScheme>"
            "</a:themeElements>"
            "<a:objectDefaults/><a:extraClrSchemeLst/>"
            "</a:theme>")


def master_xml():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<p:sldMaster xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\">"
            "<p:cSld><p:spTree>"
            "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
            "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>"
            "<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Background\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
            "<p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"12192000\" cy=\"6858000\"/></a:xfrm>"
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>"
            "<a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></p:spPr>"
            "<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"zh-CN\"/></a:p></p:txBody></p:sp>"
            "</p:spTree></p:cSld>"
            "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>"
            "<p:sldLayoutIdLst><p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/></p:sldLayoutIdLst>"
            "<p:txStyles>"
            "<p:titleStyle><a:lvl1pPr><a:defRPr sz=\"3200\"><a:latin typeface=\"Times New Roman\"/><a:ea typeface=\"SimSun\"/></a:defRPr></a:lvl1pPr></p:titleStyle>"
            "<p:bodyStyle><a:lvl1pPr><a:defRPr sz=\"2000\"><a:latin typeface=\"Times New Roman\"/><a:ea typeface=\"SimSun\"/></a:defRPr></a:lvl1pPr></p:bodyStyle>"
            "<p:otherStyle><a:lvl1pPr><a:defRPr sz=\"1200\"><a:latin typeface=\"Times New Roman\"/><a:ea typeface=\"SimSun\"/></a:defRPr></a:lvl1pPr></p:otherStyle>"
            "</p:txStyles>"
            "</p:sldMaster>")


def master_rels():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<Relationships xmlns=\"{PKG_NS}\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>"
            "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"../theme/theme1.xml\"/>"
            "</Relationships>")


def layout_xml():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<p:sldLayout xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\" type=\"blank\" preserve=\"1\">"
            "<p:cSld name=\"Blank\"><p:spTree>"
            "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
            "<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
            "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>"
            "</p:spTree></p:cSld>"
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>"
            "</p:sldLayout>")


def layout_rels():
    return ("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            f"<Relationships xmlns=\"{PKG_NS}\">"
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"../slideMasters/slideMaster1.xml\"/>"
            "</Relationships>")


if __name__ == "__main__":
    build()
