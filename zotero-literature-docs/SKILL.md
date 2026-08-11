---
name: zotero-literature-docs
description: Generate literature documents from a local Zotero library — curated Word (.docx) lists, matching PDFs, review documents, and illustrated PowerPoint decks with figures cited from the library. Use when the user asks to search or review their Zotero library and produce Word/PDF/PPT deliverables (e.g., "把检索到的文献整理成 Word 和 PDF", "按提纲生成 PPT", "做一份文献综述并出文档").
---

# Zotero Literature Docs

One pipeline turns Zotero library items into deliverables: fetch clean item data
from the Zotero local API → generate Word (python-docx) + PDF (fpdf2) reviews →
extract real figures from the library's PDFs → build an illustrated PPT on a
PowerPoint-generated skeleton. Every output is validated (re-open, garble scan,
PowerPoint open + PDF preview export).

## Prerequisites

- Zotero Desktop running with the local API on port 23119. Check with
  `python <skill>/scripts/zotero_docs.py status`.
- Python libraries next to the scripts in the project `work/` dir: copy `pylibs`
  (fpdf2, fontTools, PyMuPDF/fitz) and `pylibs2` (python-docx, lxml). See
  [references/method.md](references/method.md) for the full dependency list.
- System fonts: Deng.ttf / Dengb.ttf (PDF), Microsoft YaHei (Word/PPT).
- PowerPoint installed, used only for PPT validation (never for file conversion).

## Quick start (run from the project root)

```bash
# 1. One-time setup: copy pipeline scripts + skeleton into ./work/
python <skill>/scripts/zotero_docs.py setup

# 2. Search the library (read-only) to pick item keys
python work/zotero_docs.py search "anthocyanin"          # or the zotero skill's search

# 3. Fetch clean data
python work/zotero_docs.py fetch KEY1,KEY2,...

# 4. Read the papers (optional but recommended before writing a review)
python work/zotero_docs.py texts KEY1,KEY2,...

# 5. Write work/review_content.json (sections + [@KEY] citations), then:
python work/zotero_docs.py review work/review_content.json

# 6. Extract real figures (cropped from PDFs, not whole pages)
python work/zotero_docs.py figs KEY1,KEY2,...

# 7. Build the PPT from a deck config
python work/zotero_docs.py ppt work/deck_config.json
```

## Workflow details

1. **Search & select.** `zotero_docs.py search <query>` lists matching items
   (key, year, title) from the local API. The companion `zotero` skill also
   works (`zotero.py search`). Collect the item keys you want.
2. **Fetch data.** `zotero_docs.py fetch KEY1,KEY2,...` writes `work/data.json`.
   Stale/404 keys are skipped with a warning instead of crashing; dirty titles
   (leading numbering noise) are cleaned. Never use `export-bibtex` text as the
   data source (special characters get garbled).
3. **Read (optional).** `zotero_docs.py texts` dumps the first two pages + last
   page of each PDF to `work/paper_texts/<KEY>.txt` for abstract/intro/conclusion
   reading before writing the review.
4. **Review document.** Write `work/review_content.json` with sections and
   citation tokens like `[@ZOTERO_KEY]`, then run
   `python work/zotero_docs.py review work/review_content.json` (or
   `python work/make_review_docs.py [content.json]`). The script:
   - errors loudly if a cited `[@KEY]` is missing from `data.json`;
   - warns about fetched items that were never cited;
   - renders in-text citations in author-year style （Zhang et al.，2015）;
   - appends an alphabetically sorted reference list in Chinese journal format.
   Outputs: `outputs/<output>.docx` and `.pdf`, both garble-scanned.
   The PDF follows Chinese journal typography (SimSun/SimHei, sizes, line spacing,
   numbered headings, first-line indents). Optional `keywords`, `table`
   (caption + columns + rows) and `image` (path + caption) fields in the
   content JSON are rendered with the journal's caption style.
5. **Figures.** `python work/extract_figures.py KEY1,KEY2,...` crops embedded
   raster figures (real figures, not whole-page screenshots) to
   `work/ppt_images/<KEY>_f1.png` etc. Use `--whole-page` only when a paper has
   no extractable raster figures and you explicitly accept full-page renders.
   Figure names `<KEY>_fN.png` map to `{"type": "pic", "img": "<KEY>", "fig": N}`
   in the deck config.
6. **PPT.** Slide content lives in a JSON deck config (`work/deck_config.json`),
   not in Python code. Supported slide types:
   - `cover` (title / subtitle / note1 / note2)
   - `toc` (items)
   - `pic` (title, optional subtitle, `img` key, `fig` number; caption is
     auto-generated as "图片引自参考文献[n]：作者（年份）《标题》")
   - `text` (title, lines)
   - `refs` (count of references to list)
   `python work/assemble_pptx.py work/deck_config.json --out outputs/xxx.pptx`
   merges the slides into the PowerPoint skeleton. Any slide count is supported
   (extra parts are cloned, unused skeleton slides are dropped). See
   `scripts/deck_config.example.json`.
7. **Validate.**
   - Word: reopens with correct paragraph counts; garble scan clean.
   - PDF: text extraction contains expected titles; garble scan clean.
   - PPT: open in PowerPoint with a visible window (`WithWindow=$true`) and
     export a PDF preview. Windowless automation misreports "could not open the
     file" for decks that contain images — always validate with a window.

## Key Rules

- Data must come from the Zotero local API as JSON. Repair known GBK-style
  mojibake (汕→β, 汐→α, 汎→ζ, 坼→-) in `fetch_items.py`'s `normalize()`.
- Word: set the East Asian font explicitly (`rPr/rFonts` `eastAsia` =
  Microsoft YaHei).
- PDF: use fpdf2 with Deng fonts; call `set_x(l_margin)` after every
  `multi_cell`.
- PPT: build on the PowerPoint-created skeleton (`assets/_skeleton.pptx`), never
  a hand-written theme/master. Text alignment for center is `algn="ctr"` (not
  `"c"`). When adding PNG media, add a `<Default Extension="png" .../>` to
  `[Content_Types].xml`. Preserve each slide's original `.rels` bytes and append
  image relationships only.
- PPT figures are cropped real figures by default; do not render whole pages
  unless the user explicitly accepts full-page screenshots.
- In-text citations are author-year （Zhang et al.，2015）; consecutive citations
  merge into one parenthesis separated by semicolons. The reference list is
  alphabetical by first author in the format 作者. 年份. 标题. 期刊，卷(期)：页码.
  English titles use sentence case; gene names and Latin binomials are italic,
  naming authorities (e.g. Linden.) stay regular with a capital initial.
- Review PDFs follow Chinese journal typography (SimSun/SimHei/Times, numbered headings,
  first-line indents, three-line tables, page-break control); Word matches the
  PDF for fonts, page numbers and table pagination.
- In Word, gene symbols and Latin binomials are italic; protein/enzyme
  abbreviations are regular. Extend `ITALIC_WORDS` / `PLANT_GENERA` in
  `make_review_docs.py` when a new topic introduces new symbols.
- English mode: set `"lang": "en"` in `review_content.json` (and optionally in
  `deck_config.json`) to render English labels (Abstract / Keywords /
  References) in Word/PDF and English PPT captions ("Figure from reference [n]").
- Do not use Word COM for conversion (hangs on this machine). PowerPoint COM
  `AddPicture` is unstable; assemble pictures via XML instead.

## Resources

- `scripts/zotero_docs.py`: unified CLI (setup/status/search/fetch/texts/review/figs/ppt).
- `scripts/fetch_items.py`, `make_review_docs.py`, `make_docx_v2.py`,
  `make_pdf.py`, `make_papers_doc.py`, `extract_texts.py`,
  `extract_figures.py`, `build_pptx.py`, `assemble_pptx.py`.
- `scripts/deck_config.example.json`: deck config schema reference.
- `references/method.md`: the preserved full method (steps, dependencies,
  pitfalls, changelog).
- `assets/_skeleton.pptx`: PowerPoint-generated blank skeleton used as the PPT
  host package.
