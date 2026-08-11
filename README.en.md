# zotero-literature-docs

> English | [中文](README.md)

A Codex Skill for your **local Zotero library**: **search → analyze → review → one-click Word + PDF + PPT deliverables**.

It reads clean data from the Zotero local API, produces review documents (Word + PDF) formatted to standard Chinese academic journal conventions, crops **real figures** from the paper PDFs, and assembles an illustrated PPT with "Figure from reference [n]" captions. Every output is auto-validated (re-open check + garble scan + PowerPoint open with PDF preview export).

## Features

- **Local search**: `search` lists matching items (key / year / title) from your Zotero library
- **Clean data**: pulls JSON from the local API (`fetch`), auto-cleans dirty titles, skips stale items; never uses `export-bibtex` text (special characters get garbled)
- **Fast reading**: `texts` dumps the first two pages + last page of each PDF, so you can read abstract / intro / conclusion before writing
- **Review generation**: write sections with `[@ZoteroKey]` citations in `work/review_content.json` → one command generates Word + PDF
  - Author-year in-text citations (Zhang et al., 2015), merged consecutive citations, alphabetically sorted reference list
  - Cited-but-missing keys raise an error; fetched-but-uncited items raise a warning — no "ghost citations"
  - Chinese typography: SimSun/SimHei, numbered headings, first-line indents, three-line tables; gene names and Latin binomials in italics, proteins in regular type
- **Real figures**: `figs` crops embedded raster figures by their bounding boxes (not whole-page screenshots) → `work/ppt_images/`
- **Presentation deck**: `ppt` builds on a PowerPoint-generated skeleton, driven by a JSON config (cover / toc / pic / text / refs slide types), any slide count
- **Validation**: Word / PDF re-open check + garble scan; PPT opened in PowerPoint with a visible window and exported to a PDF preview
- **Bilingual output**: set `"lang": "en"` in the config to generate English deliverables (Abstract / Keywords / References labels and PPT captions switch to English)

## Requirements

- Windows (verified) ｜ Python 3.10+ (developed on 3.14)
- Zotero Desktop running, local API on port 23119
- Fonts: SimSun (simsun.ttc), SimHei (simhei.ttf), optional DengXian (Deng.ttf / Dengb.ttf) for PDF; Microsoft YaHei for Word / PPT
- PowerPoint (only for PPT validation, never for file conversion)
- Dependencies: fpdf2, fontTools, PyMuPDF/fitz, python-docx, lxml (the Skill places a usable offline dependency directory under `work/`)

## Quick start

Run from the project root (the first run copies the pipeline scripts and skeleton into `work/`):

```bash
# 1. One-time setup
python <skill>/scripts/zotero_docs.py setup

# 2. Search the local Zotero library (read-only)
python work/zotero_docs.py search "anthocyanin"

# 3. Fetch clean data (comma-separated keys)
python work/zotero_docs.py fetch KEY1,KEY2,...

# 4. (Optional) dump first/last pages of each PDF to read before writing
python work/zotero_docs.py texts KEY1,KEY2,...

# 5. Write work/review_content.json (sections + [@KEY] citations), then generate the review
python work/zotero_docs.py review work/review_content.json

# 6. Crop real figures from the paper PDFs
python work/zotero_docs.py figs KEY1,KEY2,...

# 7. Write work/deck_config.json, then build the PPT
python work/zotero_docs.py ppt work/deck_config.json
```

Outputs land in `outputs/`: `<name>.docx`, `<name>.pdf`, `<name>.pptx`.

## Repository layout

```text
zotero-literature-docs/
├── SKILL.md                        # Skill definition (entry point for Codex)
├── agents/openai.yaml              # trigger & behavior config
├── references/method.md            # full method, dependencies and pitfalls (Chinese)
├── references/method.en.md         # full method, dependencies and pitfalls (English)
├── scripts/                        # pipeline scripts (unified CLI: zotero_docs.py)
│   ├── zotero_docs.py              # setup/status/search/fetch/texts/review/figs/ppt
│   ├── fetch_items.py              # data fetch & cleaning
│   ├── make_review_docs.py         # review Word + PDF generation
│   ├── extract_figures.py          # real figure cropping from PDFs
│   ├── build_pptx.py / assemble_pptx.py  # PPT generation & assembly
│   └── deck_config.example.json    # deck config example
└── assets/_skeleton.pptx           # PowerPoint-generated blank skeleton
```

## Install as a Codex skill

- Option 1: copy the `zotero-literature-docs` directory to `~/.codex/skills/` and restart Codex
- Option 2: install from this repository via the Codex skill installer
- Usage: describe the task in a new chat and the skill triggers automatically, e.g. "Compile the papers on the test topic in my Zotero library into a review and produce Word / PDF / PPT"

## Notes

- Data must come from the Zotero local API as JSON; do not use `export-bibtex` text
- When a new topic introduces new gene names / plant genera, extend `ITALIC_WORDS` / `PLANT_GENERA` in `scripts/make_review_docs.py`
- Do not use Word COM for document conversion (unstable on this machine); PPT validation must use `WithWindow=$true`
- Figures are cropped real figures by default; use `--whole-page` only when a paper has no embedded raster figures and the user explicitly agrees
- For English deliverables, set `"lang": "en"` in `review_content.json` and `deck_config.json`

## License

MIT License — see [LICENSE](LICENSE).
