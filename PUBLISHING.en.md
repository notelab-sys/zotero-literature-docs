# Publishing Notes

Current status: the repository is **Private**. This document is the preparation material for a future public release.

## Version plan

### v0.1.0 (proposed initial release)

- Positioning: search → analyze → review a **local Zotero library** and generate **Word + PDF + PPT deliverables** in one pipeline
- Differentiators: one pipeline producing all three deliverables; PPT uses **real figures cropped from the paper PDFs** with "Figure from reference [n]" captions; all outputs are auto-validated
- Target users: Chinese-speaking researchers (plant science / horticulture), with output formatted to Chinese horticulture journal (园艺学报) conventions
- Compatibility: Windows + Codex (desktop / CLI); Python 3.10+; Zotero local API (port 23119)

## Pre-release checklist

1. **Visibility**: repository Settings → make Public; re-check that no sensitive information is present (already scanned — only system font paths)
2. **License**: confirm MIT (see `LICENSE`, copyright holder konjac2027); replace if another license is preferred
3. **README review**: features, quick start, requirements and layout match the actual content (Chinese `README.md` + English `README.en.md`)
4. **Script self-check**: `zotero_docs.py status` passes on a clean machine; dependency notes are complete (fpdf2 / fontTools / PyMuPDF / python-docx / lxml)
5. **Version tag**: `git tag v0.1.0` and push; optionally create a GitHub Release (with sample Word/PDF/PPT outputs)
6. **Optional**: publish to a Codex skill / plugin marketplace for one-click installation
7. **Record**: after release, register the release status and repo link in the workspace global logs / console

## How users install

- Option 1 (recommended): Codex skill installer with this repository path
- Option 2: manually copy the `zotero-literature-docs` directory to `~/.codex/skills/` and restart Codex
- Usage: describe the task in a new chat, e.g. "Compile the papers about anthocyanins in my Zotero library into a review and produce Word / PDF / PPT"

## Maintenance conventions

- After every change to the Skill content (SKILL.md / scripts / references), sync it to this repository
- For behavior changes: update the READMEs and bump the version; append pitfalls to `references/method.md`
- Prefer turning delivery logic improvements into reusable scripts instead of one-off manual processing
