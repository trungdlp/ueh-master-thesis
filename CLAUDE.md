# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

UEH MBA thesis authored in Markdown, compiled to a UEH-compliant Word `.docx` via Pandoc + python-docx post-processing. The thesis structure follows the official **Hướng dẫn Cấu trúc nội dung Đề án Thạc sĩ MBA** from Khoa Quản trị, Trường Kinh doanh, UEH (2026): Chương 1–5 + Kết luận + Tài liệu tham khảo.

## Build commands

```bash
make            # → outputs/luanvan.docx
make open       # build + open in Word
make figures    # regenerate sample figures via matplotlib
make reference  # rebuild outputs/reference-ueh.docx style template
make clean      # remove outputs/
```

First-time setup:
```bash
python3 -m venv venv
./venv/bin/pip install python-docx matplotlib
```

After opening in Word, **press F9** on Mục lục / Danh mục hình / Danh mục bảng pages to refresh the field-based listings.

## Pipeline architecture

The build is a 3-stage pipeline. Understanding the role of each stage is essential — none of them alone produces UEH-compliant output.

1. **`build_reference.py` → `outputs/reference-ueh.docx`** — Generates a Word file whose *styles* (not content) encode UEH formatting: Times New Roman 13pt black, line spacing 1.5, margins 3/3/3.5/2 cm, Heading 1 = 14pt bold UPPERCASE centered, Heading 2/3 variants, a `Table` style with full black borders + gray header row, page-number footer. Pandoc reads this file's styles and applies them to its output.

2. **`pandoc … --reference-doc=… --citeproc --csl=apa.csl`** — Concatenates `metadata.yaml` + `chapters/*.md` (sorted lexically — that's why chapters are prefixed `00-`, `01-`, …) into a single `.docx`. Citations (`[@key]`) are resolved against `references.bib` using APA style from `apa.csl`. The reference-doc supplies styling; pandoc supplies content and structure.

3. **`post_process.py outputs/luanvan.docx`** — Pandoc generates plain-text captions in `Image Caption` / `Table Caption` styles. This script walks document body in order, tracks the current chapter number from each `Heading 1` (parses `CHƯƠNG X` text), strips any pre-existing `Hình X.Y:` / `Bảng X.Y:` prefix, then prepends `Hình {chapter}.{SEQ Hình \s 1}: ` (bold) — where the SEQ field auto-resets per chapter when Word recomputes fields (F9). Caption paragraph is also centered.

Editing chapter content rarely needs touching the pipeline. Adding new caption-numbered elements (figures/tables) works automatically. Changing UEH style rules requires editing `build_reference.py` and `make reference`.

## Key conventions

- **Chapter file naming**: `chapters/NN-name.md`, prefix is sort-key only. `make` uses `$(sort $(wildcard chapters/*.md))`.
- **Headings**: `# CHƯƠNG N: ...` is parsed by `post_process.py` to extract chapter number `N` for caption prefixes. Don't reformat without updating `CHAPTER_RE` in `post_process.py`.
- **Pandoc captions**: use `: caption text` *after* a table for table captions; figure captions are the alt text in `![caption](path)`. Pandoc styles these as `Table Caption` / `Image Caption` so the TOC fields and post-processor find them.
- **TOC / List of Figures / List of Tables**: implemented as raw OOXML field codes in `chapters/04-muc-luc.md` (via pandoc's ` ```{=openxml} ` raw block). They only render real content after F9 in Word.
- **Bibliography rendering**: `chapters/11-tai-lieu-tham-khao.md` contains a `::: {#refs} :::` div — citeproc injects the bibliography into that div.
- **`luanvan.docx` and `reference-ueh.docx`** are build artifacts living in `outputs/` and are gitignored. Never commit them.

## Reference files

- `docs/UEH-MBA-huong-dan-cau-truc-de-an.pdf` — authoritative UEH MBA thesis structure spec from Khoa Quản trị, Trường Kinh doanh, UEH (2026). Consult before changing chapter outline.
