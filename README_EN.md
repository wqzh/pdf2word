[中文](https://github.com/wqzh/pdf2word/blob/main/README.md) | [English](https://github.com/wqzh/pdf2word/blob/main/README_EN.md)

# pdf2word-pro

A command-line tool that converts PDF into Word (`.docx`) while trying hard to **preserve the original layout**: paragraphs, fonts and sizes, tables (including merged cells and shading), images (including side-by-side figures and their captions), headers and footers, page breaks and page sizes.

A pure-Python implementation that only relies on `pdfplumber` for low-level PDF parsing and `python-docx` for assembling the Word document. After installation the command is `pdf2word-pro`; `python -m pdf2word_pro` also works, and the legacy `python pdf2docx.py` entry point is kept in the repository.

---

## 1. Features

| Area | What it does |
| --- | --- |
| Text | Rebuilds text as character → line → paragraph, keeping font, size, bold/italic and colour; no extra spaces for CJK, de-hyphenation for English |
| Paragraphs | Decides soft vs. hard line breaks by "does the remaining space on the previous line fit the first word of the next line"; preserves alignment, first-line indent, and maps vertical spacing to Word space-before/space-after |
| Headings | Short paragraphs noticeably larger than body text are mapped to Word Heading 1/2/3 styles (easy to restyle afterwards) |
| Columns | Detects two-column layouts by locating the gutter and emits content column by column |
| Tables | Rebuilds the grid from PDF rules, including merged cells and cell shading; cell text is scaled to the original size; tables continuing across pages are stitched back together with duplicate header rows removed |
| Images | Prefers the original image stream (JPEG as-is, lossless RGB/grey); falls back to region rendering; keeps original size, indentation and centring |
| Side-by-side figures | Figures sharing a horizontal band (e.g. an a) b) c) d) grid, or two figures side by side) are rebuilt as a borderless table at their original column positions — never split or reordered |
| Inset images | Legends / magnified details overlaid on a larger figure are merged into the outer figure's rendered region instead of being mistaken for a second column |
| Vector graphics | Charts made of curves and filled shapes are clustered, rendered as a single image, and their internal text is removed to avoid duplication |
| Headers / footers | Repeating cross-page lines become Word headers/footers, with page numbers emitted as `PAGE` / `NUMPAGES` fields; an "without header/footer" copy can be exported as well |
| Pages | Page margins are derived from the original page size and content extents; mixed portrait/landscape pages produce separate sections |
| Robustness | Encrypted or damaged PDFs are rewritten through `pypdf` and retried; a page that fails to parse is rasterised as a fallback; illegal XML characters are stripped |

---

## 2. Requirements

- Python **3.9+** (developed and verified on Python 3.11.9 / Windows 10)
- Dependencies (see `requirements.txt`):

| Package | Purpose | Verified version |
| --- | --- | --- |
| `pdfplumber` | Read PDF pages, characters, tables, rules and images | 0.11.4 |
| `python-docx` | Build the `.docx` file | 1.2.0 |
| `pillow` | Image decoding / PNG encoding | 13.0.0 |
| `pypdf` | Fallback rewrite of damaged or encrypted PDFs | 7.1.0 |

`concurrent.futures`, `argparse`, `json`, `logging`, etc. are part of the standard library and need no extra installation.

### Install

```shell
pip install pdf2word-pro
```

From source (development):

```shell
pip install -r requirements.txt
pip install -e .
```

---

## 3. Quick start

```shell
pdf2word-pro input.pdf
```

Without `-o`, the output is a `.docx` with the same name next to the input PDF, and its final path is printed to stdout.

```shell
pdf2word-pro input.pdf -o output.docx
```

---

## 4. Command-line options

```text
usage: pdf2word-pro [-h] [-o OUTPUT] [--pages PAGES] [--password ***]
                   [--keep-page-breaks] [--no-columns] [--no-figures]
                   [--no-headings] [--no-header-footer]
                   [--remove-page-head-foot] [--no-fonts] [--dpi DPI]
                   [--cache-page] [--cache-dir CACHE_DIR] [--jobs JOBS] [-v]
                   pdf
```

| Option | Default | Description |
| --- | --- | --- |
| `pdf` | required | Input PDF path (positional) |
| `-o`, `--output` | same name as input, `.docx` | Output `.docx` path; missing directories are created |
| `--pages` | all pages | Convert only the given pages, 1-based. Accepts `1-3,5,8-` (`8-` means page 8 to the end); out-of-range values are clamped |
| `--password` | none | Password for an encrypted PDF |
| `--keep-page-breaks` | off | Insert a page break wherever the original PDF breaks a page (by default content reflows) |
| `--no-columns` | off | Disable two-column detection and read full-width lines |
| `--no-figures` | off | Do not render vector graphics as images (text and bitmaps only) |
| `--no-headings` | off | Do not apply Word heading styles; everything becomes body paragraphs |
| `--no-header-footer` | off | Do not extract headers/footers (their text stays in the body) |
| `--remove-page-head-foot` | off | Besides the full document, **also** save `<name>-no-head-foot.docx` without headers/footers/page numbers, and write the per-page detection result to the console and `<name>-page-head-foot.json` |
| `--no-fonts` | off | Do not keep original font names (use the document default) |
| `--dpi` | `200` | DPI used when rasterising regions (including the whole-page fallback); higher is sharper and larger |
| `--cache-page` | off | Dump each page's parse result (JSON + plain text + images) to disk for troubleshooting |
| `--cache-dir` | `_pdf2word_pages` next to the output | Directory for the per-page cache; created if missing |
| `--jobs` | `1` | Number of parsing processes; when `>1` and more than 8 pages, pages are parsed in parallel (pages are independent, so long documents get several times faster) |
| `-v`, `--verbose` | off | DEBUG-level logging, and re-raise the full traceback on failure |

Exit codes: `0` success, `1` conversion failed (the reason is logged; `-v` shows the full stack).

---

## 5. Output files

Assuming `-o output.docx`:

| File | Produced when | Description |
| --- | --- | --- |
| `output.docx` | always | Full document, with headers/footers |
| `output-no-head-foot.docx` | with `--remove-page-head-foot` | Same document with headers/footers/page numbers removed |
| `output-page-head-foot.json` | with `--remove-page-head-foot` | Per-page header/footer/page-number detection result |
| `_pdf2word_pages\page_0001.json` … | with `--cache-page` | Structured parse result of that page (lines/characters/tables/image boxes, UTF-8) |
| `_pdf2word_pages\page_0001.txt` | with `--cache-page` | Plain-text view of that page (line text + table cell text) |
| `_pdf2word_pages\page_0001_*.jpg/png` | with `--cache-page` | Images / rendered vector figures extracted from that page |
| `_pdf2word_pages\manifest.json` | with `--cache-page` | Summary index of all pages (page number, block count, files, timing) |

---

## 6. Examples

### 6.1 Convert the whole document

```shell
pdf2word-pro input.pdf -o output.docx
```

### 6.2 Convert only a few pages (most common for layout regression)

```shell
pdf2word-pro input.pdf -o sample.docx --pages 12-14,20
```

### 6.3 Speed up a long document with multiple processes

```shell
pdf2word-pro input.pdf -o output-job8.docx --jobs 8
```

### 6.4 Remove headers and footers (keeping the full version)

```shell
pdf2word-pro input.pdf -o output.docx --remove-page-head-foot
# produces: output.docx, output-no-head-foot.docx, output-page-head-foot.json
```

### 6.5 Encrypted PDF

```shell
pdf2word-pro encrypted.pdf -o decrypted.docx --password "***"
```

### 6.6 Keep original page breaks and disable column detection (layout inspection)

```shell
pdf2word-pro input.pdf -o output-layout.docx `
  --keep-page-breaks --no-columns --cache-page --cache-dir _dbg -v
```

### 6.7 Use it as a library

```python
from pdf2word_pro import Config, pdf_to_docx

cfg = Config(pages={45, 46, 47}, image_dpi=300, jobs=4)   # pages are 0-based
out = pdf_to_docx("input.pdf", "output.docx", cfg)
print(out)
```

`Config` fields mirror the command-line options one-to-one (see the comments next to the `Config` definition): `keep_page_breaks`, `detect_columns`, `render_figures`, `image_dpi`, `jobs`, `cache_pages`, `cache_dir`, `remove_page_head_foot`, and so on.

---

## 7. How it works

1. **Per-page parsing** (`parse_page`, pdfplumber)
   - Tables: `find_tables()` → grid + merged cells + shading + text/images inside cells
   - Images: original image stream first, region rendering as a fallback; inset images (legends/details) are merged into the outer figure
   - Vector graphics: curves and filled shapes are clustered, rendered as one image, and their internal text is removed
   - Text: characters → lines → paragraphs (font/size/bold/italic/colour); two-column detection
2. **Global passes**
   - Headers/footers: repeating cross-page lines are recognised and written as Word headers/footers, page numbers as fields
   - Line-join: English de-hyphenation, no spaces for CJK
   - Cross-page merging: tables (column-aligned, duplicate header rows removed) and paragraphs (last line full and next page not indented)
3. **Writing the docx** (`write_docx`)
   - Page size/margins, alignment, indentation, line/paragraph spacing, table column widths and row heights, image size/alignment
   - Side-by-side figures become a borderless table; a figure row is never split across pages and its last row stays with the caption
4. **Robustness**: encrypted/damaged PDFs, whole-page fallback on parse failure, illegal XML character filtering

---

## 8. How side-by-side figures are laid out

This is the most error-prone part of any PDF-to-Word conversion. Key decisions:

- **Row detection**: only figures inside the same horizontal band form a row; adjacent rows with different column counts are never merged, so figures cannot be mixed up.
- **One figure per column**: each figure gets its own column, and neighbouring figures are separated at the **midpoint** between their two edges. No horizontal cell merging is performed, so even if two figures slightly overlap in the PDF, they never end up squeezed into one merged cell.
- **Row height and pagination**: table rows are set to `cantSplit` (a row never breaks across pages) and the last row's paragraphs use `keep_with_next`, so a figure row stays on the same page as the caption that follows it (`a) … b) …` or `Figure x.x-x`).
- **Inset figures**: containment (area ratio < 90 % and fully enclosed) identifies "legend / magnified detail" images, which are merged into the outer figure's rendered region rather than treated as a second side-by-side column.
- **Alignment**: for centred rows the table indent follows the original table's left edge, and scaling is computed proportionally from the available width, so nothing is stretched.

---

## 9. Known limitations and troubleshooting

- **Scanned PDFs**: there is no text layer, so each page can only become an image (the log says "no content extracted (possibly a scan — run OCR first)"). Run OCR before converting.
- **Fonts missing ToUnicode**: you will see a "N glyphs could not be mapped to Unicode" warning; those characters can only be dropped or replaced.
- **Multiple images side by side inside a table cell**: currently emitted in vertical order; per-cell side-by-side layout is not implemented yet.
- **Figures combining a bitmap base with vector overlays**: if the base is a bitmap and the coloured areas are vector-drawn, taking the image stream may leave only the base; compare with `--no-figures`, or inspect the page with `--cache-page`.
- **Slow conversion**: add `--jobs 8`; to find a stuck page, isolate it with `--pages` and add `-v` to see stage timings.
- **Images too large / too small**: adjust `--dpi` (default 200). Higher DPI is sharper and produces a larger `.docx`.

---

## 10. Repository layout

```text
pyproject.toml           packaging config (PEP 621 / setuptools)
LICENSE                  MIT license
requirements.txt         dependency list
README.md                Chinese documentation
README_EN.md             this file
pdf2docx.py              legacy entry point
src/pdf2word_pro/        package source
  core.py                converter implementation
  cli.py                 command-line entry
  __main__.py            python -m entry
  __init__.py            public API
```
