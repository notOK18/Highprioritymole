"""Convert one or more PDFs into a single Excel workbook.

For each PDF it pulls out the tables (ruled/lattice tables via pdfplumber) and
writes them into Excel columns, one sheet per PDF. Tables that continue across
pages are stitched together and a repeated header row is dropped. When a PDF has
no detectable tables (e.g. a scanned document), it falls back to plain text —
one row per line. Purely numeric cells are written as numbers so Excel can sum
them. `convert(...)` is the entry point the UI calls.
"""

import re
from pathlib import Path

_INVALID_SHEET = re.compile(r"[\[\]:*?/\\]")
_NUMERIC = re.compile(r"-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+(?:\.\d+)?")


def _clean(cell):
    """A table cell as a trimmed single-line string ('' for empty/None)."""
    if cell is None:
        return ""
    return re.sub(r"\s+", " ", str(cell)).strip()


def _clean_grid(table):
    """Clean every cell and drop fully blank rows."""
    grid = [[_clean(c) for c in row] for row in table]
    return [r for r in grid if any(r)]


def _drop_empty_columns(grid):
    """Remove columns that are empty in every row."""
    if not grid:
        return grid
    width = max(len(r) for r in grid)
    grid = [r + [""] * (width - len(r)) for r in grid]
    keep = [c for c in range(width) if any(row[c].strip() for row in grid)]
    return [[row[c] for c in keep] for row in grid]


def _page_tables(page):
    """Tables on one page: ruled tables first, else a text-aligned (borderless) one."""
    # 1) ruled/lattice tables (real cell borders)
    grids = [g for g in (_clean_grid(t) for t in page.extract_tables())
             if g and max(len(r) for r in g) > 1]
    if grids:
        return grids

    # 2) borderless table inferred from how the words line up. A column boundary
    # is only kept where enough rows align there (min_words_vertical), so the
    # description column is not split at every word gap. Tune it to the row count.
    n_rows = sum(1 for line in (page.extract_text() or "").split("\n") if line.strip())
    if n_rows < 2:
        return []
    mwv = max(3, min(8, round(n_rows * 0.4)))
    settings = {"vertical_strategy": "text", "horizontal_strategy": "text",
                "min_words_vertical": mwv, "min_words_horizontal": 1}
    grids = []
    for table in page.extract_tables(settings):
        grid = _drop_empty_columns(_clean_grid(table))
        if grid and len(grid) > 1 and max(len(r) for r in grid) > 1:
            grids.append(grid)
    return grids


def _merge_grids(grids):
    """Concatenate page grids into one (pad to max width, drop a repeated header)."""
    width = max((max((len(r) for r in g), default=0) for g in grids), default=0)
    merged, header = [], None
    for grid in grids:
        for row in grid:
            row = row + [""] * (width - len(row))
            if header is None:
                header, merged = row, [row]
            elif row != header:  # skip the header repeated on later pages
                merged.append(row)
    return merged


def extract_pdf_grids(path):
    """Return a list of grids (usually one) of string rows for one PDF."""
    import pdfplumber

    tables, text_lines = [], []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            page_grids = _page_tables(page)
            if page_grids:
                tables.extend(page_grids)
            else:
                for line in (page.extract_text() or "").split("\n"):
                    if line.strip():
                        text_lines.append([line.strip()])

    if tables:
        return [_drop_empty_columns(_merge_grids(tables))]
    if text_lines:
        return [text_lines]
    return []


def _maybe_number(value):
    """Convert a purely numeric string to a float so Excel treats it as a number."""
    if isinstance(value, str) and _NUMERIC.fullmatch(value.strip()):
        try:
            return float(value.replace(",", ""))
        except ValueError:
            return value
    return value


def _unique_sheet_name(name, used):
    """Sanitize to a valid, unique, <=31-char Excel sheet name."""
    name = _INVALID_SHEET.sub(" ", name).strip() or "Sheet"
    name = name[:31]
    base, i = name, 2
    while name.lower() in used:
        suffix = f" {i}"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name.lower())
    return name


def _write_grid(ws, book, rows):
    """Write a grid to a worksheet: bold header row, numeric cells as numbers."""
    bold = book.add_format({"bold": True})
    widths = {}
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            if r == 0:
                ws.write(r, c, value, bold)
            else:
                ws.write(r, c, _maybe_number(value))
            widths[c] = max(widths.get(c, 0), len(str(value)))
    for c, w in widths.items():
        ws.set_column(c, c, min(60, w + 2))


def convert(pdf_paths, out_path):
    """Convert every PDF in `pdf_paths` into one workbook. Returns (out_path, n_sheets)."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter

    book = xlsxwriter.Workbook(str(out_path), {"constant_memory": True})
    used, sheets = set(), 0
    for pdf in pdf_paths:
        stem = Path(pdf).stem
        grids = extract_pdf_grids(pdf)
        if not grids:
            ws = book.add_worksheet(_unique_sheet_name(f"{stem} (empty)", used))
            _write_grid(ws, book, [["(no extractable content)"]])
            sheets += 1
            continue
        for idx, grid in enumerate(grids):
            base = stem if len(grids) == 1 else f"{stem} ({idx + 1})"
            ws = book.add_worksheet(_unique_sheet_name(base, used))
            _write_grid(ws, book, grid)
            sheets += 1
    if sheets == 0:  # a workbook needs at least one sheet
        book.add_worksheet("empty")
    book.close()
    return out_path, sheets


def main():
    import sys
    if len(sys.argv) < 2:
        print("usage: python pdf_to_excel.py <file1.pdf> [file2.pdf ...]")
        return
    pdfs = sys.argv[1:]
    out = Path(pdfs[0]).with_suffix(".xlsx")
    out_path, sheets = convert(pdfs, out)
    print(f"Wrote {sheets} sheet(s) -> {out_path}")


if __name__ == "__main__":
    main()
