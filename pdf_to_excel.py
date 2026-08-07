"""Convert one or more PDFs into a single Excel workbook.

For each PDF it pulls out the tables (ruled/lattice tables via pdfplumber) and
writes them into Excel columns, one sheet per PDF. Tables that continue across
pages are stitched together and a repeated header row is dropped. When a PDF has
no detectable tables (e.g. a scanned document), it falls back to plain text —
one row per line. Purely numeric cells are written as numbers so Excel can sum
them. `convert(...)` is the entry point the UI calls.
"""

import re
from collections import OrderedDict
from pathlib import Path

_INVALID_SHEET = re.compile(r"[\[\]:*?/\\]")
_NUMERIC = re.compile(r"-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+(?:\.\d+)?")


def _clean(cell):
    """A table cell as a trimmed single-line string ('' for empty/None)."""
    if cell is None:
        return ""
    return re.sub(r"\s+", " ", str(cell)).strip()


def _merge_tables(tables):
    """Group tables by column count and concatenate each group into one grid.

    A table split across pages comes back as several same-width tables; joining
    them (and skipping a repeated header row) rebuilds the original table.
    """
    groups = OrderedDict()
    for table in tables:
        width = max((len(r) for r in table), default=0)
        groups.setdefault(width, []).append(table)

    grids = []
    for width, group in groups.items():
        merged, header = [], None
        for table in group:
            for row in table:
                row = row + [""] * (width - len(row))
                if header is None:
                    header, merged = row, [row]
                elif row != header:  # skip the header repeated on later pages
                    merged.append(row)
        grids.append(merged)
    return grids


def extract_pdf_grids(path):
    """Return a list of grids (each grid = list of string rows) for one PDF."""
    import pdfplumber

    tables, text_lines = [], []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            if page_tables:
                for table in page_tables:
                    grid = [[_clean(c) for c in row] for row in table]
                    grid = [r for r in grid if any(r)]  # drop blank rows
                    if grid:
                        tables.append(grid)
            else:
                for line in (page.extract_text() or "").split("\n"):
                    if line.strip():
                        text_lines.append([line.strip()])

    if tables:
        return _merge_tables(tables)
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
