"""Enrich supplier-common molecules with data from the Forecast folder.

Takes the list of "supplier common" molecules (the Common sheet produced by the
Supplier Common Molecules app) and every file in a Forecast folder, and for each
common molecule pulls the data it has in each source:

  * MIDAS quarterly sales  -> Units for 2023, 2024, 2025
  * Hospital consumption    -> Lebanon consumption (40%) and total (40%)
  * Tender PDF              -> minimum and maximum quantity per product

Molecules are matched on a "smart" base name (strength/form and any parenthetical
brand are ignored), consistent with the Supplier Common Molecules tool. The
output workbook has a Summary sheet (one aggregated row per molecule) plus a
detail sheet per source listing every matching product/dose row.
"""

import re
from pathlib import Path

import pandas as pd

import compare_molecules as cm

# ---------------------------------------------------------------------------
# Matching key
# ---------------------------------------------------------------------------
def base_key(text) -> str:
    """Smart base name: drop any parenthetical brand, then strength/form.

    'Nilotinib (Tasisna) 150 mA' -> 'nilotinib';  'NILOTINIB 200 mg' -> 'nilotinib'.
    """
    without_brand = re.sub(r"\([^)]*\)", " ", str(text))
    return cm._base(without_brand)


def _num(value):
    """Parse a possibly comma/space/currency-formatted value to float, or None."""
    if value is None:
        return None
    s = re.sub(r"[^\d.]", "", str(value).replace(",", ""))
    if s in ("", "."):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _ri(value):
    """Round to a whole number (consumption figures are reported as integers)."""
    return None if value is None else int(round(value))


# ---------------------------------------------------------------------------
# Source detection
# ---------------------------------------------------------------------------
def _sheet_matching(path, needles):
    """Return the first sheet name whose columns contain all `needles`, else None."""
    for sheet in cm.list_sheets(path):
        try:
            df = cm.extract_molecule_table(path, sheet_name=sheet)
        except Exception:
            continue
        cols = " | ".join(str(c).lower() for c in df.columns)
        if all(n in cols for n in needles):
            return sheet
    return None


def classify_forecast_file(path):
    """Return (kind, sheet) where kind is 'midas' | 'hospital' | 'tender' | None."""
    path = Path(path)
    if path.suffix.lower() == ".pdf":
        return "tender", None
    if path.suffix.lower() not in (".xls", ".xlsx"):
        return None, None
    hospital_sheet = _sheet_matching(path, ["lebanon consumption", "40%"])
    if hospital_sheet:
        return "hospital", hospital_sheet
    midas_sheet = _sheet_matching(path, ["molecule", "2023", "2024", "2025"])
    if midas_sheet:
        return "midas", midas_sheet
    return None, None


# ---------------------------------------------------------------------------
# Per-source extraction -> {base_key: [detail rows]}
# ---------------------------------------------------------------------------
def _col(df, *needles):
    """First column whose lowercased name contains every needle in one group."""
    for group in needles:
        for c in df.columns:
            name = str(c).lower()
            if all(n in name for n in group):
                return c
    return None


def extract_midas(path, sheet):
    df = cm.extract_molecule_table(path, sheet_name=sheet)
    mol = _col(df, ("molecule",)) or df.columns[0]
    c23 = _col(df, ("2023", "unit"), ("2023",))
    c24 = _col(df, ("2024", "unit"), ("2024",))
    c25 = _col(df, ("2025", "unit"), ("2025",))
    product = _col(df, ("international", "product"), ("product",))
    strength = _col(df, ("international", "strength"), ("strength",))
    records = {}
    for _, row in df.iterrows():
        key = base_key(row[mol])
        if not key:
            continue
        records.setdefault(key, []).append({
            "Molecule": row[mol],
            "Product": row[product] if product else "",
            "Strength": row[strength] if strength else "",
            "2023 Units": _num(row[c23]) if c23 else None,
            "2024 Units": _num(row[c24]) if c24 else None,
            "2025 Units": _num(row[c25]) if c25 else None,
        })
    return records


def extract_hospital(path, sheet):
    df = cm.extract_molecule_table(path, sheet_name=sheet)
    mol = _col(df, ("mapped", "molecule"), ("molecule",)) or df.columns[0]
    dose = _col(df, ("dose",))
    leb = _col(df, ("lebanon consumption", "40%"), ("lebanon", "40%"))
    tot = _col(df, ("total", "40%"), ("total", "consumption"))
    records = {}
    for _, row in df.iterrows():
        key = base_key(row[mol])
        if not key:
            continue
        records.setdefault(key, []).append({
            "Molecule": row[mol],
            "Dose/Strength": row[dose] if dose else "",
            "Lebanon Consumption 40%": _ri(_num(row[leb])) if leb else None,
            "Total Lebanon Consumption 40%": _ri(_num(row[tot])) if tot else None,
        })
    return records


_NUM_TOKEN = re.compile(r"\d[\d,.]*")
_ROW_START = re.compile(r"^([A-Za-z]{0,3}\d+)\b\s+(.*)$")


def extract_tender(path, sheet=None):
    """Parse the scanned tender PDF into {base_key: [ {Nb, Composition, Min, Max} ]}.

    Each data line is 'Nb Composition Unit MinQty MaxQty Group [Obs]'; the min and
    max are the last two numeric tokens on the line.
    """
    import pdfplumber
    records = {}
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            for line in (page.extract_text() or "").split("\n"):
                line = line.strip()
                m = _ROW_START.match(line)
                if not m:
                    continue
                nb, rest = m.group(1), m.group(2)
                nums = list(_NUM_TOKEN.finditer(rest))
                if len(nums) < 2:
                    continue
                min_tok, max_tok = nums[-2], nums[-1]
                composition = rest[:min_tok.start()].strip()
                key = base_key(composition)
                if not key:
                    continue
                records.setdefault(key, []).append({
                    "Nb": nb,
                    "Composition": composition,
                    "Min Qty": _num(min_tok.group()),
                    "Max Qty": _num(max_tok.group()),
                })
    return records


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------
def _supplier_common_keys(common_path, sheet="Common"):
    """(ordered base keys, {key: representative display name}) from the Common sheet."""
    try:
        df = cm.extract_molecule_table(common_path, sheet_name=sheet)
    except Exception:
        df = cm.extract_molecule_table(common_path)  # any single-table file
    name_col = next((c for c in df.columns if str(c).strip().lower() != "found in"), df.columns[0])
    keys, display = [], {}
    for value in df[name_col]:
        key = base_key(value)
        if key and key not in display:
            keys.append(key)
            display[key] = key.title()
    return keys, display


def _sum(rows, field):
    vals = [r[field] for r in rows if r.get(field) is not None]
    return round(sum(vals), 2) if vals else None


def _max(rows, field):
    vals = [r[field] for r in rows if r.get(field) is not None]
    return round(max(vals), 2) if vals else None


def enrich(common_path, forecast_dir, out_path, common_sheet="Common"):
    """Build the enriched workbook. Returns (out_path, matched_count)."""
    forecast_dir = Path(forecast_dir)
    keys, display = _supplier_common_keys(common_path, sheet=common_sheet or "Common")

    midas, hospital, tender = {}, {}, {}
    sources_found = []
    for path in sorted(forecast_dir.iterdir()):
        if not path.is_file() or path.name.startswith("~$"):
            continue
        kind, sheet = classify_forecast_file(path)
        if kind == "midas":
            midas = extract_midas(path, sheet); sources_found.append(("MIDAS", path.name))
        elif kind == "hospital":
            hospital = extract_hospital(path, sheet); sources_found.append(("Hospital", path.name))
        elif kind == "tender":
            tender = extract_tender(path); sources_found.append(("Tender", path.name))

    return _assemble_and_write(keys, display, midas, hospital, tender, sources_found, out_path)


def enrich_files(common_path, out_path, midas_path=None, hospital_path=None,
                 tender_path=None, common_sheet="Common", midas_sheet=None, hospital_sheet=None):
    """Like enrich(), but with each forecast file given explicitly (any may be None).

    The Mac app uses this: one drop per source instead of a single folder. Sheets
    are auto-detected when not pinned. Returns (out_path, matched_count).
    """
    keys, display = _supplier_common_keys(common_path, sheet=common_sheet or "Common")

    midas, hospital, tender = {}, {}, {}
    sources_found = []
    if midas_path:
        sheet = midas_sheet or _sheet_matching(midas_path, ["molecule", "2023", "2024", "2025"])
        midas = extract_midas(midas_path, sheet)
        sources_found.append(("MIDAS", Path(midas_path).name))
    if hospital_path:
        sheet = hospital_sheet or _sheet_matching(hospital_path, ["lebanon consumption", "40%"])
        hospital = extract_hospital(hospital_path, sheet)
        sources_found.append(("Hospital", Path(hospital_path).name))
    if tender_path:
        tender = extract_tender(tender_path)
        sources_found.append(("Tender", Path(tender_path).name))

    return _assemble_and_write(keys, display, midas, hospital, tender, sources_found, out_path)


def _assemble_and_write(keys, display, midas, hospital, tender, sources_found, out_path):
    """Build the summary + detail tables from the three source dicts and write them."""
    summary_rows, midas_detail, hosp_detail, tender_detail = [], [], [], []
    matched = 0
    for key in keys:
        m_rows = midas.get(key, [])
        h_rows = hospital.get(key, [])
        t_rows = tender.get(key, [])
        if not (m_rows or h_rows or t_rows):
            continue
        matched += 1
        found = " + ".join(name for name, rows in
                           (("MIDAS", m_rows), ("Hospital", h_rows), ("Tender", t_rows)) if rows)

        # MIDAS and tender are molecule-level; hospital consumption is per-dose.
        midas_cells = {
            "2023 Units": _sum(m_rows, "2023 Units"),
            "2024 Units": _sum(m_rows, "2024 Units"),
            "2025 Units": _sum(m_rows, "2025 Units"),
        }
        blank_midas = {k: None for k in midas_cells}
        tender_cells = {
            "Tender Min Qty": min((r["Min Qty"] for r in t_rows if r["Min Qty"] is not None), default=None),
            "Tender Max Qty": max((r["Max Qty"] for r in t_rows if r["Max Qty"] is not None), default=None),
            "Tender products": len(t_rows) or None,
        }
        blank_tender = {k: None for k in tender_cells}
        total_leb = _max(h_rows, "Total Lebanon Consumption 40%")  # molecule total (integer)

        if h_rows:
            # One row per dose; molecule-level MIDAS/tender sit on the first row only.
            for i, hr in enumerate(h_rows):
                summary_rows.append({
                    "Molecule": display[key],
                    "Dose": hr.get("Dose/Strength", ""),
                    "Found in": found,
                    **(midas_cells if i == 0 else blank_midas),
                    "Lebanon Consumption 40%": hr.get("Lebanon Consumption 40%"),
                    "Total Lebanon Consumption 40%": total_leb,
                    **(tender_cells if i == 0 else blank_tender),
                })
        else:
            summary_rows.append({
                "Molecule": display[key],
                "Dose": "",
                "Found in": found,
                **midas_cells,
                "Lebanon Consumption 40%": None,
                "Total Lebanon Consumption 40%": None,
                **tender_cells,
            })
        for r in m_rows:
            midas_detail.append({"Molecule": display[key], **r})
        for r in h_rows:
            hosp_detail.append({"Molecule": display[key], **r})
        for r in t_rows:
            tender_detail.append({"Molecule": display[key], **r})

    _write_workbook(out_path, summary_rows, midas_detail, hosp_detail, tender_detail, sources_found)
    return Path(out_path), matched


def _write_workbook(out_path, summary_rows, midas_detail, hosp_detail, tender_detail, sources_found):
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter
    book = xlsxwriter.Workbook(str(out_path), {"constant_memory": True})

    src_note = "  ·  ".join(f"{kind}: {name}" for kind, name in sources_found) or "no sources found"
    cm._write_sheet(book.add_worksheet("Summary"), book, pd.DataFrame(summary_rows),
                    title=f"Forecast data for {len(summary_rows)} common molecules   [{src_note}]")
    cm._write_sheet(book.add_worksheet("MIDAS detail"), book, pd.DataFrame(midas_detail))
    cm._write_sheet(book.add_worksheet("Hospital detail"), book, pd.DataFrame(hosp_detail))
    cm._write_sheet(book.add_worksheet("Tender detail"), book, pd.DataFrame(tender_detail))
    book.close()
    return out_path


def main():
    data = Path("Data")
    out, matched = enrich(
        data / "Supplier common molecules.xlsx",
        data / "Forecast",
        data / "Forecast enrichment.xlsx",
    )
    print(f"Matched {matched} molecules -> {out}")


if __name__ == "__main__":
    main()
