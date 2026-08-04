"""Find supplier molecules that also appear in the High Monopoly or Potential lists.

Given ONE supplier workbook and the two reference lists (High Monopoly and
Potential molecules), this finds every supplier molecule whose name appears in
EITHER reference list, matched on a "smart" base name that ignores strength/form.
It writes a workbook with a Supplier sheet (the full product list) and a Common
sheet (the matches, each tagged with which list it was found in).

The table extraction and name normalization are reused from compare_molecules.py.
Use `compare_supplier(...)` as the high-level entry point (the UI calls it).
"""

from pathlib import Path

import compare_molecules as cm


def _key_set(df):
    """Set of smart base names for a reference DataFrame's molecule column."""
    col = cm._name_column(df)
    keys = {cm._base(v) for v in df[col]}
    keys.discard("")
    return keys


def find_supplier_common(supplier_df, hp_keys, pot_keys):
    """Return supplier rows whose base name is in either reference set, tagged."""
    scol = cm._name_column(supplier_df)

    def found_in(name):
        base = cm._base(name)
        where = []
        if base and base in hp_keys:
            where.append("High Monopoly")
        if base and base in pot_keys:
            where.append("Potential")
        return " + ".join(where)

    tags = supplier_df[scol].apply(found_in)
    common = supplier_df[tags != ""].copy()
    common.insert(0, "Found in", tags[tags != ""].values)
    return common.reset_index(drop=True)


def compare_supplier(supplier_path, hp_path, pot_path, out_path,
                     supplier_sheet=None, hp_sheet=None, pot_sheet=None):
    """Read the three files, find supplier molecules in either list, write workbook.

    Returns (out_path, common_count). Sheet arguments optionally pin which sheet
    to read from each file; None auto-detects (preferring a hinting sheet name).
    """
    supplier_df = cm.extract_molecule_table(supplier_path, sheet_name=supplier_sheet)
    hp_df = cm.extract_molecule_table(
        hp_path, prefer_keywords=("monopoly", "priority"), sheet_name=hp_sheet)
    pot_df = cm.extract_molecule_table(
        pot_path, prefer_keywords=("potential",), sheet_name=pot_sheet)

    common = find_supplier_common(supplier_df, _key_set(hp_df), _key_set(pot_df))

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    import xlsxwriter
    book = xlsxwriter.Workbook(str(out_path), {"constant_memory": True})
    cm._write_sheet(book.add_worksheet("Supplier"), book, supplier_df,
                    title=f"Supplier: {Path(supplier_path).name}")
    cm._write_sheet(book.add_worksheet("Common"), book, common,
                    title=f"Supplier molecules found in either list: {len(common)}")
    book.close()
    return out_path, len(common)


def main():
    """CLI: compare the bundled sample files under Data/."""
    data = Path("Data")
    out, count = compare_supplier(
        data / "Suppliers" / "BETA.xlsx",
        data / "molecule_explorer (2) - High Monopoly.xlsx",
        data / "Potential molecules final.xlsx",
        data / "Supplier common molecules.xlsx",
        supplier_sheet="Sheet1",
    )
    print(f"Common (in either list): {count}\nWrote -> {out}")


if __name__ == "__main__":
    main()
