"""Change one cached cell value in an .xlsx without disturbing anything else.

openpyxl cannot do this. It loads either formulas or cached results, never both,
so saving with it writes the formulas back and discards every cached value in
the file. A validator reading the result then sees None in every formula cell,
which makes a fault injection test fail for a reason that has nothing to do with
the fault.

An .xlsx is a zip of XML. A formula cell looks like:

    <c r="B6" s="12"><f>SUM(C10:C25)</f><v>759580.0017</v></c>

The cached result is the <v> element. Rewriting just that leaves the formula,
the styles and every other cell untouched, which is what a test needs: one
deliberate change against an otherwise identical workbook.

Error values are stored the same way with a type attribute, as t="e" and a <v>
holding the error text, which is how a division by zero is represented.
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path


def _sheet_part(zf: zipfile.ZipFile, sheet_name: str) -> str:
    """Map a sheet's display name to its XML part inside the archive."""
    book = zf.read("xl/workbook.xml").decode("utf-8")
    m = re.search(rf'<sheet[^>]*name="{re.escape(sheet_name)}"[^>]*/>', book)
    if not m:
        raise KeyError(f"no sheet named {sheet_name!r}")
    rid = re.search(r'r:id="([^"]+)"', m.group(0))
    if not rid:
        raise KeyError(f"sheet {sheet_name!r} has no relationship id")

    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rel = re.search(rf'<Relationship[^>]*Id="{rid.group(1)}"[^>]*/>', rels)
    target = re.search(r'Target="([^"]+)"', rel.group(0)).group(1)
    return "xl/" + target.lstrip("/").replace("worksheets/", "worksheets/", 1)


def set_cached_value(path: Path, sheet: str, cell_ref: str, value, *, as_error: bool = False):
    """Replace the cached result of one cell. Returns the previous value."""
    path = Path(path)
    with zipfile.ZipFile(path) as zf:
        part = _sheet_part(zf, sheet)
        parts = {n: zf.read(n) for n in zf.namelist()}
        order = zf.namelist()

    xml = parts[part].decode("utf-8")
    pattern = re.compile(rf'(<c r="{cell_ref}"[^>]*?)(/>|>(.*?)</c>)', re.DOTALL)
    m = pattern.search(xml)
    if not m:
        raise KeyError(f"cell {cell_ref} not found on sheet {sheet!r}")

    attrs, _, body = m.group(1), m.group(2), m.group(3) or ""
    previous = None
    vm = re.search(r"<v>(.*?)</v>", body, re.DOTALL)
    if vm:
        previous = vm.group(1)

    # Drop any existing type attribute, then set the one this value needs.
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    if as_error:
        attrs += ' t="e"'
    elif isinstance(value, str):
        attrs += ' t="str"'

    formula = re.search(r"<f[^>]*>.*?</f>|<f[^>]*/>", body, re.DOTALL)
    keep = formula.group(0) if formula else ""
    new_cell = f"{attrs}>{keep}<v>{value}</v></c>"
    xml = xml[: m.start()] + new_cell + xml[m.end():]
    parts[part] = xml.encode("utf-8")

    tmp = path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name in order:
            out.writestr(name, parts[name])
    shutil.move(str(tmp), str(path))
    return previous


def col_letter(n: int) -> str:
    """1 -> A, 2 -> B, 27 -> AA."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ref(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"
