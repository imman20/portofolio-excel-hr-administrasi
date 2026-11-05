"""Saya memeriksa integritas workbook terhadap audit recalculation final."""
import hashlib
import json
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

ROOT = Path(__file__).resolve().parents[1]
AUDIT = json.loads((ROOT / "belajar/audit_repo.json").read_text())
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


@pytest.mark.parametrize("record", AUDIT["files"], ids=lambda r: r["file"])
def test_final_workbook_integrity(record):
    path = ROOT / record["file"]
    assert hashlib.sha256(path.read_bytes()).hexdigest() == record["sha256"]
    formulas = 0
    errors = []
    with ZipFile(path) as workbook:
        for name in workbook.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                sheet = ET.fromstring(workbook.read(name))
                formulas += len(sheet.findall(".//s:f", NS))
                errors.extend(c.attrib["r"] for c in sheet.findall(".//s:c", NS) if c.get("t") == "e")
    assert formulas == record["formulas"]
    assert not errors


def test_audit_totals_match_workbooks():
    assert len(list(ROOT.rglob("*.xlsx"))) == len(AUDIT["files"]) == 47
    assert sum(r["formulas"] for r in AUDIT["files"]) == AUDIT["total_formulas"] == 327539
    assert AUDIT["total_errors"] == 0
