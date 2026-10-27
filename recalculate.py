"""LibreOfficeで再計算し、元の書式と数式を保って計算キャッシュだけを更新する。"""
from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def recalculate(path: Path) -> None:
    path = Path(path).resolve()
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if not executable:
        raise RuntimeError("LibreOffice (soffice) が必要です。")
    with tempfile.TemporaryDirectory(prefix="furima-recalc-") as temp:
        root = Path(temp)
        target = root / "calculated"
        target.mkdir()
        result = subprocess.run(
            [executable, f"-env:UserInstallation={(root / 'profile').as_uri()}",
             "--headless", "--convert-to", "xlsx", "--outdir", str(target), str(path)],
            capture_output=True, text=True, timeout=120, check=True,
        )
        calculated = target / path.name
        if not calculated.exists():
            raise RuntimeError(f"再計算ファイルがありません: {result.stdout} {result.stderr}")
        staged = root / "cached.xlsx"
        with ZipFile(path) as original, ZipFile(calculated) as evaluated, ZipFile(staged, "w") as output:
            shared = []
            if "xl/sharedStrings.xml" in evaluated.namelist():
                shared = ["".join(x.itertext()) for x in ET.fromstring(evaluated.read("xl/sharedStrings.xml"))]
            for entry in original.infolist():
                data = original.read(entry.filename)
                if entry.filename.startswith("xl/worksheets/sheet") and entry.filename.endswith(".xml"):
                    tree = ET.fromstring(data)
                    cache = {c.attrib["r"]: c for c in ET.fromstring(evaluated.read(entry.filename)).findall(".//m:sheetData/m:row/m:c", NS)}
                    for cell in tree.findall(".//m:sheetData/m:row/m:c", NS):
                        if cell.find("m:f", NS) is None:
                            continue
                        source = cache.get(cell.attrib["r"])
                        if source is None:
                            continue
                        for old in list(cell):
                            if old.tag in {f"{{{NS['m']}}}v", f"{{{NS['m']}}}is"}:
                                cell.remove(old)
                        cell.attrib.pop("t", None)
                        value = source.find("m:v", NS)
                        kind = source.get("t")
                        if kind in ("s", "str", "inlineStr"):
                            cell.set("t", "str")
                            text = (shared[int(value.text)] if kind == "s" and value is not None
                                    else "".join(source.find("m:is", NS).itertext()) if kind == "inlineStr"
                                    else value.text if value is not None else "")
                            ET.SubElement(cell, f"{{{NS['m']}}}v").text = text
                        elif value is not None:
                            if kind:
                                cell.set("t", kind)
                            ET.SubElement(cell, f"{{{NS['m']}}}v").text = value.text
                    data = ET.tostring(tree, encoding="utf-8", xml_declaration=True)
                output.writestr(entry, data)
        shutil.copyfile(staged, path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file", type=Path)
    recalculate(parser.parse_args().file)
