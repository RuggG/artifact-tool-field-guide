"""Read-only OOXML audit. Uses the standard library; never authors workbooks."""
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parent
RUNS = ROOT / "files" / "runs"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
results = []
for file in sorted((ROOT / "files").rglob("*.xlsx")):
    record = {"file": str(file.relative_to(ROOT)), "bytes": file.stat().st_size,
              "sha256": hashlib.sha256(file.read_bytes()).hexdigest()}
    try:
        with zipfile.ZipFile(file) as z:
            names = z.namelist()
            xml = {}
            parse_errors = []
            for name in names:
                if name.endswith((".xml", ".rels")):
                    try:
                        xml[name] = ET.fromstring(z.read(name))
                    except Exception as exc:
                        parse_errors.append({"part": name, "error": str(exc)})
            charts = []
            for name, tree in xml.items():
                if re.fullmatch(r"xl/(?:drawings/)?charts/chart\d+\.xml", name):
                    charts.append({"part": name, "types": [n.tag.split("}")[-1] for n in tree.iter()
                        if n.tag.split("}")[-1].endswith("Chart")],
                        "series": len(tree.findall(".//c:ser", NS)),
                        "formulas": [n.text for n in tree.findall(".//c:f", NS)],
                        "trendlines": len(tree.findall(".//c:trendline", NS)),
                        "errorBars": len(tree.findall(".//c:errBars", NS)),
                        "dataTables": len(tree.findall(".//c:dTable", NS))})
            sheets = []
            for name, tree in xml.items():
                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name):
                    sheets.append({"part": name,
                        "cells": len(tree.findall(".//s:sheetData/s:row/s:c", NS)),
                        "formulas": [{"cell": c.attrib.get("r"), "text": f.text, "attributes": f.attrib,
                                      "cached": c.findtext("s:v", namespaces=NS)}
                                     for c in tree.findall(".//s:sheetData/s:row/s:c", NS)
                                     for f in c.findall("s:f", NS)],
                        "merges": [n.attrib for n in tree.findall(".//s:mergeCell", NS)],
                        "panes": [n.attrib for n in tree.findall(".//s:pane", NS)],
                        "conditionalFormats": [n.attrib for n in tree.findall(".//s:cfRule", NS)],
                        "validations": [dict(n.attrib, formulas=[(c.tag.split('}')[-1],c.text) for c in n])
                                        for n in tree.findall(".//s:dataValidation", NS)],
                        "sparklineGroups": sum(n.tag.split('}')[-1]=='sparklineGroup' for n in tree.iter()),
                        "images": sum(n.tag.split('}')[-1]=='pic' for n in tree.iter()),
                        "errors": [(c.attrib.get('r'),c.findtext('s:v',namespaces=NS))
                                   for c in tree.findall(".//s:c[@t='e']",NS)]})
            record.update({"xmlParseErrors": parse_errors, "parts": names, "charts": charts,
                "worksheets": sheets,
                "tableParts": [n for n in names if re.fullmatch(r"xl/tables/table\d+\.xml",n)],
                "pivotParts": [n for n in names if 'pivot' in n.lower()],
                "slicerParts": [n for n in names if 'slicer' in n.lower()],
                "commentParts": [n for n in names if 'comment' in n.lower() or 'person' in n.lower()],
                "mediaParts": [n for n in names if n.startswith('xl/media/')],
                "names": [dict(n.attrib,formula=n.text) for n in xml.get('xl/workbook.xml',ET.Element('none')).findall('.//s:definedName',NS)]})
    except Exception as exc:
        record["error"] = str(exc)
    results.append(record)
(RUNS / "export-audit.json").write_text(json.dumps(results, indent=2))
print(json.dumps({"workbooks":len(results),"parseFailures":sum(bool(x.get("error") or x.get("xmlParseErrors")) for x in results),
    "charts": {Path(x["file"]).stem:[t for c in x.get("charts",[]) for t in c["types"]] for x in results if Path(x["file"]).name.startswith('chart-')}},indent=2))
