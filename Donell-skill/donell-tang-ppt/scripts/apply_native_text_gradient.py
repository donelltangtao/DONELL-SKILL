#!/usr/bin/env python3
"""Apply native DrawingML gradient fills to text runs in a PPTX.

Usage:
  python apply_native_text_gradient.py input.pptx output.pptx config.json

Config format:
{
  "slides": {
    "1": [
      {
        "shape_name": "Title 1",
        "angle": 0,
        "rotate_with_shape": true,
        "stops": [
          {"position": 0, "color": "FFFFFF", "transparency": 0},
          {"position": 55, "color": "F3EEFF", "transparency": 0},
          {"position": 100, "color": "8A4DFF", "transparency": 0}
        ]
      }
    ]
  }
}

The script patches a:gradFill into a:rPr/a:defRPr/a:endParaRPr. It removes
existing a:solidFill, a:gradFill, a:noFill, a:pattFill and a:blipFill first.
"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}
ET.register_namespace("a", A)
ET.register_namespace("p", P)

FILL_TAGS = {
    f"{{{A}}}noFill",
    f"{{{A}}}solidFill",
    f"{{{A}}}gradFill",
    f"{{{A}}}blipFill",
    f"{{{A}}}pattFill",
    f"{{{A}}}grpFill",
}


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def percent_to_pos(value: float) -> str:
    value = max(0.0, min(100.0, float(value)))
    return str(round(value * 1000))


def transparency_to_alpha(value: float) -> str:
    value = max(0.0, min(100.0, float(value)))
    return str(round((100.0 - value) * 1000))


def angle_to_ooxml(degrees: float) -> str:
    return str(round((float(degrees) % 360.0) * 60000))


def make_gradient(spec: dict) -> ET.Element:
    grad = ET.Element(qn(A, "gradFill"), {
        "rotWithShape": "1" if spec.get("rotate_with_shape", True) else "0"
    })
    gs_lst = ET.SubElement(grad, qn(A, "gsLst"))
    stops = spec.get("stops") or []
    if len(stops) < 2:
        raise ValueError("Each gradient requires at least two stops")
    for stop in sorted(stops, key=lambda x: float(x["position"])):
        color = str(stop["color"]).replace("#", "").upper()
        if len(color) != 6:
            raise ValueError(f"Invalid RGB color: {color}")
        gs = ET.SubElement(gs_lst, qn(A, "gs"), {"pos": percent_to_pos(stop["position"])})
        clr = ET.SubElement(gs, qn(A, "srgbClr"), {"val": color})
        transparency = float(stop.get("transparency", 0))
        if transparency:
            ET.SubElement(clr, qn(A, "alpha"), {"val": transparency_to_alpha(transparency)})
    ET.SubElement(grad, qn(A, "lin"), {
        "ang": angle_to_ooxml(spec.get("angle", 0)),
        "scaled": "1" if spec.get("scaled", False) else "0",
    })
    return grad


def set_gradient_on_rpr(rpr: ET.Element, spec: dict) -> None:
    for child in list(rpr):
        if child.tag in FILL_TAGS:
            rpr.remove(child)
    # DrawingML expects fill before effects such as ln/effectLst in rPr.
    grad = make_gradient(spec)
    insert_at = 0
    for i, child in enumerate(list(rpr)):
        local = child.tag.rsplit("}", 1)[-1]
        if local in {"ln", "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl"}:
            insert_at = i
            break
        insert_at = i + 1
    rpr.insert(insert_at, grad)


def iter_named_shapes(root: ET.Element):
    for sp in root.findall(".//p:sp", NS):
        c_nv_pr = sp.find("./p:nvSpPr/p:cNvPr", NS)
        if c_nv_pr is not None:
            yield c_nv_pr.get("name", ""), sp


def apply_to_shape(shape: ET.Element, spec: dict) -> int:
    targets = []
    targets.extend(shape.findall(".//a:r/a:rPr", NS))
    targets.extend(shape.findall(".//a:p/a:endParaRPr", NS))
    targets.extend(shape.findall(".//a:p/a:pPr/a:defRPr", NS))
    # Runs may omit rPr. Create one so native fill is explicit.
    for run in shape.findall(".//a:r", NS):
        if run.find("a:rPr", NS) is None:
            run.insert(0, ET.Element(qn(A, "rPr")))
    targets = []
    targets.extend(shape.findall(".//a:r/a:rPr", NS))
    targets.extend(shape.findall(".//a:p/a:endParaRPr", NS))
    targets.extend(shape.findall(".//a:p/a:pPr/a:defRPr", NS))
    if not targets:
        # Create paragraph default properties when the shape only has plain text.
        for p_el in shape.findall(".//a:p", NS):
            ppr = p_el.find("a:pPr", NS)
            if ppr is None:
                ppr = ET.Element(qn(A, "pPr"))
                p_el.insert(0, ppr)
            defrpr = ppr.find("a:defRPr", NS)
            if defrpr is None:
                defrpr = ET.SubElement(ppr, qn(A, "defRPr"))
            targets.append(defrpr)
    for target in targets:
        set_gradient_on_rpr(target, spec)
    return len(targets)


def patch_slide(slide_path: Path, specs: list[dict]) -> list[str]:
    tree = ET.parse(slide_path)
    root = tree.getroot()
    shapes = dict(iter_named_shapes(root))
    report = []
    for spec in specs:
        name = spec["shape_name"]
        shape = shapes.get(name)
        if shape is None:
            raise KeyError(f"Shape not found in {slide_path.name}: {name}")
        count = apply_to_shape(shape, spec)
        report.append(f"{slide_path.name}: {name}: patched {count} text property nodes")
    tree.write(slide_path, encoding="UTF-8", xml_declaration=True)
    return report


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__, file=sys.stderr)
        return 2
    src, dst, cfg = map(Path, sys.argv[1:])
    if not src.exists():
        raise FileNotFoundError(src)
    config = json.loads(cfg.read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        with zipfile.ZipFile(src, "r") as zin:
            zin.extractall(td_path)
        reports = []
        for slide_no, specs in config.get("slides", {}).items():
            slide_path = td_path / "ppt" / "slides" / f"slide{int(slide_no)}.xml"
            if not slide_path.exists():
                raise FileNotFoundError(slide_path)
            reports.extend(patch_slide(slide_path, specs))
        if dst.exists():
            dst.unlink()
        with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for path in td_path.rglob("*"):
                if path.is_file():
                    zout.write(path, path.relative_to(td_path).as_posix())
    print("\n".join(reports))
    print(f"Wrote: {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
