#!/usr/bin/env python3
"""
ecore_to_docx.py

Convert an Eclipse EMF .ecore file into a Word document containing one table
per EClass. Each table lists the class's structural features (EAttributes and
EReferences) with columns:

    Name | Kind | Type | Cardinality | Default | Description

- "Kind" distinguishes attribute / reference / containment.
- "Type" resolves cross-package references via href (e.g. ecore primitives,
  or other .ecore files in the same directory, if provided).
- "Description" is pulled from EAnnotations: it looks first for a GenModel
  "documentation" detail, then falls back to any other EAnnotation detail
  whose key contains "doc" (case-insensitive), then the first detail value.

Usage:
    python ecore_to_docx.py model.ecore -o data_definitions.docx
    python ecore_to_docx.py *.ecore -o combined.docx
    python ecore_to_docx.py model.ecore --title "Data Definition Document"
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------------------
# XML namespaces used in .ecore files
# ---------------------------------------------------------------------------
NS = {
    "xmi":   "http://www.omg.org/XMI",
    "xsi":   "http://www.w3.org/2001/XMLSchema-instance",
    "ecore": "http://www.eclipse.org/emf/2002/Ecore",
}

ECORE_PRIMITIVE_RE = re.compile(
    r"http://www\.eclipse\.org/emf/2002/Ecore#//(E[A-Za-z]+)"
)

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------
@dataclass
class Feature:
    name: str
    kind: str          # "attribute", "reference", or "containment"
    type_str: str
    cardinality: str
    default: str
    description: str


@dataclass
class EClassInfo:
    name: str
    package: str
    is_abstract: bool
    is_interface: bool
    supertypes: list[str] = field(default_factory=list)
    description: str = ""
    features: list[Feature] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def _cardinality(lower: str, upper: str) -> str:
    """Format lowerBound/upperBound into a readable cardinality."""
    lo = lower if lower not in (None, "") else "0"
    up = upper if upper not in (None, "") else "1"
    if up == "-1":
        up = "*"
    if lo == up:
        return lo
    return f"{lo}..{up}"


def _resolve_type(etype_attr: Optional[str], href_attr: Optional[str]) -> str:
    """
    Turn an eType reference into a human-readable type name.
    Handles three forms:
      1. eType="#//SomeClass"           -> SomeClass (same file)
      2. <eType href="...#//EString"/>  -> EString   (ecore primitives or cross-file)
      3. eType="ecore:EDataType ...#//EString"
    """
    ref = etype_attr or href_attr or ""
    if not ref:
        return ""

    # Ecore primitive via full URL
    m = ECORE_PRIMITIVE_RE.search(ref)
    if m:
        return m.group(1)

    # Fragment-style: "#//Something" or "some.ecore#//Something"
    if "#//" in ref:
        frag = ref.split("#//", 1)[1]
        # Could be nested like "Package/Class" — take the last segment
        return frag.split("/")[-1]

    # Fallback: strip any URI prefix
    return ref.rsplit("/", 1)[-1]


def _extract_description(element: ET.Element) -> str:
    """Look for documentation-style EAnnotations on an element."""
    for ann in element.findall("eAnnotations"):
        details = ann.findall("details")
        # Priority 1: a detail with key exactly "documentation"
        for d in details:
            if d.get("key", "").lower() == "documentation":
                return (d.get("value") or "").strip()
        # Priority 2: any detail whose key contains "doc"
        for d in details:
            if "doc" in d.get("key", "").lower():
                return (d.get("value") or "").strip()
        # Priority 3: first detail's value
        if details:
            return (details[0].get("value") or "").strip()
    return ""


def _parse_feature(elem: ET.Element) -> Feature:
    """Parse an eStructuralFeatures element (attribute or reference)."""
    xsi_type = elem.get(f"{{{NS['xsi']}}}type", "")
    is_ref = xsi_type.endswith("EReference")
    containment = elem.get("containment", "false").lower() == "true"

    if is_ref:
        kind = "containment" if containment else "reference"
    else:
        kind = "attribute"

    # Type can be inline attribute OR a nested <eType href="..."/>
    etype_attr = elem.get("eType")
    href_attr = None
    etype_child = elem.find("eType")
    if etype_child is not None:
        href_attr = etype_child.get(f"{{{NS['xmi']}}}href") or etype_child.get("href")

    type_str = _resolve_type(etype_attr, href_attr)

    cardinality = _cardinality(elem.get("lowerBound"), elem.get("upperBound"))
    default = elem.get("defaultValueLiteral", "") or ""
    description = _extract_description(elem)

    return Feature(
        name=elem.get("name", ""),
        kind=kind,
        type_str=type_str,
        cardinality=cardinality,
        default=default,
        description=description,
    )


def _parse_supertypes(eclass: ET.Element) -> list[str]:
    raw = eclass.get("eSuperTypes", "")
    if not raw:
        return []
    names = []
    for token in raw.split():
        names.append(_resolve_type(token, None))
    return names


def parse_ecore(path: Path) -> list[EClassInfo]:
    """Parse a .ecore file and return all EClasses within all EPackages."""
    tree = ET.parse(path)
    root = tree.getroot()

    # Root may itself be an EPackage, or an <xmi:XMI> wrapping multiple packages
    if root.tag.endswith("EPackage"):
        packages = [root]
    else:
        packages = root.findall(".//{%s}EPackage" % NS["ecore"]) or root.findall(".//EPackage")
        if not packages:
            packages = [root]

    classes: list[EClassInfo] = []
    for pkg in packages:
        pkg_name = pkg.get("name", path.stem)
        for classifier in pkg.findall("eClassifiers"):
            xsi_type = classifier.get(f"{{{NS['xsi']}}}type", "")
            if not xsi_type.endswith("EClass"):
                continue  # skip EDataType, EEnum for now

            info = EClassInfo(
                name=classifier.get("name", ""),
                package=pkg_name,
                is_abstract=classifier.get("abstract", "false").lower() == "true",
                is_interface=classifier.get("interface", "false").lower() == "true",
                supertypes=_parse_supertypes(classifier),
                description=_extract_description(classifier),
            )
            for feat_elem in classifier.findall("eStructuralFeatures"):
                info.features.append(_parse_feature(feat_elem))
            classes.append(info)

    return classes


# ---------------------------------------------------------------------------
# Word document generation
# ---------------------------------------------------------------------------
HEADER_FILL = "2E75B6"   # blue header
ROW_ALT_FILL = "F2F2F2"  # light grey zebra
BORDER_COLOR = "BFBFBF"

COLUMNS = [
    ("Name",        Cm(3.2)),
    ("Kind",        Cm(2.2)),
    ("Type",        Cm(3.0)),
    ("Cardinality", Cm(2.0)),
    ("Default",     Cm(2.0)),
    ("Description", Cm(5.0)),
]


def _shade_cell(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def _set_cell_borders(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), BORDER_COLOR)
        borders.append(b)
    tc_pr.append(borders)


def _write_cell(cell, text: str, *, bold: bool = False, italic: bool = False,
                color: Optional[RGBColor] = None, size: int = 10) -> None:
    cell.text = ""  # clear default paragraph
    para = cell.paragraphs[0]
    run = para.add_run(text or "")
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _set_cell_borders(cell)


def _add_class_heading(doc: Document, cls: EClassInfo) -> None:
    heading = doc.add_heading(level=2)
    run = heading.add_run(cls.name)
    run.font.size = Pt(14)

    # Subtitle line with package + modifiers + supertypes
    bits = [f"Package: {cls.package}"]
    if cls.is_abstract:
        bits.append("abstract")
    if cls.is_interface:
        bits.append("interface")
    if cls.supertypes:
        bits.append("extends " + ", ".join(cls.supertypes))
    sub = doc.add_paragraph()
    sub_run = sub.add_run(" · ".join(bits))
    sub_run.italic = True
    sub_run.font.size = Pt(9)
    sub_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    if cls.description:
        desc = doc.add_paragraph(cls.description)
        desc.paragraph_format.space_after = Pt(4)


def _add_class_table(doc: Document, cls: EClassInfo) -> None:
    if not cls.features:
        note = doc.add_paragraph()
        r = note.add_run("(No structural features defined.)")
        r.italic = True
        r.font.size = Pt(9)
        return

    table = doc.add_table(rows=1 + len(cls.features), cols=len(COLUMNS))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    # Column widths
    for i, (_, width) in enumerate(COLUMNS):
        for row in table.rows:
            row.cells[i].width = width

    # Header row
    header = table.rows[0]
    for i, (label, _) in enumerate(COLUMNS):
        cell = header.cells[i]
        _write_cell(cell, label, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), size=10)
        _shade_cell(cell, HEADER_FILL)

    # Data rows
    for r_idx, feat in enumerate(cls.features, start=1):
        row = table.rows[r_idx]
        values = [
            feat.name,
            feat.kind,
            feat.type_str,
            feat.cardinality,
            feat.default,
            feat.description,
        ]
        for c_idx, val in enumerate(values):
            cell = row.cells[c_idx]
            _write_cell(cell, val, size=9)
            if r_idx % 2 == 0:
                _shade_cell(cell, ROW_ALT_FILL)

    doc.add_paragraph()  # spacing between tables


def build_document(
    classes_by_file: dict[str, list[EClassInfo]],
    title: str,
) -> Document:
    doc = Document()

    # Landscape helps the wider tables breathe
    section = doc.sections[0]
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    # Title
    title_p = doc.add_heading(title, level=0)
    for run in title_p.runs:
        run.font.size = Pt(22)

    intro = doc.add_paragraph(
        "This document defines the structural classes extracted from the provided "
        "Ecore model(s). Each section below describes an EClass and lists its "
        "attributes and references."
    )
    intro.paragraph_format.space_after = Pt(12)

    for source_file, classes in classes_by_file.items():
        if len(classes_by_file) > 1:
            h1 = doc.add_heading(level=1)
            h1.add_run(Path(source_file).name)

        # Group by package for readability
        by_package: dict[str, list[EClassInfo]] = {}
        for cls in classes:
            by_package.setdefault(cls.package, []).append(cls)

        for pkg_name, pkg_classes in by_package.items():
            if len(by_package) > 1 or len(classes_by_file) > 1:
                ph = doc.add_heading(level=1 if len(classes_by_file) == 1 else 2)
                ph.add_run(f"Package: {pkg_name}")

            for cls in sorted(pkg_classes, key=lambda c: c.name.lower()):
                _add_class_heading(doc, cls)
                _add_class_table(doc, cls)

    return doc


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(
        description="Convert .ecore file(s) into a Word document of data-definition tables."
    )
    ap.add_argument("inputs", nargs="+", help="One or more .ecore files")
    ap.add_argument("-o", "--output", default="ecore_tables.docx",
                    help="Output .docx path (default: ecore_tables.docx)")
    ap.add_argument("--title", default="Data Definitions",
                    help="Document title (default: 'Data Definitions')")
    args = ap.parse_args()

    classes_by_file: dict[str, list[EClassInfo]] = {}
    total = 0
    for raw in args.inputs:
        p = Path(raw)
        if not p.exists():
            print(f"ERROR: file not found: {p}", file=sys.stderr)
            return 2
        try:
            classes = parse_ecore(p)
        except ET.ParseError as e:
            print(f"ERROR: could not parse {p}: {e}", file=sys.stderr)
            return 2
        classes_by_file[str(p)] = classes
        total += len(classes)
        print(f"  {p.name}: {len(classes)} EClass(es)")

    if total == 0:
        print("No EClasses found in the provided file(s).", file=sys.stderr)
        return 1

    doc = build_document(classes_by_file, title=args.title)
    out_path = Path(args.output)
    doc.save(out_path)
    print(f"\nWrote {total} table(s) to {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
