"""Standard-library document extraction and lightweight conversion helpers.

These helpers support the PRD's local-first baseline without pretending to be a
full Office rendering engine. They extract enough structure for previews,
reports, and minimal OOXML conversion artifacts while leaving high-fidelity
Office, MathType, OMML writeback, and macro work to the local client contracts.
"""

from __future__ import annotations

import html
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any


def extract_docx_blocks(path: Path) -> list[dict[str, Any]]:
    """Extract visible DOCX paragraph blocks for previews and Word-to-PPT."""
    with zipfile.ZipFile(path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
    paragraphs = re.findall(r"<w:p[\s\S]*?</w:p>", document_xml)
    blocks: list[dict[str, Any]] = []
    for index, paragraph in enumerate(paragraphs, start=1):
        text = _extract_text(paragraph, "w:t")
        if not text:
            text = _strip_xml(paragraph)
        if not text:
            continue
        level = _heading_level(paragraph, index)
        blocks.append({"text": text, "level": level})
    return blocks or [{"text": path.stem, "level": 1}]


def extract_docx_object_summary(path: Path) -> dict[str, Any]:
    """Count DOCX objects that must be preserved or handed to the desktop side."""
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        try:
            document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
        except KeyError:
            document_xml = ""
    image_parts = sorted(name for name in names if name.startswith("word/media/"))
    embedded_parts = sorted(name for name in names if name.startswith("word/embeddings/"))
    omml_formulas = document_xml.count("<m:oMath")
    mathtype_objects = len(embedded_parts)
    if mathtype_objects == 0 and ("Equation Native" in document_xml or "MathType" in document_xml):
        mathtype_objects = 1
    return {
        "paragraphs": document_xml.count("<w:p"),
        "headings": len(re.findall(r"<w:pStyle[^>]+w:val=\"Heading", document_xml)),
        "images": len(image_parts),
        "tables": document_xml.count("<w:tbl"),
        "omml_formulas": omml_formulas,
        "mathtype_objects": mathtype_objects,
        "embedded_objects": len(embedded_parts),
        "formulas": omml_formulas + mathtype_objects,
        "image_parts": image_parts[:12],
        "embedded_parts": embedded_parts[:12],
    }


def extract_pptx_slides(path: Path) -> list[dict[str, Any]]:
    """Extract PPTX slide text, notes, and object counts for conversion reports."""
    with zipfile.ZipFile(path) as archive:
        names = sorted(
            [name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=_slide_number,
        )
        slides: list[dict[str, Any]] = []
        for index, name in enumerate(names, start=1):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            cleaned = _pptx_texts(xml)
            title = cleaned[0] if cleaned else f"幻灯片 {index}"
            body = cleaned[1:] if len(cleaned) > 1 else []
            related_parts = _pptx_related_parts(archive, name)
            notes = _pptx_notes(archive, related_parts)
            formula_count = _pptx_formula_count(xml)
            slides.append(
                {
                    "title": title,
                    "body": body,
                    "notes": notes,
                    "index": index,
                    "text_box_count": len(re.findall(r"<p:txBody\b", xml)),
                    "shape_count": len(re.findall(r"<p:sp\b", xml)),
                    "table_count": len(re.findall(r"<(?:\w+:)?tbl\b", xml)),
                    "image_count": len({part for part in related_parts if "/media/" in f"/{part.lower()}"}),
                    "chart_count": len({part for part in related_parts if "/charts/" in f"/{part.lower()}"}),
                    "formula_count": formula_count,
                }
            )
    return slides or [{"title": path.stem, "body": [], "index": 1}]


def extract_xlsx_sheets(path: Path) -> list[dict[str, Any]]:
    """Extract worksheet previews and formula counts from an XLSX workbook."""
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        shared_strings = _xlsx_shared_strings(archive)
        sheet_names = _xlsx_sheet_names(archive)
        worksheet_paths = sorted(
            [name for name in names if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")],
            key=_worksheet_number,
        )
        sheets: list[dict[str, Any]] = []
        for index, name in enumerate(worksheet_paths, start=1):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            cells = _xlsx_cells(xml, shared_strings)
            formulas = xml.count("<f")
            merged = xml.count("<mergeCell")
            objects = _xlsx_sheet_objects(archive, name, xml)
            sheets.append(
                {
                    "name": sheet_names.get(index, f"Sheet{index}"),
                    "index": index,
                    "cells": cells,
                    "formula_count": formulas,
                    "merged_cells": merged,
                    **objects,
                }
            )
    return sheets or [
        {
            "name": path.stem,
            "index": 1,
            "cells": [],
            "formula_count": 0,
            "merged_cells": 0,
            "chart_count": 0,
            "drawing_count": 0,
            "image_count": 0,
            "comment_count": 0,
            "table_count": 0,
        }
    ]


def build_pptx_from_docx(
    blocks: list[dict[str, Any]],
    target: Path,
    max_chars: int = 320,
    auto_pagination: bool = True,
    generate_toc: bool = False,
    object_preservation: dict[str, Any] | None = None,
) -> None:
    """Create a lightweight PPTX from DOCX blocks and preservation metadata."""
    slide_limit = max_chars if auto_pagination else 1_000_000
    slides = _slides_from_blocks(blocks, slide_limit)
    if generate_toc and slides:
        toc_items = [slide.get("title", f"幻灯片 {index}") for index, slide in enumerate(slides, start=1)]
        slides = [{"title": "目录", "body": toc_items[:18]}, *slides]
    preservation_lines = _docx_object_preservation_lines(object_preservation or {})
    if preservation_lines:
        slides.append({"title": "对象保留清单", "body": preservation_lines})
    build_pptx(slides, target)


def build_docx_from_slides(
    slides: list[dict[str, Any]],
    target: Path,
    mode: str = "逐页讲义模式",
    generate_toc: bool = True,
    template_name: str = "",
    include_notes: bool = True,
    retain_images: bool = True,
    retain_formulas: bool = True,
) -> None:
    """Create a lightweight DOCX handout from extracted PPT slide structure."""
    paragraphs: list[dict[str, Any]] = []
    if mode:
        paragraphs.append({"text": f"PPT 转 Word 模式：{mode}", "style": "Heading1"})
    if template_name:
        paragraphs.append({"text": f"Word 模板：{template_name}", "style": "Normal"})
    if generate_toc and slides:
        paragraphs.append({"text": "目录", "style": "Heading1"})
        for index, slide in enumerate(slides, start=1):
            paragraphs.append({"text": f"{index}. {slide.get('title', f'幻灯片 {index}')}", "style": "Normal"})
    if mode == "大纲模式":
        paragraphs.extend(_docx_slide_outline(slides))
        build_docx(paragraphs, target)
        return
    for slide in slides:
        paragraphs.append({"text": slide["title"], "style": "Heading1"})
        if mode == "图文混排模式":
            paragraphs.append({"text": f"幻灯片缩略图占位：第 {slide.get('index', 1)} 页", "style": "Normal"})
        summary = _pptx_slide_summary(
            slide,
            include_notes=include_notes,
            retain_images=retain_images,
            retain_formulas=retain_formulas,
        )
        if summary:
            paragraphs.append({"text": f"对象摘要：{summary}", "style": "Normal"})
        notes = [item for item in slide.get("notes", []) if item] if include_notes else []
        body = [item for item in slide.get("body", []) if item]
        if mode == "备注优先模式" and notes:
            paragraphs.append({"text": "讲者备注", "style": "Normal"})
            for item in notes:
                paragraphs.append({"text": item, "style": "Normal"})
        for item in body:
            paragraphs.append({"text": item, "style": "Normal"})
        if mode != "备注优先模式" and notes:
            paragraphs.append({"text": "备注", "style": "Normal"})
            for item in notes:
                paragraphs.append({"text": item, "style": "Normal"})
    build_docx(paragraphs, target)


def build_pdf_from_xlsx(sheets: list[dict[str, Any]], target: Path, title: str) -> None:
    """Write a text PDF summary for an Excel-to-PDF conversion artifact."""
    lines = [f"Excel to PDF: {title}", ""]
    for sheet in sheets:
        lines.append(f"Sheet {sheet.get('index')}: {sheet.get('name')}")
        if sheet.get("conversion_range"):
            lines.append(f"Range: {sheet.get('conversion_range')}")
        if sheet.get("formula_mode"):
            lines.append(f"Formula mode: {sheet.get('formula_mode')}")
        lines.append(
            "  ".join(
                [
                    f"Cells: {len(sheet.get('cells', []))}",
                    f"Formulas: {sheet.get('formula_count', 0)}",
                    f"Merged: {sheet.get('merged_cells', 0)}",
                    f"Charts: {sheet.get('chart_count', 0)}",
                    f"Images: {sheet.get('image_count', 0)}",
                    f"Comments: {sheet.get('comment_count', 0)}",
                ]
            )
        )
        for cell in sheet.get("cells", [])[:24]:
            lines.append(f"{cell.get('ref', '')}: {cell.get('value', '')}")
        lines.append("")
    build_text_pdf(lines, target, title)


def build_docx_from_xlsx(sheets: list[dict[str, Any]], target: Path, title: str) -> None:
    """Write an Excel-to-Word DOCX summary from normalized sheet previews."""
    paragraphs = [{"text": f"Excel 表格提取：{title}", "style": "Heading1"}]
    for sheet in sheets:
        paragraphs.append({"text": f"Sheet {sheet.get('index')}: {sheet.get('name')}", "style": "Heading1"})
        if sheet.get("conversion_range"):
            paragraphs.append({"text": f"转换范围：{sheet.get('conversion_range')}", "style": "Normal"})
        if sheet.get("formula_mode"):
            paragraphs.append({"text": f"公式模式：{sheet.get('formula_mode')}", "style": "Normal"})
        paragraphs.append(
            {
                "text": f"单元格 {len(sheet.get('cells', []))} 个；公式 {sheet.get('formula_count', 0)} 个；合并单元格 {sheet.get('merged_cells', 0)} 个",
                "style": "Normal",
            }
        )
        paragraphs.append(
            {
                "text": f"图表 {sheet.get('chart_count', 0)} 个；图片 {sheet.get('image_count', 0)} 个；批注 {sheet.get('comment_count', 0)} 个；表格对象 {sheet.get('table_count', 0)} 个",
                "style": "Normal",
            }
        )
        for cell in sheet.get("cells", [])[:40]:
            paragraphs.append({"text": f"{cell.get('ref', '')}: {cell.get('value', '')}", "style": "Normal"})
    build_docx(paragraphs, target)


def build_pptx_from_xlsx(sheets: list[dict[str, Any]], target: Path, title: str) -> None:
    """Write an Excel-to-PPTX summary deck from normalized sheet previews."""
    slides: list[dict[str, Any]] = []
    for sheet in sheets:
        body = [
            f"转换范围 {sheet.get('conversion_range', '全部工作表')}",
            f"公式模式 {sheet.get('formula_mode', '保留公式')}",
            f"单元格 {len(sheet.get('cells', []))} 个",
            f"公式 {sheet.get('formula_count', 0)} 个",
            f"合并单元格 {sheet.get('merged_cells', 0)} 个",
            f"图表 {sheet.get('chart_count', 0)} 个",
            f"图片 {sheet.get('image_count', 0)} 个",
            f"批注 {sheet.get('comment_count', 0)} 个",
        ]
        body.extend(f"{cell.get('ref', '')}: {cell.get('value', '')}" for cell in sheet.get("cells", [])[:8])
        slides.append({"title": f"{title} · {sheet.get('name')}", "body": body})
    build_pptx(slides or [{"title": title, "body": ["未提取到工作表"]}], target)


def build_text_pdf(lines: list[str], target: Path, title: str = "K12 Export") -> None:
    """Build a small standards-compatible PDF from plain text lines."""
    target.parent.mkdir(parents=True, exist_ok=True)
    pages = _pdf_pages(lines or [title], 42)
    objects: list[bytes] = []
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    page_refs = " ".join(f"{3 + index * 2} 0 R" for index in range(len(pages)))
    objects.append(f"<< /Type /Pages /Kids [{page_refs}] /Count {len(pages)} >>".encode("ascii"))
    for index, page_lines in enumerate(pages):
        page_obj = 3 + index * 2
        content_obj = page_obj + 1
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >> /Contents {content_obj} 0 R >>".encode(
                "ascii"
            )
        )
        stream = _pdf_stream(page_lines)
        objects.append(f"<< /Length {len(stream)} >>\nstream\n".encode("ascii") + stream + b"\nendstream")
    _write_pdf(objects, target)


def build_docx(paragraphs: list[dict[str, Any]], target: Path) -> None:
    """Package paragraphs into a minimal DOCX artifact for downloads."""
    target.parent.mkdir(parents=True, exist_ok=True)
    body = "".join(_docx_paragraph(item["text"], item.get("style", "Normal")) for item in paragraphs)
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"""
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _docx_content_types())
        archive.writestr("_rels/.rels", _package_rels("word/document.xml"))
        archive.writestr("word/document.xml", document)


def build_pptx(slides: list[dict[str, Any]], target: Path) -> None:
    """Package slide dictionaries into a minimal PPTX artifact for downloads."""
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _pptx_content_types(len(slides)))
        archive.writestr("_rels/.rels", _package_rels("ppt/presentation.xml"))
        archive.writestr("ppt/presentation.xml", _presentation_xml(len(slides)))
        archive.writestr("ppt/_rels/presentation.xml.rels", _presentation_rels(len(slides)))
        for index, slide in enumerate(slides, start=1):
            archive.writestr(f"ppt/slides/slide{index}.xml", _slide_xml(slide.get("title", f"幻灯片 {index}"), slide.get("body", [])))


def _slides_from_blocks(blocks: list[dict[str, Any]], max_chars: int) -> list[dict[str, Any]]:
    slides: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for block in blocks:
        text = block["text"]
        if block.get("level", 0) == 1 or current is None:
            if current:
                slides.append(current)
            current = {"title": text[:80], "body": []}
            continue
        if current is None:
            current = {"title": "内容", "body": []}
        if sum(len(item) for item in current["body"]) + len(text) > max_chars and current["body"]:
            slides.append(current)
            current = {"title": current["title"], "body": []}
        current["body"].append(text)
    if current:
        slides.append(current)
    return slides or [{"title": "内容", "body": ["未提取到正文"]}]


def _docx_object_preservation_lines(plan: dict[str, Any]) -> list[str]:
    source = plan.get("source") or {}
    statuses = plan.get("statuses") or {}
    object_total = sum(int(source.get(key, 0) or 0) for key in ("images", "tables", "formulas"))
    if not object_total:
        return []
    lines = [
        f"图片 {source.get('images', 0)} 个：{statuses.get('images') or '已记录源图片对象'}",
        f"表格 {source.get('tables', 0)} 个：{statuses.get('tables') or '已记录源表格结构'}",
        f"公式 {source.get('formulas', 0)} 个：{statuses.get('formulas') or '已记录源公式对象'}",
        f"OMML {source.get('omml_formulas', 0)} 个；MathType {source.get('mathtype_objects', 0)} 个",
    ]
    notes = plan.get("notes") or []
    if isinstance(notes, list):
        lines.extend(str(item) for item in notes[:4] if item)
    return lines


def _extract_text(xml: str, tag: str) -> str:
    pieces = re.findall(fr"<{tag}[^>]*>([\s\S]*?)</{tag}>", xml)
    return _clean_text("".join(pieces))


def _xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        xml = archive.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
    except KeyError:
        return []
    strings: list[str] = []
    for item in re.findall(r"<si[\s\S]*?</si>", xml):
        text = "".join(re.findall(r"<t[^>]*>([\s\S]*?)</t>", item))
        strings.append(_clean_text(text))
    return strings


def _xlsx_sheet_names(archive: zipfile.ZipFile) -> dict[int, str]:
    try:
        xml = archive.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    except KeyError:
        return {}
    names: dict[int, str] = {}
    for index, match in enumerate(re.finditer(r"<sheet[^>]*name=\"([^\"]+)\"", xml), start=1):
        names[index] = html.unescape(match.group(1))
    return names


def _xlsx_cells(xml: str, shared_strings: list[str]) -> list[dict[str, str]]:
    cells: list[dict[str, str]] = []
    for cell in re.findall(r"<c\b[\s\S]*?</c>", xml):
        ref_match = re.search(r'r="([^"]+)"', cell)
        type_match = re.search(r't="([^"]+)"', cell)
        value_match = re.search(r"<v>([\s\S]*?)</v>", cell)
        inline_match = re.search(r"<is>[\s\S]*?<t[^>]*>([\s\S]*?)</t>[\s\S]*?</is>", cell)
        value = ""
        if inline_match:
            value = _clean_text(inline_match.group(1))
        elif value_match:
            value = _clean_text(value_match.group(1))
            if type_match and type_match.group(1) == "s":
                value = shared_strings[int(value)] if value.isdigit() and int(value) < len(shared_strings) else value
        formula_match = re.search(r"<f[^>]*>([\s\S]*?)</f>", cell)
        result = value
        formula = _clean_text(formula_match.group(1)) if formula_match else ""
        if formula_match:
            value = f"={formula}" if not value else f"{value} (={formula})"
        if value:
            cells.append({"ref": ref_match.group(1) if ref_match else "", "value": value, "result": result, "formula": formula})
    return cells


def _pptx_related_parts(archive: zipfile.ZipFile, part_path: str) -> list[str]:
    rels_path = f"{posixpath.dirname(part_path)}/_rels/{posixpath.basename(part_path)}.rels"
    try:
        rels_xml = archive.read(rels_path).decode("utf-8", errors="ignore")
    except KeyError:
        return []
    return [_resolve_xlsx_target(part_path, target) for target in re.findall(r'Target="([^"]+)"', rels_xml)]


def _pptx_notes(archive: zipfile.ZipFile, related_parts: list[str]) -> list[str]:
    notes: list[str] = []
    for part in related_parts:
        if "/notesSlides/" not in f"/{part}":
            continue
        try:
            notes_xml = archive.read(part).decode("utf-8", errors="ignore")
        except KeyError:
            continue
        notes.extend(_pptx_texts(notes_xml))
    return notes


def _pptx_texts(xml: str) -> list[str]:
    texts = re.findall(r"<a:t\b[^>]*>([\s\S]*?)</a:t>", xml)
    return [_clean_text(item) for item in texts if _clean_text(item)]


def _docx_slide_outline(slides: list[dict[str, Any]]) -> list[dict[str, Any]]:
    paragraphs: list[dict[str, Any]] = []
    for slide in slides:
        paragraphs.append({"text": slide["title"], "style": "Heading1"})
        for item in slide.get("body", [])[:12]:
            paragraphs.append({"text": item, "style": "Normal"})
    return paragraphs


def _pptx_slide_summary(
    slide: dict[str, Any],
    include_notes: bool = True,
    retain_images: bool = True,
    retain_formulas: bool = True,
) -> str:
    parts = [
        ("文本框", slide.get("text_box_count", 0)),
        ("图片", slide.get("image_count", 0) if retain_images else 0),
        ("表格", slide.get("table_count", 0)),
        ("公式", slide.get("formula_count", 0) if retain_formulas else 0),
        ("图表", slide.get("chart_count", 0)),
        ("形状", slide.get("shape_count", 0)),
        ("备注", len(slide.get("notes", []) or []) if include_notes else 0),
    ]
    return "；".join(f"{label} {int(count or 0)} 个" for label, count in parts if int(count or 0) > 0)


def _pptx_formula_count(xml: str) -> int:
    math_objects = len(re.findall(r"<(?:\w+:)?oMath(?:Para)?\b", xml))
    mathtype_refs = len(re.findall(r"MathType|Equation Native", xml, flags=re.IGNORECASE))
    latex_refs = len(re.findall(r"\$[^$]{1,120}\$|\\(?:frac|sqrt|sum|int)\b", xml))
    return math_objects + mathtype_refs + latex_refs


def _xlsx_sheet_objects(archive: zipfile.ZipFile, sheet_path: str, sheet_xml: str) -> dict[str, int]:
    rels_path = f"{posixpath.dirname(sheet_path)}/_rels/{posixpath.basename(sheet_path)}.rels"
    try:
        rels_xml = archive.read(rels_path).decode("utf-8", errors="ignore")
    except KeyError:
        rels_xml = ""
    related_parts = [_resolve_xlsx_target(sheet_path, target) for target in re.findall(r'Target="([^"]+)"', rels_xml)]
    drawing_paths = sorted({path for path in related_parts if "/drawings/" in f"/{path.lower()}"})
    comment_paths = sorted({path for path in related_parts if posixpath.basename(path).lower().startswith("comments")})
    table_paths = {path for path in related_parts if "/tables/" in f"/{path.lower()}"}
    table_count = max(len(re.findall(r"<(?:\w+:)?tablePart\b", sheet_xml)), len(table_paths))
    chart_count = 0
    image_count = 0
    for drawing_path in drawing_paths:
        try:
            drawing_xml = archive.read(drawing_path).decode("utf-8", errors="ignore")
        except KeyError:
            continue
        chart_tags = len(re.findall(r"<(?:\w+:)?chart\b", drawing_xml))
        image_tags = len(re.findall(r"<(?:\w+:)?pic\b", drawing_xml))
        drawing_rels = f"{posixpath.dirname(drawing_path)}/_rels/{posixpath.basename(drawing_path)}.rels"
        try:
            drawing_rels_xml = archive.read(drawing_rels).decode("utf-8", errors="ignore")
        except KeyError:
            drawing_rels_xml = ""
        drawing_related_parts = [
            _resolve_xlsx_target(drawing_path, target)
            for target in re.findall(r'Target="([^"]+)"', drawing_rels_xml)
        ]
        chart_count += max(chart_tags, len({path for path in drawing_related_parts if "/charts/" in f"/{path.lower()}"}))
        image_count += max(image_tags, len({path for path in drawing_related_parts if "/media/" in f"/{path.lower()}"}))
    comment_count = 0
    for comment_path in comment_paths:
        try:
            comments_xml = archive.read(comment_path).decode("utf-8", errors="ignore")
        except KeyError:
            continue
        comment_count += len(re.findall(r"<(?:\w+:)?comment\b", comments_xml))
    return {
        "drawing_count": len(drawing_paths),
        "chart_count": chart_count,
        "image_count": image_count,
        "comment_count": comment_count,
        "table_count": table_count,
    }


def _resolve_xlsx_target(source_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_path), target))


def _strip_xml(xml: str) -> str:
    return _clean_text(re.sub(r"<[^>]+>", "", xml))


def _clean_text(text: str) -> str:
    return html.unescape(re.sub(r"\s+", " ", text)).strip()


def _heading_level(paragraph: str, index: int) -> int:
    style_match = re.search(r'w:val="([^"]+)"', paragraph)
    style = style_match.group(1).lower() if style_match else ""
    if "heading1" in style or style in {"title", "1"} or index == 1:
        return 1
    if "heading2" in style:
        return 2
    return 0


def _slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _worksheet_number(path: str) -> int:
    match = re.search(r"sheet(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def _pdf_pages(lines: list[str], page_size: int) -> list[list[str]]:
    pages: list[list[str]] = []
    current: list[str] = []
    for line in lines:
        for chunk in _pdf_line_chunks(line):
            current.append(chunk)
            if len(current) >= page_size:
                pages.append(current)
                current = []
    if current:
        pages.append(current)
    return pages or [["K12 Export"]]


def _pdf_line_chunks(line: str, width: int = 92) -> list[str]:
    clean = _pdf_ascii(line)
    if not clean:
        return [""]
    return [clean[index : index + width] for index in range(0, len(clean), width)]


def _pdf_ascii(value: str) -> str:
    text = html.unescape(str(value))
    return "".join(char if 32 <= ord(char) <= 126 else "?" for char in text)


def _pdf_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def _pdf_stream(lines: list[str]) -> bytes:
    commands = ["BT", "/F1 11 Tf", "50 792 Td", "14 TL"]
    for index, line in enumerate(lines):
        if index:
            commands.append("T*")
        commands.append(f"({_pdf_escape(line)}) Tj")
    commands.append("ET")
    return "\n".join(commands).encode("ascii", errors="replace")


def _write_pdf(objects: list[bytes], target: Path) -> None:
    offsets: list[int] = []
    content = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    for index, obj in enumerate(objects, start=1):
        offsets.append(len(content))
        content.extend(f"{index} 0 obj\n".encode("ascii"))
        content.extend(obj)
        content.extend(b"\nendobj\n")
    xref_offset = len(content)
    content.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
    content.extend(b"0000000000 65535 f \n")
    for offset in offsets:
        content.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
    content.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii"))
    target.write_bytes(bytes(content))


def _docx_paragraph(text: str, style: str) -> str:
    style_xml = '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' if style == "Heading1" else ""
    return f"<w:p>{style_xml}<w:r><w:t>{html.escape(text)}</w:t></w:r></w:p>"


def _docx_content_types() -> str:
    return """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""


def _package_rels(target: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{target}"/></Relationships>"""


def _pptx_content_types(slide_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for index in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>{overrides}</Types>"""


def _presentation_xml(slide_count: int) -> str:
    slide_ids = "".join(f'<p:sldId id="{255 + index}" r:id="rId{index}"/>' for index in range(1, slide_count + 1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000" type="screen4x3"/></p:presentation>"""


def _presentation_rels(slide_count: int) -> str:
    rels = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{index}.xml"/>'
        for index in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"""


def _slide_xml(title: str, body: list[str]) -> str:
    body_text = "\n".join(body) if body else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>{html.escape(title)}</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/></p:nvSpPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>{html.escape(body_text)}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"""
